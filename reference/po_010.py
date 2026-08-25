"""PO10 - diferencias de material/cantidad entre PO y PR vinculada (LBR SAP ECC)."""

from pathlib import Path
from time import perf_counter

import pandas as pd

from core.po_common import (
    get_period_suffix,
    normalize_company,
    normalize_identifier,
    normalize_text,
    parse_config_companies,
    parse_config_date,
    write_control_sheet,
)


CONTROL_ID = "PO_010"
CONTROL_NAME = "PO versus PR"
SHEET_NAME = "PO10"

KEY = ["Company", "PR Number", "PR Line"]
PO_KEY = ["Company", "PO Number", "PO Line"]

OUTPUT_COLUMNS = [
    "Company",
    "PR Number",
    "PR Line",
    "PO Number",
    "PO Line",
    "PR Item",
    "PO Item",
    "Item Match",
    "PR Quantity",
    "PO Quantity",
    "Qty Difference",
    "Vendor Name",
    "PO Material Description",
]

PO_COLUMNS = {
    "CoCd": "Company",
    "Purch.Doc.": "PO Number",
    "Item": "PO Line",
    "Doc. Date": "PO Date",
    "Material": "PO Item",
    "PO Quantity": "PO Quantity",
    "Short Text": "PO Material Description",
    "Purch.Req.": "PR Number",
    "Item.1": "PR Line",
}

PR_COLUMNS = {
    "CoCd": "Company",
    "Purch.Req.": "PR Number",
    "Item": "PR Line",
    "Req.Date": "PR Date",
    "Material": "PR Item",
    "Qty Requested": "PR Quantity",
}

VENDOR_NAME_HEADERS = (
    "Vendor Name",
    "Name 1",
    "NAME1",
    "LFA1-NAME1",
)


def _filename_key(value):
    """Compara nombres ignorando espacios, guiones y underscores."""
    return "".join(
        character
        for character in str(value).casefold()
        if character.isalnum()
    )


def _find_input(context, label):
    """Encuentra exactamente una bajada del tipo y período solicitados."""
    folder = Path(context["input_folder"])
    suffix = get_period_suffix(context["module"])
    expected = _filename_key(f"LBR {label} {suffix}")

    matches = [
        path
        for path in folder.rglob("*")
        if path.is_file()
        and path.suffix.casefold() in {".xlsx", ".xlsm", ".xls"}
        and not path.name.startswith("~$")
        and _filename_key(path.stem) == expected
    ]

    if len(matches) != 1:
        raise ValueError(
            f"{CONTROL_ID}: se esperaba un archivo '{label}' para {suffix}; "
            f"encontrados={len(matches)}: {matches}"
        )

    return matches[0]


def _read(context, label, columns):
    """Lee una bajada una sola vez y valida sus headers."""
    path = _find_input(context, label)

    try:
        raw = pd.read_excel(
            path,
            sheet_name="Sheet1",
            dtype=object,
        )
    except PermissionError as error:
        raise PermissionError(
            f"{CONTROL_ID}: el archivo está abierto o bloqueado: {path}"
        ) from error

    raw = raw.dropna(how="all").dropna(axis=1, how="all")

    if raw.empty:
        raise ValueError(f"{CONTROL_ID}: {label} no contiene datos.")

    missing = [column for column in columns if column not in raw.columns]

    if missing:
        raise ValueError(
            f"{CONTROL_ID}: faltan headers en {label}: {missing}. "
            f"Disponibles: {list(raw.columns)}"
        )

    result = raw[list(columns)].rename(columns=columns).copy()

    if label == "PO Lines":
        vendor_column = next(
            (column for column in VENDOR_NAME_HEADERS if column in raw.columns),
            None,
        )
        result["Vendor Name"] = (
            raw[vendor_column]
            if vendor_column is not None
            else pd.NA
        )

    return result, {
        "file": path,
        "rows_read": len(result),
    }


def _filter_scope(dataframe, date_column, context, source):
    """Normaliza compañía y aplica CONFIG COMPANIES/FROM/TO."""
    result = dataframe.copy()
    result["Company"] = result["Company"].map(normalize_company)

    companies = parse_config_companies(
        context["module"].get("companies", "")
    )

    if companies is None:
        excluded_company = 0
    else:
        company_mask = result["Company"].isin(companies)
        excluded_company = int((~company_mask).sum())
        result = result.loc[company_mask].copy()

    raw_date = result[date_column]
    parsed_date = pd.to_datetime(
        raw_date,
        errors="coerce",
        dayfirst=True,
    ).dt.normalize()

    invalid_date = raw_date.map(normalize_text).ne("") & parsed_date.isna()

    if invalid_date.any():
        raise ValueError(
            f"{CONTROL_ID}: {source} contiene {int(invalid_date.sum())} "
            f"fechas inválidas en {date_column}."
        )

    date_from = parse_config_date(
        context["module"].get("from", ""),
        "FROM",
    )
    date_to = parse_config_date(
        context["module"].get("to", ""),
        "TO",
    )

    if date_from > date_to:
        raise ValueError(
            f"{CONTROL_ID}: CONFIG FROM {date_from.date()} es posterior "
            f"a TO {date_to.date()}."
        )

    date_mask = parsed_date.between(
        date_from,
        date_to,
        inclusive="both",
    )
    excluded_date = int((~date_mask).sum())

    result = result.loc[date_mask].copy()
    result[date_column] = parsed_date.loc[date_mask]

    return result, {
        "excluded_company": excluded_company,
        "excluded_date": excluded_date,
    }


def _deduplicate(dataframe, key, source):
    """Quita copias exactas y falla ante claves repetidas conflictivas."""
    before = len(dataframe)
    result = dataframe.drop_duplicates().copy()
    exact_duplicates = before - len(result)

    conflicts = result.duplicated(key, keep=False)

    if conflicts.any():
        conflicting_keys = result.loc[conflicts, key].drop_duplicates()
        sample = (
            result.loc[conflicts]
            .sort_values(key, kind="stable")
            .head(20)
        )

        raise ValueError(
            f"{CONTROL_ID}: {source} contiene {len(conflicting_keys)} claves "
            f"duplicadas con datos diferentes. Ejemplos:\n"
            f"{sample.to_string(index=False)}"
        )

    return result, exact_duplicates


def _load_po(context):
    """Carga las líneas de PO que poseen una referencia completa a PR."""
    dataframe, metrics = _read(context, "PO Lines", PO_COLUMNS)

    for column in ("PO Number", "PO Line", "PR Number", "PR Line"):
        dataframe[column] = dataframe[column].map(normalize_identifier)

    pr_number = dataframe["PR Number"].map(normalize_text)
    pr_line = dataframe["PR Line"].map(normalize_text)

    # En ECC, BANFN vacío/0 y BNFPO vacío/0 representan "sin PR".
    missing_pr_number = pr_number.isin({"", "0"})
    missing_pr_line = pr_line.isin({"", "0"})

    without_pr = missing_pr_number & missing_pr_line
    incomplete_link = missing_pr_number ^ missing_pr_line

    if incomplete_link.any():
        sample = dataframe.loc[
            incomplete_link,
            PO_KEY + ["PR Number", "PR Line"],
        ].head(10)

        raise ValueError(
            f"{CONTROL_ID}: PO Lines contiene "
            f"{int(incomplete_link.sum())} referencias parciales a PR. "
            f"Ejemplos:\n{sample.to_string(index=False)}"
        )

    dataframe = dataframe.loc[~without_pr].copy()
    dataframe, scope = _filter_scope(
        dataframe,
        "PO Date",
        context,
        "PO Lines",
    )

    dataframe["PO Item"] = dataframe["PO Item"].map(normalize_text)
    dataframe["PO Material Description"] = (
        dataframe["PO Material Description"].map(normalize_text)
    )
    dataframe["Vendor Name"] = dataframe["Vendor Name"].map(normalize_text)
    dataframe["PO Quantity"] = pd.to_numeric(
        dataframe["PO Quantity"],
        errors="coerce",
    )

    dataframe, exact_duplicates = _deduplicate(
        dataframe,
        PO_KEY,
        "PO Lines",
    )

    metrics.update(scope)
    metrics.update({
        "rows_without_pr": int(without_pr.sum()),
        "exact_duplicates": exact_duplicates,
        "rows_eligible": len(dataframe),
        "vendor_name_available": dataframe["Vendor Name"].ne("").any(),
    })

    return dataframe, metrics


def _load_pr(context):
    """Carga y valida la población de líneas de PR."""
    dataframe, metrics = _read(context, "PR Lines", PR_COLUMNS)

    dataframe["Company"] = dataframe["Company"].map(normalize_company)

    for column in ("PR Number", "PR Line"):
        dataframe[column] = dataframe[column].map(normalize_identifier)

    blank = pd.DataFrame(
        {
            column: dataframe[column].map(normalize_text).eq("")
            for column in KEY
        },
        index=dataframe.index,
    )

    residual = blank.all(axis=1)
    partial = blank.any(axis=1) & ~residual

    if partial.any():
        sample = dataframe.loc[partial, KEY].head(10)

        raise ValueError(
            f"{CONTROL_ID}: PR Lines contiene {int(partial.sum())} "
            f"claves parciales. Ejemplos:\n{sample.to_string(index=False)}"
        )

    dataframe = dataframe.loc[~residual].copy()
    dataframe, scope = _filter_scope(
        dataframe,
        "PR Date",
        context,
        "PR Lines",
    )

    dataframe["PR Item"] = dataframe["PR Item"].map(normalize_text)
    dataframe["PR Quantity"] = pd.to_numeric(
        dataframe["PR Quantity"],
        errors="coerce",
    )

    dataframe, exact_duplicates = _deduplicate(
        dataframe,
        KEY,
        "PR Lines",
    )

    metrics.update(scope)
    metrics.update({
        "residual": int(residual.sum()),
        "exact_duplicates": exact_duplicates,
        "rows_eligible": len(dataframe),
    })

    return dataframe, metrics


def _exceptions(po, pr):
    """Cruza PO con PR y devuelve solamente diferencias."""
    joined = po.merge(
        pr[KEY + ["PR Item", "PR Quantity"]],
        on=KEY,
        how="inner",
        validate="many_to_one",
    )

    joined["Item Match"] = (
        joined["PR Item"]
        .eq(joined["PO Item"])
        .map({True: "Y", False: "N"})
    )
    joined["Qty Difference"] = (
        joined["PO Quantity"] - joined["PR Quantity"]
    ).round(6)

    item_difference = joined["Item Match"].eq("N")
    quantity_difference = joined["Qty Difference"].ne(0)
    exceptions = item_difference | quantity_difference

    result = (
        joined.loc[exceptions, OUTPUT_COLUMNS]
        .sort_values(
            KEY + ["PO Number", "PO Line"],
            kind="stable",
        )
        .reset_index(drop=True)
    )

    metrics = {
        "PO rows sent to merge": len(po),
        "PO rows matched with PR": len(joined),
        "PO rows without matching PR": len(po) - len(joined),
        "Rows with item difference": int(item_difference.sum()),
        "Rows with quantity difference": int(quantity_difference.sum()),
        "Exception rows": len(result),
    }

    return result, metrics


def _print_metrics(title, metrics):
    """Imprime métricas escalares sin trabajo analítico adicional."""
    print(title)

    for name, value in metrics.items():
        print(f"{name}: {value}")


def run_po_010(context):
    """Ejecuta LBR PO10 y reemplaza solamente la hoja PO10."""
    total_start = perf_counter()
    timing = {}

    started = perf_counter()
    po, po_metrics = _load_po(context)
    timing["PO load/prepare"] = perf_counter() - started

    started = perf_counter()
    pr, pr_metrics = _load_pr(context)
    timing["PR load/prepare"] = perf_counter() - started

    started = perf_counter()
    result, analytic_metrics = _exceptions(po, pr)
    timing["Analytic"] = perf_counter() - started

    # PO10 no utiliza importes ni conversión monetaria.
    timing["FX"] = 0.0

    started = perf_counter()
    output_file = write_control_sheet(
        context=context,
        sheet_name=SHEET_NAME,
        dataframe=result,
        amount_columns=[
            "PR Quantity",
            "PO Quantity",
            "Qty Difference",
        ],
    )
    timing["Workbook write"] = perf_counter() - started
    timing["Total"] = perf_counter() - total_start

    print()
    print(f"{CONTROL_ID} - {CONTROL_NAME}")
    print("-" * (len(CONTROL_ID) + len(CONTROL_NAME) + 3))

    _print_metrics("PO metrics", po_metrics)
    print()
    _print_metrics("PR metrics", pr_metrics)
    print()
    _print_metrics("Analytic metrics", analytic_metrics)

    if not po_metrics["vendor_name_available"]:
        print()
        print(
            "Audit warning: Vendor Name queda vacío porque PO Lines "
            "no contiene LFA1-NAME1/Vendor Name."
        )

    print()
    print("Timing")

    for phase, seconds in timing.items():
        suffix = " (no aplica)" if phase == "FX" else ""
        print(f"{phase}: {seconds:.3f}s{suffix}")

    print()
    print(f"Output: {output_file}")
    print(f"Sheet: {SHEET_NAME}")
    print()

    return {
        "status": "ERROR" if not result.empty else "OK",
        "output_file": output_file,
        "sheet_name": SHEET_NAME,
        "rows": len(result),
        "po_metrics": po_metrics,
        "pr_metrics": pr_metrics,
        "analytic_metrics": analytic_metrics,
        "timing": timing,
    }


run = run_po_010
