"""PO09 - Split Purchase Requisitions (LBR SAP ECC)."""

from pathlib import Path
from time import perf_counter

import pandas as pd

from core.po_common import (
    get_period_suffix,
    normalize_company,
    parse_config_companies,
    parse_config_date,
    write_control_sheet,
)

CONTROL_ID = "PO_009"
SHEET_NAME = "PO09"
DEFAULT_WINDOW_DAYS = 7

GROUP = ["Company", "Item Code", "PR Creator ID"]
DOC = GROUP + ["PR Number", "PR Doc Date"]

OUTPUT_COLUMNS = [
    "CoCo",
    "Company",
    "PR Number",
    "PR DocEntry",
    "PR Line",
    "PR Doc Date",
    "PR Canceled",
    "PR Line Status",
    "Item Code",
    "Account Code",
    "PR Material Description",
    "PR Quantity",
    "PR Creator ID",
    "PR Creator Name",
    "PR Approval Date",
    "PR Approver ID",
    "PR Approver Name",
    "PR Approval Status",
    "PR Vendor Code (raw)",
    "Vendor Code (from PO)",
    "Vendor Name (from PO)",
    "Linked PO Number",
    "Linked PO DocEntry",
    "Linked PO Lines",
    "Linked PO Quantity",
    "Has PO",
    "PR Month",
    "SPLIT_PR_KEY",
    "Split Window Days",
]

RENAME = {
    "CoCd": "Company",
    "Purch.Req.": "PR Number",
    "Item": "PR Line",
    "Req.Date": "PR Doc Date",
    "Created by": "PR Creator ID",
    "Material": "Item Code",
    "Qty Requested": "PR Quantity",
    "D": "_Deletion Indicator",
    "S": "PR Line Status",
    "PO": "Linked PO Number",
    "Item.1": "_Linked PO Line",
    "Short Text": "PR Material Description",
    "Fix. Vend.": "PR Vendor Code (raw)",
}

REQUIRED_HEADERS = set(RENAME)

SOURCE_MAP = {
    "Company": "Company",
    "PR Number": "PR Number",
    "PR Line": "PR Line",
    "PR Doc Date": "PR Doc Date",
    "PR Line Status": "PR Line Status",
    "Item Code": "Item Code",
    "PR Material Description": "PR Material Description",
    "PR Quantity": "PR Quantity",
    "PR Creator ID": "PR Creator ID",
    "PR Vendor Code (raw)": "PR Vendor Code (raw)",
    "Linked PO Number": "Linked PO Number",
    "PR Month": "PR Month",
    "SPLIT_PR_KEY": "SPLIT_PR_KEY",
    "Split Window Days": "Split Window Days",
}


def _text(series):
    """Normalización vectorizada sin convertir nulos en el texto 'nan'."""
    return (
        series.astype("string")
        .str.strip()
        .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "<NA>": pd.NA})
    )


def _identifier(series):
    """Conserva identificadores de texto y elimina el sufijo Excel '.0'."""
    return _text(series).str.replace(r"^(\d+)\.0$", r"\1", regex=True)


def _window_days(context):
    """PARAM1 vacío/NaN usa 7 días; cualquier otro valor debe ser entero >= 0."""
    raw = context.get("control", {}).get("param1")

    if raw is None or pd.isna(raw) or str(raw).strip().casefold() in {
        "",
        "nan",
        "none",
        "<na>",
    }:
        return DEFAULT_WINDOW_DAYS

    value = pd.to_numeric(raw, errors="coerce")

    if pd.isna(value) or value < 0 or float(value) != int(value):
        raise ValueError(
            f"{CONTROL_ID} PARAM1 debe ser un número entero no negativo; "
            f"valor recibido: {raw!r}."
        )

    return int(value)


def _find_input(context):
    return find_period_input_file(
        context=context,
        input_prefix=PR_LINES_INPUT_PREFIX,
        source_name="PR Lines",
    )


def _load_pr_lines(context):
    """Abre y lee PR_Lines una sola vez mediante un único ExcelFile."""
    path = _find_input(context)

    with pd.ExcelFile(path) as book:
        sheets = {
            str(sheet).strip().casefold(): sheet
            for sheet in book.sheet_names
        }
        sheet = sheets.get("sheet1")

        if sheet is None:
            raise ValueError(
                f"{path.name} no contiene una hoja Sheet1. "
                f"Hojas disponibles: {book.sheet_names}."
            )

        raw = book.parse(sheet_name=sheet, dtype=object)

    raw = raw.dropna(how="all").dropna(axis=1, how="all")
    raw.columns = raw.columns.astype(str).str.strip()

    missing = sorted(REQUIRED_HEADERS - set(raw.columns))
    if missing:
        raise ValueError(
            f"PR Lines no contiene los headers requeridos: {missing}. "
            f"Headers disponibles: {list(raw.columns)}."
        )

    data = raw.rename(columns=RENAME)[list(RENAME.values())].copy()

    key = ["Company", "PR Number", "PR Line"]
    for column in key + [
        "PR Creator ID",
        "Item Code",
        "Linked PO Number",
        "_Linked PO Line",
        "PR Vendor Code (raw)",
    ]:
        data[column] = _identifier(data[column])

    data["Company"] = data["Company"].map(
        lambda value: normalize_company(value) if pd.notna(value) else ""
    )

    blank_key = data[key].isna() | data[key].eq("")
    residual = blank_key.all(axis=1)
    partial = blank_key.any(axis=1) & ~residual

    if partial.any():
        raise ValueError(
            f"PR Lines contiene {int(partial.sum())} filas con clave parcial "
            "Company + PR Number + PR Line."
        )

    data = data.loc[~residual].copy()

    duplicates = data.duplicated(key, keep=False)
    if duplicates.any():
        raise ValueError(
            f"PR Lines contiene {int(duplicates.sum())} filas con clave "
            "Company + PR Number + PR Line duplicada."
        )

    raw_dates = _text(data["PR Doc Date"])
    parsed_dates = pd.to_datetime(
        data["PR Doc Date"],
        errors="coerce",
        dayfirst=True,
    ).dt.normalize()

    invalid_dates = raw_dates.notna() & parsed_dates.isna()
    if invalid_dates.any():
        raise ValueError(
            f"PR Lines contiene {int(invalid_dates.sum())} fechas de PR inválidas."
        )

    data["PR Doc Date"] = parsed_dates
    data["PR Quantity"] = pd.to_numeric(
        data["PR Quantity"],
        errors="coerce",
    )

    companies = parse_config_companies(
        context["module"].get("companies", "")
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
        raise ValueError("CONFIG FROM no puede ser posterior a CONFIG TO.")

    population = data["PR Doc Date"].between(
        date_from,
        date_to,
        inclusive="both",
    )

    if companies is not None:
        population &= data["Company"].isin(companies)

    data = data.loc[population].copy()

    if data.empty:
        raise ValueError("PO09 quedó sin población luego de aplicar CONFIG.")

    data["PR Month"] = data["PR Doc Date"].dt.to_period("M").astype(str)

    return data, {
        "input_file": str(path),
        "rows_read": len(raw),
        "residual_rows": int(residual.sum()),
        "rows_after_config_filters": len(data),
    }


def _find_split_prs(data, window_days):
    """Detección O(n log n), sin bucles por registro ni cruces many-to-many."""
    eligible = data.loc[
        data["Item Code"].notna()
        & data["Item Code"].ne("")
        & data["PR Doc Date"].notna()
    ].copy()

    documents = (
        eligible[DOC]
        .drop_duplicates(GROUP + ["PR Number"])
        .sort_values(GROUP + ["PR Doc Date", "PR Number"])
    )

    grouped = documents.groupby(
        GROUP,
        sort=False,
        dropna=False,
    )["PR Doc Date"]

    previous_gap = grouped.diff().dt.days
    next_gap = grouped.shift(-1).sub(documents["PR Doc Date"]).dt.days

    flagged_documents = documents.loc[
        previous_gap.le(window_days) | next_gap.le(window_days),
        GROUP + ["PR Number"],
    ].drop_duplicates()

    if flagged_documents.empty:
        result = data.iloc[0:0].copy()
        result["SPLIT_PR_KEY"] = pd.Series(dtype="object")
        result["Split Window Days"] = pd.Series(dtype="int64")
        return result

    result = eligible.merge(
        flagged_documents,
        on=GROUP + ["PR Number"],
        how="inner",
        validate="many_to_one",
    )

    result["SPLIT_PR_KEY"] = (
        result["Company"].fillna("")
        + "|"
        + result["Item Code"].fillna("")
        + "|"
        + result["PR Creator ID"].fillna("")
    )
    result["Split Window Days"] = window_days

    return result.sort_values(
        GROUP + ["PR Doc Date", "PR Number", "PR Line"]
    ).reset_index(drop=True)


def _lha_output(exceptions):
    """Mantiene el layout LHA sin fabricar campos ausentes en ECC."""
    output = pd.DataFrame("", index=exceptions.index, columns=OUTPUT_COLUMNS)

    for target, source in SOURCE_MAP.items():
        output[target] = exceptions[source]

    # Hasta contar con un maestro CoCo/Company, ambos conservan BUKRS.
    output["CoCo"] = exceptions["Company"]

    linked = (
        exceptions["Linked PO Number"].notna()
        & exceptions["Linked PO Number"].ne("")
    )
    output["Has PO"] = linked.map({True: "Y", False: "N"})

    # EBAN expone una línea de PO vinculada por línea de PR.
    output["Linked PO Lines"] = linked.astype("int64")

    return output


def run_po_009(context):
    """Ejecuta PO09 y reemplaza únicamente la hoja PO09."""
    total_start = perf_counter()

    start = perf_counter()
    pr_lines, metrics = _load_pr_lines(context)
    load_seconds = perf_counter() - start

    start = perf_counter()
    window_days = _window_days(context)
    preparation_seconds = perf_counter() - start

    start = perf_counter()
    exceptions = _find_split_prs(pr_lines, window_days)
    output = _lha_output(exceptions)
    analytic_seconds = perf_counter() - start

    # PO09 no contiene importes ni requiere conversión monetaria.
    fx_start = perf_counter()
    fx_seconds = perf_counter() - fx_start

    start = perf_counter()
    output_file = write_control_sheet(
        context=context,
        sheet_name=SHEET_NAME,
        dataframe=output,
        date_columns=["PR Doc Date", "PR Approval Date"],
        amount_columns=["PR Quantity", "Linked PO Quantity"],
    )
    write_seconds = perf_counter() - start

    total_seconds = perf_counter() - total_start
    status = "ERROR" if len(output) else "OK"

    print(f"PO09 input: {metrics['input_file']}")
    print(f"PO09 rows read: {metrics['rows_read']}")
    print(
        "PO09 timing | "
        f"load={load_seconds:.3f}s | "
        f"prepare={preparation_seconds:.3f}s | "
        f"analytic={analytic_seconds:.3f}s | "
        f"fx={fx_seconds:.3f}s | "
        f"write={write_seconds:.3f}s | "
        f"total={total_seconds:.3f}s"
    )

    return {
        "status": status,
        "output_file": output_file,
        "sheet_name": SHEET_NAME,
        "rows": len(output),
        "metrics": metrics,
    }
