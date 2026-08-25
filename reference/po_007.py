"""PO07 - Mismo usuario crea la PO y registra el GR."""

from pathlib import Path

import pandas as pd
from time import perf_counter

from core.po_common import (
    get_period_suffix,
    load_po_lines,
    normalize_identifier,
    normalize_text,
    write_control_sheet,
)


CONTROL_ID = "PO_007"
CONTROL_NAME = "Same User Creates PO and GR"
SHEET_NAME = "PO07"

PO_REQUIRED_FIELDS = (
    "Company",
    "PO Number",
    "PO Line",
    "PO Doc Date",
    "PO Creator ID",
)

OUTPUT_COLUMNS = [
    "CoCo",
    "Company",
    "PO Number",
    "PO DocEntry",
    "PO Line",
    "Vendor Code",
    "Vendor Name",
    "PO Doc Date",
    "PO Doc Currency",
    "Company Main Currency",
    "PO Canceled",
    "PO Line Status",
    "Item Code",
    "Account Code",
    "PO Material Description",
    "PO Quantity",
    "PO Unit Price",
    "PO Line Total",
    "PO Line Total USD",
    "USD Rate",
    "USD Rate Date",
    "PO Creator ID",
    "PO Creator Name",
    "PO Approval Date",
    "PO Approver ID",
    "PO Approver Name",
    "PO Approval Status",
    "GR Doc Number",
    "GR Doc Date",
    "GR First Posting Date",
    "GR Last Posting Date",
    "GR Quantity",
    "GR Creator ID",
    "GR Creator Name",
    "PO Month",
    "PR DocEntry",
    "PR Line",
    "From PR",
]

GR_ALIASES = {
    "Company": "CoCd",
    "PO Number": "Purch.Doc.",
    "PO Line": "Item",
    "GR Doc Number": "Mat. Doc.",
    "GR Doc Date": "Doc. Date",
    "GR Posting Date": "Pstng Date",
    "GR Quantity": "Quantity",
    "GR Creator ID": "User name",
}


def _load_gr(context):
    """Carga LBR PR_GR_YYYYMMDD y normaliza los campos usados por PO07."""
    suffix = get_period_suffix(context["module"])
    folder = Path(context["input_folder"])

    matches = [
        path
        for path in folder.iterdir()
        if path.is_file()
        and path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}
        and path.stem.upper() == f"LBR PR_GR_{suffix}".upper()
    ]

    if len(matches) != 1:
        raise FileNotFoundError(
            f"Expected exactly one LBR PR_GR_{suffix}.XLSX; "
            f"found {len(matches)}."
        )

    gr = pd.read_excel(matches[0], sheet_name="Sheet1", dtype=object)
    gr = gr.dropna(how="all").rename(columns={
        source: target
        for target, source in GR_ALIASES.items()
    })

    missing = [
        column
        for column in GR_ALIASES
        if column not in gr.columns
    ]
    if missing:
        raise ValueError(f"Missing PO GR columns: {missing}")

    for column in ("Company", "PO Number", "PO Line", "GR Creator ID"):
        gr[column] = gr[column].map(normalize_identifier)

    gr["GR Doc Date"] = pd.to_datetime(
        gr["GR Doc Date"],
        errors="coerce",
        dayfirst=True,
    )
    gr["GR Posting Date"] = pd.to_datetime(
        gr["GR Posting Date"],
        errors="coerce",
        dayfirst=True,
    )
    gr["GR Quantity"] = pd.to_numeric(
        gr["GR Quantity"],
        errors="coerce",
    )

    return gr, matches[0]


def _build_exceptions(po, gr):
    """Devuelve una fila por movimiento GR cuyo usuario creó también la PO."""
    keys = ["Company", "PO Number", "PO Line"]

    po = po.copy()
    gr = gr.copy()

    for column in keys + ["PO Creator ID"]:
        po[column] = po[column].map(normalize_identifier)

    for column in keys + ["GR Creator ID"]:
        gr[column] = gr[column].map(normalize_identifier)

    # PO Lines debe aportar una sola fila por posición para evitar un join
    # muchos-a-muchos. Conservamos la primera representación de la posición.
    po = po.drop_duplicates(
        subset=keys,
        keep="first",
    )

    # Elimina duplicados exactos del movimiento GR, si los hubiera.
    gr = gr.drop_duplicates(
        subset=keys + [
            "GR Doc Number",
            "GR Doc Date",
            "GR Posting Date",
            "GR Creator ID",
        ],
        keep="first",
    )

    # First/Last se calculan directamente, sin hacer otro merge.
    group = gr.groupby(keys, dropna=False)["GR Posting Date"]
    gr["GR First Posting Date"] = group.transform("min")
    gr["GR Last Posting Date"] = group.transform("max")

    # one_to_many verifica que cada posición PO sea única en el lado izquierdo.
    result = po.merge(
        gr,
        on=keys,
        how="inner",
        validate="one_to_many",
    )

    creator = (
        result["PO Creator ID"]
        .map(normalize_text)
        .str.casefold()
    )
    receiver = (
        result["GR Creator ID"]
        .map(normalize_text)
        .str.casefold()
    )

    result = result.loc[
        creator.ne("")
        & receiver.ne("")
        & creator.eq(receiver)
    ].copy()

    result["CoCo"] = result["Company"]
    result["PO Month"] = pd.to_datetime(
        result["PO Doc Date"],
        errors="coerce",
        dayfirst=True,
    ).dt.strftime("%Y-%m")

    pr = result.get(
        "PR Number",
        pd.Series("", index=result.index, dtype="object"),
    ).map(normalize_text)

    result["PR DocEntry"] = pr
    result["From PR"] = pr.ne("").map({
        True: "Y",
        False: "N",
    })

    for column in OUTPUT_COLUMNS:
        if column not in result.columns:
            result[column] = ""

    return (
        result.loc[:, OUTPUT_COLUMNS]
        .sort_values(
            ["Company", "PO Number", "PO Line", "GR Doc Number"],
            kind="stable",
        )
        .reset_index(drop=True)
    )


def run_po_007(context):
    """Ejecuta PO07 y muestra cuánto demora cada etapa."""
    started = perf_counter()

    po, po_metrics = load_po_lines(
        context,
        required_fields=PO_REQUIRED_FIELDS,
    )
    after_po = perf_counter()

    gr, gr_file = _load_gr(context)
    after_gr = perf_counter()

    exceptions = _build_exceptions(po, gr)
    after_analysis = perf_counter()

    output_file = write_control_sheet(
        context=context,
        sheet_name=SHEET_NAME,
        dataframe=exceptions,
        date_columns=[
            "PO Doc Date",
            "USD Rate Date",
            "PO Approval Date",
            "GR Doc Date",
            "GR First Posting Date",
            "GR Last Posting Date",
        ],
        amount_columns=[
            "PO Quantity",
            "PO Unit Price",
            "PO Line Total",
            "PO Line Total USD",
            "USD Rate",
            "GR Quantity",
        ],
    )
    after_excel = perf_counter()

    status = "ERROR" if not exceptions.empty else "OK"

    print(f"{CONTROL_ID} - {CONTROL_NAME}")
    print(f"PO Lines: {po_metrics['input_file']}")
    print(f"PO GR: {gr_file}")
    print(f"PO rows after CONFIG filters: {len(po)}")
    print(f"GR rows read: {len(gr)}")
    print(f"Exception rows: {len(exceptions)}")
    print()
    print("Execution times")
    print(f"Load PO: {after_po - started:.2f} seconds")
    print(f"Load GR: {after_gr - after_po:.2f} seconds")
    print(f"Analysis: {after_analysis - after_gr:.2f} seconds")
    print(f"Write Excel: {after_excel - after_analysis:.2f} seconds")
    print(f"Total PO07: {after_excel - started:.2f} seconds")
    print()
    print(f"Control result: {status}")
    print(f"Output: {output_file} [{SHEET_NAME}]")

    return {
        "status": status,
        "output_file": output_file,
        "sheet_name": SHEET_NAME,
        "rows": len(exceptions),
    }
