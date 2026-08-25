"""PO05 - GR registrado mas de N dias despues de aprobar la PO.

Control LBR independiente: lee PO Lines, PR_GR, PO_CDHDR y PO_CDPOS
originales. PARAM1 indica los dias permitidos; si esta vacio usa 30.
"""

from pathlib import Path

import pandas as pd

from modules.PO.po_003 import (
    CDHDR_ALIASES, CDHDR_INPUT_PREFIX, CDPOS_ALIASES, CDPOS_INPUT_PREFIX,
    _load_change_input, _select_lha_approval_event, build_release_events,
)
from core.po_common import (
    ALLOWED_INPUT_EXTENSIONS, get_period_suffix, inspect_input_workbook,
    load_po_lines, normalize_company, normalize_identifier, normalize_lookup,
    normalize_text, resolve_sheet_name, write_control_sheet,
)

CONTROL_ID = "PO_005"
CONTROL_NAME = "GR More Than N Days After PO Approval"
SHEET_NAME = "PO05"
DEFAULT_DAYS = 30
PO_KEY = ["Company", "PO Number"]
LINE_KEY = PO_KEY + ["PO Line"]

PO_FIELDS = (
    "Company", "PO Number", "PO Line", "Vendor Code", "PO Doc Date",
    "PO Creator ID", "Item Code", "PO Quantity", "PO Unit Price",
    "PO Doc Currency", "PO Line Total", "PO Material Description",
    "PO Line Deleted", "PO Delivery Completed", "PR Number", "PR Line",
)

GR_ALIASES = {
    "Company": ("CoCd", "Company", "BUKRS", "EKKO-BUKRS"),
    "PO Number": ("Purch.Doc.", "PO Number", "EBELN", "EKBE-EBELN"),
    "PO Line": ("Item", "PO Line", "EBELP", "EKBE-EBELP"),
    "GR Doc Number": (
        "Mat. Doc.", "Material Document", "BELNR", "EKBE-BELNR",
    ),
    "GR Fiscal Year": (
        "MatYr", "Fiscal Year", "GJAHR", "EKBE-GJAHR",
    ),
    "GR Doc Line": (
        "Item.1", "Document Line", "BUZEI", "EKBE-BUZEI",
    ),
    "GR Posting Date": (
        "Pstng Date", "Posting Date", "BUDAT", "EKBE-BUDAT",
    ),
    "GR Event Type": (
        "Tr./ev.type", "Transaction/event type", "VGABE", "EKBE-VGABE",
    ),
    "GR Quantity": (
        "Quantity", "MENGE", "EKBE-MENGE",
    ),
    "GR Creator ID": (
        "User name", "USNAM", "MKPF-USNAM",
    ),
}

OUTPUT_COLUMNS = [
    "CoCo", "Company", "PO Number", "PO DocEntry", "PO Line", "Vendor Code",
    "Vendor Name", "PO Doc Date", "PO Doc Currency", "Company Main Currency",
    "PO Canceled", "PO Line Status", "Item Code", "Account Code",
    "PO Material Description", "PO Quantity", "PO Unit Price", "PO Line Total",
    "PO Line Total USD", "USD Rate", "USD Rate Date", "PO Creator ID",
    "PO Creator Name", "PO Approval Date", "PO Approver ID", "PO Approver Name",
    "PO Approval Status", "GR Doc Number", "GR Doc Date",
    "GR First Posting Date", "GR Last Posting Date", "GR Quantity",
    "GR Creator ID", "GR Creator Name", "PO Month", "PR DocEntry", "PR Line",
    "From PR", "Days GR After Approval", "Threshold Days",
]


def _resolve_columns(df, aliases):
    """Relaciona cada campo logico con un unico header real de SAP."""
    available = {}

    for column in df.columns:
        available.setdefault(
            normalize_lookup(column),
            [],
        ).append(column)

    mapping, problems = {}, []

    for logical, candidates in aliases.items():
        matches = list(dict.fromkeys(
            match
            for candidate in candidates
            for match in available.get(
                normalize_lookup(candidate),
                [],
            )
        ))

        if len(matches) == 1:
            mapping[logical] = matches[0]
        else:
            problems.append(
                f"{logical}={matches or 'missing'}"
            )

    if problems:
        raise ValueError(
            f"PO GR header validation failed: {problems}. "
            f"Headers: {list(df.columns)}"
        )

    return mapping


def _find_gr(context):
    return find_period_input_file(
        context=context,
        input_prefix=PR_GR_INPUT_PREFIX,
        source_name="PO GR",
    )


def _load_gr(context):
    """Carga GR, elimina residuos y conserva movimientos EKBE-VGABE = 1."""
    path = _find_gr(context)

    sheet = resolve_sheet_name(
        path,
        "Sheet1",
    )

    workbook_metrics = inspect_input_workbook(
        path,
        sheet,
    )

    raw = pd.read_excel(
        path,
        sheet_name=sheet,
        dtype=object,
    )

    raw = (
        raw
        .dropna(how="all")
        .dropna(axis=1, how="all")
    )

    if raw.empty:
        raise ValueError(
            "PO GR contains no data rows."
        )

    columns = _resolve_columns(
        raw,
        GR_ALIASES,
    )

    key_names = LINE_KEY + [
        "GR Doc Number",
        "GR Fiscal Year",
        "GR Doc Line",
    ]

    physical_key = [
        columns[name]
        for name in key_names
    ]

    blank = raw[physical_key].apply(
        lambda column: column.map(
            normalize_text
        ).eq("")
    )

    residual = blank.all(axis=1)

    partial = (
        blank.any(axis=1)
        & ~residual
    )

    if partial.any():
        raise ValueError(
            f"PO GR has {int(partial.sum())} "
            "partially blank keys."
        )

    gr = (
        raw.loc[~residual]
        .rename(
            columns={
                physical: logical
                for logical, physical in columns.items()
            }
        )
        .copy()
    )

    for column in key_names[1:]:
        gr[column] = (
            gr[column]
            .map(normalize_identifier)
        )

    gr["Company"] = (
        gr["Company"]
        .map(normalize_company)
    )

    gr["GR Event Type"] = (
        gr["GR Event Type"]
        .map(normalize_identifier)
    )

    gr = gr.loc[
        gr["GR Event Type"].eq("1")
    ].copy()

    raw_date = gr["GR Posting Date"]

    gr["GR Posting Date"] = pd.to_datetime(
        raw_date,
        errors="coerce",
        dayfirst=True,
    ).dt.normalize()

    invalid = (
        raw_date.map(normalize_text).ne("")
        & gr["GR Posting Date"].isna()
    )

    if invalid.any():
        raise ValueError(
            f"PO GR has {int(invalid.sum())} "
            "invalid posting dates."
        )

    gr = gr.loc[
        gr["GR Posting Date"].notna()
    ].copy()

    gr["GR Quantity"] = pd.to_numeric(
        gr["GR Quantity"],
        errors="coerce",
    )

    return gr, {
        **workbook_metrics,
        "rows_read": len(raw),
        "gr_rows": len(gr),
    }


def _days(context):
    raw = (
        context
        .get("control", {})
        .get("param1", "")
    )

    if normalize_text(raw) == "":
        return DEFAULT_DAYS

    value = pd.to_numeric(
        raw,
        errors="coerce",
    )

    if (
        pd.isna(value)
        or value < 0
        or not float(value).is_integer()
    ):
        raise ValueError(
            "PO05 PARAM1 must be a "
            "non-negative whole number."
        )

    return int(value)


def _aggregate_gr(gr):
    ordered = gr.sort_values(
        LINE_KEY
        + [
            "GR Posting Date",
            "GR Doc Number",
        ]
    )

    result = (
        ordered
        .groupby(
            LINE_KEY,
            as_index=False,
        )
        .agg(
            **{
                "GR Doc Number": (
                    "GR Doc Number",
                    "last",
                ),
                "GR First Posting Date": (
                    "GR Posting Date",
                    "min",
                ),
                "GR Last Posting Date": (
                    "GR Posting Date",
                    "max",
                ),
                "GR Quantity": (
                    "GR Quantity",
                    "sum",
                ),
                "GR Creator ID": (
                    "GR Creator ID",
                    "last",
                ),
            }
        )
    )

    result["GR Doc Date"] = (
        result["GR Last Posting Date"]
    )

    return result


def _output(
    po,
    approvals,
    receipts,
    days,
):
    approval = approvals[
        PO_KEY
        + [
            "Approval Timestamp",
            "PO Approver ID",
        ]
    ].rename(
        columns={
            "Approval Timestamp": "PO Approval Date",
        }
    )

    data = po.merge(
        approval,
        on=PO_KEY,
        how="left",
        validate="many_to_one",
    )

    data = data.merge(
        receipts,
        on=LINE_KEY,
        how="left",
        validate="many_to_one",
    )

    difference = (
        data["GR Last Posting Date"]
        - data["PO Approval Date"]
    ).dt.days

    data = data.loc[
        difference.gt(days)
    ].copy()

    data["Days GR After Approval"] = (
        difference.loc[data.index]
    )

    data["Threshold Days"] = days

    out = pd.DataFrame(
        index=data.index
    )

    def source(name):
        return data.get(
            name,
            pd.Series(
                "",
                index=data.index,
                dtype=object,
            ),
        )

    for column in OUTPUT_COLUMNS:
        out[column] = source(column)

    out["CoCo"] = source("Company")
    out["PO Approval Status"] = "EFFECTIVE RELEASE"

    out["PO Month"] = (
        pd.to_datetime(
            source("PO Doc Date"),
            errors="coerce",
        )
        .dt.strftime("%Y-%m")
        .fillna("")
    )

    out["PR DocEntry"] = source(
        "PR Number"
    )

    out["From PR"] = (
        source("PR Number")
        .map(normalize_text)
        .ne("")
        .map({
            True: "Y",
            False: "N",
        })
    )

    out["PO Line Status"] = ""

    out.loc[
        source("PO Line Deleted")
        .map(normalize_text)
        .ne(""),
        "PO Line Status",
    ] = "DELETED"

    completed = (
        source("PO Line Deleted")
        .map(normalize_text)
        .eq("")
        & source("PO Delivery Completed")
        .map(normalize_text)
        .ne("")
    )

    out.loc[
        completed,
        "PO Line Status",
    ] = "COMPLETED"

    return (
        out
        .sort_values(LINE_KEY)
        .reset_index(drop=True)
    )


def run_po_005(context):
    """Ejecuta PO05 y reemplaza solamente la hoja PO05."""
    days = _days(context)

    po, po_metrics = load_po_lines(
        context,
        required_fields=PO_FIELDS,
    )

    gr, gr_metrics = _load_gr(
        context
    )

    cdhdr, _ = _load_change_input(
        context,
        CDHDR_INPUT_PREFIX,
        CDHDR_ALIASES,
    )

    cdpos, _ = _load_change_input(
        context,
        CDPOS_INPUT_PREFIX,
        CDPOS_ALIASES,
    )

    _, eligible, _ = build_release_events(
        cdhdr,
        cdpos,
    )

    po_headers = (
        po[
            PO_KEY
            + [
                "PO Creator ID",
            ]
        ]
        .drop_duplicates(
            PO_KEY
        )
    )

    approvals, _ = _select_lha_approval_event(
        eligible,
        po_headers,
    )

    result = _output(
        po,
        approvals,
        _aggregate_gr(gr),
        days,
    )

    output_file = write_control_sheet(
        context=context,
        sheet_name=SHEET_NAME,
        dataframe=result,
        date_columns=[
            "PO Doc Date",
            "PO Approval Date",
            "USD Rate Date",
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
            "Days GR After Approval",
            "Threshold Days",
        ],
    )

    exception_pos = (
        result[PO_KEY]
        .drop_duplicates()
        .shape[0]
        if not result.empty
        else 0
    )

    status = (
        "ERROR"
        if exception_pos
        else "OK"
    )

    summary = {
        "Control": CONTROL_ID,
        "Control Result": status,
        "Threshold Days": days,
        "PO Lines Rows Read": po_metrics["rows_read"],
        "GR Rows Read": gr_metrics["rows_read"],
        "GR Rows Used": gr_metrics["gr_rows"],
        "Exception Rows": len(result),
        "Exception POs": exception_pos,
    }

    print(
        f"\n{CONTROL_ID} - {CONTROL_NAME}"
    )

    for label, value in summary.items():
        print(
            f"{label}: {value}"
        )

    print(
        f"Output: {output_file} [{SHEET_NAME}]\n"
    )

    return {
        "status": status,
        "output_file": output_file,
        "sheet_name": SHEET_NAME,
        "rows": len(result),
        "summary": summary,
    }


run = run_po_005
