"""PO11 - líneas de PO sin Solicitud de Compra."""

from time import perf_counter

import pandas as pd

from core.gl_common import (
    load_gl_fx_rates_data,
    normalize_fx_rates,
    select_fx_rate_to_usd,
)
from core.po_common import (
    load_po_lines,
    normalize_text,
    write_control_sheet,
)


CONTROL_ID = "PO_011"
CONTROL_NAME = "Purchase Orders Without Purchase Requisition"
SHEET_NAME = "PO11"

PO11_REQUIRED_FIELDS = (
    "Company",
    "PO Number",
    "PO Line",
    "PO Doc Date",
    "PR Number",
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


def _elapsed(start):
    return perf_counter() - start


def _measure_key_duplicates(dataframe):
    """
    Mide claves repetidas sin excluirlas ni detener PO11.

    PO11 conserva todas las filas de la bajada para mantener la misma
    población y granularidad efectiva que LHA.
    """
    key = ["Company", "PO Number", "PO Line"]
    duplicated = dataframe.duplicated(key, keep=False)

    return {
        "duplicate_key_rows": int(duplicated.sum()),
        "duplicate_keys": int(
            dataframe.loc[duplicated, key]
            .drop_duplicates()
            .shape[0]
        ),
    }


def _add_document_date_fx(dataframe, fx_dataframe):
    """
    Convierte una vez por moneda + fecha de documento.

    select_fx_rate_to_usd conserva la prioridad de rate type, búsqueda de
    fecha anterior, tasa directa/inversa, factores y tratamiento de USD
    definidos por la infraestructura GL.
    """
    result = dataframe.copy()

    result["PO Line Total"] = pd.to_numeric(
        result.get("PO Line Total"),
        errors="coerce",
    )
    result["PO Doc Currency"] = (
        result.get(
            "PO Doc Currency",
            pd.Series("", index=result.index, dtype="object"),
        )
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )
    result["PO Doc Date"] = pd.to_datetime(
        result["PO Doc Date"],
        errors="coerce",
        dayfirst=True,
    ).dt.normalize()

    normalized_fx = normalize_fx_rates(fx_dataframe)

    fx_keys = (
        result[["PO Doc Currency", "PO Doc Date"]]
        .drop_duplicates()
        .reset_index(drop=True)
    )

    fx_rows = []

    # Loop solo por claves FX únicas, nunca por línea de PO.
    for currency, document_date in fx_keys.itertuples(index=False, name=None):
        details = select_fx_rate_to_usd(
            normalized_fx_dataframe=normalized_fx,
            currency=currency,
            requested_date=document_date,
        )

        fx_rows.append(
            {
                "PO Doc Currency": currency,
                "PO Doc Date": document_date,
                "_FX_TO_USD": (
                    details["fx_to_usd"] if details is not None else pd.NA
                ),
                "USD Rate": (
                    details["usd_rate"] if details is not None else pd.NA
                ),
                "USD Rate Date": (
                    details["rate_date"] if details is not None else pd.NaT
                ),
            }
        )

    fx_lookup = pd.DataFrame(
        fx_rows,
        columns=[
            "PO Doc Currency",
            "PO Doc Date",
            "_FX_TO_USD",
            "USD Rate",
            "USD Rate Date",
        ],
    )

    result = result.merge(
        fx_lookup,
        on=["PO Doc Currency", "PO Doc Date"],
        how="left",
        validate="many_to_one",
    )

    result["PO Line Total USD"] = (
        result["PO Line Total"]
        * pd.to_numeric(result["_FX_TO_USD"], errors="coerce")
    )

    return result.drop(columns="_FX_TO_USD")


def _build_output(exceptions):
    """Entrega exactamente las columnas LHA, en el mismo orden."""
    details = exceptions.copy()

    details["CoCo"] = details["Company"]
    details["From PR"] = "N"
    details["PO Month"] = (
        pd.to_datetime(
            details["PO Doc Date"],
            errors="coerce",
            dayfirst=True,
        )
        .dt.strftime("%Y-%m")
        .fillna("")
    )

    for column in OUTPUT_COLUMNS:
        if column not in details.columns:
            details[column] = ""

    return details.loc[:, OUTPUT_COLUMNS].reset_index(drop=True)


def run_po_011(context):
    timings = {}

    started = perf_counter()
    po_lines, input_metrics = load_po_lines(
        context,
        required_fields=PO11_REQUIRED_FIELDS,
    )
    timings["load_po_lines"] = _elapsed(started)

    started = perf_counter()

    key_metrics = _measure_key_duplicates(po_lines)

    pr_number = po_lines["PR Number"].map(normalize_text)
    from_pr = pr_number.ne("")

    prepared = po_lines.copy()
    prepared["From PR"] = from_pr.map({True: "Y", False: "N"})

    timings["prepare_and_validate"] = _elapsed(started)

    started = perf_counter()
    exceptions = prepared.loc[~from_pr].copy()
    timings["analytic"] = _elapsed(started)

    started = perf_counter()
    fx_dataframe = load_gl_fx_rates_data(context)
    exceptions = _add_document_date_fx(exceptions, fx_dataframe)
    timings["fx"] = _elapsed(started)

    started = perf_counter()
    output = _build_output(exceptions)
    output_file = write_control_sheet(
        context=context,
        sheet_name=SHEET_NAME,
        dataframe=output,
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
    timings["write_workbook"] = _elapsed(started)

    exception_pos = (
        exceptions[["Company", "PO Number"]]
        .drop_duplicates()
        .shape[0]
    )

    print()
    print(f"{CONTROL_ID} - {CONTROL_NAME}")
    print("-" * (len(CONTROL_ID) + len(CONTROL_NAME) + 3))
    print(f"Input file: {input_metrics['input_file']}")
    print(f"Rows read: {input_metrics['rows_read']}")
    print(f"Residual rows: {input_metrics['residual_rows']}")
    print(
        "Rows after CONFIG filters: "
        f"{input_metrics['rows_after_config_filters']}"
    )
    print(
    "Rows belonging to duplicated PO-line keys: "
    f"{key_metrics['duplicate_key_rows']}"
    )
    print(
        "Distinct duplicated PO-line keys: "
        f"{key_metrics['duplicate_keys']}"
    )

    if key_metrics["duplicate_keys"]:
        print(
            "WARNING: duplicated Company + PO Number + PO Line keys "
            "were retained; PO11 does not silently remove source rows."
        )
    print(f"Lines with PR: {int(from_pr.sum())}")
    print(f"Lines without PR: {len(exceptions)}")
    print(f"POs without PR: {exception_pos}")
    print(
        "Timings (seconds): "
        + ", ".join(
            f"{name}={seconds:.3f}"
            for name, seconds in timings.items()
        )
    )
    print(f"PO11 output file: {output_file}")
    print(f"PO11 output sheet: {SHEET_NAME}")
    print()

    return {
    "status": "ERROR" if not exceptions.empty else "OK",
    "output_file": output_file,
    "sheet_name": SHEET_NAME,
    "rows": len(exceptions),
    "exception_pos": exception_pos,
    "duplicate_key_rows": key_metrics["duplicate_key_rows"],
    "duplicate_keys": key_metrics["duplicate_keys"],
    "timings": timings,
    }
