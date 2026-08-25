"""PO06 - diferencias de precio para el mismo proveedor/material."""

import pandas as pd

from core.po_common import load_po_lines, normalize_text, write_control_sheet


CONTROL_ID = "PO_006"
SHEET_NAME = "PO06"

REQUIRED_FIELDS = (
    "Company",
    "PO Number",
    "PO Line",
    "Vendor Code",
    "Item Code",
    "PO Unit Price",
    "PO Price Unit",
    "PO Doc Currency",
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
    "Min Unit Price",
    "Max Unit Price",
    "Price Difference",
]

def _threshold(context):
    """PARAM1 vacío = cualquier diferencia; acepta coma decimal."""
    value = normalize_text(context.get("control", {}).get("param1", ""))
    return 0.0 if not value else float(value.replace(",", "."))


def _build_output(exceptions):
    """Entrega PO06 con el mismo orden de columnas que LHA."""
    output = exceptions.copy()

    # En ECC, CoCd es el código de compañía.
    if "CoCo" not in output.columns:
        output["CoCo"] = output.get("Company", "")

    # PO Month no viene directamente en la bajada.
    if "PO Month" not in output.columns:
        po_date = pd.to_datetime(
            output.get("PO Doc Date"),
            errors="coerce",
            dayfirst=True,
        )
        output["PO Month"] = po_date.dt.strftime("%Y-%m").fillna("")

    # En SAP ECC tenemos PR Number, no el DocEntry interno de SAP B1.
    if "PR DocEntry" not in output.columns:
        output["PR DocEntry"] = output.get("PR Number", "")

    if "From PR" not in output.columns:
        pr_number = output.get(
            "PR Number",
            pd.Series("", index=output.index, dtype="object"),
        ).map(normalize_text)

        output["From PR"] = pr_number.ne("").map(
            {True: "Y", False: "N"}
        )

    # Crea vacías las columnas LHA que la bajada LBR no contiene.
    for column in OUTPUT_COLUMNS:
        if column not in output.columns:
            output[column] = ""

    return output.loc[:, OUTPUT_COLUMNS]


def run_po_006(context):
    po, metrics = load_po_lines(context, required_fields=REQUIRED_FIELDS)

    for column in ("Vendor Code", "Item Code", "PO Doc Currency"):
        po[column] = po[column].map(normalize_text)

    po["PO Unit Price"] = pd.to_numeric(po["PO Unit Price"], errors="coerce")
    po["PO Price Unit"] = pd.to_numeric(po["PO Price Unit"], errors="coerce")

    # SAP NETPR está expresado por PEINH unidades.
    po["Comparable Unit Price"] = (
        po["PO Unit Price"] / po["PO Price Unit"]
    ).round(6)

    keys = ["Company", "Vendor Code", "Item Code", "PO Doc Currency"]
    eligible = po[
        po["Vendor Code"].ne("")
        & po["Item Code"].ne("")
        & po["PO Doc Currency"].ne("")
        & po["Comparable Unit Price"].notna()
        & po["PO Price Unit"].gt(0)
    ].copy()

    grouped = eligible.groupby(keys, dropna=False)["Comparable Unit Price"]
    eligible["Min Unit Price"] = grouped.transform("min")
    eligible["Max Unit Price"] = grouped.transform("max")
    eligible["Price Difference"] = (
        eligible["Max Unit Price"] - eligible["Min Unit Price"]
    ).round(6)

    threshold = _threshold(context)
    exceptions = eligible[
        eligible["Price Difference"].gt(threshold)
    ].copy()

    exceptions["PRICE_DIFF_KEY"] = (
        exceptions[keys].astype(str).agg("|".join, axis=1)
    )
    exceptions["Threshold"] = threshold

    exceptions = exceptions.sort_values(
        keys + ["Comparable Unit Price", "PO Number", "PO Line"],
        kind="stable",
    ).reset_index(drop=True)

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
            "Min Unit Price",
            "Max Unit Price",
            "Price Difference",
        ],
    )

    print(
        f"{CONTROL_ID}: {len(exceptions)} líneas con diferencia de precio "
        f"(umbral={threshold}; población={metrics['rows_after_config_filters']})."
    )

    return {
        "status": "ERROR" if not exceptions.empty else "OK",
        "output_file": output_file,
        "sheet_name": SHEET_NAME,
        "rows": len(exceptions),
    }
