"""PO02 - potential duplicate purchase orders.

The control identifies lines belonging to different purchase orders with the
same company, vendor, material and order quantity.  The comparison covers the
configured period and deliberately remains independent from PO01. USD output
uses the closing FX rate requested at the CONFIG TO date.
"""

import pandas as pd

from core.gl_common import (
    load_gl_fx_rates_data,
    normalize_fx_rates,
    select_fx_rate_to_usd,
)
from core.po_common import (
    PO02_REQUIRED_FIELDS,
    load_po_lines,
    normalize_text,
    parse_config_date,
    write_control_sheet,
)


CONTROL_ID = "PO_002"
CONTROL_NAME = "Duplicate Purchase Orders"
SHEET_NAME = "PO02"

GROUP_FIELDS = [
    "Company",
    "Vendor Code",
    "Item Code",
    "PO Quantity",
]

LHA_OUTPUT_COLUMNS = [
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
    "DUP_PO_KEY",
]


def prepare_duplicate_population(dataframe):
    """Normalize PO02 fields and separate rows that cannot be compared."""
    prepared = dataframe.copy()

    prepared["Vendor Code"] = prepared["Vendor Code"].map(normalize_text)
    prepared["Item Code"] = prepared["Item Code"].map(normalize_text)

    raw_quantity = prepared["PO Quantity"]
    quantity_text = raw_quantity.map(normalize_text)
    quantity = pd.to_numeric(raw_quantity, errors="coerce")

    missing_vendor = prepared["Vendor Code"].eq("")
    missing_material = prepared["Item Code"].eq("")
    missing_quantity = quantity_text.eq("")
    invalid_quantity = quantity_text.ne("") & quantity.isna()

    prepared["PO Quantity"] = quantity

    eligible_mask = ~(
        missing_vendor
        | missing_material
        | missing_quantity
        | invalid_quantity
    )

    eligible = prepared.loc[eligible_mask].copy()

    metrics = {
        "rows_without_vendor": int(missing_vendor.sum()),
        "rows_without_material": int(missing_material.sum()),
        "rows_without_quantity": int(missing_quantity.sum()),
        "rows_with_invalid_quantity": int(invalid_quantity.sum()),
        "rows_eligible": len(eligible),
        "rows_excluded": int((~eligible_mask).sum()),
    }

    return eligible, metrics


def build_po02_exceptions(eligible):
    """Return details for keys shared by more than one distinct PO."""
    if eligible.empty:
        return eligible.copy()

    po_counts = (
        eligible.groupby(
            GROUP_FIELDS,
            dropna=False,
        )["PO Number"]
        .transform("nunique")
    )

    exceptions = eligible.loc[po_counts.gt(1)].copy()

    if exceptions.empty:
        return exceptions

    exceptions["Duplicate PO Count"] = po_counts.loc[exceptions.index].astype(int)
    exceptions["DUP_PO_KEY"] = (
        exceptions[GROUP_FIELDS]
        .astype(str)
        .agg("|".join, axis=1)
    )

    sort_columns = GROUP_FIELDS + [
        "PO Number",
        "PO Line",
    ]

    return exceptions.sort_values(
        sort_columns,
        kind="stable",
    ).reset_index(drop=True)


def count_duplicate_groups(exceptions):
    """Count distinct duplicate keys without relying on display strings."""
    if exceptions.empty:
        return 0

    return len(
        exceptions[GROUP_FIELDS]
        .drop_duplicates()
    )


def add_period_end_usd_fields(exceptions, fx_dataframe, module_config):
    """Convert PO line totals to USD using the CONFIG TO closing rate."""
    result = exceptions.copy()
    result["PO Line Total USD"] = pd.NA
    result["USD Rate"] = pd.NA
    result["USD Rate Date"] = pd.NaT

    if result.empty or fx_dataframe.empty:
        return result

    rate_date = parse_config_date(
        module_config.get("to", ""),
        "TO",
    )
    normalized_fx = normalize_fx_rates(fx_dataframe)
    rate_cache = {}
    amount_usd_values = []
    usd_rate_values = []
    usd_rate_date_values = []

    for _, row in result.iterrows():
        currency = normalize_text(
            row.get("PO Doc Currency", "")
        ).upper()

        if currency not in rate_cache:
            rate_cache[currency] = select_fx_rate_to_usd(
                normalized_fx_dataframe=normalized_fx,
                currency=currency,
                requested_date=rate_date,
            )

        fx_details = rate_cache[currency]
        line_total = pd.to_numeric(
            pd.Series([row.get("PO Line Total", pd.NA)]),
            errors="coerce",
        ).iloc[0]

        if fx_details is None:
            amount_usd_values.append(pd.NA)
            usd_rate_values.append(pd.NA)
            usd_rate_date_values.append(pd.NaT)
            continue

        amount_usd_values.append(
            pd.NA
            if pd.isna(line_total)
            else line_total * fx_details["fx_to_usd"]
        )
        usd_rate_values.append(fx_details["usd_rate"])
        usd_rate_date_values.append(fx_details["rate_date"])

    result["PO Line Total USD"] = amount_usd_values
    result["USD Rate"] = usd_rate_values
    result["USD Rate Date"] = usd_rate_date_values

    return result


def build_summary_row(input_metrics, logic_metrics, exceptions, distinct_pos):
    """Build the auditable PO02 summary record."""
    exception_pos = (
        exceptions["PO Number"].nunique()
        if not exceptions.empty
        else 0
    )
    duplicate_groups = count_duplicate_groups(exceptions)
    control_result = "ERROR" if not exceptions.empty else "OK"
    message = (
        f"Potential duplicate POs found: {duplicate_groups} groups, "
        f"{exception_pos} POs and {len(exceptions)} detail rows."
        if not exceptions.empty
        else "No potential duplicate POs found in the eligible population."
    )

    return {
        "Record Type": "SUMMARY",
        "Control": CONTROL_ID,
        "Control Name": CONTROL_NAME,
        "Control Result": control_result,
        "Message": message,
        "Rows Read": input_metrics["rows_read"],
        "Residual Rows": input_metrics["residual_rows"],
        "Rows After Config Filters": input_metrics["rows_after_config_filters"],
        "Distinct POs": distinct_pos,
        "Rows Without Vendor": logic_metrics["rows_without_vendor"],
        "Rows Without Material": logic_metrics["rows_without_material"],
        "Rows Without Quantity": logic_metrics["rows_without_quantity"],
        "Rows With Invalid Quantity": logic_metrics["rows_with_invalid_quantity"],
        "Rows Eligible For Duplicate Logic": logic_metrics["rows_eligible"],
        "Rows Excluded By Duplicate Logic": logic_metrics["rows_excluded"],
        "Duplicate Groups": duplicate_groups,
        "Exception Rows": len(exceptions),
        "Exception POs": exception_pos,
    }


def build_lha_output(exceptions):
    """Return exception detail with the exact LHA PO column contract."""
    details = exceptions.copy()

    if "CoCo" not in details.columns:
        details["CoCo"] = details.get("Company", "")

    if "PO Month" not in details.columns:
        document_dates = pd.to_datetime(
            details.get("PO Doc Date"),
            errors="coerce",
            dayfirst=True,
        )
        details["PO Month"] = document_dates.dt.strftime("%Y-%m").fillna("")

    if "From PR" not in details.columns:
        pr_number = details.get(
            "PR Number",
            pd.Series("", index=details.index, dtype="object"),
        ).map(normalize_text)
        details["From PR"] = pr_number.ne("").map({True: "Y", False: "N"})

    for column in LHA_OUTPUT_COLUMNS:
        if column not in details.columns:
            details[column] = ""

    return details.loc[:, LHA_OUTPUT_COLUMNS]


def print_po02_metrics(summary_row, input_metrics):
    """Print population and exception diagnostics to the run log."""
    print()
    print(f"{CONTROL_ID} - {CONTROL_NAME}")
    print("-" * (len(CONTROL_ID) + len(CONTROL_NAME) + 3))
    print(f"Input file: {input_metrics['input_file']}")
    print(f"Input sheet: {input_metrics['input_sheet']}")
    print(f"Rows read: {summary_row['Rows Read']}")
    print(f"Residual rows: {summary_row['Residual Rows']}")
    print(f"Rows after CONFIG filters: {summary_row['Rows After Config Filters']}")
    print(f"Rows without vendor: {summary_row['Rows Without Vendor']}")
    print(f"Rows without material: {summary_row['Rows Without Material']}")
    print(f"Rows without quantity: {summary_row['Rows Without Quantity']}")
    print(f"Rows with invalid quantity: {summary_row['Rows With Invalid Quantity']}")
    print(f"Rows eligible for logic: {summary_row['Rows Eligible For Duplicate Logic']}")
    print(f"Duplicate groups: {summary_row['Duplicate Groups']}")
    print(f"Exception rows: {summary_row['Exception Rows']}")
    print(f"Exception POs: {summary_row['Exception POs']}")
    print(f"Control result: {summary_row['Control Result']}")
    print(f"Message: {summary_row['Message']}")
    print()


def run_po_002(context):
    """Execute PO02 independently and replace only the PO02 output sheet."""
    po_lines, input_metrics = load_po_lines(
        context,
        required_fields=PO02_REQUIRED_FIELDS,
    )
    eligible, logic_metrics = prepare_duplicate_population(po_lines)
    exceptions = build_po02_exceptions(eligible)
    fx_dataframe = load_gl_fx_rates_data(context)
    exceptions = add_period_end_usd_fields(
        exceptions,
        fx_dataframe,
        context["module"],
    )

    summary_row = build_summary_row(
        input_metrics,
        logic_metrics,
        exceptions,
        int(po_lines["PO Number"].nunique()),
    )

    output = build_lha_output(exceptions)
    output_file = write_control_sheet(
        context=context,
        sheet_name=SHEET_NAME,
        dataframe=output,
        date_columns=["PO Doc Date", "USD Rate Date"],
        amount_columns=[
            "PO Quantity",
            "PO Unit Price",
            "PO Line Total",
            "PO Line Total USD",
            "USD Rate",
        ],
    )

    print_po02_metrics(summary_row, input_metrics)
    print(f"PO02 output file: {output_file}")
    print(f"PO02 output sheet: {SHEET_NAME}")

    return {
        "status": summary_row["Control Result"],
        "output_file": output_file,
        "sheet_name": SHEET_NAME,
        "rows": len(exceptions),
        "summary": summary_row,
    }
