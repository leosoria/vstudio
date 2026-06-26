"""
CD_003 - Duplicate Cash Disbursements.

Analysis:
- Module: Cash Disbursements
- Analysis Code: CD_03
- Analysis Title: Duplicate Cash Disbursements

Description:
Identifies potential duplicate cash disbursements.

Procedure:
Detect potential duplicate cash disbursements.

Analytic logic:
Flag payments with the same vendor, payment currency and payment amount that
occur on the same date or within a configurable duplicate review window.
Reversals are excluded by reusing the CD_001 payment universe, which keeps
payment rows only and derives positive payment amounts.

Design notes:
- CD_003 uses the same payment universe as CD_001 by reusing the CD_001
  detail dataframe builder.
- It does not read the CD01 sheet from the output workbook.
- It writes/replaces only the CD03 sheet.
- It does not delete sheets from other controls.
"""

import pandas as pd

from core.cd_common import (
    apply_standard_cd_formatting,
    build_key,
    get_cd_output_file,
    load_cd_base_data,
    normalize_text,
    open_or_create_cd_output_workbook,
    recreate_cd_sheet,
    save_cd_output_workbook,
    write_dataframe_to_sheet,
)
from modules.CD.cd_001 import build_cd_001_dataframe


SHEET_NAME = "CD03"
DUP_WINDOW_DAYS = 30

OUTPUT_COLUMNS = [
    "CoCo",
    "Company",
    "KEY_DOC",
    "Payment DocEntry",
    "Payment DocNum",
    "Journal TransId",
    "Vendor Code",
    "Vendor Name",
    "KEY_VENDOR",
    "Payment Date",
    "Tax Date",
    "Payment Currency",
    "Company Main Currency",
    "Payment Amount",
    "Payment Amount USD",
    "Counter Reference",
    "Canceled",
    "Transfer Amount",
    "Cash Amount",
    "Check Amount",
    "Credit Card Amount",
    "DUP_PAYMENT_KEY",
    "Dup Window Days",
    "Nearest Duplicate Days",
]

DATE_COLUMNS = {
    "Payment Date",
    "Tax Date",
}

AMOUNT_COLUMNS = {
    "Payment Amount",
    "Payment Amount USD",
    "Transfer Amount",
    "Cash Amount",
    "Check Amount",
    "Credit Card Amount",
}

INTEGER_COLUMNS = {
    "Payment DocEntry",
    "Payment DocNum",
    "Journal TransId",
    "Dup Window Days",
    "Nearest Duplicate Days",
}


def format_amount_for_key(value):
    """
    Format an amount for duplicate matching keys.

    The duplicate match itself uses a rounded numeric amount. This display key
    removes unnecessary trailing zeros so values like 1168.8600 are shown as
    1168.86 and values like 108175.00 are shown as 108175.
    """
    amount = pd.to_numeric(value, errors="coerce")

    if pd.isna(amount):
        return ""

    amount_text = f"{float(amount):.2f}".rstrip("0").rstrip(".")

    return amount_text


def build_duplicate_payment_key(row):
    """
    Build the CD_003 duplicate key used for review and grouping.
    """
    return build_key(
        row.get("Company", ""),
        row.get("Vendor Code", ""),
        row.get("Payment Currency", ""),
        row.get("Payment Amount Key", ""),
    )


def get_dup_window_days(context):
    """
    Return duplicate review window days for CD_003.

    Priority:
    1. Control PARAM1 from the CD sheet in config.xlsx.
    2. Default DUP_WINDOW_DAYS when PARAM1 is blank.
    """
    control_config = context.get("control", {})
    param1_value = normalize_text(control_config.get("param1", ""))

    if param1_value == "":
        return DUP_WINDOW_DAYS

    try:
        window_days = int(float(param1_value))
    except ValueError as error:
        raise ValueError(
            "Invalid CD_003 PARAM1. "
            "Expected a positive number of days for duplicate review window. "
            f"Received: {param1_value}"
        ) from error

    if window_days <= 0:
        raise ValueError(
            "Invalid CD_003 PARAM1. "
            "Duplicate review window days must be greater than zero. "
            f"Received: {param1_value}"
        )

    return window_days


def get_nearest_duplicate_days(payment_dates, window_days):
    """
    Return the nearest duplicate distance in days for each payment date.

    A row receives a value only when another payment in the same duplicate group
    exists on the same date or within the configured review window.
    """
    result = pd.Series(pd.NA, index=payment_dates.index, dtype="object")
    sorted_dates = payment_dates.sort_values()

    for position, (row_index, payment_date) in enumerate(sorted_dates.items()):
        if pd.isna(payment_date):
            continue

        nearest_days = None

        for other_position, other_date in enumerate(sorted_dates):
            if other_position == position:
                continue

            if pd.isna(other_date):
                continue

            days_difference = abs((payment_date - other_date).days)

            if days_difference > window_days:
                continue

            if nearest_days is None or days_difference < nearest_days:
                nearest_days = days_difference

        if nearest_days is not None:
            result.loc[row_index] = nearest_days

    return result


def build_cd_003_dataframe(detail_dataframe, window_days=DUP_WINDOW_DAYS):
    """
    Build CD_003 duplicate disbursements dataframe from CD_001 detail rows.

    Duplicate rules:
    - Start from the same CD_001 payment universe.
    - Exclude rows flagged as canceled when a canceled flag exists.
    - Match by Company, Vendor Code, Payment Currency and rounded Payment Amount.
    - Flag rows that have another payment in the same match group on the same
      date or within window_days.
    - Exclude repeated lines of the same KEY_DOC from creating a duplicate by
      de-duplicating at payment document level for the matching step.
    """
    if detail_dataframe.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    working_dataframe = detail_dataframe.copy()

    working_dataframe["Payment Date"] = pd.to_datetime(
        working_dataframe["Payment Date"],
        errors="coerce",
    )

    working_dataframe["Payment Amount"] = pd.to_numeric(
        working_dataframe["Payment Amount"],
        errors="coerce",
    )

    working_dataframe["Canceled"] = (
        working_dataframe["Canceled"]
        .fillna("N")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    working_dataframe = working_dataframe[working_dataframe["Canceled"] != "Y"].copy()

    working_dataframe["Payment Amount Rounded"] = working_dataframe[
        "Payment Amount"
    ].round(2)
    working_dataframe["Payment Amount Key"] = working_dataframe[
        "Payment Amount Rounded"
    ].apply(format_amount_for_key)

    working_dataframe["Duplicate Group Key"] = [
        build_key(company, vendor_code, payment_currency, amount_key)
        for company, vendor_code, payment_currency, amount_key in zip(
            working_dataframe["Company"].apply(normalize_text),
            working_dataframe["Vendor Code"].apply(normalize_text),
            working_dataframe["Payment Currency"].apply(normalize_text),
            working_dataframe["Payment Amount Key"].apply(normalize_text),
        )
    ]

    matching_dataframe = working_dataframe.drop_duplicates(
        subset=["Duplicate Group Key", "KEY_DOC"],
        keep="first",
    ).copy()

    matching_dataframe["Nearest Duplicate Days"] = pd.NA

    for _, group in matching_dataframe.groupby("Duplicate Group Key", dropna=False):
        if len(group) < 2:
            continue

        nearest_duplicate_days = get_nearest_duplicate_days(
            payment_dates=group["Payment Date"],
            window_days=window_days,
        )

        matching_dataframe.loc[
            nearest_duplicate_days.dropna().index,
            "Nearest Duplicate Days",
        ] = nearest_duplicate_days.dropna()

    matching_dataframe["Is Duplicate Candidate"] = matching_dataframe[
        "Nearest Duplicate Days"
    ].notna()

    nearest_days_by_key_doc = (
        matching_dataframe.dropna(subset=["Nearest Duplicate Days"])
        .set_index("KEY_DOC")["Nearest Duplicate Days"]
        .to_dict()
    )

    duplicate_key_docs = set(
        matching_dataframe.loc[
            matching_dataframe["Is Duplicate Candidate"],
            "KEY_DOC",
        ]
    )

    output_dataframe = working_dataframe[
        working_dataframe["KEY_DOC"].isin(duplicate_key_docs)
    ].copy()

    if output_dataframe.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    output_dataframe["DUP_PAYMENT_KEY"] = output_dataframe.apply(
        build_duplicate_payment_key,
        axis=1,
    )
    output_dataframe["Dup Window Days"] = window_days
    output_dataframe["Nearest Duplicate Days"] = output_dataframe["KEY_DOC"].map(
        nearest_days_by_key_doc
    )

    output_dataframe = output_dataframe[OUTPUT_COLUMNS]

    output_dataframe = output_dataframe.sort_values(
        by=[
            "Company",
            "Vendor Code",
            "Payment Currency",
            "Payment Amount",
            "Payment Date",
            "KEY_DOC",
        ],
        ascending=[True, True, True, True, True, True],
        na_position="last",
    ).reset_index(drop=True)

    return output_dataframe


def print_cd_003_summary(detail_dataframe, output_dataframe, output_file, window_days):
    """
    Print CD_003 validation summary.
    """
    print("CD_003 validation summary")
    print("-------------------------")
    print(f"CD_003 source/detail rows: {len(detail_dataframe)}")
    print(f"CD_003 duplicate candidate rows: {len(output_dataframe)}")
    print(f"CD_003 duplicate keys: {output_dataframe['DUP_PAYMENT_KEY'].nunique()}")
    print(f"Duplicate window days: {window_days}")
    print(f"Output file: {output_file}")
    print(f"Output sheet: {SHEET_NAME}")
    print()


def run_cd_003(context):
    """
    Run CD_003.
    """
    window_days = get_dup_window_days(context)

    source_dataframe = load_cd_base_data(context)

    detail_dataframe = build_cd_001_dataframe(
        source_dataframe=source_dataframe,
        context=context,
    )

    output_dataframe = build_cd_003_dataframe(
        detail_dataframe=detail_dataframe,
        window_days=window_days,
    )

    if len(output_dataframe) == 0:
        print()
        print("WARNING: CD_003 generated 0 rows.")
        print("Possible causes:")
        print(
            f"- No duplicate same vendor/currency/amount payments were found "
            f"within {window_days} days."
        )
        print("- No rows with D/C = S were found.")
        print("- COMPANIES filter does not match Empr values.")
        print("- The input file is empty or the wrong sheet was read.")
        print("- The module PARAM1 matched the wrong input file.")
        print()

    output_file = get_cd_output_file(context)

    workbook = open_or_create_cd_output_workbook(output_file)

    worksheet = recreate_cd_sheet(
        workbook=workbook,
        sheet_name=SHEET_NAME,
    )

    write_dataframe_to_sheet(
        worksheet=worksheet,
        dataframe=output_dataframe,
    )

    apply_standard_cd_formatting(
        worksheet=worksheet,
        dataframe=output_dataframe,
        date_columns=DATE_COLUMNS,
        amount_columns=AMOUNT_COLUMNS,
        integer_columns=INTEGER_COLUMNS,
    )

    save_cd_output_workbook(
        workbook=workbook,
        output_file=output_file,
    )

    print_cd_003_summary(
        detail_dataframe=detail_dataframe,
        output_dataframe=output_dataframe,
        output_file=output_file,
        window_days=window_days,
    )
