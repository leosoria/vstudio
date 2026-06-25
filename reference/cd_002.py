"""
CD_002 - Summary Of Cash Disbursements By Vendor.

Analysis:
- Module: Cash Disbursements
- Analysis Code: CD_02
- Analysis Title: Summary Of Cash Disbursements By Vendor

Description:
Aggregates / summarizes disbursements by vendor.

Procedure:
Summarize and trend disbursements by vendor.

Analytic logic:
Aggregate payments by vendor, amount/count over period.

Context:
Payment visibility for vendor oversight.

Design notes:
- CD_002 uses the same payment universe as CD_001 by reusing the CD_001
  detail dataframe builder.
- It does not read the CD01 sheet from the output workbook.
- It writes/replaces only the CD02 sheet.
- It does not delete sheets from other controls.
"""

import pandas as pd

from core.cd_common import (
    apply_standard_cd_formatting,
    get_cd_output_file,
    load_cd_base_data,
    open_or_create_cd_output_workbook,
    recreate_cd_sheet,
    save_cd_output_workbook,
    write_dataframe_to_sheet,
)
from modules.CD.cd_001 import build_cd_001_dataframe


SHEET_NAME = "CD02"

OUTPUT_COLUMNS = [
    "Company",
    "Vendor Code",
    "Vendor Name",
    "Qty Payments",
    "Total Amount USD",
    "First Payment",
    "Last Payment",
]

DATE_COLUMNS = {
    "First Payment",
    "Last Payment",
}

AMOUNT_COLUMNS = {
    "Total Amount USD",
}

INTEGER_COLUMNS = {
    "Qty Payments",
}


def build_cd_002_dataframe(detail_dataframe):
    """
    Build CD_002 vendor summary dataframe from the CD payment detail dataframe.

    The detail dataframe is expected to follow CD_001's payment-style output.

    Aggregation rules:
    - Group by Company, Vendor Code and Vendor Name.
    - Qty Payments counts unique KEY_DOC values to avoid overcounting payments
      that have more than one line.
    - Total Amount USD sums Payment Amount USD.
    - First Payment is the earliest Payment Date.
    - Last Payment is the latest Payment Date.

    If a future regional expectation requires line counts instead of unique
    payment documents, change Qty Payments from nunique(KEY_DOC) to count rows.
    """
    if detail_dataframe.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    working_dataframe = detail_dataframe.copy()

    working_dataframe["Payment Date"] = pd.to_datetime(
        working_dataframe["Payment Date"],
        errors="coerce",
    )

    working_dataframe["Payment Amount USD"] = pd.to_numeric(
        working_dataframe["Payment Amount USD"],
        errors="coerce",
    )

    summary_dataframe = (
        working_dataframe
        .groupby(
            [
                "Company",
                "Vendor Code",
                "Vendor Name",
            ],
            dropna=False,
        )
        .agg(
            **{
                "Qty Payments": ("KEY_DOC", "nunique"),
                "Total Amount USD": ("Payment Amount USD", "sum"),
                "First Payment": ("Payment Date", "min"),
                "Last Payment": ("Payment Date", "max"),
            }
        )
        .reset_index()
    )

    summary_dataframe = summary_dataframe[OUTPUT_COLUMNS]

    summary_dataframe = summary_dataframe.sort_values(
        by=[
            "Company",
            "Total Amount USD",
            "Vendor Code",
            "Vendor Name",
        ],
        ascending=[True, False, True, True],
        na_position="last",
    ).reset_index(drop=True)

    return summary_dataframe


def print_cd_002_summary(detail_dataframe, summary_dataframe, output_file):
    """
    Print CD_002 validation summary.
    """
    total_amount_usd = summary_dataframe["Total Amount USD"].sum()

    first_payment = pd.to_datetime(
        summary_dataframe["First Payment"],
        errors="coerce",
    ).min()

    last_payment = pd.to_datetime(
        summary_dataframe["Last Payment"],
        errors="coerce",
    ).max()

    if pd.isna(first_payment):
        first_payment_text = ""
    else:
        first_payment_text = first_payment.strftime("%d/%m/%Y")

    if pd.isna(last_payment):
        last_payment_text = ""
    else:
        last_payment_text = last_payment.strftime("%d/%m/%Y")

    print("CD_002 validation summary")
    print("-------------------------")
    print(f"CD_002 source/detail rows: {len(detail_dataframe)}")
    print(f"CD_002 vendors summarized: {len(summary_dataframe)}")
    print(f"Total Amount USD: {total_amount_usd:,.2f}")
    print(f"First Payment: {first_payment_text}")
    print(f"Last Payment: {last_payment_text}")
    print(f"Output file: {output_file}")
    print(f"Output sheet: {SHEET_NAME}")
    print()


def run_cd_002(context):
    """
    Run CD_002.
    """
    source_dataframe = load_cd_base_data(context)

    detail_dataframe = build_cd_001_dataframe(
        source_dataframe=source_dataframe,
        context=context,
    )

    summary_dataframe = build_cd_002_dataframe(
        detail_dataframe=detail_dataframe,
    )

    if len(summary_dataframe) == 0:
        print()
        print("WARNING: CD_002 generated 0 rows.")
        print("Possible causes:")
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
        dataframe=summary_dataframe,
    )

    apply_standard_cd_formatting(
        worksheet=worksheet,
        dataframe=summary_dataframe,
        date_columns=DATE_COLUMNS,
        amount_columns=AMOUNT_COLUMNS,
        integer_columns=INTEGER_COLUMNS,
    )

    save_cd_output_workbook(
        workbook=workbook,
        output_file=output_file,
    )

    print_cd_002_summary(
        detail_dataframe=detail_dataframe,
        summary_dataframe=summary_dataframe,
        output_file=output_file,
    )
