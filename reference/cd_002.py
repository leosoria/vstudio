"""
CD_002 - Summary Of Cash Disbursements By Vendor.

Analysis:
- Module: Cash Disbursements
- Analysis Code: CD_02
- Analysis Title: Summary Of Cash Disbursements By Vendor

Description:
Aggregates / summarizes cash disbursements by vendor.

Procedure:
Summarize and trend disbursements by vendor.

Analytic logic:
Aggregate payments by vendor, amount/count over period, using the same clean
payment universe as CD_001.

Context:
Payment visibility for vendor oversight.

Design notes:
- CD_002 uses the same clean payment universe as CD_001 by reusing the CD_001
  detail dataframe builder.
- It does not read the CD01 sheet from the output workbook.
- It does not exclude intercompanies because that is not part of the CD_02
  control objective.
- It writes/replaces only the CD02 sheet.
- It does not delete sheets from other controls.
"""

import pandas as pd

from core.cd_common import (
    apply_standard_cd_formatting,
    find_column,
    get_cd_output_file,
    load_cd_base_data,
    normalize_text,
    open_or_create_cd_output_workbook,
    recreate_cd_sheet,
    require_columns,
    save_cd_output_workbook,
    write_dataframe_to_sheet,
)
from modules.CD.cd_001 import build_cd_001_dataframe


SHEET_NAME = "CD02"

SOURCE_REQUIRED_COLUMNS = {
    "debit_credit_indicator": [
        "D/C",
    ],
}

REVERSAL_COLUMNS = {
    "reversal_with": [
        "Estorno c/",
        "Estorno c",
        "Estorno c.",
    ],
    "reversal_document": [
        "Estorno",
    ],
}

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


def get_optional_reversal_column(source_dataframe, logical_name):
    """
    Return an optional reversal column by logical name.
    """
    return find_column(
        dataframe=source_dataframe,
        possible_names=REVERSAL_COLUMNS[logical_name],
    )


def is_blank_series(series):
    """
    Return True for blank/null values in a pandas series.
    """
    return series.fillna("").apply(normalize_text) == ""


def build_cd_002_clean_source_dataframe(source_dataframe):
    """
    Return source dataframe cleaned for reversal fields plus validation counts.

    CD_002 ultimately uses build_cd_001_dataframe() for the detail universe.
    These counts make the CD_002 run transparent for audit review and show the
    expected clean CD filters:
    - D/C = S
    - Estorno c/ blank
    - Estorno blank
    """
    required_columns = require_columns(
        source_dataframe,
        SOURCE_REQUIRED_COLUMNS,
    )

    debit_credit_column = required_columns["debit_credit_indicator"]

    payment_rows_mask = (
        source_dataframe[debit_credit_column]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        == "S"
    )

    payment_rows = source_dataframe[payment_rows_mask].copy()

    reversal_with_column = get_optional_reversal_column(
        source_dataframe=payment_rows,
        logical_name="reversal_with",
    )
    reversal_document_column = get_optional_reversal_column(
        source_dataframe=payment_rows,
        logical_name="reversal_document",
    )

    if reversal_with_column is None:
        reversal_with_blank_mask = pd.Series(
            [True] * len(payment_rows),
            index=payment_rows.index,
        )
    else:
        reversal_with_blank_mask = is_blank_series(
            payment_rows[reversal_with_column]
        )

    if reversal_document_column is None:
        reversal_document_blank_mask = pd.Series(
            [True] * len(payment_rows),
            index=payment_rows.index,
        )
    else:
        reversal_document_blank_mask = is_blank_series(
            payment_rows[reversal_document_column]
        )

    valid_payment_rows_mask = reversal_with_blank_mask & reversal_document_blank_mask
    clean_source_dataframe = payment_rows[valid_payment_rows_mask].copy()

    source_counts = {
        "source_rows_loaded": len(source_dataframe),
        "payment_rows_dc_s": len(payment_rows),
        "reversed_payment_rows_excluded": int(
            len(payment_rows) - valid_payment_rows_mask.sum()
        ),
    }

    return clean_source_dataframe, source_counts


def build_cd_002_dataframe(detail_dataframe):
    """
    Build CD_002 vendor summary dataframe from the CD payment detail dataframe.

    The detail dataframe is expected to follow CD_001's clean payment-style
    output and must already exclude reversed/canceled payments.

    Aggregation rules:
    - Group by Company, Vendor Code and Vendor Name.
    - Qty Payments counts unique KEY_DOC values to avoid overcounting payments
      that have more than one line.
    - Total Amount USD sums Payment Amount USD.
    - First Payment is the earliest Payment Date.
    - Last Payment is the latest Payment Date.
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


def print_cd_002_summary(source_counts, detail_dataframe, summary_dataframe, output_file):
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
    print(f"CD_002 source rows loaded: {source_counts['source_rows_loaded']}")
    print(f"CD_002 payment rows D/C = S: {source_counts['payment_rows_dc_s']}")
    print(
        "CD_002 reversed payment rows excluded: "
        f"{source_counts['reversed_payment_rows_excluded']}"
    )
    print(f"CD_002 valid payment rows: {len(detail_dataframe)}")
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

    clean_source_dataframe, source_counts = build_cd_002_clean_source_dataframe(
        source_dataframe=source_dataframe,
    )

    detail_dataframe = build_cd_001_dataframe(
        source_dataframe=clean_source_dataframe,
        context=context,
    )

    summary_dataframe = build_cd_002_dataframe(
        detail_dataframe=detail_dataframe,
    )

    if len(summary_dataframe) == 0:
        print()
        print("WARNING: CD_002 generated 0 rows.")
        print("Possible causes:")
        print("- No valid rows with D/C = S were found.")
        print("- All D/C = S rows were reversed/canceled.")
        print("- COMPANIES filter does not match Empr values.")
        print("- The input file is empty or the wrong sheet was read.")
        print("- The LBR CA input file for the module TO date was not found.")
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
        source_counts=source_counts,
        detail_dataframe=detail_dataframe,
        summary_dataframe=summary_dataframe,
        output_file=output_file,
    )
