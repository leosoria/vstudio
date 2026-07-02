"""
FAM_002 - Fixed Assets With Zero Book Value.

Analysis:
- Module: Fixed Asset Management
- Analysis Code: FAM_02
- Analysis Title: Fixed Assets With Zero Book Value

Description:
Identifies assets with zero book value.

Procedure:
List assets where cost including revaluation equals accumulated depreciation.

Analytic Logic:
Find assets where net book value is zero, using a small tolerance to avoid
rounding differences.

Context:
Listing assists monitoring of asset base and control activities.

Input:
- AR01 export saved as:
    input/LBR FAM AR_YYYYMMDD.xlsx
- FX rates saved as:
    input/FxRates_YYYYMMDD.xlsx

Output:
- Workbook:
    output/LBR_Results_FAM_YYYYMMDD.xlsx
- Sheet:
    FAM02

Rules:
- This control writes/replaces only sheet FAM02.
- It does not delete FAM01 or sheets from other controls.
- The source universe and output layout are intentionally aligned with FAM01.
"""

import pandas as pd

from modules.FAM.fam_001 import (
    AMOUNT_COLUMNS,
    DATE_COLUMNS,
    INTEGER_COLUMNS,
    REQUIRED_COLUMNS,
    build_fam_001_dataframe,
    remove_ar01_summary_rows,
    resolve_optional_columns,
)
from core.fam_common import (
    apply_standard_fam_formatting,
    build_fx_rate_lookup,
    filter_by_company,
    get_fam_output_file,
    load_fam_ar01_data,
    load_fx_rates_data,
    open_or_create_fam_output_workbook,
    parse_number,
    recreate_fam_sheet,
    require_columns,
    save_fam_output_workbook,
    write_dataframe_to_sheet,
)


SHEET_NAME = "FAM02"
DEFAULT_ZERO_BOOK_VALUE_TOLERANCE = 0.01
BOOK_VALUE_OUTPUT_COLUMN = "VNC en fecha de fin (LC)"


def get_zero_book_value_tolerance(context):
    """
    Return the FAM02 tolerance.

    FAM02 accepts an optional numeric tolerance in the control PARAM1 field.
    If PARAM1 is blank or cannot be parsed, the default 0.01 tolerance is used.
    """
    control = context.get("control", {})
    raw_tolerance = control.get("param1", "")

    parsed_tolerance = parse_number(raw_tolerance)

    if pd.isna(parsed_tolerance):
        return DEFAULT_ZERO_BOOK_VALUE_TOLERANCE

    return abs(float(parsed_tolerance))


def filter_zero_book_value_assets(output_dataframe, tolerance):
    """
    Keep assets where net book value is zero within tolerance.

    The control uses FAM01's normalized output dataframe instead of reading the
    FAM01 worksheet from Excel, keeping the control independent and avoiding
    dependencies between output sheets.
    """
    book_value = output_dataframe[BOOK_VALUE_OUTPUT_COLUMN].apply(parse_number)
    zero_book_value_mask = book_value.fillna(0).abs() <= tolerance

    return output_dataframe[zero_book_value_mask].copy()


def build_fam_002_dataframe(source_dataframe, required_columns, optional_columns, fx_lookup, tolerance):
    """
    Build the FAM02 output dataframe.

    FAM02 starts from the same base universe and output structure as FAM01 and
    then filters assets with zero net book value.
    """
    fam_001_dataframe = build_fam_001_dataframe(
        source_dataframe=source_dataframe,
        required_columns=required_columns,
        optional_columns=optional_columns,
        fx_lookup=fx_lookup,
    )

    return filter_zero_book_value_assets(
        output_dataframe=fam_001_dataframe,
        tolerance=tolerance,
    )


def run_fam_002(context):
    """
    Run FAM_002 and write the FAM02 sheet.
    """
    source_dataframe = load_fam_ar01_data(context)

    fx_dataframe = load_fx_rates_data(context)
    fx_lookup = build_fx_rate_lookup(fx_dataframe)

    required_columns = require_columns(
        source_dataframe,
        REQUIRED_COLUMNS,
    )
    optional_columns = resolve_optional_columns(source_dataframe)

    source_dataframe = remove_ar01_summary_rows(
        source_dataframe=source_dataframe,
        required_columns=required_columns,
    )

    source_dataframe = filter_by_company(
        dataframe=source_dataframe,
        company_column=required_columns["company_code"],
        companies_filter=context["module"].get("companies", ""),
    )

    tolerance = get_zero_book_value_tolerance(context)

    output_dataframe = build_fam_002_dataframe(
        source_dataframe=source_dataframe,
        required_columns=required_columns,
        optional_columns=optional_columns,
        fx_lookup=fx_lookup,
        tolerance=tolerance,
    )

    output_file = get_fam_output_file(context)

    workbook = open_or_create_fam_output_workbook(output_file)
    worksheet = recreate_fam_sheet(workbook, SHEET_NAME)

    write_dataframe_to_sheet(
        worksheet=worksheet,
        dataframe=output_dataframe,
    )

    apply_standard_fam_formatting(
        worksheet=worksheet,
        dataframe=output_dataframe,
        date_columns=DATE_COLUMNS,
        amount_columns=AMOUNT_COLUMNS,
        integer_columns=INTEGER_COLUMNS,
    )

    save_fam_output_workbook(
        workbook=workbook,
        output_file=output_file,
    )

    print(f"FAM_002 output file: {output_file}")
    print(f"FAM_002 sheet: {SHEET_NAME}")
    print(f"FAM_002 zero book value tolerance: {tolerance}")
    print(f"FAM_002 rows: {len(output_dataframe)}")
