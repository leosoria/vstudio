"""
FAM_003 - Fixed Assets With No Depreciation.

Analysis:
- Module: Fixed Asset Management
- Analysis Code: FAM_03
- Analysis Title: Fixed Assets With No Depreciation

Description:
Identifies assets with acquisition/capitalized value but no accumulated
depreciation.

Procedure:
List active/capitalized assets where accumulated depreciation is zero or within
tolerance.

Analytic Logic:
Start from the same base universe as FAM01, build the normalized FAM01-like
dataframe without reading the FAM01 worksheet, and keep assets where:
- abs(Amortización acumulada en fecha fin (LC)) <= tolerance
- abs(CAP en fecha de fin (LC)) > tolerance

Context:
Listing assists monitoring assets that may not be depreciating as expected.

Input:
- AR01 export saved as:
    input/LBR FAM AR_YYYYMMDD.xlsx
- FX rates saved as:
    input/FxRates_YYYYMMDD.xlsx

Output:
- Workbook:
    output/LBR_Results_FAM_YYYYMMDD.xlsx
- Sheet:
    FAM03

Rules:
- This control writes/replaces only sheet FAM03.
- It does not delete FAM01, FAM02 or sheets from other controls.
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


SHEET_NAME = "FAM03"
DEFAULT_NO_DEPRECIATION_TOLERANCE = 0.01
ACCUMULATED_DEPRECIATION_OUTPUT_COLUMN = "Amortización acumulada en fecha fin (LC)"
CAP_OUTPUT_COLUMN = "CAP en fecha de fin (LC)"


def get_no_depreciation_tolerance(context):
    """
    Return the FAM03 tolerance.

    FAM03 accepts an optional numeric tolerance in the control PARAM1 field.
    If PARAM1 is blank or cannot be parsed, the default 0.01 tolerance is used.
    """
    control = context.get("control", {})
    raw_tolerance = control.get("param1", "")

    parsed_tolerance = parse_number(raw_tolerance)

    if pd.isna(parsed_tolerance):
        return DEFAULT_NO_DEPRECIATION_TOLERANCE

    return abs(float(parsed_tolerance))


def filter_no_depreciation_assets(output_dataframe, tolerance):
    """
    Keep capitalized assets with no accumulated depreciation within tolerance.

    The control uses FAM01's normalized output dataframe instead of reading the
    FAM01 or FAM02 worksheets from Excel, keeping the control independent and
    avoiding dependencies between output sheets.
    """
    accumulated_depreciation = output_dataframe[
        ACCUMULATED_DEPRECIATION_OUTPUT_COLUMN
    ].apply(parse_number)
    cap_value = output_dataframe[CAP_OUTPUT_COLUMN].apply(parse_number)

    no_depreciation_mask = accumulated_depreciation.fillna(0).abs() <= tolerance
    has_capitalized_value_mask = cap_value.fillna(0).abs() > tolerance

    return output_dataframe[
        no_depreciation_mask & has_capitalized_value_mask
    ].copy()


def build_fam_003_dataframe(
    source_dataframe,
    required_columns,
    optional_columns,
    fx_lookup,
    tolerance,
):
    """
    Build the FAM03 output dataframe.

    FAM03 starts from the same base universe and output structure as FAM01 and
    then filters capitalized assets with no accumulated depreciation.
    """
    fam_001_dataframe = build_fam_001_dataframe(
        source_dataframe=source_dataframe,
        required_columns=required_columns,
        optional_columns=optional_columns,
        fx_lookup=fx_lookup,
    )

    return filter_no_depreciation_assets(
        output_dataframe=fam_001_dataframe,
        tolerance=tolerance,
    )


def run_fam_003(context):
    """
    Run FAM_003 and write the FAM03 sheet.
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

    tolerance = get_no_depreciation_tolerance(context)

    output_dataframe = build_fam_003_dataframe(
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

    print(f"FAM_003 output file: {output_file}")
    print(f"FAM_003 sheet: {SHEET_NAME}")
    print(f"FAM_003 no depreciation tolerance: {tolerance}")
    print(f"FAM_003 rows: {len(output_dataframe)}")
