"""FAM_003 - capitalized fixed assets with no accumulated depreciation.

This control is independent from FAM_001. It reads the AR and FX inputs,
writes an executive summary to FAM03 and writes selected assets to
FAM03_Detail. It replaces only those two worksheets.
"""

from pathlib import Path

import pandas as pd

from openpyxl.styles import Alignment, Font

from core.fam_common import (
    apply_standard_fam_formatting,
    build_fx_rate_lookup,
    clean_dataframe,
    filter_by_company,
    get_fam_output_file,
    load_fx_rates_data,
    normalize_company_output,
    normalize_currency,
    normalize_text,
    open_or_create_fam_output_workbook,
    parse_number,
    recreate_fam_sheet,
    require_columns,
    save_fam_output_workbook,
    to_datetime_value,
    write_dataframe_to_sheet,
)


SUMMARY_SHEET_NAME = "FAM03"
DETAIL_SHEET_NAME = "FAM03_Detail"

DEFAULT_NO_DEPRECIATION_TOLERANCE = 0.01


REQUIRED_COLUMNS = {
    "asset_number": [
        "Asset",
        "Imobilizado",
    ],
    "asset_subnumber": [
        "Subnumber",
        "Subnº",
        "Subn°",
        "Subno",
        "Subnro",
    ],
    "company_code": [
        "Company Code",
        "Empresa",
    ],
    "asset_description": [
        "Asset description",
        "Denominação do imobilizado",
        "Denominacao do imobilizado",
    ],
    "capitalization_date": [
        "Capitalized on",
        "Incorporação em",
        "Incorporacao em",
    ],
    "acquisition_value": [
        "Acquis.val.",
        "ValAquis.",
        "ValAquis",
    ],
    "accumulated_depreciation": [
        "Accum.dep.",
        "Depreciação ac.",
        "Depreciacao ac.",
    ],
    "book_value": [
        "Book val.",
        "Valor contábil",
        "Valor contabil",
    ],
    "asset_class": [
        "Asset Class",
        "Classe imobilizado",
    ],
    "balance_account": [
        "Bal.sh.acct APC",
        "Conta do balanço CAP",
        "Conta do balanco CAP",
    ],
    "depreciation_start_date": [
        "Ord.dep.start date",
        "Depreciação normal",
        "Depreciacao normal",
    ],
    "currency": [
        "Currency",
        "Moeda",
    ],
}


CLASSIFICATION_COLUMN = "Clasificación"

CAP_LC_COLUMN = "CAP en fecha de fin (LC)"
DEPRECIATION_LC_COLUMN = (
    "Amortización acumulada en fecha fin (LC)"
)
BOOK_VALUE_LC_COLUMN = "VNC en fecha de fin (LC)"

FX_DIRECTION_COLUMN = "Dirección FX original"
FX_RAW_RATE_COLUMN = "Tasa FX original"
FX_APPLIED_RATE_COLUMN = "Tasa aplicada USD por LC"
FX_DATE_COLUMN = "Fecha rate USD"
FX_STATUS_COLUMN = "FX Status"

CAP_USD_COLUMN = "CAP en fecha de fin (USD)"
DEPRECIATION_USD_COLUMN = (
    "Amortización acumulada en fecha fin (USD)"
)
BOOK_VALUE_USD_COLUMN = "VNC en fecha de fin (USD)"


DETAIL_COLUMNS = [
    CLASSIFICATION_COLUMN,
    CAP_LC_COLUMN,
    DEPRECIATION_LC_COLUMN,
    BOOK_VALUE_LC_COLUMN,
    "Fecha inicio depreciación",
    "CoCo",
    "Núm.AF",
    "Subnº",
    "Descripción de activo",
    "Clase AF",
    "Fecha capitalización",
    "Cuenta de balance",
    "Moneda",
    FX_DIRECTION_COLUMN,
    FX_RAW_RATE_COLUMN,
    FX_APPLIED_RATE_COLUMN,
    FX_DATE_COLUMN,
    FX_STATUS_COLUMN,
    CAP_USD_COLUMN,
    DEPRECIATION_USD_COLUMN,
    BOOK_VALUE_USD_COLUMN,
]


DATE_COLUMNS = {
    "Fecha inicio depreciación",
    "Fecha capitalización",
    FX_DATE_COLUMN,
}


AMOUNT_COLUMNS = {
    CAP_LC_COLUMN,
    DEPRECIATION_LC_COLUMN,
    BOOK_VALUE_LC_COLUMN,
    CAP_USD_COLUMN,
    DEPRECIATION_USD_COLUMN,
    BOOK_VALUE_USD_COLUMN,
}


INTEGER_COLUMNS = set()


def get_no_depreciation_tolerance(context):
    """PARAM2 overrides the default tolerance; PARAM1 lists excluded accounts."""
    raw = context.get("control", {}).get("param2", "")
    if pd.isna(raw) or str(raw).strip() == "":
        return DEFAULT_NO_DEPRECIATION_TOLERANCE
    value = parse_number(raw)
    import math
    if pd.isna(value) or not math.isfinite(float(value)) or float(value) <= 0:
        raise ValueError("FAM03 PARAM2 must be positive or blank for tolerance 0.01.")
    return float(value)


def get_fam03_excluded_accounts(context):
    """Read comma-separated account identifiers."""
    import re
    raw = context.get("control", {}).get("param1", "")
    if pd.isna(raw) or str(raw).strip() == "":
        return set()
    accounts = set()
    for token in str(raw).split(","):
        token = token.strip()
        if not re.fullmatch(r"[0-9]+(?:\.0)?", token):
            raise ValueError(
                "FAM03 PARAM1: use account numbers separated by commas, "
                "e.g. 1203010022,1203010011."
            )
        if token.endswith(".0"):
            token = token[:-2]
        accounts.add(token.lstrip("0") or "0")
    return accounts


def exclude_fam03_accounts(dataframe, columns, context):
    accounts = get_fam03_excluded_accounts(context)
    if not accounts:
        print("FAM_003 account exclusions (PARAM1): none")
        return dataframe

    def account_key(value):
        value = normalize_identifier(value)
        return value.lstrip("0") or "0"

    excluded = dataframe[columns["balance_account"]].map(account_key).isin(accounts)
    print("FAM_003 excluded accounts: " + ", ".join(sorted(accounts)))
    print(f"FAM_003 asset rows excluded by account: {int(excluded.sum())}")
    return dataframe.loc[~excluded].copy()


def get_analysis_date(context):
    """
    Return the normalized module end date.
    """
    value = to_datetime_value(
        context["module"].get(
            "to",
            "",
        )
    )

    if pd.isna(value):
        raise ValueError(
            "FAM_003 requires a valid module TO date."
        )

    return value.normalize()


def get_ar_input_file(context):
    """
    Resolve the AR input filename.

    Supported names:
    - LBR FAM_AR_YYYYMMDD.xlsx
    - LBR FAM AR_YYYYMMDD.xlsx
    """
    suffix = get_analysis_date(
        context
    ).strftime("%Y%m%d")

    folder = Path(
        context["input_folder"]
    )

    candidates = [
        folder / f"LBR FAM_AR_{suffix}.xlsx",
        folder / f"LBR FAM AR_{suffix}.xlsx",
    ]

    for candidate in candidates:
        if candidate.exists():
            return candidate

    raise FileNotFoundError(
        "FAM AR input file not found. "
        "Expected one of: "
        + ", ".join(
            str(candidate)
            for candidate in candidates
        )
    )


def read_ar_dataframe(context):
    """
    Read the AR input directly without depending on FAM_001.
    """
    input_file = get_ar_input_file(
        context
    )

    dataframe = pd.read_excel(
        input_file,
        dtype=object,
    )

    dataframe = clean_dataframe(
        dataframe
    )

    print(
        f"FAM_003 AR input file: "
        f"{input_file}"
    )

    print(
        f"FAM_003 AR rows loaded: "
        f"{len(dataframe)}"
    )

    return dataframe


def normalize_identifier(value):
    """
    Remove Excel's trailing .0 while preserving identifier text.
    """
    value = normalize_text(value)

    if value.endswith(".0"):
        return value[:-2]

    return value


def remove_summary_rows(
    dataframe,
    columns,
):
    """
    Keep asset-level AR rows only.
    """
    has_company = dataframe[
        columns["company_code"]
    ].apply(
        lambda value: normalize_text(value) != ""
    )

    has_asset = dataframe[
        columns["asset_number"]
    ].apply(
        lambda value: normalize_identifier(value) != ""
    )

    has_description = dataframe[
        columns["asset_description"]
    ].apply(
        lambda value: normalize_text(value) != ""
    )

    valid_rows = (
        has_company
        & has_asset
        & has_description
    )

    return dataframe.loc[
        valid_rows
    ].copy()


def get_fx_details(
    currency,
    fx_lookup,
):
    """
    Return traceable FX details.

    Applied rate always represents:
        USD per one unit of local currency.
    """
    currency = normalize_currency(
        currency
    )

    details = fx_lookup.get(
        currency
    )

    if (
        details is None
        or pd.isna(
            details.get(
                "rate",
                pd.NA,
            )
        )
    ):
        return {
            "status": "RATE_NOT_FOUND",
            "direction": "",
            "raw_rate": pd.NA,
            "applied_rate": pd.NA,
            "date": pd.NaT,
        }

    applied_rate = details.get(
        "rate",
        pd.NA,
    )

    if (
        pd.isna(applied_rate)
        or float(applied_rate) == 0
    ):
        return {
            "status": "INVALID_RATE",
            "direction": details.get(
                "direction",
                "",
            ),
            "raw_rate": details.get(
                "sap_raw_rate",
                pd.NA,
            ),
            "applied_rate": pd.NA,
            "date": details.get(
                "date",
                pd.NaT,
            ),
        }

    return {
        "status": "OK",
        "direction": details.get(
            "direction",
            f"USD->{currency}",
        ),
        "raw_rate": details.get(
            "sap_raw_rate",
            applied_rate,
        ),
        "applied_rate": 1.0 / float(applied_rate),
        "date": details.get(
            "date",
            pd.NaT,
        ),
    }


def convert_to_usd(
    value,
    fx_details,
):
    """
    Convert LC to USD with the normalized USD-per-LC multiplier.
    """
    value = parse_number(value)

    rate = fx_details.get(
        "applied_rate",
        pd.NA,
    )

    if (
        pd.isna(value)
        or pd.isna(rate)
        or float(rate) == 0
    ):
        return pd.NA

    return (
        float(value)
        * float(rate)
    )


def classify_asset(
    cap,
    depreciation,
    book_value,
    depreciation_start,
    cutoff,
    tolerance,
):
    """
    Return the FAM03 classification.

    NO_ACCUMULATED_DEPRECIATION:
    - asset has acquisition/capitalized value;
    - accumulated depreciation is zero within tolerance;
    - depreciation start date is on or before the cut-off date.

    MISSING_DEPRECIATION_START_DATE:
    - asset has acquisition/capitalized value;
    - accumulated depreciation is zero;
    - depreciation start date is missing.

    INCOMPLETE_OR_INVALID_DATA:
    - CAP, accumulated depreciation or NBV is missing/invalid.
    """
    if (
        pd.isna(cap)
        or pd.isna(depreciation)
        or pd.isna(book_value)
    ):
        return "INCOMPLETE_OR_INVALID_DATA"

    if abs(float(cap)) <= tolerance:
        return ""

    if abs(float(depreciation)) > tolerance:
        return ""

    if pd.isna(depreciation_start):
        return "MISSING_DEPRECIATION_START_DATE"

    if depreciation_start.normalize() > cutoff:
        return ""

    return "NO_ACCUMULATED_DEPRECIATION"


def build_fam_003_dataframe(
    source_dataframe,
    columns,
    fx_lookup,
    cutoff,
    tolerance,
):
    """
    Build FAM03 directly from the AR and FX inputs.
    """
    records = []

    for _, row in source_dataframe.iterrows():
        cap = parse_number(
            row[
                columns["acquisition_value"]
            ]
        )

        depreciation = parse_number(
            row[
                columns["accumulated_depreciation"]
            ]
        )

        book_value = parse_number(
            row[
                columns["book_value"]
            ]
        )

        depreciation_start = to_datetime_value(
            row[
                columns["depreciation_start_date"]
            ]
        )

        classification = classify_asset(
            cap=cap,
            depreciation=depreciation,
            book_value=book_value,
            depreciation_start=depreciation_start,
            cutoff=cutoff,
            tolerance=tolerance,
        )

        if classification == "":
            continue

        currency = normalize_currency(
            row[
                columns["currency"]
            ]
        )

        fx_details = get_fx_details(
            currency=currency,
            fx_lookup=fx_lookup,
        )

        records.append(
            {
                CLASSIFICATION_COLUMN: classification,
                CAP_LC_COLUMN: cap,
                DEPRECIATION_LC_COLUMN: depreciation,
                BOOK_VALUE_LC_COLUMN: book_value,
                "Fecha inicio depreciación": (
                    depreciation_start
                ),
                "CoCo": normalize_company_output(
                    row[
                        columns["company_code"]
                    ]
                ),
                "Núm.AF": normalize_identifier(
                    row[
                        columns["asset_number"]
                    ]
                ),
                "Subnº": normalize_identifier(
                    row[
                        columns["asset_subnumber"]
                    ]
                ),
                "Descripción de activo": normalize_text(
                    row[
                        columns["asset_description"]
                    ]
                ),
                "Clase AF": normalize_text(
                    row[
                        columns["asset_class"]
                    ]
                ),
                "Fecha capitalización": to_datetime_value(
                    row[
                        columns["capitalization_date"]
                    ]
                ),
                "Cuenta de balance": normalize_identifier(
                    row[
                        columns["balance_account"]
                    ]
                ),
                "Moneda": currency,
                FX_DIRECTION_COLUMN: (
                    fx_details["direction"]
                ),
                FX_RAW_RATE_COLUMN: (
                    fx_details["raw_rate"]
                ),
                FX_APPLIED_RATE_COLUMN: (
                    fx_details["applied_rate"]
                ),
                FX_DATE_COLUMN: (
                    fx_details["date"]
                ),
                FX_STATUS_COLUMN: (
                    fx_details["status"]
                ),
                CAP_USD_COLUMN: convert_to_usd(
                    cap,
                    fx_details,
                ),
                DEPRECIATION_USD_COLUMN: convert_to_usd(
                    depreciation,
                    fx_details,
                ),
                BOOK_VALUE_USD_COLUMN: convert_to_usd(
                    book_value,
                    fx_details,
                ),
            }
        )

    return pd.DataFrame(
        records,
        columns=DETAIL_COLUMNS,
    )


def build_summary(
    total_assets,
    detail_dataframe,
    cutoff,
    tolerance,
):
    """
    Build the FAM03 executive summary in English.
    """
    if detail_dataframe.empty:
        classification_counts = pd.Series(
            dtype="int64"
        )

        fx_ok_count = 0
        fx_not_found_count = 0
        fx_invalid_count = 0
    else:
        classification_counts = (
            detail_dataframe[
                CLASSIFICATION_COLUMN
            ]
            .value_counts()
        )

        fx_ok_count = int(
            detail_dataframe[
                FX_STATUS_COLUMN
            ]
            .eq("OK")
            .sum()
        )

        fx_not_found_count = int(
            detail_dataframe[
                FX_STATUS_COLUMN
            ]
            .eq("RATE_NOT_FOUND")
            .sum()
        )

        fx_invalid_count = int(
            detail_dataframe[
                FX_STATUS_COLUMN
            ]
            .eq("INVALID_RATE")
            .sum()
        )

    selected_count = len(
        detail_dataframe
    )

    if selected_count == 0:
        overall_result = (
            "No potential FAM03 exceptions identified"
        )
    else:
        overall_result = (
            f"{selected_count} potential FAM03 "
            f"exception(s) identified"
        )

    rows = [
        {
            "Metric": "Analysis cut-off date",
            "Result": cutoff.date(),
        },
        {
            "Metric": "Applied tolerance",
            "Result": tolerance,
        },
        {
            "Metric": "Total assets analyzed",
            "Result": total_assets,
        },
        {
            "Metric": (
                "Records included in FAM03_Detail"
            ),
            "Result": selected_count,
        },
        {
            "Metric": (
                "Potential FAM03 exceptions"
            ),
            "Result": selected_count,
        },
        {
            "Metric": (
                "Successful FX conversions"
            ),
            "Result": fx_ok_count,
        },
        {
            "Metric": "FX rates not found",
            "Result": fx_not_found_count,
        },
        {
            "Metric": "Invalid FX rates",
            "Result": fx_invalid_count,
        },
        {
            "Metric": "Overall result",
            "Result": overall_result,
        },
    ]

    classification_labels = {
        "NO_ACCUMULATED_DEPRECIATION": (
            "Capitalized assets with no "
            "accumulated depreciation"
        ),
        "MISSING_DEPRECIATION_START_DATE": (
            "Assets with no depreciation "
            "start date"
        ),
        "INCOMPLETE_OR_INVALID_DATA": (
            "Assets with incomplete or "
            "invalid data"
        ),
    }

    for classification, count in (
        classification_counts.items()
    ):
        display_label = classification_labels.get(
            classification,
            classification,
        )

        rows.append(
            {
                "Metric": (
                    f"Classification: "
                    f"{display_label}"
                ),
                "Result": int(count),
            }
        )

    return pd.DataFrame(
        rows,
        columns=[
            "Metric",
            "Result",
        ],
    )

def format_fx_rates(
    worksheet,
    dataframe,
):
    """
    Display rates with enough precision for audit recalculation.
    """
    column_positions = {
        column_name: column_index
        for column_index, column_name
        in enumerate(
            dataframe.columns,
            start=1,
        )
    }

    for column_name in (
        FX_RAW_RATE_COLUMN,
        FX_APPLIED_RATE_COLUMN,
    ):
        column_index = column_positions.get(
            column_name
        )

        if column_index is None:
            continue

        for row_index in range(
            2,
            len(dataframe) + 2,
        ):
            worksheet.cell(
                row=row_index,
                column=column_index,
            ).number_format = "0.00000000"

def format_summary_sheet(worksheet):
    """
    Add title, description and formatting to the FAM03 summary.
    """
    worksheet.insert_rows(
        1,
        amount=3,
    )

    worksheet.merge_cells(
        "A1:B1"
    )

    worksheet.merge_cells(
        "A2:B2"
    )

    worksheet["A1"] = (
        "FAM03 - Fixed Assets Control Summary"
    )

    worksheet["A2"] = (
        "Identification of capitalized assets with no "
        "accumulated depreciation after the depreciation "
        "start date"
    )

    worksheet["A1"].font = Font(
        bold=True,
        size=14,
    )

    worksheet["A2"].font = Font(
        italic=True,
        size=11,
    )

    worksheet["A1"].alignment = Alignment(
        horizontal="left",
        vertical="center",
    )

    worksheet["A2"].alignment = Alignment(
        horizontal="left",
        vertical="center",
        wrap_text=True,
    )

    worksheet.row_dimensions[1].height = 22
    worksheet.row_dimensions[2].height = 32

    worksheet.column_dimensions["A"].width = 58
    worksheet.column_dimensions["B"].width = 55

    worksheet.freeze_panes = "A5"

def run_fam_003(context):
    """
    Run independent FAM03.

    Replaces only:
    - FAM03
    - FAM03_Detail
    """
    cutoff = get_analysis_date(
        context
    )

    tolerance = get_no_depreciation_tolerance(
        context
    )

    source_dataframe = read_ar_dataframe(
        context
    )

    columns = require_columns(
        source_dataframe,
        REQUIRED_COLUMNS,
    )

    source_dataframe = remove_summary_rows(
        dataframe=source_dataframe,
        columns=columns,
    )

    source_dataframe = filter_by_company(
        dataframe=source_dataframe,
        company_column=(
            columns["company_code"]
        ),
        companies_filter=(
            context["module"].get(
                "companies",
                "",
            )
        ),
    )

    source_dataframe = exclude_fam03_accounts(source_dataframe, columns, context)

    fx_dataframe = load_fx_rates_data(
        context
    )

    fx_lookup = build_fx_rate_lookup(
        fx_dataframe
    )

    detail_dataframe = build_fam_003_dataframe(
        source_dataframe=source_dataframe,
        columns=columns,
        fx_lookup=fx_lookup,
        cutoff=cutoff,
        tolerance=tolerance,
    )

    summary_dataframe = build_summary(
        total_assets=len(source_dataframe),
        detail_dataframe=detail_dataframe,
        cutoff=cutoff,
        tolerance=tolerance,
    )

    output_file = get_fam_output_file(
        context
    )

    workbook = open_or_create_fam_output_workbook(
        output_file
    )

    summary_worksheet = recreate_fam_sheet(
        workbook,
        SUMMARY_SHEET_NAME,
    )

    detail_worksheet = recreate_fam_sheet(
        workbook,
        DETAIL_SHEET_NAME,
    )

    write_dataframe_to_sheet(
        worksheet=summary_worksheet,
        dataframe=summary_dataframe,
    )

    
    format_summary_sheet(
        worksheet=summary_worksheet,
    )
    write_dataframe_to_sheet(
        worksheet=detail_worksheet,
        dataframe=detail_dataframe,
    )

    apply_standard_fam_formatting(
        worksheet=detail_worksheet,
        dataframe=detail_dataframe,
        date_columns=DATE_COLUMNS,
        amount_columns=AMOUNT_COLUMNS,
        integer_columns=INTEGER_COLUMNS,
    )

    summary_worksheet.column_dimensions[
        "A"
    ].width = 55

    summary_worksheet.column_dimensions[
        "B"
    ].width = 42

    format_fx_rates(
        worksheet=detail_worksheet,
        dataframe=detail_dataframe,
    )

    save_fam_output_workbook(
        workbook=workbook,
        output_file=output_file,
    )

    print(
        f"FAM_003 output file: "
        f"{output_file}"
    )

    print(
        f"FAM_003 summary sheet: "
        f"{SUMMARY_SHEET_NAME}"
    )

    print(
        f"FAM_003 detail sheet: "
        f"{DETAIL_SHEET_NAME}"
    )

    print(
        f"FAM_003 tolerance: "
        f"{tolerance}"
    )

    print(
        f"FAM_003 assets analyzed: "
        f"{len(source_dataframe)}"
    )

    print(
        f"FAM_003 exception rows: "
        f"{len(detail_dataframe)}"
    )


