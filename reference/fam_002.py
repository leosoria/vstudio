"""
FAM_002 - Fixed Assets Depreciation and Net Book Value Analysis.

This control is independent from FAM_001.

It reads directly:
- input/LBR FAM AR_YYYYMMDD.xlsx
- input/FxRates_YYYYMMDD.xlsx

It writes/replaces only:
- output/LBR_Results_FAM_YYYYMMDD.xlsx
- sheet FAM02

Accounting logic:
- Expected NBV = acquisition value + accumulated depreciation
- Difference = reported NBV - expected NBV

Output population:
- accounting exceptions;
- correctly reconciled zero-book-value assets as informational records.

FX convention:
- Applied FX rate always means USD per one unit of local currency.
- USD amount = local-currency amount * applied FX rate.
- A USD -> local-currency rate is inverted exactly once.
"""

from pathlib import Path

import pandas as pd

from core.fam_common import (
    apply_standard_fam_formatting,
    filter_by_company,
    get_fam_output_file,
    open_or_create_fam_output_workbook,
    parse_number,
    recreate_fam_sheet,
    save_fam_output_workbook,
    write_dataframe_to_sheet,
)


SHEET_NAME = "FAM02"
DEFAULT_TOLERANCE = 0.01
MAX_PREVIOUS_FX_DAYS = 10
PREFERRED_EXCHANGE_RATE_TYPES = ("EN", "M")


AR_REQUIRED_COLUMNS = {
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
        "Valor aquisição",
        "Valor aquisicao",
    ],
    "accumulated_depreciation": [
        "Accum.dep.",
        "Depreciação ac.",
        "Depreciacao ac.",
        "Depreciação acumulada",
        "Depreciacao acumulada",
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


FX_REQUIRED_COLUMNS = {
    "rate_type": [
        "ExRt",
        "TCot",
        "CgCâ",
        "CgC",
    ],
    "from_currency": [
        "From",
        "De",
        "From Currency",
    ],
    "to_currency": [
        "To",
        "Para",
        "A",
        "To Currency",
    ],
    "valid_from": [
        "Valid from",
        "Vál.desde",
        "Val.desde",
        "Válido de",
        "Valido de",
    ],
    "raw_rate": [
        "Exch. Rate",
        "Exchange Rate",
        "Taxa câmbio",
        "Taxa cambio",
        "Tipo cambio",
    ],
    "factor_from": [
        "Ratio (from)",
        "Factor From",
        "Fator (origem)",
        "Factor (de)",
    ],
    "factor_to": [
        "Ratio (to)",
        "Factor To",
        "Fator (para)",
        "Factor (a)",
    ],
}


CONTROL_RESULT_COLUMN = "Resultado del control"
CLASSIFICATION_COLUMN = "Classification"

ACQUISITION_VALUE_LC_COLUMN = "CAP en fecha de fin (LC)"
ACCUMULATED_DEPRECIATION_LC_COLUMN = (
    "Amortización acumulada en fecha fin (LC)"
)
EXPECTED_BOOK_VALUE_LC_COLUMN = "VNC esperado en fecha de fin (LC)"
REPORTED_BOOK_VALUE_LC_COLUMN = "VNC en fecha de fin (LC)"
CONTROL_DIFFERENCE_LC_COLUMN = "Diferencia de control (LC)"

FX_ORIGINAL_DIRECTION_COLUMN = "Dirección FX original"
FX_ORIGINAL_RATE_COLUMN = "Tasa FX original"
FX_APPLIED_RATE_COLUMN = "Tasa aplicada USD por LC"
FX_DATE_COLUMN = "Fecha rate USD"
FX_STATUS_COLUMN = "FX Status"

ACQUISITION_VALUE_USD_COLUMN = "CAP en fecha de fin (USD)"
ACCUMULATED_DEPRECIATION_USD_COLUMN = (
    "Amortización acumulada en fecha fin (USD)"
)
EXPECTED_BOOK_VALUE_USD_COLUMN = "VNC esperado en fecha de fin (USD)"
REPORTED_BOOK_VALUE_USD_COLUMN = "VNC en fecha de fin (USD)"
CONTROL_DIFFERENCE_USD_COLUMN = "Diferencia de control (USD)"


OUTPUT_COLUMNS = [
    CONTROL_RESULT_COLUMN,
    CLASSIFICATION_COLUMN,
    ACQUISITION_VALUE_LC_COLUMN,
    ACCUMULATED_DEPRECIATION_LC_COLUMN,
    EXPECTED_BOOK_VALUE_LC_COLUMN,
    REPORTED_BOOK_VALUE_LC_COLUMN,
    CONTROL_DIFFERENCE_LC_COLUMN,
    "CoCo",
    "Núm.AF",
    "Subnº",
    "Descripción de activo",
    "Clase AF",
    "Fecha capitalización",
    "Fecha inicio depreciación",
    "Cuenta de balance",
    "Moneda",
    FX_ORIGINAL_DIRECTION_COLUMN,
    FX_ORIGINAL_RATE_COLUMN,
    FX_APPLIED_RATE_COLUMN,
    FX_DATE_COLUMN,
    FX_STATUS_COLUMN,
    ACQUISITION_VALUE_USD_COLUMN,
    ACCUMULATED_DEPRECIATION_USD_COLUMN,
    EXPECTED_BOOK_VALUE_USD_COLUMN,
    REPORTED_BOOK_VALUE_USD_COLUMN,
    CONTROL_DIFFERENCE_USD_COLUMN,
]


DATE_COLUMNS = {
    "Fecha capitalización",
    "Fecha inicio depreciación",
    FX_DATE_COLUMN,
}


AMOUNT_COLUMNS = {
    ACQUISITION_VALUE_LC_COLUMN,
    ACCUMULATED_DEPRECIATION_LC_COLUMN,
    EXPECTED_BOOK_VALUE_LC_COLUMN,
    REPORTED_BOOK_VALUE_LC_COLUMN,
    CONTROL_DIFFERENCE_LC_COLUMN,
    FX_ORIGINAL_RATE_COLUMN,
    FX_APPLIED_RATE_COLUMN,
    ACQUISITION_VALUE_USD_COLUMN,
    ACCUMULATED_DEPRECIATION_USD_COLUMN,
    EXPECTED_BOOK_VALUE_USD_COLUMN,
    REPORTED_BOOK_VALUE_USD_COLUMN,
    CONTROL_DIFFERENCE_USD_COLUMN,
}


INTEGER_COLUMNS = set()


def normalize_text(value):
    """
    Normalize a text value.
    """
    if value is None or pd.isna(value):
        return ""

    text = str(value).strip()

    if text.lower() in {"nan", "none"}:
        return ""

    return text


def normalize_header(value):
    """
    Normalize a column header for case-insensitive matching.
    """
    return normalize_text(value).casefold()


def normalize_currency(value):
    """
    Normalize a currency code.
    """
    return normalize_text(value).upper()


def normalize_identifier(value):
    """
    Normalize identifiers read by pandas as numeric values.

    Examples:
    - 60000000054.0 -> 60000000054
    - 0.0 -> 0
    """
    text = normalize_text(value)

    if text.endswith(".0"):
        text = text[:-2]

    return text


def normalize_company(value):
    """
    Normalize company code for output and comparison.
    """
    return normalize_identifier(value)


def parse_date(value):
    """
    Parse Excel, ISO and day-first date values.
    """
    if value is None or pd.isna(value):
        return pd.NaT

    if isinstance(value, pd.Timestamp):
        return value

    text = normalize_text(value)

    if text == "":
        return pd.NaT

    if (
        len(text) >= 10
        and text[4:5] == "-"
        and text[7:8] == "-"
    ):
        return pd.to_datetime(
            text,
            errors="coerce",
            dayfirst=False,
        )

    return pd.to_datetime(
        text,
        errors="coerce",
        dayfirst=True,
    )


def get_analysis_date(context):
    """
    Return the FAM analysis end date.
    """
    raw_date = context["module"].get("to", "")
    analysis_date = parse_date(raw_date)

    if pd.isna(analysis_date):
        raise ValueError(
            f"FAM_002 could not parse module TO date: {raw_date}"
        )

    return analysis_date.normalize()


def get_period_suffix(context):
    """
    Return YYYYMMDD from the configured analysis date.
    """
    return get_analysis_date(context).strftime("%Y%m%d")


def get_input_folder(context):
    """
    Return the configured input folder as a Path.
    """
    return Path(context["input_folder"])


def get_ar_input_file(context):
    """
    Return the expected FAM AR input file.
    """
    input_file = (
        get_input_folder(context)
        / f"LBR FAM_AR_{get_period_suffix(context)}.xlsx"
    )

    if not input_file.exists():
        raise FileNotFoundError(
            "FAM AR input file not found. "
            f"Expected: {input_file}"
        )

    return input_file


def get_fx_input_file(context):
    """
    Return the expected FX input file.
    """
    input_file = (
        get_input_folder(context)
        / f"FxRates_{get_period_suffix(context)}.xlsx"
    )

    if not input_file.exists():
        raise FileNotFoundError(
            "FX rates input file not found. "
            f"Expected: {input_file}"
        )

    return input_file


def find_column(dataframe, aliases):
    """
    Find a dataframe column from a list of aliases.
    """
    normalized_columns = {
        normalize_header(column): column
        for column in dataframe.columns
    }

    for alias in aliases:
        normalized_alias = normalize_header(alias)

        if normalized_alias in normalized_columns:
            return normalized_columns[normalized_alias]

    return None


def resolve_required_columns(
    dataframe,
    required_column_definitions,
    source_name,
):
    """
    Resolve required source columns and raise a clear error when missing.
    """
    resolved = {}
    missing = []

    for logical_name, aliases in required_column_definitions.items():
        source_column = find_column(
            dataframe,
            aliases,
        )

        if source_column is None:
            missing.append(
                f"{logical_name}: {aliases}"
            )
        else:
            resolved[logical_name] = source_column

    if missing:
        raise ValueError(
            f"{source_name} is missing required columns. "
            f"Missing definitions: {missing}. "
            f"Available columns: {list(dataframe.columns)}"
        )

    return resolved


def read_ar_dataframe(context):
    """
    Read the FAM AR file directly.
    """
    input_file = get_ar_input_file(context)

    dataframe = pd.read_excel(
        input_file,
        dtype=object,
    )

    dataframe = dataframe.dropna(
        axis=0,
        how="all",
    )

    dataframe = dataframe.dropna(
        axis=1,
        how="all",
    )

    print(f"FAM_002 AR input file: {input_file}")
    print(f"FAM_002 AR rows loaded: {len(dataframe)}")

    return dataframe


def read_fx_dataframe(context):
    """
    Read and normalize the SAP FX file directly.
    """
    input_file = get_fx_input_file(context)

    raw_dataframe = pd.read_excel(
        input_file,
        dtype=object,
    )

    raw_dataframe = raw_dataframe.dropna(
        axis=0,
        how="all",
    )

    raw_dataframe = raw_dataframe.dropna(
        axis=1,
        how="all",
    )

    columns = resolve_required_columns(
        raw_dataframe,
        FX_REQUIRED_COLUMNS,
        "FX rates file",
    )

    fx_dataframe = pd.DataFrame(
        index=raw_dataframe.index
    )

    fx_dataframe["rate_type"] = (
        raw_dataframe[columns["rate_type"]]
        .apply(normalize_text)
        .str.upper()
    )

    fx_dataframe["from_currency"] = (
        raw_dataframe[columns["from_currency"]]
        .apply(normalize_currency)
    )

    fx_dataframe["to_currency"] = (
        raw_dataframe[columns["to_currency"]]
        .apply(normalize_currency)
    )

    fx_dataframe["valid_from"] = (
        raw_dataframe[columns["valid_from"]]
        .apply(parse_date)
    )

    fx_dataframe["raw_rate"] = (
        raw_dataframe[columns["raw_rate"]]
        .apply(parse_number)
    )

    fx_dataframe["factor_from"] = (
        raw_dataframe[columns["factor_from"]]
        .apply(parse_number)
    )

    fx_dataframe["factor_to"] = (
        raw_dataframe[columns["factor_to"]]
        .apply(parse_number)
    )

    fx_dataframe = fx_dataframe.dropna(
        subset=[
            "valid_from",
            "raw_rate",
        ]
    ).copy()

    print(f"FAM_002 FX input file: {input_file}")
    print(f"FAM_002 FX rows loaded: {len(fx_dataframe)}")

    return fx_dataframe


def calculate_adjusted_fx_rate(
    raw_rate,
    factor_from,
    factor_to,
):
    """
    Apply SAP FX factors.

    Rules:
    - 0/0 and 1/1 use the raw rate.
    - Otherwise: raw rate * factor_to / factor_from.
    """
    raw_rate = parse_number(raw_rate)
    factor_from = parse_number(factor_from)
    factor_to = parse_number(factor_to)

    if pd.isna(raw_rate):
        return pd.NA

    raw_rate = float(raw_rate)

    if raw_rate == 0:
        return pd.NA

    if pd.isna(factor_from):
        factor_from = 0.0
    else:
        factor_from = float(factor_from)

    if pd.isna(factor_to):
        factor_to = 0.0
    else:
        factor_to = float(factor_to)

    if (
        factor_from == 0
        and factor_to == 0
    ):
        return raw_rate

    if (
        factor_from == 1
        and factor_to == 1
    ):
        return raw_rate

    if factor_from == 0:
        return pd.NA

    return (
        raw_rate
        * factor_to
        / factor_from
    )


def select_fx_rate(
    fx_dataframe,
    currency,
    analysis_date,
):
    """
    Select an FX rate and return USD per unit of local currency.

    Search order:
    1. EN rate type, then M.
    2. Exact date, then up to 10 previous days.
    3. Currency -> USD.
    4. USD -> Currency, inverted once.
    """
    currency = normalize_currency(currency)

    if currency == "":
        return {
            "status": "RATE_NOT_FOUND",
            "direction": "",
            "raw_rate": pd.NA,
            "applied_rate": pd.NA,
            "date": pd.NaT,
        }

    if currency == "USD":
        return {
            "status": "OK",
            "direction": "USD->USD",
            "raw_rate": 1.0,
            "applied_rate": 1.0,
            "date": analysis_date,
        }

    for rate_type in PREFERRED_EXCHANGE_RATE_TYPES:
        for days_back in range(
            MAX_PREVIOUS_FX_DAYS + 1
        ):
            candidate_date = (
                analysis_date
                - pd.Timedelta(days=days_back)
            )

            candidates = fx_dataframe[
                (
                    fx_dataframe["rate_type"]
                    == rate_type
                )
                & (
                    fx_dataframe["valid_from"]
                    == candidate_date
                )
            ]

            if candidates.empty:
                continue

            direct_rows = candidates[
                (
                    candidates["from_currency"]
                    == currency
                )
                & (
                    candidates["to_currency"]
                    == "USD"
                )
            ]

            if not direct_rows.empty:
                row = direct_rows.iloc[0]

                adjusted_rate = calculate_adjusted_fx_rate(
                    row["raw_rate"],
                    row["factor_from"],
                    row["factor_to"],
                )

                if (
                    pd.isna(adjusted_rate)
                    or float(adjusted_rate) == 0
                ):
                    return {
                        "status": "INVALID_RATE",
                        "direction": f"{currency}->USD",
                        "raw_rate": row["raw_rate"],
                        "applied_rate": pd.NA,
                        "date": row["valid_from"],
                    }

                return {
                    "status": "OK",
                    "direction": f"{currency}->USD",
                    "raw_rate": float(row["raw_rate"]),
                    "applied_rate": float(adjusted_rate),
                    "date": row["valid_from"],
                }

            inverse_rows = candidates[
                (
                    candidates["from_currency"]
                    == "USD"
                )
                & (
                    candidates["to_currency"]
                    == currency
                )
            ]

            if not inverse_rows.empty:
                row = inverse_rows.iloc[0]

                adjusted_rate = calculate_adjusted_fx_rate(
                    row["raw_rate"],
                    row["factor_from"],
                    row["factor_to"],
                )

                if (
                    pd.isna(adjusted_rate)
                    or float(adjusted_rate) == 0
                ):
                    return {
                        "status": "INVALID_RATE",
                        "direction": f"USD->{currency}",
                        "raw_rate": row["raw_rate"],
                        "applied_rate": pd.NA,
                        "date": row["valid_from"],
                    }

                return {
                    "status": "OK",
                    "direction": f"USD->{currency}",
                    "raw_rate": float(row["raw_rate"]),
                    "applied_rate": (
                        1 / float(adjusted_rate)
                    ),
                    "date": row["valid_from"],
                }

    return {
        "status": "RATE_NOT_FOUND",
        "direction": "",
        "raw_rate": pd.NA,
        "applied_rate": pd.NA,
        "date": pd.NaT,
    }


def build_fx_lookup(
    fx_dataframe,
    currencies,
    analysis_date,
):
    """
    Build an independent currency-to-USD lookup for FAM02.
    """
    lookup = {}

    normalized_currencies = sorted(
        {
            normalize_currency(currency)
            for currency in currencies
            if normalize_currency(currency) != ""
        }
    )

    for currency in normalized_currencies:
        lookup[currency] = select_fx_rate(
            fx_dataframe=fx_dataframe,
            currency=currency,
            analysis_date=analysis_date,
        )

    return lookup


def convert_to_usd(
    local_amount,
    fx_details,
):
    """
    Convert a local amount to USD using USD-per-local-currency rate.
    """
    amount = parse_number(local_amount)

    if pd.isna(amount):
        return pd.NA

    if fx_details["status"] != "OK":
        return pd.NA

    applied_rate = fx_details["applied_rate"]

    if (
        pd.isna(applied_rate)
        or float(applied_rate) == 0
    ):
        return pd.NA

    return (
        float(amount)
        * float(applied_rate)
    )


def get_tolerance(context):
    """
    Return positive PARAM1 or the default 0.01 tolerance.
    """
    raw_tolerance = (
        context
        .get("control", {})
        .get("param1", "")
    )

    tolerance = parse_number(raw_tolerance)

    if pd.isna(tolerance):
        return DEFAULT_TOLERANCE

    return abs(float(tolerance))


def classify_asset(
    acquisition_value,
    accumulated_depreciation,
    reported_book_value,
    expected_book_value,
    control_difference,
    tolerance,
):
    """
    Classify one asset.

    Exceptions take priority over informational classification.
    """
    required_values = [
        acquisition_value,
        accumulated_depreciation,
        reported_book_value,
    ]

    if any(
        pd.isna(value)
        for value in required_values
    ):
        return (
            "EXCEPTION",
            "INCOMPLETE_OR_INVALID_DATA",
        )

    exceptions = []

    if abs(control_difference) > tolerance:
        exceptions.append(
            "BOOK_VALUE_MISMATCH"
        )

    if (
        abs(accumulated_depreciation)
        > abs(acquisition_value) + tolerance
    ):
        exceptions.append(
            "OVER_DEPRECIATED"
        )

    if reported_book_value < -tolerance:
        exceptions.append(
            "NEGATIVE_BOOK_VALUE"
        )

    if (
        acquisition_value > tolerance
        and accumulated_depreciation > tolerance
    ):
        exceptions.append(
            "UNEXPECTED_DEPRECIATION_SIGN"
        )

    if exceptions:
        return (
            "EXCEPTION",
            "; ".join(exceptions),
        )

    if abs(reported_book_value) <= tolerance:
        return (
            "INFORMATIONAL",
            "ZERO_BOOK_VALUE_INFORMATIONAL",
        )

    return "", ""


def build_fam_002_dataframe(
    source_dataframe,
    source_columns,
    fx_lookup,
    tolerance,
):
    """
    Build FAM02 directly from the AR input.
    """
    records = []

    for _, source_row in source_dataframe.iterrows():
        acquisition_value = parse_number(
            source_row[
                source_columns["acquisition_value"]
            ]
        )

        accumulated_depreciation = parse_number(
            source_row[
                source_columns[
                    "accumulated_depreciation"
                ]
            ]
        )

        reported_book_value = parse_number(
            source_row[
                source_columns["book_value"]
            ]
        )

        if (
            pd.isna(acquisition_value)
            or pd.isna(accumulated_depreciation)
        ):
            expected_book_value = pd.NA
        else:
            expected_book_value = (
                float(acquisition_value)
                + float(accumulated_depreciation)
            )

        if (
            pd.isna(reported_book_value)
            or pd.isna(expected_book_value)
        ):
            control_difference = pd.NA
        else:
            control_difference = (
                float(reported_book_value)
                - float(expected_book_value)
            )

        control_result, classification = classify_asset(
            acquisition_value=acquisition_value,
            accumulated_depreciation=(
                accumulated_depreciation
            ),
            reported_book_value=reported_book_value,
            expected_book_value=expected_book_value,
            control_difference=control_difference,
            tolerance=tolerance,
        )

        if classification == "":
            continue

        currency = normalize_currency(
            source_row[
                source_columns["currency"]
            ]
        )

        fx_details = fx_lookup.get(
            currency,
            {
                "status": "RATE_NOT_FOUND",
                "direction": "",
                "raw_rate": pd.NA,
                "applied_rate": pd.NA,
                "date": pd.NaT,
            },
        )

        acquisition_value_usd = convert_to_usd(
            acquisition_value,
            fx_details,
        )

        accumulated_depreciation_usd = convert_to_usd(
            accumulated_depreciation,
            fx_details,
        )

        expected_book_value_usd = convert_to_usd(
            expected_book_value,
            fx_details,
        )

        reported_book_value_usd = convert_to_usd(
            reported_book_value,
            fx_details,
        )

        control_difference_usd = convert_to_usd(
            control_difference,
            fx_details,
        )

        records.append(
            {
                CONTROL_RESULT_COLUMN: control_result,
                CLASSIFICATION_COLUMN: classification,
                ACQUISITION_VALUE_LC_COLUMN: acquisition_value,
                ACCUMULATED_DEPRECIATION_LC_COLUMN: (
                    accumulated_depreciation
                ),
                EXPECTED_BOOK_VALUE_LC_COLUMN: (
                    expected_book_value
                ),
                REPORTED_BOOK_VALUE_LC_COLUMN: (
                    reported_book_value
                ),
                CONTROL_DIFFERENCE_LC_COLUMN: (
                    control_difference
                ),
                "CoCo": normalize_company(
                    source_row[
                        source_columns["company_code"]
                    ]
                ),
                "Núm.AF": normalize_identifier(
                    source_row[
                        source_columns["asset_number"]
                    ]
                ),
                "Subnº": normalize_identifier(
                    source_row[
                        source_columns["asset_subnumber"]
                    ]
                ),
                "Descripción de activo": normalize_text(
                    source_row[
                        source_columns["asset_description"]
                    ]
                ),
                "Clase AF": normalize_text(
                    source_row[
                        source_columns["asset_class"]
                    ]
                ),
                "Fecha capitalización": parse_date(
                    source_row[
                        source_columns["capitalization_date"]
                    ]
                ),
                "Fecha inicio depreciación": parse_date(
                    source_row[
                        source_columns[
                            "depreciation_start_date"
                        ]
                    ]
                ),
                "Cuenta de balance": normalize_identifier(
                    source_row[
                        source_columns["balance_account"]
                    ]
                ),
                "Moneda": currency,
                FX_ORIGINAL_DIRECTION_COLUMN: (
                    fx_details["direction"]
                ),
                FX_ORIGINAL_RATE_COLUMN: (
                    fx_details["raw_rate"]
                ),
                FX_APPLIED_RATE_COLUMN: (
                    fx_details["applied_rate"]
                ),
                FX_DATE_COLUMN: fx_details["date"],
                FX_STATUS_COLUMN: fx_details["status"],
                ACQUISITION_VALUE_USD_COLUMN: (
                    acquisition_value_usd
                ),
                ACCUMULATED_DEPRECIATION_USD_COLUMN: (
                    accumulated_depreciation_usd
                ),
                EXPECTED_BOOK_VALUE_USD_COLUMN: (
                    expected_book_value_usd
                ),
                REPORTED_BOOK_VALUE_USD_COLUMN: (
                    reported_book_value_usd
                ),
                CONTROL_DIFFERENCE_USD_COLUMN: (
                    control_difference_usd
                ),
            }
        )

    return pd.DataFrame(
        records,
        columns=OUTPUT_COLUMNS,
    )


def remove_summary_rows(
    dataframe,
    source_columns,
):
    """
    Remove AR subtotal and total rows.
    """
    has_company = dataframe[
        source_columns["company_code"]
    ].apply(
        lambda value: normalize_text(value) != ""
    )

    has_asset = dataframe[
        source_columns["asset_number"]
    ].apply(
        lambda value: normalize_identifier(value) != ""
    )

    has_description = dataframe[
        source_columns["asset_description"]
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

def build_fam_002_summary(
    total_assets,
    detail_dataframe,
    analysis_date,
    tolerance,
):
    """
    Build the FAM02 executive summary.

    FAM02 contains only control totals and conclusions.
    The asset-level population is written to FAM02_Detail.
    """
    if detail_dataframe.empty:
        exception_count = 0
        informational_count = 0
        zero_book_value_count = 0
        fx_ok_count = 0
        fx_not_found_count = 0
        fx_invalid_count = 0
        classification_counts = pd.Series(dtype="int64")
    else:
        exception_count = int(
            detail_dataframe[
                CONTROL_RESULT_COLUMN
            ]
            .eq("EXCEPTION")
            .sum()
        )

        informational_count = int(
            detail_dataframe[
                CONTROL_RESULT_COLUMN
            ]
            .eq("INFORMATIONAL")
            .sum()
        )

        zero_book_value_count = int(
            detail_dataframe[
                CLASSIFICATION_COLUMN
            ]
            .eq("ZERO_BOOK_VALUE_INFORMATIONAL")
            .sum()
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

        classification_counts = (
            detail_dataframe[
                CLASSIFICATION_COLUMN
            ]
            .fillna("")
            .astype(str)
            .str.split("; ")
            .explode()
        )

        classification_counts = (
            classification_counts[
                classification_counts != ""
            ]
            .value_counts()
        )

    selected_count = len(detail_dataframe)

    if exception_count == 0:
        overall_result = (
            "No accounting exceptions identified"
        )
    else:
        overall_result = (
            f"{exception_count} accounting exception(s) identified"
        )

    summary_rows = [
        {
            "Metric": "Analysis cut-off date",
            "Result": analysis_date.date(),
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
            "Metric": "Records included in FAM02_Detail",
            "Result": selected_count,
        },
        {
            "Metric": "Zero-NBV assets correctly reconciled",
            "Result": zero_book_value_count,
        },
        {
            "Metric": "Accounting exceptions",
            "Result": exception_count,
        },
        {
            "Metric": "Informational records",
            "Result": informational_count,
        },
        {
            "Metric": "Successful FX conversions",
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

    summary_dataframe = pd.DataFrame(
        summary_rows,
        columns=[
            "Metric",
            "Result",
        ],
    )

    classification_rows = [
        {
            "Classification": classification,
            "Count": int(count),
        }
        for classification, count
        in classification_counts.items()
    ]

    classification_dataframe = pd.DataFrame(
        classification_rows,
        columns=[
            "Classification",
            "Count",
        ],
    )

    return (
        summary_dataframe,
        classification_dataframe,
    )


def write_fam_002_summary(
    worksheet,
    summary_dataframe,
    classification_dataframe,
):
    """
    Write the executive summary and classification totals to FAM02.
    """
    worksheet.cell(
        row=1,
        column=1,
        value="FAM02 - Fixed Assets Control Summary",
    )

    worksheet.cell(
        row=2,
        column=1,
        value=(
            "Reconciliation of acquisition value, "
            "accumulated depreciation and net book value"
        ),
    )

    summary_header_row = 4

    worksheet.cell(
        row=summary_header_row,
        column=1,
        value="Metric",
    )

    worksheet.cell(
        row=summary_header_row,
        column=2,
        value="Result",
    )

    for row_offset, row in enumerate(
        summary_dataframe.itertuples(index=False),
        start=1,
    ):
        worksheet.cell(
            row=summary_header_row + row_offset,
            column=1,
            value=row[0],
        )

        worksheet.cell(
            row=summary_header_row + row_offset,
            column=2,
            value=row[1],
        )

    classification_start_row = (
        summary_header_row
        + len(summary_dataframe)
        + 3
    )

    worksheet.cell(
        row=classification_start_row,
        column=1,
        value="Classification summary",
    )

    worksheet.cell(
        row=classification_start_row + 1,
        column=1,
        value="Classification",
    )

    worksheet.cell(
        row=classification_start_row + 1,
        column=2,
        value="Count",
    )

    if classification_dataframe.empty:
        worksheet.cell(
            row=classification_start_row + 2,
            column=1,
            value="Sin registros seleccionados",
        )

        worksheet.cell(
            row=classification_start_row + 2,
            column=2,
            value=0,
        )
    else:
        for row_offset, row in enumerate(
            classification_dataframe.itertuples(index=False),
            start=2,
        ):
            worksheet.cell(
                row=classification_start_row + row_offset,
                column=1,
                value=row[0],
            )

            worksheet.cell(
                row=classification_start_row + row_offset,
                column=2,
                value=row[1],
            )

    # Basic summary presentation.
    worksheet.freeze_panes = "A5"
    worksheet.column_dimensions["A"].width = 52
    worksheet.column_dimensions["B"].width = 38

    # Title.
    worksheet["A1"].font = worksheet["A1"].font.copy(
        bold=True,
        size=14,
    )

    # Main summary header.
    for cell in worksheet[summary_header_row]:
        if cell.column <= 2:
            cell.font = cell.font.copy(
                bold=True,
            )

    # Classification title and header.
    worksheet.cell(
        row=classification_start_row,
        column=1,
    ).font = worksheet.cell(
        row=classification_start_row,
        column=1,
    ).font.copy(
        bold=True,
        size=12,
    )

    for column_index in range(1, 3):
        worksheet.cell(
            row=classification_start_row + 1,
            column=column_index,
        ).font = worksheet.cell(
            row=classification_start_row + 1,
            column=column_index,
        ).font.copy(
            bold=True,
        )

    # Date format for the cut-off date.
    worksheet.cell(
        row=summary_header_row + 1,
        column=2,
    ).number_format = "dd/mm/yyyy"

    # Integer format for population totals.
    for row_index in range(
        summary_header_row + 3,
        summary_header_row + 11,
    ):
        worksheet.cell(
            row=row_index,
            column=2,
        ).number_format = "#,##0"

    for row_index in range(
        classification_start_row + 2,
        classification_start_row + 2
        + max(len(classification_dataframe), 1),
    ):
        worksheet.cell(
            row=row_index,
            column=2,
        ).number_format = "#,##0"


def run_fam_002(context):
    """
    Run the independent FAM02 analysis.

    Output sheets:
    - FAM02: executive summary.
    - FAM02_Detail: selected asset-level records.

    No other workbook sheets are removed or modified.
    """
    summary_sheet_name = "FAM02"
    detail_sheet_name = "FAM02_Detail"

    analysis_date = get_analysis_date(context)
    tolerance = get_tolerance(context)

    source_dataframe = read_ar_dataframe(context)

    source_columns = resolve_required_columns(
        source_dataframe,
        AR_REQUIRED_COLUMNS,
        "FAM AR file",
    )

    source_dataframe = remove_summary_rows(
        dataframe=source_dataframe,
        source_columns=source_columns,
    )

    source_dataframe = filter_by_company(
        dataframe=source_dataframe,
        company_column=(
            source_columns["company_code"]
        ),
        companies_filter=(
            context["module"].get(
                "companies",
                "",
            )
        ),
    )

    total_assets = len(source_dataframe)

    fx_dataframe = read_fx_dataframe(context)

    currency_values = source_dataframe[
        source_columns["currency"]
    ]

    fx_lookup = build_fx_lookup(
        fx_dataframe=fx_dataframe,
        currencies=currency_values,
        analysis_date=analysis_date,
    )

    internal_detail_dataframe = build_fam_002_dataframe(
        source_dataframe=source_dataframe,
        source_columns=source_columns,
        fx_lookup=fx_lookup,
        tolerance=tolerance,
    )

    (
        summary_dataframe,
        classification_dataframe,
    ) = build_fam_002_summary(
        total_assets=total_assets,
        detail_dataframe=internal_detail_dataframe,
        analysis_date=analysis_date,
        tolerance=tolerance,
    )

    # Resultado del control is used internally for summary counts.
    # It is removed from the asset-level output because Clasificación
    # already communicates whether the row is informational or exceptional.
    detail_dataframe = (
        internal_detail_dataframe
        .drop(
            columns=[
                CONTROL_RESULT_COLUMN,
            ],
            errors="ignore",
        )
        .copy()
    )

    output_file = get_fam_output_file(context)

    workbook = open_or_create_fam_output_workbook(
        output_file
    )

    summary_worksheet = recreate_fam_sheet(
        workbook,
        summary_sheet_name,
    )

    detail_worksheet = recreate_fam_sheet(
        workbook,
        detail_sheet_name,
    )

    write_fam_002_summary(
        worksheet=summary_worksheet,
        summary_dataframe=summary_dataframe,
        classification_dataframe=classification_dataframe,
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

    # Display FX rates with sufficient precision for audit recalculation.
    detail_column_positions = {
        column_name: column_index
        for column_index, column_name
        in enumerate(
            detail_dataframe.columns,
            start=1,
        )
    }

    for fx_column in [
        FX_ORIGINAL_RATE_COLUMN,
        FX_APPLIED_RATE_COLUMN,
    ]:
        column_index = detail_column_positions.get(
            fx_column
        )

        if column_index is None:
            continue

        for row_index in range(
            2,
            len(detail_dataframe) + 2,
        ):
            detail_worksheet.cell(
                row=row_index,
                column=column_index,
            ).number_format = "0.00000000"

    save_fam_output_workbook(
        workbook=workbook,
        output_file=output_file,
    )

    exception_count = int(
        internal_detail_dataframe[
            CONTROL_RESULT_COLUMN
        ]
        .eq("EXCEPTION")
        .sum()
    )

    informational_count = int(
        internal_detail_dataframe[
            CONTROL_RESULT_COLUMN
        ]
        .eq("INFORMATIONAL")
        .sum()
    )

    print(f"FAM_002 output file: {output_file}")
    print(f"FAM_002 summary sheet: {summary_sheet_name}")
    print(f"FAM_002 detail sheet: {detail_sheet_name}")
    print(f"FAM_002 tolerance: {tolerance}")
    print(f"FAM_002 assets analyzed: {total_assets}")
    print(
        f"FAM_002 selected rows: "
        f"{len(internal_detail_dataframe)}"
    )
    print(
        f"FAM_002 exception rows: "
        f"{exception_count}"
    )
    print(
        f"FAM_002 informational rows: "
        f"{informational_count}"
    )

