"""
FAM_001 - Fixed Assets Listing.

Analysis:
- Module: Fixed Asset Management
- Analysis Code: AM_ANALYTIC_01_AMCS101
- Analysis Title: Fixed Assets Listing

Description:
Extracts a current listing of fixed assets.

Procedure:
Produce current list of all assets.

Analytic Logic:
Extract full asset register as of extract date.

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
    FAM01

Rules:
- This control writes/replaces only sheet FAM01.
- It does not delete sheets from other controls.
"""

import pandas as pd

from core.fam_common import (
    apply_standard_fam_formatting,
    build_fx_rate_lookup,
    filter_by_company,
    get_fam_output_file,
    get_optional_column,
    load_fam_ar01_data,
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


SHEET_NAME = "FAM01"


REQUIRED_COLUMNS = {
    "asset_number": [
        "Imobilizado",
    ],
    "asset_subnumber": [
        "Subnº",
        "Subn°",
        "Subno",
        "Subnro",
    ],
    "company_code": [
        "Empresa",
    ],
    "asset_description": [
        "Denominação do imobilizado",
        "Denominacao do imobilizado",
    ],
    "capitalization_date": [
        "Incorporação em",
        "Incorporacao em",
    ],
    "acquisition_value": [
        "ValAquis.",
        "ValAquis",
        "Valor aquisição",
        "Valor aquisicao",
    ],
    "accumulated_depreciation": [
        "Depreciação ac.",
        "Depreciacao ac.",
        "Depreciação acumulada",
        "Depreciacao acumulada",
    ],
    "book_value": [
        "Valor contábil",
        "Valor contabil",
    ],
    "asset_class": [
        "Classe imobilizado",
    ],
    "currency": [
        "Moeda",
    ],
}


OPTIONAL_COLUMNS = {
    "division": [
        "Divisão",
        "Divisao",
    ],
    "balance_item": [
        "Item de balanço",
        "Item de balanco",
    ],
    "balance_account_cap": [
        "Conta do balanço CAP",
        "Conta do balanco CAP",
    ],
    "normal_depreciation": [
        "Depreciação normal",
        "Depreciacao normal",
    ],
}


OUTPUT_COLUMNS = [
    "CoCo",
    "Company",
    "Rate USD (To)",
    "Fecha rate USD",
    "Cuenta de balance",
    "Nombre de cuenta",
    "Clase AF",
    "Núm.AF",
    "Descripción de activo",
    "CAP histórico",
    "CAP histórico (moneda)",
    "Fecha capitalización",
    "Vida útil",
    "Resto vida útil",
    "Cl.amortiz.",
    "CAP en fecha inicio (LC)",
    "CAP en fecha inicio (USD)",
    "CAP en fecha inicio (moneda)",
    "Amortiz.acumul.en fecha de inicio",
    "Amortiz.acumul.en fecha de inicio (moneda)",
    "Revaloración acumulada fecha inicio",
    "Revaloración acumulada fecha inicio (moneda)",
    "VNC en fecha inicio",
    "VNC en fecha inicio (moneda)",
    "Capitalización",
    "Capitalización (moneda)",
    "CAP retirado",
    "CAP retirado (moneda)",
    "VNC retirado",
    "VNC retirado (moneda)",
    "CAP transferidos",
    "CAP transferidos (moneda)",
    "VNC transferido",
    "VNC transferido (moneda)",
    "Revaloración",
    "Revaloración (moneda)",
    "Valoración",
    "Valoración (moneda)",
    "CAP en fecha de fin (LC)",
    "CAP en fecha de fin (USD)",
    "CAP en fecha de fin (moneda)",
    "VNC en fecha de fin (LC)",
    "VNC en fecha de fin (USD)",
    "VNC en fecha de fin (moneda)",
    "Depreciación en fecha de fin",
    "Depreciación en fecha de fin (moneda)",
    "Amortización acumulada en fecha fin (LC)",
    "Amortización acumulada en fecha fin (USD)",
    "Amortización acumulada en fecha fin (moneda)",
    "Subnº",
    "División",
    "Item de balance",
    "Fecha inicio depreciación",
]


DATE_COLUMNS = {
    "Fecha rate USD",
    "Fecha capitalización",
    "Fecha inicio depreciación",
}


AMOUNT_COLUMNS = {
    "Rate USD (To)",
    "CAP histórico",
    "CAP en fecha inicio (LC)",
    "CAP en fecha inicio (USD)",
    "Amortiz.acumul.en fecha de inicio",
    "Revaloración acumulada fecha inicio",
    "VNC en fecha inicio",
    "Capitalización",
    "CAP retirado",
    "VNC retirado",
    "CAP transferidos",
    "VNC transferido",
    "Revaloración",
    "Valoración",
    "CAP en fecha de fin (LC)",
    "CAP en fecha de fin (USD)",
    "VNC en fecha de fin (LC)",
    "VNC en fecha de fin (USD)",
    "Depreciación en fecha de fin",
    "Amortización acumulada en fecha fin (LC)",
    "Amortización acumulada en fecha fin (USD)",
}


INTEGER_COLUMNS = {
    "Vida útil",
    "Resto vida útil",
}


def resolve_optional_columns(source_dataframe):
    """
    Resolve optional columns.

    Missing optional columns are allowed and will be output as blank.
    """
    resolved_columns = {}

    for logical_name, possible_names in OPTIONAL_COLUMNS.items():
        resolved_columns[logical_name] = get_optional_column(
            source_dataframe,
            possible_names,
        )

    return resolved_columns


def blank_series(source_dataframe):
    """
    Return a blank series matching the dataframe index.
    """
    return pd.Series(
        [""] * len(source_dataframe),
        index=source_dataframe.index,
    )


def blank_date_series(source_dataframe):
    """
    Return a blank date series matching the dataframe index.
    """
    return pd.Series(
        [pd.NaT] * len(source_dataframe),
        index=source_dataframe.index,
    )


def blank_number_series(source_dataframe):
    """
    Return a blank numeric series matching the dataframe index.
    """
    return pd.Series(
        [pd.NA] * len(source_dataframe),
        index=source_dataframe.index,
    )


def zero_number_series(source_dataframe):
    """
    Return a zero numeric series matching the dataframe index.
    """
    return pd.Series(
        [0.0] * len(source_dataframe),
        index=source_dataframe.index,
    )


def get_text_series(source_dataframe, column_name):
    """
    Return normalized text series from column or blank series if missing.
    """
    if column_name is None:
        return blank_series(source_dataframe)

    return source_dataframe[column_name].apply(normalize_text)


def get_date_series(source_dataframe, column_name):
    """
    Return date series from column or blank date series if missing.
    """
    if column_name is None:
        return blank_date_series(source_dataframe)

    return source_dataframe[column_name].apply(to_datetime_value)


def normalize_asset_number(value):
    """
    Normalize asset number as text.

    Examples:
    - 60000000054.0 -> 60000000054
    - 60000000054 -> 60000000054
    """
    value_text = normalize_text(value)

    if value_text == "":
        return ""

    if value_text.endswith(".0"):
        value_text = value_text[:-2]

    return value_text


def normalize_account_number(value):
    """
    Normalize account values.

    Examples:
    - 1203010006.0 -> 1203010006
    - 1203010006 -> 1203010006
    """
    value_text = normalize_text(value)

    if value_text == "":
        return ""

    if value_text.endswith(".0"):
        value_text = value_text[:-2]

    return value_text


def get_asset_number_series(source_dataframe, column_name):
    """
    Return normalized asset number series.
    """
    return source_dataframe[column_name].apply(normalize_asset_number)


def remove_ar01_summary_rows(source_dataframe, required_columns):
    """
    Remove AR01 subtotal / total rows.

    AR01 exports can include summary rows at the end of the report.
    Those rows usually have amounts but do not have:
    - Empresa
    - Imobilizado
    - Denominação do imobilizado

    FAM01 must keep asset-level rows only.
    """
    result = source_dataframe.copy()

    company_column = required_columns["company_code"]
    asset_number_column = required_columns["asset_number"]
    asset_description_column = required_columns["asset_description"]

    has_company = result[company_column].apply(
        lambda value: normalize_text(value) != ""
    )
    has_asset_number = result[asset_number_column].apply(
        lambda value: normalize_asset_number(value) != ""
    )
    has_asset_description = result[asset_description_column].apply(
        lambda value: normalize_text(value) != ""
    )

    valid_asset_rows = has_company & has_asset_number & has_asset_description

    removed_rows = len(result) - int(valid_asset_rows.sum())

    if removed_rows > 0:
        print(f"FAM_001 AR01 summary/total rows removed: {removed_rows}")

    return result[valid_asset_rows].copy()


def get_fx_rate(currency, fx_lookup):
    """
    Return FX rate data for a currency.

    Rate meaning:
        local currency -> USD

    Formula:
        USD = LC * FxRate
    """
    currency = normalize_currency(currency)

    if currency == "":
        return None

    if currency in fx_lookup:
        return fx_lookup[currency]

    return None


def convert_lc_to_usd(value, currency, fx_lookup):
    """
    Convert local currency value to USD.

    Formula:
        USD = LC * FxRate

    FxRate follows the same convention used by CD:
        Currency -> USD
    """
    amount = value

    if pd.isna(amount):
        return pd.NA

    fx_rate_data = get_fx_rate(
        currency=currency,
        fx_lookup=fx_lookup,
    )

    if fx_rate_data is None:
        return pd.NA

    rate = fx_rate_data.get("rate", pd.NA)

    if pd.isna(rate):
        return pd.NA

    if float(rate) == 0:
        return pd.NA

    return float(amount) * float(rate)


def get_fx_rate_value(currency, fx_lookup):
    """
    Return rate value for a currency.
    """
    fx_rate_data = get_fx_rate(
        currency=currency,
        fx_lookup=fx_lookup,
    )

    if fx_rate_data is None:
        return pd.NA

    return fx_rate_data.get("rate", pd.NA)


def get_fx_rate_date(currency, fx_lookup):
    """
    Return rate date for a currency.
    """
    fx_rate_data = get_fx_rate(
        currency=currency,
        fx_lookup=fx_lookup,
    )

    if fx_rate_data is None:
        return pd.NaT

    return fx_rate_data.get("date", pd.NaT)


def build_fam_001_dataframe(source_dataframe, required_columns, optional_columns, fx_lookup):
    """
    Build the FAM01 fixed assets listing output dataframe.
    """
    output_dataframe = pd.DataFrame(index=source_dataframe.index)

    currency_series = source_dataframe[required_columns["currency"]].apply(normalize_currency)

    acquisition_value_lc = source_dataframe[
        required_columns["acquisition_value"]
    ].apply(parse_number)

    accumulated_depreciation_lc = source_dataframe[
        required_columns["accumulated_depreciation"]
    ].apply(parse_number)

    book_value_lc = source_dataframe[
        required_columns["book_value"]
    ].apply(parse_number)

    output_dataframe["CoCo"] = source_dataframe[required_columns["company_code"]].apply(
        normalize_company_output
    )
    output_dataframe["Company"] = ""
    output_dataframe["Rate USD (To)"] = currency_series.apply(
        lambda currency: get_fx_rate_value(currency, fx_lookup)
    )
    output_dataframe["Fecha rate USD"] = currency_series.apply(
        lambda currency: get_fx_rate_date(currency, fx_lookup)
    )
    output_dataframe["Cuenta de balance"] = get_text_series(
        source_dataframe,
        optional_columns["balance_account_cap"],
    ).apply(normalize_account_number)
    output_dataframe["Nombre de cuenta"] = ""
    output_dataframe["Clase AF"] = source_dataframe[required_columns["asset_class"]].apply(
        normalize_text
    )
    output_dataframe["Núm.AF"] = get_asset_number_series(
        source_dataframe,
        required_columns["asset_number"],
    )
    output_dataframe["Descripción de activo"] = source_dataframe[
        required_columns["asset_description"]
    ].apply(normalize_text)

    output_dataframe["CAP histórico"] = zero_number_series(source_dataframe)
    output_dataframe["CAP histórico (moneda)"] = currency_series

    output_dataframe["Fecha capitalización"] = source_dataframe[
        required_columns["capitalization_date"]
    ].apply(to_datetime_value)

    output_dataframe["Vida útil"] = blank_number_series(source_dataframe)
    output_dataframe["Resto vida útil"] = blank_number_series(source_dataframe)
    output_dataframe["Cl.amortiz."] = ""

    output_dataframe["CAP en fecha inicio (LC)"] = blank_number_series(source_dataframe)
    output_dataframe["CAP en fecha inicio (USD)"] = blank_number_series(source_dataframe)
    output_dataframe["CAP en fecha inicio (moneda)"] = currency_series

    output_dataframe["Amortiz.acumul.en fecha de inicio"] = blank_number_series(source_dataframe)
    output_dataframe["Amortiz.acumul.en fecha de inicio (moneda)"] = currency_series

    output_dataframe["Revaloración acumulada fecha inicio"] = zero_number_series(source_dataframe)
    output_dataframe["Revaloración acumulada fecha inicio (moneda)"] = currency_series

    output_dataframe["VNC en fecha inicio"] = blank_number_series(source_dataframe)
    output_dataframe["VNC en fecha inicio (moneda)"] = currency_series

    output_dataframe["Capitalización"] = zero_number_series(source_dataframe)
    output_dataframe["Capitalización (moneda)"] = currency_series

    output_dataframe["CAP retirado"] = zero_number_series(source_dataframe)
    output_dataframe["CAP retirado (moneda)"] = currency_series

    output_dataframe["VNC retirado"] = zero_number_series(source_dataframe)
    output_dataframe["VNC retirado (moneda)"] = currency_series

    output_dataframe["CAP transferidos"] = zero_number_series(source_dataframe)
    output_dataframe["CAP transferidos (moneda)"] = currency_series

    output_dataframe["VNC transferido"] = zero_number_series(source_dataframe)
    output_dataframe["VNC transferido (moneda)"] = currency_series

    output_dataframe["Revaloración"] = zero_number_series(source_dataframe)
    output_dataframe["Revaloración (moneda)"] = currency_series

    output_dataframe["Valoración"] = zero_number_series(source_dataframe)
    output_dataframe["Valoración (moneda)"] = currency_series

    output_dataframe["CAP en fecha de fin (LC)"] = acquisition_value_lc
    output_dataframe["CAP en fecha de fin (USD)"] = [
        convert_lc_to_usd(value, currency, fx_lookup)
        for value, currency in zip(acquisition_value_lc, currency_series)
    ]
    output_dataframe["CAP en fecha de fin (moneda)"] = currency_series

    output_dataframe["VNC en fecha de fin (LC)"] = book_value_lc
    output_dataframe["VNC en fecha de fin (USD)"] = [
        convert_lc_to_usd(value, currency, fx_lookup)
        for value, currency in zip(book_value_lc, currency_series)
    ]
    output_dataframe["VNC en fecha de fin (moneda)"] = currency_series

    output_dataframe["Depreciación en fecha de fin"] = blank_number_series(source_dataframe)
    output_dataframe["Depreciación en fecha de fin (moneda)"] = currency_series

    output_dataframe["Amortización acumulada en fecha fin (LC)"] = accumulated_depreciation_lc
    output_dataframe["Amortización acumulada en fecha fin (USD)"] = [
        convert_lc_to_usd(value, currency, fx_lookup)
        for value, currency in zip(accumulated_depreciation_lc, currency_series)
    ]
    output_dataframe["Amortización acumulada en fecha fin (moneda)"] = currency_series

    output_dataframe["Subnº"] = get_asset_number_series(
        source_dataframe,
        required_columns["asset_subnumber"],
    )
    output_dataframe["División"] = get_text_series(
        source_dataframe,
        optional_columns["division"],
    )
    output_dataframe["Item de balance"] = get_text_series(
        source_dataframe,
        optional_columns["balance_item"],
    )
    output_dataframe["Fecha inicio depreciación"] = get_date_series(
        source_dataframe,
        optional_columns["normal_depreciation"],
    )

    output_dataframe = output_dataframe[OUTPUT_COLUMNS]

    return output_dataframe


def run_fam_001(context):
    """
    Run FAM_001 and write the FAM01 sheet.
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

    output_dataframe = build_fam_001_dataframe(
        source_dataframe=source_dataframe,
        required_columns=required_columns,
        optional_columns=optional_columns,
        fx_lookup=fx_lookup,
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

    print(f"FAM_001 output file: {output_file}")
    print(f"FAM_001 sheet: {SHEET_NAME}")
    print(f"FAM_001 rows: {len(output_dataframe)}")
