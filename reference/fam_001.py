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
- Optional SQVI balance support saved as:
    input/LBR_FAM_BAL__YYYYMMDD.xlsx
- Optional SQVI movement support saved as:
    input/LBR_FAM_MOV_YYYYMMDD.xlsx

Output:
- Workbook:
    output/LBR_Results_FAM_YYYYMMDD.xlsx
- Sheet:
    FAM01

Rules:
- This control writes/replaces only sheet FAM01.
- It does not delete sheets from other controls.
- AR01 remains the primary source; BAL/MOV files enrich missing metrics only.
"""

import re
import unicodedata
from pathlib import Path

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
DEFAULT_DEPRECIATION_AREA = "01"
FAM01_BALANCE_FILE_KEYWORD = "BAL"
FAM01_MOVEMENT_FILE_KEYWORD = "MOV"


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


SUPPORT_COLUMN_ALIASES = {
    "EMPR": "BUKRS",
    "EMPRESA": "BUKRS",
    "BUKRS": "BUKRS",
    "IMOBILIZADO": "ANLN1",
    "ANLN1": "ANLN1",
    "SBN": "ANLN2",
    "SUBNUMERO": "ANLN2",
    "ANLN2": "ANLN2",
    "CLASSE": "ANLKL",
    "CLASE": "ANLKL",
    "ANLKL": "ANLKL",
    "DENOMINACAO DO IMOBILIZADO": "TXT50",
    "DENOMINACION DEL ACTIVO FIJO": "TXT50",
    "TXT50": "TXT50",
    "AR": "AFABE",
    "AREA": "AFABE",
    "AFABE": "AFABE",
    "ANO": "GJAHR",
    "ANIO": "GJAHR",
    "EJERCICIO": "GJAHR",
    "GJAHR": "GJAHR",
    "VAL AQUIS ACUM": "KANSW",
    "VALOR AQUISICAO ACUM": "KANSW",
    "VALOR ADQUISICION ACUM": "KANSW",
    "KANSW": "KANSW",
    "DEPR NORM ACUM": "KNAFA",
    "DEPRECIACION NORMAL ACUM": "KNAFA",
    "KNAFA": "KNAFA",
    "DEPR ESPE ACUM": "KSAFA",
    "DEPRECIACION ESPECIAL ACUM": "KSAFA",
    "KSAFA": "KSAFA",
    "DEPR EXTR ACUM": "KAAFA",
    "DEPRECIACION EXTRAORD ACUM": "KAAFA",
    "KAAFA": "KAAFA",
    "RESERVA ACUM": "KMAFA",
    "KMAFA": "KMAFA",
    "N DOC": "BELNR",
    "NUM DOC": "BELNR",
    "DOCUMENTO": "BELNR",
    "BELNR": "BELNR",
    "ITM": "BUZEI",
    "ITEM": "BUZEI",
    "BUZEI": "BUZEI",
    "TIPO DE MOVIMENTO": "BWASL",
    "TIPO DE MOVIMIENTO": "BWASL",
    "BWASL": "BWASL",
    "DATA REF": "BZDAT",
    "FECHA REF": "BZDAT",
    "BZDAT": "BZDAT",
    "MONTANTE LANCADO": "ANBTR",
    "IMPORTE CONTABILIZADO": "ANBTR",
    "ANBTR": "ANBTR",
    "DEPR NORM MOV": "NAFAB",
    "DEPRECIACION NORMAL MOV": "NAFAB",
    "NAFAB": "NAFAB",
    "DEPR ESPE MOV": "SAFAB",
    "DEPRECIACION ESPECIAL MOV": "SAFAB",
    "SAFAB": "SAFAB",
}


BALANCE_REQUIRED_COLUMNS = {
    "company_code": "BUKRS",
    "asset_number": "ANLN1",
    "asset_subnumber": "ANLN2",
    "depreciation_area": "AFABE",
    "fiscal_year": "GJAHR",
    "apc_opening": "KANSW",
}

BALANCE_DEPRECIATION_COLUMNS = [
    "KNAFA",
    "KSAFA",
    "KAAFA",
    "KMAFA",
]

MOVEMENT_REQUIRED_COLUMNS = {
    "company_code": "BUKRS",
    "asset_number": "ANLN1",
    "asset_subnumber": "ANLN2",
    "depreciation_area": "AFABE",
    "fiscal_year": "GJAHR",
    "transaction_type": "BWASL",
    "amount": "ANBTR",
}

MOVEMENT_DEPRECIATION_COLUMNS = [
    "NAFAB",
    "SAFAB",
]

CAPITALIZATION_TRANSACTION_PREFIXES = ("1",)
VALUATION_TRANSACTION_PREFIXES = ("7", "8")


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


def normalize_lookup_text(value):
    """
    Normalize text for tolerant SQVI filename and header matching.
    """
    value_text = normalize_text(value).upper()
    value_text = unicodedata.normalize("NFKD", value_text)
    value_text = "".join(character for character in value_text if not unicodedata.combining(character))

    return re.sub(r"[^A-Z0-9]+", " ", value_text).strip()


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


def normalize_company_code(value):
    """
    Normalize company code for matching FAM AR01 and SQVI support files.
    """
    value_text = normalize_text(value).upper()

    if value_text.endswith(".0"):
        value_text = value_text[:-2]

    if value_text.isdigit():
        return str(int(value_text))

    return value_text


def normalize_depreciation_area(value):
    """
    Normalize depreciation area values such as 1, 01 or 1.0 to 2 digits.
    """
    value_text = normalize_text(value)

    if value_text.endswith(".0"):
        value_text = value_text[:-2]

    if value_text.isdigit():
        return value_text.zfill(2)

    return value_text


def normalize_support_column_name(column):
    """
    Normalize SQVI column headers to their SAP technical names when needed.
    """
    column_text = normalize_text(column).upper()
    column_key = normalize_lookup_text(column_text)

    if column_text in SUPPORT_COLUMN_ALIASES:
        return SUPPORT_COLUMN_ALIASES[column_text]

    if column_key in SUPPORT_COLUMN_ALIASES:
        return SUPPORT_COLUMN_ALIASES[column_key]

    return column_text


def normalize_support_columns(dataframe):
    """
    Rename Portuguese/Spanish SQVI headers to SAP technical field names.
    """
    normalized_columns = []
    used_columns = set()

    for column in dataframe.columns:
        normalized_column = normalize_support_column_name(column)

        if normalized_column in used_columns:
            suffix = 1
            unique_column = f"{normalized_column}_{suffix}"
            while unique_column in used_columns:
                suffix += 1
                unique_column = f"{normalized_column}_{suffix}"
            normalized_column = unique_column

        normalized_columns.append(normalized_column)
        used_columns.add(normalized_column)

    dataframe = dataframe.copy()
    dataframe.columns = normalized_columns

    return dataframe


def get_analysis_date_to(context):
    """
    Return module TO date as a Python date.
    """
    raw_date = context["module"].get("to", "")
    parsed_date = pd.to_datetime(raw_date, errors="coerce")

    if pd.isna(parsed_date):
        raise ValueError(f"FAM_001 could not parse module TO date: {raw_date}")

    return parsed_date.date()


def get_analysis_date_from(context):
    """
    Return module FROM date as a Python date when configured.
    """
    raw_date = context["module"].get("from", "")
    parsed_date = pd.to_datetime(raw_date, errors="coerce")

    if pd.isna(parsed_date):
        return None

    return parsed_date.date()


def get_period_suffix(context):
    """
    Return YYYYMMDD suffix from the module TO date.
    """
    return get_analysis_date_to(context).strftime("%Y%m%d")


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


def find_fam01_support_file(context, keyword):
    """
    Find optional SQVI support files saved in input/ for FAM01 enrichment.

    Expected patterns:
    - LBR_FAM_BAL__YYYYMMDD.xlsx
    - LBR_FAM_MOV_YYYYMMDD.xlsx
    """
    input_folder = Path(context["input_folder"])
    period_suffix = get_period_suffix(context)
    keyword_normalized = normalize_lookup_text(keyword)

    candidates = []
    for file_path in input_folder.glob("*.xls*"):
        if file_path.name.startswith("~$"):
            continue
        if file_path.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
            continue

        normalized_name = normalize_lookup_text(file_path.stem)
        normalized_tokens = normalized_name.split()
        compact_name = normalized_name.replace(" ", "")

        if "FAM" not in normalized_tokens:
            continue
        if keyword_normalized not in normalized_tokens:
            continue
        if period_suffix not in compact_name:
            continue

        candidates.append(file_path)

    if not candidates:
        return None

    return sorted(candidates, key=lambda path: (len(path.name), path.name.lower()))[0]


def read_optional_support_file(context, keyword):
    """
    Read an optional FAM SQVI support workbook if it exists.
    """
    support_file = find_fam01_support_file(context, keyword)

    if support_file is None:
        print(f"FAM_001 support file {keyword}: not found in input/ (optional)")
        return pd.DataFrame()

    dataframe = pd.read_excel(support_file)
    dataframe = dataframe.dropna(axis=0, how="all")
    dataframe = dataframe.dropna(axis=1, how="all")
    dataframe = normalize_support_columns(dataframe)

    print(f"FAM_001 support file {keyword}: {support_file}")
    print(f"FAM_001 support file {keyword} rows loaded: {len(dataframe)}")
    print(f"FAM_001 support file {keyword} columns resolved: {list(dataframe.columns)}")

    return dataframe


def filter_support_dataframe(dataframe, context):
    """
    Filter SQVI support dataframe by depreciation area, fiscal year and date range.
    """
    if dataframe.empty:
        return dataframe

    result = dataframe.copy()
    fiscal_year_value = get_analysis_date_to(context).year

    if "AFABE" in result.columns:
        result = result[
            result["AFABE"].apply(normalize_depreciation_area) == DEFAULT_DEPRECIATION_AREA
        ].copy()

    if "GJAHR" in result.columns:
        result = result[
            pd.to_numeric(result["GJAHR"], errors="coerce") == fiscal_year_value
        ].copy()

    if "BZDAT" in result.columns:
        date_to = get_analysis_date_to(context)
        date_from = get_analysis_date_from(context)
        posting_dates = pd.to_datetime(result["BZDAT"], errors="coerce").dt.date
        date_mask = posting_dates <= date_to

        if date_from is not None:
            date_mask = date_mask & (posting_dates >= date_from)

        result = result[date_mask.fillna(False)].copy()

    return result


def sum_existing_columns(dataframe, columns):
    """
    Sum numeric values from columns that exist in a dataframe.
    """
    existing_columns = [column for column in columns if column in dataframe.columns]

    if not existing_columns:
        return pd.Series(index=dataframe.index, dtype="float64")

    numeric_dataframe = dataframe[existing_columns].apply(
        lambda column: pd.to_numeric(column.apply(parse_number), errors="coerce")
    )

    return numeric_dataframe.sum(axis=1, min_count=1)


def normalize_depreciation_amount(value):
    """
    Return depreciation amounts as positive values for presentation.
    """
    parsed_value = parse_number(value)

    if pd.isna(parsed_value):
        return pd.NA

    return abs(float(parsed_value))


def classify_movement_amount(row, prefixes):
    """
    Return ANBTR when BWASL starts with one of the requested prefixes.
    """
    transaction_type = normalize_text(row.get(MOVEMENT_REQUIRED_COLUMNS["transaction_type"], ""))

    if transaction_type.endswith(".0"):
        transaction_type = transaction_type[:-2]

    if not transaction_type.startswith(prefixes):
        return pd.NA

    return parse_number(row.get(MOVEMENT_REQUIRED_COLUMNS["amount"], None))


def build_support_key(company_code, asset_number, asset_subnumber):
    """
    Build the asset-level key used to enrich FAM01 with SQVI BAL/MOV data.
    """
    return (
        normalize_company_code(company_code),
        normalize_asset_number(asset_number),
        normalize_asset_number(asset_subnumber),
    )


def support_dataframe_has_columns(dataframe, required_columns, support_name):
    """
    Return True when support dataframe has all required technical columns.
    """
    missing_columns = sorted(set(required_columns) - set(dataframe.columns))

    if missing_columns:
        print(f"FAM_001 {support_name} support skipped. Missing columns: {missing_columns}")
        return False

    return True


def build_balance_support_by_asset(context):
    """
    Build FAM01 asset-level metrics from SQVI balance file ANLA + ANLC.
    """
    dataframe = read_optional_support_file(context, FAM01_BALANCE_FILE_KEYWORD)
    dataframe = filter_support_dataframe(dataframe, context)

    required_columns = set(BALANCE_REQUIRED_COLUMNS.values())
    output_columns = [
        "Support Key",
        "CAP en fecha inicio (LC)",
        "Amortiz.acumul.en fecha de inicio",
        "VNC en fecha inicio",
    ]

    if dataframe.empty:
        return pd.DataFrame(columns=output_columns)

    if not support_dataframe_has_columns(dataframe, required_columns, "BAL"):
        return pd.DataFrame(columns=output_columns)

    working = pd.DataFrame(index=dataframe.index)
    working["Support Key"] = [
        build_support_key(company_code, asset_number, asset_subnumber)
        for company_code, asset_number, asset_subnumber in zip(
            dataframe[BALANCE_REQUIRED_COLUMNS["company_code"]],
            dataframe[BALANCE_REQUIRED_COLUMNS["asset_number"]],
            dataframe[BALANCE_REQUIRED_COLUMNS["asset_subnumber"]],
        )
    ]
    working["CAP en fecha inicio (LC)"] = pd.to_numeric(
        dataframe[BALANCE_REQUIRED_COLUMNS["apc_opening"]].apply(parse_number),
        errors="coerce",
    )
    working["Amortiz.acumul.en fecha de inicio"] = sum_existing_columns(
        dataframe,
        BALANCE_DEPRECIATION_COLUMNS,
    ).apply(normalize_depreciation_amount)
    working["VNC en fecha inicio"] = (
        working["CAP en fecha inicio (LC)"]
        - working["Amortiz.acumul.en fecha de inicio"]
    )

    summary = (
        working
        .groupby("Support Key", dropna=False)[[
            "CAP en fecha inicio (LC)",
            "Amortiz.acumul.en fecha de inicio",
            "VNC en fecha inicio",
        ]]
        .sum(min_count=1)
        .reset_index()
    )

    print(f"FAM_001 BAL support asset rows: {len(summary)}")

    return summary[output_columns]


def build_movement_support_by_asset(context):
    """
    Build FAM01 asset-level metrics from SQVI movement file ANLA + ANEP.
    """
    dataframe = read_optional_support_file(context, FAM01_MOVEMENT_FILE_KEYWORD)
    dataframe = filter_support_dataframe(dataframe, context)

    required_columns = set(MOVEMENT_REQUIRED_COLUMNS.values())
    output_columns = [
        "Support Key",
        "Capitalización",
        "Valoración",
        "Depreciación en fecha de fin",
    ]

    if dataframe.empty:
        return pd.DataFrame(columns=output_columns)

    if not support_dataframe_has_columns(dataframe, required_columns, "MOV"):
        return pd.DataFrame(columns=output_columns)

    working = pd.DataFrame(index=dataframe.index)
    working["Support Key"] = [
        build_support_key(company_code, asset_number, asset_subnumber)
        for company_code, asset_number, asset_subnumber in zip(
            dataframe[MOVEMENT_REQUIRED_COLUMNS["company_code"]],
            dataframe[MOVEMENT_REQUIRED_COLUMNS["asset_number"]],
            dataframe[MOVEMENT_REQUIRED_COLUMNS["asset_subnumber"]],
        )
    ]
    working["Capitalización"] = dataframe.apply(
        lambda row: classify_movement_amount(row, CAPITALIZATION_TRANSACTION_PREFIXES),
        axis=1,
    )
    working["Valoración"] = dataframe.apply(
        lambda row: classify_movement_amount(row, VALUATION_TRANSACTION_PREFIXES),
        axis=1,
    )
    working["Depreciación en fecha de fin"] = sum_existing_columns(
        dataframe,
        MOVEMENT_DEPRECIATION_COLUMNS,
    ).apply(normalize_depreciation_amount)

    summary = (
        working
        .groupby("Support Key", dropna=False)[[
            "Capitalización",
            "Valoración",
            "Depreciación en fecha de fin",
        ]]
        .sum(min_count=1)
        .reset_index()
    )

    print(f"FAM_001 MOV support asset rows: {len(summary)}")

    return summary[output_columns]


def build_fam01_support_metrics(context):
    """
    Combine optional SQVI BAL/MOV asset-level metrics for FAM01 enrichment.
    """
    balance_support = build_balance_support_by_asset(context)
    movement_support = build_movement_support_by_asset(context)

    support_metrics = {
        "balance": balance_support,
        "movement": movement_support,
    }

    if balance_support.empty and movement_support.empty:
        print("FAM_001 SQVI support: no BAL/MOV enrichment applied.")
    else:
        print(
            "FAM_001 SQVI support applied: "
            f"BAL rows={len(balance_support)}, MOV rows={len(movement_support)}"
        )

    return support_metrics


def build_support_lookup(dataframe, metric_column):
    """
    Build lookup dictionary Support Key -> metric value.
    """
    if dataframe is None or dataframe.empty:
        return {}

    if "Support Key" not in dataframe.columns or metric_column not in dataframe.columns:
        return {}

    return dataframe.set_index("Support Key")[metric_column].to_dict()


def lookup_support_metric(support_keys, support_dataframe, metric_column, default_series):
    """
    Return metric series using SQVI support values when available.
    """
    lookup = build_support_lookup(support_dataframe, metric_column)

    if not lookup:
        return default_series

    values = [lookup.get(key, pd.NA) for key in support_keys]
    return pd.Series(values, index=default_series.index)


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


def convert_series_lc_to_usd(local_currency_series, currency_series, fx_lookup):
    """
    Convert an LC series to USD using row-level currency values.
    """
    return pd.Series(
        [
            convert_lc_to_usd(value, currency, fx_lookup)
            for value, currency in zip(local_currency_series, currency_series)
        ],
        index=local_currency_series.index,
    )


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


def build_fam_001_dataframe(source_dataframe, required_columns, optional_columns, fx_lookup, support_metrics=None):
    """
    Build the FAM01 fixed assets listing output dataframe.
    """
    support_metrics = support_metrics or {}
    balance_support = support_metrics.get("balance", pd.DataFrame())
    movement_support = support_metrics.get("movement", pd.DataFrame())

    output_dataframe = pd.DataFrame(index=source_dataframe.index)

    currency_series = source_dataframe[required_columns["currency"]].apply(normalize_currency)
    company_series = source_dataframe[required_columns["company_code"]].apply(normalize_company_output)
    asset_number_series = get_asset_number_series(source_dataframe, required_columns["asset_number"])
    asset_subnumber_series = get_asset_number_series(source_dataframe, required_columns["asset_subnumber"])

    support_keys = [
        build_support_key(company_code, asset_number, asset_subnumber)
        for company_code, asset_number, asset_subnumber in zip(
            company_series,
            asset_number_series,
            asset_subnumber_series,
        )
    ]

    acquisition_value_lc = source_dataframe[
        required_columns["acquisition_value"]
    ].apply(parse_number)

    accumulated_depreciation_lc = source_dataframe[
        required_columns["accumulated_depreciation"]
    ].apply(parse_number)

    book_value_lc = source_dataframe[
        required_columns["book_value"]
    ].apply(parse_number)

    cap_begin_lc = lookup_support_metric(
        support_keys=support_keys,
        support_dataframe=balance_support,
        metric_column="CAP en fecha inicio (LC)",
        default_series=blank_number_series(source_dataframe),
    )
    beginning_accumulated_depreciation = lookup_support_metric(
        support_keys=support_keys,
        support_dataframe=balance_support,
        metric_column="Amortiz.acumul.en fecha de inicio",
        default_series=blank_number_series(source_dataframe),
    )
    beginning_book_value = lookup_support_metric(
        support_keys=support_keys,
        support_dataframe=balance_support,
        metric_column="VNC en fecha inicio",
        default_series=blank_number_series(source_dataframe),
    )
    capitalization = lookup_support_metric(
        support_keys=support_keys,
        support_dataframe=movement_support,
        metric_column="Capitalización",
        default_series=zero_number_series(source_dataframe),
    ).fillna(0)
    valuation = lookup_support_metric(
        support_keys=support_keys,
        support_dataframe=movement_support,
        metric_column="Valoración",
        default_series=zero_number_series(source_dataframe),
    ).fillna(0)
    depreciation_end = lookup_support_metric(
        support_keys=support_keys,
        support_dataframe=movement_support,
        metric_column="Depreciación en fecha de fin",
        default_series=blank_number_series(source_dataframe),
    )

    output_dataframe["CoCo"] = company_series
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
    output_dataframe["Núm.AF"] = asset_number_series
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

    output_dataframe["CAP en fecha inicio (LC)"] = cap_begin_lc
    output_dataframe["CAP en fecha inicio (USD)"] = convert_series_lc_to_usd(
        cap_begin_lc,
        currency_series,
        fx_lookup,
    )
    output_dataframe["CAP en fecha inicio (moneda)"] = currency_series

    output_dataframe["Amortiz.acumul.en fecha de inicio"] = beginning_accumulated_depreciation
    output_dataframe["Amortiz.acumul.en fecha de inicio (moneda)"] = currency_series

    output_dataframe["Revaloración acumulada fecha inicio"] = zero_number_series(source_dataframe)
    output_dataframe["Revaloración acumulada fecha inicio (moneda)"] = currency_series

    output_dataframe["VNC en fecha inicio"] = beginning_book_value
    output_dataframe["VNC en fecha inicio (moneda)"] = currency_series

    output_dataframe["Capitalización"] = capitalization
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

    output_dataframe["Valoración"] = valuation
    output_dataframe["Valoración (moneda)"] = currency_series

    output_dataframe["CAP en fecha de fin (LC)"] = acquisition_value_lc
    output_dataframe["CAP en fecha de fin (USD)"] = convert_series_lc_to_usd(
        acquisition_value_lc,
        currency_series,
        fx_lookup,
    )
    output_dataframe["CAP en fecha de fin (moneda)"] = currency_series

    output_dataframe["VNC en fecha de fin (LC)"] = book_value_lc
    output_dataframe["VNC en fecha de fin (USD)"] = convert_series_lc_to_usd(
        book_value_lc,
        currency_series,
        fx_lookup,
    )
    output_dataframe["VNC en fecha de fin (moneda)"] = currency_series

    output_dataframe["Depreciación en fecha de fin"] = depreciation_end
    output_dataframe["Depreciación en fecha de fin (moneda)"] = currency_series

    output_dataframe["Amortización acumulada en fecha fin (LC)"] = accumulated_depreciation_lc
    output_dataframe["Amortización acumulada en fecha fin (USD)"] = convert_series_lc_to_usd(
        accumulated_depreciation_lc,
        currency_series,
        fx_lookup,
    )
    output_dataframe["Amortización acumulada en fecha fin (moneda)"] = currency_series

    output_dataframe["Subnº"] = asset_subnumber_series
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

    support_metrics = build_fam01_support_metrics(context)

    output_dataframe = build_fam_001_dataframe(
        source_dataframe=source_dataframe,
        required_columns=required_columns,
        optional_columns=optional_columns,
        fx_lookup=fx_lookup,
        support_metrics=support_metrics,
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
