"""GL_006 - Frequently Reversed General Journals."""
from __future__ import annotations

import pandas as pd

from core.gl_common import (
    apply_standard_gl_formatting,
    build_company_name_map,
    build_gl_account_name_map,
    filter_by_company,
    get_gl_output_file,
    get_optional_column,
    load_gl_bsas_data,
    load_gl_bsis_data,
    load_gl_fx_rates_data,
    load_gl_master_data,
    normalize_code_keep_leading_zeroes,
    normalize_company_output,
    normalize_text,
    open_or_create_gl_output_workbook,
    recreate_gl_sheet,
    require_columns,
    save_gl_output_workbook,
    to_datetime_value,
    write_dataframe_to_sheet,
    write_single_sheet_workbook_fast,
)

SHEET_NAME = "GL06"

REQUIRED_COLUMNS = {
    "company_code": ["Empr", "BUKRS", "BKPF-BUKRS", "BSIS-BUKRS", "BSAS-BUKRS"],
    "document_number": ["Nº doc.", "Nº doc", "BELNR", "BKPF-BELNR", "BSIS-BELNR", "BSAS-BELNR"],
    "fiscal_year": ["Ano", "GJAHR", "BKPF-GJAHR", "BSIS-GJAHR", "BSAS-GJAHR"],
    "posting_date": ["Dt.lçto.", "Dt.lçto", "Dt.lcto.", "Dt.lcto", "BUDAT", "BKPF-BUDAT"],
    "user": ["Nome do usuário", "Nome do usuario", "USNAM", "BKPF-USNAM"],
    "currency": ["Moeda", "WAERS", "BKPF-WAERS"],
    "line_item": ["Itm", "BUZEI", "BSIS-BUZEI", "BSAS-BUZEI"],
    "gl_account": ["Razão", "Razao", "HKONT", "BSIS-HKONT", "BSAS-HKONT"],
    "debit_credit_indicator": ["D/C", "SHKZG", "BSIS-SHKZG", "BSAS-SHKZG"],
    "amount_local_currency": ["Montante em MI", "DMBTR", "BSIS-DMBTR", "BSAS-DMBTR"],
    "amount_document_currency": ["Montante", "WRBTR", "BSIS-WRBTR", "BSAS-WRBTR"],
}

OPTIONAL_COLUMNS = {
    "reversal_document": ["Estorno c/", "Estorno c/.1", "Estorno c/.2", "Estorno c/.3", "STBLG", "BKPF-STBLG"],
    "reversal_year": ["Ano.4", "Ano.3", "Ano.2", "Ano.1", "STJAH", "BKPF-STJAH"],
}

OUTPUT_COLUMNS = [
    "CoCo", "Company", "TransId", "Journal Number", "Posting Date", "Reverses TransId",
    "Auto Reversal", "Line", "Account Code", "Account Name", "Create User ID", "Create User Name",
    "Debit", "Credit", "Line Amount Local", "Line Amount Abs", "Amount in Reporting Currency",
    "Report Currency", "Amount in Document Currency", "Document Currency", "Reversal Count",
]

DATE_COLUMNS = {"Posting Date"}
AMOUNT_COLUMNS = {"Debit", "Credit", "Line Amount Local", "Line Amount Abs", "Amount in Reporting Currency", "Amount in Document Currency"}
INTEGER_COLUMNS = {"Reversal Count"}



def clean_text_series(series: pd.Series) -> pd.Series:
    result = series.copy()
    result = result.where(result.notna(), "")
    result = result.astype(str).str.strip()
    result = result.mask(result.str.lower().isin(["nan", "nat", "none"]), "")
    return result.str.replace(r"\.0$", "", regex=True)


def parse_date_series(series: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, errors="coerce")

    text_series = clean_text_series(series)
    result = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")

    blank_mask = text_series == ""
    yyyymmdd_mask = ~blank_mask & text_series.str.fullmatch(r"\d{8}", na=False)
    yyyy_mm_dd_mask = ~blank_mask & text_series.str.fullmatch(r"\d{4}-\d{2}-\d{2}.*", na=False)
    remaining_mask = ~blank_mask & ~yyyymmdd_mask & ~yyyy_mm_dd_mask

    if yyyymmdd_mask.any():
        result.loc[yyyymmdd_mask] = pd.to_datetime(
            text_series.loc[yyyymmdd_mask],
            format="%Y%m%d",
            errors="coerce",
        )

    if yyyy_mm_dd_mask.any():
        result.loc[yyyy_mm_dd_mask] = pd.to_datetime(
            text_series.loc[yyyy_mm_dd_mask].str[:10],
            format="%Y-%m-%d",
            errors="coerce",
        )

    if remaining_mask.any():
        result.loc[remaining_mask] = pd.to_datetime(
            text_series.loc[remaining_mask],
            errors="coerce",
            dayfirst=True,
        )

    return result


def parse_number_series(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    text_series = clean_text_series(series)
    text_series = text_series.str.replace("\u00a0", "", regex=False)
    text_series = text_series.str.replace(" ", "", regex=False)

    comma_decimal_mask = text_series.str.contains(",", regex=False, na=False)
    dot_decimal_mask = text_series.str.contains(r"\.\d{1,2}$", regex=True, na=False)

    comma_values = text_series[comma_decimal_mask].str.replace(".", "", regex=False)
    comma_values = comma_values.str.replace(",", ".", regex=False)

    dot_values = text_series[~comma_decimal_mask & dot_decimal_mask].str.replace(
        ",",
        "",
        regex=False,
    )

    plain_values = text_series[~comma_decimal_mask & ~dot_decimal_mask]
    plain_values = plain_values.str.replace(",", "", regex=False)

    normalized = pd.Series("", index=series.index, dtype="object")
    normalized.loc[comma_values.index] = comma_values
    normalized.loc[dot_values.index] = dot_values
    normalized.loc[plain_values.index] = plain_values

    return pd.to_numeric(normalized, errors="coerce")


def build_company_currency_map(master_dataframe: pd.DataFrame) -> dict:
    if master_dataframe.empty:
        return {}

    company_column = get_optional_column(
        master_dataframe,
        ["BUKRS", "Empr", "BKPF-BUKRS"],
    )
    currency_column = get_optional_column(
        master_dataframe,
        ["WAERS", "Moeda", "Currency"],
    )

    if company_column is None or currency_column is None:
        return {}

    result = {}

    for _, row in master_dataframe.iterrows():
        company_code = normalize_company_output(row.get(company_column, ""))
        currency = normalize_text(row.get(currency_column, "")).upper()

        if company_code == "":
            continue

        if company_code not in result:
            result[company_code] = currency

    return result

def print_header(title: str):
    print(title)
    print("-" * len(title))


def get_optional_series(dataframe: pd.DataFrame, column_name: str | None) -> pd.Series:
    if column_name is None:
        return pd.Series("", index=dataframe.index, dtype="object")
    return dataframe[column_name]


def normalize_company_series(series: pd.Series) -> pd.Series:
    return clean_text_series(series).map(normalize_company_output)


def normalize_code_series(series: pd.Series) -> pd.Series:
    return clean_text_series(series).map(normalize_code_keep_leading_zeroes)


def build_trans_id_series(company_code: pd.Series, fiscal_year: pd.Series, document_number: pd.Series) -> pd.Series:
    return company_code.astype(str) + "|" + fiscal_year.astype(str) + "|" + document_number.astype(str)


def build_reverses_trans_id_series(company_code: pd.Series, reversal_year: pd.Series, reversal_document: pd.Series) -> pd.Series:
    reversal_document = reversal_document.astype(str).str.strip()
    reversal_year = reversal_year.astype(str).str.strip()
    return pd.Series("", index=company_code.index, dtype="object").mask(
        reversal_document != "",
        company_code.astype(str).str.strip() + "|" + reversal_year + "|" + reversal_document,
    )


def resolve_optional_columns(dataframe: pd.DataFrame) -> dict[str, str | None]:
    return {name: get_optional_column(dataframe, aliases) for name, aliases in OPTIONAL_COLUMNS.items()}


def prepare_source_dataframe(source_dataframe: pd.DataFrame, context: dict, source_name: str):
    if source_dataframe.empty:
        return pd.DataFrame(), {}, {}
    print(f"Resolving columns for {source_name}...")
    required = require_columns(source_dataframe, REQUIRED_COLUMNS, f"GL {source_name}")
    optional = resolve_optional_columns(source_dataframe)
    result = filter_by_company(source_dataframe, required["company_code"], context.get("module", {}).get("companies", context.get("companies", "")))
    print(f"{source_name} rows after company filter: {len(result)}")
    posting_date = parse_date_series(result[required["posting_date"]])
    from_date = to_datetime_value(context.get("module", {}).get("from", context.get("from", "")))
    to_date = to_datetime_value(context.get("module", {}).get("to", context.get("to", "")))
    if not pd.isna(from_date):
        result = result[posting_date >= from_date].copy()
        posting_date = posting_date.loc[result.index]
    if not pd.isna(to_date):
        result = result[posting_date <= to_date].copy()
        posting_date = posting_date.loc[result.index]
    print(f"{source_name} rows after posting date filter: {len(result)}")
    if result.empty:
        return result, required, optional
    company_code = normalize_company_series(result[required["company_code"]])
    fiscal_year = normalize_code_series(result[required["fiscal_year"]])
    document_number = normalize_code_series(result[required["document_number"]])
    result["_SAP_COMPANY_CODE"] = company_code
    result["_SAP_FISCAL_YEAR"] = fiscal_year
    result["_SAP_DOCUMENT_NUMBER"] = document_number
    result["_TRANS_ID"] = build_trans_id_series(company_code, fiscal_year, document_number)
    result["_POSTING_DATE"] = posting_date
    result["_SOURCE"] = source_name
    return result, required, optional


def build_reversed_lines_for_source(source_dataframe: pd.DataFrame, master_dataframe: pd.DataFrame, context: dict, source_name: str) -> pd.DataFrame:
    prepared, required, optional = prepare_source_dataframe(source_dataframe, context, source_name)
    if prepared.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)
    reversal_document = normalize_code_series(get_optional_series(prepared, optional.get("reversal_document")))
    reversal_year = normalize_code_series(get_optional_series(prepared, optional.get("reversal_year")))
    reversed_mask = reversal_document != ""
    prepared = prepared[reversed_mask].copy()
    reversal_document = reversal_document.loc[prepared.index]
    reversal_year = reversal_year.loc[prepared.index]
    print(f"{source_name} reversed journal lines kept: {len(prepared)}")
    if prepared.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)
    company_name_map = build_company_name_map(master_dataframe)
    company_currency_map = build_company_currency_map(master_dataframe)
    account_name_map = build_gl_account_name_map(master_dataframe)
    indicator = clean_text_series(prepared[required["debit_credit_indicator"]]).str.upper()
    amount_abs = parse_number_series(prepared[required["amount_local_currency"]]).abs()
    debit = amount_abs.where(indicator == "S", pd.NA)
    credit = amount_abs.where(indicator == "H", pd.NA)
    line_amount_local = amount_abs.where(indicator != "H", amount_abs * -1)
    document_amount_abs = parse_number_series(prepared[required["amount_document_currency"]]).abs()
    account_code = normalize_code_series(prepared[required["gl_account"]])
    company_code = prepared["_SAP_COMPANY_CODE"]
    output = pd.DataFrame({
        "CoCo": company_code,
        "Company": company_code.map(company_name_map).fillna(""),
        "TransId": prepared["_TRANS_ID"],
        "Journal Number": prepared["_SAP_DOCUMENT_NUMBER"],
        "Posting Date": prepared["_POSTING_DATE"],
        "Reverses TransId": build_reverses_trans_id_series(company_code, reversal_year, reversal_document),
        "Auto Reversal": "N",
        "Line": normalize_code_series(prepared[required["line_item"]]),
        "Account Code": account_code,
        "Account Name": account_code.map(account_name_map).fillna(""),
        "Create User ID": clean_text_series(prepared[required["user"]]),
        "Create User Name": "",
        "Debit": debit,
        "Credit": credit,
        "Line Amount Local": line_amount_local,
        "Line Amount Abs": amount_abs,
        "Amount in Reporting Currency": pd.NA,
        "Report Currency": company_code.map(company_currency_map).fillna(""),
        "Amount in Document Currency": document_amount_abs,
        "Document Currency": clean_text_series(prepared[required["currency"]]),
    })
    return output[OUTPUT_COLUMNS[:-1]].copy()


def add_reporting_currency_from_fx(output_dataframe: pd.DataFrame, fx_dataframe: pd.DataFrame) -> pd.DataFrame:
    result = output_dataframe.copy()
    if result.empty:
        return result
    result["Amount in Reporting Currency"] = pd.NA
    if fx_dataframe.empty:
        print("FxRates input was not found. GL06 reporting-currency amount will remain blank.")
        return result
    # LBR GL uses USD as system/reporting currency when rates are available. Conservative optional fill.
    result["Amount in Reporting Currency"] = result["Line Amount Local"].where(result["Report Currency"].str.upper() == "USD", pd.NA)
    return result


def create_gl06_frequently_reversed_journals(bsis_dataframe: pd.DataFrame, bsas_dataframe: pd.DataFrame, master_dataframe: pd.DataFrame, fx_dataframe: pd.DataFrame, context: dict) -> pd.DataFrame:
    frames = []
    for source_name, source_dataframe in (("BSIS", bsis_dataframe), ("BSAS", bsas_dataframe)):
        frame = build_reversed_lines_for_source(source_dataframe, master_dataframe, context, source_name)
        if not frame.empty:
            frames.append(frame)
    if not frames:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)
    result = pd.concat(frames, ignore_index=True)
    result = result[result["Line Amount Abs"].notna() & (result["Line Amount Abs"] != 0)].copy()
    group_columns = ["CoCo", "Account Code", "Line Amount Abs"]
    counts = result.groupby(group_columns, dropna=False).size().rename("Reversal Count").reset_index()
    min_count_raw = context.get("params", {}).get("min_count", context.get("min_count", "2"))
    try:
        min_count = int(float(normalize_text(min_count_raw) or "2"))
    except ValueError:
        min_count = 2
    flagged = counts[counts["Reversal Count"] > min_count].copy()
    result = result.merge(flagged, on=group_columns, how="inner")
    if result.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)
    result = add_reporting_currency_from_fx(result, fx_dataframe)
    result = result.sort_values(["CoCo", "Account Code", "Line Amount Abs", "Posting Date", "Journal Number", "Line"], kind="stable").reset_index(drop=True)
    return result[OUTPUT_COLUMNS].copy()


def write_gl06_output(output_dataframe: pd.DataFrame, context: dict):
    output_file = get_gl_output_file(context)
    print(f"Output workbook: {output_file}")
    if not output_file.exists():
        fast_written = write_single_sheet_workbook_fast(output_file, SHEET_NAME, output_dataframe, DATE_COLUMNS, AMOUNT_COLUMNS, INTEGER_COLUMNS)
        if fast_written:
            print(f"GL06 rows written: {len(output_dataframe)}")
            return output_file
    workbook = open_or_create_gl_output_workbook(output_file)
    worksheet = recreate_gl_sheet(workbook, SHEET_NAME)
    write_dataframe_to_sheet(worksheet, output_dataframe)
    apply_standard_gl_formatting(worksheet, output_dataframe, DATE_COLUMNS, AMOUNT_COLUMNS, INTEGER_COLUMNS)
    save_gl_output_workbook(workbook, output_file)
    print(f"GL06 rows written: {len(output_dataframe)}")
    return output_file


def run_gl_006(context: dict):
    print_header("Running GL_006 - Frequently Reversed General Journals")
    bsis_dataframe = load_gl_bsis_data(context)
    bsas_dataframe = load_gl_bsas_data(context)
    master_dataframe = load_gl_master_data(context)
    fx_dataframe = load_gl_fx_rates_data(context)
    if bsis_dataframe.empty and bsas_dataframe.empty:
        raise FileNotFoundError("GL06 requires BSIS or BSAS GL journal input files.")
    output_dataframe = create_gl06_frequently_reversed_journals(bsis_dataframe, bsas_dataframe, master_dataframe, fx_dataframe, context)
    write_gl06_output(output_dataframe, context)
    return output_dataframe
