"""
GL_007 - General Journals Posted To The Prior Fiscal Period.

SAP LBR defensible logic:
- Posting Date = BKPF-BUDAT / Portuguese header "Dt.lçto.".
- Entry Date = BKPF-CPUDT / Portuguese header "Dt.entr.".
- A line is flagged when both dates are valid and Posting Date is before the first
  day of the Entry Date month. This identifies journals created/entered in the
  current open period but posted back to a prior fiscal period.
- The control is line-level and reads BSIS + BSAS. It writes/replaces only GL07.
- FxRates is optional: if unavailable, USD columns are left blank.
"""

from pathlib import Path

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
    normalize_company_output,
    normalize_fx_rates,
    normalize_text,
    open_or_create_gl_output_workbook,
    recreate_gl_sheet,
    require_columns,
    save_gl_output_workbook,
    select_fx_rate_to_usd,
    to_datetime_value,
    write_dataframe_to_sheet,
    write_single_sheet_workbook_fast,
)

SHEET_NAME = "GL07"

REQUIRED_COLUMNS = {
    "company_code": ["Empr", "BUKRS", "BKPF-BUKRS", "BSIS-BUKRS", "BSAS-BUKRS"],
    "document_number": ["Nº doc.", "Nº doc", "BELNR", "BKPF-BELNR", "BSIS-BELNR", "BSAS-BELNR"],
    "fiscal_year": ["Ano", "GJAHR", "BKPF-GJAHR", "BSIS-GJAHR", "BSAS-GJAHR"],
    "document_type": ["Tp.doc.", "Tp.doc", "BLART", "BKPF-BLART"],
    "document_date": ["Data doc.", "Data doc", "BLDAT", "BKPF-BLDAT"],
    "posting_date": ["Dt.lçto.", "Dt.lçto", "Dt.lcto.", "Dt.lcto", "BUDAT", "BKPF-BUDAT"],
    "entry_date": ["Dt.entr.", "Dt.entr", "CPUDT", "BKPF-CPUDT"],
    "entry_time": ["Hora", "CPUTM", "BKPF-CPUTM"],
    "user": ["Nome do usuário", "Nome do usuario", "USNAM", "BKPF-USNAM"],
    "transaction_code": ["CódT", "CodT", "TCODE", "BKPF-TCODE"],
    "header_text": ["Texto cabeçalho documento", "Texto cabecalho documento", "BKTXT", "BKPF-BKTXT"],
    "currency": ["Moeda", "WAERS", "BKPF-WAERS"],
    "line_item": ["Itm", "BUZEI", "BSIS-BUZEI", "BSAS-BUZEI"],
    "gl_account": ["Razão", "Razao", "HKONT", "BSIS-HKONT", "BSAS-HKONT"],
    "debit_credit_indicator": ["D/C", "SHKZG", "BSIS-SHKZG", "BSAS-SHKZG"],
    "amount_local_currency": ["Montante em MI", "DMBTR", "BSIS-DMBTR", "BSAS-DMBTR"],
    "amount_document_currency": ["Montante", "WRBTR", "BSIS-WRBTR", "BSAS-WRBTR"],
    "line_text": ["Texto", "SGTXT", "BSIS-SGTXT", "BSAS-SGTXT"],
}

OPTIONAL_COLUMNS = {
    "reference": ["Referência", "Referencia", "XBLNR", "BKPF-XBLNR"],
    "reversal_document": ["Estorno c/", "Estorno c/.1", "Estorno c/.2", "Estorno c/.3", "STBLG", "BKPF-STBLG"],
    "creator_user": ["Pré-edição", "Pre-edição", "Pré-edicao", "Pre-edicao", "PPNAM", "BKPF-PPNAM"],
}

OUTPUT_COLUMNS = [
    "CoCo", "Company", "Company Main Currency", "Company System Currency", "Journal Number", "TransId", "Line",
    "Journal Type", "Posting Date", "Document Date", "Entry Date", "Update Date", "Entry Weekday", "Journal Memo",
    "Journal Entry Status", "Reverses TransId", "Auto Reversal", "Account Code", "Account Name", "Line Memo", "Debit",
    "Credit", "Line Amount Local", "FC Currency", "Debit Credit Indicator", "Line Amount USD", "USD Method", "USD Rate",
    "USD Rate Date", "Header Total Local", "Header Total USD", "Creator ID", "Creator Name", "Approver ID", "Approver Name",
    "Period Code", "Period Name", "Period From", "Period To", "Current Period Status", "Latest Period Log Status",
    "Latest Period Log Date", "Period Open Date", "Period Close Date", "Latest Period Open Date", "Latest Period Close Date",
    "Days From Period End", "Posted Before Period Open", "Posting Month", "Prior Period Reason",
]

DATE_COLUMNS = {"Posting Date", "Document Date", "Entry Date", "Update Date", "USD Rate Date", "Period From", "Period To", "Latest Period Log Date", "Period Open Date", "Period Close Date", "Latest Period Open Date", "Latest Period Close Date"}
AMOUNT_COLUMNS = {"Debit", "Credit", "Line Amount Local", "Line Amount USD", "USD Rate", "Header Total Local", "Header Total USD"}
INTEGER_COLUMNS = {"Entry Weekday", "Days From Period End"}


def print_header(title):
    print(title)
    print("-" * len(title))


def blank_series(index):
    return pd.Series("", index=index, dtype="object")


def clean_text_series(series):
    result = series.where(series.notna(), "").astype(str).str.strip()
    result = result.mask(result.str.lower() == "nan", "")
    return result.str.replace(r"\.0$", "", regex=True)


def normalize_company_series(series):
    result = clean_text_series(series)
    numeric_mask = result.str.fullmatch(r"\d+")
    return result.where(~numeric_mask, result.astype("Int64").astype(str))


def parse_date_series(series):
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, errors="coerce")
    text_series = clean_text_series(series)
    result = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    yyyymmdd = text_series.str.fullmatch(r"\d{8}", na=False)
    iso = text_series.str.fullmatch(r"\d{4}-\d{2}-\d{2}.*", na=False)
    other = (text_series != "") & ~yyyymmdd & ~iso
    result.loc[yyyymmdd] = pd.to_datetime(text_series.loc[yyyymmdd], format="%Y%m%d", errors="coerce")
    result.loc[iso] = pd.to_datetime(text_series.loc[iso].str[:10], format="%Y-%m-%d", errors="coerce")
    result.loc[other] = pd.to_datetime(text_series.loc[other], errors="coerce", dayfirst=True)
    return result


def parse_number_series(series):
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")
    text = clean_text_series(series).str.replace("\u00a0", "", regex=False).str.replace(" ", "", regex=False)
    comma = text.str.contains(",", regex=False, na=False)
    dot_dec = text.str.contains(r"\.\d{1,2}$", regex=True, na=False)
    normalized = pd.Series("", index=series.index, dtype="object")
    normalized.loc[comma] = text.loc[comma].str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    normalized.loc[~comma & dot_dec] = text.loc[~comma & dot_dec].str.replace(",", "", regex=False)
    normalized.loc[~comma & ~dot_dec] = text.loc[~comma & ~dot_dec].str.replace(",", "", regex=False)
    return pd.to_numeric(normalized, errors="coerce")


def get_optional_series(dataframe, column_name):
    if column_name is None:
        return blank_series(dataframe.index)
    return dataframe[column_name]


def build_company_currency_map(master_dataframe):
    if master_dataframe.empty:
        return {}
    company_col = get_optional_column(master_dataframe, ["BUKRS", "Empr"])
    currency_col = get_optional_column(master_dataframe, ["WAERS", "Moeda", "Currency"])
    if company_col is None or currency_col is None:
        return {}
    result = {}
    for _, row in master_dataframe.iterrows():
        company = normalize_company_output(row.get(company_col, ""))
        currency = normalize_text(row.get(currency_col, "")).upper()
        if company and company not in result:
            result[company] = currency
    return result


def resolve_optional_columns(dataframe):
    return {name: get_optional_column(dataframe, aliases) for name, aliases in OPTIONAL_COLUMNS.items()}


def build_trans_id(company_code, fiscal_year, document_number):
    return company_code.astype(str).str.strip() + "|" + fiscal_year.astype(str).str.strip() + "|" + document_number.astype(str).str.strip()


def prepare_source_dataframe(source_dataframe, context, source_name):
    if source_dataframe.empty:
        return pd.DataFrame(), {}, {}
    print(f"Resolving columns for {source_name}...")
    required = require_columns(source_dataframe, REQUIRED_COLUMNS, f"GL {source_name}")
    optional = resolve_optional_columns(source_dataframe)
    module_config = context["module"]
    print(f"{source_name} rows before company filter: {len(source_dataframe)}")
    result = filter_by_company(source_dataframe, required["company_code"], module_config.get("companies", ""))
    print(f"{source_name} rows after company filter: {len(result)}")
    entry_date = parse_date_series(result[required["entry_date"]])
    from_date = to_datetime_value(module_config.get("from", ""))
    to_date = to_datetime_value(module_config.get("to", ""))
    if from_date is not None and to_date is not None:
        mask = (entry_date >= from_date) & (entry_date <= to_date)
        result = result[mask].copy()
        entry_date = entry_date.loc[result.index]
    print(f"{source_name} rows after entry date filter: {len(result)}")
    if result.empty:
        return result, required, optional
    company_code = normalize_company_series(result[required["company_code"]])
    fiscal_year = clean_text_series(result[required["fiscal_year"]])
    document_number = clean_text_series(result[required["document_number"]])
    result["_SAP_COMPANY_CODE"] = company_code
    result["_SAP_FISCAL_YEAR"] = fiscal_year
    result["_SAP_DOCUMENT_NUMBER"] = document_number
    result["_SAP_LINE_ITEM"] = clean_text_series(result[required["line_item"]])
    result["_POSTING_DATE"] = parse_date_series(result[required["posting_date"]])
    result["_ENTRY_DATE"] = entry_date
    result["_DOCUMENT_DATE"] = parse_date_series(result[required["document_date"]])
    result["_TRANS_ID"] = build_trans_id(company_code, fiscal_year, document_number)
    result["_SOURCE"] = source_name
    return result, required, optional


def get_prior_period_mask(dataframe):
    entry_month_start = dataframe["_ENTRY_DATE"].dt.to_period("M").dt.to_timestamp()
    return dataframe["_POSTING_DATE"].notna() & dataframe["_ENTRY_DATE"].notna() & (dataframe["_POSTING_DATE"] < entry_month_start)


def build_header_total_map(dataframe, required):
    if dataframe.empty:
        return pd.Series(dtype="float64")
    indicator = clean_text_series(dataframe[required["debit_credit_indicator"]]).str.upper()
    amount_abs = parse_number_series(dataframe[required["amount_local_currency"]]).abs().fillna(0)
    base = pd.DataFrame({"TransId": dataframe["_TRANS_ID"], "Debit": amount_abs.where(indicator == "S", 0), "Credit": amount_abs.where(indicator == "H", 0)})
    grouped = base.groupby("TransId", dropna=False)[["Debit", "Credit"]].sum()
    return grouped[["Debit", "Credit"]].max(axis=1)


def build_output_lines(dataframe, required, optional, master_dataframe):
    if dataframe.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)
    company_name_map = build_company_name_map(master_dataframe)
    company_currency_map = build_company_currency_map(master_dataframe)
    account_name_map = build_gl_account_name_map(master_dataframe)
    company = dataframe["_SAP_COMPANY_CODE"]
    account = clean_text_series(dataframe[required["gl_account"]])
    indicator = clean_text_series(dataframe[required["debit_credit_indicator"]]).str.upper()
    local_abs = parse_number_series(dataframe[required["amount_local_currency"]]).abs().fillna(0)
    debit = local_abs.where(indicator == "S", 0)
    credit = local_abs.where(indicator == "H", 0)
    signed = debit - credit
    period_from = dataframe["_POSTING_DATE"].dt.to_period("M").dt.to_timestamp()
    period_to = dataframe["_POSTING_DATE"].dt.to_period("M").dt.to_timestamp("M")
    output = pd.DataFrame({
        "CoCo": company,
        "Company": company.map(company_name_map).fillna(""),
        "Company Main Currency": company.map(company_currency_map).fillna(""),
        "Company System Currency": "",
        "Journal Number": dataframe["_SAP_DOCUMENT_NUMBER"],
        "TransId": dataframe["_TRANS_ID"],
        "Line": dataframe["_SAP_LINE_ITEM"],
        "Journal Type": clean_text_series(dataframe[required["document_type"]]),
        "Posting Date": dataframe["_POSTING_DATE"],
        "Document Date": dataframe["_DOCUMENT_DATE"],
        "Entry Date": dataframe["_ENTRY_DATE"],
        "Update Date": dataframe["_ENTRY_DATE"],
        "Entry Weekday": dataframe["_ENTRY_DATE"].dt.dayofweek,
        "Journal Memo": clean_text_series(dataframe[required["header_text"]]),
        "Journal Entry Status": get_optional_series(dataframe, optional["reversal_document"]).pipe(clean_text_series).mask(lambda s: s == "", "Normal Entry").mask(lambda s: s != "Normal Entry", "Reversal Entry"),
        "Reverses TransId": clean_text_series(get_optional_series(dataframe, optional["reversal_document"])),
        "Auto Reversal": "N",
        "Account Code": account,
        "Account Name": account.map(account_name_map).fillna(""),
        "Line Memo": clean_text_series(dataframe[required["line_text"]]),
        "Debit": debit,
        "Credit": credit,
        "Line Amount Local": signed,
        "FC Currency": clean_text_series(dataframe[required["currency"]]),
        "Debit Credit Indicator": indicator,
        "Line Amount USD": "",
        "USD Method": "",
        "USD Rate": "",
        "USD Rate Date": "",
        "Header Total Local": dataframe["_TRANS_ID"].map(build_header_total_map(dataframe, required)).fillna(0),
        "Header Total USD": "",
        "Creator ID": clean_text_series(get_optional_series(dataframe, optional["creator_user"])),
        "Creator Name": "",
        "Approver ID": clean_text_series(dataframe[required["user"]]),
        "Approver Name": "",
        "Period Code": dataframe["_ENTRY_DATE"].dt.strftime("%Y-%m").fillna(""),
        "Period Name": dataframe["_ENTRY_DATE"].dt.strftime("%Y-%m").fillna(""),
        "Period From": period_from,
        "Period To": period_to,
        "Current Period Status": "Y",
        "Latest Period Log Status": "C",
        "Latest Period Log Date": dataframe["_ENTRY_DATE"],
        "Period Open Date": "",
        "Period Close Date": "",
        "Latest Period Open Date": "",
        "Latest Period Close Date": "",
        "Days From Period End": (dataframe["_POSTING_DATE"] - period_to).dt.days,
        "Posted Before Period Open": "N",
        "Posting Month": dataframe["_POSTING_DATE"].dt.strftime("%Y-%m").fillna(""),
        "Prior Period Reason": "Posted to period before entry period",
    })
    return output[OUTPUT_COLUMNS]


def enrich_usd(output_dataframe, fx_rates_dataframe):
    if output_dataframe.empty or fx_rates_dataframe.empty:
        return output_dataframe

    normalized_fx_dataframe = normalize_fx_rates(fx_rates_dataframe)
    if normalized_fx_dataframe.empty:
        return output_dataframe

    fx_cache = {}
    line_amount_usd_values = []
    usd_method_values = []
    usd_rate_values = []
    usd_rate_date_values = []
    header_total_usd_values = []

    for _, row in output_dataframe.iterrows():
        currency = normalize_text(row.get("FC Currency", "")).upper()

        if currency == "":
            currency = normalize_text(row.get("Company Main Currency", "")).upper()

        fx_date = row.get("Posting Date", pd.NaT)

        if pd.isna(fx_date):
            fx_date = row.get("Document Date", pd.NaT)

        if pd.isna(fx_date):
            cache_key = (currency, "")
        else:
            cache_key = (currency, fx_date.strftime("%Y-%m-%d"))

        if cache_key in fx_cache:
            fx_details = fx_cache[cache_key]
        else:
            fx_details = select_fx_rate_to_usd(
                normalized_fx_dataframe=normalized_fx_dataframe,
                currency=currency,
                requested_date=fx_date,
            )
            fx_cache[cache_key] = fx_details

        if fx_details is None:
            line_amount_usd_values.append(pd.NA)
            usd_method_values.append("")
            usd_rate_values.append(pd.NA)
            usd_rate_date_values.append(pd.NaT)
            header_total_usd_values.append(pd.NA)
            continue

        line_amount_local = pd.to_numeric(
            pd.Series([row.get("Line Amount Local", pd.NA)]),
            errors="coerce",
        ).iloc[0]
        header_total_local = pd.to_numeric(
            pd.Series([row.get("Header Total Local", pd.NA)]),
            errors="coerce",
        ).iloc[0]

        if pd.isna(line_amount_local):
            line_amount_usd_values.append(pd.NA)
        else:
            line_amount_usd_values.append(line_amount_local * fx_details["fx_to_usd"])

        if pd.isna(header_total_local):
            header_total_usd_values.append(pd.NA)
        else:
            header_total_usd_values.append(header_total_local * fx_details["fx_to_usd"])

        usd_method_values.append(fx_details["method"])
        usd_rate_values.append(fx_details["usd_rate"])
        usd_rate_date_values.append(fx_details["rate_date"])

    output_dataframe["Company System Currency"] = "USD"
    output_dataframe["Line Amount USD"] = line_amount_usd_values
    output_dataframe["USD Method"] = usd_method_values
    output_dataframe["USD Rate"] = usd_rate_values
    output_dataframe["USD Rate Date"] = usd_rate_date_values
    output_dataframe["Header Total USD"] = header_total_usd_values

    missing_usd = output_dataframe["Line Amount USD"].isna().sum()

    if missing_usd > 0:
        print()
        print("WARNING: Some GL FX rates were not found.")
        print("USD columns are blank for those rows.")
        print(f"Rows without USD amount: {missing_usd}")
        print()

    return output_dataframe


def run_gl_007(context):
    print_header("GL_007 - General Journals Posted To The Prior Fiscal Period")
    bsis = load_gl_bsis_data(context)
    bsas = load_gl_bsas_data(context)
    master = load_gl_master_data(context)
    fx_rates = load_gl_fx_rates_data(context)
    frames = []
    for source_name, source in (("BSIS", bsis), ("BSAS", bsas)):
        prepared, required, optional = prepare_source_dataframe(source, context, source_name)
        if prepared.empty:
            continue
        flagged = prepared[get_prior_period_mask(prepared)].copy()
        print(f"{source_name} prior-period rows: {len(flagged)}")
        frames.append(build_output_lines(flagged, required, optional, master))
    output = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=OUTPUT_COLUMNS)
    output = enrich_usd(output, fx_rates)
    output_file = get_gl_output_file(context)
    print(f"Writing {len(output)} rows to {output_file} [{SHEET_NAME}]...")
    if not Path(output_file).exists():
        write_single_sheet_workbook_fast(output_file, SHEET_NAME, output, DATE_COLUMNS, AMOUNT_COLUMNS, INTEGER_COLUMNS)
    else:
        workbook = open_or_create_gl_output_workbook(output_file)
        worksheet = recreate_gl_sheet(workbook, SHEET_NAME)
        write_dataframe_to_sheet(worksheet, output)
        apply_standard_gl_formatting(worksheet, output, DATE_COLUMNS, AMOUNT_COLUMNS, INTEGER_COLUMNS)
        save_gl_output_workbook(workbook, output_file)
    print("GL_007 completed successfully.")
    return output
