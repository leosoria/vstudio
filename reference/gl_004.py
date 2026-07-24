"""
GL_004 - General Journals Created And Approved By The Same User.

SAP LBR proxy logic:
- Creator ID = BKPF-PPNAM, exported in Portuguese as "Pré-edição".
- Approver ID = BKPF-USNAM, exported in Portuguese as "Nome do usuário".
- Reports journal lines where Creator ID and Approver ID are the same non-blank user.

This control is independent and writes/replaces only the GL04 worksheet.
"""

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


SHEET_NAME = "GL04"


REQUIRED_COLUMNS = {
    "company_code": ["Empr", "BUKRS", "BKPF-BUKRS"],
    "document_number": ["Nº doc.", "Nº doc", "BELNR", "BKPF-BELNR"],
    "fiscal_year": ["Ano", "GJAHR", "BKPF-GJAHR"],
    "document_type": ["Tp.doc.", "Tp.doc", "BLART", "BKPF-BLART"],
    "document_date": ["Data doc.", "Data doc", "BLDAT", "BKPF-BLDAT"],
    "posting_date": ["Dt.lçto.", "Dt.lçto", "Dt.lcto.", "Dt.lcto", "BUDAT", "BKPF-BUDAT"],
    "entry_date": ["Dt.entr.", "Dt.entr", "CPUDT", "BKPF-CPUDT"],
    "entry_time": ["Hora", "CPUTM", "BKPF-CPUTM"],
    "approver_user": ["Nome do usuário", "Nome do usuario", "USNAM", "BKPF-USNAM"],
    "creator_user": [
        "Pré-edição",
        "Pre-edição",
        "Pré-edicao",
        "Pre-edicao",
        "Pré edição",
        "Pre edição",
        "Pré edicao",
        "Pre edicao",
        "PPNAM",
        "BKPF-PPNAM",
    ],
    "transaction_code": ["CódT", "CodT", "TCODE", "BKPF-TCODE"],
    "header_text": ["Texto cabeçalho documento", "Texto cabecalho documento", "BKTXT", "BKPF-BKTXT"],
    "currency": ["Moeda", "WAERS", "BKPF-WAERS"],
    "reference": ["Referência", "Referencia", "XBLNR", "BKPF-XBLNR"],
    "line_item": ["Itm", "BUZEI", "BSIS-BUZEI", "BSAS-BUZEI"],
    "gl_account": ["Razão", "Razao", "HKONT", "BSIS-HKONT", "BSAS-HKONT"],
    "debit_credit_indicator": ["D/C", "SHKZG", "BSIS-SHKZG", "BSAS-SHKZG"],
    "amount_local_currency": ["Montante em MI", "DMBTR", "BSIS-DMBTR", "BSAS-DMBTR"],
    "amount_document_currency": ["Montante", "WRBTR", "BSIS-WRBTR", "BSAS-WRBTR"],
    "line_text": ["Texto", "SGTXT", "BSIS-SGTXT", "BSAS-SGTXT"],
}


OPTIONAL_COLUMNS = {
    "cost_center": ["Centro cst", "KOSTL", "BSIS-KOSTL", "BSAS-KOSTL"],
    "profit_center": ["Cen.lucro", "PRCTR", "BSIS-PRCTR", "BSAS-PRCTR"],
    "order": ["Ordem", "AUFNR", "BSIS-AUFNR", "BSAS-AUFNR"],
    "assignment": ["Atribuição", "Atribuicao", "ZUONR", "BSIS-ZUONR", "BSAS-ZUONR"],
    "clearing_document": ["DocCompens", "AUGBL", "BSIS-AUGBL", "BSAS-AUGBL"],
    "clearing_date": ["Compensaç.", "Compensac.", "Compensação", "Compensacao", "AUGDT", "BSIS-AUGDT", "BSAS-AUGDT"],
    "reversal_document": [
        "Estorno c/",
        "Estorno c/.1",
        "Estorno c/.2",
        "Estorno c/.3",
        "Estorno c",
        "Estorno c.1",
        "Estorno c.2",
        "Estorno c.3",
        "STBLG",
        "BKPF-STBLG",
    ],
    "reversal_year": ["Ano.4", "Ano.3", "Ano.2", "Ano.1", "STJAH", "BKPF-STJAH"],
    "reversal_reason": [
        "Motiv.est",
        "Motiv.est.1",
        "Motiv.est.2",
        "Motiv.est.3",
        "Motiv est",
        "Motiv est.1",
        "Motiv est.2",
        "Motiv est.3",
        "STGRD",
        "BKPF-STGRD",
    ],
}


OUTPUT_COLUMNS = [
    "CoCo",
    "Company",
    "Company Main Currency",
    "Company System Currency",
    "Journal Number",
    "TransId",
    "Line",
    "Journal Type",
    "Posting Date",
    "Document Date",
    "Entry Date",
    "Update Date",
    "Entry Weekday",
    "Journal Memo",
    "Journal Entry Status",
    "Reverses TransId",
    "Auto Reversal",
    "Account Code",
    "Account Name",
    "Line Memo",
    "Debit",
    "Credit",
    "Line Amount Local",
    "FC Currency",
    "Debit Credit Indicator",
    "Line Amount USD",
    "USD Method",
    "USD Rate",
    "USD Rate Date",
    "Header Total Local",
    "Header Total USD",
    "Creator ID",
    "Creator Name",
    "Approver ID",
    "Approver Name",
    "Period Code",
    "Period Name",
    "Period From",
    "Period To",
    "Current Period Status",
    "Latest Period Log Status",
    "Latest Period Log Date",
    "Period Open Date",
    "Period Close Date",
    "Latest Period Open Date",
    "Latest Period Close Date",
    "Days From Period End",
    "Posted Before Period Open",
    "Posting Month",
]


DATE_COLUMNS = {
    "Posting Date",
    "Document Date",
    "Entry Date",
    "Update Date",
    "USD Rate Date",
    "Period From",
    "Period To",
    "Latest Period Log Date",
    "Period Open Date",
    "Period Close Date",
    "Latest Period Open Date",
    "Latest Period Close Date",
}


AMOUNT_COLUMNS = {
    "Debit",
    "Credit",
    "Line Amount Local",
    "Line Amount USD",
    "USD Rate",
    "Header Total Local",
    "Header Total USD",
}


INTEGER_COLUMNS = {
    "Entry Weekday",
    "Days From Period End",
}


def print_header(title):
    print(title)
    print("-" * len(title))


def blank_series(index):
    return pd.Series("", index=index, dtype="object")


def get_optional_series(dataframe, column_name):
    if column_name is None:
        return blank_series(dataframe.index)

    return dataframe[column_name]


def clean_text_series(series):
    result = series.copy()
    result = result.where(result.notna(), "")
    result = result.astype(str).str.strip()
    result = result.mask(result.str.lower() == "nan", "")
    result = result.str.replace(r"\.0$", "", regex=True)

    return result


def normalize_company_series(series):
    result = clean_text_series(series)
    numeric_mask = result.str.fullmatch(r"\d+")
    result = result.where(
        ~numeric_mask,
        result.astype("Int64").astype(str),
    )

    return result


def normalize_code_series(series):
    return clean_text_series(series)


def parse_date_series(series):
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, errors="coerce")

    result = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    text_series = clean_text_series(series)

    blank_mask = text_series == ""
    yyyymmdd_mask = ~blank_mask & text_series.str.fullmatch(r"\d{8}", na=False)
    yyyy_mm_dd_mask = ~blank_mask & text_series.str.fullmatch(
        r"\d{4}-\d{2}-\d{2}.*",
        na=False,
    )
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


def parse_number_series(series):
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


def build_trans_id_series(company_code, fiscal_year, document_number):
    return (
        company_code.astype(str).str.strip()
        + "|"
        + fiscal_year.astype(str).str.strip()
        + "|"
        + document_number.astype(str).str.strip()
    )


def build_reverses_trans_id_series(company_code, reversal_year, reversal_document):
    reversal_document = reversal_document.astype(str).str.strip()
    reversal_year = reversal_year.astype(str).str.strip()

    return pd.Series(
        "",
        index=company_code.index,
        dtype="object",
    ).mask(
        reversal_document != "",
        company_code.astype(str).str.strip()
        + "|"
        + reversal_year.astype(str).str.strip()
        + "|"
        + reversal_document,
    )


def format_period_code_series(date_series):
    return date_series.dt.strftime("%Y-%m").fillna("")


def get_month_start_series(date_series):
    return date_series.dt.to_period("M").dt.to_timestamp()


def get_month_end_series(date_series):
    return date_series.dt.to_period("M").dt.to_timestamp("M")


def calculate_days_from_period_end_series(posting_date_series, period_to_series):
    return (posting_date_series - period_to_series).dt.days


def resolve_optional_columns(source_dataframe):
    resolved_columns = {}

    for logical_name, possible_names in OPTIONAL_COLUMNS.items():
        resolved_columns[logical_name] = get_optional_column(
            source_dataframe,
            possible_names,
        )

    return resolved_columns


def build_company_currency_map(master_dataframe):
    if master_dataframe.empty:
        return {}

    bukrs_column = get_optional_column(
        master_dataframe,
        ["BUKRS", "Empr"],
    )
    currency_column = get_optional_column(
        master_dataframe,
        ["WAERS", "Moeda", "Currency"],
    )

    if bukrs_column is None or currency_column is None:
        return {}

    result = {}

    for _, row in master_dataframe.iterrows():
        company_code = normalize_company_output(row.get(bukrs_column, ""))
        currency = normalize_text(row.get(currency_column, "")).upper()

        if company_code == "":
            continue

        if company_code not in result:
            result[company_code] = currency

    return result


def prepare_source_dataframe(source_dataframe, context, source_name):
    if source_dataframe.empty:
        return pd.DataFrame(), {}, {}

    print(f"Resolving columns for {source_name}...")

    required_columns = require_columns(
        dataframe=source_dataframe,
        required_columns=REQUIRED_COLUMNS,
        source_name=f"GL {source_name}",
    )
    optional_columns = resolve_optional_columns(source_dataframe)
    module_config = context["module"]

    print(f"{source_name} rows before company filter: {len(source_dataframe)}")

    result = filter_by_company(
        dataframe=source_dataframe,
        company_column=required_columns["company_code"],
        companies_filter=module_config.get("companies", ""),
    )

    print(f"{source_name} rows after company filter: {len(result)}")

    posting_date = parse_date_series(result[required_columns["posting_date"]])
    from_date = to_datetime_value(module_config.get("from", ""))
    to_date = to_datetime_value(module_config.get("to", ""))

    posting_date_filter = (posting_date >= from_date) & (posting_date <= to_date)

    result = result[posting_date_filter].copy()
    posting_date = posting_date.loc[result.index]

    print(f"{source_name} rows after posting date filter: {len(result)}")

    if result.empty:
        return result, required_columns, optional_columns

    company_code = normalize_company_series(result[required_columns["company_code"]])
    fiscal_year = normalize_code_series(result[required_columns["fiscal_year"]])
    document_number = normalize_code_series(result[required_columns["document_number"]])
    line_item = normalize_code_series(result[required_columns["line_item"]])
    entry_date = parse_date_series(result[required_columns["entry_date"]])

    result["_SAP_COMPANY_CODE"] = company_code
    result["_SAP_FISCAL_YEAR"] = fiscal_year
    result["_SAP_DOCUMENT_NUMBER"] = document_number
    result["_SAP_LINE_ITEM"] = line_item
    result["_POSTING_DATE"] = posting_date
    result["_ENTRY_DATE"] = entry_date
    result["_ENTRY_WEEKDAY"] = entry_date.dt.dayofweek
    result["_TRANS_ID"] = build_trans_id_series(
        company_code=company_code,
        fiscal_year=fiscal_year,
        document_number=document_number,
    )
    result["_SOURCE"] = source_name

    return result, required_columns, optional_columns


def build_header_total_base(prepared_dataframe, required_columns):
    if prepared_dataframe.empty:
        return pd.DataFrame(columns=["TransId", "Debit Total", "Credit Total"])

    sap_indicator = clean_text_series(
        prepared_dataframe[required_columns["debit_credit_indicator"]]
    ).str.upper()
    amount_local_abs = parse_number_series(
        prepared_dataframe[required_columns["amount_local_currency"]]
    ).abs()

    debit = amount_local_abs.where(sap_indicator == "S", 0).fillna(0)
    credit = amount_local_abs.where(sap_indicator == "H", 0).fillna(0)

    return pd.DataFrame(
        {
            "TransId": prepared_dataframe["_TRANS_ID"],
            "Debit Total": debit,
            "Credit Total": credit,
        }
    )


def build_header_total_map(header_total_base_dataframe):
    if header_total_base_dataframe.empty:
        return pd.Series(dtype="float64")

    grouped = header_total_base_dataframe.groupby("TransId", dropna=False)[
        ["Debit Total", "Credit Total"]
    ].sum()

    grouped["Header Total Local"] = grouped[["Debit Total", "Credit Total"]].max(axis=1)

    return grouped["Header Total Local"]


def get_same_creator_approver_mask(prepared_dataframe, required_columns):
    creator_user = clean_text_series(prepared_dataframe[required_columns["creator_user"]]).str.upper()
    approver_user = clean_text_series(prepared_dataframe[required_columns["approver_user"]]).str.upper()

    return (creator_user != "") & (approver_user != "") & (creator_user == approver_user)


def build_lha_like_gl04_lines(
    prepared_dataframe,
    required_columns,
    optional_columns,
    master_dataframe,
):
    if prepared_dataframe.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    company_name_map = build_company_name_map(master_dataframe)
    company_currency_map = build_company_currency_map(master_dataframe)
    gl_account_name_map = build_gl_account_name_map(master_dataframe)

    sap_company_code = prepared_dataframe["_SAP_COMPANY_CODE"]
    sap_document_number = prepared_dataframe["_SAP_DOCUMENT_NUMBER"]
    sap_line_item = prepared_dataframe["_SAP_LINE_ITEM"]
    trans_id = prepared_dataframe["_TRANS_ID"]

    posting_date = prepared_dataframe["_POSTING_DATE"]
    document_date = parse_date_series(prepared_dataframe[required_columns["document_date"]])
    entry_date = prepared_dataframe["_ENTRY_DATE"]

    period_code = format_period_code_series(posting_date)
    period_from = get_month_start_series(posting_date)
    period_to = get_month_end_series(posting_date)
    posting_month = format_period_code_series(posting_date)

    sap_document_type = clean_text_series(prepared_dataframe[required_columns["document_type"]])
    sap_indicator = clean_text_series(
        prepared_dataframe[required_columns["debit_credit_indicator"]]
    ).str.upper()

    amount_local_abs = parse_number_series(
        prepared_dataframe[required_columns["amount_local_currency"]]
    ).abs()

    debit = amount_local_abs.where(sap_indicator == "S", pd.NA)
    credit = amount_local_abs.where(sap_indicator == "H", pd.NA)

    line_amount_local = amount_local_abs.copy()
    line_amount_local = line_amount_local.where(
        sap_indicator != "H",
        line_amount_local * -1,
    )

    debit_credit_indicator = sap_indicator.map({"S": "D", "H": "C"}).fillna(sap_indicator)
    gl_account = normalize_code_series(prepared_dataframe[required_columns["gl_account"]])

    reversal_document = normalize_code_series(
        get_optional_series(prepared_dataframe, optional_columns["reversal_document"])
    )
    reversal_year = normalize_code_series(
        get_optional_series(prepared_dataframe, optional_columns["reversal_year"])
    )

    reverses_trans_id = build_reverses_trans_id_series(
        company_code=sap_company_code,
        reversal_year=reversal_year,
        reversal_document=reversal_document,
    )

    journal_entry_status = pd.Series(
        "Normal Entry",
        index=prepared_dataframe.index,
        dtype="object",
    )
    journal_entry_status = journal_entry_status.mask(reversal_document != "", "Reversed")

    output_dataframe = pd.DataFrame(
        {
            "CoCo": sap_company_code,
            "Company": sap_company_code.map(company_name_map).fillna(""),
            "Company Main Currency": sap_company_code.map(company_currency_map).fillna(""),
            "Company System Currency": "",
            "Journal Number": sap_document_number,
            "TransId": trans_id,
            "Line": sap_line_item,
            "Journal Type": sap_document_type,
            "Posting Date": posting_date,
            "Document Date": document_date,
            "Entry Date": entry_date,
            "Update Date": entry_date,
            "Entry Weekday": prepared_dataframe["_ENTRY_WEEKDAY"],
            "Journal Memo": clean_text_series(prepared_dataframe[required_columns["header_text"]]),
            "Journal Entry Status": journal_entry_status,
            "Reverses TransId": reverses_trans_id,
            "Auto Reversal": "N",
            "Account Code": gl_account,
            "Account Name": gl_account.map(gl_account_name_map).fillna(""),
            "Line Memo": clean_text_series(prepared_dataframe[required_columns["line_text"]]),
            "Debit": debit,
            "Credit": credit,
            "Line Amount Local": line_amount_local,
            "FC Currency": clean_text_series(prepared_dataframe[required_columns["currency"]]),
            "Debit Credit Indicator": debit_credit_indicator,
            "Line Amount USD": pd.NA,
            "USD Method": "",
            "USD Rate": pd.NA,
            "USD Rate Date": pd.NaT,
            "Header Total Local": pd.NA,
            "Header Total USD": pd.NA,
            "Creator ID": clean_text_series(prepared_dataframe[required_columns["creator_user"]]),
            "Creator Name": "",
            "Approver ID": clean_text_series(prepared_dataframe[required_columns["approver_user"]]),
            "Approver Name": "",
            "Period Code": period_code,
            "Period Name": period_code,
            "Period From": period_from,
            "Period To": period_to,
            "Current Period Status": "",
            "Latest Period Log Status": "",
            "Latest Period Log Date": pd.NaT,
            "Period Open Date": pd.NaT,
            "Period Close Date": pd.NaT,
            "Latest Period Open Date": pd.NaT,
            "Latest Period Close Date": pd.NaT,
            "Days From Period End": calculate_days_from_period_end_series(
                posting_date_series=posting_date,
                period_to_series=period_to,
            ),
            "Posted Before Period Open": "",
            "Posting Month": posting_month,
        }
    )

    return output_dataframe[OUTPUT_COLUMNS].copy()


def add_usd_fields(line_dataframe, fx_dataframe):
    result = line_dataframe.copy()

    if result.empty:
        return result

    if fx_dataframe.empty:
        return result

    normalized_fx_dataframe = normalize_fx_rates(fx_dataframe)
    fx_cache = {}

    line_amount_usd_values = []
    usd_method_values = []
    usd_rate_values = []
    usd_rate_date_values = []
    header_total_usd_values = []

    for _, row in result.iterrows():
        currency = normalize_text(row.get("FC Currency", "")).upper()
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

    result["Company System Currency"] = "USD"
    result["Line Amount USD"] = line_amount_usd_values
    result["USD Method"] = usd_method_values
    result["USD Rate"] = usd_rate_values
    result["USD Rate Date"] = usd_rate_date_values
    result["Header Total USD"] = header_total_usd_values

    missing_usd = result["Line Amount USD"].isna().sum()

    if missing_usd > 0:
        print()
        print("WARNING: Some GL FX rates were not found.")
        print("USD columns are blank for those rows.")
        print(f"Rows without USD amount: {missing_usd}")
        print()

    return result


def create_gl04_lines_for_source(source_dataframe, master_dataframe, context, source_name):
    if source_dataframe.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS), pd.DataFrame()

    prepared_dataframe, required_columns, optional_columns = prepare_source_dataframe(
        source_dataframe=source_dataframe,
        context=context,
        source_name=source_name,
    )

    if prepared_dataframe.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS), pd.DataFrame()

    header_total_base_dataframe = build_header_total_base(
        prepared_dataframe=prepared_dataframe,
        required_columns=required_columns,
    )

    print(f"{source_name} identifying same creator and approver journal lines before output build...")

    same_user_mask = get_same_creator_approver_mask(
        prepared_dataframe=prepared_dataframe,
        required_columns=required_columns,
    )
    same_user_dataframe = prepared_dataframe[same_user_mask].copy()

    print(f"{source_name} same creator and approver journal lines kept: {len(same_user_dataframe)}")

    if same_user_dataframe.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS), header_total_base_dataframe

    print(f"{source_name} creating LHA-like GL04 line rows...")

    output_dataframe = build_lha_like_gl04_lines(
        prepared_dataframe=same_user_dataframe,
        required_columns=required_columns,
        optional_columns=optional_columns,
        master_dataframe=master_dataframe,
    )

    print(f"{source_name} LHA-like GL04 line rows created: {len(output_dataframe)}")

    return output_dataframe, header_total_base_dataframe


def create_gl04_same_creator_approver_journals(
    bsis_dataframe,
    bsas_dataframe,
    master_dataframe,
    fx_dataframe,
    context,
):
    output_frames = []
    header_total_base_frames = []

    for source_name, source_dataframe in (
        ("BSIS", bsis_dataframe),
        ("BSAS", bsas_dataframe),
    ):
        source_output_dataframe, source_header_total_base_dataframe = create_gl04_lines_for_source(
            source_dataframe=source_dataframe,
            master_dataframe=master_dataframe,
            context=context,
            source_name=source_name,
        )

        if not source_output_dataframe.empty:
            output_frames.append(source_output_dataframe)

        if not source_header_total_base_dataframe.empty:
            header_total_base_frames.append(source_header_total_base_dataframe)

    if output_frames:
        result = pd.concat(output_frames, ignore_index=True)
    else:
        result = pd.DataFrame(columns=OUTPUT_COLUMNS)

    if header_total_base_frames:
        header_total_base_dataframe = pd.concat(header_total_base_frames, ignore_index=True)
        header_total_map = build_header_total_map(header_total_base_dataframe)

        if not result.empty:
            result["Header Total Local"] = result["TransId"].map(header_total_map)

    if not result.empty:
        result = result.sort_values(
            by=["CoCo", "Posting Date", "Journal Number", "Line"],
            kind="stable",
        ).reset_index(drop=True)

    if not result.empty and not fx_dataframe.empty:
        print("Applying optional FxRates to GL04 result rows...")
        result = add_usd_fields(result, fx_dataframe)
    elif fx_dataframe.empty:
        print("FxRates input was not found. GL04 USD columns will remain blank.")

    return result[OUTPUT_COLUMNS].copy()


def write_gl04_output(output_dataframe, context):
    output_file = get_gl_output_file(context)

    print(f"Output workbook: {output_file}")

    if not output_file.exists():
        print("Output workbook does not exist. Using fast GL04 writer...")

        fast_written = write_single_sheet_workbook_fast(
            output_file=output_file,
            sheet_name=SHEET_NAME,
            dataframe=output_dataframe,
            date_columns=DATE_COLUMNS,
            amount_columns=AMOUNT_COLUMNS,
            integer_columns=INTEGER_COLUMNS,
        )

        if fast_written:
            print(f"GL04 rows written: {len(output_dataframe)}")
            print(f"GL output workbook: {output_file}")
            print()
            return output_file

    print("Output workbook already exists or fast writer is unavailable.")
    print("Using preserve-sheets openpyxl writer...")

    workbook = open_or_create_gl_output_workbook(output_file)
    worksheet = recreate_gl_sheet(workbook, SHEET_NAME)

    print("Writing GL04 rows to worksheet...")
    write_dataframe_to_sheet(
        worksheet=worksheet,
        dataframe=output_dataframe,
    )

    print("Applying GL04 formatting...")
    apply_standard_gl_formatting(
        worksheet=worksheet,
        dataframe=output_dataframe,
        date_columns=DATE_COLUMNS,
        amount_columns=AMOUNT_COLUMNS,
        integer_columns=INTEGER_COLUMNS,
    )

    print("Saving GL output workbook...")
    save_gl_output_workbook(workbook, output_file)

    print(f"GL04 rows written: {len(output_dataframe)}")
    print(f"GL output workbook: {output_file}")
    print()

    return output_file


def run_gl_004(context):
    """
    Execute GL_004 and write only the GL04 sheet.
    """
    print_header("Running GL_004 - General Journals Created And Approved By The Same User")
    print("Logic: Creator ID = Pré-edição / PPNAM; Approver ID = Nome do usuário / USNAM")

    bsis_dataframe = load_gl_bsis_data(context)
    bsas_dataframe = load_gl_bsas_data(context)
    master_dataframe = load_gl_master_data(context)
    fx_dataframe = load_gl_fx_rates_data(context)

    if bsis_dataframe.empty and bsas_dataframe.empty:
        raise FileNotFoundError(
            "GL04 requires at least one GL journal input file:\n"
            "- input/LBR GL_JE_BSIS_YYYYMMDD.xlsx\n"
            "- input/LBR GL_JE_BSAS_YYYYMMDD.xlsx"
        )

    output_dataframe = create_gl04_same_creator_approver_journals(
        bsis_dataframe=bsis_dataframe,
        bsas_dataframe=bsas_dataframe,
        master_dataframe=master_dataframe,
        fx_dataframe=fx_dataframe,
        context=context,
    )

    write_gl04_output(
        output_dataframe=output_dataframe,
        context=context,
    )

    print(f"GL04 rows written: {len(output_dataframe)}")

    return output_dataframe
