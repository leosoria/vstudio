"""
GL_001 - Suspicious General Journals: Entered On The Weekend.
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


SHEET_NAME = "GL01"


REQUIRED_COLUMNS = {
    "company_code": ["Empr", "BUKRS"],
    "document_number": ["Nº doc.", "Nº doc", "BELNR"],
    "fiscal_year": ["Ano", "GJAHR"],
    "document_type": ["Tp.doc.", "Tp.doc", "BLART"],
    "document_date": ["Data doc.", "Data doc", "BLDAT"],
    "posting_date": ["Dt.lçto.", "Dt.lçto", "Dt.lcto.", "Dt.lcto", "BUDAT"],
    "entry_date": ["Dt.entr.", "Dt.entr", "CPUDT"],
    "entry_time": ["Hora", "CPUTM"],
    "user": ["Nome do usuário", "Nome do usuario", "USNAM"],
    "transaction_code": ["CódT", "CodT", "TCODE"],
    "header_text": ["Texto cabeçalho documento", "Texto cabecalho documento", "BKTXT"],
    "currency": ["Moeda", "WAERS"],
    "reference": ["Referência", "Referencia", "XBLNR"],
    "line_item": ["Itm", "BUZEI"],
    "gl_account": ["Razão", "Razao", "HKONT"],
    "debit_credit_indicator": ["D/C", "SHKZG"],
    "amount_local_currency": ["Montante em MI", "DMBTR"],
    "amount_document_currency": ["Montante", "WRBTR"],
    "line_text": ["Texto", "SGTXT"],
}


OPTIONAL_COLUMNS = {
    "cost_center": ["Centro cst", "KOSTL"],
    "profit_center": ["Cen.lucro", "PRCTR"],
    "order": ["Ordem", "AUFNR"],
    "assignment": ["Atribuição", "Atribuicao", "ZUONR"],
    "clearing_document": ["DocCompens", "AUGBL"],
    "clearing_date": ["Compensaç.", "Compensac.", "Compensação", "Compensacao", "AUGDT"],
    "reversal_document": ["Estorno c/", "Estorno c", "STBLG"],
    "reversal_year": ["Ano.2", "STJAH"],
    "reversal_reason": ["Motiv.est", "Motiv est", "STGRD"],
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
    "Source",
    "SAP Company Code",
    "SAP Fiscal Year",
    "SAP Document Number",
    "SAP Line Item",
    "SAP Document Type",
    "SAP Reference",
    "SAP Header Text",
    "SAP Clearing Document",
    "SAP Clearing Date",
    "SAP Reversal Document",
    "SAP Reversal Year",
    "SAP Reversal Reason",
    "Document Key",
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
    "SAP Clearing Date",
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
    return pd.to_datetime(
        series,
        errors="coerce",
        dayfirst=True,
    )


def parse_number_series(series):
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    text_series = clean_text_series(series)
    text_series = text_series.str.replace(" ", "", regex=False)

    has_comma = text_series.str.contains(",", regex=False, na=False)
    has_dot = text_series.str.contains(".", regex=False, na=False)

    comma_decimal_mask = has_comma & ~has_dot
    text_series = text_series.where(
        ~comma_decimal_mask,
        text_series.str.replace(",", ".", regex=False),
    )

    both_mask = has_comma & has_dot
    last_comma = text_series.str.rfind(",")
    last_dot = text_series.str.rfind(".")

    european_mask = both_mask & (last_comma > last_dot)
    text_series = text_series.where(
        ~european_mask,
        text_series.str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
    )

    us_mask = both_mask & (last_dot > last_comma)
    text_series = text_series.where(
        ~us_mask,
        text_series.str.replace(",", "", regex=False),
    )

    return pd.to_numeric(text_series, errors="coerce")


def build_trans_id_series(company_code, fiscal_year, document_number):
    return (
        company_code.astype(str)
        + "|"
        + fiscal_year.astype(str)
        + "|"
        + document_number.astype(str)
    )


def build_document_key_series(company_code, fiscal_year, document_number, line_item):
    return (
        company_code.astype(str)
        + "|"
        + fiscal_year.astype(str)
        + "|"
        + document_number.astype(str)
        + "|"
        + line_item.astype(str)
    )


def build_reverses_trans_id_series(company_code, reversal_year, reversal_document):
    reversal_document = normalize_code_series(reversal_document)
    reversal_year = normalize_code_series(reversal_year)

    reverses_trans_id = (
        company_code.astype(str)
        + "|"
        + reversal_year.astype(str)
        + "|"
        + reversal_document.astype(str)
    )

    reverses_trans_id = reverses_trans_id.mask(reversal_document == "", "")

    return reverses_trans_id


def format_period_code_series(date_series):
    return date_series.dt.strftime("%Y-%m").fillna("")


def get_month_start_series(date_series):
    return date_series.dt.to_period("M").dt.start_time


def get_month_end_series(date_series):
    return date_series.dt.to_period("M").dt.end_time.dt.normalize()


def calculate_days_from_period_end_series(posting_date_series, period_to_series):
    return (
        posting_date_series.dt.normalize()
        - period_to_series.dt.normalize()
    ).dt.days


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

    result = result[
        (posting_date >= from_date)
        & (posting_date <= to_date)
    ].copy()

    posting_date = parse_date_series(result[required_columns["posting_date"]])

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
    result["_DOCUMENT_KEY"] = build_document_key_series(
        company_code=company_code,
        fiscal_year=fiscal_year,
        document_number=document_number,
        line_item=line_item,
    )
    result["_SOURCE"] = source_name

    return result, required_columns, optional_columns


def keep_weekend_trans_id_lines(prepared_dataframe):
    if prepared_dataframe.empty:
        return prepared_dataframe

    header_dataframe = prepared_dataframe.drop_duplicates(
        subset=["_TRANS_ID"],
        keep="first",
    )

    weekend_trans_ids = set(
        header_dataframe.loc[
            header_dataframe["_ENTRY_WEEKDAY"].isin([5, 6]),
            "_TRANS_ID",
        ]
    )

    return prepared_dataframe[
        prepared_dataframe["_TRANS_ID"].isin(weekend_trans_ids)
    ].copy()


def build_lha_like_weekend_lines(
    prepared_dataframe,
    required_columns,
    optional_columns,
    master_dataframe,
    context,
):
    if prepared_dataframe.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    company_name_map = build_company_name_map(master_dataframe)
    company_currency_map = build_company_currency_map(master_dataframe)
    gl_account_name_map = build_gl_account_name_map(master_dataframe)

    sap_company_code = prepared_dataframe["_SAP_COMPANY_CODE"]
    sap_fiscal_year = prepared_dataframe["_SAP_FISCAL_YEAR"]
    sap_document_number = prepared_dataframe["_SAP_DOCUMENT_NUMBER"]
    sap_line_item = prepared_dataframe["_SAP_LINE_ITEM"]

    trans_id = prepared_dataframe["_TRANS_ID"]
    document_key = prepared_dataframe["_DOCUMENT_KEY"]

    posting_date = prepared_dataframe["_POSTING_DATE"]
    document_date = parse_date_series(prepared_dataframe[required_columns["document_date"]])
    entry_date = prepared_dataframe["_ENTRY_DATE"]

    period_from = get_month_start_series(posting_date)
    period_to = get_month_end_series(posting_date)

    sap_document_type = clean_text_series(prepared_dataframe[required_columns["document_type"]])
    sap_indicator = clean_text_series(
        prepared_dataframe[required_columns["debit_credit_indicator"]]
    ).str.upper()

    amount_local_raw = prepared_dataframe[required_columns["amount_local_currency"]]
    amount_local_abs = parse_number_series(amount_local_raw).abs()

    debit = amount_local_abs.where(sap_indicator == "S", pd.NA)
    credit = amount_local_abs.where(sap_indicator == "H", pd.NA)

    line_amount_local = amount_local_abs.copy()
    line_amount_local = line_amount_local.where(sap_indicator != "H", line_amount_local * -1)

    debit_credit_indicator = sap_indicator.map(
        {
            "S": "D",
            "H": "C",
        }
    ).fillna(sap_indicator)

    gl_account = normalize_code_series(prepared_dataframe[required_columns["gl_account"]])

    reversal_document = normalize_code_series(
        get_optional_series(prepared_dataframe, optional_columns["reversal_document"])
    )
    reversal_year = normalize_code_series(
        get_optional_series(prepared_dataframe, optional_columns["reversal_year"])
    )
    reversal_reason = clean_text_series(
        get_optional_series(prepared_dataframe, optional_columns["reversal_reason"])
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
    journal_entry_status = journal_entry_status.mask(
        reversal_document != "",
        "Reversed",
    )

    period_code = format_period_code_series(posting_date)
    posting_month = format_period_code_series(posting_date)

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
            "Creator ID": clean_text_series(prepared_dataframe[required_columns["user"]]),
            "Creator Name": "",
            "Approver ID": "",
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
            "Posted Before Period Open": "N",
            "Posting Month": posting_month,
            "Source": prepared_dataframe["_SOURCE"],
            "SAP Company Code": sap_company_code,
            "SAP Fiscal Year": sap_fiscal_year,
            "SAP Document Number": sap_document_number,
            "SAP Line Item": sap_line_item,
            "SAP Document Type": sap_document_type,
            "SAP Reference": clean_text_series(prepared_dataframe[required_columns["reference"]]),
            "SAP Header Text": clean_text_series(prepared_dataframe[required_columns["header_text"]]),
            "SAP Clearing Document": normalize_code_series(
                get_optional_series(prepared_dataframe, optional_columns["clearing_document"])
            ),
            "SAP Clearing Date": parse_date_series(
                get_optional_series(prepared_dataframe, optional_columns["clearing_date"])
            ),
            "SAP Reversal Document": reversal_document,
            "SAP Reversal Year": reversal_year,
            "SAP Reversal Reason": reversal_reason,
            "Document Key": document_key,
        }
    )

    return output_dataframe[OUTPUT_COLUMNS].copy()


def add_header_totals(line_dataframe):
    result = line_dataframe.copy()

    if result.empty:
        return result

    debit_by_trans_id = (
        pd.to_numeric(result["Debit"], errors="coerce")
        .fillna(0)
        .groupby(result["TransId"])
        .sum()
    )

    credit_by_trans_id = (
        pd.to_numeric(result["Credit"], errors="coerce")
        .fillna(0)
        .groupby(result["TransId"])
        .sum()
    )

    header_total_by_trans_id = pd.concat(
        [
            debit_by_trans_id.rename("Debit Total"),
            credit_by_trans_id.rename("Credit Total"),
        ],
        axis=1,
    ).fillna(0)

    header_total_by_trans_id["Header Total Local"] = header_total_by_trans_id[
        ["Debit Total", "Credit Total"]
    ].max(axis=1)

    result["Header Total Local"] = result["TransId"].map(
        header_total_by_trans_id["Header Total Local"]
    )

    return result


def create_weekend_lines_for_source(
    source_dataframe,
    master_dataframe,
    context,
    source_name,
):
    if source_dataframe.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    prepared_dataframe, required_columns, optional_columns = prepare_source_dataframe(
        source_dataframe=source_dataframe,
        context=context,
        source_name=source_name,
    )

    if prepared_dataframe.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    print(f"{source_name} identifying weekend journals...")
    weekend_lines = keep_weekend_trans_id_lines(prepared_dataframe)

    print(f"{source_name} weekend journal lines kept: {len(weekend_lines)}")

    if weekend_lines.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    print(f"{source_name} creating LHA-like weekend line rows...")
    output_dataframe = build_lha_like_weekend_lines(
        prepared_dataframe=weekend_lines,
        required_columns=required_columns,
        optional_columns=optional_columns,
        master_dataframe=master_dataframe,
        context=context,
    )

    print(f"{source_name} LHA-like weekend line rows created: {len(output_dataframe)}")

    return output_dataframe


def create_gl01_weekend_journals(bsis_dataframe, bsas_dataframe, master_dataframe, context):
    if bsis_dataframe.empty and bsas_dataframe.empty:
        raise FileNotFoundError(
            "No GL input files were found. Expected at least one of:\n"
            "- input/LBR GL_JE_BSIS_YYYYMMDD.xlsx\n"
            "- input/LBR GL_JE_BSAS_YYYYMMDD.xlsx"
        )

    source_dataframes = []

    bsis_weekend_lines = create_weekend_lines_for_source(
        source_dataframe=bsis_dataframe,
        master_dataframe=master_dataframe,
        context=context,
        source_name="BSIS",
    )

    if not bsis_weekend_lines.empty:
        source_dataframes.append(bsis_weekend_lines)

    bsas_weekend_lines = create_weekend_lines_for_source(
        source_dataframe=bsas_dataframe,
        master_dataframe=master_dataframe,
        context=context,
        source_name="BSAS",
    )

    if not bsas_weekend_lines.empty:
        source_dataframes.append(bsas_weekend_lines)

    if not source_dataframes:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    weekend_line_dataframe = pd.concat(
        source_dataframes,
        ignore_index=True,
    )

    weekend_line_dataframe = add_header_totals(weekend_line_dataframe)

    weekend_header_dataframe = weekend_line_dataframe.drop_duplicates(
        subset=["TransId"],
        keep="first",
    ).copy()

    weekend_header_dataframe = weekend_header_dataframe.sort_values(
        by=[
            "Entry Date",
            "CoCo",
            "Journal Number",
        ],
        kind="stable",
    ).reset_index(drop=True)

    return weekend_header_dataframe[OUTPUT_COLUMNS].copy()


def run_gl_001(context):
    print_header("Running GL_001 - Suspicious General Journals: Entered On The Weekend")

    bsis_dataframe = load_gl_bsis_data(context)
    print(f"BSIS rows loaded: {len(bsis_dataframe)}")

    bsas_dataframe = load_gl_bsas_data(context)
    print(f"BSAS rows loaded: {len(bsas_dataframe)}")

    master_dataframe = load_gl_master_data(context)
    print(f"GL master rows loaded: {len(master_dataframe)}")

    print("Creating GL01 weekend journals with optimized logic...")
    output_dataframe = create_gl01_weekend_journals(
        bsis_dataframe=bsis_dataframe,
        bsas_dataframe=bsas_dataframe,
        master_dataframe=master_dataframe,
        context=context,
    )

    print(f"GL01 weekend journal rows prepared: {len(output_dataframe)}")

    output_file = get_gl_output_file(context)

    print(f"Output workbook: {output_file}")

    if not output_file.exists():
        print("Output workbook does not exist. Using fast GL01 writer...")
        fast_written = write_single_sheet_workbook_fast(
            output_file=output_file,
            sheet_name=SHEET_NAME,
            dataframe=output_dataframe,
            date_columns=DATE_COLUMNS,
            amount_columns=AMOUNT_COLUMNS,
            integer_columns=INTEGER_COLUMNS,
        )

        if fast_written:
            print(f"GL01 rows written: {len(output_dataframe)}")
            print(f"GL output workbook: {output_file}")
            print()
            return

    print("Output workbook already exists or fast writer is unavailable.")
    print("Using preserve-sheets openpyxl writer...")

    workbook = open_or_create_gl_output_workbook(output_file)
    worksheet = recreate_gl_sheet(workbook, SHEET_NAME)

    print("Writing GL01 rows to worksheet...")
    write_dataframe_to_sheet(worksheet, output_dataframe)

    print("Applying GL01 formatting...")
    apply_standard_gl_formatting(
        worksheet=worksheet,
        dataframe=output_dataframe,
        date_columns=DATE_COLUMNS,
        amount_columns=AMOUNT_COLUMNS,
        integer_columns=INTEGER_COLUMNS,
    )

    print("Saving GL output workbook...")
    save_gl_output_workbook(workbook, output_file)

    print(f"GL01 rows written: {len(output_dataframe)}")
    print(f"GL output workbook: {output_file}")
    print()
