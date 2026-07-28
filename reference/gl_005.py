"""
GL_005 - General Ledger Accounts Created In The Last Three Months.

LHA logic:
- Extracts general ledger accounts created in the last three months.
- Uses LBR GL_MD as the main source.
- FxRates is optional and is only used if balance columns are available in the master data.

This control is independent and writes/replaces only the GL05 worksheet.
"""

import pandas as pd

from core.gl_common import (
    apply_standard_gl_formatting,
    get_gl_output_file,
    get_optional_column,
    load_gl_fx_rates_data,
    load_gl_master_data,
    normalize_fx_rates,
    normalize_text,
    open_or_create_gl_output_workbook,
    recreate_gl_sheet,
    save_gl_output_workbook,
    select_fx_rate_to_usd,
    to_datetime_value,
    write_dataframe_to_sheet,
    write_single_sheet_workbook_fast,
)


SHEET_NAME = "GL05"


OUTPUT_COLUMNS = [
    "CoCo",
    "Company",
    "Account Code",
    "Account Name",
    "Dimension 1 Relevant",
    "Loading Factor Code",
    "Creation Date",
    "Date of Update",
    "Details",
    "Control Account",
    "Active",
    "Active From",
    "Active To",
    "Inactive",
    "Inactive From",
    "Inactive To",
    "User Signature",
    "User Code",
    "User Name",
    "Account Currency",
    "Company Main Currency",
    "Company System Currency",
    "Balance",
    "Balance System Currency",
    "Balance USD",
    "USD Rate",
    "USD Rate Date",
    "USD Method",
]


DATE_COLUMNS = {
    "Creation Date",
    "Date of Update",
    "Active From",
    "Active To",
    "Inactive From",
    "Inactive To",
    "USD Rate Date",
}


AMOUNT_COLUMNS = {
    "Balance",
    "Balance System Currency",
    "Balance USD",
    "USD Rate",
}


REQUIRED_COLUMNS = {
    "company_code": ["Empr", "BUKRS", "T001-BUKRS", "SKB1-BUKRS"],
    "company_name": ["Nome da firma", "BUTXT", "T001-BUTXT", "Company", "Company Name"],
    "company_currency": ["Moeda", "WAERS", "T001-WAERS", "Company Main Currency"],
    "account_code": [
        "Cta.Razão",
        "Cta.Razão.1",
        "Cta.Razão.2",
        "Cta.Razao",
        "Cta.Razao.1",
        "Cta.Razao.2",
        "SAKNR",
        "SKA1-SAKNR",
        "SKB1-SAKNR",
        "SKAT-SAKNR",
        "Account Code",
    ],
    "creation_date": ["Data", "ERDAT", "SKA1-ERDAT", "Creation Date", "Create Date"],
}


OPTIONAL_COLUMNS = {
    "chart_of_accounts": ["PlCt", "PlCt.1", "PlCt.2", "KTOPL", "SKA1-KTOPL", "SKB1-KTOPL"],
    "short_text": ["Texto breve", "TXT20", "SKAT-TXT20", "Short Text"],
    "long_text": ["TxtDescr", "TXT50", "SKAT-TXT50", "Description", "Account Name"],
    "language": ["Idioma", "SPRAS", "SKAT-SPRAS"],
    "created_by": ["Autor", "ERNAM", "SKA1-ERNAM", "User Code", "Create User ID"],
    "update_date": [
        "Data modificação",
        "Data modificacao",
        "Data de modificação",
        "Data de modificacao",
        "AEDAT",
        "SKA1-AEDAT",
        "Date of Update",
    ],
    "balance_account_indicator": ["CtaBalnç", "CtaBalnc", "XBILK", "SKA1-XBILK"],
    "account_group": ["Nº cta.grp", "Nº cta grp", "No cta.grp", "KTOKS", "SKA1-KTOKS"],
    "account_currency": ["Moeda conta", "Moeda da conta", "SKB1-WAERS", "Account Currency"],
    "system_currency": ["Moeda sistema", "Moeda do sistema", "Company System Currency", "System Currency"],
    "dimension_1_relevant": ["Dimension 1 Relevant", "Dimensão 1 relevante", "Dimensao 1 relevante"],
    "loading_factor_code": ["Loading Factor Code", "Código fator carga", "Codigo fator carga"],
    "details": ["Details", "Detalhes"],
    "control_account": ["Control Account", "Conta controle", "Conta de controle", "MITKZ", "SKB1-MITKZ"],
    "active": ["Active", "Ativo"],
    "active_from": ["Active From", "Ativo desde"],
    "active_to": ["Active To", "Ativo até", "Ativo ate"],
    "inactive": ["Inactive", "Inativo"],
    "inactive_from": ["Inactive From", "Inativo desde"],
    "inactive_to": ["Inactive To", "Inativo até", "Inativo ate"],
    "user_signature": ["User Signature", "Assinatura usuário", "Assinatura usuario"],
    "user_name": ["User Name", "Nome usuário", "Nome usuario", "Nome do usuário", "Nome do usuario"],
    "balance": ["Balance", "Saldo", "Saldo ML", "Balance Local", "Saldo moeda local"],
    "balance_system_currency": [
        "Balance System Currency",
        "Saldo moeda sistema",
        "Saldo em moeda sistema",
        "Saldo USD",
    ],
}


def print_header(title):
    print(title)
    print("-" * len(title))


def blank_series(index):
    return pd.Series("", index=index, dtype="object")


def blank_number_series(index):
    return pd.Series(pd.NA, index=index, dtype="Float64")


def blank_date_series(index):
    return pd.Series(pd.NaT, index=index, dtype="datetime64[ns]")


def clean_text_series(series):
    result = series.copy()
    result = result.where(result.notna(), "")
    result = result.astype(str).str.strip()
    result = result.mask(result.str.lower() == "nan", "")
    result = result.str.replace(r"\.0$", "", regex=True)

    return result


def normalize_company_series(series):
    result = clean_text_series(series)
    numeric_mask = result.str.fullmatch(r"\d+", na=False)

    if numeric_mask.any():
        result = result.where(
            ~numeric_mask,
            result.astype("Int64").astype(str),
        )

    return result


def normalize_code_series(series):
    return clean_text_series(series)


def parse_date_series(series):
    if series is None:
        return pd.Series(dtype="datetime64[ns]")

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
    if series is None:
        return pd.Series(dtype="Float64")

    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    text_series = clean_text_series(series)
    text_series = text_series.str.replace("\u00a0", "", regex=False)
    text_series = text_series.str.replace(" ", "", regex=False)

    negative_parentheses_mask = text_series.str.fullmatch(r"\(.*\)", na=False)
    text_series = text_series.str.replace("(", "-", regex=False)
    text_series = text_series.str.replace(")", "", regex=False)
    text_series = text_series.mask(negative_parentheses_mask & text_series.eq("-"), "")

    dash_mask = text_series.isin(["-", "--"])
    text_series = text_series.mask(dash_mask, "")

    comma_decimal_mask = text_series.str.contains(",", regex=False, na=False)
    dot_decimal_mask = text_series.str.contains(r"\.\d{1,6}$", regex=True, na=False)

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


def require_columns(dataframe, required_columns, source_name):
    resolved_columns = {}
    missing_columns = []

    for logical_name, possible_names in required_columns.items():
        column_name = get_optional_column(dataframe, possible_names)

        if column_name is None:
            missing_columns.append(f"{logical_name}: {', '.join(possible_names)}")
        else:
            resolved_columns[logical_name] = column_name

    if missing_columns:
        raise ValueError(
            f"{source_name} is missing required columns:\n- "
            + "\n- ".join(missing_columns)
        )

    return resolved_columns


def resolve_optional_columns(dataframe):
    resolved_columns = {}

    for logical_name, possible_names in OPTIONAL_COLUMNS.items():
        resolved_columns[logical_name] = get_optional_column(dataframe, possible_names)

    return resolved_columns


def get_optional_text_series(dataframe, column_name):
    if column_name is None:
        return blank_series(dataframe.index)

    return clean_text_series(dataframe[column_name])


def get_optional_date_series(dataframe, column_name):
    if column_name is None:
        return blank_date_series(dataframe.index)

    return parse_date_series(dataframe[column_name])


def get_optional_number_series(dataframe, column_name):
    if column_name is None:
        return blank_number_series(dataframe.index)

    return parse_number_series(dataframe[column_name])


def first_non_blank_text(*series_list):
    if not series_list:
        return pd.Series(dtype="object")

    result = blank_series(series_list[0].index)

    for series in series_list:
        text_series = clean_text_series(series)
        result = result.mask(result == "", text_series)

    return result


def get_analysis_date(context):
    module_config = context.get("module", {})

    analysis_date = to_datetime_value(module_config.get("to", ""))

    if analysis_date is not None and not pd.isna(analysis_date):
        return pd.Timestamp(analysis_date).normalize()

    return pd.Timestamp.today().normalize()


def filter_by_company_if_needed(dataframe, company_column, context):
    module_config = context.get("module", {})
    companies_filter = normalize_text(module_config.get("companies", ""))

    if companies_filter == "":
        return dataframe.copy()

    normalized_filter = companies_filter

    for separator in [";", "|", "\n", "\r", "\t"]:
        normalized_filter = normalized_filter.replace(separator, ",")

    if "," not in normalized_filter and " " in normalized_filter:
        normalized_filter = ",".join(normalized_filter.split())

    requested_companies = set()

    for value in normalized_filter.split(","):
        value_text = normalize_text(value)

        if value_text == "":
            continue

        if value_text.upper() in ["ALL", "TODAS", "TODOS"]:
            return dataframe.copy()

        requested_companies.add(value_text.upper())

        if value_text.isdigit():
            requested_companies.add(str(int(value_text)))
            requested_companies.add(value_text.zfill(4))

    if not requested_companies:
        return dataframe.copy()

    raw_company_values = clean_text_series(dataframe[company_column]).str.upper()
    normalized_company_values = normalize_company_series(dataframe[company_column]).str.upper()
    padded_company_values = normalized_company_values.mask(
        ~normalized_company_values.str.fullmatch(r"\d+", na=False),
        normalized_company_values,
    )
    padded_company_values = padded_company_values.mask(
        normalized_company_values.str.fullmatch(r"\d+", na=False),
        normalized_company_values.str.zfill(4),
    )

    company_filter = raw_company_values.isin(requested_companies)
    company_filter = company_filter | normalized_company_values.isin(requested_companies)
    company_filter = company_filter | padded_company_values.isin(requested_companies)

    result = dataframe[company_filter].copy()

    if result.empty:
        print()
        print("WARNING: GL05 company filter returned no rows.")
        print(f"Configured companies filter: {companies_filter}")
        print("GL05 will continue without applying the company filter to avoid a headers-only output.")
        print("Please review config.xlsx COMPANIES for the GL module.")
        print()

        return dataframe.copy()

    return result


def filter_recent_accounts_with_fallback(filtered_master_dataframe, creation_date, context):
    analysis_date = get_analysis_date(context)
    threshold_date = analysis_date - pd.DateOffset(months=3)

    print(f"GL05 analysis date from config: {analysis_date.strftime('%d/%m/%Y')}")
    print(f"GL05 creation date threshold from config: {threshold_date.strftime('%d/%m/%Y')}")

    creation_date_filter = creation_date.notna()
    creation_date_filter = creation_date_filter & (creation_date >= threshold_date)
    creation_date_filter = creation_date_filter & (creation_date <= analysis_date)

    if creation_date_filter.any():
        return filtered_master_dataframe[creation_date_filter].copy(), analysis_date, threshold_date

    latest_creation_date = creation_date.dropna().max()

    if pd.isna(latest_creation_date):
        return filtered_master_dataframe[creation_date_filter].copy(), analysis_date, threshold_date

    fallback_analysis_date = pd.Timestamp(latest_creation_date).normalize()
    fallback_threshold_date = fallback_analysis_date - pd.DateOffset(months=3)
    fallback_filter = creation_date.notna()
    fallback_filter = fallback_filter & (creation_date >= fallback_threshold_date)
    fallback_filter = fallback_filter & (creation_date <= fallback_analysis_date)

    if fallback_filter.any():
        print()
        print("WARNING: GL05 found no accounts using the config analysis date.")
        print("GL05 is falling back to the latest Creation Date found in the GL master input.")
        print(f"Fallback analysis date: {fallback_analysis_date.strftime('%d/%m/%Y')}")
        print(f"Fallback creation date threshold: {fallback_threshold_date.strftime('%d/%m/%Y')}")
        print("Please review config.xlsx TO date if this is not expected.")
        print()

        return (
            filtered_master_dataframe[fallback_filter].copy(),
            fallback_analysis_date,
            fallback_threshold_date,
        )

    return filtered_master_dataframe[creation_date_filter].copy(), analysis_date, threshold_date


def choose_account_name(master_dataframe, optional_columns):
    long_text = get_optional_text_series(master_dataframe, optional_columns["long_text"])
    short_text = get_optional_text_series(master_dataframe, optional_columns["short_text"])

    return first_non_blank_text(long_text, short_text)


def choose_details(master_dataframe, optional_columns, account_name):
    details = get_optional_text_series(master_dataframe, optional_columns["details"])

    if optional_columns["details"] is not None:
        return details

    short_text = get_optional_text_series(master_dataframe, optional_columns["short_text"])

    return short_text.mask(short_text == account_name, "")


def derive_active_series(master_dataframe, optional_columns):
    active = get_optional_text_series(master_dataframe, optional_columns["active"]).str.upper()
    inactive = get_optional_text_series(master_dataframe, optional_columns["inactive"]).str.upper()

    if optional_columns["active"] is not None:
        return active.replace({"TRUE": "Y", "YES": "Y", "SIM": "Y", "S": "Y", "1": "Y"})

    if optional_columns["inactive"] is not None:
        inactive_normalized = inactive.replace(
            {"TRUE": "Y", "YES": "Y", "SIM": "Y", "S": "Y", "1": "Y"}
        )
        return pd.Series("Y", index=master_dataframe.index, dtype="object").mask(
            inactive_normalized == "Y",
            "N",
        )

    return pd.Series("Y", index=master_dataframe.index, dtype="object")


def derive_inactive_series(master_dataframe, optional_columns, active):
    inactive = get_optional_text_series(master_dataframe, optional_columns["inactive"]).str.upper()

    if optional_columns["inactive"] is not None:
        return inactive.replace({"TRUE": "Y", "YES": "Y", "SIM": "Y", "S": "Y", "1": "Y"})

    return pd.Series("N", index=master_dataframe.index, dtype="object").mask(active == "N", "Y")


def build_gl05_accounts(master_dataframe, context):
    if master_dataframe.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    print(f"GL master rows before column resolution: {len(master_dataframe)}")

    required_columns = require_columns(
        dataframe=master_dataframe,
        required_columns=REQUIRED_COLUMNS,
        source_name="GL05 master data",
    )
    optional_columns = resolve_optional_columns(master_dataframe)

    print("Resolving GL05 master data fields...")

    filtered_master_dataframe = filter_by_company_if_needed(
        dataframe=master_dataframe,
        company_column=required_columns["company_code"],
        context=context,
    )

    print(f"GL master rows after company filter: {len(filtered_master_dataframe)}")

    if filtered_master_dataframe.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    company_code = normalize_company_series(filtered_master_dataframe[required_columns["company_code"]])
    company = clean_text_series(filtered_master_dataframe[required_columns["company_name"]])
    company_main_currency = clean_text_series(
        filtered_master_dataframe[required_columns["company_currency"]]
    ).str.upper()
    account_code = normalize_code_series(filtered_master_dataframe[required_columns["account_code"]])
    creation_date = parse_date_series(filtered_master_dataframe[required_columns["creation_date"]])

    result_source, analysis_date, threshold_date = filter_recent_accounts_with_fallback(
        filtered_master_dataframe=filtered_master_dataframe,
        creation_date=creation_date,
        context=context,
    )
    creation_date = creation_date.loc[result_source.index]
    company_code = company_code.loc[result_source.index]
    company = company.loc[result_source.index]
    company_main_currency = company_main_currency.loc[result_source.index]
    account_code = account_code.loc[result_source.index]

    print(f"GL05 accounts created in last three months kept: {len(result_source)}")

    if result_source.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    account_name = choose_account_name(result_source, optional_columns)
    details = choose_details(result_source, optional_columns, account_name)
    active = derive_active_series(result_source, optional_columns)
    inactive = derive_inactive_series(result_source, optional_columns, active)

    control_account = get_optional_text_series(result_source, optional_columns["control_account"])
    balance_account_indicator = get_optional_text_series(
        result_source,
        optional_columns["balance_account_indicator"],
    )
    control_account = first_non_blank_text(control_account, balance_account_indicator).str.upper()
    control_account = control_account.replace(
        {"X": "Y", "TRUE": "Y", "YES": "Y", "SIM": "Y", "S": "Y", "1": "Y"}
    )
    control_account = control_account.mask(control_account == "", "N")

    account_currency = get_optional_text_series(result_source, optional_columns["account_currency"]).str.upper()
    account_currency = account_currency.mask(account_currency == "", "##")
    company_system_currency = get_optional_text_series(
        result_source,
        optional_columns["system_currency"],
    ).str.upper()
    company_system_currency = company_system_currency.mask(company_system_currency == "", "USD")

    balance = get_optional_number_series(result_source, optional_columns["balance"])
    balance_system_currency = get_optional_number_series(
        result_source,
        optional_columns["balance_system_currency"],
    )

    output_dataframe = pd.DataFrame(
        {
            "CoCo": company_code,
            "Company": company,
            "Account Code": account_code,
            "Account Name": account_name,
            "Dimension 1 Relevant": get_optional_text_series(
                result_source,
                optional_columns["dimension_1_relevant"],
            ).mask(
                get_optional_text_series(result_source, optional_columns["dimension_1_relevant"]) == "",
                "N",
            ),
            "Loading Factor Code": get_optional_text_series(
                result_source,
                optional_columns["loading_factor_code"],
            ),
            "Creation Date": creation_date,
            "Date of Update": get_optional_date_series(result_source, optional_columns["update_date"]),
            "Details": details,
            "Control Account": control_account,
            "Active": active,
            "Active From": get_optional_date_series(result_source, optional_columns["active_from"]),
            "Active To": get_optional_date_series(result_source, optional_columns["active_to"]),
            "Inactive": inactive,
            "Inactive From": get_optional_date_series(result_source, optional_columns["inactive_from"]),
            "Inactive To": get_optional_date_series(result_source, optional_columns["inactive_to"]),
            "User Signature": get_optional_text_series(
                result_source,
                optional_columns["user_signature"],
            ),
            "User Code": get_optional_text_series(result_source, optional_columns["created_by"]),
            "User Name": get_optional_text_series(result_source, optional_columns["user_name"]),
            "Account Currency": account_currency,
            "Company Main Currency": company_main_currency,
            "Company System Currency": company_system_currency,
            "Balance": balance,
            "Balance System Currency": balance_system_currency,
            "Balance USD": pd.NA,
            "USD Rate": pd.NA,
            "USD Rate Date": pd.NaT,
            "USD Method": "",
        }
    )

    output_dataframe = output_dataframe.drop_duplicates(
        subset=["CoCo", "Account Code"],
        keep="first",
    )

    output_dataframe = output_dataframe.sort_values(
        by=["CoCo", "Creation Date", "Account Code"],
        kind="stable",
    ).reset_index(drop=True)

    return output_dataframe[OUTPUT_COLUMNS].copy()


def add_usd_fields(output_dataframe, fx_dataframe, context):
    result = output_dataframe.copy()

    if result.empty:
        return result

    if fx_dataframe.empty:
        print("FxRates input was not found. GL05 USD rate columns will remain blank.")
        return result

    normalized_fx_dataframe = normalize_fx_rates(fx_dataframe)
    analysis_date = get_analysis_date(context)
    fx_cache = {}

    balance_usd_values = []
    usd_method_values = []
    usd_rate_values = []
    usd_rate_date_values = []

    for _, row in result.iterrows():
        company_main_currency = normalize_text(row.get("Company Main Currency", "")).upper()
        company_system_currency = normalize_text(row.get("Company System Currency", "")).upper()

        balance_system_currency = pd.to_numeric(
            pd.Series([row.get("Balance System Currency", pd.NA)]),
            errors="coerce",
        ).iloc[0]
        balance = pd.to_numeric(
            pd.Series([row.get("Balance", pd.NA)]),
            errors="coerce",
        ).iloc[0]

        if company_main_currency == "":
            balance_usd_values.append(pd.NA)
            usd_method_values.append("")
            usd_rate_values.append(pd.NA)
            usd_rate_date_values.append(pd.NaT)
            continue

        if company_main_currency == "USD":
            if not pd.isna(balance):
                balance_usd_values.append(balance)
            elif company_system_currency == "USD" and not pd.isna(balance_system_currency):
                balance_usd_values.append(balance_system_currency)
            else:
                balance_usd_values.append(pd.NA)

            usd_method_values.append("SysTotal (sistema=USD)")
            usd_rate_values.append(1)
            usd_rate_date_values.append(analysis_date)
            continue

        cache_key = (company_main_currency, analysis_date.strftime("%Y-%m-%d"))

        if cache_key in fx_cache:
            fx_details = fx_cache[cache_key]
        else:
            fx_details = select_fx_rate_to_usd(
                normalized_fx_dataframe=normalized_fx_dataframe,
                currency=company_main_currency,
                requested_date=analysis_date,
            )
            fx_cache[cache_key] = fx_details

        if fx_details is None:
            balance_usd_values.append(pd.NA)
            usd_method_values.append("")
            usd_rate_values.append(pd.NA)
            usd_rate_date_values.append(pd.NaT)
            continue

        if not pd.isna(balance):
            balance_usd_values.append(balance * fx_details["fx_to_usd"])
        elif company_system_currency == "USD" and not pd.isna(balance_system_currency):
            balance_usd_values.append(balance_system_currency)
        else:
            balance_usd_values.append(pd.NA)

        usd_method_values.append(fx_details["method"])
        usd_rate_values.append(fx_details["usd_rate"])
        usd_rate_date_values.append(fx_details["rate_date"])

    result["Company System Currency"] = "USD"
    result["Balance USD"] = balance_usd_values
    result["USD Method"] = usd_method_values
    result["USD Rate"] = usd_rate_values
    result["USD Rate Date"] = usd_rate_date_values

    missing_rate = result["USD Rate"].isna().sum()

    if missing_rate > 0:
        print()
        print("WARNING: Some GL05 USD rates could not be found.")
        print("USD rate columns are blank for those rows.")
        print(f"Rows without USD rate: {missing_rate}")
        print()

    return result[OUTPUT_COLUMNS].copy()


def write_gl05_output(output_dataframe, context):
    output_file = get_gl_output_file(context)

    print(f"Output workbook: {output_file}")

    if not output_file.exists():
        print("Output workbook does not exist. Using fast GL05 writer...")

        fast_written = write_single_sheet_workbook_fast(
            output_file=output_file,
            sheet_name=SHEET_NAME,
            dataframe=output_dataframe,
            date_columns=DATE_COLUMNS,
            amount_columns=AMOUNT_COLUMNS,
            integer_columns=set(),
        )

        if fast_written:
            print(f"GL05 rows written: {len(output_dataframe)}")
            print(f"GL output workbook: {output_file}")
            print()
            return output_file

    print("Output workbook already exists or fast writer is unavailable.")
    print("Using preserve-sheets openpyxl writer...")

    workbook = open_or_create_gl_output_workbook(output_file)
    worksheet = recreate_gl_sheet(workbook, SHEET_NAME)

    print("Writing GL05 rows to worksheet...")
    write_dataframe_to_sheet(
        worksheet=worksheet,
        dataframe=output_dataframe,
    )

    print("Applying GL05 formatting...")
    apply_standard_gl_formatting(
        worksheet=worksheet,
        dataframe=output_dataframe,
        date_columns=DATE_COLUMNS,
        amount_columns=AMOUNT_COLUMNS,
        integer_columns=set(),
    )

    print("Saving GL output workbook...")
    save_gl_output_workbook(workbook, output_file)

    print(f"GL05 rows written: {len(output_dataframe)}")
    print(f"GL output workbook: {output_file}")
    print()

    return output_file


def run_gl_005(context):
    """
    Execute GL_005 and write only the GL05 sheet.
    """
    print_header("Running GL_005 - General Ledger Accounts Created In The Last Three Months")
    print("Logic: Creation Date = Data / ERDAT within the three months before analysis date")

    master_dataframe = load_gl_master_data(context)
    fx_dataframe = load_gl_fx_rates_data(context)

    if master_dataframe.empty:
        raise FileNotFoundError(
            "GL05 requires the GL master data input file:\n"
            "- input/LBR GL_MD_YYYYMMDD.xlsx"
        )

    output_dataframe = build_gl05_accounts(
        master_dataframe=master_dataframe,
        context=context,
    )

    if not output_dataframe.empty:
        output_dataframe = add_usd_fields(
            output_dataframe=output_dataframe,
            fx_dataframe=fx_dataframe,
            context=context,
        )
    elif fx_dataframe.empty:
        print("FxRates input was not found. GL05 USD columns will remain blank.")

    write_gl05_output(
        output_dataframe=output_dataframe,
        context=context,
    )

    print(f"GL05 rows written: {len(output_dataframe)}")

    return output_dataframe
