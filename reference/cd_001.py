"""
CD_001 - Cash Disbursements By Vendor.

Analysis:
- Module: Cash Disbursements
- Analysis Code: CD_ANALYTIC_01_CDCS101
- Analysis Title: Cash Disbursements By Vendor

Description:
Extracts cash disbursements by vendor using the common CD SAP extract.

Output style:
This control produces a regional-style cash disbursements output with columns similar to:
- CoCo
- Company
- KEY_DOC
- Payment DocEntry
- Payment DocNum
- Journal TransId
- Vendor Code
- Vendor Name
- KEY_VENDOR
- Payment Date
- Tax Date
- Payment Currency
- Company Main Currency
- Payment Amount
- Payment Amount USD
- Counter Reference
- Canceled
- Transfer Amount
- Cash Amount
- Check Amount
- Credit Card Amount
- FxRate

SAP logic:
- The input extract includes BKPF + BSAK + T001 + LFA1 data.
- For this payment-style output, only payment rows are kept:
    D/C = S
- Payment Amount is shown as positive using abs(amount).
- Payment Amount USD is calculated using SAP FX rates when available.
- FxRate is included as the last output column for auditability.

Payment method mapping for Brazil:
- D = Cheque -> Check Amount
- P = Cheque Administrativo -> Check Amount
- T = Cartão de crédito -> Credit Card Amount
- All other payment methods -> Transfer Amount
- No cash/dinheiro method has been identified yet, so Cash Amount remains 0.
"""

import pandas as pd

from core.ar_common import (
    find_fx_rates_file,
    read_sap_fx_rates_file,
    select_fx_rate_row,
)
from core.cd_common import (
    apply_standard_cd_formatting,
    build_key,
    filter_by_company,
    get_cd_output_file,
    get_optional_column,
    load_cd_base_data,
    normalize_company_output,
    normalize_text,
    open_or_create_cd_output_workbook,
    parse_number,
    recreate_cd_sheet,
    require_columns,
    save_cd_output_workbook,
    to_datetime_value,
    write_dataframe_to_sheet,
)


SHEET_NAME = "CD01"


REQUIRED_COLUMNS = {
    "company_code": [
        "Empr",
    ],
    "header_accounting_document": [
        "Nº doc.",
        "Nº doc",
    ],
    "vendor_code": [
        "Fornecedor",
    ],
    "clearing_date": [
        "Compensaç.",
        "Compensaç",
        "Compensac.",
        "Compensac",
    ],
    "clearing_document": [
        "DocCompens",
    ],
    "debit_credit_indicator": [
        "D/C",
    ],
    "amount_document_original": [
        "Montante",
    ],
}


OPTIONAL_COLUMNS = {
    "company_name": [
        "Nome da firma",
    ],
    "vendor_name": [
        "Nome 1",
    ],
    "line_document_date": [
        "Data doc..1",
        "Data doc.1",
        "Data doc1",
    ],
    "header_document_date": [
        "Data doc.",
        "Data doc",
    ],
    "line_currency": [
        "Moeda",
    ],
    "company_main_currency": [
        "Moeda.4",
        "Moeda.3",
        "Moeda.2",
        "Moeda.1",
    ],
    "line_reference": [
        "Referência.1",
        "Referencia.1",
    ],
    "header_reference": [
        "Referência",
        "Referencia",
    ],
    "reference_key": [
        "Chv.ref.",
        "Chv.ref",
    ],
    "transaction_code": [
        "CódT",
        "CodT",
    ],
    "payment_method": [
        "MP",
    ],
    "line_accounting_document": [
        "Nº doc..1",
        "Nº doc.1",
    ],
}


OUTPUT_COLUMNS = [
    "CoCo",
    "Company",
    "KEY_DOC",
    "Payment DocEntry",
    "Payment DocNum",
    "Journal TransId",
    "Vendor Code",
    "Vendor Name",
    "KEY_VENDOR",
    "Payment Date",
    "Tax Date",
    "Payment Currency",
    "Company Main Currency",
    "Payment Amount",
    "Payment Amount USD",
    "Counter Reference",
    "Canceled",
    "Transfer Amount",
    "Cash Amount",
    "Check Amount",
    "Credit Card Amount",
    "FxRate",
]


DATE_COLUMNS = {
    "Payment Date",
    "Tax Date",
}


AMOUNT_COLUMNS = {
    "Payment Amount",
    "Payment Amount USD",
    "Transfer Amount",
    "Cash Amount",
    "Check Amount",
    "Credit Card Amount",
}


INTEGER_COLUMNS = {
    "Payment DocEntry",
    "Payment DocNum",
    "Journal TransId",
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
        return pd.Series(
            [pd.NaT] * len(source_dataframe),
            index=source_dataframe.index,
        )

    return source_dataframe[column_name].apply(to_datetime_value)


def normalize_document_number(value):
    """
    Normalize document-like values.

    Examples:
    - 2000006632.0 -> 2000006632
    - 10903.0 -> 10903
    - blank -> blank
    """
    value_text = normalize_text(value)

    if value_text == "":
        return ""

    if value_text.endswith(".0"):
        value_text = value_text[:-2]

    try:
        number = float(value_text)

        if number.is_integer():
            return str(int(number))
    except ValueError:
        pass

    return value_text


def get_document_series(source_dataframe, column_name):
    """
    Return normalized document series from column or blank series if missing.
    """
    if column_name is None:
        return blank_series(source_dataframe)

    return source_dataframe[column_name].apply(normalize_document_number)


def get_payment_amount_series(source_dataframe, amount_column):
    """
    Return payment amount as positive numeric value.
    """
    def parse_payment_amount(value):
        parsed_amount = parse_number(value)

        if pd.isna(parsed_amount):
            return pd.NA

        return abs(float(parsed_amount))

    return source_dataframe[amount_column].apply(parse_payment_amount)


def choose_counter_reference(row):
    """
    Return the best available counter reference.

    Priority:
    1. Line Reference
    2. Header Reference
    3. Reference Key
    """
    for column_name in [
        "Line Reference",
        "Header Reference",
        "Reference Key",
    ]:
        value = normalize_text(row.get(column_name, ""))

        if value != "":
            return value

    return ""


def derive_journal_trans_id(reference_key, line_accounting_document):
    """
    Derive Journal TransId for the regional-style output.

    SAP ECC does not have a direct SAP B1 Journal TransId equivalent in this extract.

    Priority:
    1. Reference Key, because it is a more technical transaction reference.
    2. Line Accounting Document.
    """
    reference_key = normalize_document_number(reference_key)

    if reference_key != "":
        return reference_key

    return normalize_document_number(line_accounting_document)


def load_fx_rate_cache(context):
    """
    Load SAP FX rates file for CD conversion.

    Expected file:
        input/FxRates_YYYYMMDD.xlsx

    If the file does not exist, return None and the caller will leave USD blank.
    """
    input_folder = context["input_folder"]
    module_config = context["module"]

    fx_rates_file = find_fx_rates_file(
        input_folder=input_folder,
        module_config=module_config,
    )

    if fx_rates_file is None:
        print()
        print("WARNING: FxRates file was not found.")
        print("Payment Amount USD will be blank.")
        print("Expected file pattern: FxRates_YYYYMMDD.xlsx")
        print()
        return None

    fx_df, raw_fx_df = read_sap_fx_rates_file(fx_rates_file)

    print(f"CD_001 FX rates file: {fx_rates_file}")
    print(f"CD_001 FX raw rows: {len(raw_fx_df)}")
    print(f"CD_001 FX normalized rows: {len(fx_df)}")

    return {
        "fx_rates_file": fx_rates_file,
        "fx_df": fx_df,
        "rate_cache": {},
        "missing_rates": set(),
    }


def get_fx_rate_for_payment(fx_context, currency, payment_date):
    """
    Return currency -> USD FX rate for a payment date.

    Rules:
    - USD -> USD = 1
    - Otherwise use SAP FX rates.
    - requested_rate_type='daily' uses TCot M according to AR FX logic.
    """
    currency = normalize_text(currency).upper()

    if currency == "":
        return None

    if currency == "USD":
        return 1.0

    if fx_context is None:
        return None

    payment_date = to_datetime_value(payment_date)

    if pd.isna(payment_date):
        return None

    payment_date_key = payment_date.strftime("%Y-%m-%d")
    cache_key = (currency, payment_date_key)

    if cache_key in fx_context["rate_cache"]:
        return fx_context["rate_cache"][cache_key]

    result = select_fx_rate_row(
        fx_df=fx_context["fx_df"],
        currency=currency,
        requested_date=payment_date,
        requested_rate_type="daily",
        allow_previous_date=True,
        max_previous_days=10,
        allow_tcot_fallback=True,
        allow_future_date=False,
    )

    if result["status"] != "Found":
        fx_context["missing_rates"].add(cache_key)
        fx_context["rate_cache"][cache_key] = None
        return None

    final_fx_rate = result["final_fx_rate"]

    fx_context["rate_cache"][cache_key] = final_fx_rate

    return final_fx_rate


def calculate_fx_rate(row, fx_context):
    """
    Calculate FX rate for output.

    Formula:
        Payment Amount USD = Payment Amount * FxRate
    """
    payment_currency = row.get("Payment Currency", "")
    payment_date = row.get("Payment Date", pd.NaT)

    fx_rate = get_fx_rate_for_payment(
        fx_context=fx_context,
        currency=payment_currency,
        payment_date=payment_date,
    )

    if fx_rate is None:
        return pd.NA

    return fx_rate


def calculate_payment_amount_usd(row):
    """
    Calculate Payment Amount USD.

    Formula:
        Payment Amount USD = Payment Amount * FxRate
    """
    payment_amount = row.get("Payment Amount", pd.NA)
    fx_rate = row.get("FxRate", pd.NA)

    if pd.isna(payment_amount):
        return pd.NA

    if pd.isna(fx_rate):
        return pd.NA

    return float(payment_amount) * float(fx_rate)


def get_payment_method(row):
    """
    Return normalized payment method from the row.
    """
    return normalize_text(row.get("Payment Method", "")).upper()


def derive_transfer_amount(row):
    """
    Derive Transfer Amount based on payment method.

    Brazil payment method mapping from T042Z:
    - D = Cheque -> Check Amount
    - P = Cheque Administrativo -> Check Amount
    - T = Cartão de crédito -> Credit Card Amount
    - Other methods are treated as bank/electronic/transfer/boleto payments.
    """
    payment_method = get_payment_method(row)

    if payment_method in ["D", "P", "T"]:
        return 0

    return row["Payment Amount"]


def derive_cash_amount(row):
    """
    Derive Cash Amount based on payment method.

    No Brazil payment method has been identified as cash/dinheiro yet.
    """
    return 0


def derive_check_amount(row):
    """
    Derive Check Amount based on payment method.

    Brazil payment method mapping from T042Z:
    - D = Cheque
    - P = Cheque Administrativo
    """
    payment_method = get_payment_method(row)

    if payment_method in ["D", "P"]:
        return row["Payment Amount"]

    return 0


def derive_credit_card_amount(row):
    """
    Derive Credit Card Amount based on payment method.

    Brazil payment method mapping from T042Z:
    - T = Cartão de crédito
    """
    payment_method = get_payment_method(row)

    if payment_method == "T":
        return row["Payment Amount"]

    return 0


def build_cd_001_dataframe(source_dataframe, context):
    """
    Build CD_001 regional-style output dataframe from the CD base input.
    """
    required_columns = require_columns(source_dataframe, REQUIRED_COLUMNS)
    optional_columns = resolve_optional_columns(source_dataframe)

    module_config = context["module"]

    filtered_dataframe = filter_by_company(
        dataframe=source_dataframe,
        company_column=required_columns["company_code"],
        companies_filter=module_config.get("companies", ""),
    )

    print(f"CD_001 source rows loaded: {len(source_dataframe)}")
    print(f"CD_001 rows after company filter: {len(filtered_dataframe)}")

    # Keep payment rows only.
    filtered_dataframe = filtered_dataframe[
        filtered_dataframe[required_columns["debit_credit_indicator"]]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        == "S"
    ].copy()

    print(f"CD_001 payment rows D/C = S: {len(filtered_dataframe)}")

    output_dataframe = pd.DataFrame(index=filtered_dataframe.index)

    company_code = filtered_dataframe[required_columns["company_code"]].apply(
        normalize_company_output
    )

    company_name = get_text_series(
        filtered_dataframe,
        optional_columns["company_name"],
    )

    vendor_code = get_document_series(
        filtered_dataframe,
        required_columns["vendor_code"],
    )

    vendor_name = get_text_series(
        filtered_dataframe,
        optional_columns["vendor_name"],
    )

    payment_doc_entry = get_document_series(
        filtered_dataframe,
        required_columns["clearing_document"],
    )

    payment_doc_num = get_document_series(
        filtered_dataframe,
        required_columns["header_accounting_document"],
    )

    line_accounting_document = get_document_series(
        filtered_dataframe,
        optional_columns["line_accounting_document"],
    )

    reference_key = get_document_series(
        filtered_dataframe,
        optional_columns["reference_key"],
    )

    journal_trans_id = [
        derive_journal_trans_id(reference_value, line_document_value)
        for reference_value, line_document_value in zip(
            reference_key,
            line_accounting_document,
        )
    ]

    payment_date = filtered_dataframe[
        required_columns["clearing_date"]
    ].apply(to_datetime_value)

    tax_date = get_date_series(
        filtered_dataframe,
        optional_columns["line_document_date"],
    )

    # If line tax date is missing, use header document date.
    if tax_date.isna().all():
        tax_date = get_date_series(
            filtered_dataframe,
            optional_columns["header_document_date"],
        )

    payment_currency = get_text_series(
        filtered_dataframe,
        optional_columns["line_currency"],
    )

    company_main_currency = get_text_series(
        filtered_dataframe,
        optional_columns["company_main_currency"],
    )

    # If company main currency is blank, fallback to payment currency.
    company_main_currency = company_main_currency.mask(
        company_main_currency == "",
        payment_currency,
    )

    payment_amount = get_payment_amount_series(
        filtered_dataframe,
        required_columns["amount_document_original"],
    )

    output_dataframe["CoCo"] = company_code
    output_dataframe["Company"] = company_name

    output_dataframe["KEY_DOC"] = [
        build_key(company, payment_doc)
        for company, payment_doc in zip(
            output_dataframe["Company"],
            payment_doc_entry,
        )
    ]

    output_dataframe["Payment DocEntry"] = payment_doc_entry
    output_dataframe["Payment DocNum"] = payment_doc_num
    output_dataframe["Journal TransId"] = journal_trans_id

    output_dataframe["Vendor Code"] = vendor_code
    output_dataframe["Vendor Name"] = vendor_name

    output_dataframe["KEY_VENDOR"] = [
        build_key(company, vendor)
        for company, vendor in zip(
            output_dataframe["Company"],
            output_dataframe["Vendor Code"],
        )
    ]

    output_dataframe["Payment Date"] = payment_date
    output_dataframe["Tax Date"] = tax_date
    output_dataframe["Payment Currency"] = payment_currency
    output_dataframe["Company Main Currency"] = company_main_currency
    output_dataframe["Payment Amount"] = payment_amount

    output_dataframe["Line Reference"] = get_text_series(
        filtered_dataframe,
        optional_columns["line_reference"],
    )
    output_dataframe["Header Reference"] = get_text_series(
        filtered_dataframe,
        optional_columns["header_reference"],
    )
    output_dataframe["Reference Key"] = reference_key

    output_dataframe["Counter Reference"] = output_dataframe.apply(
        choose_counter_reference,
        axis=1,
    )

    # SAP extract does not currently provide a direct canceled flag.
    output_dataframe["Canceled"] = "N"

    output_dataframe["Payment Method"] = get_text_series(
        filtered_dataframe,
        optional_columns["payment_method"],
    )

    output_dataframe["Transfer Amount"] = output_dataframe.apply(
        derive_transfer_amount,
        axis=1,
    )

    output_dataframe["Cash Amount"] = output_dataframe.apply(
        derive_cash_amount,
        axis=1,
    )

    output_dataframe["Check Amount"] = output_dataframe.apply(
        derive_check_amount,
        axis=1,
    )

    output_dataframe["Credit Card Amount"] = output_dataframe.apply(
        derive_credit_card_amount,
        axis=1,
    )

    fx_context = load_fx_rate_cache(context)

    output_dataframe["FxRate"] = output_dataframe.apply(
        lambda row: calculate_fx_rate(row, fx_context),
        axis=1,
    )

    output_dataframe["Payment Amount USD"] = output_dataframe.apply(
        calculate_payment_amount_usd,
        axis=1,
    )

    if fx_context is not None and len(fx_context["missing_rates"]) > 0:
        print()
        print("WARNING: Some FX rates were not found.")
        print("Payment Amount USD is blank for those rows.")
        print("Missing rates:")

        for currency, payment_date_key in sorted(fx_context["missing_rates"]):
            print(f"  Currency: {currency} | Payment date: {payment_date_key}")

        print()

    output_dataframe = output_dataframe[OUTPUT_COLUMNS]

    return output_dataframe


def print_cd_001_summary(output_dataframe):
    """
    Print CD_001 validation summary.
    """
    print("CD_001 validation summary")
    print("-------------------------")
    print(f"Rows written: {len(output_dataframe)}")
    print(f"Unique vendors: {output_dataframe['KEY_VENDOR'].nunique()}")
    print(f"Unique payment documents: {output_dataframe['KEY_DOC'].nunique()}")

    total_payment_amount = output_dataframe["Payment Amount"].sum()
    total_payment_amount_usd = output_dataframe["Payment Amount USD"].sum()
    total_transfer_amount = output_dataframe["Transfer Amount"].sum()
    total_cash_amount = output_dataframe["Cash Amount"].sum()
    total_check_amount = output_dataframe["Check Amount"].sum()
    total_credit_card_amount = output_dataframe["Credit Card Amount"].sum()

    print(f"Total Payment Amount: {total_payment_amount:,.2f}")
    print(f"Total Payment Amount USD: {total_payment_amount_usd:,.2f}")
    print(f"Total Transfer Amount: {total_transfer_amount:,.2f}")
    print(f"Total Cash Amount: {total_cash_amount:,.2f}")
    print(f"Total Check Amount: {total_check_amount:,.2f}")
    print(f"Total Credit Card Amount: {total_credit_card_amount:,.2f}")
    print()


def run_cd_001(context):
    """
    Run CD_001.
    """
    source_dataframe = load_cd_base_data(context)

    output_dataframe = build_cd_001_dataframe(
        source_dataframe=source_dataframe,
        context=context,
    )

    if len(output_dataframe) == 0:
        print()
        print("WARNING: CD_001 generated 0 rows.")
        print("Possible causes:")
        print("- No rows with D/C = S were found.")
        print("- COMPANIES filter does not match Empr values.")
        print("- The input file is empty or the wrong sheet was read.")
        print("- The module PARAM1 matched the wrong input file.")
        print()

    output_file = get_cd_output_file(context)

    workbook = open_or_create_cd_output_workbook(output_file)

    worksheet = recreate_cd_sheet(
        workbook=workbook,
        sheet_name=SHEET_NAME,
    )

    write_dataframe_to_sheet(
        worksheet=worksheet,
        dataframe=output_dataframe,
    )

    apply_standard_cd_formatting(
        worksheet=worksheet,
        dataframe=output_dataframe,
        date_columns=DATE_COLUMNS,
        amount_columns=AMOUNT_COLUMNS,
        integer_columns=INTEGER_COLUMNS,
    )

    save_cd_output_workbook(
        workbook=workbook,
        output_file=output_file,
    )

    print_cd_001_summary(output_dataframe)

    print(f"CD_001 output file: {output_file}")
    print(f"CD_001 output sheet: {SHEET_NAME}")
