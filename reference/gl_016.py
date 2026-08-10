"""GL_016 - General journals posted to a selected GL account.

The control is intentionally independent from the other GL controls. It reads
the selected account from the GL_016 control row (PARAM1), with the GL module's
PARAM1 as fallback, filters BSIS and BSAS before enrichment/FX, and writes or
replaces only the GL16 worksheet.
"""

import shutil
import tempfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from time import perf_counter
from zipfile import ZIP_DEFLATED, ZipFile

import pandas as pd
from openpyxl.writer.excel import ExcelWriter

from core.gl_common import (
    apply_standard_gl_formatting,
    build_company_name_map,
    filter_by_company,
    get_gl_output_file,
    get_optional_column,
    load_gl_bsas_data,
    load_gl_bsis_data,
    load_gl_fx_rates_data,
    load_gl_master_data,
    normalize_code_keep_leading_zeroes,
    normalize_company_output,
    normalize_fx_rates,
    normalize_text,
    open_or_create_gl_output_workbook,
    recreate_gl_sheet,
    require_columns,
    select_fx_rate_to_usd,
    to_datetime_value,
    write_dataframe_to_sheet,
    write_single_sheet_workbook_fast,
)


SHEET_NAME = "GL16"
REPORT_CURRENCY = "USD"

OUTPUT_COLUMNS = [
    "Company Code",
    "Company Name",
    "GL Account",
    "GL Account Description",
    "Document Type",
    "Document Number",
    "Document Text",
    "Line Number",
    "Report Amount",
    "Debit",
    "Credit",
    "Report Currency",
    "Amount in Document Currency",
    "Document Date",
    "Transaction Code",
    "Date Entered",
    "Posting Date",
    "Create User ID",
    "Create User Name",
    "Approver User ID",
    "Fiscal Period",
    "Approver User Name",
    "Transaction Code Description",
    "Fiscal Year",
    "Amount in Reporting Currency",
]

DATE_COLUMNS = ["Document Date", "Date Entered", "Posting Date"]

AMOUNT_COLUMNS = [
    "Report Amount",
    "Debit",
    "Credit",
    "Amount in Document Currency",
    "Amount in Reporting Currency",
]

REQUIRED_COLUMNS = {
    "company": ["Company Code", "BKPF-BUKRS", "BUKRS", "Empr", "Empr.1"],
    "year": ["Fiscal Year", "BKPF-GJAHR", "GJAHR", "Ano"],
    "document": [
        "Document Number",
        "BKPF-BELNR",
        "BELNR",
        "Nº doc.",
        "Nº doc..1",
        "Nº doc",
    ],
    "line": ["Line Number", "BSIS-BUZEI", "BSAS-BUZEI", "BUZEI", "Itm"],
    "account": [
        "GL Account",
        "BSIS-HKONT",
        "BSAS-HKONT",
        "HKONT",
        "Razão",
        "Razao",
        "Cta.Razão",
        "Cta.Razao",
        "Cta.Razão.1",
        "Cta.Razao.1",
    ],
    "indicator": [
        "Debit/Credit Indicator",
        "BSIS-SHKZG",
        "BSAS-SHKZG",
        "SHKZG",
        "D/C",
    ],
    "document_amount": [
        "Amount in Document Currency",
        "BSIS-WRBTR",
        "BSAS-WRBTR",
        "WRBTR",
        "Montante",
    ],
    "document_currency": [
        "Document Currency",
        "BKPF-WAERS",
        "WAERS",
        "Moeda",
        "Moeda.1",
    ],
    "posting_date": [
        "Posting Date",
        "BKPF-BUDAT",
        "BUDAT",
        "Dt.lçto.",
        "Dt.lçto",
        "Dt.lcto.",
        "Dt.lcto",
    ],
}

OPTIONAL_COLUMNS = {
    "document_type": ["Document Type", "BKPF-BLART", "BLART", "Tipo doc."],
    "document_text": [
        "Document Text",
        "BKPF-BKTXT",
        "BKTXT",
        "Txt.cabeç.doc.",
        "Texto cab.documento",
    ],
    "document_date": ["Document Date", "BKPF-BLDAT", "BLDAT", "Data doc."],
    "entry_date": [
        "Entry Date",
        "Date Entered",
        "BKPF-CPUDT",
        "CPUDT",
        "Dt.entr.",
        "Dt.entr",
    ],
    "create_user": [
        "Create User ID",
        "BKPF-PPNAM",
        "PPNAM",
        "Usuário responsável",
    ],
    "approver_user": [
        "Approver User ID",
        "BKPF-USNAM",
        "USNAM",
        "Nome do usuário",
    ],
    "transaction": ["Transaction Code", "BKPF-TCODE", "TCODE", "Transação"],
    "fiscal_period": [
        "Fiscal Period",
        "BKPF-MONAT",
        "MONAT",
        "Período",
        "Periodo",
    ],
}

MASTER_ACCOUNT_ALIASES = [
    "GL Account",
    "SAKNR",
    "HKONT",
    "Account",
    "Conta",
    "Razão",
    "Razao",
    "Cta.Razão",
    "Cta.Razao",
    "Cta.Razão.1",
    "Cta.Razao.1",
]

MASTER_COMPANY_ALIASES = ["Company Code", "BUKRS", "Empr", "Empr.1"]

MASTER_DESCRIPTION_ALIASES = [
    "GL Account Description",
    "TXT50",
    "TxtDescr",
    "TXT20",
    "Account Name",
    "Description",
    "Texto breve",
    "Texto",
]


def _log_stage(name, started):
    print(f"GL16 timing - {name}: {perf_counter() - started:.2f} seconds")
    return perf_counter()


def _blank(index):
    return pd.Series("", index=index, dtype="object")


def _clean(series):
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .replace({"nan": "", "None": "", "<NA>": ""})
    )


def _codes(series):
    return _clean(series).map(normalize_code_keep_leading_zeroes)


def _numbers(series):
    text = _clean(series)
    parentheses = text.str.match(r"^\(.*\)$")
    text = text.str.replace(r"[()]", "", regex=True).str.replace(
        " ", "", regex=False
    )
    both = text.str.contains(",", regex=False) & text.str.contains(
        ".", regex=False
    )
    comma_decimal = text.str.contains(r",\d{1,6}$", regex=True)

    text = text.where(
        ~both,
        text.str.replace(".", "", regex=False).str.replace(
            ",", ".", regex=False
        ),
    )
    text = text.where(
        both | ~comma_decimal,
        text.str.replace(".", "", regex=False).str.replace(
            ",", ".", regex=False
        ),
    )
    text = text.where(
        both | comma_decimal,
        text.str.replace(",", "", regex=False),
    )

    result = pd.to_numeric(text, errors="coerce")
    return result.where(~parentheses, -result.abs())


def _dates(series):
    parsed = pd.to_datetime(series, errors="coerce", dayfirst=True)
    numeric = pd.to_numeric(series, errors="coerce")
    excel = pd.to_datetime(
        numeric,
        unit="D",
        origin="1899-12-30",
        errors="coerce",
    )
    return parsed.where(parsed.notna(), excel)


def _optional(dataframe, aliases, transform=None):
    column = get_optional_column(dataframe, aliases)

    if column is None:
        return _blank(dataframe.index)

    series = dataframe[column]

    if transform is None:
        return _clean(series)

    return transform(series)


def _selected_account(context):
    control_value = context.get("control", {}).get("param1", "")
    module_value = context.get("module", {}).get("param1", "")

    if normalize_text(control_value) != "":
        original = control_value
        source = "GL sheet / GL_016 row / PARAM1"
    else:
        original = module_value
        source = "CONFIG sheet / GL module row / PARAM1"

    normalized = normalize_code_keep_leading_zeroes(original)

    print(f"GL16 PARAM1 source: {source}")
    print(f"GL16 PARAM1 original: {normalize_text(original)}")
    print(f"GL16 PARAM1 normalized: {normalized}")

    return original, normalized


def _period(context):
    period_from = to_datetime_value(context["module"].get("from", ""))
    period_to = to_datetime_value(context["module"].get("to", ""))

    if pd.isna(period_from) or pd.isna(period_to):
        raise ValueError(
            "GL16 requires valid FROM and TO dates in config.xlsx."
        )

    return period_from.normalize(), period_to.normalize()


def _prepare_and_filter(
    dataframe,
    source,
    account,
    context,
    period_from,
    period_to,
):
    if dataframe.empty:
        return pd.DataFrame(), 0, 0

    required = require_columns(
        dataframe=dataframe,
        required_columns=REQUIRED_COLUMNS,
        source_name=f"GL16 {source}",
    )

    prepared = pd.DataFrame(index=dataframe.index)
    prepared["Company Code"] = _codes(
        dataframe[required["company"]]
    ).map(normalize_company_output)
    prepared["Fiscal Year"] = _codes(dataframe[required["year"]])
    prepared["Document Number"] = _codes(
        dataframe[required["document"]]
    )
    prepared["Line Number"] = _codes(dataframe[required["line"]])
    prepared["GL Account"] = _codes(dataframe[required["account"]])
    prepared["Debit/Credit Indicator"] = _clean(
        dataframe[required["indicator"]]
    ).str.upper()
    prepared["Document Amount Raw"] = _numbers(
        dataframe[required["document_amount"]]
    )
    prepared["Document Currency"] = _clean(
        dataframe[required["document_currency"]]
    ).str.upper()
    prepared["Posting Date"] = _dates(
        dataframe[required["posting_date"]]
    )
    prepared["Document Type"] = _optional(
        dataframe,
        OPTIONAL_COLUMNS["document_type"],
    )
    prepared["Document Text"] = _optional(
        dataframe,
        OPTIONAL_COLUMNS["document_text"],
    )
    prepared["Document Date"] = _optional(
        dataframe,
        OPTIONAL_COLUMNS["document_date"],
        _dates,
    )
    prepared["Date Entered"] = _optional(
        dataframe,
        OPTIONAL_COLUMNS["entry_date"],
        _dates,
    )
    prepared["Create User ID"] = _optional(
        dataframe,
        OPTIONAL_COLUMNS["create_user"],
    )
    prepared["Approver User ID"] = _optional(
        dataframe,
        OPTIONAL_COLUMNS["approver_user"],
    )
    prepared["Transaction Code"] = _optional(
        dataframe,
        OPTIONAL_COLUMNS["transaction"],
    )
    prepared["Fiscal Period"] = _optional(
        dataframe,
        OPTIONAL_COLUMNS["fiscal_period"],
        _codes,
    )
    prepared["Source"] = source

    prepared = filter_by_company(
        dataframe=prepared,
        company_column="Company Code",
        companies_filter=context["module"].get("companies", ""),
    )

    date_filter = (
        prepared["Posting Date"].notna()
        & (prepared["Posting Date"].dt.normalize() >= period_from)
        & (prepared["Posting Date"].dt.normalize() <= period_to)
    )

    prepared = prepared.loc[date_filter].copy()
    config_count = len(prepared)
    prepared = prepared.loc[
        prepared["GL Account"].eq(account)
    ].copy()

    return prepared, config_count, len(prepared)


def _deduplicate(dataframe):
    if dataframe.empty:
        return dataframe, 0

    key = [
        "Company Code",
        "Fiscal Year",
        "Document Number",
        "Line Number",
    ]

    dataframe = dataframe.copy()
    dataframe["_source_order"] = (
        dataframe["Source"]
        .map({"BSAS": 0, "BSIS": 1})
        .fillna(2)
    )
    dataframe = dataframe.sort_values(
        key + ["_source_order"],
        kind="mergesort",
    )

    duplicates = int(
        dataframe.duplicated(key, keep="first").sum()
    )

    dataframe = dataframe.drop_duplicates(
        key,
        keep="first",
    ).drop(columns="_source_order")

    return dataframe, duplicates


def _account_description_map(master):
    if master.empty:
        return {}

    account_column = get_optional_column(
        master,
        MASTER_ACCOUNT_ALIASES,
    )
    company_column = get_optional_column(
        master,
        MASTER_COMPANY_ALIASES,
    )
    description_column = get_optional_column(
        master,
        MASTER_DESCRIPTION_ALIASES,
    )

    if account_column is None or description_column is None:
        return {}

    result = {}

    for row in master.itertuples(index=False, name=None):
        values = dict(zip(master.columns, row))
        account = normalize_code_keep_leading_zeroes(
            values.get(account_column, "")
        )
        company = (
            normalize_company_output(
                values.get(company_column, "")
            )
            if company_column
            else ""
        )
        description = normalize_text(
            values.get(description_column, "")
        )
        key = (company, account)

        if account and description and key not in result:
            result[key] = description

    return result


def _enrich(dataframe, master):
    before = len(dataframe)
    company_names = build_company_name_map(master)
    account_descriptions = _account_description_map(master)

    dataframe["Company Name"] = (
        dataframe["Company Code"]
        .map(company_names)
        .fillna("")
    )

    keys = zip(
        dataframe["Company Code"],
        dataframe["GL Account"],
    )
    dataframe["GL Account Description"] = [
        account_descriptions.get(key, "")
        for key in keys
    ]

    dataframe["Create User Name"] = ""
    dataframe["Approver User Name"] = ""
    dataframe["Transaction Code Description"] = ""

    if len(dataframe) != before:
        raise ValueError(
            "GL16 GL master enrichment changed the journal line count."
        )

    return dataframe


def _amounts_and_fx(dataframe, fx_rates):
    indicator = dataframe["Debit/Credit Indicator"]
    absolute = dataframe["Document Amount Raw"].abs()
    credit_mask = indicator.isin(["H", "C", "CREDIT"])
    debit_mask = indicator.isin(["S", "D", "DEBIT"])

    dataframe["Debit"] = absolute.where(debit_mask, 0.0)
    dataframe["Credit"] = absolute.where(credit_mask, 0.0)
    dataframe["Amount in Document Currency"] = absolute.where(
        ~credit_mask,
        -absolute,
    )
    dataframe["Report Currency"] = REPORT_CURRENCY
    dataframe["Amount in Reporting Currency"] = pd.NA

    normalized_fx = normalize_fx_rates(fx_rates)

    conversion_keys = dataframe[
        ["Document Currency", "Posting Date"]
    ].drop_duplicates()

    for currency, posting_date in conversion_keys.itertuples(
        index=False,
        name=None,
    ):
        rate = select_fx_rate_to_usd(
            normalized_fx_dataframe=normalized_fx,
            currency=currency,
            requested_date=posting_date,
        )

        if rate is None:
            continue

        conversion_filter = (
            dataframe["Document Currency"].eq(currency)
            & (
                dataframe["Posting Date"].dt.normalize()
                == posting_date.normalize()
            )
        )

        dataframe.loc[
            conversion_filter,
            "Amount in Reporting Currency",
        ] = (
            dataframe.loc[
                conversion_filter,
                "Amount in Document Currency",
            ]
            * rate["fx_to_usd"]
        )

    dataframe["Report Amount"] = dataframe[
        "Amount in Reporting Currency"
    ]

    return dataframe


def _empty_output():
    return pd.DataFrame(columns=OUTPUT_COLUMNS)


def _save_existing_workbook_fast(workbook, output_file):
    """Save locally with low compression before copying to OneDrive."""
    output_file = Path(output_file)
    output_file.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.NamedTemporaryFile(
        suffix=".xlsx",
        delete=False,
    ) as handle:
        temporary_file = Path(handle.name)

    try:
        with ZipFile(
            temporary_file,
            mode="w",
            compression=ZIP_DEFLATED,
            compresslevel=1,
            allowZip64=True,
        ) as archive:
            ExcelWriter(
                workbook,
                archive,
            ).save()

        shutil.copy2(
            temporary_file,
            output_file,
        )
    except PermissionError as error:
        raise PermissionError(
            f"Could not save output workbook: {output_file}. "
            "Close the workbook and run again."
        ) from error
    finally:
        workbook.close()
        temporary_file.unlink(missing_ok=True)


def _timed_load(loader, context):
    started = perf_counter()
    dataframe = loader(context)

    return (
        dataframe,
        perf_counter() - started,
    )


def _load_journal_sources(context):
    """Read the independent BSIS and BSAS workbooks concurrently."""
    with ThreadPoolExecutor(max_workers=2) as executor:
        bsis_future = executor.submit(
            _timed_load,
            load_gl_bsis_data,
            context,
        )
        bsas_future = executor.submit(
            _timed_load,
            load_gl_bsas_data,
            context,
        )

        bsis, bsis_seconds = bsis_future.result()
        bsas, bsas_seconds = bsas_future.result()

    print(
        "GL16 timing - lectura BSIS: "
        f"{bsis_seconds:.2f} seconds"
    )
    print(
        "GL16 timing - lectura BSAS: "
        f"{bsas_seconds:.2f} seconds"
    )

    return bsis, bsas


def _write(context, dataframe):
    output_file = get_gl_output_file(context)

    created = write_single_sheet_workbook_fast(
        output_file=output_file,
        sheet_name=SHEET_NAME,
        dataframe=dataframe,
        date_columns=DATE_COLUMNS,
        amount_columns=AMOUNT_COLUMNS,
        integer_columns=[],
    )

    if created:
        print(
            f"GL16 output file created: {output_file}"
        )
        return output_file

    workbook = open_or_create_gl_output_workbook(
        output_file
    )
    worksheet = recreate_gl_sheet(
        workbook,
        SHEET_NAME,
    )

    write_dataframe_to_sheet(
        worksheet,
        dataframe,
    )
    apply_standard_gl_formatting(
        worksheet=worksheet,
        dataframe=dataframe,
        date_columns=DATE_COLUMNS,
        amount_columns=AMOUNT_COLUMNS,
        integer_columns=[],
    )

    _save_existing_workbook_fast(
        workbook,
        output_file,
    )

    print(
        "GL16 output sheet written: "
        f"{output_file} [{SHEET_NAME}]"
    )

    return output_file


def run_gl_016(context):
    """Run GL_016."""
    total_start = perf_counter()
    stage_start = total_start

    print(
        "Running GL16 - General Journal Entries Posted "
        "For A Selected General Ledger Account"
    )
    print(
        "GL16 input sufficiency: BSIS, BSAS, GL_MD, "
        "optional FxRates and PARAM1 are sufficient."
    )

    original, account = _selected_account(context)

    if normalize_text(original) == "" or account == "":
        print(
            "GL16 requires a selected GL Account in PARAM1."
        )

        output = _empty_output()
        output_file = _write(context, output)

        print("GL16 rows written: 0")
        print(
            "GL16 total elapsed: "
            f"{perf_counter() - total_start:.2f} seconds"
        )

        return {
            "status": "OK",
            "output_file": output_file,
            "sheet_name": SHEET_NAME,
            "rows": 0,
        }

    period_from, period_to = _period(context)

    bsis, bsas = _load_journal_sources(context)

    print(f"GL16 BSIS rows read: {len(bsis)}")
    print(f"GL16 BSAS rows read: {len(bsas)}")

    stage_start = _log_stage(
        "lectura paralela BSIS/BSAS",
        stage_start,
    )

    if bsis.empty and bsas.empty:
        raise ValueError(
            "GL16 requires at least one BSIS or BSAS input file."
        )

    (
        bsis_filtered,
        bsis_config,
        bsis_matches,
    ) = _prepare_and_filter(
        dataframe=bsis,
        source="BSIS",
        account=account,
        context=context,
        period_from=period_from,
        period_to=period_to,
    )

    (
        bsas_filtered,
        bsas_config,
        bsas_matches,
    ) = _prepare_and_filter(
        dataframe=bsas,
        source="BSAS",
        account=account,
        context=context,
        period_from=period_from,
        period_to=period_to,
    )

    print(
        "GL16 rows after CONFIG filters - "
        f"BSIS: {bsis_config}; BSAS: {bsas_config}"
    )
    print(
        "GL16 account matches - "
        f"BSIS: {bsis_matches}; BSAS: {bsas_matches}"
    )

    stage_start = _log_stage(
        "filtros CONFIG y cuenta",
        stage_start,
    )

    journals = pd.concat(
        [bsas_filtered, bsis_filtered],
        ignore_index=True,
    )

    journals, duplicates = _deduplicate(
        journals
    )

    print(
        "GL16 BSIS/BSAS duplicate lines removed: "
        f"{duplicates}"
    )

    stage_start = _log_stage(
        "deduplicación",
        stage_start,
    )

    if (
        not journals.empty
        and not journals["GL Account"].eq(account).all()
    ):
        raise ValueError(
            "GL16 validation failed: output contains "
            "a different GL account."
        )

    journal_line_key = [
        "Company Code",
        "Fiscal Year",
        "Document Number",
        "Line Number",
    ]

    if journals.duplicated(journal_line_key).any():
        raise ValueError(
            "GL16 validation failed: duplicate "
            "journal line keys remain."
        )

    if journals.empty:
        print(
            "GL16 found no exact GL Account matches "
            "in BSIS or BSAS. Select PARAM1 from the "
            "Razão column of BSIS/BSAS, not only from GL_MD."
        )
        print(
            "GL16 GL Master Data skipped: "
            "no matching journal lines."
        )
        print(
            "GL16 FxRates used: no "
            "(no matching journal lines)."
        )

        output = _empty_output()

        stage_start = _log_stage(
            "construcción del output",
            stage_start,
        )
        output_file = _write(
            context,
            output,
        )
        _log_stage(
            "escritura",
            stage_start,
        )

        print("GL16 distinct documents: 0")
        print("GL16 rows written: 0")
        print(
            "GL16 total elapsed: "
            f"{perf_counter() - total_start:.2f} seconds"
        )

        return {
            "status": "OK",
            "output_file": output_file,
            "sheet_name": SHEET_NAME,
            "rows": 0,
        }

    master = load_gl_master_data(context)
    journals = _enrich(
        journals,
        master,
    )

    stage_start = _log_stage(
        "GL Master Data",
        stage_start,
    )

    fx_rates = load_gl_fx_rates_data(context)
    journals = _amounts_and_fx(
        journals,
        fx_rates,
    )

    print(
        "GL16 FxRates used: "
        f"{'yes' if not fx_rates.empty else 'no'}"
    )

    stage_start = _log_stage(
        "FX",
        stage_start,
    )

    invalid_debit = (
        journals["Debit"].fillna(0) < 0
    ).any()
    invalid_credit = (
        journals["Credit"].fillna(0) < 0
    ).any()

    if invalid_debit or invalid_credit:
        raise ValueError(
            "GL16 validation failed: Debit or Credit "
            "contains a negative value."
        )

    missing_period = journals[
        "Fiscal Period"
    ].eq("")

    if missing_period.any():
        derived_period = (
            journals["Posting Date"]
            .dt.month
            .astype("Int64")
            .astype("string")
        )

        journals["Fiscal Period"] = (
            journals["Fiscal Period"].mask(
                missing_period,
                derived_period,
            )
        )

    output = (
        journals.reindex(
            columns=OUTPUT_COLUMNS
        )
        .sort_values(
            journal_line_key,
            kind="mergesort",
        )
    )

    if list(output.columns) != OUTPUT_COLUMNS:
        raise ValueError(
            "GL16 validation failed: output columns "
            "do not match LHA."
        )

    distinct_documents = journals[
        [
            "Company Code",
            "Fiscal Year",
            "Document Number",
        ]
    ].drop_duplicates().shape[0]

    print(
        "GL16 distinct documents: "
        f"{distinct_documents}"
    )

    stage_start = _log_stage(
        "construcción del output",
        stage_start,
    )

    output_file = _write(
        context,
        output,
    )

    _log_stage(
        "escritura",
        stage_start,
    )

    print(
        f"GL16 rows written: {len(output)}"
    )
    print(
        "GL16 total elapsed: "
        f"{perf_counter() - total_start:.2f} seconds"
    )

    return {
        "status": "OK",
        "output_file": output_file,
        "sheet_name": SHEET_NAME,
        "rows": len(output),
    }
