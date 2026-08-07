"""
GL_010 - Potential Duplicate General Journals: Same GL Account & Description.

The control compares accounting lines inside the active CONFIG posting-date range.
It reproduces the strict LHA description rule: trimmed Line Text, falling back to
trimmed Document Text, with case, accents, punctuation, and internal whitespace
preserved.  A result requires at least two different journals.

The control is intentionally independent.  It filters and detects exceptions before
master-data enrichment and optional FX conversion, and writes/replaces only GL10.
"""

from pathlib import Path
from time import perf_counter

import pandas as pd

from core.gl_common import (
    apply_standard_gl_formatting,
    build_company_name_map,
    build_gl_account_name_map,
    filter_by_company,
    get_gl_output_file,
    load_gl_bsas_data,
    load_gl_bsis_data,
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


SHEET_NAME = "GL10"

OUTPUT_COLUMNS = [
    "Company Code",
    "Company Name",
    "GL Account",
    "GL Account Description",
    "Document Type",
    "Document Number",
    "Document Text",
    "Line Text",
    "Effective Description",
    "Line Number",
    "Report Amount",
    "Debit",
    "Credit",
    "Report Currency",
    "Amount in Document Currency",
    "Document Currency",
    "Amount in Reporting Currency",
    "Document Date",
    "Date Entered",
    "Posting Date",
    "Create User ID",
    "Create User Name",
    "Approver User ID",
    "Approver User Name",
    "Fiscal Period",
    "Transaction Code",
    "Transaction Code Description",
    "Fiscal Year",
    "Duplicate Journal Count",
    "Duplicate Line Count",
    "Duplicate Key",
    "FX Method",
    "FX Rate",
    "FX Rate Date",
    "Source",
]

DATE_COLUMNS = {"Document Date", "Date Entered", "Posting Date", "FX Rate Date"}
AMOUNT_COLUMNS = {
    "Report Amount",
    "Debit",
    "Credit",
    "Amount in Document Currency",
    "Amount in Reporting Currency",
    "FX Rate",
}
INTEGER_COLUMNS = {"Duplicate Journal Count", "Duplicate Line Count"}

ALIASES = {
    "company": ["Empr", "BUKRS", "Company Code"],
    "fiscal_year": ["Ano", "GJAHR", "Fiscal Year"],
    "document": ["Nº doc.", "Nº doc", "BELNR", "Document Number"],
    "line": ["Itm", "BUZEI", "Line Number", "Line"],
    "account": ["Razão", "Razao", "Cta.Razão", "Cta.Razao", "HKONT", "GL Account", "Account Code"],
    "posting_date": ["Dt.lçto.", "Dt.lçto", "Dt.lcto.", "Dt.lcto", "BUDAT", "Posting Date"],
    "line_text": ["Texto", "SGTXT", "Line Text", "Line Memo"],
    "document_text": [
        "Texto cabeçalho documento",
        "Texto cabecalho documento",
        "BKTXT",
        "Document Text",
        "Journal Memo",
    ],
    "document_type": ["Tp.doc.", "Tp.doc", "BLART", "Document Type", "Journal Type"],
    "document_date": ["Data doc.", "Data doc", "BLDAT", "Document Date"],
    "entry_date": ["Dt.entr.", "Dt.entr", "CPUDT", "Date Entered", "Entry Date"],
    "create_user": ["Nome do usuário", "Nome do usuario", "USNAM", "Create User ID", "Creator ID"],
    "approver_user": ["Pré-edição", "Pre-edicao", "PPNAM", "Approver User ID", "Approver ID"],
    "transaction_code": ["CódT", "CodT", "TCODE", "Transaction Code"],
    "currency": ["Moeda", "Moeda.1", "WAERS", "Document Currency", "FC Currency"],
    "dc_indicator": ["D/C", "SHKZG", "Debit Credit Indicator"],
    "local_amount": ["Montante em MI", "DMBTR", "Line Amount Local"],
    "document_amount": ["Montante", "WRBTR", "Amount in Document Currency"],
}

REQUIRED_FIELDS = {
    "company",
    "fiscal_year",
    "document",
    "line",
    "account",
    "posting_date",
    "line_text",
    "document_text",
}


def print_header(title):
    print(title)
    print("-" * len(title))


def _elapsed(start):
    return perf_counter() - start


def _header_base(value):
    """Return a case-insensitive header without a pandas duplicate suffix."""
    text = str(value).strip().casefold()
    parts = text.rsplit(".", 1)
    if len(parts) == 2 and parts[1].isdigit():
        return parts[0]
    return text


def _matching_columns(columns, aliases):
    exact = {str(alias).strip().casefold() for alias in aliases}
    result = []
    for column in columns:
        normalized = str(column).strip().casefold()
        if normalized in exact or _header_base(column) in exact:
            result.append(column)
    return result


def _text(series):
    """Null-safe vectorized text cleanup used for keys and visible text."""
    result = series.where(series.notna(), "").astype("string")
    result = result.str.replace("\u00a0", " ", regex=False).str.strip()
    result = result.mask(result.str.casefold().isin(["nan", "none", "<na>"]), "")
    return result.fillna("").astype("object")


def _code(series):
    result = _text(series)
    return result.str.replace(r"\.0$", "", regex=True)


def _date(series):
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, errors="coerce")

    text = _text(series)
    result = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    compact = text.str.fullmatch(r"\d{8}", na=False)
    iso = text.str.fullmatch(r"\d{4}-\d{2}-\d{2}.*", na=False)

    if compact.any():
        result.loc[compact] = pd.to_datetime(text.loc[compact], format="%Y%m%d", errors="coerce")
    if iso.any():
        result.loc[iso] = pd.to_datetime(text.loc[iso].str[:10], format="%Y-%m-%d", errors="coerce")
    other = (text != "") & ~compact & ~iso
    if other.any():
        result.loc[other] = pd.to_datetime(text.loc[other], errors="coerce", dayfirst=True)
    return result


def _number(series):
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    text = _text(series).str.replace(" ", "", regex=False)
    parenthesized_negative = text.str.startswith("(") & text.str.endswith(")")
    text = text.mask(parenthesized_negative, text.str[1:-1])

    comma = text.str.contains(",", regex=False, na=False)
    dot = text.str.contains(".", regex=False, na=False)
    text = text.where(~(comma & ~dot), text.str.replace(",", ".", regex=False))
    both = comma & dot
    european = both & (text.str.rfind(",") > text.str.rfind("."))
    text = text.where(
        ~european,
        text.str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
    )
    us = both & ~european
    text = text.where(~us, text.str.replace(",", "", regex=False))

    number = pd.to_numeric(text, errors="coerce")
    return number.mask(parenthesized_negative, number * -1)


def _candidate_score(dataframe, column, logical_name, companies_filter=""):
    """Score duplicate Portuguese headers by expected content, never by position alone."""
    sample = dataframe[column].dropna().head(5000)
    if sample.empty:
        return -1.0
    values = _text(sample)
    nonblank = values != ""
    if not nonblank.any():
        return -1.0

    score = float(nonblank.mean())
    if logical_name == "company":
        configured = {
            item.strip().casefold()
            for item in str(companies_filter).replace(";", ",").split(",")
            if item.strip()
        }
        if configured:
            score += 10.0 * values.str.casefold().isin(configured).mean()
        score += 2.0 * values.str.len().between(1, 8).mean()
    elif logical_name == "fiscal_year":
        score += 5.0 * values.str.fullmatch(r"(?:19|20)\d{2}(?:\.0)?", na=False).mean()
    elif logical_name == "line":
        numeric = pd.to_numeric(values, errors="coerce")
        score += 2.0 * numeric.notna().mean() + 2.0 * numeric.between(0, 99999).mean()
    elif logical_name == "account":
        score += 2.0 * values.str.fullmatch(r"[A-Za-z0-9._/-]+", na=False).mean()
    elif logical_name in {"posting_date", "document_date", "entry_date"}:
        score += 5.0 * _date(sample).notna().mean()
    elif logical_name in {"local_amount", "document_amount"}:
        score += 5.0 * _number(sample).notna().mean()
    elif logical_name == "dc_indicator":
        score += 5.0 * values.str.upper().isin(["S", "H", "D", "C"]).mean()
    return score


def resolve_columns(dataframe, source_name, companies_filter=""):
    resolved = {}
    for logical_name, aliases in ALIASES.items():
        candidates = _matching_columns(dataframe.columns, aliases)
        if not candidates:
            if logical_name in REQUIRED_FIELDS:
                raise ValueError(
                    f"GL10 could not resolve required field '{logical_name}' in {source_name}. "
                    f"Expected one of: {', '.join(aliases)}"
                )
            resolved[logical_name] = None
            continue
        if len(candidates) == 1:
            resolved[logical_name] = candidates[0]
            continue
        scores = {
            column: _candidate_score(dataframe, column, logical_name, companies_filter)
            for column in candidates
        }
        selected = max(candidates, key=lambda column: scores[column])
        resolved[logical_name] = selected
        print(f"{source_name} {logical_name}: selected '{selected}' from {candidates}")
    return resolved


def _optional_series(dataframe, column, default=""):
    if column is None:
        return pd.Series(default, index=dataframe.index, dtype="object")
    return dataframe[column]


def _module_dates(context):
    module_config = context["module"]
    date_from = to_datetime_value(module_config.get("from", ""))
    date_to = to_datetime_value(module_config.get("to", ""))
    if pd.isna(date_from) or pd.isna(date_to):
        raise ValueError("GL10 requires valid CONFIG FROM and CONFIG TO dates.")
    return pd.Timestamp(date_from).normalize(), pd.Timestamp(date_to).normalize()


def prepare_source(source_dataframe, context, source_name):
    """Resolve, filter, and reduce one source before description processing."""
    if source_dataframe.empty:
        return pd.DataFrame(), {"read": 0, "historical": 0, "in_period": 0}

    module_config = context["module"]
    columns = resolve_columns(
        source_dataframe,
        source_name,
        module_config.get("companies", ""),
    )
    rows_read = len(source_dataframe)
    filtered = filter_by_company(
        dataframe=source_dataframe,
        company_column=columns["company"],
        companies_filter=module_config.get("companies", ""),
    )
    posting_date = _date(filtered[columns["posting_date"]])
    date_from, date_to = _module_dates(context)
    in_period_mask = posting_date.between(date_from, date_to, inclusive="both")
    historical_rows = int((~in_period_mask).sum())
    filtered = filtered.loc[in_period_mask].copy()
    posting_date = posting_date.loc[filtered.index]

    if filtered.empty:
        return pd.DataFrame(), {
            "read": rows_read,
            "historical": historical_rows,
            "in_period": 0,
        }

    company = _code(filtered[columns["company"]])
    fiscal_year = _code(filtered[columns["fiscal_year"]])
    document = _code(filtered[columns["document"]])
    line = _code(filtered[columns["line"]])
    account = _code(filtered[columns["account"]])

    result = pd.DataFrame(index=filtered.index)
    result["Company Code"] = company
    result["Fiscal Year"] = fiscal_year
    result["Document Number"] = document
    result["Line Number"] = line
    result["GL Account"] = account
    result["Posting Date"] = posting_date
    result["Line Text"] = _text(filtered[columns["line_text"]])
    result["Document Text"] = _text(filtered[columns["document_text"]])
    result["Document Type"] = _text(_optional_series(filtered, columns["document_type"]))
    result["Document Date"] = _date(_optional_series(filtered, columns["document_date"]))
    result["Date Entered"] = _date(_optional_series(filtered, columns["entry_date"]))
    result["Create User ID"] = _text(_optional_series(filtered, columns["create_user"]))
    result["Approver User ID"] = _text(_optional_series(filtered, columns["approver_user"]))
    result["Transaction Code"] = _text(_optional_series(filtered, columns["transaction_code"]))
    result["Document Currency"] = _text(_optional_series(filtered, columns["currency"])).str.upper()
    result["D/C"] = _text(_optional_series(filtered, columns["dc_indicator"])).str.upper()
    result["Local Amount Raw"] = _number(_optional_series(filtered, columns["local_amount"]))
    result["Document Amount Raw"] = _number(_optional_series(filtered, columns["document_amount"]))
    result["Source"] = source_name

    result["_JOURNAL_KEY"] = (
        company + "|" + fiscal_year + "|" + document
    )
    result["_LINE_KEY"] = result["_JOURNAL_KEY"] + "|" + line

    return result.reset_index(drop=True), {
        "read": rows_read,
        "historical": historical_rows,
        "in_period": len(result),
    }


def deduplicate_technical_lines(dataframe):
    """Keep one accounting line; BSAS prevails over BSIS as the later cleared state."""
    if dataframe.empty:
        return dataframe, 0
    result = dataframe.copy()
    result["_SOURCE_PRIORITY"] = result["Source"].str.upper().map({"BSIS": 1, "BSAS": 2}).fillna(0)
    result = result.sort_values(["_LINE_KEY", "_SOURCE_PRIORITY"], kind="stable")
    duplicate_mask = result.duplicated("_LINE_KEY", keep="last")
    removed = int(duplicate_mask.sum())
    result = result.loc[~duplicate_mask].drop(columns="_SOURCE_PRIORITY").reset_index(drop=True)
    return result, removed


def detect_duplicate_lines(dataframe):
    """Apply strict LHA description selection and return qualifying line-level rows."""
    metrics = {
        "missing_company": 0,
        "missing_account": 0,
        "missing_document_key": 0,
        "line_text": 0,
        "document_fallback": 0,
        "blank_description": 0,
        "duplicate_keys": 0,
    }
    if dataframe.empty:
        return dataframe, metrics

    result = dataframe.copy()
    missing_company = result["Company Code"] == ""
    missing_account = result["GL Account"] == ""
    missing_document = (
        (result["Fiscal Year"] == "")
        | (result["Document Number"] == "")
        | (result["Line Number"] == "")
    )
    metrics["missing_company"] = int(missing_company.sum())
    metrics["missing_account"] = int(missing_account.sum())
    metrics["missing_document_key"] = int(missing_document.sum())
    result = result.loc[~(missing_company | missing_account | missing_document)].copy()

    has_line_text = result["Line Text"] != ""
    has_document_text = result["Document Text"] != ""
    metrics["line_text"] = int(has_line_text.sum())
    metrics["document_fallback"] = int((~has_line_text & has_document_text).sum())
    metrics["blank_description"] = int((~has_line_text & ~has_document_text).sum())

    # Strict LHA: no case-folding, accent removal, punctuation changes, or internal
    # whitespace collapse.  _text already performed only null handling and strip.
    result["Effective Description"] = result["Line Text"].where(
        has_line_text,
        result["Document Text"],
    )
    result = result.loc[result["Effective Description"] != ""].copy()
    if result.empty:
        return result, metrics

    group_columns = ["Company Code", "GL Account", "Effective Description"]
    result["Duplicate Journal Count"] = result.groupby(
        group_columns,
        sort=False,
        dropna=False,
    )["_JOURNAL_KEY"].transform("nunique")
    result = result.loc[result["Duplicate Journal Count"] > 1].copy()
    if result.empty:
        return result, metrics

    result["Duplicate Line Count"] = result.groupby(
        group_columns,
        sort=False,
        dropna=False,
    )["_LINE_KEY"].transform("size")
    metrics["duplicate_keys"] = int(result[group_columns].drop_duplicates().shape[0])
    result["Duplicate Key"] = (
        result["Company Code"]
        + " | "
        + result["GL Account"]
        + " | "
        + result["Effective Description"]
    )
    return result, metrics


def _safe_map(master_dataframe, builder):
    if master_dataframe.empty:
        return {}
    try:
        return builder(master_dataframe)
    except (KeyError, ValueError, TypeError):
        return {}


def enrich_output(dataframe, master_dataframe):
    if dataframe.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)
    result = dataframe.copy()
    company_map = _safe_map(master_dataframe, build_company_name_map)
    account_map = _safe_map(master_dataframe, build_gl_account_name_map)
    result["Company Name"] = result["Company Code"].map(company_map).fillna("")
    result["GL Account Description"] = result["GL Account"].map(account_map).fillna("")

    local_abs = result["Local Amount Raw"].abs()
    is_debit = result["D/C"].isin(["S", "D"])
    is_credit = result["D/C"].isin(["H", "C"])
    result["Debit"] = local_abs.where(is_debit, pd.NA)
    result["Credit"] = local_abs.where(is_credit, pd.NA)
    result["_SIGNED_LOCAL"] = local_abs.where(~is_credit, -local_abs)

    document_abs = result["Document Amount Raw"].abs()
    result["Amount in Document Currency"] = document_abs.where(~is_credit, -document_abs)
    result["Create User Name"] = ""
    result["Approver User Name"] = ""
    result["Fiscal Period"] = result["Posting Date"].dt.strftime("%Y-%m").fillna("")
    result["Transaction Code Description"] = ""
    result["Report Currency"] = ""
    result["Amount in Reporting Currency"] = pd.NA
    result["Report Amount"] = pd.NA
    result["FX Method"] = ""
    result["FX Rate"] = pd.NA
    result["FX Rate Date"] = pd.NaT
    return result


def add_optional_fx(dataframe, fx_dataframe):
    """Resolve FX once per currency/date pair and map it only to exception rows."""
    if dataframe.empty or fx_dataframe.empty:
        return dataframe, len(dataframe)

    result = dataframe.copy()
    normalized_fx = normalize_fx_rates(fx_dataframe)
    fx_date = result["Document Date"].where(result["Document Date"].notna(), result["Posting Date"])
    result["_FX_DATE"] = pd.to_datetime(fx_date, errors="coerce").dt.normalize()
    pairs = result[["Document Currency", "_FX_DATE"]].drop_duplicates()
    lookup_rows = []

    # The loop is over distinct FX keys after exception reduction, never over the GL
    # population or individual output lines.
    for pair in pairs.to_dict("records"):
        currency = pair["Document Currency"]
        requested_date = pair["_FX_DATE"]
        details = select_fx_rate_to_usd(
            normalized_fx_dataframe=normalized_fx,
            currency=normalize_text(currency).upper(),
            requested_date=requested_date,
        )
        lookup_rows.append(
            {
                "Document Currency": currency,
                "_FX_DATE": requested_date,
                "_FX_TO_REPORT": pd.NA if details is None else details["fx_to_usd"],
                "FX Method": "" if details is None else details["method"],
                "FX Rate": pd.NA if details is None else details["usd_rate"],
                "FX Rate Date": pd.NaT if details is None else details["rate_date"],
            }
        )

    lookup = pd.DataFrame(lookup_rows)
    result = result.drop(columns=["FX Method", "FX Rate", "FX Rate Date"]).merge(
        lookup,
        on=["Document Currency", "_FX_DATE"],
        how="left",
        validate="many_to_one",
    )
    fx_factor = pd.to_numeric(result["_FX_TO_REPORT"], errors="coerce")
    signed_document_amount = pd.to_numeric(result["Amount in Document Currency"], errors="coerce")
    result["Amount in Reporting Currency"] = signed_document_amount * fx_factor
    result["Report Amount"] = result["Amount in Reporting Currency"]
    converted = result["Amount in Reporting Currency"].notna()
    result["Report Currency"] = ""
    result.loc[converted, "Report Currency"] = "USD"
    missing = int((~converted).sum())
    return result.drop(columns=["_FX_DATE", "_FX_TO_REPORT"]), missing


def finalize_output(dataframe):
    if dataframe.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)
    result = dataframe.sort_values(
        [
            "Company Code",
            "GL Account",
            "Effective Description",
            "Fiscal Year",
            "Document Number",
            "Line Number",
        ],
        kind="stable",
    ).reset_index(drop=True)
    for column in OUTPUT_COLUMNS:
        if column not in result.columns:
            result[column] = ""
    return result[OUTPUT_COLUMNS].copy()


def create_gl10_duplicate_journals(
    bsis_dataframe,
    bsas_dataframe,
    master_dataframe,
    fx_dataframe,
    context,
):
    stage = perf_counter()
    bsis, bsis_metrics = prepare_source(bsis_dataframe, context, "BSIS")
    bsas, bsas_metrics = prepare_source(bsas_dataframe, context, "BSAS")
    print(f"GL10 filter time: {_elapsed(stage):.2f} seconds")
    print(f"BSIS rows read: {bsis_metrics['read']}")
    print(f"BSAS rows read: {bsas_metrics['read']}")
    print(f"Historical/out-of-config rows excluded: {bsis_metrics['historical'] + bsas_metrics['historical']}")
    print(f"Rows inside CONFIG period: {bsis_metrics['in_period'] + bsas_metrics['in_period']}")

    if bsis.empty and bsas.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS), {"technical_duplicates": 0, "fx_missing": 0}

    stage = perf_counter()
    population = pd.concat([bsis, bsas], ignore_index=True)
    population, technical_duplicates = deduplicate_technical_lines(population)
    exceptions, detection_metrics = detect_duplicate_lines(population)
    print(f"GL10 detection time: {_elapsed(stage):.2f} seconds")
    print(f"Technical BSIS/BSAS duplicates removed: {technical_duplicates}")
    print(f"Rows without Company Code: {detection_metrics['missing_company']}")
    print(f"Rows without GL Account: {detection_metrics['missing_account']}")
    print(f"Rows without document/line key: {detection_metrics['missing_document_key']}")
    print(f"Rows using Line Text: {detection_metrics['line_text']}")
    print(f"Rows using Document Text fallback: {detection_metrics['document_fallback']}")
    print(f"Rows with both descriptions blank: {detection_metrics['blank_description']}")
    print(f"Duplicate keys detected: {detection_metrics['duplicate_keys']}")

    stage = perf_counter()
    output = enrich_output(exceptions, master_dataframe)
    output, fx_missing = add_optional_fx(output, fx_dataframe)
    output = finalize_output(output)
    print(f"GL10 enrichment and FX time: {_elapsed(stage):.2f} seconds")

    metrics = {
        "technical_duplicates": technical_duplicates,
        "fx_missing": fx_missing,
        **detection_metrics,
    }
    return output, metrics


def write_gl10_output(output_dataframe, context):
    output_file = get_gl_output_file(context)
    print(f"Output workbook: {output_file}")
    stage = perf_counter()

    if not Path(output_file).exists():
        fast_written = write_single_sheet_workbook_fast(
            output_file=output_file,
            sheet_name=SHEET_NAME,
            dataframe=output_dataframe,
            date_columns=DATE_COLUMNS,
            amount_columns=AMOUNT_COLUMNS,
            integer_columns=INTEGER_COLUMNS,
        )
        if fast_written:
            print(f"GL10 write time: {_elapsed(stage):.2f} seconds")
            print(f"GL output workbook: {output_file}")
            print("Sheet replaced: GL10")
            return output_file

    workbook = open_or_create_gl_output_workbook(output_file)
    worksheet = recreate_gl_sheet(workbook, SHEET_NAME)
    write_dataframe_to_sheet(worksheet=worksheet, dataframe=output_dataframe)
    apply_standard_gl_formatting(
        worksheet=worksheet,
        dataframe=output_dataframe,
        date_columns=DATE_COLUMNS,
        amount_columns=AMOUNT_COLUMNS,
        integer_columns=INTEGER_COLUMNS,
    )
    try:
        save_gl_output_workbook(workbook, output_file)
    except PermissionError as error:
        raise PermissionError("Close the workbook and run again.") from error

    print(f"GL10 write time: {_elapsed(stage):.2f} seconds")
    print(f"GL output workbook: {output_file}")
    print("Sheet replaced: GL10")
    return output_file


def run_gl_010(context):
    """Execute GL_010 and write or replace only the GL10 worksheet."""
    total_start = perf_counter()
    print_header(
        "Running GL_010 - Potential Duplicate General Journals: "
        "Same General Ledger Account & Same Description"
    )
    date_from, date_to = _module_dates(context)
    print(f"CONFIG FROM: {date_from:%Y-%m-%d}")
    print(f"CONFIG TO: {date_to:%Y-%m-%d}")
    print("Date basis: Posting Date (BUDAT / Dt.lçto.)")
    print("Description normalization: strict LHA (strip only; no case/accent/punctuation/internal-space changes)")

    read_start = perf_counter()
    bsis_dataframe = load_gl_bsis_data(context)
    bsas_dataframe = load_gl_bsas_data(context)
    master_dataframe = load_gl_master_data(context)
    fx_dataframe = load_gl_fx_rates_data(context)
    print(f"GL10 input read time: {_elapsed(read_start):.2f} seconds")
    print(f"GL master rows loaded: {len(master_dataframe)}")
    print(f"GL FxRates rows loaded: {len(fx_dataframe)}")

    if bsis_dataframe.empty and bsas_dataframe.empty:
        raise FileNotFoundError(
            "GL10 requires at least one GL journal input file:\n"
            "- input/LBR GL_JE_BSIS_YYYYMMDD.xlsx\n"
            "- input/LBR GL_JE_BSAS_YYYYMMDD.xlsx"
        )

    output_dataframe, metrics = create_gl10_duplicate_journals(
        bsis_dataframe=bsis_dataframe,
        bsas_dataframe=bsas_dataframe,
        master_dataframe=master_dataframe,
        fx_dataframe=fx_dataframe,
        context=context,
    )
    write_gl10_output(output_dataframe, context)

    journals = 0 if output_dataframe.empty else output_dataframe[
        ["Company Code", "Fiscal Year", "Document Number"]
    ].drop_duplicates().shape[0]
    print(f"Distinct journals reported: {journals}")
    print(f"Final GL10 lines: {len(output_dataframe)}")
    print(f"Affected companies: {output_dataframe['Company Code'].nunique() if not output_dataframe.empty else 0}")
    print(f"Affected GL accounts: {output_dataframe['GL Account'].nunique() if not output_dataframe.empty else 0}")
    print(f"Rows without FX conversion: {metrics['fx_missing']}")
    print(f"GL10 total time: {_elapsed(total_start):.2f} seconds")
    print()
    return output_dataframe
