"""
GL_003 - General Journal Narrations With Suspicious Words.

Optimized version:
- Filters suspicious narration lines before building the full LHA-like output.
- Searches line and header narrations using normalized text.
- Applies FX only to GL03 result rows, not to all BSIS/BSAS rows.
"""

import re
import unicodedata
from pathlib import Path

import pandas as pd
from core.gl_common import (
    apply_standard_gl_formatting,
    get_gl_output_file,
    normalize_fx_rates,
    normalize_text,
    load_gl_fx_rates_data,
    load_gl_master_data,
    open_or_create_gl_output_workbook,
    recreate_gl_sheet,
    save_gl_output_workbook,
    select_fx_rate_to_usd,
    write_dataframe_to_sheet,
    write_single_sheet_workbook_fast,
)
from modules.GL.gl_002 import (
    AMOUNT_COLUMNS,
    DATE_COLUMNS,
    INTEGER_COLUMNS,
    OPTIONAL_COLUMNS,
    OUTPUT_COLUMNS as GL02_OUTPUT_COLUMNS,
    build_header_total_base,
    build_header_total_map,
    REQUIRED_COLUMNS,
    build_lha_like_gl_lines,
    clean_text_series,
    prepare_source_dataframe,
    print_header,
)


SHEET_NAME = "GL03"


OUTPUT_COLUMNS = GL02_OUTPUT_COLUMNS + ["Matched Word"]


SUSPICIOUS_WORDS = [
    "miscellaneous",
    "misc",
    "general",
    "manual",
    "adjustment",
    "adjust",
    "regularization",
    "regularisation",
    "reclassification",
    "reclass",
    "correction",
    "fix",
    "write off",
    "write-off",
    "others",
    "other",
    "various",
    "varios",
    "varias",
    "otros",
    "otras",
    "diversos",
    "diversas",
    "ajuste",
    "ajustes",
    "regularizacion",
    "regularizacao",
    "reclasificacion",
    "reclassificacao",
    "correccion",
    "correcao",
    "geral",
]


BLANK_NARRATION_LABEL = "(narracion en blanco)"


def normalize_header_name(value):
    """
    Normalize an Excel header for usecols matching.
    """
    return str(value).strip().lower()


def get_gl03_journal_usecols():
    """
    Return the SAP journal columns needed by GL03.

    Reading only needed columns reduces Excel load time on wide SAP extracts while
    keeping the control independent and preserving the GL02 output-building helpers.
    """
    needed_columns = set()

    for possible_names in REQUIRED_COLUMNS.values():
        for possible_name in possible_names:
            needed_columns.add(normalize_header_name(possible_name))

    for possible_names in OPTIONAL_COLUMNS.values():
        for possible_name in possible_names:
            needed_columns.add(normalize_header_name(possible_name))

    return needed_columns


def find_latest_gl_input_file(context, source_name):
    """
    Find the latest GL BSIS/BSAS input file for the requested source.
    """
    input_folder = Path(context["input_folder"])
    source_name = str(source_name).strip().upper()
    expected_prefix = f"lbr gl_je_{source_name.lower()}_"

    candidates = [
        file_path
        for file_path in input_folder.iterdir()
        if file_path.is_file()
        and file_path.name.lower().startswith(expected_prefix)
        and file_path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}
    ]

    if not candidates:
        print(f"GL {source_name} input file: not found")
        return None

    return sorted(candidates, key=lambda file_path: file_path.name.lower())[-1]


def load_gl_journal_data_optimized(context, source_name):
    """
    Load a BSIS/BSAS journal extract with only GL03-relevant columns.
    """
    input_file = find_latest_gl_input_file(context, source_name)

    if input_file is None:
        return pd.DataFrame()

    print(f"GL {source_name} input file: {input_file}")

    needed_columns = get_gl03_journal_usecols()

    return pd.read_excel(
        input_file,
        usecols=lambda column_name: normalize_header_name(column_name) in needed_columns,
    )


def normalize_text_for_matching(value):
    """
    Normalize narration text for suspicious-word matching.
    """
    if value is None or pd.isna(value):
        return ""

    text = str(value).strip().lower()

    if text == "" or text == "nan":
        return ""

    text = unicodedata.normalize("NFKD", text)
    text = "".join(character for character in text if not unicodedata.combining(character))
    text = re.sub(r"\s+", " ", text).strip()

    return text


def normalize_text_series_for_matching(series):
    """
    Vectorized normalization for narration matching.
    """
    return clean_text_series(series).map(normalize_text_for_matching)


def normalize_suspicious_words(words):
    """
    Normalize and de-duplicate the suspicious word list preserving order.
    """
    normalized_words = []
    seen_words = set()

    for word in words:
        normalized_word = normalize_text_for_matching(word)

        if normalized_word == "" or normalized_word in seen_words:
            continue

        normalized_words.append(normalized_word)
        seen_words.add(normalized_word)

    return normalized_words


def build_suspicious_word_pattern(suspicious_word):
    """
    Build a regex pattern for one normalized suspicious word.

    Multi-word terms are matched as normalized substrings. Single-word terms are
    matched with token boundaries to reduce false positives such as matching
    'fix' inside a longer unrelated word.
    """
    escaped_word = re.escape(suspicious_word)

    if " " in suspicious_word or "-" in suspicious_word:
        return escaped_word

    return rf"(?<!\w){escaped_word}(?!\w)"


def apply_suspicious_word_matches(matched_word, text_series, suspicious_words, label_prefix):
    """
    Fill matched_word with the first suspicious word found in text_series.

    The function only evaluates rows not already matched by a higher-priority
    condition and uses pandas string matching instead of row-by-row Python loops.
    """
    unmatched_mask = matched_word == ""

    for suspicious_word in suspicious_words:
        if not unmatched_mask.any():
            break

        pattern = build_suspicious_word_pattern(suspicious_word)
        word_mask = text_series.loc[unmatched_mask].str.contains(
            pattern,
            regex=True,
            na=False,
        )

        if not word_mask.any():
            continue

        matched_index = word_mask[word_mask].index
        matched_word.loc[matched_index] = f"{label_prefix}: {suspicious_word}"
        unmatched_mask.loc[matched_index] = False

    return matched_word


def build_matched_word_series(prepared_dataframe, required_columns):
    """
    Build GL03 matched-word labels from line and header narrations.

    Priority:
    1. Blank line narration.
    2. Suspicious word in line narration.
    3. Suspicious word in header narration.

    Header-level blanks are not reported by themselves because SAP documents can
    have a blank header text while retaining a meaningful line text. Reporting
    header blanks alone creates a very large low-value output and slows Excel
    generation without improving review quality.
    """
    suspicious_words = normalize_suspicious_words(SUSPICIOUS_WORDS)

    line_memo = normalize_text_series_for_matching(
        prepared_dataframe[required_columns["line_text"]]
    )
    journal_memo = normalize_text_series_for_matching(
        prepared_dataframe[required_columns["header_text"]]
    )

    matched_word = pd.Series("", index=prepared_dataframe.index, dtype="object")

    line_blank_mask = line_memo == ""
    matched_word.loc[line_blank_mask] = BLANK_NARRATION_LABEL

    matched_word = apply_suspicious_word_matches(
        matched_word=matched_word,
        text_series=line_memo,
        suspicious_words=suspicious_words,
        label_prefix="Line Memo",
    )
    matched_word = apply_suspicious_word_matches(
        matched_word=matched_word,
        text_series=journal_memo,
        suspicious_words=suspicious_words,
        label_prefix="Journal Memo",
    )

    return matched_word


def get_suspicious_narration_mask(prepared_dataframe, required_columns):
    """
    Identify rows with blank or suspicious GL narrations.
    """
    matched_word = build_matched_word_series(
        prepared_dataframe=prepared_dataframe,
        required_columns=required_columns,
    )

    return matched_word != "", matched_word


def create_gl03_lines_for_source(source_dataframe, master_dataframe, context, source_name):
    """
    Create GL03 output rows for one SAP source, preserving source-level totals.
    """
    if source_dataframe.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS), pd.DataFrame()

    prepared_dataframe, required_columns, optional_columns = prepare_source_dataframe(
        source_dataframe=source_dataframe,
        context=context,
        source_name=source_name,
    )

    if prepared_dataframe.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS), pd.DataFrame()

    print(f"{source_name} identifying suspicious narration journal lines before output build...")

    suspicious_narration_mask, matched_word = get_suspicious_narration_mask(
        prepared_dataframe=prepared_dataframe,
        required_columns=required_columns,
    )
    suspicious_narration_dataframe = prepared_dataframe[suspicious_narration_mask].copy()

    print(
        f"{source_name} suspicious narration journal lines kept: "
        f"{len(suspicious_narration_dataframe)}"
    )

    if suspicious_narration_dataframe.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS), pd.DataFrame()

    suspicious_trans_ids = suspicious_narration_dataframe["_TRANS_ID"].drop_duplicates()
    header_total_population = prepared_dataframe[
        prepared_dataframe["_TRANS_ID"].isin(suspicious_trans_ids)
    ]
    header_total_base_dataframe = build_header_total_base(
        prepared_dataframe=header_total_population,
        required_columns=required_columns,
    )

    print(f"{source_name} creating LHA-like GL03 line rows...")

    output_dataframe = build_lha_like_gl_lines(
        prepared_dataframe=suspicious_narration_dataframe,
        required_columns=required_columns,
        optional_columns=optional_columns,
        master_dataframe=master_dataframe,
    )
    output_dataframe["Matched Word"] = matched_word.loc[
        suspicious_narration_dataframe.index
    ].values

    print(f"{source_name} LHA-like GL03 line rows created: {len(output_dataframe)}")

    return output_dataframe[OUTPUT_COLUMNS], header_total_base_dataframe


def add_usd_fields_fast(line_dataframe, fx_dataframe):
    """
    Add GL03 USD fields using one FX lookup per currency/date pair.

    GL03 can return a large population. Reusing the GL02 row-by-row FX helper is
    correct but slow for more than one hundred thousand output rows, so this
    version keeps the same Posting Date primary / Document Date fallback logic
    and applies the selected rate with vectorized pandas operations.
    """
    result = line_dataframe.copy()

    if result.empty or fx_dataframe.empty:
        return result

    normalized_fx_dataframe = normalize_fx_rates(fx_dataframe)

    fx_date = pd.to_datetime(result["Posting Date"], errors="coerce")
    document_date = pd.to_datetime(result["Document Date"], errors="coerce")
    fx_date = fx_date.where(fx_date.notna(), document_date)

    currency = result["FC Currency"].map(lambda value: normalize_text(value).upper())
    rate_key = pd.DataFrame(
        {
            "Currency": currency,
            "FX Date": fx_date.dt.strftime("%Y-%m-%d").fillna(""),
        },
        index=result.index,
    )

    unique_rate_keys = rate_key.drop_duplicates()
    fx_lookup = {}

    for _, key_row in unique_rate_keys.iterrows():
        key_currency = key_row["Currency"]
        key_date_text = key_row["FX Date"]

        if key_date_text == "":
            requested_date = pd.NaT
        else:
            requested_date = pd.to_datetime(key_date_text, errors="coerce")

        fx_details = select_fx_rate_to_usd(
            normalized_fx_dataframe=normalized_fx_dataframe,
            currency=key_currency,
            requested_date=requested_date,
        )
        fx_lookup[(key_currency, key_date_text)] = fx_details

    usd_method = []
    usd_rate = []
    usd_rate_date = []
    fx_to_usd = []

    for key_currency, key_date_text in zip(rate_key["Currency"], rate_key["FX Date"]):
        fx_details = fx_lookup.get((key_currency, key_date_text))

        if fx_details is None:
            usd_method.append("")
            usd_rate.append(pd.NA)
            usd_rate_date.append(pd.NaT)
            fx_to_usd.append(pd.NA)
            continue

        usd_method.append(fx_details["method"])
        usd_rate.append(fx_details["usd_rate"])
        usd_rate_date.append(fx_details["rate_date"])
        fx_to_usd.append(fx_details["fx_to_usd"])

    fx_to_usd = pd.to_numeric(pd.Series(fx_to_usd, index=result.index), errors="coerce")
    line_amount_local = pd.to_numeric(result["Line Amount Local"], errors="coerce")
    header_total_local = pd.to_numeric(result["Header Total Local"], errors="coerce")

    result["Company System Currency"] = "USD"
    result["Line Amount USD"] = line_amount_local * fx_to_usd
    result["USD Method"] = usd_method
    result["USD Rate"] = usd_rate
    result["USD Rate Date"] = usd_rate_date
    result["Header Total USD"] = header_total_local * fx_to_usd

    missing_usd = result["Line Amount USD"].isna().sum()

    if missing_usd > 0:
        print()
        print("WARNING: Some GL FX rates were not found.")
        print("USD columns are blank for those rows.")
        print(f"Rows without USD amount: {missing_usd}")
        print()

    return result


def create_gl03_suspicious_narration_journals(
    bsis_dataframe,
    bsas_dataframe,
    master_dataframe,
    fx_dataframe,
    context,
):
    """
    Create the consolidated GL03 suspicious-narration output.
    """
    if bsis_dataframe.empty and bsas_dataframe.empty:
        raise FileNotFoundError(
            "No GL input files were found. Expected at least one of:\n"
            "- input/LBR GL_JE_BSIS_YYYYMMDD.xlsx\n"
            "- input/LBR GL_JE_BSAS_YYYYMMDD.xlsx"
        )

    source_dataframes = []
    header_total_base_dataframes = []

    bsis_lines, bsis_header_total_base = create_gl03_lines_for_source(
        source_dataframe=bsis_dataframe,
        master_dataframe=master_dataframe,
        context=context,
        source_name="BSIS",
    )

    if not bsis_lines.empty:
        source_dataframes.append(bsis_lines)

    if not bsis_header_total_base.empty:
        header_total_base_dataframes.append(bsis_header_total_base)

    bsas_lines, bsas_header_total_base = create_gl03_lines_for_source(
        source_dataframe=bsas_dataframe,
        master_dataframe=master_dataframe,
        context=context,
        source_name="BSAS",
    )

    if not bsas_lines.empty:
        source_dataframes.append(bsas_lines)

    if not bsas_header_total_base.empty:
        header_total_base_dataframes.append(bsas_header_total_base)

    if not source_dataframes:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    output_dataframe = pd.concat(source_dataframes, ignore_index=True)

    if header_total_base_dataframes:
        header_total_base_dataframe = pd.concat(
            header_total_base_dataframes,
            ignore_index=True,
        )
        header_total_map = build_header_total_map(header_total_base_dataframe)
        output_dataframe["Header Total Local"] = output_dataframe["TransId"].map(
            header_total_map
        )

    output_dataframe = add_usd_fields_fast(
        line_dataframe=output_dataframe,
        fx_dataframe=fx_dataframe,
    )

    output_dataframe = output_dataframe.sort_values(
        by=["Posting Date", "CoCo", "Journal Number", "Line"],
        kind="stable",
    ).reset_index(drop=True)

    return output_dataframe[OUTPUT_COLUMNS].copy()


def run_gl_003(context):
    """
    Execute GL_003 and write only the GL03 sheet.
    """
    print_header("Running GL_003 - General Journal Narrations With Suspicious Words")

    bsis_dataframe = load_gl_journal_data_optimized(context, "BSIS")
    print(f"BSIS rows loaded: {len(bsis_dataframe)}")

    bsas_dataframe = load_gl_journal_data_optimized(context, "BSAS")
    print(f"BSAS rows loaded: {len(bsas_dataframe)}")

    master_dataframe = load_gl_master_data(context)
    print(f"GL master rows loaded: {len(master_dataframe)}")

    fx_dataframe = load_gl_fx_rates_data(context)
    print(f"GL FxRates rows loaded: {len(fx_dataframe)}")

    print("Creating optimized GL03 suspicious narration journal lines...")

    output_dataframe = create_gl03_suspicious_narration_journals(
        bsis_dataframe=bsis_dataframe,
        bsas_dataframe=bsas_dataframe,
        master_dataframe=master_dataframe,
        fx_dataframe=fx_dataframe,
        context=context,
    )

    print(f"GL03 suspicious narration line rows prepared: {len(output_dataframe)}")

    output_file = get_gl_output_file(context)

    print(f"Output workbook: {output_file}")

    if not output_file.exists():
        print("Output workbook does not exist. Using fast GL03 writer...")

        fast_written = write_single_sheet_workbook_fast(
            output_file=output_file,
            sheet_name=SHEET_NAME,
            dataframe=output_dataframe,
            date_columns=DATE_COLUMNS,
            amount_columns=AMOUNT_COLUMNS,
            integer_columns=INTEGER_COLUMNS,
        )

        if fast_written:
            print(f"GL03 rows written: {len(output_dataframe)}")
            print(f"GL output workbook: {output_file}")
            print()
            return

    print("Output workbook already exists or fast writer is unavailable.")
    print("Using preserve-sheets openpyxl writer...")

    workbook = open_or_create_gl_output_workbook(output_file)
    worksheet = recreate_gl_sheet(workbook, SHEET_NAME)

    print("Writing GL03 rows to worksheet...")
    write_dataframe_to_sheet(worksheet, output_dataframe)

    print("Applying GL03 formatting...")
    apply_standard_gl_formatting(
        worksheet=worksheet,
        dataframe=output_dataframe,
        date_columns=DATE_COLUMNS,
        amount_columns=AMOUNT_COLUMNS,
        integer_columns=INTEGER_COLUMNS,
    )

    print("Saving GL output workbook...")
    save_gl_output_workbook(workbook, output_file)

    print(f"GL03 rows written: {len(output_dataframe)}")
    print(f"GL output workbook: {output_file}")
    print()
