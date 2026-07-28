"""
GL_008 - General Journals To Accounts Inactive For More Than Six Months.

The control reports journal lines posted during the configured period when the
same Company Code + GL Account had no movement during the preceding 180 days.

Historical inputs are read independently from files whose names contain all of:
"LBR", "GL_JE", the source name (BSIS or BSAS), and "Historic". Normal BSIS
and BSAS inputs continue to be loaded through core.gl_common.

Important functional choices:
- The threshold is exactly 180 days, not six calendar months.
- Activity is evaluated sequentially for each posting event. All lines of a
  qualifying Company + Account + Document event are reported.
- A missing preceding movement is reportable only because the supplied historic
  files are expected to cover the complete 180-day lookback. The reason does
  not claim that the account never had older movements.
- Formal account blocking is not evaluated because no reliable inactive-from
  master-data field is available in the current SAP LBR inputs.
"""

from pathlib import Path
import re
import unicodedata

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
    open_or_create_gl_output_workbook,
    recreate_gl_sheet,
    save_gl_output_workbook,
    select_fx_rate_to_usd,
    write_dataframe_to_sheet,
    write_single_sheet_workbook_fast,
)


SHEET_NAME = "GL08"
DEFAULT_INACTIVE_DAYS = 180
HISTORIC_TOKEN = "historic"
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}


FIELD_ALIASES = {
    "company_code": ["Empr", "BKPF-BUKRS", "BUKRS"],
    "document_number": ["Nº doc.", "Nº doc", "BKPF-BELNR", "BELNR"],
    "fiscal_year": ["Ano", "BKPF-GJAHR", "GJAHR"],
    "document_type": ["Tp.doc.", "Tp.doc", "BKPF-BLART", "BLART"],
    "document_date": ["Data doc.", "Data doc", "BKPF-BLDAT", "BLDAT"],
    "posting_date": [
        "Dt.lçto.",
        "Dt.lçto",
        "Dt.lcto.",
        "Dt.lcto",
        "BKPF-BUDAT",
        "BUDAT",
    ],
    "entry_date": ["Dt.entr.", "Dt.entr", "BKPF-CPUDT", "CPUDT"],
    "approver_user": [
        "Nome do usuário",
        "Nome do usuario",
        "BKPF-USNAM",
        "USNAM",
    ],
    "creator_user": [
        "Pré-edição",
        "Pre-edição",
        "Pré-edicao",
        "Pre-edicao",
        "Pré edição",
        "Pre edição",
        "BKPF-PPNAM",
        "PPNAM",
    ],
    "transaction_code": ["CódT", "CodT", "BKPF-TCODE", "TCODE"],
    "header_text": [
        "Texto cabeçalho documento",
        "Texto cabecalho documento",
        "BKPF-BKTXT",
        "BKTXT",
    ],
    "currency": ["Moeda", "BKPF-WAERS", "WAERS"],
    "line_item": ["Itm", "BSIS-BUZEI", "BSAS-BUZEI", "BUZEI"],
    "gl_account": [
        "Razão",
        "Razao",
        "BSIS-HKONT",
        "BSAS-HKONT",
        "HKONT",
    ],
    "debit_credit_indicator": [
        "D/C",
        "BSIS-SHKZG",
        "BSAS-SHKZG",
        "SHKZG",
    ],
    "amount_local_currency": [
        "Montante em MI",
        "BSIS-DMBTR",
        "BSAS-DMBTR",
        "DMBTR",
    ],
    "amount_document_currency": [
        "Montante",
        "BSIS-WRBTR",
        "BSAS-WRBTR",
        "WRBTR",
    ],
    "line_text": [
        "Texto",
        "BSIS-SGTXT",
        "BSAS-SGTXT",
        "SGTXT",
    ],
    "fiscal_period": [
        "Período",
        "Periodo",
        "BKPF-MONAT",
        "MONAT",
    ],
}


REQUIRED_FIELDS = {
    "company_code",
    "document_number",
    "fiscal_year",
    "posting_date",
    "line_item",
    "gl_account",
}


OUTPUT_COLUMNS = [
    "Company Code",
    "Company Name",
    "GL Account",
    "GL Account Description",
    "Document Type",
    "Document Number",
    "Document Text",
    "Line Text",
    "Line Number",
    "Debit",
    "Credit",
    "Amount in Local Currency",
    "Amount in Reporting Currency",
    "Report Currency",
    "Amount in Document Currency",
    "Document Currency",
    "Document Date",
    "Fiscal Year",
    "Date Entered",
    "Posting Date",
    "Create User ID",
    "Create User Name",
    "Approver User ID",
    "Approver User Name",
    "Fiscal Period",
    "Transaction Code",
    "Transaction Code Description",
    "Last Movement Before Current Posting",
    "Inactive Days",
    "Months Since Last Movement",
    "History Coverage Status",
    "Inactive Flag (b) No Movement",
    "GL08 Reason",
]


DATE_COLUMNS = {
    "Document Date",
    "Date Entered",
    "Posting Date",
    "Last Movement Before Current Posting",
}


AMOUNT_COLUMNS = {
    "Debit",
    "Credit",
    "Amount in Local Currency",
    "Amount in Reporting Currency",
    "Amount in Document Currency",
}


INTEGER_COLUMNS = {
    "Inactive Days",
}


def print_header(title):
    print(title)
    print("-" * len(title))


def normalize_header(value):
    text = unicodedata.normalize(
        "NFKD",
        str(value).strip().lower(),
    )

    return "".join(
        character
        for character in text
        if not unicodedata.combining(character)
    )


def normalize_text_series(series):
    result = series.where(
        series.notna(),
        "",
    ).astype(str).str.strip()

    result = result.mask(
        result.str.lower().isin({"nan", "none", "nat"}),
        "",
    )

    return result.str.replace(
        r"\.0$",
        "",
        regex=True,
    )


def normalize_code_series(series):
    return normalize_text_series(series)


def normalize_company_series(series):
    result = normalize_text_series(series)
    numeric_mask = result.str.fullmatch(
        r"\d+",
        na=False,
    )

    numeric_values = pd.to_numeric(
        result.where(numeric_mask),
        errors="coerce",
    )

    normalized_numeric = numeric_values.astype(
        "Int64",
    ).astype(str)

    return result.where(
        ~numeric_mask,
        normalized_numeric,
    )


def parse_date_series(series):
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(
            series,
            errors="coerce",
        ).dt.normalize()

    text = normalize_text_series(series)

    result = pd.Series(
        pd.NaT,
        index=series.index,
        dtype="datetime64[ns]",
    )

    yyyymmdd = text.str.fullmatch(
        r"\d{8}",
        na=False,
    )

    iso_date = text.str.match(
        r"\d{4}-\d{2}-\d{2}",
        na=False,
    )

    remaining = (
        (text != "")
        & ~yyyymmdd
        & ~iso_date
    )

    if yyyymmdd.any():
        result.loc[yyyymmdd] = pd.to_datetime(
            text.loc[yyyymmdd],
            format="%Y%m%d",
            errors="coerce",
        )

    if iso_date.any():
        result.loc[iso_date] = pd.to_datetime(
            text.loc[iso_date].str[:10],
            format="%Y-%m-%d",
            errors="coerce",
        )

    if remaining.any():
        result.loc[remaining] = pd.to_datetime(
            text.loc[remaining],
            errors="coerce",
            dayfirst=True,
        )

    return result.dt.normalize()


def parse_number_series(series):
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(
            series,
            errors="coerce",
        )

    text = normalize_text_series(series)
    text = text.str.replace(
        " ",
        "",
        regex=False,
    )

    both = (
        text.str.contains(",", regex=False)
        & text.str.contains(".", regex=False)
    )

    comma_decimal = (
        both
        & (text.str.rfind(",") > text.str.rfind("."))
    )

    dot_decimal = both & ~comma_decimal

    text.loc[comma_decimal] = (
        text.loc[comma_decimal]
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )

    text.loc[dot_decimal] = (
        text.loc[dot_decimal]
        .str.replace(",", "", regex=False)
    )

    only_comma = (
        text.str.contains(",", regex=False)
        & ~text.str.contains(".", regex=False)
    )

    text.loc[only_comma] = (
        text.loc[only_comma]
        .str.replace(",", ".", regex=False)
    )

    return pd.to_numeric(
        text,
        errors="coerce",
    )


def empty_series(dataframe, dtype="object"):
    return pd.Series(
        index=dataframe.index,
        dtype=dtype,
    )


def get_series(
    dataframe,
    columns,
    field_name,
    dtype="object",
):
    column_name = columns.get(field_name)

    if column_name is None:
        if dtype == "datetime64[ns]":
            return pd.Series(
                pd.NaT,
                index=dataframe.index,
                dtype=dtype,
            )

        return pd.Series(
            "",
            index=dataframe.index,
            dtype=dtype,
        )

    return dataframe[column_name]


def find_column(dataframe, aliases):
    normalized_columns = {}

    for column in dataframe.columns:
        normalized_columns.setdefault(
            normalize_header(column),
            [],
        ).append(column)

    for alias in aliases:
        matches = normalized_columns.get(
            normalize_header(alias),
            [],
        )

        if matches:
            return matches[0]

    technical_aliases = [
        normalize_header(alias)
        for alias in aliases
        if "-" in alias
    ]

    for column in dataframe.columns:
        normalized = normalize_header(column)
        base = re.sub(
            r"\.\d+$",
            "",
            normalized,
        )

        if base in technical_aliases:
            return column

    return None


def resolve_columns(
    dataframe,
    source_name,
    require_all=True,
):
    columns = {
        field_name: find_column(
            dataframe,
            aliases,
        )
        for field_name, aliases in FIELD_ALIASES.items()
    }

    if require_all:
        missing = sorted(
            field
            for field in REQUIRED_FIELDS
            if columns.get(field) is None
        )

        if missing:
            raise ValueError(
                f"GL08 could not resolve required fields in "
                f"{source_name}: "
                + ", ".join(missing)
                + f". Available headers: "
                f"{list(dataframe.columns)}"
            )

    return columns


def find_historic_input_files(
    context,
    source_name,
):
    input_folder = Path(context["input_folder"])
    source_token = source_name.lower()
    candidates = []

    for file_path in input_folder.iterdir():
        lower_name = file_path.name.lower()

        if (
            file_path.is_file()
            and file_path.suffix.lower() in EXCEL_EXTENSIONS
            and "lbr" in lower_name
            and "gl_je" in lower_name
            and source_token in lower_name
            and HISTORIC_TOKEN in lower_name
        ):
            candidates.append(file_path)

    return sorted(
        candidates,
        key=lambda path: path.name.lower(),
    )


def load_historic_source(
    context,
    source_name,
):
    files = find_historic_input_files(
        context,
        source_name,
    )

    if not files:
        print(
            f"GL08 {source_name} Historic input file: "
            f"not found"
        )
        return pd.DataFrame()

    frames = []

    for file_path in files:
        print(
            f"GL08 {source_name} Historic input file: "
            f"{file_path}"
        )

        frame = pd.read_excel(file_path)
        frame["_INPUT_FILE"] = file_path.name
        frames.append(frame)

    if frames:
        result = pd.concat(
            frames,
            ignore_index=True,
        )
    else:
        result = pd.DataFrame()

    print(
        f"GL08 {source_name} Historic files read: "
        f"{len(files)}"
    )

    print(
        f"GL08 {source_name} Historic rows loaded: "
        f"{len(result)}"
    )

    return result


def prepare_activity_source(
    dataframe,
    context,
    source_name,
    is_historic,
):
    if dataframe.empty:
        return pd.DataFrame()

    columns = resolve_columns(
        dataframe,
        source_name,
    )

    module_config = context["module"]

    filtered = filter_by_company(
        dataframe=dataframe,
        company_column=columns["company_code"],
        companies_filter=module_config.get(
            "companies",
            "",
        ),
    ).copy()

    posting_date = parse_date_series(
        filtered[columns["posting_date"]]
    )

    result = pd.DataFrame(
        index=filtered.index,
    )

    result["_COMPANY"] = normalize_company_series(
        filtered[columns["company_code"]]
    )

    result["_ACCOUNT"] = normalize_code_series(
        filtered[columns["gl_account"]]
    )

    result["_FISCAL_YEAR"] = normalize_code_series(
        filtered[columns["fiscal_year"]]
    )

    result["_DOCUMENT"] = normalize_code_series(
        filtered[columns["document_number"]]
    )

    result["_LINE"] = normalize_code_series(
        filtered[columns["line_item"]]
    )

    result["_POSTING_DATE"] = posting_date
    result["_SOURCE"] = source_name
    result["_IS_HISTORIC"] = is_historic

    result["_SOURCE_PRIORITY"] = {
        ("BSIS", True): 1,
        ("BSAS", True): 2,
        ("BSIS", False): 3,
        ("BSAS", False): 4,
    }[(source_name, is_historic)]

    result["_RAW_INDEX"] = filtered.index

    valid = (
        (result["_COMPANY"] != "")
        & (result["_ACCOUNT"] != "")
        & (result["_FISCAL_YEAR"] != "")
        & (result["_DOCUMENT"] != "")
        & (result["_LINE"] != "")
        & result["_POSTING_DATE"].notna()
    )

    invalid_rows = int(
        (~valid).sum()
    )

    if invalid_rows:
        print(
            f"WARNING: {source_name} rows excluded for "
            f"incomplete GL08 key/date: {invalid_rows}"
        )

    result = result.loc[valid].reset_index(
        drop=True,
    )

    result["_RAW_ROW"] = list(
        filtered.loc[valid].to_dict("records")
    )

    result["_COLUMN_MAP"] = [
        columns
    ] * len(result)

    return result


def build_activity_population(
    normal_sources,
    historic_sources,
    context,
):
    frames = []

    for source_name in ("BSIS", "BSAS"):
        historic = prepare_activity_source(
            historic_sources[source_name],
            context,
            source_name,
            True,
        )

        normal = prepare_activity_source(
            normal_sources[source_name],
            context,
            source_name,
            False,
        )

        if not historic.empty:
            frames.append(historic)

        if not normal.empty:
            frames.append(normal)

    if not frames:
        return pd.DataFrame()

    activity = pd.concat(
        frames,
        ignore_index=True,
    )

    line_key = [
        "_COMPANY",
        "_FISCAL_YEAR",
        "_DOCUMENT",
        "_LINE",
    ]

    activity = activity.sort_values(
        "_SOURCE_PRIORITY",
        kind="stable",
    )

    before = len(activity)

    activity = activity.drop_duplicates(
        line_key,
        keep="last",
    ).reset_index(drop=True)

    print(
        "GL08 duplicated BSIS/BSAS lines removed: "
        f"{before - len(activity)}"
    )

    return activity


def identify_exceptions(
    activity,
    context,
):
    if activity.empty:
        return activity

    module_config = context["module"]

    from_date = pd.to_datetime(
        module_config.get("from", ""),
        errors="coerce",
    )

    to_date = pd.to_datetime(
        module_config.get("to", ""),
        errors="coerce",
    )

    if pd.isna(from_date) or pd.isna(to_date):
        raise ValueError(
            "GL08 requires valid FROM and TO dates "
            "in config.xlsx."
        )

    from_date = from_date.normalize()
    to_date = to_date.normalize()

    history = activity[
        activity["_POSTING_DATE"] < from_date
    ]

    if history.empty:
        raise ValueError(
            "GL08 Historic inputs contain no valid "
            "movements before CONFIG FROM."
        )

    print(
        "GL08 historic posting date range found: "
        f"{history['_POSTING_DATE'].min().date()} to "
        f"{history['_POSTING_DATE'].max().date()}"
    )

    in_period = activity["_POSTING_DATE"].between(
        from_date,
        to_date,
        inclusive="both",
    )

    print(
        "GL08 rows inside configured period: "
        f"{int(in_period.sum())}"
    )

    event_key = [
        "_COMPANY",
        "_ACCOUNT",
        "_POSTING_DATE",
        "_FISCAL_YEAR",
        "_DOCUMENT",
    ]

    events = (
        activity[event_key]
        .drop_duplicates()
        .sort_values(
            event_key,
            kind="stable",
        )
    )

    events["_PREVIOUS_MOVEMENT"] = events.groupby(
        [
            "_COMPANY",
            "_ACCOUNT",
        ],
        sort=False,
    )["_POSTING_DATE"].shift(1)

    events["_INACTIVE_DAYS"] = (
        events["_POSTING_DATE"]
        - events["_PREVIOUS_MOVEMENT"]
    ).dt.days

    events["_IN_PERIOD"] = events[
        "_POSTING_DATE"
    ].between(
        from_date,
        to_date,
        inclusive="both",
    )

    events["_NO_PRIOR_IN_WINDOW"] = events[
        "_PREVIOUS_MOVEMENT"
    ].isna()

    events["_QUALIFIES"] = (
        events["_IN_PERIOD"]
        & (
            events["_NO_PRIOR_IN_WINDOW"]
            | (
                events["_INACTIVE_DAYS"]
                >= DEFAULT_INACTIVE_DAYS
            )
        )
    )

    qualifying = events.loc[
        events["_QUALIFIES"],
        event_key
        + [
            "_PREVIOUS_MOVEMENT",
            "_INACTIVE_DAYS",
            "_NO_PRIOR_IN_WINDOW",
        ],
    ]

    exceptions = activity.merge(
        qualifying,
        on=event_key,
        how="inner",
    )

    exceptions = exceptions[
        ~exceptions["_IS_HISTORIC"]
    ].copy()

    exceptions["_HISTORY_STATUS"] = (
        "Previous movement identified"
    )

    exceptions.loc[
        exceptions["_NO_PRIOR_IN_WINDOW"],
        "_HISTORY_STATUS",
    ] = (
        "No movement found in complete "
        "180-day lookback"
    )

    exceptions["_REASON"] = (
        "No movement for 180 days or more"
    )

    exceptions.loc[
        exceptions["_NO_PRIOR_IN_WINDOW"],
        "_REASON",
    ] = (
        "No movement found in complete "
        "180-day lookback"
    )

    return exceptions.sort_values(
        [
            "_COMPANY",
            "_POSTING_DATE",
            "_DOCUMENT",
            "_LINE",
        ],
        kind="stable",
    ).reset_index(drop=True)


def master_maps(master_dataframe):
    if master_dataframe.empty:
        return {}, {}

    return (
        build_company_name_map(
            master_dataframe
        ),
        build_gl_account_name_map(
            master_dataframe
        ),
    )


def raw_value(
    row,
    field_name,
):
    raw = row["_RAW_ROW"]
    columns = row["_COLUMN_MAP"]
    column = columns.get(field_name)

    if column is None:
        return ""

    return raw.get(
        column,
        "",
    )


def build_output(
    exceptions,
    master_dataframe,
):
    if exceptions.empty:
        return pd.DataFrame(
            columns=OUTPUT_COLUMNS
        )

    company_names, account_names = master_maps(
        master_dataframe
    )

    rows = []

    for _, row in exceptions.iterrows():
        indicator = str(
            raw_value(
                row,
                "debit_credit_indicator",
            )
        ).strip().upper()

        local_abs = parse_number_series(
            pd.Series(
                [
                    raw_value(
                        row,
                        "amount_local_currency",
                    )
                ]
            )
        ).abs().iloc[0]

        document_abs = parse_number_series(
            pd.Series(
                [
                    raw_value(
                        row,
                        "amount_document_currency",
                    )
                ]
            )
        ).abs().iloc[0]

        if (
            indicator == "H"
            and pd.notna(local_abs)
        ):
            local_signed = -local_abs
        else:
            local_signed = local_abs

        if (
            indicator == "H"
            and pd.notna(document_abs)
        ):
            document_signed = -document_abs
        else:
            document_signed = document_abs

        inactive_days = row["_INACTIVE_DAYS"]

        rows.append(
            {
                "Company Code": row["_COMPANY"],
                "Company Name": company_names.get(
                    row["_COMPANY"],
                    "",
                ),
                "GL Account": row["_ACCOUNT"],
                "GL Account Description": (
                    account_names.get(
                        row["_ACCOUNT"],
                        "",
                    )
                ),
                "Document Type": str(
                    raw_value(
                        row,
                        "document_type",
                    )
                ).strip(),
                "Document Number": row["_DOCUMENT"],
                "Document Text": str(
                    raw_value(
                        row,
                        "header_text",
                    )
                ).strip(),
                "Line Text": str(
                    raw_value(
                        row,
                        "line_text",
                    )
                ).strip(),
                "Line Number": row["_LINE"],
                "Debit": (
                    local_abs
                    if indicator == "S"
                    else pd.NA
                ),
                "Credit": (
                    local_abs
                    if indicator == "H"
                    else pd.NA
                ),
                "Amount in Local Currency": (
                    local_signed
                ),
                "Amount in Reporting Currency": (
                    pd.NA
                ),
                "Report Currency": "",
                "Amount in Document Currency": (
                    document_signed
                ),
                "Document Currency": str(
                    raw_value(
                        row,
                        "currency",
                    )
                ).strip().upper(),
                "Document Date": parse_date_series(
                    pd.Series(
                        [
                            raw_value(
                                row,
                                "document_date",
                            )
                        ]
                    )
                ).iloc[0],
                "Fiscal Year": row["_FISCAL_YEAR"],
                "Date Entered": parse_date_series(
                    pd.Series(
                        [
                            raw_value(
                                row,
                                "entry_date",
                            )
                        ]
                    )
                ).iloc[0],
                "Posting Date": row["_POSTING_DATE"],
                "Create User ID": str(
                    raw_value(
                        row,
                        "creator_user",
                    )
                ).strip(),
                "Create User Name": "",
                "Approver User ID": str(
                    raw_value(
                        row,
                        "approver_user",
                    )
                ).strip(),
                "Approver User Name": "",
                "Fiscal Period": str(
                    raw_value(
                        row,
                        "fiscal_period",
                    )
                ).strip(),
                "Transaction Code": str(
                    raw_value(
                        row,
                        "transaction_code",
                    )
                ).strip(),
                "Transaction Code Description": "",
                "Last Movement Before Current Posting": (
                    row["_PREVIOUS_MOVEMENT"]
                ),
                "Inactive Days": (
                    inactive_days
                    if pd.notna(inactive_days)
                    else pd.NA
                ),
                "Months Since Last Movement": (
                    round(
                        float(inactive_days) / 30,
                        1,
                    )
                    if pd.notna(inactive_days)
                    else pd.NA
                ),
                "History Coverage Status": (
                    row["_HISTORY_STATUS"]
                ),
                "Inactive Flag (b) No Movement": "Y",
                "GL08 Reason": row["_REASON"],
            }
        )

    return pd.DataFrame(
        rows,
        columns=OUTPUT_COLUMNS,
    )


def add_reporting_currency(
    output_dataframe,
    fx_dataframe,
):
    result = output_dataframe.copy()

    if result.empty or fx_dataframe.empty:
        if fx_dataframe.empty:
            print(
                "FxRates input was not found. "
                "GL08 reporting-currency columns "
                "remain blank."
            )

        return result

    normalized_fx = normalize_fx_rates(
        fx_dataframe
    )

    fx_cache = {}

    for index, row in result.iterrows():
        currency = str(
            row["Document Currency"]
        ).strip().upper()

        requested_date = row["Posting Date"]

        cache_key = (
            currency,
            requested_date.strftime("%Y-%m-%d"),
        )

        if cache_key not in fx_cache:
            fx_cache[cache_key] = (
                select_fx_rate_to_usd(
                    normalized_fx_dataframe=(
                        normalized_fx
                    ),
                    currency=currency,
                    requested_date=requested_date,
                )
            )

        fx_details = fx_cache[cache_key]

        amount = pd.to_numeric(
            pd.Series(
                [
                    row[
                        "Amount in Document Currency"
                    ]
                ]
            ),
            errors="coerce",
        ).iloc[0]

        if (
            fx_details is None
            or pd.isna(amount)
        ):
            continue

        result.at[
            index,
            "Amount in Reporting Currency",
        ] = (
            amount
            * fx_details["fx_to_usd"]
        )

        result.at[
            index,
            "Report Currency",
        ] = "USD"

    return result


def write_output(
    output_dataframe,
    context,
):
    output_file = get_gl_output_file(
        context
    )

    print(
        f"Output workbook: {output_file}"
    )

    if not output_file.exists():
        written = write_single_sheet_workbook_fast(
            output_file=output_file,
            sheet_name=SHEET_NAME,
            dataframe=output_dataframe,
            date_columns=DATE_COLUMNS,
            amount_columns=AMOUNT_COLUMNS,
            integer_columns=INTEGER_COLUMNS,
        )

        if written:
            return output_file

    workbook = open_or_create_gl_output_workbook(
        output_file
    )

    worksheet = recreate_gl_sheet(
        workbook,
        SHEET_NAME,
    )

    write_dataframe_to_sheet(
        worksheet=worksheet,
        dataframe=output_dataframe,
    )

    apply_standard_gl_formatting(
        worksheet=worksheet,
        dataframe=output_dataframe,
        date_columns=DATE_COLUMNS,
        amount_columns=AMOUNT_COLUMNS,
        integer_columns=INTEGER_COLUMNS,
    )

    save_gl_output_workbook(
        workbook,
        output_file,
    )

    return output_file


def run_gl_008(context):
    print_header(
        "Running GL_008 - General Journals To "
        "Accounts Inactive For More Than Six Months"
    )

    print(
        f"GL08 threshold: "
        f"{DEFAULT_INACTIVE_DAYS} exact days"
    )

    normal_sources = {
        "BSIS": load_gl_bsis_data(context),
        "BSAS": load_gl_bsas_data(context),
    }

    historic_sources = {
        "BSIS": load_historic_source(
            context,
            "BSIS",
        ),
        "BSAS": load_historic_source(
            context,
            "BSAS",
        ),
    }

    print(
        "BSIS current rows loaded: "
        f"{len(normal_sources['BSIS'])}"
    )

    print(
        "BSAS current rows loaded: "
        f"{len(normal_sources['BSAS'])}"
    )

    if (
        normal_sources["BSIS"].empty
        and normal_sources["BSAS"].empty
    ):
        raise FileNotFoundError(
            "GL08 requires at least one current "
            "BSIS or BSAS input file."
        )

    if (
        historic_sources["BSIS"].empty
        and historic_sources["BSAS"].empty
    ):
        raise FileNotFoundError(
            "GL08 requires BSIS Historic and/or "
            "BSAS Historic inputs covering the "
            "complete 180-day lookback."
        )

    master_dataframe = load_gl_master_data(
        context
    )

    fx_dataframe = load_gl_fx_rates_data(
        context
    )

    activity = build_activity_population(
        normal_sources,
        historic_sources,
        context,
    )

    print(
        "GL08 unique accounting lines analyzed: "
        f"{len(activity)}"
    )

    exceptions = identify_exceptions(
        activity,
        context,
    )

    print(
        f"GL08 exception lines: "
        f"{len(exceptions)}"
    )

    distinct_accounts = (
        exceptions["_ACCOUNT"].nunique()
        if not exceptions.empty
        else 0
    )

    print(
        "GL08 distinct accounts reported: "
        f"{distinct_accounts}"
    )

    output_dataframe = build_output(
        exceptions,
        master_dataframe,
    )

    output_dataframe = add_reporting_currency(
        output_dataframe,
        fx_dataframe,
    )

    missing_reporting = int(
        output_dataframe[
            "Amount in Reporting Currency"
        ].isna().sum()
    )

    print(
        "GL08 rows without reporting-currency "
        f"conversion: {missing_reporting}"
    )

    if not output_dataframe.empty:
        print("GL08 rows by reason:")

        reason_counts = output_dataframe[
            "GL08 Reason"
        ].value_counts()

        for reason, count in reason_counts.items():
            print(
                f"- {reason}: {count}"
            )

    output_file = write_output(
        output_dataframe,
        context,
    )

    print(
        f"GL08 rows written: "
        f"{len(output_dataframe)}"
    )

    print(
        f"GL output workbook: "
        f"{output_file}"
    )

    print()

    return output_dataframe
