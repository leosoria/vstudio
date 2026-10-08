# VM_017 objective: identify valid vendors whose eligible accounting documents
# were posted by exactly one nonblank Posting User during the configured period.
# The control counts unique SAP accounting documents, not independent economic
# events. The detail sheet exposes document types, amounts, currencies and
# clearing references so Audit can identify payments, fiscal compensations,
# offsetting pairs and other related postings.
# Audit may revise the minimum document threshold or eligible document scope
# in the future, but this implementation does not exclude document types,
# technical vendors, tax vendors or offsetting pairs without approval.
"""
VM_017 - Vendors with a single accounting-document posting user.

This LBR/SAP ECC control identifies valid vendors whose eligible accounting
documents were posted by exactly one nonblank Posting User during the
configured period.

Accounting-document lines are reduced to unique documents before counting.
Posting User is obtained exclusively from the accounting-document header
population by Company, Fiscal Year and Accounting Document.
"""

import re
from time import perf_counter
from typing import Any

import pandas as pd

from core.vm_common import (
    build_vendor_master_population,
    get_valid_vendor_population,
    get_vm_period,
    load_vm_posting_headers,
    load_vm_vendor_postings,
    load_vm_vendors,
    normalize_company,
    normalize_document_number,
    normalize_identifier,
    normalize_vendor_code,
    safe_text,
    write_vm_control_sheet,
)


CONTROL_ID = "VM_017"
SHEET_NAME = "VM17"
DETAIL_SHEET_NAME = "VM17 Detail"
DEFAULT_MINIMUM_DOCUMENTS = 2

VENDOR_KEY_COLUMNS = [
    "Company",
    "Vendor Code",
]

DOCUMENT_KEY_COLUMNS = [
    "Company",
    "Fiscal Year",
    "Accounting Document",
]

VENDOR_REQUIRED_COLUMNS = [
    "Company",
    "Company Name",
    "Vendor Code",
    "Vendor Name",
]

POSTING_REQUIRED_COLUMNS = [
    "Company",
    "Vendor Code",
    "Fiscal Year",
    "Accounting Document",
    "Accounting Document Line",
    "Posting Date",
    "Document Type",
    "Debit/Credit Indicator",
    "Document Currency",
    "Amount in Document Currency",
    "Amount in Local Currency",
    "Clearing Document",
    "Clearing Date",
    "Posting Source",
]

POSTING_HEADER_REQUIRED_COLUMNS = [
    "Company",
    "Fiscal Year",
    "Accounting Document",
    "Posting User",
]

OUTPUT_COLUMNS = [
    "Company",
    "CoCo",
    "Vendor Code",
    "Vendor Name",
    "Posting User",
    "Posting Document Count",
    "First Posting Date",
    "Last Posting Date",
]


DETAIL_OUTPUT_COLUMNS = [
    "Company",
    "CoCo",
    "Vendor Code",
    "Vendor Name",
    "Posting Source",
    "Fiscal Year",
    "Accounting Document",
    "Accounting Document Line",
    "Document Type",
    "Posting Date",
    "Posting User",
    "Debit/Credit Indicator",
    "Document Currency",
    "Amount in Document Currency",
    "Amount in Local Currency",
    "Clearing Document",
    "Clearing Date",
    "Possible Offsetting Pair",
]


def _print_timing(
    stage_name: str,
    started: float,
) -> float:
    finished = perf_counter()

    print(
        f"{CONTROL_ID} {stage_name}: "
        f"{finished - started:.2f} seconds"
    )

    return finished


def _require_columns(
    dataframe: pd.DataFrame,
    required_columns: list[str],
    population_name: str,
) -> None:
    if not isinstance(dataframe, pd.DataFrame):
        raise TypeError(
            f"{CONTROL_ID}: {population_name} must be a pandas DataFrame."
        )

    missing_columns = sorted(
        set(required_columns).difference(
            dataframe.columns
        )
    )

    if missing_columns:
        raise ValueError(
            f"{CONTROL_ID}: {population_name} is missing columns: "
            f"{missing_columns}."
        )


def _configured_companies(
    context: dict[str, Any],
) -> set[str]:
    module_config = context.get("module")

    if not isinstance(module_config, dict):
        raise ValueError(
            "VM_017 requires context['module'] configuration."
        )

    raw_companies = module_config.get(
        "companies",
        "",
    )

    if raw_companies is None:
        return set()

    if isinstance(
        raw_companies,
        (
            list,
            tuple,
            set,
        ),
    ):
        raw_values = list(raw_companies)
    else:
        text = safe_text(raw_companies)

        if (
            text == ""
            or text.upper() in {
                "ALL",
                "*",
                "TODAS",
                "TODOS",
            }
        ):
            return set()

        raw_values = (
            text.replace(";", ",")
            .replace("|", ",")
            .split(",")
        )

    normalized_values = {
        normalize_company(value)
        for value in raw_values
        if safe_text(value) != ""
    }

    if normalized_values.intersection(
        {
            "ALL",
            "*",
            "TODAS",
            "TODOS",
        }
    ):
        return set()

    return {
        value
        for value in normalized_values
        if value != ""
    }


def _filter_configured_companies(
    vendor_master: pd.DataFrame,
    context: dict[str, Any],
) -> tuple[pd.DataFrame, int]:
    companies = _configured_companies(context)

    if not companies:
        return vendor_master.copy(), 0

    included = (
        vendor_master["Company"]
        .map(normalize_company)
        .isin(companies)
    )

    excluded_rows = int(
        (~included).sum()
    )

    return (
        vendor_master.loc[included]
        .copy()
        .reset_index(drop=True),
        excluded_rows,
    )


def _control_parameter(
    context: dict[str, Any],
    name: str,
) -> Any:
    control_config = context.get("control")

    if isinstance(control_config, dict):
        value = control_config.get(name, "")

        if safe_text(value) != "":
            return value

    module_config = context.get("module")

    if isinstance(module_config, dict):
        return module_config.get(name, "")

    return ""


def _minimum_document_count(
    context: dict[str, Any],
) -> int:
    raw_value = _control_parameter(
        context,
        "param1",
    )
    text = safe_text(raw_value)

    if text == "":
        return DEFAULT_MINIMUM_DOCUMENTS

    if not re.fullmatch(
        r"\d+(?:\.0+)?",
        text,
    ):
        raise ValueError(
            f"{CONTROL_ID}: PARAM1 must be an integer greater than or "
            f"equal to 2; received {raw_value!r}."
        )

    value = int(float(text))

    if value < 2:
        raise ValueError(
            f"{CONTROL_ID}: PARAM1 must be greater than or equal to 2; "
            f"received {raw_value!r}."
        )

    return value


def _parameter_warnings(
    context: dict[str, Any],
) -> list[str]:
    warnings = []
    param2 = _control_parameter(
        context,
        "param2",
    )

    if safe_text(param2) != "":
        warnings.append(
            "VM_017 PARAM2 is reserved and was not applied."
        )

    return warnings


def _parse_posting_dates(
    dataframe: pd.DataFrame,
    population_name: str,
) -> pd.Series:
    text = (
        dataframe["Posting Date"]
        .astype("string")
        .fillna("")
        .str.strip()
    )

    parsed = pd.to_datetime(
        text,
        format="%Y-%m-%d",
        errors="coerce",
    )

    blank = text.eq("")

    if blank.any():
        examples = (
            dataframe.loc[
                blank,
                [
                    column
                    for column in DOCUMENT_KEY_COLUMNS
                    if column in dataframe.columns
                ],
            ]
            .head(20)
            .to_dict("records")
        )

        raise ValueError(
            f"{CONTROL_ID}: {population_name} contains blank "
            f"Posting Date values. Examples: {examples}"
        )

    invalid = (
        text.ne("")
        & parsed.isna()
    )

    if invalid.any():
        examples = (
            text.loc[invalid]
            .drop_duplicates()
            .head(20)
            .tolist()
        )

        raise ValueError(
            f"{CONTROL_ID}: {population_name} contains invalid "
            f"Posting Date values. Examples: {examples}"
        )

    return pd.Series(
        parsed,
        index=dataframe.index,
        dtype="datetime64[ns]",
    )


def _filter_postings_to_period(
    posting_population: pd.DataFrame,
    context: dict[str, Any],
) -> tuple[pd.DataFrame, dict[str, int]]:
    _require_columns(
        posting_population,
        POSTING_REQUIRED_COLUMNS,
        "posting population",
    )

    data = posting_population.copy()
    posting_dates = _parse_posting_dates(
        data,
        "posting population",
    )

    date_from, date_to = get_vm_period(context)

    in_period = (
        posting_dates.ge(date_from)
        & posting_dates.le(date_to)
    )

    metrics = {
        "posting_rows_before_period_filter": len(data),
        "posting_rows_outside_period": int(
            (~in_period).sum()
        ),
        "posting_rows_in_period": int(
            in_period.sum()
        ),
    }

    data = data.loc[in_period].copy()
    data["Posting Date"] = posting_dates.loc[
        in_period
    ]

    return data.reset_index(drop=True), metrics


def _parse_optional_dates(
    series: pd.Series,
    population_name: str,
    column_name: str,
) -> pd.Series:
    text = (
        series.astype("string")
        .fillna("")
        .str.strip()
    )

    parsed = pd.to_datetime(
        text,
        format="%Y-%m-%d",
        errors="coerce",
    )

    invalid = text.ne("") & parsed.isna()

    if invalid.any():
        examples = (
            text.loc[invalid]
            .drop_duplicates()
            .head(20)
            .tolist()
        )

        raise ValueError(
            f"{CONTROL_ID}: {population_name} contains invalid "
            f"{column_name} values. Examples: {examples}"
        )

    return pd.Series(
        parsed,
        index=series.index,
        dtype="datetime64[ns]",
    )


def _parse_amounts(
    series: pd.Series,
    population_name: str,
    column_name: str,
) -> pd.Series:
    text = (
        series.astype("string")
        .fillna("")
        .str.strip()
    )

    numeric = pd.to_numeric(
        text,
        errors="coerce",
    )

    unresolved = text.ne("") & numeric.isna()

    if unresolved.any():
        localized = (
            text.loc[unresolved]
            .str.replace(" ", "", regex=False)
        )

        both_separators = (
            localized.str.contains(".", regex=False)
            & localized.str.contains(",", regex=False)
        )

        localized.loc[both_separators] = (
            localized.loc[both_separators]
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
        )

        comma_only = (
            ~both_separators
            & localized.str.contains(",", regex=False)
        )

        localized.loc[comma_only] = (
            localized.loc[comma_only]
            .str.replace(",", ".", regex=False)
        )

        numeric.loc[unresolved] = pd.to_numeric(
            localized,
            errors="coerce",
        )

    invalid = text.ne("") & numeric.isna()

    if invalid.any():
        examples = (
            text.loc[invalid]
            .drop_duplicates()
            .head(20)
            .tolist()
        )

        raise ValueError(
            f"{CONTROL_ID}: {population_name} contains nonconvertible "
            f"{column_name} values. Examples: {examples}"
        )

    return numeric


def _possible_offsetting_pair_row_ids(
    detail: pd.DataFrame,
) -> set[int]:
    if detail.empty:
        return set()

    candidates = detail.loc[
        detail["Document Currency"].ne("")
        & detail["Clearing Document"].ne("")
        & detail["Clearing Date"].notna()
        & detail["Amount in Document Currency"].notna()
        & detail["Amount in Document Currency"].ne(0)
    ].copy()

    if candidates.empty:
        return set()

    candidates["_Detail Row ID"] = candidates.index
    candidates["_Absolute Document Amount"] = candidates[
        "Amount in Document Currency"
    ].abs()

    pair_key = [
        "Company",
        "Vendor Code",
        "Document Currency",
        "Clearing Document",
        "Clearing Date",
        "_Absolute Document Amount",
    ]

    positive = candidates.loc[
        candidates["Amount in Document Currency"].gt(0),
        [
            "_Detail Row ID",
            "Accounting Document",
            *pair_key,
        ],
    ].rename(
        columns={
            "_Detail Row ID": "_Positive Row ID",
            "Accounting Document": "_Positive Document",
        }
    )

    negative = candidates.loc[
        candidates["Amount in Document Currency"].lt(0),
        [
            "_Detail Row ID",
            "Accounting Document",
            *pair_key,
        ],
    ].rename(
        columns={
            "_Detail Row ID": "_Negative Row ID",
            "Accounting Document": "_Negative Document",
        }
    )

    if positive.empty or negative.empty:
        return set()

    # The candidate join is intentionally many-to-many because every positive
    # and negative line sharing the explicit clearing key must be evaluated.
    pairs = positive.merge(
        negative,
        how="inner",
        on=pair_key,
    )

    pairs = pairs.loc[
        pairs["_Positive Document"].ne(
            pairs["_Negative Document"]
        )
    ]

    if pairs.empty:
        return set()

    return set(
        pairs["_Positive Row ID"].tolist()
        + pairs["_Negative Row ID"].tolist()
    )


def _build_vm_017_detail(
    postings: pd.DataFrame,
    headers: pd.DataFrame,
    vendors: pd.DataFrame,
    exceptions: pd.DataFrame,
) -> pd.DataFrame:
    if exceptions.empty:
        return pd.DataFrame(
            columns=DETAIL_OUTPUT_COLUMNS
        )

    exception_keys = (
        exceptions[
            [
                "CoCo",
                "Vendor Code",
            ]
        ]
        .rename(
            columns={
                "CoCo": "Company",
            }
        )
        .copy()
    )

    exception_keys["Company"] = exception_keys[
        "Company"
    ].map(normalize_company)

    exception_keys["Vendor Code"] = exception_keys[
        "Vendor Code"
    ].map(normalize_vendor_code)

    duplicate_exceptions = exception_keys.duplicated(
        subset=VENDOR_KEY_COLUMNS,
        keep=False,
    )

    if duplicate_exceptions.any():
        examples = (
            exception_keys.loc[
                duplicate_exceptions,
                VENDOR_KEY_COLUMNS,
            ]
            .head(20)
            .to_dict("records")
        )

        raise ValueError(
            f"{CONTROL_ID}: exception population is not unique by "
            f"Company/Vendor Code. Examples: {examples}"
        )

    detail = postings.merge(
        exception_keys,
        how="inner",
        on=VENDOR_KEY_COLUMNS,
        validate="many_to_one",
    )

    detail = detail.merge(
        headers[
            [
                *DOCUMENT_KEY_COLUMNS,
                "Posting User",
            ]
        ],
        how="left",
        on=DOCUMENT_KEY_COLUMNS,
        validate="many_to_one",
        indicator=True,
    )

    missing_headers = detail["_merge"].ne("both")

    if missing_headers.any():
        examples = (
            detail.loc[
                missing_headers,
                DOCUMENT_KEY_COLUMNS,
            ]
            .drop_duplicates()
            .head(20)
            .to_dict("records")
        )

        raise ValueError(
            f"{CONTROL_ID}: detail population contains documents "
            f"without posting headers. Examples: {examples}"
        )

    detail = detail.drop(
        columns="_merge"
    )

    blank_user = detail["Posting User"].map(
        safe_text
    ).eq("")

    if blank_user.any():
        examples = (
            detail.loc[
                blank_user,
                DOCUMENT_KEY_COLUMNS,
            ]
            .drop_duplicates()
            .head(20)
            .to_dict("records")
        )

        raise ValueError(
            f"{CONTROL_ID}: detail population contains blank "
            f"Posting User values. Examples: {examples}"
        )

    detail = detail.merge(
        vendors[
            [
                *VENDOR_KEY_COLUMNS,
                "Company Name",
                "Vendor Name",
            ]
        ],
        how="left",
        on=VENDOR_KEY_COLUMNS,
        validate="many_to_one",
    )

    detail = detail.reset_index(drop=True)

    pair_row_ids = _possible_offsetting_pair_row_ids(
        detail
    )

    detail["Possible Offsetting Pair"] = ""

    detail.loc[
        detail.index.isin(pair_row_ids),
        "Possible Offsetting Pair",
    ] = "YES"

    output = pd.DataFrame(
        {
            "Company": detail[
                "Company Name"
            ].map(safe_text),
            "CoCo": detail[
                "Company"
            ].map(normalize_company),
            "Vendor Code": detail[
                "Vendor Code"
            ].map(normalize_vendor_code),
            "Vendor Name": detail[
                "Vendor Name"
            ].map(safe_text),
            "Posting Source": detail[
                "Posting Source"
            ].map(safe_text),
            "Fiscal Year": detail[
                "Fiscal Year"
            ].map(normalize_identifier),
            "Accounting Document": detail[
                "Accounting Document"
            ].map(normalize_document_number),
            "Accounting Document Line": detail[
                "Accounting Document Line"
            ].map(normalize_identifier),
            "Document Type": detail[
                "Document Type"
            ].map(safe_text),
            "Posting Date": detail[
                "Posting Date"
            ],
            "Posting User": detail[
                "Posting User"
            ].map(safe_text),
            "Debit/Credit Indicator": detail[
                "Debit/Credit Indicator"
            ].map(safe_text),
            "Document Currency": detail[
                "Document Currency"
            ].map(safe_text),
            "Amount in Document Currency": detail[
                "Amount in Document Currency"
            ],
            "Amount in Local Currency": detail[
                "Amount in Local Currency"
            ],
            "Clearing Document": detail[
                "Clearing Document"
            ].map(normalize_document_number),
            "Clearing Date": detail[
                "Clearing Date"
            ],
            "Possible Offsetting Pair": detail[
                "Possible Offsetting Pair"
            ],
        },
        columns=DETAIL_OUTPUT_COLUMNS,
    )

    return (
        output.sort_values(
            [
                "Company",
                "CoCo",
                "Vendor Name",
                "Vendor Code",
                "Posting Date",
                "Fiscal Year",
                "Accounting Document",
                "Accounting Document Line",
                "Posting Source",
            ],
            kind="mergesort",
        )
        .loc[
            :,
            DETAIL_OUTPUT_COLUMNS,
        ]
        .reset_index(drop=True)
    )

def build_vm_017(
    vendor_population: pd.DataFrame,
    posting_population: pd.DataFrame,
    posting_header_population: pd.DataFrame,
    minimum_document_count: int = DEFAULT_MINIMUM_DOCUMENTS,
) -> pd.DataFrame:
    _require_columns(
        vendor_population,
        VENDOR_REQUIRED_COLUMNS,
        "vendor population",
    )
    _require_columns(
        posting_population,
        POSTING_REQUIRED_COLUMNS,
        "posting population",
    )
    _require_columns(
        posting_header_population,
        POSTING_HEADER_REQUIRED_COLUMNS,
        "posting header population",
    )

    if (
        isinstance(minimum_document_count, bool)
        or not isinstance(minimum_document_count, int)
        or minimum_document_count < 2
    ):
        raise ValueError(
            f"{CONTROL_ID}: minimum_document_count must be an integer "
            "greater than or equal to 2."
        )

    vendors = vendor_population.loc[
        :,
        VENDOR_REQUIRED_COLUMNS,
    ].copy()

    vendors["Company"] = vendors[
        "Company"
    ].map(normalize_company)

    vendors["Vendor Code"] = vendors[
        "Vendor Code"
    ].map(normalize_vendor_code)

    vendors["Company Name"] = vendors[
        "Company Name"
    ].map(safe_text)

    vendors["Vendor Name"] = vendors[
        "Vendor Name"
    ].map(safe_text)

    invalid_vendor_key = (
        vendors["Company"].eq("")
        | vendors["Vendor Code"].eq("")
    )

    if invalid_vendor_key.any():
        examples = (
            vendors.loc[
                invalid_vendor_key,
                VENDOR_KEY_COLUMNS,
            ]
            .head(20)
            .to_dict("records")
        )

        raise ValueError(
            f"{CONTROL_ID}: vendor population contains blank normalized "
            f"Company/Vendor Code keys. Examples: {examples}"
        )

    duplicate_vendors = vendors.duplicated(
        subset=VENDOR_KEY_COLUMNS,
        keep=False,
    )

    if duplicate_vendors.any():
        examples = (
            vendors.loc[
                duplicate_vendors,
                VENDOR_KEY_COLUMNS,
            ]
            .drop_duplicates()
            .head(20)
            .to_dict("records")
        )

        raise ValueError(
            f"{CONTROL_ID}: vendor population is not unique by "
            f"Company/Vendor Code. Examples: {examples}"
        )

    postings = posting_population.loc[
        :,
        POSTING_REQUIRED_COLUMNS,
    ].copy()

    postings["Company"] = postings[
        "Company"
    ].map(normalize_company)

    postings["Vendor Code"] = postings[
        "Vendor Code"
    ].map(normalize_vendor_code)

    postings["Fiscal Year"] = postings[
        "Fiscal Year"
    ].map(normalize_identifier)

    postings["Accounting Document"] = postings[
        "Accounting Document"
    ].map(normalize_document_number)

    postings["Accounting Document Line"] = (
        postings["Accounting Document Line"]
        .map(normalize_identifier)
    )

    postings["Posting Date"] = _parse_posting_dates(
        postings,
        "posting population",
    )

    invalid_posting_key = (
        postings[
            [
                "Company",
                "Vendor Code",
                "Fiscal Year",
                "Accounting Document",
                "Accounting Document Line",
            ]
        ]
        .eq("")
        .any(axis=1)
    )

    if invalid_posting_key.any():
        examples = (
            postings.loc[
                invalid_posting_key,
                POSTING_REQUIRED_COLUMNS,
            ]
            .head(20)
            .to_dict("records")
        )

        raise ValueError(
            f"{CONTROL_ID}: posting population contains blank normalized "
            f"mandatory keys. Examples: {examples}"
        )

    postings = postings.merge(
        vendors[VENDOR_KEY_COLUMNS],
        how="inner",
        on=VENDOR_KEY_COLUMNS,
        validate="many_to_one",
    )

    document_vendor_key = [
        "Company",
        "Vendor Code",
        "Fiscal Year",
        "Accounting Document",
    ]

    posting_date_counts = (
        postings.groupby(
            document_vendor_key,
            sort=False,
            observed=True,
            dropna=False,
        )["Posting Date"]
        .nunique()
    )

    conflicting_posting_dates = posting_date_counts.loc[
        posting_date_counts.gt(1)
    ]

    if not conflicting_posting_dates.empty:
        examples = (
            conflicting_posting_dates
            .head(20)
            .reset_index()
            .loc[
                :,
                document_vendor_key,
            ]
            .to_dict("records")
        )

        raise ValueError(
            f"{CONTROL_ID}: posting population contains conflicting "
            "Posting Date values for the same vendor document. "
            f"Examples: {examples}"
        )

    documents = (
        postings.sort_values(
            [
                *document_vendor_key,
                "Accounting Document Line",
            ],
            kind="mergesort",
        )
        .drop_duplicates(
            subset=document_vendor_key,
            keep="first",
        )
        .loc[
            :,
            [
                *document_vendor_key,
                "Posting Date",
            ],
        ]
        .reset_index(drop=True)
    )

    headers = posting_header_population.loc[
        :,
        POSTING_HEADER_REQUIRED_COLUMNS,
    ].copy()

    headers["Company"] = headers[
        "Company"
    ].map(normalize_company)

    headers["Fiscal Year"] = headers[
        "Fiscal Year"
    ].map(normalize_identifier)

    headers["Accounting Document"] = headers[
        "Accounting Document"
    ].map(normalize_document_number)

    headers["Posting User"] = headers[
        "Posting User"
    ].map(safe_text)

    invalid_header_key = (
        headers[DOCUMENT_KEY_COLUMNS]
        .eq("")
        .any(axis=1)
    )

    if invalid_header_key.any():
        examples = (
            headers.loc[
                invalid_header_key,
                DOCUMENT_KEY_COLUMNS,
            ]
            .head(20)
            .to_dict("records")
        )

        raise ValueError(
            f"{CONTROL_ID}: posting header population contains blank "
            f"normalized document keys. Examples: {examples}"
        )

    nonblank_headers = headers.loc[
        headers["Posting User"].ne("")
    ]

    header_user_counts = (
        nonblank_headers.groupby(
            DOCUMENT_KEY_COLUMNS,
            sort=False,
            observed=True,
            dropna=False,
        )["Posting User"]
        .nunique()
    )

    conflicting_users = header_user_counts.loc[
        header_user_counts.gt(1)
    ]

    if not conflicting_users.empty:
        examples = (
            conflicting_users
            .head(20)
            .reset_index()
            .loc[
                :,
                DOCUMENT_KEY_COLUMNS,
            ]
            .to_dict("records")
        )

        raise ValueError(
            f"{CONTROL_ID}: posting header population contains "
            "conflicting Posting User values for the same document. "
            f"Examples: {examples}"
        )

    headers["_Posting User Blank"] = headers[
        "Posting User"
    ].eq("")

    headers = (
        headers.sort_values(
            [
                *DOCUMENT_KEY_COLUMNS,
                "_Posting User Blank",
                "Posting User",
            ],
            kind="mergesort",
        )
        .drop_duplicates(
            subset=DOCUMENT_KEY_COLUMNS,
            keep="first",
        )
        .drop(
            columns="_Posting User Blank"
        )
        .reset_index(drop=True)
    )

    attributed = documents.merge(
        headers[
            [
                *DOCUMENT_KEY_COLUMNS,
                "Posting User",
            ]
        ],
        how="left",
        on=DOCUMENT_KEY_COLUMNS,
        validate="many_to_one",
        indicator=True,
    )

    missing_headers = attributed[
        "_merge"
    ].ne("both")

    if missing_headers.any():
        examples = (
            attributed.loc[
                missing_headers,
                DOCUMENT_KEY_COLUMNS,
            ]
            .head(20)
            .to_dict("records")
        )

        raise ValueError(
            f"{CONTROL_ID}: {int(missing_headers.sum())} unique posting "
            "documents have no matching posting header. "
            f"Examples: {examples}"
        )

    attributed = attributed.drop(
        columns="_merge"
    )

    blank_posting_user = attributed[
        "Posting User"
    ].map(safe_text).eq("")

    if blank_posting_user.any():
        examples = (
            attributed.loc[
                blank_posting_user,
                DOCUMENT_KEY_COLUMNS,
            ]
            .head(20)
            .to_dict("records")
        )

        raise ValueError(
            f"{CONTROL_ID}: {int(blank_posting_user.sum())} unique "
            "posting documents have blank Posting User. "
            f"Examples: {examples}"
        )

    if attributed.empty:
        output = pd.DataFrame(
            columns=OUTPUT_COLUMNS
        )
        output.attrs["metrics"] = {
            "unique_documents_before_header_join": 0,
            "documents_with_header": 0,
            "documents_without_header": 0,
            "documents_with_posting_user": 0,
            "documents_with_blank_posting_user": 0,
            "valid_vendors_with_documents": 0,
            "vendors_with_single_posting_user": 0,
            "vendors_with_multiple_posting_users": 0,
            "vendors_meeting_document_minimum": 0,
            "exception_rows": 0,
        }
        return output

    vendor_summary = (
        attributed.groupby(
            VENDOR_KEY_COLUMNS,
            as_index=False,
            sort=False,
            observed=True,
            dropna=False,
        )
        .agg(
            Posting_User_Count=(
                "Posting User",
                "nunique",
            ),
            Posting_User=(
                "Posting User",
                "first",
            ),
            Posting_Document_Count=(
                "Accounting Document",
                "size",
            ),
            First_Posting_Date=(
                "Posting Date",
                "min",
            ),
            Last_Posting_Date=(
                "Posting Date",
                "max",
            ),
        )
    )

    single_user = vendor_summary[
        "Posting_User_Count"
    ].eq(1)

    multiple_users = vendor_summary[
        "Posting_User_Count"
    ].gt(1)

    meets_minimum = vendor_summary[
        "Posting_Document_Count"
    ].ge(minimum_document_count)

    exceptions = vendor_summary.loc[
        single_user & meets_minimum
    ].copy()

    exceptions = exceptions.merge(
        vendors,
        how="left",
        on=VENDOR_KEY_COLUMNS,
        validate="one_to_one",
    )

    output = pd.DataFrame(
        {
            "Company": exceptions[
                "Company Name"
            ].map(safe_text),
            "CoCo": exceptions[
                "Company"
            ].map(normalize_company),
            "Vendor Code": exceptions[
                "Vendor Code"
            ].map(normalize_vendor_code),
            "Vendor Name": exceptions[
                "Vendor Name"
            ].map(safe_text),
            "Posting User": exceptions[
                "Posting_User"
            ].map(safe_text),
            "Posting Document Count": exceptions[
                "Posting_Document_Count"
            ].astype("int64"),
            "First Posting Date": exceptions[
                "First_Posting_Date"
            ],
            "Last Posting Date": exceptions[
                "Last_Posting_Date"
            ],
        },
        columns=OUTPUT_COLUMNS,
    )

    output = (
        output.sort_values(
            [
                "Company",
                "CoCo",
                "Posting User",
                "Vendor Name",
                "Vendor Code",
            ],
            kind="mergesort",
        )
        .drop_duplicates(
            subset=[
                "CoCo",
                "Vendor Code",
            ],
            keep="first",
        )
        .loc[
            :,
            OUTPUT_COLUMNS,
        ]
        .reset_index(drop=True)
    )

    output.attrs["metrics"] = {
        "unique_documents_before_header_join": len(
            documents
        ),
        "documents_with_header": len(
            attributed
        ),
        "documents_without_header": 0,
        "documents_with_posting_user": len(
            attributed
        ),
        "documents_with_blank_posting_user": 0,
        "valid_vendors_with_documents": len(
            vendor_summary
        ),
        "vendors_with_single_posting_user": int(
            single_user.sum()
        ),
        "vendors_with_multiple_posting_users": int(
            multiple_users.sum()
        ),
        "vendors_meeting_document_minimum": int(
            meets_minimum.sum()
        ),
        "exception_rows": len(output),
    }

    return output


def build_vm_017_detail(
    vendor_population: pd.DataFrame,
    posting_population: pd.DataFrame,
    posting_header_population: pd.DataFrame,
    exceptions: pd.DataFrame,
) -> pd.DataFrame:
    _require_columns(
        vendor_population,
        VENDOR_REQUIRED_COLUMNS,
        "vendor population",
    )
    _require_columns(
        posting_population,
        POSTING_REQUIRED_COLUMNS,
        "posting population",
    )
    _require_columns(
        posting_header_population,
        POSTING_HEADER_REQUIRED_COLUMNS,
        "posting header population",
    )
    _require_columns(
        exceptions,
        OUTPUT_COLUMNS,
        "exception summary",
    )

    vendors = vendor_population.loc[
        :,
        VENDOR_REQUIRED_COLUMNS,
    ].copy()

    vendors["Company"] = vendors[
        "Company"
    ].map(normalize_company)

    vendors["Vendor Code"] = vendors[
        "Vendor Code"
    ].map(normalize_vendor_code)

    vendors["Company Name"] = vendors[
        "Company Name"
    ].map(safe_text)

    vendors["Vendor Name"] = vendors[
        "Vendor Name"
    ].map(safe_text)

    invalid_vendor_key = (
        vendors["Company"].eq("")
        | vendors["Vendor Code"].eq("")
    )

    if invalid_vendor_key.any():
        examples = (
            vendors.loc[
                invalid_vendor_key,
                VENDOR_KEY_COLUMNS,
            ]
            .head(20)
            .to_dict("records")
        )

        raise ValueError(
            f"{CONTROL_ID}: detail vendor population contains blank "
            f"Company/Vendor Code keys. Examples: {examples}"
        )

    duplicate_vendors = vendors.duplicated(
        subset=VENDOR_KEY_COLUMNS,
        keep=False,
    )

    if duplicate_vendors.any():
        examples = (
            vendors.loc[
                duplicate_vendors,
                VENDOR_KEY_COLUMNS,
            ]
            .drop_duplicates()
            .head(20)
            .to_dict("records")
        )

        raise ValueError(
            f"{CONTROL_ID}: detail vendor population is not unique by "
            f"Company/Vendor Code. Examples: {examples}"
        )

    postings = posting_population.loc[
        :,
        POSTING_REQUIRED_COLUMNS,
    ].copy()

    postings["Company"] = postings[
        "Company"
    ].map(normalize_company)

    postings["Vendor Code"] = postings[
        "Vendor Code"
    ].map(normalize_vendor_code)

    postings["Fiscal Year"] = postings[
        "Fiscal Year"
    ].map(normalize_identifier)

    postings["Accounting Document"] = postings[
        "Accounting Document"
    ].map(normalize_document_number)

    postings["Accounting Document Line"] = postings[
        "Accounting Document Line"
    ].map(normalize_identifier)

    postings["Document Type"] = postings[
        "Document Type"
    ].map(safe_text)

    postings["Debit/Credit Indicator"] = postings[
        "Debit/Credit Indicator"
    ].map(safe_text)

    postings["Document Currency"] = postings[
        "Document Currency"
    ].map(safe_text)

    postings["Clearing Document"] = postings[
        "Clearing Document"
    ].map(normalize_document_number)

    postings["Posting Source"] = postings[
        "Posting Source"
    ].map(safe_text)

    postings["Posting Date"] = _parse_posting_dates(
        postings,
        "detail posting population",
    )

    postings["Clearing Date"] = _parse_optional_dates(
        postings["Clearing Date"],
        "detail posting population",
        "Clearing Date",
    )

    postings["Amount in Document Currency"] = _parse_amounts(
        postings["Amount in Document Currency"],
        "detail posting population",
        "Amount in Document Currency",
    )

    postings["Amount in Local Currency"] = _parse_amounts(
        postings["Amount in Local Currency"],
        "detail posting population",
        "Amount in Local Currency",
    )

    detail_grain_columns = [
        "Company",
        "Vendor Code",
        "Fiscal Year",
        "Accounting Document",
        "Accounting Document Line",
        "Posting Source",
    ]

    invalid_posting_key = (
        postings[
            detail_grain_columns
        ]
        .eq("")
        .any(axis=1)
    )

    if invalid_posting_key.any():
        examples = (
            postings.loc[
                invalid_posting_key,
                detail_grain_columns,
            ]
            .head(20)
            .to_dict("records")
        )

        raise ValueError(
            f"{CONTROL_ID}: detail posting population contains blank "
            f"keys. Examples: {examples}"
        )

    duplicate_detail_grain = postings.duplicated(
        subset=detail_grain_columns,
        keep=False,
    )

    if duplicate_detail_grain.any():
        duplicated = postings.loc[
            duplicate_detail_grain
        ].copy()

        value_columns = [
            column
            for column in POSTING_REQUIRED_COLUMNS
            if column not in detail_grain_columns
        ]

        conflicts = (
            duplicated.groupby(
                detail_grain_columns,
                sort=False,
                observed=True,
                dropna=False,
            )[value_columns]
            .nunique(dropna=False)
            .gt(1)
            .any(axis=1)
        )

        if conflicts.any():
            examples = (
                conflicts.loc[
                    conflicts
                ]
                .head(20)
                .reset_index()
                .loc[
                    :,
                    detail_grain_columns,
                ]
                .to_dict("records")
            )

            raise ValueError(
                f"{CONTROL_ID}: detail posting population contains "
                f"conflicting rows. Examples: {examples}"
            )

        postings = postings.drop_duplicates(
            subset=detail_grain_columns,
            keep="first",
        )

    headers = posting_header_population.loc[
        :,
        POSTING_HEADER_REQUIRED_COLUMNS,
    ].copy()

    headers["Company"] = headers[
        "Company"
    ].map(normalize_company)

    headers["Fiscal Year"] = headers[
        "Fiscal Year"
    ].map(normalize_identifier)

    headers["Accounting Document"] = headers[
        "Accounting Document"
    ].map(normalize_document_number)

    headers["Posting User"] = headers[
        "Posting User"
    ].map(safe_text)

    invalid_header_key = (
        headers[
            DOCUMENT_KEY_COLUMNS
        ]
        .eq("")
        .any(axis=1)
    )

    if invalid_header_key.any():
        examples = (
            headers.loc[
                invalid_header_key,
                DOCUMENT_KEY_COLUMNS,
            ]
            .head(20)
            .to_dict("records")
        )

        raise ValueError(
            f"{CONTROL_ID}: detail posting header population contains "
            f"blank document keys. Examples: {examples}"
        )

    nonblank_headers = headers.loc[
        headers["Posting User"].ne("")
    ]

    user_counts = (
        nonblank_headers.groupby(
            DOCUMENT_KEY_COLUMNS,
            sort=False,
            observed=True,
            dropna=False,
        )["Posting User"]
        .nunique()
    )

    conflicting_users = user_counts.loc[
        user_counts.gt(1)
    ]

    if not conflicting_users.empty:
        examples = (
            conflicting_users
            .head(20)
            .reset_index()
            .loc[
                :,
                DOCUMENT_KEY_COLUMNS,
            ]
            .to_dict("records")
        )

        raise ValueError(
            f"{CONTROL_ID}: detail posting header population contains "
            f"conflicting Posting User values. Examples: {examples}"
        )

    headers["_Posting User Blank"] = headers[
        "Posting User"
    ].eq("")

    headers = (
        headers.sort_values(
            [
                *DOCUMENT_KEY_COLUMNS,
                "_Posting User Blank",
                "Posting User",
            ],
            kind="mergesort",
        )
        .drop_duplicates(
            subset=DOCUMENT_KEY_COLUMNS,
            keep="first",
        )
        .drop(
            columns="_Posting User Blank"
        )
        .reset_index(drop=True)
    )

    return _build_vm_017_detail(
        postings=postings,
        headers=headers,
        vendors=vendors,
        exceptions=exceptions,
    )

def run_vm_017(
    context: dict[str, Any],
) -> dict[str, Any]:
    started = perf_counter()

    module_config = context.get("module")

    if not isinstance(module_config, dict):
        raise ValueError(
            "VM_017 requires context['module'] configuration."
        )

    date_from, date_to = get_vm_period(context)
    minimum_documents = _minimum_document_count(
        context
    )

    warnings = _parameter_warnings(
        context
    )

    vendor_source = load_vm_vendors(
        context
    )
    source_rows = len(vendor_source)

    vendor_master = build_vendor_master_population(
        vendor_source
    )
    master_rows = len(vendor_master)

    configured_vendor_master, excluded_company_rows = (
        _filter_configured_companies(
            vendor_master,
            context,
        )
    )

    valid_population_input_rows = len(
        configured_vendor_master
    )

    valid_population, population_metrics = (
        get_valid_vendor_population(
            configured_vendor_master
        )
    )

    if valid_population.empty:
        raise ValueError(
            "VM_017: valid vendor population is empty after CONFIG "
            "company and common VM exclusion rules."
        )

    posting_population, posting_metadata = (
        load_vm_vendor_postings(
            context
        )
    )

    warnings.extend(
        posting_metadata.get(
            "warnings",
            [],
        )
    )

    if (
        posting_population is None
        or not posting_metadata.get(
            "available",
            False,
        )
    ):
        raise ValueError(
            "VM_017: vendor posting population is unavailable. "
            "Both BSIK and BSAK extracts are required."
        )

    posting_population, period_metrics = (
        _filter_postings_to_period(
            posting_population,
            context,
        )
    )

    if posting_population.empty:
        raise ValueError(
            "VM_017: posting sources are available but contain no "
            "documents within CONFIG FROM/TO."
        )

    posting_headers, header_metadata = (
        load_vm_posting_headers(
            context
        )
    )

    warnings.extend(
        header_metadata.get(
            "warnings",
            [],
        )
    )

    if not header_metadata.get(
        "available",
        False,
    ):
        raise ValueError(
            "VM_017: posting header population is unavailable."
        )

    load_finished = _print_timing(
        "input load and population validation",
        started,
    )

    output = build_vm_017(
        vendor_population=valid_population,
        posting_population=posting_population,
        posting_header_population=posting_headers,
        minimum_document_count=minimum_documents,
    )

    analytic_metrics = output.attrs.get(
        "metrics",
        {},
    )

    detail = build_vm_017_detail(
        vendor_population=valid_population,
        posting_population=posting_population,
        posting_header_population=posting_headers,
        exceptions=output,
    )
    analytic_finished = _print_timing(
        "preparation and analytic logic",
        load_finished,
    )

    fx_finished = _print_timing(
        "FX conversion (not applicable)",
        analytic_finished,
    )

    output_file = write_vm_control_sheet(
        context=context,
        sheet_name=SHEET_NAME,
        dataframe=output,
        date_columns=[
            "First Posting Date",
            "Last Posting Date",
        ],
        integer_columns=[
            "Posting Document Count",
        ],
    )

    output_file = write_vm_control_sheet(
        context=context,
        sheet_name=DETAIL_SHEET_NAME,
        dataframe=detail,
        date_columns=[
            "Posting Date",
            "Clearing Date",
        ],
        amount_columns=[
            "Amount in Document Currency",
            "Amount in Local Currency",
        ],
    )
    _print_timing(
        "workbook write",
        fx_finished,
    )

    valid_rows = int(
        population_metrics.get(
            "output_rows",
            len(valid_population),
        )
    )

    net_common_exclusion = max(
        valid_population_input_rows
        - valid_rows,
        0,
    )

    valid_vendor_keys = valid_population[
        VENDOR_KEY_COLUMNS
    ].drop_duplicates()

    posting_vendor_keys = posting_population[
        VENDOR_KEY_COLUMNS
    ].copy()

    posting_vendor_keys["Company"] = (
        posting_vendor_keys["Company"]
        .map(normalize_company)
    )

    posting_vendor_keys["Vendor Code"] = (
        posting_vendor_keys["Vendor Code"]
        .map(normalize_vendor_code)
    )

    valid_vendors_with_documents = len(
        valid_vendor_keys.merge(
            posting_vendor_keys.drop_duplicates(),
            how="inner",
            on=VENDOR_KEY_COLUMNS,
            validate="one_to_one",
        )
    )

    valid_vendors_without_documents = max(
        len(valid_vendor_keys)
        - valid_vendors_with_documents,
        0,
    )

    print(f"{CONTROL_ID} period FROM: {date_from.date()}")
    print(f"{CONTROL_ID} period TO: {date_to.date()}")
    print(
        f"{CONTROL_ID} functional definition: at least "
        f"{minimum_documents} unique documents and exactly one "
        "nonblank Posting User per Company/Vendor Code."
    )
    print(
        f"{CONTROL_ID} Posting User source: "
        f"{header_metadata.get('source', 'BKPF-USNAM')}"
    )
    print(f"{CONTROL_ID} original vm_vendors rows: {source_rows}")
    print(
        f"{CONTROL_ID} vendor master rows before company filter: "
        f"{master_rows}"
    )
    print(
        f"{CONTROL_ID} rows excluded by CONFIG companies: "
        f"{excluded_company_rows}"
    )
    print(
        f"{CONTROL_ID} rows delivered to valid population rules: "
        f"{valid_population_input_rows}"
    )
    print(f"{CONTROL_ID} valid vendor rows: {valid_rows}")

    # Individual exclusion metrics may overlap and must not be added to
    # reconcile the valid population.
    print(
        f"{CONTROL_ID} central deletion flag matches: "
        f"{population_metrics.get('excluded_central_deletion_flag', 0)}"
    )
    print(
        f"{CONTROL_ID} company deletion flag matches: "
        f"{population_metrics.get('excluded_company_deletion_flag', 0)}"
    )
    print(
        f"{CONTROL_ID} vendor prefix E/T matches: "
        f"{population_metrics.get('excluded_vendor_prefix_e_or_t', 0)}"
    )
    print(
        f"{CONTROL_ID} Account Group ZFUN matches: "
        f"{population_metrics.get('excluded_employee_account_group_zfun', 0)}"
    )
    print(
        f"{CONTROL_ID} configured intercompany Vendor Codes: "
        f"{population_metrics.get('configured_intercompany_vendor_codes', 0)}"
    )
    print(
        f"{CONTROL_ID} intercompany Vendor Code matches: "
        f"{population_metrics.get('excluded_intercompany_vendor_code', 0)}"
    )
    print(
        f"{CONTROL_ID} net common-rule exclusion: "
        f"{net_common_exclusion}"
    )
    print(
        f"{CONTROL_ID} rows with Trading Partner: "
        f"{population_metrics.get('trading_partner_nonblank_rows', 0)}"
    )
    print(
        f"{CONTROL_ID} BSIK/BSAK available: "
        f"{posting_metadata.get('available', False)}"
    )
    print(
        f"{CONTROL_ID} BSIK rows: "
        f"{posting_metadata.get('bsik_rows', 0)}"
    )
    print(
        f"{CONTROL_ID} BSAK rows: "
        f"{posting_metadata.get('bsak_rows', 0)}"
    )
    print(
        f"{CONTROL_ID} combined posting rows: "
        f"{posting_metadata.get('posting_rows', 0)}"
    )
    print(
        f"{CONTROL_ID} posting rows outside period: "
        f"{period_metrics.get('posting_rows_outside_period', 0)}"
    )
    print(
        f"{CONTROL_ID} posting rows within period: "
        f"{period_metrics.get('posting_rows_in_period', 0)}"
    )
    print(
        f"{CONTROL_ID} unique documents before header join: "
        f"{analytic_metrics.get('unique_documents_before_header_join', 0)}"
    )
    print(
        f"{CONTROL_ID} available posting headers: "
        f"{header_metadata.get('header_rows', 0)}"
    )
    print(
        f"{CONTROL_ID} documents with header: "
        f"{analytic_metrics.get('documents_with_header', 0)}"
    )
    print(
        f"{CONTROL_ID} documents without header: "
        f"{analytic_metrics.get('documents_without_header', 0)}"
    )
    print(
        f"{CONTROL_ID} documents with Posting User: "
        f"{analytic_metrics.get('documents_with_posting_user', 0)}"
    )
    print(
        f"{CONTROL_ID} documents with blank Posting User: "
        f"{analytic_metrics.get('documents_with_blank_posting_user', 0)}"
    )
    print(
        f"{CONTROL_ID} Posting User conflicts: "
        f"{header_metadata.get('posting_user_conflicts', 0)}"
    )
    print(
        f"{CONTROL_ID} valid vendors with documents: "
        f"{valid_vendors_with_documents}"
    )
    print(
        f"{CONTROL_ID} valid vendors without documents: "
        f"{valid_vendors_without_documents}"
    )
    print(
        f"{CONTROL_ID} vendors with one distinct Posting User: "
        f"{analytic_metrics.get('vendors_with_single_posting_user', 0)}"
    )
    print(
        f"{CONTROL_ID} vendors with multiple Posting Users: "
        f"{analytic_metrics.get('vendors_with_multiple_posting_users', 0)}"
    )
    print(
        f"{CONTROL_ID} vendors meeting document minimum: "
        f"{analytic_metrics.get('vendors_meeting_document_minimum', 0)}"
    )
    print(f"{CONTROL_ID} final exceptions: {len(output)}")

    detail_unique_documents = (
        detail[
            [
                "CoCo",
                "Vendor Code",
                "Fiscal Year",
                "Accounting Document",
            ]
        ]
        .drop_duplicates()
        .shape[0]
    )

    possible_pair_rows = int(
        detail[
            "Possible Offsetting Pair"
        ].eq("YES").sum()
    )

    possible_pair_vendors = (
        detail.loc[
            detail["Possible Offsetting Pair"].eq("YES"),
            [
                "CoCo",
                "Vendor Code",
            ],
        ]
        .drop_duplicates()
        .shape[0]
    )

    document_currency_values = sorted(
        value
        for value in detail[
            "Document Currency"
        ].map(safe_text).unique()
        if value != ""
    )

    blank_currency_rows = int(
        detail[
            "Document Currency"
        ].map(safe_text).eq("").sum()
    )

    if blank_currency_rows:
        warnings.append(
            f"VM_017 detail contains {blank_currency_rows} row(s) "
            "with blank Document Currency."
        )

    warnings.extend(
        [
            "VM_017 detail cannot display Transaction Code because the "
            "current common posting populations do not expose it.",
            "VM_017 detail cannot display Entry Date or Entry Time because "
            "the current common posting populations do not expose them.",
            "VM_017 optional posting-header fields require approved "
            "canonical aliases in core.vm_common before use.",
        ]
    )

    print(f"{CONTROL_ID} detail sheet: {DETAIL_SHEET_NAME}")
    print(f"{CONTROL_ID} detail rows: {len(detail)}")
    print(
        f"{CONTROL_ID} detail unique documents: "
        f"{detail_unique_documents}"
    )
    print(
        f"{CONTROL_ID} detail BSIK rows: "
        f"{int(detail['Posting Source'].eq('BSIK').sum())}"
    )
    print(
        f"{CONTROL_ID} detail BSAK rows: "
        f"{int(detail['Posting Source'].eq('BSAK').sum())}"
    )
    print(
        f"{CONTROL_ID} possible offsetting pair rows: "
        f"{possible_pair_rows}"
    )
    print(
        f"{CONTROL_ID} exception vendors with possible offsetting pairs: "
        f"{possible_pair_vendors}"
    )
    print(
        f"{CONTROL_ID} document currency values: "
        f"{document_currency_values}"
    )
    print(
        f"{CONTROL_ID} document count definition: "
        "unique SAP accounting documents"
    )
    print(
        f"{CONTROL_ID} offsetting pair indicator is diagnostic only "
        "and does not change exceptions"
    )

    for warning in [
        *population_metrics.get(
            "warnings",
            [],
        ),
        *warnings,
    ]:
        print(f"WARNING: {warning}")

    # The current orchestration layer may display a technical OK whenever no
    # Python exception occurs, even though this dictionary returns ERROR when
    # the control identifies exceptions.
    return {
        "status": (
            "ERROR"
            if not output.empty
            else "OK"
        ),
        "output_file": output_file,
        "sheet_name": SHEET_NAME,
        "rows": len(output),
    }










