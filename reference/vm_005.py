"""
VM_005 - Vendors sharing the same Tax/Business Number.

Objective
---------
Identify different Vendor Codes that share the same normalized
Tax/Business Number within the same company.

Population key:
    Company + Vendor Code

Analytical key:
    Company + normalized Tax/Business Number

Final granularity:
    CoCo + Vendor Code

The input workbook resolution, reading, header mapping, footer handling,
canonicalization, posting logic, last-invoice calculation and workbook
writing are delegated exclusively to core.vm_common.
"""

from time import perf_counter
from typing import Any

import pandas as pd

from core.vm_common import (
    build_vendor_master_population,
    build_vm_last_invoice_population,
    get_valid_vendor_population,
    load_vm_vendor_postings,
    load_vm_vendors,
    normalize_company,
    normalize_tax_id,
    resolve_tax_business_number,
    safe_text,
    write_vm_control_sheet,
)


CONTROL_ID = "VM_005"
SHEET_NAME = "VM05"

TAX_BUSINESS_NUMBER_PRIORITY = (
    "Tax Number 1",
    "Tax Number 2",
    "Tax Number 3",
    "Tax Number 4",
    "Tax Number 5",
    "VAT Registration Number",
)

VENDOR_KEY_COLUMNS = [
    "Company",
    "Vendor Code",
]

LAST_INVOICE_COLUMNS = [
    "Last Invoice Number",
    "Last Transaction Date",
    "Last Inv Amt Doc Currency",
    "Last Inv Amt Doc Currency Indicator",
]

OUTPUT_COLUMNS = [
    "Company",
    "CoCo",
    "Vendor Code",
    "Vendor Name",
    "Tax/Business Number",
    "Last Invoice Number",
    "Last Transaction Date",
    "Last Inv Amt Doc Currency",
    "Last Inv Amt Doc Currency Indicator",
    "Group",
]


def _configured_companies(
    context: dict[str, Any],
) -> set[str]:
    """
    Return normalized company codes configured for VM.

    An empty value, ALL, *, TODAS or TODOS means that every company in the
    validated vendor population is included.

    This is the same local fallback used by the approved LBR VM controls.
    """
    module_config = context.get(
        "module"
    )

    if not isinstance(
        module_config,
        dict,
    ):
        raise ValueError(
            f"{CONTROL_ID} requires context['module'] configuration."
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
        raw_values = list(
            raw_companies
        )
    else:
        text = safe_text(
            raw_companies
        )

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

    return normalized_values


def _filter_configured_companies(
    vendor_master: pd.DataFrame,
    context: dict[str, Any],
) -> tuple[pd.DataFrame, int]:
    """
    Apply the VM CONFIG company population after common normalization.
    """
    companies = _configured_companies(
        context
    )

    if not companies:
        return vendor_master.copy(), 0

    company_values = vendor_master[
        "Company"
    ].map(normalize_company)

    included = company_values.isin(
        companies
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


def _require_columns(
    dataframe: pd.DataFrame,
    required_columns: list[str] | tuple[str, ...] | set[str],
    source_name: str,
) -> None:
    """
    Raise when a DataFrame does not contain all required columns.
    """
    missing_columns = sorted(
        set(required_columns).difference(
            dataframe.columns
        )
    )

    if missing_columns:
        raise ValueError(
            f"{CONTROL_ID}: {source_name} is missing columns: "
            f"{missing_columns}."
        )


def _validate_unique_key(
    dataframe: pd.DataFrame,
    key_columns: list[str],
    source_name: str,
) -> int:
    """
    Validate one-row-per-key granularity without hiding duplicates.
    """
    duplicate_rows = dataframe.duplicated(
        subset=key_columns,
        keep=False,
    )

    duplicate_count = int(
        duplicate_rows.sum()
    )

    if duplicate_count:
        examples = (
            dataframe.loc[
                duplicate_rows,
                key_columns,
            ]
            .head(20)
            .to_dict("records")
        )

        raise ValueError(
            f"{CONTROL_ID}: {source_name} is not unique by "
            f"{key_columns}. Duplicate rows: {duplicate_count}. "
            f"Examples: {examples}"
        )

    return duplicate_count


def _normalize_tax_business_number_once(
    dataframe: pd.DataFrame,
) -> pd.DataFrame:
    """
    Normalize Tax/Business Number once per distinct source value.

    The original Tax/Business Number is retained for display. The normalized
    value is stored in a separate technical column.
    """
    result = dataframe.copy()

    source = (
        result["Tax/Business Number"]
        .astype("string")
        .fillna("")
    )

    unique_values = pd.Index(
        source.unique(),
        dtype="string",
    )

    normalized_lookup = pd.Series(
        (
            normalize_tax_id(value)
            for value in unique_values
        ),
        index=unique_values,
        dtype="string",
    )

    result["_Normalized Tax/Business Number"] = (
        source.map(normalized_lookup)
        .fillna("")
        .astype("string")
    )

    return result


def _prepare_last_invoices(
    postings: pd.DataFrame | None,
) -> pd.DataFrame:
    """
    Build and validate the common last-invoice population.
    """
    if postings is None:
        raise FileNotFoundError(
            f"{CONTROL_ID}: VPBSIK and VPBSAK are mandatory for the "
            "required VM05 output layout."
        )

    last_invoices = build_vm_last_invoice_population(
        postings
    )

    if not isinstance(
        last_invoices,
        pd.DataFrame,
    ):
        raise TypeError(
            f"{CONTROL_ID}: build_vm_last_invoice_population() must "
            "return a pandas DataFrame."
        )

    _require_columns(
        last_invoices,
        [
            *VENDOR_KEY_COLUMNS,
            *LAST_INVOICE_COLUMNS,
        ],
        "last-invoice population",
    )

    _validate_unique_key(
        last_invoices,
        VENDOR_KEY_COLUMNS,
        "last-invoice population",
    )

    return (
        last_invoices.loc[
            :,
            [
                *VENDOR_KEY_COLUMNS,
                *LAST_INVOICE_COLUMNS,
            ],
        ]
        .copy()
        .reset_index(drop=True)
    )


def build_vm_005(
    valid_population: pd.DataFrame,
    last_invoices: pd.DataFrame,
) -> tuple[pd.DataFrame, dict[str, int]]:
    """
    Build the VM05 exception population.

    Exception rule
    --------------
    Two or more distinct Vendor Codes share the same nonblank normalized
    Tax/Business Number within the same Company.

    No fuzzy comparison, minimum length, FX conversion or additional vendor
    exclusion is applied.
    """
    required_vendor_columns = {
        "Company",
        "Company Name",
        "Vendor Code",
        "Vendor Name",
        *TAX_BUSINESS_NUMBER_PRIORITY,
    }

    _require_columns(
        valid_population,
        required_vendor_columns,
        "valid vendor population",
    )

    _require_columns(
        last_invoices,
        [
            *VENDOR_KEY_COLUMNS,
            *LAST_INVOICE_COLUMNS,
        ],
        "last-invoice population",
    )

    _validate_unique_key(
        valid_population,
        VENDOR_KEY_COLUMNS,
        "valid vendor population",
    )

    _validate_unique_key(
        last_invoices,
        VENDOR_KEY_COLUMNS,
        "last-invoice population",
    )

    if valid_population.empty:
        return (
            pd.DataFrame(
                columns=OUTPUT_COLUMNS
            ),
            {
                "eligible_nonblank_tax_rows": 0,
                "exception_rows": 0,
                "groups": 0,
                "duplicate_output_key_rows": 0,
            },
        )

    vendors = valid_population.copy()

    vendors["Company"] = vendors[
        "Company"
    ].map(normalize_company)

    vendors = resolve_tax_business_number(
        vendors,
        priority_columns=TAX_BUSINESS_NUMBER_PRIORITY,
    )

    vendors = _normalize_tax_business_number_once(
        vendors
    )

    eligible = vendors.loc[
        vendors[
            "_Normalized Tax/Business Number"
        ].ne("")
    ].copy()

    eligible_nonblank_tax_rows = len(
        eligible
    )

    if eligible.empty:
        return (
            pd.DataFrame(
                columns=OUTPUT_COLUMNS
            ),
            {
                "eligible_nonblank_tax_rows": 0,
                "exception_rows": 0,
                "groups": 0,
                "duplicate_output_key_rows": 0,
            },
        )

    analytical_key_columns = [
        "Company",
        "_Normalized Tax/Business Number",
    ]

    distinct_vendor_count = (
        eligible.groupby(
            analytical_key_columns,
            sort=False,
            observed=True,
        )["Vendor Code"]
        .transform("nunique")
    )

    exceptions = eligible.loc[
        distinct_vendor_count.ge(2)
    ].copy()

    if exceptions.empty:
        return (
            pd.DataFrame(
                columns=OUTPUT_COLUMNS
            ),
            {
                "eligible_nonblank_tax_rows": eligible_nonblank_tax_rows,
                "exception_rows": 0,
                "groups": 0,
                "duplicate_output_key_rows": 0,
            },
        )

    exceptions["_Analytical Key"] = (
        exceptions["Company"]
        .astype("string")
        .str.cat(
            exceptions[
                "_Normalized Tax/Business Number"
            ].astype("string"),
            sep="\u00a6",
        )
    )

    exceptions["Group"] = (
        pd.factorize(
            exceptions["_Analytical Key"],
            sort=True,
        )[0]
        + 1
    )

    exceptions = exceptions.merge(
        last_invoices,
        how="left",
        on=VENDOR_KEY_COLUMNS,
        validate="one_to_one",
    )

    # Company remains the company code through every analytical merge.
    exceptions["CoCo"] = exceptions[
        "Company"
    ]

    # Only after all merges are complete is Company replaced with its
    # descriptive name for display.
    exceptions["Company"] = exceptions[
        "Company Name"
    ]

    output = (
        exceptions.sort_values(
            [
                "Group",
                "CoCo",
                "Vendor Code",
            ],
            kind="stable",
        )
        .loc[
            :,
            OUTPUT_COLUMNS,
        ]
        .reset_index(drop=True)
    )

    duplicate_output_key_rows = int(
        output.duplicated(
            subset=[
                "CoCo",
                "Vendor Code",
            ],
            keep=False,
        ).sum()
    )

    metrics = {
        "eligible_nonblank_tax_rows": eligible_nonblank_tax_rows,
        "exception_rows": len(output),
        "groups": int(output["Group"].nunique()),
        "duplicate_output_key_rows": duplicate_output_key_rows,
    }

    return output, metrics


def validate_vm_005_output(
    output: pd.DataFrame,
) -> None:
    """
    Validate the complete VM05 analytical contract vectorially.
    """
    if list(output.columns) != OUTPUT_COLUMNS:
        raise ValueError(
            f"{CONTROL_ID}: invalid output columns or order. "
            f"Expected {OUTPUT_COLUMNS}; received {list(output.columns)}."
        )

    if output.empty:
        return

    if output[
        [
            "CoCo",
            "Vendor Code",
        ]
    ].isna().any().any():
        raise ValueError(
            f"{CONTROL_ID}: output contains null CoCo or Vendor Code keys."
        )

    blank_coco = (
        output["CoCo"]
        .astype("string")
        .fillna("")
        .str.strip()
        .eq("")
    )

    blank_vendor = (
        output["Vendor Code"]
        .astype("string")
        .fillna("")
        .str.strip()
        .eq("")
    )

    if blank_coco.any() or blank_vendor.any():
        raise ValueError(
            f"{CONTROL_ID}: output contains blank CoCo or Vendor Code keys."
        )

    duplicate_output = output.duplicated(
        subset=[
            "CoCo",
            "Vendor Code",
        ],
        keep=False,
    )

    if duplicate_output.any():
        examples = (
            output.loc[
                duplicate_output,
                [
                    "CoCo",
                    "Vendor Code",
                ],
            ]
            .head(20)
            .to_dict("records")
        )

        raise ValueError(
            f"{CONTROL_ID}: output is not unique by CoCo + Vendor Code. "
            f"Examples: {examples}"
        )

    normalized_tax = (
        _normalize_tax_business_number_once(
            output
        )["_Normalized Tax/Business Number"]
    )

    if normalized_tax.eq("").any():
        raise ValueError(
            f"{CONTROL_ID}: output contains blank normalized "
            "Tax/Business Number keys."
        )

    validation = output.loc[
        :,
        [
            "CoCo",
            "Vendor Code",
            "Group",
        ],
    ].copy()

    validation["_Normalized Tax/Business Number"] = (
        normalized_tax
    )

    vendors_per_group = (
        validation.groupby(
            "Group",
            sort=False,
            observed=True,
        )["Vendor Code"]
        .nunique()
    )

    if vendors_per_group.lt(2).any():
        invalid_groups = (
            vendors_per_group.loc[
                vendors_per_group.lt(2)
            ]
            .index.tolist()
        )

        raise ValueError(
            f"{CONTROL_ID}: every Group must contain at least two "
            f"distinct Vendor Codes. Invalid groups: {invalid_groups[:20]}"
        )

    companies_per_group = (
        validation.groupby(
            "Group",
            sort=False,
            observed=True,
        )["CoCo"]
        .nunique()
    )

    if companies_per_group.ne(1).any():
        invalid_groups = (
            companies_per_group.loc[
                companies_per_group.ne(1)
            ]
            .index.tolist()
        )

        raise ValueError(
            f"{CONTROL_ID}: cross-company Groups detected: "
            f"{invalid_groups[:20]}"
        )

    tax_keys_per_group = (
        validation.groupby(
            "Group",
            sort=False,
            observed=True,
        )["_Normalized Tax/Business Number"]
        .nunique()
    )

    if tax_keys_per_group.ne(1).any():
        invalid_groups = (
            tax_keys_per_group.loc[
                tax_keys_per_group.ne(1)
            ]
            .index.tolist()
        )

        raise ValueError(
            f"{CONTROL_ID}: Groups contain more than one normalized "
            f"Tax/Business Number: {invalid_groups[:20]}"
        )

    group_values = (
        pd.to_numeric(
            output["Group"],
            errors="coerce",
        )
    )

    if group_values.isna().any():
        raise ValueError(
            f"{CONTROL_ID}: Group contains nonnumeric or null values."
        )

    if not group_values.mod(1).eq(0).all():
        raise ValueError(
            f"{CONTROL_ID}: Group contains noninteger values."
        )

    actual_groups = sorted(
        group_values.astype(int).unique().tolist()
    )

    expected_groups = list(
        range(
            1,
            len(actual_groups) + 1,
        )
    )

    if actual_groups != expected_groups:
        raise ValueError(
            f"{CONTROL_ID}: Group values must be consecutive from 1. "
            f"Received: {actual_groups[:20]}"
        )

    deterministic_key = (
        output["CoCo"]
        .astype("string")
        .str.cat(
            normalized_tax.astype("string"),
            sep="\u00a6",
        )
    )

    expected_group = (
        pd.factorize(
            deterministic_key,
            sort=True,
        )[0]
        + 1
    )

    if not (
        group_values.astype(int).to_numpy()
        == expected_group
    ).all():
        raise ValueError(
            f"{CONTROL_ID}: Group assignment is not deterministic "
            "according to sorted Company + normalized Tax/Business Number."
        )

    expected_order = (
        output.sort_values(
            [
                "Group",
                "CoCo",
                "Vendor Code",
            ],
            kind="stable",
        )
        .index.to_numpy()
    )

    if not (
        expected_order
        == output.index.to_numpy()
    ).all():
        raise ValueError(
            f"{CONTROL_ID}: output is not ordered by "
            "Group, CoCo, Vendor Code."
        )


def _metadata_count(
    metadata: dict[str, Any],
    possible_keys: tuple[str, ...],
) -> int | None:
    """
    Return the first available integer metric without assuming optional keys.
    """
    for key in possible_keys:
        value = metadata.get(
            key
        )

        if value is None:
            continue

        try:
            return int(
                value
            )
        except (
            TypeError,
            ValueError,
        ):
            continue

    return None


def run_vm_005(
    context: dict[str, Any],
) -> dict[str, Any]:
    """
    Execute VM05 and replace only the VM05 result worksheet.
    """
    total_started = perf_counter()
    timings: dict[str, float] = {}

    stage_started = perf_counter()

    vendor_source = load_vm_vendors(
        context
    )

    timings["vendor load"] = (
        perf_counter()
        - stage_started
    )

    source_vendor_rows = len(
        vendor_source
    )

    stage_started = perf_counter()

    vendor_master = build_vendor_master_population(
        vendor_source
    )

    master_vendor_rows = len(
        vendor_master
    )

    vendor_master, excluded_company = (
        _filter_configured_companies(
            vendor_master,
            context,
        )
    )

    valid_population, population_metrics = (
        get_valid_vendor_population(
            vendor_master
        )
    )

    timings["population build and filter"] = (
        perf_counter()
        - stage_started
    )

    stage_started = perf_counter()

    postings, posting_metadata = load_vm_vendor_postings(
        context
    )

    last_invoices = _prepare_last_invoices(
        postings
    )

    timings["posting load and last invoice"] = (
        perf_counter()
        - stage_started
    )

    stage_started = perf_counter()

    resolved_population = resolve_tax_business_number(
        valid_population,
        priority_columns=TAX_BUSINESS_NUMBER_PRIORITY,
    )

    normalized_population = _normalize_tax_business_number_once(
        resolved_population
    )

    timings["tax normalization"] = (
        perf_counter()
        - stage_started
    )

    stage_started = perf_counter()

    # build_vm_005 intentionally repeats no source load or last-invoice build.
    # The already-resolved display and technical values are carried in a
    # temporary analytical input, while the public builder still validates
    # all required tax source columns.
    analytical_input = valid_population.copy()

    analytical_input["Tax/Business Number"] = (
        normalized_population["Tax/Business Number"]
    )

    output, analytical_metrics = build_vm_005(
        analytical_input,
        last_invoices,
    )

    timings["analytical logic"] = (
        perf_counter()
        - stage_started
    )

    # No currency conversion applies to VM05.
    timings["FX conversion (not applicable)"] = 0.0

    stage_started = perf_counter()

    validate_vm_005_output(
        output
    )

    timings["validations"] = (
        perf_counter()
        - stage_started
    )

    stage_started = perf_counter()

    output_file = write_vm_control_sheet(
        context=context,
        sheet_name="VM05",
        dataframe=output,
        date_columns=[
            "Last Transaction Date",
        ],
        amount_columns=[
            "Last Inv Amt Doc Currency",
        ],
        integer_columns=[
            "Group",
        ],
    )

    timings["workbook write"] = (
        perf_counter()
        - stage_started
    )

    timings["total"] = (
        perf_counter()
        - total_started
    )

    posting_rows = (
        0
        if postings is None
        else len(postings)
    )

    re_kr_posting_rows = _metadata_count(
        posting_metadata,
        (
            "re_kr_rows",
            "invoice_rows",
            "eligible_invoice_rows",
            "re_kr_posting_rows",
        ),
    )

    warnings = [
        *population_metrics.get(
            "warnings",
            [],
        ),
        *posting_metadata.get(
            "warnings",
            [],
        ),
    ]

    print(
        f"{CONTROL_ID} source vendor rows: "
        f"{source_vendor_rows}"
    )
    print(
        f"{CONTROL_ID} master vendor rows: "
        f"{master_vendor_rows}"
    )
    print(
        f"{CONTROL_ID} rows excluded by CONFIG: "
        f"{excluded_company}"
    )
    print(
        f"{CONTROL_ID} valid vendor rows: "
        f"{len(valid_population)}"
    )
    print(
        f"{CONTROL_ID} posting rows: "
        f"{posting_rows}"
    )

    if re_kr_posting_rows is not None:
        print(
            f"{CONTROL_ID} RE/KR posting rows: "
            f"{re_kr_posting_rows}"
        )

    print(
        f"{CONTROL_ID} vendors with last invoice: "
        f"{len(last_invoices)}"
    )
    print(
        f"{CONTROL_ID} eligible nonblank "
        f"Tax/Business Number rows: "
        f"{analytical_metrics['eligible_nonblank_tax_rows']}"
    )
    print(
        f"{CONTROL_ID} exception rows: "
        f"{analytical_metrics['exception_rows']}"
    )
    print(
        f"{CONTROL_ID} groups: "
        f"{analytical_metrics['groups']}"
    )
    print(
        f"{CONTROL_ID} duplicate output-key rows: "
        f"{analytical_metrics['duplicate_output_key_rows']}"
    )

    for warning in warnings:
        print(
            f"WARNING: {warning}"
        )

    for timing_name, elapsed in timings.items():
        print(
            f"{CONTROL_ID} {timing_name}: "
            f"{elapsed:.2f}s"
        )

    return {
        "status": (
            "ERROR"
            if not output.empty
            else "OK"
        ),
        "output_file": output_file,
        "sheet_name": "VM05",
        "rows": len(output),
    }
