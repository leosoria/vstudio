# VM_013 objective: identify valid vendors without a resolved Tax Number.
# Functional Tax Number definition for the initial LBR/SAP ECC implementation:
# Tax Number 1 (LFA1-STCD1), followed by Tax Number 2 (LFA1-STCD2).
# The first nonblank value wins. Audit may revise this explicit definition
# in the future.
#
# Functional and audit notes:
# - The priority was confirmed from the LBR/SAP ECC vm_vendors extract dated
#   2026-07-31.
# - The diagnostic population contained 24,276 unique Company + Vendor rows.
# - Tax Number 1 resolved 18,999 rows and Tax Number 2 resolved another 5,101
#   rows. The fields were complementary in the observed population.
# - Adding Tax Number 3, Tax Number 4, Tax Number 5 or VAT Registration Number
#   did not resolve any additional row after Tax Number 1 and Tax Number 2.
# - Tax Number 3 is deliberately excluded because it represents a different
#   fiscal attribute in the observed data and may contain values such as
#   ISENTO.
# - VM_013 identifies absence only. It does not validate format, length,
#   checksum, country structure, uniqueness, completeness or legitimacy.
# - A nonblank value such as "0", a zero sequence, an alphanumeric identifier
#   or an identifier containing punctuation is not considered absent.
# - VM_012 and VM_013 may produce the same exceptions in the current data, but
#   they are not functionally identical: VM_012 currently uses a broader
#   Business Number priority.
# - The global run_analysis.py currently treats a runner without a Python
#   exception as technically OK, even when this runner returns status ERROR
#   because audit exceptions were found.

"""
VM_013 - Valid vendors without a resolved Tax Number.

Objective
---------
Identify valid vendors whose canonical Tax Number is empty after applying the
approved initial LBR/SAP ECC functional resolution:

    1. Tax Number 1 (LFA1-STCD1)
    2. Tax Number 2 (LFA1-STCD2)

The first nonblank value wins.

Population
----------
The control operates only on the population returned by
get_valid_vendor_population(). Common deletion, employee/functionary,
Vendor Code prefix and intercompany exclusions remain the responsibility of
core.vm_common.

Analytic scope
--------------
VM_013 detects absence only. It does not perform:

- length validation;
- checksum validation;
- country-specific validation;
- format validation;
- duplicate-Tax-Number detection;
- legitimacy validation;
- transformations that turn an informed value into an empty value.

Output grain
------------
One row per unique CoCo + Vendor Code, written only to worksheet VM13.
"""

import re
from time import perf_counter
from typing import Any

import pandas as pd

from core.vm_common import (
    build_vendor_master_population,
    get_valid_vendor_population,
    load_vm_vendors,
    normalize_company,
    normalize_vendor_code,
    resolve_tax_business_number,
    safe_text,
    write_vm_control_sheet,
)


CONTROL_ID = "VM_013"
SHEET_NAME = "VM13"

# This tuple is an ordered functional definition. Do not convert it to a set,
# sort it alphabetically or append other fiscal fields merely to reduce the
# number of exceptions.
TAX_NUMBER_PRIORITY = (
    "Tax Number 1",
    "Tax Number 2",
)

TAX_NUMBER_PRIORITY_SOURCE = (
    "default confirmed from LBR/SAP ECC vm_vendors profiling dated 2026-07-31"
)

ALL_COMPANY_MARKERS = frozenset(
    {
        "ALL",
        "*",
        "TODAS",
        "TODOS",
    }
)

OUTPUT_COLUMNS = [
    "Company",
    "CoCo",
    "Vendor Code",
    "Vendor Name",
    "Tax Number",
]

REQUIRED_VENDOR_COLUMNS = {
    "Company",
    "Company Name",
    "Vendor Code",
    "Vendor Name",
    "Tax Number",
}


def _print_timing(
    stage_name: str,
    started: float,
) -> float:
    """Print one execution-stage timing and return its finish timestamp."""

    finished = perf_counter()

    print(
        f"{CONTROL_ID} {stage_name}: "
        f"{finished - started:.2f} seconds"
    )

    return finished


def _configured_companies(
    context: dict[str, Any],
) -> set[str]:
    """
    Return normalized company codes configured for the VM module.

    An empty set means all companies.

    Supported all-company values:
        None, blank, ALL, *, TODAS and TODOS.

    Supported multiple-company inputs:
        comma-delimited text;
        semicolon-delimited text;
        pipe-delimited text;
        list;
        tuple;
        set.

    If a collection contains an all-company marker, the complete population is
    included.
    """

    module_config = context.get("module")

    if not isinstance(module_config, dict):
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
        raw_values = list(raw_companies)
    else:
        text = safe_text(raw_companies)

        if (
            text == ""
            or text.upper() in ALL_COMPANY_MARKERS
        ):
            return set()

        raw_values = re.split(
            r"[,;|]",
            text,
        )

    cleaned_values = [
        safe_text(value)
        for value in raw_values
    ]

    if any(
        value.upper() in ALL_COMPANY_MARKERS
        for value in cleaned_values
        if value != ""
    ):
        return set()

    normalized_values = {
        normalize_company(value)
        for value in cleaned_values
        if value != ""
    }

    normalized_values.discard("")

    return normalized_values


def _filter_configured_companies(
    vendor_master: pd.DataFrame,
    context: dict[str, Any],
) -> tuple[pd.DataFrame, int]:
    """
    Apply the configured-company filter to the normalized vendor master.

    This function must be called after build_vendor_master_population() and
    before get_valid_vendor_population().
    """

    companies = _configured_companies(
        context
    )

    if not companies:
        return (
            vendor_master.copy().reset_index(drop=True),
            0,
        )

    if "Company" not in vendor_master.columns:
        raise ValueError(
            f"{CONTROL_ID}: vendor master is missing Company."
        )

    normalized_company = vendor_master[
        "Company"
    ].map(normalize_company)

    included = normalized_company.isin(
        companies
    )

    excluded_rows = int(
        (~included).sum()
    )

    filtered = (
        vendor_master.loc[included]
        .copy()
        .reset_index(drop=True)
    )

    return filtered, excluded_rows


def _prepare_valid_vendor_population(
    context: dict[str, Any],
) -> tuple[pd.DataFrame, dict[str, Any]]:
    """
    Load and prepare the VM_013 valid-vendor population.

    Processing order:

    1. Load vm_vendors exactly once.
    2. Build one vendor-master row per Company + Vendor Code.
    3. Capture master rows before the company filter.
    4. Filter configured companies.
    5. Apply common valid-vendor exclusions.
    """

    module_config = context.get("module")

    if not isinstance(module_config, dict):
        raise ValueError(
            f"{CONTROL_ID} requires context['module'] configuration."
        )

    vendor_source = load_vm_vendors(
        context
    )

    vendor_master = build_vendor_master_population(
        vendor_source
    )

    # Capture this metric before filtering configured companies.
    master_rows = len(
        vendor_master
    )

    (
        company_population,
        excluded_company,
    ) = _filter_configured_companies(
        vendor_master,
        context,
    )

    valid_population_input_rows = len(
        company_population
    )

    (
        valid_population,
        population_metrics,
    ) = get_valid_vendor_population(
        company_population
    )

    metrics = {
        "source_rows": len(vendor_source),
        "master_rows": master_rows,
        "excluded_company": excluded_company,
        "valid_population_input_rows": valid_population_input_rows,
        **population_metrics,
    }

    if valid_population.empty:
        raise ValueError(
            "VM_013: valid vendor population is empty after CONFIG "
            "company and common VM exclusion rules."
        )

    return valid_population, metrics


def _resolve_tax_number(
    vendor_population: pd.DataFrame,
) -> pd.DataFrame:
    """
    Add the canonical Tax Number using the approved ordered priority.

    Resolution is intentionally performed before build_vm_013(), keeping the
    pure analytic function independent from the fiscal-priority decision.

    No normalized-output helper column is requested because VM_013 detects
    absence only, not identifier format.
    """

    return resolve_tax_business_number(
        vendor_population,
        TAX_NUMBER_PRIORITY,
        output_column="Tax Number",
    )


def build_vm_013(
    vendor_population: pd.DataFrame,
) -> pd.DataFrame:
    """
    Return valid vendors whose already-resolved Tax Number is blank.

    This function is pure:

    - it does not read files;
    - it does not write files;
    - it does not access context;
    - it does not resolve fiscal priority;
    - it does not reapply common exclusions;
    - it does not mutate vendor_population.

    Parameters
    ----------
    vendor_population:
        Valid-vendor population containing a canonical Tax Number column
        resolved before this function is called.

    Returns
    -------
    pandas.DataFrame
        VM13 exceptions with exactly OUTPUT_COLUMNS.
    """

    missing_columns = sorted(
        REQUIRED_VENDOR_COLUMNS.difference(
            vendor_population.columns
        )
    )

    if missing_columns:
        raise ValueError(
            f"{CONTROL_ID}: missing vendor population columns: "
            f"{missing_columns}"
        )

    if vendor_population.empty:
        return pd.DataFrame(
            columns=OUTPUT_COLUMNS
        )

    data = vendor_population.copy()

    # CoCo is the normalized SAP company code.
    data["CoCo"] = data[
        "Company"
    ].map(normalize_company)

    data["Vendor Code"] = data[
        "Vendor Code"
    ].map(normalize_vendor_code)

    # The LBR display column named Company contains the clean company name,
    # while CoCo contains the normalized company code.
    data["Company"] = data[
        "Company Name"
    ].map(safe_text)

    data["Vendor Name"] = data[
        "Vendor Name"
    ].map(safe_text)

    data["Tax Number"] = data[
        "Tax Number"
    ].map(safe_text)

    exceptions = data.loc[
        data["Tax Number"].eq("")
        & data["CoCo"].ne("")
        & data["Vendor Code"].ne("")
    ].copy()

    if exceptions.empty:
        return pd.DataFrame(
            columns=OUTPUT_COLUMNS
        )

    # Defensive protection against artificial source multiplication. Vendor
    # master integrity conflicts must already have been detected by
    # build_vendor_master_population() and are not suppressed here.
    exceptions = exceptions.drop_duplicates(
        subset=[
            "CoCo",
            "Vendor Code",
        ],
        keep="first",
    )

    return (
        exceptions.sort_values(
            [
                "Company",
                "CoCo",
                "Vendor Name",
                "Vendor Code",
            ],
            kind="mergesort",
        )
        .loc[:, OUTPUT_COLUMNS]
        .reset_index(drop=True)
    )


def _print_metrics(
    context: dict[str, Any],
    metrics: dict[str, Any],
    output: pd.DataFrame,
) -> None:
    """
    Print period, population and Tax Number metrics.

    Individual exclusion metrics may overlap. For example, one vendor may have
    both a deletion flag and Account Group ZFUN. Therefore their individual
    counts must not be added to reconcile the population.

    The net common-rule exclusion is the authoritative reconciliation between
    the company-filtered input and valid-vendor output.
    """

    module_config = context["module"]

    print(
        f"{CONTROL_ID} period FROM: "
        f"{safe_text(module_config.get('from', ''))}"
    )
    print(
        f"{CONTROL_ID} period TO: "
        f"{safe_text(module_config.get('to', ''))}"
    )

    print(
        f"{CONTROL_ID} original vendor source rows: "
        f"{metrics.get('source_rows', 0)}"
    )
    print(
        f"{CONTROL_ID} vendor master rows before company filter: "
        f"{metrics.get('master_rows', 0)}"
    )
    print(
        f"{CONTROL_ID} rows excluded by CONFIG company: "
        f"{metrics.get('excluded_company', 0)}"
    )
    print(
        f"{CONTROL_ID} rows delivered to "
        f"get_valid_vendor_population: "
        f"{metrics.get('valid_population_input_rows', 0)}"
    )

    print(
        f"{CONTROL_ID} rows matching Central Deletion Flag exclusion: "
        f"{metrics.get('excluded_central_deletion_flag', 0)}"
    )
    print(
        f"{CONTROL_ID} rows matching Company Deletion Flag exclusion: "
        f"{metrics.get('excluded_company_deletion_flag', 0)}"
    )
    print(
        f"{CONTROL_ID} rows matching Vendor Code prefix E or T exclusion: "
        f"{metrics.get('excluded_vendor_prefix_e_or_t', 0)}"
    )
    print(
        f"{CONTROL_ID} rows matching employee Account Group ZFUN exclusion: "
        f"{metrics.get('excluded_employee_account_group_zfun', 0)}"
    )
    print(
        f"{CONTROL_ID} configured intercompany Vendor Codes: "
        f"{metrics.get('configured_intercompany_vendor_codes', 0)}"
    )
    print(
        f"{CONTROL_ID} rows matching intercompany Vendor Code exclusion: "
        f"{metrics.get('excluded_intercompany_vendor_code', 0)}"
    )

    valid_input_rows = metrics.get(
        "valid_population_input_rows",
        0,
    )
    valid_output_rows = metrics.get(
        "output_rows",
        0,
    )

    net_common_exclusions = max(
        valid_input_rows - valid_output_rows,
        0,
    )

    print(
        f"{CONTROL_ID} net rows excluded by common VM rules: "
        f"{net_common_exclusions}"
    )
    print(
        f"{CONTROL_ID} Trading Partner nonblank rows: "
        f"{metrics.get('trading_partner_nonblank_rows', 0)}"
    )
    print(
        f"{CONTROL_ID} valid vendor population rows: "
        f"{valid_output_rows}"
    )
    print(
        f"{CONTROL_ID} vendors without Tax Number: "
        f"{len(output)}"
    )

    print(
        f"{CONTROL_ID} Tax Number definition: "
        "first nonblank value from the approved ordered priority"
    )
    print(
        f"{CONTROL_ID} Tax Number priority source: "
        f"{TAX_NUMBER_PRIORITY_SOURCE}"
    )
    print(
        f"{CONTROL_ID} Tax Number priority: "
        f"{' > '.join(TAX_NUMBER_PRIORITY)}"
    )

    print(
        f"{CONTROL_ID} VM_012 comparison note: "
        "VM_012 and VM_013 may currently return the same vendor keys, "
        "but their approved Business Number and Tax Number priorities "
        "are different."
    )


def _print_warnings(
    metrics: dict[str, Any],
) -> None:
    """Print every warning returned by common VM population processing."""

    for warning in metrics.get(
        "warnings",
        [],
    ):
        print(
            f"WARNING: {warning}"
        )


def run_vm_013(
    context: dict[str, Any],
) -> dict[str, Any]:
    """
    Execute VM_013 and replace only worksheet VM13.

    Exact execution flow:

    1. Validate context["module"].
    2. Interpret configured companies.
    3. Load vendors exactly once.
    4. Build the vendor master.
    5. Capture master rows before company filtering.
    6. Apply the configured-company filter.
    7. Apply get_valid_vendor_population().
    8. Resolve the canonical Tax Number.
    9. Execute build_vm_013().
    10. Write only VM13.
    11. Print metrics.
    12. Print warnings.
    13. Return the standard result dictionary.

    Status semantics returned by this runner:

        OK:
            no vendor without Tax Number was found.

        ERROR:
            one or more vendors without Tax Number were found.

        Python exception:
            the control could not execute correctly.

    Note that the current global run_analysis.py may report technical execution
    as OK whenever no Python exception occurs, even if this dictionary returns
    status ERROR.
    """

    started = perf_counter()

    module_config = context.get("module")

    if not isinstance(module_config, dict):
        raise ValueError(
            f"{CONTROL_ID} requires context['module'] configuration."
        )

    # Validate company configuration before reading any input file.
    _configured_companies(
        context
    )

    (
        valid_population,
        metrics,
    ) = _prepare_valid_vendor_population(
        context
    )

    stage_started = _print_timing(
        "input load and population validation",
        started,
    )

    resolved_population = _resolve_tax_number(
        valid_population
    )

    output = build_vm_013(
        resolved_population
    )

    stage_started = _print_timing(
        "analytic logic",
        stage_started,
    )

    # VM13 has no date, amount or integer formatting requirements. Passing only
    # these three arguments also ensures that no unrelated worksheet is opened
    # or rewritten by this control.
    output_file = write_vm_control_sheet(
        context=context,
        sheet_name=SHEET_NAME,
        dataframe=output,
    )

    _print_timing(
        "workbook write",
        stage_started,
    )

    _print_metrics(
        context=context,
        metrics=metrics,
        output=output,
    )

    _print_warnings(
        metrics
    )

    return {
        "status": "ERROR" if not output.empty else "OK",
        "output_file": output_file,
        "sheet_name": SHEET_NAME,
        "rows": len(output),
    }
