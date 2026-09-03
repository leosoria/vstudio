# VM_012 objective: identify valid vendors without a resolved Business Number.
# Functional priority agreed for the initial LBR/SAP ECC implementation:
# Tax Number 1, Tax Number 2, Tax Number 3, Tax Number 4, Tax Number 5,
# VAT Registration Number. The first nonblank value wins. Audit may revise
# this shared, explicit order in the future.

"""VM_012 - Valid vendors without a Business Number."""

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


CONTROL_ID = "VM_012"
SHEET_NAME = "VM12"

BUSINESS_NUMBER_PRIORITY = (
    "Tax Number 1",
    "Tax Number 2",
    "Tax Number 3",
    "Tax Number 4",
    "Tax Number 5",
    "VAT Registration Number",
)

OUTPUT_COLUMNS = [
    "Company",
    "CoCo",
    "Vendor Code",
    "Vendor Name",
    "Tax/Business Number",
]

_ALL_COMPANIES = {
    "ALL",
    "*",
    "TODAS",
    "TODOS",
}


def _print_timing(
    stage_name: str,
    started: float,
) -> float:
    """Print one stage duration and return its finish timestamp."""
    finished = perf_counter()

    print(
        f"{CONTROL_ID} {stage_name}: "
        f"{finished - started:.2f} seconds"
    )

    return finished


def _configured_companies(
    context: dict[str, Any],
) -> tuple[str, ...]:
    """
    Return configured company codes.

    An empty tuple means that every company in the vendor master is included.
    """
    module_config = context.get("module")

    if not isinstance(module_config, dict):
        raise ValueError(
            f"{CONTROL_ID} requires context['module'] configuration."
        )

    raw_companies = module_config.get("companies")

    if raw_companies is None:
        return ()

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

        if text == "" or text.upper() in _ALL_COMPANIES:
            return ()

        raw_values = re.split(
            r"[,;|]",
            text,
        )

    companies: list[str] = []

    for raw_value in raw_values:
        text = safe_text(raw_value)

        if text == "":
            continue

        if text.upper() in _ALL_COMPANIES:
            return ()

        company = normalize_company(text)

        if company and company not in companies:
            companies.append(company)

    return tuple(companies)


def _filter_configured_companies(
    vendor_master: pd.DataFrame,
    context: dict[str, Any],
) -> tuple[pd.DataFrame, int]:
    """Filter the normalized vendor master by configured companies."""
    companies = _configured_companies(context)

    if not companies:
        return (
            vendor_master.copy().reset_index(drop=True),
            0,
        )

    included = (
        vendor_master["Company"]
        .map(normalize_company)
        .isin(companies)
    )

    excluded_rows = int((~included).sum())

    return (
        vendor_master.loc[included]
        .copy()
        .reset_index(drop=True),
        excluded_rows,
    )


def build_vm_012(
    vendor_population: pd.DataFrame,
) -> pd.DataFrame:
    """
    Return valid vendors whose resolved Business Number is blank.

    The input must already contain the canonical Tax/Business Number column.
    This function only detects absence; it does not validate format, length,
    checksum, country rules, completeness or duplication of tax identifiers.
    """
    required_columns = {
        "Company",
        "Company Name",
        "Vendor Code",
        "Vendor Name",
        "Tax/Business Number",
    }

    missing_columns = sorted(
        required_columns.difference(
            vendor_population.columns
        )
    )

    if missing_columns:
        raise ValueError(
            f"{CONTROL_ID}: missing vendor population columns: "
            f"{missing_columns}."
        )

    if vendor_population.empty:
        return pd.DataFrame(
            columns=OUTPUT_COLUMNS
        )

    data = vendor_population.copy()

    data["CoCo"] = data[
        "Company"
    ].map(normalize_company)

    data["Vendor Code"] = data[
        "Vendor Code"
    ].map(normalize_vendor_code)

    data["Company"] = data[
        "Company Name"
    ].map(safe_text)

    data["Vendor Name"] = data[
        "Vendor Name"
    ].map(safe_text)

    data["Tax/Business Number"] = data[
        "Tax/Business Number"
    ].map(safe_text)

    exceptions = data.loc[
        data["CoCo"].ne("")
        & data["Vendor Code"].ne("")
        & data["Tax/Business Number"].eq("")
    ].copy()

    if exceptions.empty:
        return pd.DataFrame(
            columns=OUTPUT_COLUMNS
        )

    return (
        exceptions.drop_duplicates(
            subset=[
                "CoCo",
                "Vendor Code",
            ],
            keep="first",
        )
        .sort_values(
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


def run_vm_012(
    context: dict[str, Any],
) -> dict[str, Any]:
    """Execute VM_012 and replace only the VM12 result worksheet."""
    started = perf_counter()

    # Validate module configuration before loading any source file.
    _configured_companies(context)

    vendor_source = load_vm_vendors(
        context
    )

    vendor_master = build_vendor_master_population(
        vendor_source
    )

    master_rows = len(vendor_master)

    (
        filtered_master,
        excluded_company_rows,
    ) = _filter_configured_companies(
        vendor_master,
        context,
    )

    (
        valid_population,
        population_metrics,
    ) = get_valid_vendor_population(
        filtered_master
    )

    if valid_population.empty:
        raise ValueError(
            f"{CONTROL_ID}: valid vendor population is empty after CONFIG "
            "company and common VM exclusion rules."
        )

    resolved_population = resolve_tax_business_number(
        valid_population,
        priority_columns=BUSINESS_NUMBER_PRIORITY,
        output_column="Tax/Business Number",
    )

    stage_started = _print_timing(
        "input load and population validation",
        started,
    )

    output = build_vm_012(
        resolved_population
    )

    stage_started = _print_timing(
        "analytic logic",
        stage_started,
    )

    output_file = write_vm_control_sheet(
        context=context,
        sheet_name=SHEET_NAME,
        dataframe=output,
    )

    _print_timing(
        "workbook write",
        stage_started,
    )

    module_config = context["module"]

    print(
        f"{CONTROL_ID} period FROM/TO: "
        f"{safe_text(module_config.get('from'))} / "
        f"{safe_text(module_config.get('to'))}"
    )

    print(
        f"{CONTROL_ID} vendor source rows: "
        f"{len(vendor_source)}"
    )

    print(
        f"{CONTROL_ID} vendor master rows: "
        f"{master_rows}"
    )

    print(
        f"{CONTROL_ID} rows excluded by CONFIG company: "
        f"{excluded_company_rows}"
    )

    print(
        f"{CONTROL_ID} rows entering valid vendor population: "
        f"{population_metrics.get('input_rows', len(filtered_master))}"
    )

    print(
        f"{CONTROL_ID} rows excluded by Central Deletion Flag: "
        f"{population_metrics.get('excluded_central_deletion_flag', 0)}"
    )

    print(
        f"{CONTROL_ID} rows excluded by Company Deletion Flag: "
        f"{population_metrics.get('excluded_company_deletion_flag', 0)}"
    )

    print(
        f"{CONTROL_ID} rows excluded by vendor prefix E or T: "
        f"{population_metrics.get('excluded_vendor_prefix_e_or_t', 0)}"
    )

    print(
        f"{CONTROL_ID} rows excluded by intercompany Account Group: "
        f"{population_metrics.get('excluded_intercompany_account_group', 0)}"
    )

    print(
        f"{CONTROL_ID} valid vendor rows: "
        f"{population_metrics.get('output_rows', len(valid_population))}"
    )

    print(
        f"{CONTROL_ID} vendors without Business Number: "
        f"{len(output)}"
    )

    print(
        f"{CONTROL_ID} Business Number priority: "
        f"{', '.join(BUSINESS_NUMBER_PRIORITY)}"
    )

    for warning in population_metrics.get(
        "warnings",
        [],
    ):
        print(
            f"WARNING: {warning}"
        )

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
