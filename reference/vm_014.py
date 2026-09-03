# VM_014 objective: identify valid vendors without a reported primary phone.
# Functional phone definition for the initial LBR/SAP ECC implementation:
# Phone1, sourced from Telephone 1 / LFA1-TELF1 in vm_vendors.
# A vendor is an exception when Phone1 is blank after safe_text.
# VM_014 identifies absence only; it does not validate phone format.
# Audit may revise this explicit definition in the future.
# Phone1 is not combined with other telephone fields, and no phone is inferred
# from addresses, banks, contacts, or any other source.
# No minimum digit count or country-specific validation is applied, and a
# reported value is never transformed into blank to determine absence.
# Nonblank signs, letters, or zeroes remain reported values: this control tests
# absence, not phone-format quality.
# The historical control used Phone1 and selected blank Phone1 values.
# The initial LBR/SAP ECC definition uses canonical Phone1, already mapped by
# core.vm_common from Telephone 1 / TELF1 / LFA1-TELF1.

"""VM_014 - Valid vendors without a reported primary phone."""

from time import perf_counter
from typing import Any

import pandas as pd

from core.vm_common import (
    build_vendor_master_population,
    get_valid_vendor_population,
    load_vm_vendors,
    normalize_company,
    normalize_vendor_code,
    safe_text,
    write_vm_control_sheet,
)


CONTROL_ID = "VM_014"
SHEET_NAME = "VM14"

REQUIRED_COLUMNS = [
    "Company",
    "Company Name",
    "Vendor Code",
    "Vendor Name",
    "Phone1",
]

OUTPUT_COLUMNS = [
    "Company",
    "CoCo",
    "Vendor Code",
    "Vendor Name",
    "Phone1",
]

_ALL_COMPANY_VALUES = {
    "ALL",
    "*",
    "TODAS",
    "TODOS",
}


def _print_timing(
    stage_name: str,
    started: float,
) -> float:
    """Print one lightweight stage timing and return the finish time."""
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
    Return configured company codes.

    An empty set means that every company present in the normalized vendor
    master is included.
    """
    module_config = context.get("module")

    if not isinstance(module_config, dict):
        raise ValueError(
            "VM_014 requires context['module'] configuration."
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
            or text.upper() in _ALL_COMPANY_VALUES
        ):
            return set()

        raw_values = (
            text.replace(";", ",")
            .replace("|", ",")
            .split(",")
        )

    if any(
        safe_text(value).upper() in _ALL_COMPANY_VALUES
        for value in raw_values
    ):
        return set()

    return {
        normalize_company(value)
        for value in raw_values
        if safe_text(value) != ""
    }


def _filter_configured_companies(
    vendor_master: pd.DataFrame,
    companies: set[str],
) -> tuple[pd.DataFrame, int]:
    """
    Filter the normalized master after master construction.

    Company filtering is intentionally performed after
    build_vendor_master_population and before get_valid_vendor_population.
    """
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

    excluded_rows = int(
        (~included).sum()
    )

    return (
        vendor_master.loc[included]
        .copy()
        .reset_index(drop=True),
        excluded_rows,
    )


def build_vm_014(
    vendor_population: pd.DataFrame,
) -> pd.DataFrame:
    """
    Return valid vendors whose canonical Phone1 is absent after safe_text.

    The input must already be the result of get_valid_vendor_population.
    This function does not read files, write files, access runner context,
    select alternative phone sources, or repeat common exclusion rules.
    """
    missing = sorted(
        set(REQUIRED_COLUMNS).difference(
            vendor_population.columns
        )
    )

    if missing:
        raise ValueError(
            f"{CONTROL_ID}: missing vendor population columns: "
            f"{missing}"
        )

    working = vendor_population.loc[
        :,
        REQUIRED_COLUMNS,
    ].copy()

    working["CoCo"] = working[
        "Company"
    ].map(normalize_company)

    working["Company"] = working[
        "Company Name"
    ].map(safe_text)

    working["Vendor Code"] = working[
        "Vendor Code"
    ].map(normalize_vendor_code)

    working["Vendor Name"] = working[
        "Vendor Name"
    ].map(safe_text)

    working["Phone1"] = working[
        "Phone1"
    ].map(safe_text)

    # Deliberately do not use normalize_phone: nonblank content is reported,
    # regardless of punctuation, letters, digit count, zeroes, or country.
    exceptions = working.loc[
        working["CoCo"].ne("")
        & working["Vendor Code"].ne("")
        & working["Phone1"].eq("")
    ].copy()

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


def run_vm_014(
    context: dict[str, Any],
) -> dict[str, Any]:
    """Execute VM14 and replace only the VM14 result worksheet."""
    started = perf_counter()

    # Validate module configuration and interpret companies before file access.
    companies = _configured_companies(
        context
    )

    vendor_source = load_vm_vendors(
        context
    )

    vendor_master = build_vendor_master_population(
        vendor_source
    )

    master_rows = len(
        vendor_master
    )

    (
        company_population,
        excluded_company_rows,
    ) = _filter_configured_companies(
        vendor_master,
        companies,
    )

    valid_population_input_rows = len(
        company_population
    )

    (
        valid_population,
        metrics,
    ) = get_valid_vendor_population(
        company_population
    )

    if valid_population.empty:
        raise ValueError(
            "VM_014: valid vendor population is empty after CONFIG "
            "company and common VM exclusion rules."
        )

    stage_started = _print_timing(
        "input load and population validation",
        started,
    )

    output = build_vm_014(
        valid_population
    )

    stage_started = _print_timing(
        "preparation and analytic logic",
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

    output_rows = int(
        metrics.get(
            "output_rows",
            0,
        )
    )

    net_common_exclusion = max(
        valid_population_input_rows - output_rows,
        0,
    )

    print(
        f"{CONTROL_ID} period FROM: "
        f"{safe_text(context['module'].get('from', ''))}"
    )
    print(
        f"{CONTROL_ID} period TO: "
        f"{safe_text(context['module'].get('to', ''))}"
    )
    print(
        f"{CONTROL_ID} original vendor source rows: "
        f"{len(vendor_source)}"
    )
    print(
        f"{CONTROL_ID} master rows before CONFIG company filter: "
        f"{master_rows}"
    )
    print(
        f"{CONTROL_ID} rows excluded by CONFIG company: "
        f"{excluded_company_rows}"
    )
    print(
        f"{CONTROL_ID} rows submitted to "
        f"get_valid_vendor_population: "
        f"{valid_population_input_rows}"
    )
    print(
        f"{CONTROL_ID} rows matching Central Deletion Flag: "
        f"{metrics.get('excluded_central_deletion_flag', 0)}"
    )
    print(
        f"{CONTROL_ID} rows matching Company Deletion Flag: "
        f"{metrics.get('excluded_company_deletion_flag', 0)}"
    )
    print(
        f"{CONTROL_ID} rows matching vendor prefix E or T: "
        f"{metrics.get('excluded_vendor_prefix_e_or_t', 0)}"
    )
    print(
        f"{CONTROL_ID} rows matching Account Group ZFUN: "
        f"{metrics.get('excluded_employee_account_group_zfun', 0)}"
    )
    print(
        f"{CONTROL_ID} configured intercompany Vendor Codes: "
        f"{metrics.get('configured_intercompany_vendor_codes', 0)}"
    )
    print(
        f"{CONTROL_ID} rows matching intercompany Vendor Code: "
        f"{metrics.get('excluded_intercompany_vendor_code', 0)}"
    )

    # Individual exclusion metrics can overlap and must not be summed to
    # reconcile the population; use the input-to-output difference instead.
    print(
        f"{CONTROL_ID} net common-rule exclusion rows: "
        f"{net_common_exclusion}"
    )
    print(
        f"{CONTROL_ID} rows with nonblank Trading Partner: "
        f"{metrics.get('trading_partner_nonblank_rows', 0)}"
    )
    print(
        f"{CONTROL_ID} valid population rows: "
        f"{output_rows}"
    )
    print(
        f"{CONTROL_ID} Phone definition: "
        "Phone1 after safe_text"
    )
    print(
        f"{CONTROL_ID} Phone source: "
        "Telephone 1 / LFA1-TELF1"
    )
    print(
        f"{CONTROL_ID} vendors without Phone1: "
        f"{len(output)}"
    )

    for warning in metrics.get(
        "warnings",
        [],
    ):
        print(
            f"WARNING: {warning}"
        )

    # run_analysis.py currently records technical OK whenever no Python
    # exception occurs, even when this result reports audit findings as ERROR.
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
