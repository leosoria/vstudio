# VM_015 objective: identify valid vendors without reported bank master data.
# Functional bank-data definition for the initial LBR/SAP ECC implementation:
# vendor bank data sourced from LFBK through the canonical bank columns already
# mapped in vm_vendors by core.vm_common.
# A vendor is an exception when it has no bank record according to the official
# build_vendor_bank_population semantics.
# VM_015 identifies absence only; it does not validate bank-data quality,
# completeness, legitimacy, format, or ownership.
# Audit may revise this explicit definition in the future.
# Bank data is not inferred from invoices, payments, text, addresses, contacts,
# or any other source; BSIK and BSAK are not loaded, and company-owned bank data
# is not used.
# Historical OCRD, OCRB, DflAccount, HouseBank, CBU, and other SAP Business One
# rules are intentionally not incorporated into this LBR/SAP ECC control.
# A bank record need not have every component completed. VM_015 imposes no
# minimum Bank Account length and validates neither IBAN, SWIFT/BIC, country
# format, nor the existence of Bank Code in a bank catalogue.
# Account Holder Name need not match Vendor Name, and Bank Valid From/To validity
# is not used to determine absence.
# A reported bank value is never turned into blank by destructive normalization.
# VM_015 controls bank-data absence, not quality, completeness, legitimacy, or
# validity. The historical reference defines VM15 as vendors with no bank
# data in any source; this initial LBR/SAP ECC implementation instead uses only
# the canonical LFBK columns already mapped by core.vm_common in vm_vendors.

"""VM_015 - Valid vendors without reported LFBK bank master data."""

from time import perf_counter
from typing import Any

import pandas as pd

from core.vm_common import (
    build_vendor_bank_population,
    build_vendor_master_population,
    get_valid_vendor_population,
    load_vm_vendors,
    normalize_company,
    normalize_vendor_code,
    safe_text,
    write_vm_control_sheet,
)


CONTROL_ID = "VM_015"
SHEET_NAME = "VM15"

OUTPUT_COLUMNS = [
    "Company",
    "CoCo",
    "Vendor Code",
    "Vendor Name",
]

_VENDOR_REQUIRED_COLUMNS = {
    "Company",
    "Company Name",
    "Vendor Code",
    "Vendor Name",
}
_BANK_REQUIRED_COLUMNS = {
    "Company",
    "Vendor Code",
}
_ALL_COMPANY_TOKENS = {
    "ALL",
    "*",
    "TODAS",
    "TODOS",
}


def _print_timing(stage_name: str, started: float) -> float:
    """Print one lightweight stage timing and return its finish time."""
    finished = perf_counter()
    print(f"{CONTROL_ID} {stage_name}: {finished - started:.2f} seconds")
    return finished


def _configured_companies(module_config: dict[str, Any]) -> set[str]:
    """Return normalized configured companies; an empty set means all."""
    raw_companies = module_config.get("companies", "")

    if raw_companies is None:
        return set()

    if isinstance(raw_companies, (list, tuple, set)):
        raw_values = list(raw_companies)
    else:
        text = safe_text(raw_companies)
        if text == "" or text.upper() in _ALL_COMPANY_TOKENS:
            return set()
        raw_values = text.replace(";", ",").replace("|", ",").split(",")

    if any(
        safe_text(value).upper() in _ALL_COMPANY_TOKENS
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
    """Filter the normalized master after it has been built."""
    if not companies:
        return vendor_master.copy().reset_index(drop=True), 0

    included = vendor_master["Company"].map(
        normalize_company
    ).isin(companies)
    excluded_rows = int((~included).sum())

    return (
        vendor_master.loc[included].copy().reset_index(drop=True),
        excluded_rows,
    )


def _normalized_vendor_keys(
    dataframe: pd.DataFrame,
) -> set[tuple[str, str]]:
    """Return unique, nonblank normalized Company + Vendor Code keys."""
    companies = dataframe["Company"].map(normalize_company)
    vendor_codes = dataframe["Vendor Code"].map(normalize_vendor_code)

    return {
        (company, vendor_code)
        for company, vendor_code in zip(companies, vendor_codes)
        if company != "" and vendor_code != ""
    }


def build_vm_015(
    vendor_population: pd.DataFrame,
    vendor_bank_population: pd.DataFrame,
) -> pd.DataFrame:
    """Return valid vendors absent from the official vendor bank population."""
    missing_vendor_columns = sorted(
        _VENDOR_REQUIRED_COLUMNS.difference(vendor_population.columns)
    )

    if missing_vendor_columns:
        raise ValueError(
            f"{CONTROL_ID}: missing vendor population columns: "
            f"{missing_vendor_columns}"
        )

    missing_bank_columns = sorted(
        _BANK_REQUIRED_COLUMNS.difference(
            vendor_bank_population.columns
        )
    )

    if missing_bank_columns:
        raise ValueError(
            f"{CONTROL_ID}: missing vendor bank population columns: "
            f"{missing_bank_columns}"
        )

    vendors = vendor_population.loc[
        :,
        [
            "Company",
            "Company Name",
            "Vendor Code",
            "Vendor Name",
        ],
    ].copy()

    vendors["CoCo"] = vendors["Company"].map(normalize_company)
    vendors["Vendor Code"] = vendors["Vendor Code"].map(
        normalize_vendor_code
    )
    vendors["Company"] = vendors["Company Name"].map(safe_text)
    vendors["Vendor Name"] = vendors["Vendor Name"].map(safe_text)

    bank_keys = _normalized_vendor_keys(vendor_bank_population)

    vendor_keys = pd.Series(
        list(zip(vendors["CoCo"], vendors["Vendor Code"])),
        index=vendors.index,
        dtype=object,
    )

    exceptions = vendors.loc[
        vendors["CoCo"].ne("")
        & vendors["Vendor Code"].ne("")
        & ~vendor_keys.isin(bank_keys)
    ].copy()

    return (
        exceptions.drop_duplicates(
            subset=["CoCo", "Vendor Code"],
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


def run_vm_015(
    context: dict[str, Any],
) -> dict[str, Any]:
    """Execute VM15 and replace only the VM15 result worksheet."""
    started = perf_counter()

    module_config = context.get("module")

    if not isinstance(module_config, dict):
        raise ValueError(
            "VM_015 requires context['module'] configuration."
        )

    companies = _configured_companies(module_config)

    vendor_source = load_vm_vendors(context)

    vendor_master = build_vendor_master_population(vendor_source)
    master_rows = len(vendor_master)

    vendor_bank_population = build_vendor_bank_population(
        vendor_source
    )

    filtered_master, excluded_company_rows = (
        _filter_configured_companies(
            vendor_master,
            companies,
        )
    )

    valid_population_input_rows = len(filtered_master)

    valid_population, metrics = get_valid_vendor_population(
        filtered_master
    )

    if valid_population.empty:
        raise ValueError(
            "VM_015: valid vendor population is empty after CONFIG "
            "company and common VM exclusion rules."
        )

    stage_started = _print_timing(
        "input load and population validation",
        started,
    )

    output = build_vm_015(
        valid_population,
        vendor_bank_population,
    )

    valid_keys = _normalized_vendor_keys(valid_population)
    bank_keys = _normalized_vendor_keys(vendor_bank_population)

    valid_with_bank = len(
        valid_keys.intersection(bank_keys)
    )
    valid_without_bank = len(
        valid_keys.difference(bank_keys)
    )
    unique_bank_keys = len(bank_keys)

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

    print(
        f"{CONTROL_ID} period FROM: "
        f"{safe_text(module_config.get('from', ''))}"
    )
    print(
        f"{CONTROL_ID} period TO: "
        f"{safe_text(module_config.get('to', ''))}"
    )
    print(
        f"{CONTROL_ID} source rows: "
        f"{len(vendor_source)}"
    )
    print(
        f"{CONTROL_ID} master rows before CONFIG company filter: "
        f"{master_rows}"
    )
    print(
        f"{CONTROL_ID} bank population rows: "
        f"{len(vendor_bank_population)}"
    )
    print(
        f"{CONTROL_ID} unique Company + Vendor Code keys "
        f"with bank data: {unique_bank_keys}"
    )
    print(
        f"{CONTROL_ID} rows excluded by CONFIG company: "
        f"{excluded_company_rows}"
    )
    print(
        f"{CONTROL_ID} rows passed to "
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

    # Individual exclusion metrics may overlap and must not be summed to
    # reconcile the population; only input minus output is the net
    # common-rule exclusion.
    net_common_exclusion = max(
        valid_population_input_rows
        - metrics.get("output_rows", 0),
        0,
    )

    print(
        f"{CONTROL_ID} net rows excluded by common VM rules: "
        f"{net_common_exclusion}"
    )
    print(
        f"{CONTROL_ID} rows with Trading Partner populated: "
        f"{metrics.get('trading_partner_nonblank_rows', 0)}"
    )
    print(
        f"{CONTROL_ID} valid population rows: "
        f"{metrics.get('output_rows', 0)}"
    )
    print(
        "VM_015 Bank definition: presence in "
        "build_vendor_bank_population"
    )
    print(
        "VM_015 Bank source: LFBK canonical columns in vm_vendors"
    )
    print(
        f"VM_015 valid vendors with bank data: "
        f"{valid_with_bank}"
    )
    print(
        f"VM_015 valid vendors without bank data: "
        f"{valid_without_bank}"
    )

    for warning in metrics.get("warnings", []):
        print(f"WARNING: {warning}")

    # The current run_analysis.py can report technical OK whenever no Python
    # exception occurs, even when this dictionary reports audit findings as
    # ERROR.
    return {
        "status": "ERROR" if not output.empty else "OK",
        "output_file": output_file,
        "sheet_name": SHEET_NAME,
        "rows": len(output),
    }
