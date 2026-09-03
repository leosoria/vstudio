# VM_016 objective: identify valid vendors whose only reported central address
# is a post-office-box address rather than a physical street address.
# For the initial LBR/SAP ECC implementation, the address source is the canonical
# LFA1 central-address data already mapped in vm_vendors by core.vm_common.
# LBR currently provides one central vendor address, so a valid vendor is an
# exception when its canonical Street matches an approved PO Box pattern.
# Blank addresses do not qualify, and no address is inferred from invoices,
# payments, contacts, bank data, BSIK, BSAK, or other sources.

"""VM_016 - Valid vendors with only a post-office-box central address."""

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
    safe_text,
    write_vm_control_sheet,
)


CONTROL_ID = "VM_016"
SHEET_NAME = "VM16"

# Approved VM_016 PO Box patterns. The generic BOX alternative is retained as
# part of the explicit initial LBR functional definition.
POBOX_PATTERNS = (
    r"\bP\.?\s*O\.?\s*BOX\b",
    r"\bPOST\s*OFFICE\s*BOX\b",
    r"\bCASILLA\b",
    r"\bAPARTADO\b",
    r"\bAP\.?\s*POSTAL\b",
    r"\bBOX\b",
)

_POBOX_RE = re.compile(
    "|".join(POBOX_PATTERNS),
    re.IGNORECASE,
)

OUTPUT_COLUMNS = [
    "CoCo",
    "Vendor Code",
    "Address Type",
    "Street",
    "Building",
    "City",
    "Country",
    "Company",
    "Vendor Name",
    "Last Invoice Number",
    "Last Transaction Date",
    "Last Inv Amt Doc Currency",
    "Last Inv Amt Doc Currency Indicator",
]

_REQUIRED_COLUMNS = {
    "Company",
    "Company Name",
    "Vendor Code",
    "Vendor Name",
    "Street",
    "City",
    "Country",
}

_ALL_COMPANY_TOKENS = {
    "ALL",
    "*",
    "TODAS",
    "TODOS",
}


def _print_timing(
    stage_name: str,
    started: float,
) -> float:
    """Print one lightweight stage timing and return its finish time."""
    finished = perf_counter()

    print(
        f"{CONTROL_ID} {stage_name}: "
        f"{finished - started:.2f} seconds"
    )

    return finished


def _configured_companies(
    module_config: dict[str, Any],
) -> set[str]:
    """
    Return normalized configured companies.

    An empty set means every company present in the validated vendor
    population. ALL, *, TODAS, TODOS, empty text and None all mean every
    company.
    """
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
            or text.upper() in _ALL_COMPANY_TOKENS
        ):
            return set()

        raw_values = (
            text.replace(
                ";",
                ",",
            )
            .replace(
                "|",
                ",",
            )
            .split(
                ","
            )
        )

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
    """Filter the normalized vendor master after it has been built."""
    if not companies:
        return (
            vendor_master.copy().reset_index(
                drop=True
            ),
            0,
        )

    included = (
        vendor_master["Company"]
        .map(
            normalize_company
        )
        .isin(
            companies
        )
    )

    excluded_rows = int(
        (~included).sum()
    )

    return (
        vendor_master.loc[
            included
        ]
        .copy()
        .reset_index(
            drop=True
        ),
        excluded_rows,
    )


def build_vm_016(
    vendor_population: pd.DataFrame,
) -> pd.DataFrame:
    """
    Return valid vendors whose canonical Street reports a PO Box only.

    VM_016 evaluates whether the only reported central vendor address is a PO
    Box rather than a physical street. LBR/SAP ECC exposes one central LFA1
    address per Company + Vendor Code after build_vendor_master_population, so
    a matching Street is the vendor's only reported central address.

    The required VM_016 output layout is preserved. Fields unavailable in the
    LBR source remain blank and are not inferred from transactional sources.
    """
    missing_columns = sorted(
        _REQUIRED_COLUMNS.difference(
            vendor_population.columns
        )
    )

    if missing_columns:
        raise ValueError(
            f"{CONTROL_ID}: missing vendor population columns: "
            f"{missing_columns}"
        )

    data = vendor_population.loc[
        :,
        [
            "Company",
            "Company Name",
            "Vendor Code",
            "Vendor Name",
            "Street",
            "City",
            "Country",
        ],
    ].copy()

    data["CoCo"] = data[
        "Company"
    ].map(
        normalize_company
    )

    data["Vendor Code"] = data[
        "Vendor Code"
    ].map(
        normalize_vendor_code
    )

    data["Company"] = data[
        "Company Name"
    ].map(
        safe_text
    )

    for column in [
        "Vendor Name",
        "Street",
        "City",
        "Country",
    ]:
        data[column] = data[
            column
        ].map(
            safe_text
        )

    # ECC LFA1 supplies one central address rather than separate address rows.
    # B identifies the reported central business address in the output.
    data["Address Type"] = "B"

    # These output fields are unavailable in the current vm_vendors source.
    # No Building or invoice data is inferred from alternative sources.
    data["Building"] = ""
    data["Last Invoice Number"] = ""
    data["Last Transaction Date"] = ""
    data["Last Inv Amt Doc Currency"] = ""
    data[
        "Last Inv Amt Doc Currency Indicator"
    ] = ""

    po_box = data[
        "Street"
    ].map(
        lambda value: (
            bool(
                _POBOX_RE.search(
                    value
                )
            )
            if value != ""
            else False
        )
    )

    exceptions = data.loc[
        data["CoCo"].ne("")
        & data["Vendor Code"].ne("")
        & po_box
    ].copy()

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
        .loc[
            :,
            OUTPUT_COLUMNS,
        ]
        .reset_index(
            drop=True
        )
    )


def run_vm_016(
    context: dict[str, Any],
) -> dict[str, Any]:
    """Execute VM16 and replace only the VM16 result worksheet."""
    started = perf_counter()

    module_config = context.get(
        "module"
    )

    if not isinstance(
        module_config,
        dict,
    ):
        raise ValueError(
            "VM_016 requires context['module'] configuration."
        )

    companies = _configured_companies(
        module_config
    )

    # vm_vendors is loaded exactly once. The master and address population both
    # derive from this same validated DataFrame.
    vendor_source = load_vm_vendors(
        context
    )

    vendor_master = build_vendor_master_population(
        vendor_source
    )

    master_rows = len(
        vendor_master
    )

    filtered_master, excluded_company_rows = (
        _filter_configured_companies(
            vendor_master,
            companies,
        )
    )

    valid_population_input_rows = len(
        filtered_master
    )

    valid_population, metrics = (
        get_valid_vendor_population(
            filtered_master
        )
    )

    if valid_population.empty:
        raise ValueError(
            "VM_016: valid vendor population is empty after CONFIG "
            "company and common VM exclusion rules."
        )

    stage_started = _print_timing(
        "input load and population validation",
        started,
    )

    output = build_vm_016(
        valid_population
    )

    valid_street_rows = int(
        valid_population[
            "Street"
        ]
        .map(
            safe_text
        )
        .ne(
            ""
        )
        .sum()
    )

    stage_started = _print_timing(
        "preparation and analytic logic",
        stage_started,
    )

    output_file = write_vm_control_sheet(
        context=context,
        sheet_name=SHEET_NAME,
        dataframe=output,
        date_columns=[
            "Last Transaction Date",
        ],
        amount_columns=[
            "Last Inv Amt Doc Currency",
        ],
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
    # reconcile the population; input minus output is the net exclusion.
    net_common_exclusion = max(
        valid_population_input_rows
        - metrics.get(
            "output_rows",
            0,
        ),
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
        f"{CONTROL_ID} valid vendors with nonblank Street: "
        f"{valid_street_rows}"
    )

    print(
        "VM_016 Address definition: canonical LFA1 Street matches "
        "an approved VM_016 PO Box pattern"
    )

    print(
        "VM_016 Address source: LFA1 central address canonical "
        "columns in vm_vendors"
    )

    print(
        f"VM_016 valid vendors with PO Box only: "
        f"{len(output)}"
    )

    for warning in metrics.get(
        "warnings",
        [],
    ):
        print(
            f"WARNING: {warning}"
        )

    # The current run_analysis.py can report technical OK whenever no Python
    # exception occurs, even when this dictionary reports audit findings as
    # ERROR.
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
