"""
VM_002 - Vendors sharing the same address.

Objective
---------
Identify different vendors that share the same normalized address within the
same company.

The input workbook resolution, reading, header mapping, key validation and
population normalization are delegated exclusively to core.vm_common.
"""

from time import perf_counter
from typing import Any

import pandas as pd

from core.vm_common import (
    build_vendor_master_population,
    get_valid_vendor_population,
    load_vm_vendors,
    normalize_company,
    normalize_exact_key,
    safe_text,
    write_vm_control_sheet,
)


CONTROL_ID = "VM_002"
SHEET_NAME = "VM02"

IDENTITY_COLUMNS = [
    "Company",
    "Vendor Code",
]

ADDRESS_KEY_COLUMNS = [
    "Street",
    "City",
    "ZipCode",
    "Country",
]

OUTPUT_COLUMNS = [
    "CoCo",
    "Vendor Code",
    "Address Type",
    "Street",
    "City",
    "ZipCode",
    "State",
    "Country",
    "Group",
    "Company",
    "Vendor Name",
    "Last Invoice Number",
    "Last Transaction Date",
    "Last Inv Amt Doc Currency",
    "Last Inv Amt Doc Currency Indicator",
]


def _print_timing(
    stage_name: str,
    started: float,
) -> float:
    """Print one lightweight stage timing."""
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
    Return normalized company codes configured for VM.

    An empty value or ALL means that every company present in the validated
    vendor population is included.
    """
    module_config = context.get(
        "module"
    )

    if not isinstance(
        module_config,
        dict,
    ):
        raise ValueError(
            "VM_002 requires context['module'] configuration."
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
    Apply the VM CONFIG company population.

    Company filtering is performed after common normalization, so values such
    as 15, 0015 and 15.0 resolve consistently to 0015.
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


def _load_vm02_population(
    context: dict[str, Any],
) -> tuple[pd.DataFrame, dict[str, Any]]:
    """
    Load and validate the VM02 population.

    The vendor workbook is read exactly once through load_vm_vendors().
    """
    vendor_source = load_vm_vendors(
        context
    )

    vendor_master = build_vendor_master_population(
        vendor_source
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

    metrics = {
        "source_rows": len(vendor_source),
        "master_rows": len(vendor_master),
        "excluded_company": excluded_company,
        **population_metrics,
    }

    if valid_population.empty:
        raise ValueError(
            f"{CONTROL_ID}: valid vendor population is empty after "
            "CONFIG company and common VM exclusion rules."
        )

    return valid_population, metrics


def _normalize_address_columns(
    dataframe: pd.DataFrame,
) -> pd.DataFrame:
    """
    Create normalized address components once per unique source value.

    Unicode and regular-expression normalization are Python operations. Running
    them once per distinct value is cheaper than repeating them for every row.
    """
    normalized = pd.DataFrame(
        index=dataframe.index
    )

    for column in ADDRESS_KEY_COLUMNS:
        source = (
            dataframe[column]
            .astype("string")
            .fillna("")
        )

        unique_values = pd.Index(
            source.unique(),
            dtype="string",
        )

        lookup = pd.Series(
            (
                normalize_exact_key(value)
                for value in unique_values
            ),
            index=unique_values,
            dtype="string",
        )

        normalized[column] = (
            source.map(lookup)
            .fillna("")
        )

    return normalized


def _add_lha_display_columns(
    dataframe: pd.DataFrame,
) -> pd.DataFrame:
    """
    Map LBR vendor-master columns to the required LHA VM02 output layout.

    The current VM vendor extract does not contain invoice history, so the four
    last-invoice display columns remain blank. No unsupported transaction or FX
    inference is performed silently.
    """
    result = dataframe.copy()

    result["CoCo"] = result[
        "Company"
    ].map(normalize_company)

    result["Company"] = (
        result["Company Name"]
        .astype("string")
        .fillna("")
        .str.strip()
    )

    # ECC LFA1 provides one central vendor address rather than separate B/S
    # address rows. It is represented as B for compatibility with LHA output.
    result["Address Type"] = "B"

    result["Last Invoice Number"] = ""
    result["Last Transaction Date"] = ""
    result["Last Inv Amt Doc Currency"] = ""
    result[
        "Last Inv Amt Doc Currency Indicator"
    ] = ""

    return result


def build_vm_002(
    vendor_population: pd.DataFrame,
) -> pd.DataFrame:
    """
    Return vendors sharing the same address within the same company.

    Address key:
        normalized Street | City | ZipCode | Country

    Exception rule:
        at least two distinct Vendor Codes share the same nonblank address key
        within one company.
    """
    required_columns = {
        "Company",
        "Company Name",
        "Vendor Code",
        "Vendor Name",
        "Street",
        "City",
        "ZipCode",
        "State",
        "Country",
    }

    missing_columns = sorted(
        required_columns.difference(
            vendor_population.columns
        )
    )

    if missing_columns:
        raise ValueError(
            f"{CONTROL_ID}: vendor population is missing columns: "
            f"{missing_columns}."
        )

    if vendor_population.empty:
        return pd.DataFrame(
            columns=OUTPUT_COLUMNS
        )

    data = vendor_population.copy()

    duplicate_vendor = data.duplicated(
        subset=IDENTITY_COLUMNS,
        keep=False,
    )

    if duplicate_vendor.any():
        duplicated_keys = (
            data.loc[
                duplicate_vendor,
                IDENTITY_COLUMNS,
            ]
            .drop_duplicates()
            .head(20)
            .to_dict("records")
        )

        raise ValueError(
            f"{CONTROL_ID}: vendor master is not unique by "
            f"{IDENTITY_COLUMNS}. Examples: {duplicated_keys}"
        )

    normalized = _normalize_address_columns(
        data
    )

    address_key = (
        normalized["Street"]
        .str.cat(
            normalized["City"],
            sep="|",
        )
        .str.cat(
            normalized["ZipCode"],
            sep="|",
        )
        .str.cat(
            normalized["Country"],
            sep="|",
        )
    )

    address_content = address_key.str.replace(
        "|",
        "",
        regex=False,
    )

    data["_Address Key"] = (
        data["Company"]
        .map(normalize_company)
        .astype("string")
        .str.cat(
            address_key,
            sep="\u00a6",
        )
    )

    # LHA excludes an address only when Street, City, ZipCode and Country are
    # all empty after normalization.
    comparable = data.loc[
        address_content.ne("")
    ].copy()

    if comparable.empty:
        return pd.DataFrame(
            columns=OUTPUT_COLUMNS
        )

    distinct_vendor_count = (
        comparable.groupby(
            "_Address Key",
            sort=False,
            observed=True,
        )["Vendor Code"]
        .transform("nunique")
    )

    exceptions = comparable.loc[
        distinct_vendor_count.ge(2)
    ].copy()

    if exceptions.empty:
        return pd.DataFrame(
            columns=OUTPUT_COLUMNS
        )

    # Deterministic groups numbered from 1.
    exceptions["Group"] = (
        pd.factorize(
            exceptions["_Address Key"],
            sort=True,
        )[0]
        + 1
    )

    output = _add_lha_display_columns(
        exceptions
    )

    for column in OUTPUT_COLUMNS:
        if column not in output.columns:
            output[column] = ""

    return (
        output.sort_values(
            [
                "Group",
                "CoCo",
                "Vendor Code",
            ],
            kind="stable",
        )
        .loc[:, OUTPUT_COLUMNS]
        .reset_index(drop=True)
    )


def run_vm_002(
    context: dict[str, Any],
) -> dict[str, Any]:
    """Execute VM02 and replace only the VM02 result worksheet."""
    started = perf_counter()

    vendor_population, metrics = (
        _load_vm02_population(
            context
        )
    )

    stage_started = _print_timing(
        "input load and population validation",
        started,
    )

    output = build_vm_002(
        vendor_population
    )

    stage_started = _print_timing(
        "preparation and analytic logic",
        stage_started,
    )

    # VM02 performs no monetary conversion.
    stage_started = _print_timing(
        "FX conversion (not applicable)",
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
        integer_columns=[
            "Group",
        ],
    )

    _print_timing(
        "workbook write",
        stage_started,
    )

    print(f"{CONTROL_ID} source rows: {metrics['source_rows']}")
    print(f"{CONTROL_ID} master rows: {metrics['master_rows']}")
    print(
        f"{CONTROL_ID} rows excluded by CONFIG company: "
        f"{metrics['excluded_company']}"
    )
    print(
        f"{CONTROL_ID} valid population rows: "
        f"{metrics['output_rows']}"
    )
    print(f"{CONTROL_ID} exception rows: {len(output)}")

    for warning in metrics.get(
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
