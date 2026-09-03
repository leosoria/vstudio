"""
VM_008 - Vendors whose name matches an active employee name.

Objective
---------
Identify vendors whose normalized name is identical to the normalized name of
an active employee belonging to the same company.

Vendor and employee file discovery, source validation, SAP identifier
normalization, employee-period validation and common vendor exclusions are
delegated to core.vm_common.
"""

import re
import unicodedata
from time import perf_counter
from typing import Any

import pandas as pd

from core.vm_common import (
    build_vendor_master_population,
    get_valid_vendor_population,
    load_vm_employees,
    load_vm_vendors,
    normalize_company,
    normalize_employee_code,
    normalize_vendor_code,
    safe_text,
    write_vm_control_sheet,
)


CONTROL_ID = "VM_008"
SHEET_NAME = "VM08"

OUTPUT_COLUMNS = [
    "Company",
    "CoCo",
    "Vendor Code",
    "Vendor Name",
    "Employee Code",
    "Employee Name",
]

_VENDOR_REQUIRED_COLUMNS = {
    "Company",
    "Company Name",
    "Vendor Code",
    "Vendor Name",
}

_EMPLOYEE_REQUIRED_COLUMNS = {
    "Company",
    "Employee Code",
    "Employee Name",
}

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
    """Print one lightweight stage timing and return its completion time."""
    finished = perf_counter()

    print(
        f"{CONTROL_ID} {stage_name}: "
        f"{finished - started:.2f} seconds"
    )

    return finished


def _normalize_name(
    value: Any,
) -> str:
    """
    Build an exact-comparison name key.

    The key is uppercase, accent-free and contains only letters, numbers and
    single spaces. Punctuation is replaced with spaces. No name components,
    corporate words or other tokens are removed.
    """
    text = safe_text(value).upper()
    decomposed = unicodedata.normalize(
        "NFKD",
        text,
    )

    without_diacritics = "".join(
        character
        for character in decomposed
        if not unicodedata.combining(character)
    )

    punctuation_as_spaces = "".join(
        character
        if character.isalnum() or character.isspace()
        else " "
        for character in without_diacritics
    )

    return re.sub(
        r"\s+",
        " ",
        punctuation_as_spaces,
    ).strip()


def _normalized_names(
    series: pd.Series,
) -> pd.Series:
    """
    Normalize each distinct source name once and map it back to the population.

    This avoids repeating Python Unicode normalization for every row while
    retaining a vectorized pandas comparison and merge.
    """
    source = (
        series.astype("string")
        .fillna("")
    )

    unique_values = pd.Index(
        source.unique(),
        dtype="string",
    )

    lookup = pd.Series(
        (
            _normalize_name(value)
            for value in unique_values
        ),
        index=unique_values,
        dtype="string",
    )

    return (
        source.map(lookup)
        .fillna("")
    )


def _configured_companies(
    context: dict[str, Any],
) -> set[str]:
    """
    Return normalized company codes configured for VM.

    An empty set means every company is included. Empty values, ALL, *,
    TODAS and TODOS are treated as all companies.
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
            or text.upper() in _ALL_COMPANIES
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
        _ALL_COMPANIES
    ):
        return set()

    return normalized_values


def _filter_companies(
    dataframe: pd.DataFrame,
    companies: set[str],
) -> tuple[pd.DataFrame, int]:
    """
    Apply the CONFIG company filter.

    Return the filtered population and the number of rows excluded by the
    configured companies.
    """
    if not companies:
        return (
            dataframe.copy()
            .reset_index(drop=True),
            0,
        )

    company_values = dataframe[
        "Company"
    ].map(
        normalize_company
    )

    included = company_values.isin(
        companies
    )

    excluded_rows = int(
        (~included).sum()
    )

    return (
        dataframe.loc[included]
        .copy()
        .reset_index(drop=True),
        excluded_rows,
    )


def _empty_output() -> pd.DataFrame:
    """Return an empty VM08 output with the exact required schema."""
    return pd.DataFrame(
        columns=OUTPUT_COLUMNS
    )


def build_vm_008(
    vendor_population: pd.DataFrame,
    employee_population: pd.DataFrame,
) -> pd.DataFrame:
    """
    Return vendors matching active employees by company and normalized name.

    This is the pure analytical portion of VM_008. It does not read files,
    write workbooks or use the runner context.

    Parameters
    ----------
    vendor_population:
        Valid vendor population after common VM exclusions.
    employee_population:
        Active employee population whose validity overlaps the configured
        period.

    Returns
    -------
    pandas.DataFrame
        One row per unique CoCo + Vendor Code + Employee Code, containing
        exactly OUTPUT_COLUMNS.
    """
    missing_vendor_columns = sorted(
        _VENDOR_REQUIRED_COLUMNS.difference(
            vendor_population.columns
        )
    )

    missing_employee_columns = sorted(
        _EMPLOYEE_REQUIRED_COLUMNS.difference(
            employee_population.columns
        )
    )

    if (
        missing_vendor_columns
        or missing_employee_columns
    ):
        raise ValueError(
            f"{CONTROL_ID}: "
            f"missing vendor columns: {missing_vendor_columns}; "
            f"missing employee columns: {missing_employee_columns}."
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

    employees = employee_population.loc[
        :,
        [
            "Company",
            "Employee Code",
            "Employee Name",
        ],
    ].copy()

    vendors["Company"] = vendors[
        "Company"
    ].map(
        normalize_company
    )

    vendors["Vendor Code"] = vendors[
        "Vendor Code"
    ].map(
        normalize_vendor_code
    )

    vendors["Company Name"] = vendors[
        "Company Name"
    ].map(
        safe_text
    )

    vendors["Vendor Name"] = vendors[
        "Vendor Name"
    ].map(
        safe_text
    )

    vendors["_Normalized Name"] = _normalized_names(
        vendors["Vendor Name"]
    )

    employees["Company"] = employees[
        "Company"
    ].map(
        normalize_company
    )

    employees["Employee Code"] = employees[
        "Employee Code"
    ].map(
        normalize_employee_code
    )

    employees["Employee Name"] = employees[
        "Employee Name"
    ].map(
        safe_text
    )

    employees["_Normalized Name"] = _normalized_names(
        employees["Employee Name"]
    )

    comparable_vendors = vendors.loc[
        vendors["Company"].ne("")
        & vendors["Vendor Code"].ne("")
        & vendors["_Normalized Name"].ne("")
    ].copy()

    comparable_employees = employees.loc[
        employees["Company"].ne("")
        & employees["Employee Code"].ne("")
        & employees["_Normalized Name"].ne("")
    ].copy()

    comparable_vendors = (
        comparable_vendors.drop_duplicates(
            subset=[
                "Company",
                "Company Name",
                "Vendor Code",
                "Vendor Name",
                "_Normalized Name",
            ],
            keep="first",
        )
        .reset_index(drop=True)
    )

    comparable_employees = (
        comparable_employees.drop_duplicates(
            subset=[
                "Company",
                "Employee Code",
                "Employee Name",
                "_Normalized Name",
            ],
            keep="first",
        )
        .reset_index(drop=True)
    )

    if (
        comparable_vendors.empty
        or comparable_employees.empty
    ):
        return _empty_output()

    matches = comparable_vendors.merge(
        comparable_employees,
        on=[
            "Company",
            "_Normalized Name",
        ],
        how="inner",
        validate="many_to_many",
    )

    if matches.empty:
        return _empty_output()

    output = pd.DataFrame(
        {
            "Company": matches[
                "Company Name"
            ],
            "CoCo": matches[
                "Company"
            ],
            "Vendor Code": matches[
                "Vendor Code"
            ],
            "Vendor Name": matches[
                "Vendor Name"
            ],
            "Employee Code": matches[
                "Employee Code"
            ],
            "Employee Name": matches[
                "Employee Name"
            ],
        }
    )

    output = output.drop_duplicates(
        subset=[
            "CoCo",
            "Vendor Code",
            "Employee Code",
        ],
        keep="first",
    )

    return (
        output.sort_values(
            [
                "Company",
                "CoCo",
                "Vendor Name",
                "Vendor Code",
                "Employee Name",
                "Employee Code",
            ],
            kind="mergesort",
        )
        .loc[:, OUTPUT_COLUMNS]
        .reset_index(drop=True)
    )


def run_vm_008(
    context: dict[str, Any],
) -> dict[str, Any]:
    """Execute VM08 and replace only the VM08 result worksheet."""
    started = perf_counter()

    companies = _configured_companies(
        context
    )

    vendor_source = load_vm_vendors(
        context
    )

    vendor_master = build_vendor_master_population(
        vendor_source
    )

    vendor_master_rows = len(
        vendor_master
    )

    (
        vendor_master,
        excluded_vendor_company_rows,
    ) = _filter_companies(
        vendor_master,
        companies,
    )

    (
        vendor_population,
        vendor_metrics,
    ) = get_valid_vendor_population(
        vendor_master
    )

    (
        employee_population,
        employee_metrics,
    ) = load_vm_employees(
        context
    )

    if (
        employee_population is None
        or not employee_metrics.get(
            "available",
            False,
        )
    ):
        raise FileNotFoundError(
            f"{CONTROL_ID} requires the employee file for the configured "
            "period, but the expected VM employee workbook was not found."
        )

    (
        employee_population,
        excluded_employee_company_rows,
    ) = _filter_companies(
        employee_population,
        companies,
    )

    stage_started = _print_timing(
        "input load and population validation",
        started,
    )

    output = build_vm_008(
        vendor_population,
        employee_population,
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

    module_config = context.get(
        "module",
        {},
    )

    print(
        f"{CONTROL_ID} period FROM/TO: "
        f"{safe_text(module_config.get('from', ''))} / "
        f"{safe_text(module_config.get('to', ''))}"
    )

    print(
        f"{CONTROL_ID} vendor source rows: "
        f"{len(vendor_source)}"
    )

    print(
        f"{CONTROL_ID} vendor master rows: "
        f"{vendor_master_rows}"
    )

    print(
        f"{CONTROL_ID} vendor rows excluded by CONFIG company: "
        f"{excluded_vendor_company_rows}"
    )

    print(
        f"{CONTROL_ID} valid vendor rows: "
        f"{vendor_metrics.get('output_rows', len(vendor_population))}"
    )

    print(
        f"{CONTROL_ID} employee source rows: "
        f"{employee_metrics.get('source_rows', 0)}"
    )

    print(
        f"{CONTROL_ID} active period employee rows: "
        f"{employee_metrics.get('valid_period_rows', 0)}"
    )

    print(
        f"{CONTROL_ID} distinct employees: "
        f"{employee_metrics.get('distinct_employees', 0)}"
    )

    print(
        f"{CONTROL_ID} employee rows excluded by CONFIG company: "
        f"{excluded_employee_company_rows}"
    )

    print(
        f"{CONTROL_ID} exception rows: "
        f"{len(output)}"
    )

    warnings = [
        *vendor_metrics.get(
            "warnings",
            [],
        ),
        *employee_metrics.get(
            "warnings",
            [],
        ),
    ]

    for warning in warnings:
        print(
            f"WARNING: {warning}"
        )

    return {
        "status": "ERROR" if not output.empty else "OK",
        "output_file": output_file,
        "sheet_name": SHEET_NAME,
        "rows": len(output),
    }
