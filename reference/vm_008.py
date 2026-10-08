"""
VM_008 - Vendors whose name matches a Logicalis employee name.

Logicalis employees are records in the VM vendor master whose Vendor Code
starts with F and whose Account Group is ZFUN. Employee records with central
or company deletion flags are excluded.

The supplier population is obtained through the common VM valid-population
rules. Matching is exact after name normalization and is restricted to the
same company.
"""

import re
import unicodedata
from time import perf_counter
from typing import Any

import pandas as pd

from core.vm_common import (
    build_vendor_master_population,
    get_valid_vendor_population,
    is_blank,
    load_vm_vendors,
    normalize_company,
    normalize_upper_text,
    normalize_vendor_code,
    safe_text,
    write_vm_control_sheet,
)


CONTROL_ID = "VM_008"
SHEET_NAME = "VM08"

EMPLOYEE_CODE_PREFIX = "F"
EMPLOYEE_ACCOUNT_GROUP = "ZFUN"

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

_EMPLOYEE_SOURCE_REQUIRED_COLUMNS = {
    "Company",
    "Vendor Code",
    "Vendor Name",
    "Account Group",
    "Central Deletion Flag",
    "Company Deletion Flag",
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
    """Print one execution-stage duration and return its completion time."""
    finished = perf_counter()

    print(
        f"{CONTROL_ID} {stage_name}: "
        f"{finished - started:.2f} seconds"
    )

    return finished


def _empty_output() -> pd.DataFrame:
    """Return an empty VM08 output with the exact required schema."""
    return pd.DataFrame(
        columns=OUTPUT_COLUMNS
    )


def _normalize_name(
    value: Any,
) -> str:
    """
    Build an exact-comparison name key.

    The key is uppercase and accent-free. Punctuation is replaced with spaces
    and consecutive whitespace is collapsed. No words or name components are
    removed and no approximate comparison is performed.
    """
    text = safe_text(
        value
    ).upper()

    decomposed = unicodedata.normalize(
        "NFKD",
        text,
    )

    without_diacritics = "".join(
        character
        for character in decomposed
        if not unicodedata.combining(
            character
        )
    )

    punctuation_as_spaces = "".join(
        character
        if (
            character.isalnum()
            or character.isspace()
        )
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
    """Normalize each distinct source name once and map it back."""
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
        source.map(
            lookup
        )
        .fillna("")
    )


def _configured_companies(
    context: dict[str, Any],
) -> set[str]:
    """
    Return normalized configured companies.

    An empty set means all companies. Empty values, None, ALL, *, TODAS and
    TODOS are interpreted as all companies.
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
            text.replace(
                ";",
                ",",
            )
            .replace(
                "|",
                ",",
            )
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
    """Filter a population by configured company codes."""
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

    return (
        dataframe.loc[
            included
        ]
        .copy()
        .reset_index(drop=True),
        int(
            (~included).sum()
        ),
    )


def _logicalis_employee_mask(
    vendor_master: pd.DataFrame,
) -> pd.Series:
    """
    Identify Logicalis employees by the functional F-code rule.

    Account Group is not used as an employee-identification condition because
    valid employee records may belong to an Account Group other than ZFUN.
    """
    required = {
        "Company",
        "Vendor Code",
        "Vendor Name",
        "Central Deletion Flag",
        "Company Deletion Flag",
    }

    missing = sorted(
        required.difference(
            vendor_master.columns
        )
    )

    if missing:
        raise ValueError(
            f"{CONTROL_ID}: employee source is missing columns: "
            f"{missing}."
        )

    vendor_codes = vendor_master[
        "Vendor Code"
    ].map(
        normalize_vendor_code
    )

    return vendor_codes.str.startswith(
        EMPLOYEE_CODE_PREFIX,
        na=False,
    )


def _exclude_employee_codes_from_suppliers(
    vendor_population: pd.DataFrame,
) -> tuple[pd.DataFrame, int]:
    """
    Remove every F-code record from the supplier population.

    The common valid-vendor rules exclude the usual employee Account Group,
    but F-code records belonging to another Account Group may survive those
    rules. VM08-VM11 must never treat an F-code record as a supplier.
    """
    if "Vendor Code" not in vendor_population.columns:
        raise ValueError(
            f"{CONTROL_ID}: supplier population is missing "
            "'Vendor Code'."
        )

    vendor_codes = vendor_population[
        "Vendor Code"
    ].map(
        normalize_vendor_code
    )

    employee_code = vendor_codes.str.startswith(
        EMPLOYEE_CODE_PREFIX,
        na=False,
    )

    return (
        vendor_population.loc[
            ~employee_code
        ]
        .copy()
        .reset_index(drop=True),
        int(
            employee_code.sum()
        ),
    )


def _build_employee_population(
    vendor_master: pd.DataFrame,
) -> tuple[pd.DataFrame, dict[str, int]]:
    """
    Build active Logicalis employees from the VM vendor master.

    Employee records with central or company deletion flags are excluded.
    """
    employee_mask = _logicalis_employee_mask(
        vendor_master
    )

    candidates = vendor_master.loc[
        employee_mask
    ].copy()

    central_deleted = candidates[
        "Central Deletion Flag"
    ].map(
        lambda value: not is_blank(
            value
        )
    )

    company_deleted = candidates[
        "Company Deletion Flag"
    ].map(
        lambda value: not is_blank(
            value
        )
    )

    deleted = (
        central_deleted
        | company_deleted
    )

    employees = candidates.loc[
        ~deleted,
        [
            "Company",
            "Vendor Code",
            "Vendor Name",
        ],
    ].copy()

    employees = employees.rename(
        columns={
            "Vendor Code": "Employee Code",
            "Vendor Name": "Employee Name",
        }
    )

    employees["Company"] = employees[
        "Company"
    ].map(
        normalize_company
    )

    employees["Employee Code"] = employees[
        "Employee Code"
    ].map(
        normalize_vendor_code
    )

    employees["Employee Name"] = employees[
        "Employee Name"
    ].map(
        safe_text
    )

    employees = (
        employees.drop_duplicates(
            subset=[
                "Company",
                "Employee Code",
                "Employee Name",
            ],
            keep="first",
        )
        .sort_values(
            [
                "Company",
                "Employee Code",
            ],
            kind="mergesort",
        )
        .reset_index(drop=True)
    )

    metrics = {
        "employee_candidate_rows": len(
            candidates
        ),
        "employee_excluded_central_deletion": int(
            central_deleted.sum()
        ),
        "employee_excluded_company_deletion": int(
            company_deleted.sum()
        ),
        "employee_excluded_any_deletion": int(
            deleted.sum()
        ),
        "employee_output_rows": len(
            employees
        ),
        "distinct_employees": employees[
            [
                "Company",
                "Employee Code",
            ]
        ].drop_duplicates().shape[0],
    }

    return (
        employees,
        metrics,
    )


def build_vm_008(
    vendor_population: pd.DataFrame,
    employee_population: pd.DataFrame,
) -> pd.DataFrame:
    """
    Return exact supplier/employee name matches within the same company.

    This pure analytical function does not read files, write workbooks or
    access the runner context.
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

    vendors["Company Name"] = vendors[
        "Company Name"
    ].map(
        safe_text
    )

    vendors["Vendor Code"] = vendors[
        "Vendor Code"
    ].map(
        normalize_vendor_code
    )

    vendors["Vendor Name"] = vendors[
        "Vendor Name"
    ].map(
        safe_text
    )

    vendors["_Normalized Name"] = _normalized_names(
        vendors[
            "Vendor Name"
        ]
    )

    employees["Company"] = employees[
        "Company"
    ].map(
        normalize_company
    )

    employees["Employee Code"] = employees[
        "Employee Code"
    ].map(
        normalize_vendor_code
    )

    employees["Employee Name"] = employees[
        "Employee Name"
    ].map(
        safe_text
    )

    employees["_Normalized Name"] = _normalized_names(
        employees[
            "Employee Name"
        ]
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

    matches = matches.loc[
        matches["Vendor Code"].ne(
            matches["Employee Code"]
        )
    ].copy()

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
        .loc[
            :,
            OUTPUT_COLUMNS,
        ]
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

    vendor_master_all_companies = (
        build_vendor_master_population(
            vendor_source
        )
    )

    vendor_master_rows = len(
        vendor_master_all_companies
    )

    all_company_employee_candidates = int(
        _logicalis_employee_mask(
            vendor_master_all_companies
        ).sum()
    )

    (
        vendor_master,
        excluded_company_rows,
    ) = _filter_companies(
        vendor_master_all_companies,
        companies,
    )

    configured_employee_candidates = int(
        _logicalis_employee_mask(
            vendor_master
        ).sum()
    )

    excluded_employee_company_rows = (
        all_company_employee_candidates
        - configured_employee_candidates
    )

    (
        employee_population,
        employee_metrics,
    ) = _build_employee_population(
        vendor_master
    )

    (
        vendor_population_before_employee_separation,
        vendor_metrics,
    ) = get_valid_vendor_population(
        vendor_master
    )

    (
        vendor_population,
        employee_codes_removed_from_suppliers,
    ) = _exclude_employee_codes_from_suppliers(
        vendor_population_before_employee_separation
    )

    if employee_population.empty:
        raise ValueError(
            f"{CONTROL_ID}: no active Logicalis employees were found "
            f"using Vendor Code prefix {EMPLOYEE_CODE_PREFIX!r} after "
            "applying configured-company and deletion-flag rules."
        )

    if vendor_population.empty:
        raise ValueError(
            f"{CONTROL_ID}: valid supplier population is empty after "
            "the configured company and common VM exclusion rules."
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
        f"{CONTROL_ID} rows excluded by CONFIG company: "
        f"{excluded_company_rows}"
    )

    print(
        f"{CONTROL_ID} valid population before F separation: "
        f"{vendor_metrics.get('output_rows', len(vendor_population_before_employee_separation))}"
    )

    print(
        f"{CONTROL_ID} F-code rows removed from suppliers: "
        f"{employee_codes_removed_from_suppliers}"
    )

    print(
        f"{CONTROL_ID} valid supplier rows: "
        f"{len(vendor_population)}"
    )

    print(
        f"{CONTROL_ID} employee identification rule: "
        f"Vendor Code starts with {EMPLOYEE_CODE_PREFIX!r}"
    )

    print(
        f"{CONTROL_ID} employee candidate rows before CONFIG company: "
        f"{all_company_employee_candidates}"
    )

    print(
        f"{CONTROL_ID} employee rows excluded by CONFIG company: "
        f"{excluded_employee_company_rows}"
    )

    print(
        f"{CONTROL_ID} employee candidate rows after CONFIG company: "
        f"{employee_metrics.get('employee_candidate_rows', 0)}"
    )

    print(
        f"{CONTROL_ID} employees excluded by Central Deletion Flag: "
        f"{employee_metrics.get('employee_excluded_central_deletion', 0)}"
    )

    print(
        f"{CONTROL_ID} employees excluded by Company Deletion Flag: "
        f"{employee_metrics.get('employee_excluded_company_deletion', 0)}"
    )

    print(
        f"{CONTROL_ID} employees excluded by any deletion flag: "
        f"{employee_metrics.get('employee_excluded_any_deletion', 0)}"
    )

    print(
        f"{CONTROL_ID} active employee rows: "
        f"{employee_metrics.get('employee_output_rows', len(employee_population))}"
    )

    print(
        f"{CONTROL_ID} distinct active employees: "
        f"{employee_metrics.get('distinct_employees', 0)}"
    )

    print(
        f"{CONTROL_ID} exception rows: "
        f"{len(output)}"
    )

    for warning in vendor_metrics.get(
        "warnings",
        [],
    ):
        print(
            f"WARNING: {warning}"
        )

    return {
        "status": "ERROR" if not output.empty else "OK",
        "output_file": output_file,
        "sheet_name": SHEET_NAME,
        "rows": len(output),
    }
