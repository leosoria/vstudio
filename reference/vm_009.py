"""
VM_009 - Vendors sharing an address with a Logicalis employee.

Objective
---------
Identify valid suppliers whose normalized address is identical to the
normalized address of a Logicalis employee in the same company.

Logicalis employees are records in the VM vendor master whose Vendor Code
starts with F and whose Account Group is ZFUN. Employee records with central
or company deletion flags are excluded.

The comparison uses the exact normalized combination:

    Company + Street + City + ZipCode + Country

State is retained for presentation but does not form part of the comparison
key. All supplier/employee combinations are preserved in one output sheet and
classified by address concentration for audit review.
"""

from time import perf_counter
from typing import Any

import pandas as pd

from core.vm_common import (
    build_vendor_master_population,
    get_valid_vendor_population,
    is_blank,
    load_vm_vendors,
    normalize_company,
    normalize_exact_key,
    normalize_upper_text,
    normalize_vendor_code,
    safe_text,
    write_vm_control_sheet,
)


CONTROL_ID = "VM_009"
SHEET_NAME = "VM09"

EMPLOYEE_CODE_PREFIX = "F"
EMPLOYEE_ACCOUNT_GROUP = "ZFUN"

MASSIVE_EMPLOYEE_THRESHOLD = 10

ADDRESS_COLUMNS = [
    "Street",
    "City",
    "ZipCode",
    "Country",
]

OUTPUT_COLUMNS = [
    "Address Group",
    "Company",
    "CoCo",
    "Vendor Code",
    "Vendor Name",
    "Vendor Street",
    "Vendor City",
    "Vendor ZipCode",
    "Vendor State",
    "Vendor Country",
    "Employee Code",
    "Employee Name",
    "Employee Street",
    "Employee City",
    "Employee ZipCode",
    "Employee State",
    "Employee Country",
    "Suppliers at Address",
    "Employees at Address",
    "Match Rows at Address",
    "Address Concentration",
    "Potential Corporate Address",
    "Audit Treatment",
]

_VENDOR_REQUIRED_COLUMNS = {
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

_EMPLOYEE_REQUIRED_COLUMNS = {
    "Company",
    "Employee Code",
    "Employee Name",
    "Street",
    "City",
    "ZipCode",
    "State",
    "Country",
}

_EMPLOYEE_SOURCE_REQUIRED_COLUMNS = {
    "Company",
    "Vendor Code",
    "Vendor Name",
    "Account Group",
    "Central Deletion Flag",
    "Company Deletion Flag",
    "Street",
    "City",
    "ZipCode",
    "State",
    "Country",
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
    """Return an empty VM09 output with the exact required schema."""
    return pd.DataFrame(
        columns=OUTPUT_COLUMNS
    )


def _configured_companies(
    context: dict[str, Any],
) -> set[str]:
    """
    Return normalized configured company codes.

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
            "Street",
            "City",
            "ZipCode",
            "State",
            "Country",
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

    for column in [
        "Street",
        "City",
        "ZipCode",
        "State",
        "Country",
    ]:
        employees[column] = employees[
            column
        ].map(
            safe_text
        )

    employees = (
        employees.drop_duplicates(
            subset=[
                "Company",
                "Employee Code",
                "Employee Name",
                "Street",
                "City",
                "ZipCode",
                "State",
                "Country",
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


def _prepare_addresses(
    dataframe: pd.DataFrame,
    *,
    code_column: str,
    name_column: str,
) -> tuple[pd.DataFrame, dict[str, int]]:
    """
    Normalize complete addresses for exact comparison.

    All four comparison components must be nonblank after normalization.
    State is retained for presentation but is not part of the address key.
    """
    required = {
        "Company",
        code_column,
        name_column,
        "Street",
        "City",
        "ZipCode",
        "State",
        "Country",
    }

    missing = sorted(
        required.difference(
            dataframe.columns
        )
    )

    if missing:
        raise ValueError(
            f"{CONTROL_ID}: address population for {code_column!r} "
            f"is missing columns: {missing}."
        )

    result = dataframe.loc[
        :,
        [
            "Company",
            code_column,
            name_column,
            "Street",
            "City",
            "ZipCode",
            "State",
            "Country",
        ],
    ].copy()

    result["Company"] = result[
        "Company"
    ].map(
        normalize_company
    )

    result[code_column] = result[
        code_column
    ].map(
        normalize_vendor_code
    )

    result[name_column] = result[
        name_column
    ].map(
        safe_text
    )

    for column in [
        "Street",
        "City",
        "ZipCode",
        "State",
        "Country",
    ]:
        result[column] = result[
            column
        ].map(
            safe_text
        )

    normalized_columns = []

    for column in ADDRESS_COLUMNS:
        normalized_column = (
            f"_Normalized {column}"
        )

        result[normalized_column] = result[
            column
        ].map(
            normalize_exact_key
        )

        normalized_columns.append(
            normalized_column
        )

    valid_identity = (
        result["Company"].ne("")
        & result[code_column].ne("")
    )

    complete_address = pd.Series(
        True,
        index=result.index,
        dtype=bool,
    )

    for column in normalized_columns:
        complete_address &= result[
            column
        ].ne("")

    comparable = result.loc[
        valid_identity
        & complete_address
    ].copy()

    comparable["_Address Key"] = (
        comparable[
            "_Normalized Street"
        ]
        + "|"
        + comparable[
            "_Normalized City"
        ]
        + "|"
        + comparable[
            "_Normalized ZipCode"
        ]
        + "|"
        + comparable[
            "_Normalized Country"
        ]
    )

    comparable = (
        comparable.drop_duplicates(
            subset=[
                "Company",
                code_column,
                "_Address Key",
            ],
            keep="first",
        )
        .reset_index(drop=True)
    )

    metrics = {
        "input_rows": len(
            result
        ),
        "invalid_identity_rows": int(
            (~valid_identity).sum()
        ),
        "incomplete_address_rows": int(
            (
                valid_identity
                & ~complete_address
            ).sum()
        ),
        "comparable_rows": len(
            comparable
        ),
        "distinct_entities": comparable[
            [
                "Company",
                code_column,
            ]
        ].drop_duplicates().shape[0],
        "distinct_addresses": comparable[
            [
                "Company",
                "_Address Key",
            ]
        ].drop_duplicates().shape[0],
    }

    return (
        comparable,
        metrics,
    )


def _address_concentration(
    employee_count: pd.Series,
) -> pd.Series:
    """Classify address concentration without excluding any result."""
    result = pd.Series(
        "INDIVIDUAL",
        index=employee_count.index,
        dtype=object,
    )

    result.loc[
        employee_count.between(
            2,
            MASSIVE_EMPLOYEE_THRESHOLD - 1,
            inclusive="both",
        )
    ] = "SHARED"

    result.loc[
        employee_count.ge(
            MASSIVE_EMPLOYEE_THRESHOLD
        )
    ] = "MASSIVE"

    return result


def build_vm_009(
    vendor_population: pd.DataFrame,
    employee_population: pd.DataFrame,
) -> pd.DataFrame:
    """
    Return exact supplier/employee address matches within the same company.

    All matches are preserved. Address concentration and audit-treatment
    columns differentiate massive, shared and individual address groups.
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
            "Street",
            "City",
            "ZipCode",
            "State",
            "Country",
        ],
    ].copy()

    vendors["Company Name"] = vendors[
        "Company Name"
    ].map(
        safe_text
    )

    (
        vendor_addresses,
        _,
    ) = _prepare_addresses(
        vendors,
        code_column="Vendor Code",
        name_column="Vendor Name",
    )

    (
        employee_addresses,
        _,
    ) = _prepare_addresses(
        employee_population,
        code_column="Employee Code",
        name_column="Employee Name",
    )

    if (
        vendor_addresses.empty
        or employee_addresses.empty
    ):
        return _empty_output()

    supplier_counts = (
        vendor_addresses.groupby(
            [
                "Company",
                "_Address Key",
            ],
            sort=False,
            observed=True,
        )["Vendor Code"]
        .nunique()
        .rename(
            "Suppliers at Address"
        )
        .reset_index()
    )

    employee_counts = (
        employee_addresses.groupby(
            [
                "Company",
                "_Address Key",
            ],
            sort=False,
            observed=True,
        )["Employee Code"]
        .nunique()
        .rename(
            "Employees at Address"
        )
        .reset_index()
    )

    matches = vendor_addresses.merge(
        employee_addresses[
            [
                "Company",
                "Employee Code",
                "Employee Name",
                "Street",
                "City",
                "ZipCode",
                "State",
                "Country",
                "_Address Key",
            ]
        ],
        on=[
            "Company",
            "_Address Key",
        ],
        how="inner",
        suffixes=(
            " Vendor",
            " Employee",
        ),
        validate="many_to_many",
    )

    matches = matches.loc[
        matches["Vendor Code"].ne(
            matches["Employee Code"]
        )
    ].copy()

    matches = matches.drop_duplicates(
        subset=[
            "Company",
            "Vendor Code",
            "Employee Code",
            "_Address Key",
        ],
        keep="first",
    )

    if matches.empty:
        return _empty_output()

    matches = matches.merge(
        supplier_counts,
        on=[
            "Company",
            "_Address Key",
        ],
        how="left",
        validate="many_to_one",
    )

    matches = matches.merge(
        employee_counts,
        on=[
            "Company",
            "_Address Key",
        ],
        how="left",
        validate="many_to_one",
    )

    matches["Match Rows at Address"] = (
        matches.groupby(
            [
                "Company",
                "_Address Key",
            ],
            sort=False,
            observed=True,
        )["Vendor Code"]
        .transform(
            "size"
        )
        .astype(int)
    )

    address_group_key = (
        matches["Company"]
        + "|"
        + matches["_Address Key"]
    )

    matches["Address Group"] = (
        pd.factorize(
            address_group_key,
            sort=True,
        )[0]
        + 1
    )

    matches["Address Concentration"] = (
        _address_concentration(
            matches[
                "Employees at Address"
            ]
        )
    )

    matches["Potential Corporate Address"] = (
        matches[
            "Address Concentration"
        ]
        .map(
            {
                "MASSIVE": "YES",
                "SHARED": "REVIEW",
                "INDIVIDUAL": "NO",
            }
        )
    )

    matches["Audit Treatment"] = (
        matches[
            "Address Concentration"
        ]
        .map(
            {
                "MASSIVE": (
                    "PENDING CORPORATE ADDRESS VALIDATION"
                ),
                "SHARED": (
                    "PENDING CORPORATE ADDRESS VALIDATION"
                ),
                "INDIVIDUAL": (
                    "PENDING INDIVIDUAL REVIEW"
                ),
            }
        )
    )

    company_names = vendors.loc[
        :,
        [
            "Company",
            "Company Name",
        ],
    ].copy()

    company_names["Company"] = company_names[
        "Company"
    ].map(
        normalize_company
    )

    company_names["Company Name"] = company_names[
        "Company Name"
    ].map(
        safe_text
    )

    company_name_counts = (
        company_names.loc[
            company_names[
                "Company Name"
            ].ne("")
        ]
        .groupby(
            "Company",
            sort=False,
            observed=True,
        )["Company Name"]
        .nunique()
    )

    conflicting_company_names = (
        company_name_counts.loc[
            company_name_counts.gt(
                1
            )
        ]
    )

    if not conflicting_company_names.empty:
        raise ValueError(
            f"{CONTROL_ID}: multiple Company Name values were found "
            f"for the same Company: "
            f"{conflicting_company_names.to_dict()}."
        )

    company_names = (
        company_names.sort_values(
            [
                "Company",
                "Company Name",
            ],
            kind="mergesort",
        )
        .drop_duplicates(
            subset=[
                "Company",
            ],
            keep="first",
        )
        .reset_index(drop=True)
    )

    matches = (
        matches.drop(
            columns=[
                "Company Name",
            ],
            errors="ignore",
        )
        .merge(
            company_names,
            on="Company",
            how="left",
            validate="many_to_one",
        )
    )

    output = pd.DataFrame(
        {
            "Address Group": matches[
                "Address Group"
            ],
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
            "Vendor Street": matches[
                "Street Vendor"
            ],
            "Vendor City": matches[
                "City Vendor"
            ],
            "Vendor ZipCode": matches[
                "ZipCode Vendor"
            ],
            "Vendor State": matches[
                "State Vendor"
            ],
            "Vendor Country": matches[
                "Country Vendor"
            ],
            "Employee Code": matches[
                "Employee Code"
            ],
            "Employee Name": matches[
                "Employee Name"
            ],
            "Employee Street": matches[
                "Street Employee"
            ],
            "Employee City": matches[
                "City Employee"
            ],
            "Employee ZipCode": matches[
                "ZipCode Employee"
            ],
            "Employee State": matches[
                "State Employee"
            ],
            "Employee Country": matches[
                "Country Employee"
            ],
            "Suppliers at Address": matches[
                "Suppliers at Address"
            ],
            "Employees at Address": matches[
                "Employees at Address"
            ],
            "Match Rows at Address": matches[
                "Match Rows at Address"
            ],
            "Address Concentration": matches[
                "Address Concentration"
            ],
            "Potential Corporate Address": matches[
                "Potential Corporate Address"
            ],
            "Audit Treatment": matches[
                "Audit Treatment"
            ],
        }
    )

    output = output.drop_duplicates(
        subset=[
            "CoCo",
            "Vendor Code",
            "Employee Code",
            "Address Group",
        ],
        keep="first",
    )

    concentration_priority = {
        "MASSIVE": 1,
        "SHARED": 2,
        "INDIVIDUAL": 3,
    }

    output["_Concentration Priority"] = (
        output[
            "Address Concentration"
        ]
        .map(
            concentration_priority
        )
        .fillna(99)
    )

    return (
        output.sort_values(
            [
                "_Concentration Priority",
                "Match Rows at Address",
                "Address Group",
                "Company",
                "CoCo",
                "Vendor Name",
                "Vendor Code",
                "Employee Name",
                "Employee Code",
            ],
            ascending=[
                True,
                False,
                True,
                True,
                True,
                True,
                True,
                True,
                True,
            ],
            kind="mergesort",
        )
        .loc[
            :,
            OUTPUT_COLUMNS,
        ]
        .reset_index(drop=True)
    )


def run_vm_009(
    context: dict[str, Any],
) -> dict[str, Any]:
    """Execute VM09 and replace only the VM09 result worksheet."""
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

    (
        comparable_vendors,
        vendor_address_metrics,
    ) = _prepare_addresses(
        vendor_population,
        code_column="Vendor Code",
        name_column="Vendor Name",
    )

    (
        comparable_employees,
        employee_address_metrics,
    ) = _prepare_addresses(
        employee_population,
        code_column="Employee Code",
        name_column="Employee Name",
    )

    stage_started = _print_timing(
        "input load and population validation",
        started,
    )

    output = build_vm_009(
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
        integer_columns=[
            "Address Group",
            "Suppliers at Address",
            "Employees at Address",
            "Match Rows at Address",
        ],
    )

    _print_timing(
        "workbook write",
        stage_started,
    )

    module_config = context.get(
        "module",
        {},
    )

    distinct_matching_addresses = (
        output[
            [
                "CoCo",
                "Address Group",
            ]
        ]
        .drop_duplicates()
        .shape[0]
        if not output.empty
        else 0
    )

    massive_addresses = (
        output.loc[
            output[
                "Address Concentration"
            ].eq(
                "MASSIVE"
            ),
            [
                "CoCo",
                "Address Group",
            ],
        ]
        .drop_duplicates()
        .shape[0]
        if not output.empty
        else 0
    )

    shared_addresses = (
        output.loc[
            output[
                "Address Concentration"
            ].eq(
                "SHARED"
            ),
            [
                "CoCo",
                "Address Group",
            ],
        ]
        .drop_duplicates()
        .shape[0]
        if not output.empty
        else 0
    )

    individual_addresses = (
        output.loc[
            output[
                "Address Concentration"
            ].eq(
                "INDIVIDUAL"
            ),
            [
                "CoCo",
                "Address Group",
            ],
        ]
        .drop_duplicates()
        .shape[0]
        if not output.empty
        else 0
    )

    massive_match_rows = int(
        output[
            "Address Concentration"
        ].eq(
            "MASSIVE"
        ).sum()
    )

    pending_corporate_rows = int(
        output[
            "Address Concentration"
        ].isin(
            [
                "MASSIVE",
                "SHARED",
            ]
        ).sum()
    )

    largest_employee_count = (
        int(
            output[
                "Employees at Address"
            ].max()
        )
        if not output.empty
        else 0
    )

    largest_match_rows = (
        int(
            output[
                "Match Rows at Address"
            ].max()
        )
        if not output.empty
        else 0
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
        f"{CONTROL_ID} active employee rows: "
        f"{employee_metrics.get('employee_output_rows', len(employee_population))}"
    )

    print(
        f"{CONTROL_ID} distinct active employees: "
        f"{employee_metrics.get('distinct_employees', 0)}"
    )

    print(
        f"{CONTROL_ID} suppliers with complete address: "
        f"{vendor_address_metrics.get('comparable_rows', len(comparable_vendors))}"
    )

    print(
        f"{CONTROL_ID} suppliers excluded by incomplete address: "
        f"{vendor_address_metrics.get('incomplete_address_rows', 0)}"
    )

    print(
        f"{CONTROL_ID} employees with complete address: "
        f"{employee_address_metrics.get('comparable_rows', len(comparable_employees))}"
    )

    print(
        f"{CONTROL_ID} employees excluded by incomplete address: "
        f"{employee_address_metrics.get('incomplete_address_rows', 0)}"
    )

    print(
        f"{CONTROL_ID} distinct matching addresses: "
        f"{distinct_matching_addresses}"
    )

    print(
        f"{CONTROL_ID} massive matching addresses: "
        f"{massive_addresses}"
    )

    print(
        f"{CONTROL_ID} shared matching addresses: "
        f"{shared_addresses}"
    )

    print(
        f"{CONTROL_ID} individual matching addresses: "
        f"{individual_addresses}"
    )

    print(
        f"{CONTROL_ID} massive address match rows: "
        f"{massive_match_rows}"
    )

    print(
        f"{CONTROL_ID} rows pending corporate address validation: "
        f"{pending_corporate_rows}"
    )

    print(
        f"{CONTROL_ID} largest address employee count: "
        f"{largest_employee_count}"
    )

    print(
        f"{CONTROL_ID} largest address match rows: "
        f"{largest_match_rows}"
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
