"""
VM_011 - Vendors sharing a bank account with a Logicalis employee.

Objective
---------
Identify valid suppliers whose complete normalized bank combination is
identical to that of an active Logicalis employee in the same company.

The comparison key is:

    Company + Bank Country + Bank Code + Bank Account

Logicalis employees are VM vendor-master records whose Vendor Code starts
with F. Employee records with central or company deletion flags are excluded.

All complete matches are preserved. Concentration and bank-account quality
columns distinguish individual matches, shared keys, massive keys and
possible textual placeholder accounts.
"""

import re
from time import perf_counter
from typing import Any

import pandas as pd

from core.vm_common import (
    build_vendor_bank_population,
    build_vendor_master_population,
    get_valid_vendor_population,
    is_blank,
    load_vm_vendors,
    normalize_company,
    normalize_exact_key,
    normalize_vendor_code,
    safe_text,
    write_vm_control_sheet,
)


CONTROL_ID = "VM_011"
SHEET_NAME = "VM11"

EMPLOYEE_CODE_PREFIX = "F"
MASSIVE_ENTITY_THRESHOLD = 10

OUTPUT_COLUMNS = [
    "Bank Group",
    "Company",
    "CoCo",
    "Vendor Code",
    "Vendor Name",
    "Vendor Bank Country",
    "Vendor Bank Code",
    "Vendor Bank Account",
    "Vendor Account Holder Name",
    "Vendor Bank Valid From",
    "Vendor Bank Valid To",
    "Employee Code",
    "Employee Name",
    "Employee Bank Country",
    "Employee Bank Code",
    "Employee Bank Account",
    "Employee Account Holder Name",
    "Employee Bank Valid From",
    "Employee Bank Valid To",
    "Suppliers at Bank",
    "Employees at Bank",
    "Match Rows at Bank",
    "Bank Concentration",
    "Bank Account Quality",
    "Potential Placeholder Account",
    "Audit Treatment",
]

_BANK_DETAIL_COLUMNS = [
    "Bank Country",
    "Bank Code",
    "Bank Account",
    "Account Holder Name",
    "Bank Valid From",
    "Bank Valid To",
]

_VENDOR_REQUIRED_COLUMNS = {
    "Company",
    "Company Name",
    "Vendor Code",
    "Vendor Name",
    *_BANK_DETAIL_COLUMNS,
}

_EMPLOYEE_REQUIRED_COLUMNS = {
    "Company",
    "Employee Code",
    "Employee Name",
    *_BANK_DETAIL_COLUMNS,
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
    """Return an empty VM11 output with the exact required schema."""
    return pd.DataFrame(
        columns=OUTPUT_COLUMNS
    )


def _configured_companies(
    context: dict[str, Any],
) -> set[str]:
    """Return normalized configured companies; an empty set means all."""
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

    normalized = {
        normalize_company(value)
        for value in raw_values
        if safe_text(value) != ""
    }

    if normalized.intersection(
        _ALL_COMPANIES
    ):
        return set()

    return normalized


def _filter_companies(
    dataframe: pd.DataFrame,
    companies: set[str],
) -> tuple[pd.DataFrame, int]:
    """Apply the configured company filter."""
    if not companies:
        return (
            dataframe.copy()
            .reset_index(drop=True),
            0,
        )

    included = dataframe[
        "Company"
    ].map(
        normalize_company
    ).isin(
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


def _employee_mask(
    vendor_master: pd.DataFrame,
) -> pd.Series:
    """Identify every F-code employee, independently of Account Group."""
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

    return vendor_master[
        "Vendor Code"
    ].map(
        normalize_vendor_code
    ).str.startswith(
        EMPLOYEE_CODE_PREFIX,
        na=False,
    )


def _build_employee_identity(
    vendor_master: pd.DataFrame,
) -> tuple[pd.DataFrame, dict[str, int]]:
    """Build active employee identities from F-code master records."""
    candidates = vendor_master.loc[
        _employee_mask(
            vendor_master
        )
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
            ],
            keep="first",
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
        "distinct_employees": len(
            employees
        ),
    }

    return (
        employees,
        metrics,
    )


def _exclude_f_codes_from_suppliers(
    vendor_population: pd.DataFrame,
) -> tuple[pd.DataFrame, int]:
    """Remove every F-code record from the supplier population."""
    f_code = vendor_population[
        "Vendor Code"
    ].map(
        normalize_vendor_code
    ).str.startswith(
        EMPLOYEE_CODE_PREFIX,
        na=False,
    )

    return (
        vendor_population.loc[
            ~f_code
        ]
        .copy()
        .reset_index(drop=True),
        int(
            f_code.sum()
        ),
    )


def _entity_keys(
    dataframe: pd.DataFrame,
    code_column: str,
) -> set[str]:
    """Return normalized Company + entity-code keys."""
    return set(
        dataframe["Company"].map(
            normalize_company
        )
        + "|"
        + dataframe[code_column].map(
            normalize_vendor_code
        )
    )


def _attach_bank_identities(
    bank_population: pd.DataFrame,
    vendor_population: pd.DataFrame,
    employee_population: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Restrict bank rows to valid suppliers and active employees."""
    banks = bank_population.copy()

    banks["Company"] = banks[
        "Company"
    ].map(
        normalize_company
    )

    banks["Vendor Code"] = banks[
        "Vendor Code"
    ].map(
        normalize_vendor_code
    )

    banks["_Entity Key"] = (
        banks["Company"]
        + "|"
        + banks["Vendor Code"]
    )

    vendor_identity = vendor_population.loc[
        :,
        [
            "Company",
            "Company Name",
            "Vendor Code",
            "Vendor Name",
        ],
    ].copy()

    vendor_identity["Company"] = vendor_identity[
        "Company"
    ].map(
        normalize_company
    )

    vendor_identity["Vendor Code"] = vendor_identity[
        "Vendor Code"
    ].map(
        normalize_vendor_code
    )

    vendor_identity["Company Name"] = vendor_identity[
        "Company Name"
    ].map(
        safe_text
    )

    vendor_identity["Vendor Name"] = vendor_identity[
        "Vendor Name"
    ].map(
        safe_text
    )

    vendor_identity = vendor_identity.drop_duplicates(
        subset=[
            "Company",
            "Vendor Code",
        ]
    )

    employee_identity = employee_population.copy()

    supplier_keys = _entity_keys(
        vendor_identity,
        "Vendor Code",
    )

    employee_keys = _entity_keys(
        employee_identity,
        "Employee Code",
    )

    supplier_banks = banks.loc[
        banks[
            "_Entity Key"
        ].isin(
            supplier_keys
        )
    ].copy()

    supplier_banks = supplier_banks.merge(
        vendor_identity,
        on=[
            "Company",
            "Vendor Code",
        ],
        how="left",
        validate="many_to_one",
    )

    employee_banks = banks.loc[
        banks[
            "_Entity Key"
        ].isin(
            employee_keys
        )
    ].copy()

    employee_banks = employee_banks.rename(
        columns={
            "Vendor Code": "Employee Code",
        }
    )

    employee_banks = employee_banks.merge(
        employee_identity,
        on=[
            "Company",
            "Employee Code",
        ],
        how="left",
        validate="many_to_one",
    )

    return (
        supplier_banks.reset_index(
            drop=True
        ),
        employee_banks.reset_index(
            drop=True
        ),
    )


def _prepare_banks(
    dataframe: pd.DataFrame,
    *,
    code_column: str,
    name_column: str,
) -> tuple[pd.DataFrame, dict[str, int]]:
    """Normalize complete bank combinations for exact comparison."""
    required = {
        "Company",
        code_column,
        name_column,
        *_BANK_DETAIL_COLUMNS,
    }

    missing = sorted(
        required.difference(
            dataframe.columns
        )
    )

    if missing:
        raise ValueError(
            f"{CONTROL_ID}: bank population for {code_column!r} "
            f"is missing columns: {missing}."
        )

    selected_columns = [
        "Company",
        code_column,
        name_column,
        *_BANK_DETAIL_COLUMNS,
    ]

    if "Company Name" in dataframe.columns:
        selected_columns.insert(
            1,
            "Company Name",
        )

    result = dataframe.loc[
        :,
        selected_columns,
    ].copy()

    result["Company"] = result[
        "Company"
    ].map(
        normalize_company
    )

    if "Company Name" in result.columns:
        result["Company Name"] = result[
            "Company Name"
        ].map(
            safe_text
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

    for column in _BANK_DETAIL_COLUMNS:
        result[column] = result[
            column
        ].map(
            safe_text
        )

    for column in [
        "Bank Country",
        "Bank Code",
        "Bank Account",
    ]:
        result[
            f"_Normalized {column}"
        ] = result[
            column
        ].map(
            normalize_exact_key
        )

    valid_identity = (
        result["Company"].ne("")
        & result[code_column].ne("")
    )

    complete_bank = (
        result["_Normalized Bank Country"].ne("")
        & result["_Normalized Bank Code"].ne("")
        & result["_Normalized Bank Account"].ne("")
    )

    comparable = result.loc[
        valid_identity
        & complete_bank
    ].copy()

    comparable["_Bank Key"] = (
        comparable[
            "_Normalized Bank Country"
        ]
        + "|"
        + comparable[
            "_Normalized Bank Code"
        ]
        + "|"
        + comparable[
            "_Normalized Bank Account"
        ]
    )

    comparable = (
        comparable.drop_duplicates(
            subset=[
                "Company",
                code_column,
                "_Bank Key",
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
        "incomplete_bank_rows": int(
            (
                valid_identity
                & ~complete_bank
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
        "distinct_bank_keys": comparable[
            [
                "Company",
                "_Bank Key",
            ]
        ].drop_duplicates().shape[0],
        "valid_from_populated": int(
            comparable[
                "Bank Valid From"
            ].ne("").sum()
        ),
        "valid_to_populated": int(
            comparable[
                "Bank Valid To"
            ].ne("").sum()
        ),
    }

    return (
        comparable,
        metrics,
    )


def _bank_account_quality(
    account_key: pd.Series,
) -> pd.Series:
    """Classify normalized bank-account content."""
    quality = pd.Series(
        "ALPHANUMERIC",
        index=account_key.index,
        dtype=object,
    )

    numeric = account_key.str.fullmatch(
        r"\d+",
        na=False,
    )

    has_digit = account_key.str.contains(
        r"\d",
        regex=True,
        na=False,
    )

    has_letter = account_key.str.contains(
        r"[A-Z]",
        regex=True,
        na=False,
    )

    quality.loc[
        numeric
    ] = "NUMERIC"

    quality.loc[
        has_letter
        & ~has_digit
    ] = "TEXT PLACEHOLDER"

    quality.loc[
        account_key.eq("")
    ] = "INVALID"

    return quality


def _bank_concentration(
    supplier_count: pd.Series,
    employee_count: pd.Series,
) -> pd.Series:
    """Classify how many entities share a bank key."""
    concentration = pd.Series(
        "INDIVIDUAL",
        index=supplier_count.index,
        dtype=object,
    )

    shared = (
        supplier_count.gt(
            1
        )
        | employee_count.gt(
            1
        )
    )

    massive = (
        supplier_count.ge(
            MASSIVE_ENTITY_THRESHOLD
        )
        | employee_count.ge(
            MASSIVE_ENTITY_THRESHOLD
        )
    )

    concentration.loc[
        shared
    ] = "SHARED"

    concentration.loc[
        massive
    ] = "MASSIVE"

    return concentration


def build_vm_011(
    vendor_bank_population: pd.DataFrame,
    employee_bank_population: pd.DataFrame,
) -> pd.DataFrame:
    """Return exact supplier/employee bank matches within each company."""
    missing_vendors = sorted(
        _VENDOR_REQUIRED_COLUMNS.difference(
            vendor_bank_population.columns
        )
    )

    missing_employees = sorted(
        _EMPLOYEE_REQUIRED_COLUMNS.difference(
            employee_bank_population.columns
        )
    )

    if (
        missing_vendors
        or missing_employees
    ):
        raise ValueError(
            f"{CONTROL_ID}: "
            f"missing vendor-bank columns: {missing_vendors}; "
            f"missing employee-bank columns: {missing_employees}."
        )

    vendor_codes = vendor_bank_population[
        "Vendor Code"
    ].map(
        normalize_vendor_code
    )

    vendors = vendor_bank_population.loc[
        ~vendor_codes.str.startswith(
            EMPLOYEE_CODE_PREFIX,
            na=False,
        )
    ].copy()

    vendor_banks, _ = _prepare_banks(
        vendors,
        code_column="Vendor Code",
        name_column="Vendor Name",
    )

    employee_banks, _ = _prepare_banks(
        employee_bank_population,
        code_column="Employee Code",
        name_column="Employee Name",
    )

    if (
        vendor_banks.empty
        or employee_banks.empty
    ):
        return _empty_output()

    supplier_counts = (
        vendor_banks.groupby(
            [
                "Company",
                "_Bank Key",
            ],
            sort=False,
            observed=True,
        )["Vendor Code"]
        .nunique()
        .rename(
            "Suppliers at Bank"
        )
        .reset_index()
    )

    employee_counts = (
        employee_banks.groupby(
            [
                "Company",
                "_Bank Key",
            ],
            sort=False,
            observed=True,
        )["Employee Code"]
        .nunique()
        .rename(
            "Employees at Bank"
        )
        .reset_index()
    )

    employee_columns = [
        "Company",
        "Employee Code",
        "Employee Name",
        *_BANK_DETAIL_COLUMNS,
        "_Normalized Bank Account",
        "_Bank Key",
    ]

    matches = vendor_banks.merge(
        employee_banks[
            employee_columns
        ],
        on=[
            "Company",
            "_Bank Key",
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
            "_Bank Key",
        ],
        keep="first",
    )

    if matches.empty:
        return _empty_output()

    matches = matches.merge(
        supplier_counts,
        on=[
            "Company",
            "_Bank Key",
        ],
        how="left",
        validate="many_to_one",
    )

    matches = matches.merge(
        employee_counts,
        on=[
            "Company",
            "_Bank Key",
        ],
        how="left",
        validate="many_to_one",
    )

    matches["Match Rows at Bank"] = (
        matches.groupby(
            [
                "Company",
                "_Bank Key",
            ],
            sort=False,
            observed=True,
        )["Vendor Code"]
        .transform(
            "size"
        )
        .astype(int)
    )

    group_key = (
        matches["Company"]
        + "|"
        + matches["_Bank Key"]
    )

    matches["Bank Group"] = (
        pd.factorize(
            group_key,
            sort=True,
        )[0]
        + 1
    )

    matches["Bank Concentration"] = (
        _bank_concentration(
            matches[
                "Suppliers at Bank"
            ],
            matches[
                "Employees at Bank"
            ],
        )
    )

    matches["Bank Account Quality"] = (
        _bank_account_quality(
            matches[
                "_Normalized Bank Account Vendor"
            ]
        )
    )

    matches["Potential Placeholder Account"] = (
        matches[
            "Bank Account Quality"
        ]
        .map(
            {
                "TEXT PLACEHOLDER": "YES",
                "ALPHANUMERIC": "REVIEW",
                "NUMERIC": "NO",
                "INVALID": "YES",
            }
        )
    )

    matches["Audit Treatment"] = (
        matches[
            "Bank Concentration"
        ]
        .map(
            {
                "MASSIVE": (
                    "PENDING SHARED BANK KEY VALIDATION"
                ),
                "SHARED": (
                    "PENDING SHARED BANK REVIEW"
                ),
                "INDIVIDUAL": (
                    "PENDING INDIVIDUAL REVIEW"
                ),
            }
        )
    )

    placeholder = matches[
        "Bank Account Quality"
    ].isin(
        [
            "TEXT PLACEHOLDER",
            "INVALID",
        ]
    )

    matches.loc[
        placeholder,
        "Audit Treatment",
    ] = "PENDING BANK MASTER DATA VALIDATION"

    output = pd.DataFrame(
        {
            "Bank Group": matches[
                "Bank Group"
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
            "Vendor Bank Country": matches[
                "Bank Country Vendor"
            ],
            "Vendor Bank Code": matches[
                "Bank Code Vendor"
            ],
            "Vendor Bank Account": matches[
                "Bank Account Vendor"
            ],
            "Vendor Account Holder Name": matches[
                "Account Holder Name Vendor"
            ],
            "Vendor Bank Valid From": matches[
                "Bank Valid From Vendor"
            ],
            "Vendor Bank Valid To": matches[
                "Bank Valid To Vendor"
            ],
            "Employee Code": matches[
                "Employee Code"
            ],
            "Employee Name": matches[
                "Employee Name"
            ],
            "Employee Bank Country": matches[
                "Bank Country Employee"
            ],
            "Employee Bank Code": matches[
                "Bank Code Employee"
            ],
            "Employee Bank Account": matches[
                "Bank Account Employee"
            ],
            "Employee Account Holder Name": matches[
                "Account Holder Name Employee"
            ],
            "Employee Bank Valid From": matches[
                "Bank Valid From Employee"
            ],
            "Employee Bank Valid To": matches[
                "Bank Valid To Employee"
            ],
            "Suppliers at Bank": matches[
                "Suppliers at Bank"
            ],
            "Employees at Bank": matches[
                "Employees at Bank"
            ],
            "Match Rows at Bank": matches[
                "Match Rows at Bank"
            ],
            "Bank Concentration": matches[
                "Bank Concentration"
            ],
            "Bank Account Quality": matches[
                "Bank Account Quality"
            ],
            "Potential Placeholder Account": matches[
                "Potential Placeholder Account"
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
            "Bank Group",
        ],
        keep="first",
    )

    treatment_priority = {
        "PENDING INDIVIDUAL REVIEW": 1,
        "PENDING SHARED BANK REVIEW": 2,
        "PENDING SHARED BANK KEY VALIDATION": 3,
        "PENDING BANK MASTER DATA VALIDATION": 4,
    }

    output["_Priority"] = output[
        "Audit Treatment"
    ].map(
        treatment_priority
    ).fillna(
        99
    )

    return (
        output.sort_values(
            [
                "_Priority",
                "Bank Group",
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


def run_vm_011(
    context: dict[str, Any],
) -> dict[str, Any]:
    """Execute VM11 and replace only the VM11 result worksheet."""
    started = perf_counter()

    companies = _configured_companies(
        context
    )

    vendor_source = load_vm_vendors(
        context
    )

    vendor_master_all = build_vendor_master_population(
        vendor_source
    )

    vendor_bank_population = build_vendor_bank_population(
        vendor_source
    )

    vendor_master_rows = len(
        vendor_master_all
    )

    all_employee_candidates = int(
        _employee_mask(
            vendor_master_all
        ).sum()
    )

    (
        vendor_master,
        excluded_company_rows,
    ) = _filter_companies(
        vendor_master_all,
        companies,
    )

    configured_employee_candidates = int(
        _employee_mask(
            vendor_master
        ).sum()
    )

    (
        employee_identity,
        employee_metrics,
    ) = _build_employee_identity(
        vendor_master
    )

    (
        population_before_f_separation,
        vendor_metrics,
    ) = get_valid_vendor_population(
        vendor_master
    )

    (
        vendor_population,
        f_codes_removed,
    ) = _exclude_f_codes_from_suppliers(
        population_before_f_separation
    )

    (
        supplier_bank_rows,
        employee_bank_rows,
    ) = _attach_bank_identities(
        vendor_bank_population,
        vendor_population,
        employee_identity,
    )

    (
        comparable_supplier_banks,
        supplier_bank_metrics,
    ) = _prepare_banks(
        supplier_bank_rows,
        code_column="Vendor Code",
        name_column="Vendor Name",
    )

    (
        comparable_employee_banks,
        employee_bank_metrics,
    ) = _prepare_banks(
        employee_bank_rows,
        code_column="Employee Code",
        name_column="Employee Name",
    )

    stage_started = _print_timing(
        "input load and population validation",
        started,
    )

    output = build_vm_011(
        supplier_bank_rows,
        employee_bank_rows,
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
            "Vendor Bank Valid From",
            "Vendor Bank Valid To",
            "Employee Bank Valid From",
            "Employee Bank Valid To",
        ],
        integer_columns=[
            "Bank Group",
            "Suppliers at Bank",
            "Employees at Bank",
            "Match Rows at Bank",
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

    distinct_bank_groups = (
        output[
            [
                "CoCo",
                "Bank Group",
            ]
        ]
        .drop_duplicates()
        .shape[0]
        if not output.empty
        else 0
    )

    def distinct_group_count(
        column: str,
        value: str,
    ) -> int:
        if output.empty:
            return 0

        return (
            output.loc[
                output[column].eq(
                    value
                ),
                [
                    "CoCo",
                    "Bank Group",
                ],
            ]
            .drop_duplicates()
            .shape[0]
        )

    placeholder_rows = int(
        output[
            "Bank Account Quality"
        ].eq(
            "TEXT PLACEHOLDER"
        ).sum()
    )

    largest_match_rows = (
        int(
            output[
                "Match Rows at Bank"
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
        f"{CONTROL_ID} vendor bank rows: "
        f"{len(vendor_bank_population)}"
    )
    print(
        f"{CONTROL_ID} rows excluded by CONFIG company: "
        f"{excluded_company_rows}"
    )
    print(
        f"{CONTROL_ID} valid population before F separation: "
        f"{vendor_metrics.get('output_rows', len(population_before_f_separation))}"
    )
    print(
        f"{CONTROL_ID} F-code rows removed from suppliers: "
        f"{f_codes_removed}"
    )
    print(
        f"{CONTROL_ID} valid supplier rows: "
        f"{len(vendor_population)}"
    )
    print(
        f"{CONTROL_ID} employee candidate rows before CONFIG company: "
        f"{all_employee_candidates}"
    )
    print(
        f"{CONTROL_ID} employee rows excluded by CONFIG company: "
        f"{all_employee_candidates - configured_employee_candidates}"
    )
    print(
        f"{CONTROL_ID} employees excluded by deletion flags: "
        f"{employee_metrics.get('employee_excluded_any_deletion', 0)}"
    )
    print(
        f"{CONTROL_ID} active employee rows: "
        f"{employee_metrics.get('employee_output_rows', len(employee_identity))}"
    )
    print(
        f"{CONTROL_ID} raw supplier bank rows: "
        f"{len(supplier_bank_rows)}"
    )
    print(
        f"{CONTROL_ID} raw employee bank rows: "
        f"{len(employee_bank_rows)}"
    )
    print(
        f"{CONTROL_ID} supplier complete bank rows: "
        f"{supplier_bank_metrics.get('comparable_rows', len(comparable_supplier_banks))}"
    )
    print(
        f"{CONTROL_ID} employee complete bank rows: "
        f"{employee_bank_metrics.get('comparable_rows', len(comparable_employee_banks))}"
    )
    print(
        f"{CONTROL_ID} distinct matching bank keys: "
        f"{distinct_bank_groups}"
    )
    print(
        f"{CONTROL_ID} numeric bank groups: "
        f"{distinct_group_count('Bank Account Quality', 'NUMERIC')}"
    )
    print(
        f"{CONTROL_ID} alphanumeric bank groups: "
        f"{distinct_group_count('Bank Account Quality', 'ALPHANUMERIC')}"
    )
    print(
        f"{CONTROL_ID} text-placeholder bank groups: "
        f"{distinct_group_count('Bank Account Quality', 'TEXT PLACEHOLDER')}"
    )
    print(
        f"{CONTROL_ID} massive bank groups: "
        f"{distinct_group_count('Bank Concentration', 'MASSIVE')}"
    )
    print(
        f"{CONTROL_ID} shared bank groups: "
        f"{distinct_group_count('Bank Concentration', 'SHARED')}"
    )
    print(
        f"{CONTROL_ID} individual bank groups: "
        f"{distinct_group_count('Bank Concentration', 'INDIVIDUAL')}"
    )
    print(
        f"{CONTROL_ID} text-placeholder match rows: "
        f"{placeholder_rows}"
    )
    print(
        f"{CONTROL_ID} largest bank match rows: "
        f"{largest_match_rows}"
    )
    print(
        f"{CONTROL_ID} exception rows: "
        f"{len(output)}"
    )

    validity_available = (
        supplier_bank_metrics.get(
            "valid_from_populated",
            0,
        )
        + supplier_bank_metrics.get(
            "valid_to_populated",
            0,
        )
        + employee_bank_metrics.get(
            "valid_from_populated",
            0,
        )
        + employee_bank_metrics.get(
            "valid_to_populated",
            0,
        )
    )

    if validity_available == 0:
        print(
            "WARNING: Bank validity dates are unavailable; "
            "VM_011 compares all complete bank combinations "
            "present in the source."
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
