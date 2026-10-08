"""
VM_010 - Vendors sharing a telephone number with a Logicalis employee.

Objective
---------
Identify valid suppliers whose normalized telephone number is identical to
the normalized telephone number of an active Logicalis employee in the same
company.

Logicalis employees are records in the VM vendor master whose Vendor Code
starts with F. Employee records with central or company deletion flags are
excluded.

The supplier population is obtained through the common VM valid-population
rules and every F-code record is then removed from the supplier side.

Telephone comparison is exact after retaining only digits. All matches are
preserved and classified for audit review.
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
    normalize_phone,
    normalize_vendor_code,
    safe_text,
    write_vm_control_sheet,
)


CONTROL_ID = "VM_010"
SHEET_NAME = "VM10"

EMPLOYEE_CODE_PREFIX = "F"
MASSIVE_ENTITY_THRESHOLD = 10

OUTPUT_COLUMNS = [
    "Phone Group",
    "Company",
    "CoCo",
    "Vendor Code",
    "Vendor Name",
    "Vendor Phone",
    "Employee Code",
    "Employee Name",
    "Employee Phone",
    "Phone Digit Length",
    "Suppliers at Phone",
    "Employees at Phone",
    "Match Rows at Phone",
    "Phone Concentration",
    "Phone Quality",
    "Potential Corporate Phone",
    "Audit Treatment",
]

_VENDOR_REQUIRED_COLUMNS = {
    "Company",
    "Company Name",
    "Vendor Code",
    "Vendor Name",
    "Phone1",
}

_EMPLOYEE_REQUIRED_COLUMNS = {
    "Company",
    "Employee Code",
    "Employee Name",
    "Phone1",
}

_EMPLOYEE_SOURCE_REQUIRED_COLUMNS = {
    "Company",
    "Vendor Code",
    "Vendor Name",
    "Phone1",
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
    """Return an empty VM10 output with the exact required schema."""
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


def _logicalis_employee_mask(
    vendor_master: pd.DataFrame,
) -> pd.Series:
    """Identify Logicalis employees by the functional F-code rule."""
    missing = sorted(
        _EMPLOYEE_SOURCE_REQUIRED_COLUMNS.difference(
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


def _build_employee_population(
    vendor_master: pd.DataFrame,
) -> tuple[pd.DataFrame, dict[str, int]]:
    """Build active Logicalis employees from F-code master records."""
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
            "Phone1",
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

    employees["Phone1"] = employees[
        "Phone1"
    ].map(
        safe_text
    )

    employees = (
        employees.drop_duplicates(
            subset=[
                "Company",
                "Employee Code",
                "Phone1",
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


def _exclude_employee_codes_from_suppliers(
    vendor_population: pd.DataFrame,
) -> tuple[pd.DataFrame, int]:
    """Remove every F-code record from the supplier population."""
    if "Vendor Code" not in vendor_population.columns:
        raise ValueError(
            f"{CONTROL_ID}: supplier population is missing "
            "'Vendor Code'."
        )

    employee_code = vendor_population[
        "Vendor Code"
    ].map(
        normalize_vendor_code
    ).str.startswith(
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


def _prepare_phones(
    dataframe: pd.DataFrame,
    *,
    code_column: str,
    name_column: str,
) -> tuple[pd.DataFrame, dict[str, int]]:
    """Normalize nonblank telephone numbers for exact comparison."""
    required = {
        "Company",
        code_column,
        name_column,
        "Phone1",
    }

    missing = sorted(
        required.difference(
            dataframe.columns
        )
    )

    if missing:
        raise ValueError(
            f"{CONTROL_ID}: phone population for {code_column!r} "
            f"is missing columns: {missing}."
        )

    selected_columns = [
        "Company",
        code_column,
        name_column,
        "Phone1",
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

    result["Phone1"] = result[
        "Phone1"
    ].map(
        safe_text
    )

    result["_Phone Key"] = result[
        "Phone1"
    ].map(
        normalize_phone
    )

    valid_identity = (
        result["Company"].ne("")
        & result[code_column].ne("")
    )

    comparable_phone = result[
        "_Phone Key"
    ].ne("")

    comparable = (
        result.loc[
            valid_identity
            & comparable_phone
        ]
        .drop_duplicates(
            subset=[
                "Company",
                code_column,
                "_Phone Key",
            ],
            keep="first",
        )
        .reset_index(drop=True)
    )

    comparable["_Phone Length"] = comparable[
        "_Phone Key"
    ].str.len()

    metrics = {
        "input_rows": len(
            result
        ),
        "invalid_identity_rows": int(
            (~valid_identity).sum()
        ),
        "blank_phone_rows": int(
            (
                valid_identity
                & ~comparable_phone
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
        "distinct_phones": comparable[
            [
                "Company",
                "_Phone Key",
            ]
        ].drop_duplicates().shape[0],
    }

    return (
        comparable,
        metrics,
    )


def _phone_quality(
    phone_length: pd.Series,
) -> pd.Series:
    """Classify phone quality by normalized digit length."""
    quality = pd.Series(
        "STANDARD",
        index=phone_length.index,
        dtype=object,
    )

    quality.loc[
        phone_length.le(
            6
        )
    ] = "VERY SHORT"

    quality.loc[
        phone_length.between(
            7,
            9,
            inclusive="both",
        )
    ] = "SHORT"

    quality.loc[
        phone_length.gt(
            15
        )
    ] = "LONG"

    return quality


def _phone_concentration(
    supplier_count: pd.Series,
    employee_count: pd.Series,
) -> pd.Series:
    """Classify the number of entities sharing a telephone."""
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


def build_vm_010(
    vendor_population: pd.DataFrame,
    employee_population: pd.DataFrame,
) -> pd.DataFrame:
    """Return exact supplier/employee telephone matches by company."""
    missing_vendors = sorted(
        _VENDOR_REQUIRED_COLUMNS.difference(
            vendor_population.columns
        )
    )

    missing_employees = sorted(
        _EMPLOYEE_REQUIRED_COLUMNS.difference(
            employee_population.columns
        )
    )

    if (
        missing_vendors
        or missing_employees
    ):
        raise ValueError(
            f"{CONTROL_ID}: "
            f"missing vendor columns: {missing_vendors}; "
            f"missing employee columns: {missing_employees}."
        )

    vendor_codes = vendor_population[
        "Vendor Code"
    ].map(
        normalize_vendor_code
    )

    vendors = vendor_population.loc[
        ~vendor_codes.str.startswith(
            EMPLOYEE_CODE_PREFIX,
            na=False,
        )
    ].copy()

    vendor_phones, _ = _prepare_phones(
        vendors,
        code_column="Vendor Code",
        name_column="Vendor Name",
    )

    employee_phones, _ = _prepare_phones(
        employee_population,
        code_column="Employee Code",
        name_column="Employee Name",
    )

    if (
        vendor_phones.empty
        or employee_phones.empty
    ):
        return _empty_output()

    supplier_counts = (
        vendor_phones.groupby(
            [
                "Company",
                "_Phone Key",
            ],
            sort=False,
            observed=True,
        )["Vendor Code"]
        .nunique()
        .rename(
            "Suppliers at Phone"
        )
        .reset_index()
    )

    employee_counts = (
        employee_phones.groupby(
            [
                "Company",
                "_Phone Key",
            ],
            sort=False,
            observed=True,
        )["Employee Code"]
        .nunique()
        .rename(
            "Employees at Phone"
        )
        .reset_index()
    )

    matches = vendor_phones.merge(
        employee_phones[
            [
                "Company",
                "Employee Code",
                "Employee Name",
                "Phone1",
                "_Phone Key",
            ]
        ],
        on=[
            "Company",
            "_Phone Key",
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
            "_Phone Key",
        ],
        keep="first",
    )

    if matches.empty:
        return _empty_output()

    matches = matches.merge(
        supplier_counts,
        on=[
            "Company",
            "_Phone Key",
        ],
        how="left",
        validate="many_to_one",
    )

    matches = matches.merge(
        employee_counts,
        on=[
            "Company",
            "_Phone Key",
        ],
        how="left",
        validate="many_to_one",
    )

    matches["Match Rows at Phone"] = (
        matches.groupby(
            [
                "Company",
                "_Phone Key",
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
        + matches["_Phone Key"]
    )

    matches["Phone Group"] = (
        pd.factorize(
            group_key,
            sort=True,
        )[0]
        + 1
    )

    matches["Phone Digit Length"] = matches[
        "_Phone Key"
    ].str.len()

    matches["Phone Concentration"] = (
        _phone_concentration(
            matches[
                "Suppliers at Phone"
            ],
            matches[
                "Employees at Phone"
            ],
        )
    )

    matches["Phone Quality"] = _phone_quality(
        matches[
            "Phone Digit Length"
        ]
    )

    matches["Potential Corporate Phone"] = (
        matches[
            "Phone Concentration"
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
            "Phone Concentration"
        ]
        .map(
            {
                "MASSIVE": (
                    "PENDING CORPORATE PHONE VALIDATION"
                ),
                "SHARED": (
                    "PENDING SHARED PHONE REVIEW"
                ),
                "INDIVIDUAL": (
                    "PENDING INDIVIDUAL REVIEW"
                ),
            }
        )
    )

    data_quality = matches[
        "Phone Quality"
    ].isin(
        [
            "VERY SHORT",
            "LONG",
        ]
    )

    matches.loc[
        data_quality,
        "Audit Treatment",
    ] = "PENDING PHONE DATA QUALITY REVIEW"

    output = pd.DataFrame(
        {
            "Phone Group": matches[
                "Phone Group"
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
            "Vendor Phone": matches[
                "Phone1 Vendor"
            ],
            "Employee Code": matches[
                "Employee Code"
            ],
            "Employee Name": matches[
                "Employee Name"
            ],
            "Employee Phone": matches[
                "Phone1 Employee"
            ],
            "Phone Digit Length": matches[
                "Phone Digit Length"
            ],
            "Suppliers at Phone": matches[
                "Suppliers at Phone"
            ],
            "Employees at Phone": matches[
                "Employees at Phone"
            ],
            "Match Rows at Phone": matches[
                "Match Rows at Phone"
            ],
            "Phone Concentration": matches[
                "Phone Concentration"
            ],
            "Phone Quality": matches[
                "Phone Quality"
            ],
            "Potential Corporate Phone": matches[
                "Potential Corporate Phone"
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
            "Phone Group",
        ],
        keep="first",
    )

    priority = {
        "INDIVIDUAL": 1,
        "SHARED": 2,
        "MASSIVE": 3,
    }

    output["_Priority"] = output[
        "Phone Concentration"
    ].map(
        priority
    ).fillna(
        99
    )

    return (
        output.sort_values(
            [
                "_Priority",
                "Phone Group",
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


def run_vm_010(
    context: dict[str, Any],
) -> dict[str, Any]:
    """Execute VM10 and replace only the VM10 result worksheet."""
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

    all_employee_candidates = int(
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

    (
        employee_population,
        employee_metrics,
    ) = _build_employee_population(
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
    ) = _exclude_employee_codes_from_suppliers(
        population_before_f_separation
    )

    if employee_population.empty:
        raise ValueError(
            f"{CONTROL_ID}: no active F-code employees were found."
        )

    if vendor_population.empty:
        raise ValueError(
            f"{CONTROL_ID}: valid supplier population is empty."
        )

    (
        comparable_vendors,
        vendor_phone_metrics,
    ) = _prepare_phones(
        vendor_population,
        code_column="Vendor Code",
        name_column="Vendor Name",
    )

    (
        comparable_employees,
        employee_phone_metrics,
    ) = _prepare_phones(
        employee_population,
        code_column="Employee Code",
        name_column="Employee Name",
    )

    stage_started = _print_timing(
        "input load and population validation",
        started,
    )

    output = build_vm_010(
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
            "Phone Group",
            "Phone Digit Length",
            "Suppliers at Phone",
            "Employees at Phone",
            "Match Rows at Phone",
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

    distinct_matching_phones = (
        output[
            [
                "CoCo",
                "Phone Group",
            ]
        ]
        .drop_duplicates()
        .shape[0]
        if not output.empty
        else 0
    )

    def group_count(
        value: str,
    ) -> int:
        if output.empty:
            return 0

        return (
            output.loc[
                output[
                    "Phone Concentration"
                ].eq(
                    value
                ),
                [
                    "CoCo",
                    "Phone Group",
                ],
            ]
            .drop_duplicates()
            .shape[0]
        )

    massive_rows = int(
        output[
            "Phone Concentration"
        ].eq(
            "MASSIVE"
        ).sum()
    )

    data_quality_rows = int(
        output[
            "Audit Treatment"
        ].eq(
            "PENDING PHONE DATA QUALITY REVIEW"
        ).sum()
    )

    largest_supplier_count = (
        int(
            output[
                "Suppliers at Phone"
            ].max()
        )
        if not output.empty
        else 0
    )

    largest_employee_count = (
        int(
            output[
                "Employees at Phone"
            ].max()
        )
        if not output.empty
        else 0
    )

    largest_match_rows = (
        int(
            output[
                "Match Rows at Phone"
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
        f"{CONTROL_ID} employee candidate rows after CONFIG company: "
        f"{employee_metrics.get('employee_candidate_rows', 0)}"
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
        f"{CONTROL_ID} suppliers with normalized phone: "
        f"{vendor_phone_metrics.get('comparable_rows', len(comparable_vendors))}"
    )
    print(
        f"{CONTROL_ID} employees with normalized phone: "
        f"{employee_phone_metrics.get('comparable_rows', len(comparable_employees))}"
    )
    print(
        f"{CONTROL_ID} distinct matching phones: "
        f"{distinct_matching_phones}"
    )
    print(
        f"{CONTROL_ID} massive phone groups: "
        f"{group_count('MASSIVE')}"
    )
    print(
        f"{CONTROL_ID} shared phone groups: "
        f"{group_count('SHARED')}"
    )
    print(
        f"{CONTROL_ID} individual phone groups: "
        f"{group_count('INDIVIDUAL')}"
    )
    print(
        f"{CONTROL_ID} massive phone match rows: "
        f"{massive_rows}"
    )
    print(
        f"{CONTROL_ID} phone data-quality rows: "
        f"{data_quality_rows}"
    )
    print(
        f"{CONTROL_ID} largest phone supplier count: "
        f"{largest_supplier_count}"
    )
    print(
        f"{CONTROL_ID} largest phone employee count: "
        f"{largest_employee_count}"
    )
    print(
        f"{CONTROL_ID} largest phone match rows: "
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
