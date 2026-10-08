"""
VM_007 - Vendors with frequent bank-data changes.

Objective
---------
Identify valid vendors with at least the number of distinct bank-change events
configured in PARAM1 during the VM analysis period.

PARAM1 defines the minimum number of distinct change events. When PARAM1 is
blank, DEFAULT_MINIMUM_CHANGES is used.

Initial LFBK record creations identified by CHNGIND=I are excluded because
they represent the first registration of bank data rather than a subsequent
modification.

The threshold is calculated using distinct Change Event Key values. Multiple
CDPOS positions belonging to the same change document count as one event for
threshold purposes, but every changed field remains visible in the output.

The input workbook resolution, SAP CDHDR/CDPOS reconciliation, normalization,
common vendor-population validation and workbook writing are delegated to
core.vm_common.

Output
------
One row per qualifying vendor, company and changed bank-data position:

    Company
    CoCo
    Vendor Code
    Vendor Name
    Cambios
    Change Date
    Change Time
    Changed By
    Transaction Code
    Change Document
    Changed Record Key
    Changed Field
    Change Type
    Old Value
    New Value
"""

from time import perf_counter
from typing import Any

import pandas as pd

from core.vm_common import (
    build_vendor_master_population,
    get_valid_vendor_population,
    get_vm_period,
    load_vm_bank_changes,
    load_vm_vendors,
    normalize_company,
    normalize_identifier,
    normalize_upper_text,
    normalize_vendor_code,
    safe_text,
    write_vm_control_sheet,
)


CONTROL_ID = "VM_007"
SHEET_NAME = "VM07"

# This value is used only when PARAM1 is blank.
# Configure PARAM1=3 in config.xlsx for the requested threshold.
DEFAULT_MINIMUM_CHANGES = 2

OUTPUT_COLUMNS = [
    "Company",
    "CoCo",
    "Vendor Code",
    "Vendor Name",
    "Cambios",
    "Change Date",
    "Change Time",
    "Changed By",
    "Transaction Code",
    "Change Document",
    "Changed Record Key",
    "Changed Field",
    "Change Type",
    "Old Value",
    "New Value",
]

CHANGE_DETAIL_COLUMNS = [
    "Vendor Code",
    "Change Event Key",
    "Change Date",
    "Change Time",
    "Changed By",
    "Transaction Code",
    "Change Document",
    "Changed Record Key",
    "Changed Field",
    "Change Type",
    "Old Value",
    "New Value",
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


def _minimum_changes(
    context: dict[str, Any],
) -> int:
    """
    Return the minimum number of distinct bank-change events.

    PARAM1 must be a positive whole number. A blank PARAM1 uses
    DEFAULT_MINIMUM_CHANGES.
    """
    control_config = context.get(
        "control",
        {},
    )

    if not isinstance(
        control_config,
        dict,
    ):
        raise ValueError(
            f"{CONTROL_ID} requires context['control'] configuration."
        )

    raw_value = control_config.get(
        "param1",
        "",
    )

    text_value = safe_text(
        raw_value
    )

    if text_value == "":
        return DEFAULT_MINIMUM_CHANGES

    try:
        numeric_value = float(
            text_value.replace(
                ",",
                ".",
            )
        )

        threshold = int(
            numeric_value
        )

    except (
        TypeError,
        ValueError,
    ) as error:
        raise ValueError(
            f"{CONTROL_ID} PARAM1 must be a positive whole number; "
            f"received {raw_value!r}."
        ) from error

    if (
        numeric_value != threshold
        or threshold < 1
    ):
        raise ValueError(
            f"{CONTROL_ID} PARAM1 must be a positive whole number; "
            f"received {raw_value!r}."
        )

    return threshold


def _configured_companies(
    context: dict[str, Any],
) -> set[str]:
    """
    Return normalized company codes configured for VM.

    An empty value, ALL, *, TODAS or TODOS means that every company present in
    the validated vendor population is included.
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
            or text.upper() in {
                "ALL",
                "*",
                "TODAS",
                "TODOS",
            }
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
        {
            "ALL",
            "*",
            "TODAS",
            "TODOS",
        }
    ):
        return set()

    return normalized_values


def _load_valid_vendor_population(
    context: dict[str, Any],
) -> tuple[pd.DataFrame, dict[str, Any]]:
    """
    Load the vendor workbook and return the valid VM07 vendor population.

    The configured-company filter and the common VM population exclusions are
    applied before matching vendors to bank-change events.
    """
    vendor_source = load_vm_vendors(
        context
    )

    vendor_master = build_vendor_master_population(
        vendor_source
    )

    master_rows_before_company = len(
        vendor_master
    )

    configured_companies = _configured_companies(
        context
    )

    excluded_company = 0

    if configured_companies:
        company_values = vendor_master[
            "Company"
        ].map(
            normalize_company
        )

        included = company_values.isin(
            configured_companies
        )

        excluded_company = int(
            (~included).sum()
        )

        vendor_master = (
            vendor_master.loc[
                included
            ]
            .copy()
            .reset_index(drop=True)
        )

    valid_population, population_metrics = (
        get_valid_vendor_population(
            vendor_master
        )
    )

    metrics = {
        "source_rows": len(
            vendor_source
        ),
        "master_rows_before_company": (
            master_rows_before_company
        ),
        "master_rows": len(
            vendor_master
        ),
        "excluded_company": excluded_company,
        **population_metrics,
    }

    return (
        valid_population,
        metrics,
    )


def _validate_input_columns(
    bank_changes: pd.DataFrame,
    vendor_population: pd.DataFrame,
) -> None:
    """Validate all columns required by the VM07 detailed analytic."""
    required_change_columns = set(
        CHANGE_DETAIL_COLUMNS
    )

    required_vendor_columns = {
        "Company",
        "Company Name",
        "Vendor Code",
        "Vendor Name",
    }

    missing_change_columns = sorted(
        required_change_columns.difference(
            bank_changes.columns
        )
    )

    missing_vendor_columns = sorted(
        required_vendor_columns.difference(
            vendor_population.columns
        )
    )

    if (
        missing_change_columns
        or missing_vendor_columns
    ):
        raise ValueError(
            f"{CONTROL_ID} missing required columns. "
            f"Bank changes: {missing_change_columns}; "
            f"vendors: {missing_vendor_columns}."
        )


def _prepare_change_details(
    bank_changes: pd.DataFrame,
    *,
    date_from: Any,
    date_to: Any,
) -> pd.DataFrame:
    """
    Prepare eligible bank-change detail positions.

    Initial LFBK insertions are removed. The returned DataFrame retains every
    changed field so the result explains when, who and what changed.
    """
    changes = bank_changes.loc[
        :,
        CHANGE_DETAIL_COLUMNS,
    ].copy()

    changes["Vendor Code"] = changes[
        "Vendor Code"
    ].map(
        normalize_vendor_code
    )

    changes["Change Event Key"] = changes[
        "Change Event Key"
    ].map(
        safe_text
    )

    changes["Change Date"] = pd.to_datetime(
        changes["Change Date"],
        errors="coerce",
    )

    changes["Change Type"] = changes[
        "Change Type"
    ].map(
        normalize_upper_text
    )

    changes["Change Document"] = changes[
        "Change Document"
    ].map(
        normalize_identifier
    )

    for column in [
        "Change Time",
        "Changed By",
        "Transaction Code",
        "Changed Record Key",
        "Changed Field",
        "Old Value",
        "New Value",
    ]:
        changes[column] = changes[
            column
        ].map(
            safe_text
        )

    start_date = pd.Timestamp(
        date_from
    ).normalize()

    end_date = pd.Timestamp(
        date_to
    ).normalize()

    if start_date > end_date:
        raise ValueError(
            f"{CONTROL_ID}: date_from cannot be later than date_to."
        )

    in_analysis_period = changes[
        "Change Date"
    ].between(
        start_date,
        end_date,
        inclusive="both",
    )

    # SAP CDPOS CHNGIND=I represents initial LFBK record creation.
    is_bank_modification = changes[
        "Change Type"
    ].ne("I")

    has_vendor_code = changes[
        "Vendor Code"
    ].ne("")

    has_event_key = changes[
        "Change Event Key"
    ].ne("")

    eligible = (
        in_analysis_period
        & is_bank_modification
        & has_vendor_code
        & has_event_key
    )

    changes = changes.loc[
        eligible
    ].copy()

    if changes.empty:
        return pd.DataFrame(
            columns=CHANGE_DETAIL_COLUMNS
        )

    # Avoid duplicating identical CDPOS positions while preserving different
    # fields belonging to the same change event.
    changes = changes.drop_duplicates(
        subset=CHANGE_DETAIL_COLUMNS,
        keep="first",
    )

    return (
        changes.sort_values(
            [
                "Vendor Code",
                "Change Date",
                "Change Time",
                "Change Document",
                "Changed Record Key",
                "Changed Field",
            ],
            kind="mergesort",
        )
        .reset_index(drop=True)
    )


def _summarize_vendor_changes(
    change_details: pd.DataFrame,
    *,
    minimum_changes: int,
) -> pd.DataFrame:
    """
    Count distinct bank-change events by Vendor Code.

    Multiple changed fields belonging to one Change Event Key count as one
    event. Only vendors meeting or exceeding PARAM1 are returned.
    """
    if minimum_changes < 1:
        raise ValueError(
            f"{CONTROL_ID}: minimum_changes must be at least 1."
        )

    if change_details.empty:
        return pd.DataFrame(
            columns=[
                "Vendor Code",
                "Cambios",
            ]
        )

    event_rows = (
        change_details.loc[
            :,
            [
                "Vendor Code",
                "Change Event Key",
            ],
        ]
        .drop_duplicates(
            subset=[
                "Vendor Code",
                "Change Event Key",
            ],
            keep="first",
        )
    )

    summary = (
        event_rows.groupby(
            "Vendor Code",
            sort=False,
            observed=True,
        )
        .size()
        .rename("Cambios")
        .reset_index()
    )

    summary = summary.loc[
        summary["Cambios"].ge(
            minimum_changes
        )
    ].copy()

    if summary.empty:
        return pd.DataFrame(
            columns=[
                "Vendor Code",
                "Cambios",
            ]
        )

    summary["Cambios"] = pd.to_numeric(
        summary["Cambios"],
        errors="raise",
    ).astype("int64")

    return (
        summary.sort_values(
            [
                "Cambios",
                "Vendor Code",
            ],
            ascending=[
                False,
                True,
            ],
            kind="mergesort",
        )
        .reset_index(drop=True)
    )


def _prepare_vendor_display(
    vendor_population: pd.DataFrame,
) -> pd.DataFrame:
    """
    Prepare Company, CoCo, Vendor Code and Vendor Name for the output.

    The common VM vendor population uses:
        Company      -> company code
        Company Name -> company description

    Therefore:
        CoCo    <- Company
        Company <- Company Name
    """
    vendors = vendor_population.loc[
        :,
        [
            "Company",
            "Company Name",
            "Vendor Code",
            "Vendor Name",
        ],
    ].copy()

    vendors["CoCo"] = vendors[
        "Company"
    ].map(
        normalize_company
    )

    vendors["Company"] = vendors[
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

    has_company = vendors[
        "CoCo"
    ].ne("")

    has_vendor = vendors[
        "Vendor Code"
    ].ne("")

    vendors = vendors.loc[
        has_company
        & has_vendor,
        [
            "Company",
            "CoCo",
            "Vendor Code",
            "Vendor Name",
        ],
    ].copy()

    return (
        vendors.drop_duplicates(
            subset=[
                "CoCo",
                "Vendor Code",
            ],
            keep="first",
        )
        .sort_values(
            [
                "CoCo",
                "Vendor Code",
            ],
            kind="mergesort",
        )
        .reset_index(drop=True)
    )


def build_vm_007(
    bank_changes: pd.DataFrame,
    vendor_population: pd.DataFrame,
    *,
    date_from: Any,
    date_to: Any,
    minimum_changes: int = DEFAULT_MINIMUM_CHANGES,
) -> pd.DataFrame:
    """
    Return detailed bank changes for vendors meeting the VM07 threshold.

    Threshold grain:
        one distinct Change Event Key per Vendor Code

    Output grain:
        one row per company and changed CDPOS bank-data position

    Cambios is repeated on every detail row and represents the number of
    distinct change events for the vendor, not the number of output rows.
    """
    _validate_input_columns(
        bank_changes,
        vendor_population,
    )

    if minimum_changes < 1:
        raise ValueError(
            f"{CONTROL_ID}: minimum_changes must be at least 1."
        )

    if (
        bank_changes.empty
        or vendor_population.empty
    ):
        return pd.DataFrame(
            columns=OUTPUT_COLUMNS
        )

    change_details = _prepare_change_details(
        bank_changes,
        date_from=date_from,
        date_to=date_to,
    )

    qualifying_vendors = _summarize_vendor_changes(
        change_details,
        minimum_changes=minimum_changes,
    )

    if qualifying_vendors.empty:
        return pd.DataFrame(
            columns=OUTPUT_COLUMNS
        )

    qualifying_details = change_details.merge(
        qualifying_vendors,
        on="Vendor Code",
        how="inner",
        validate="many_to_one",
    )

    vendors = _prepare_vendor_display(
        vendor_population
    )

    if vendors.empty:
        return pd.DataFrame(
            columns=OUTPUT_COLUMNS
        )

    result = qualifying_details.merge(
        vendors,
        on="Vendor Code",
        how="inner",
        validate="many_to_many",
    )

    if result.empty:
        return pd.DataFrame(
            columns=OUTPUT_COLUMNS
        )

    result["Cambios"] = pd.to_numeric(
        result["Cambios"],
        errors="raise",
    ).astype("int64")

    return (
        result.loc[
            :,
            OUTPUT_COLUMNS,
        ]
        .sort_values(
            [
                "Cambios",
                "Company",
                "CoCo",
                "Vendor Code",
                "Change Date",
                "Change Time",
                "Change Document",
                "Changed Record Key",
                "Changed Field",
            ],
            ascending=[
                False,
                True,
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
        .reset_index(drop=True)
    )


def run_vm_007(
    context: dict[str, Any],
) -> dict[str, Any]:
    """Execute VM07 and replace only the VM07 result worksheet."""
    started = perf_counter()

    threshold = _minimum_changes(
        context
    )

    date_from, date_to = get_vm_period(
        context
    )

    vendor_population, vendor_metrics = (
        _load_valid_vendor_population(
            context
        )
    )

    bank_changes, change_metrics = (
        load_vm_bank_changes(
            context
        )
    )

    if bank_changes is None:
        raise FileNotFoundError(
            f"{CONTROL_ID} requires both period-specific "
            "VM BANK CDHDR and CDPOS files."
        )

    stage_started = _print_timing(
        "input load and population validation",
        started,
    )

    output = build_vm_007(
        bank_changes,
        vendor_population,
        date_from=date_from,
        date_to=date_to,
        minimum_changes=threshold,
    )

    stage_started = _print_timing(
        "preparation and analytic logic",
        stage_started,
    )

    # VM07 performs no monetary conversion.
    stage_started = _print_timing(
        "FX conversion (not applicable)",
        stage_started,
    )

    output_file = write_vm_control_sheet(
        context=context,
        sheet_name=SHEET_NAME,
        dataframe=output,
        date_columns=[
            "Change Date",
        ],
        integer_columns=[
            "Cambios",
        ],
    )

    _print_timing(
        "workbook write",
        stage_started,
    )

    qualifying_vendor_count = (
        output["Vendor Code"].nunique()
        if not output.empty
        else 0
    )

    qualifying_company_vendor_count = (
        output.loc[
            :,
            [
                "CoCo",
                "Vendor Code",
            ],
        ]
        .drop_duplicates()
        .shape[0]
        if not output.empty
        else 0
    )

    print(
        f"{CONTROL_ID} analysis period: "
        f"{date_from:%Y-%m-%d} to {date_to:%Y-%m-%d}"
    )
    print(
        f"{CONTROL_ID} minimum distinct change events from PARAM1: "
        f"{threshold}"
    )
    print(
        f"{CONTROL_ID} vendor source rows: "
        f"{vendor_metrics.get('source_rows', 0)}"
    )
    print(
        f"{CONTROL_ID} vendor master rows before company filter: "
        f"{vendor_metrics.get('master_rows_before_company', 0)}"
    )
    print(
        f"{CONTROL_ID} vendor master rows after company filter: "
        f"{vendor_metrics.get('master_rows', 0)}"
    )
    print(
        f"{CONTROL_ID} rows excluded by CONFIG company: "
        f"{vendor_metrics.get('excluded_company', 0)}"
    )
    print(
        f"{CONTROL_ID} valid vendor rows: "
        f"{vendor_metrics.get('output_rows', 0)}"
    )
    print(
        f"{CONTROL_ID} CDHDR rows: "
        f"{change_metrics.get('header_rows', 0)}"
    )
    print(
        f"{CONTROL_ID} CDPOS rows: "
        f"{change_metrics.get('position_rows', 0)}"
    )
    print(
        f"{CONTROL_ID} source change events: "
        f"{change_metrics.get('change_events', 0)}"
    )
    print(
        f"{CONTROL_ID} qualifying distinct vendors in output: "
        f"{qualifying_vendor_count}"
    )
    print(
        f"{CONTROL_ID} qualifying company/vendor combinations: "
        f"{qualifying_company_vendor_count}"
    )
    print(
        f"{CONTROL_ID} exception detail rows: "
        f"{len(output)}"
    )

    for warning in [
        *vendor_metrics.get(
            "warnings",
            [],
        ),
        *change_metrics.get(
            "warnings",
            [],
        ),
    ]:
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