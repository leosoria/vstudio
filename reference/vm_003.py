"""
VM_003 - Vendors sharing the same Phone1 within the same company.

All input resolution, reading, canonicalization, common population rules,
last-invoice construction and workbook writing are delegated to core.vm_common.
"""

from time import perf_counter
from typing import Any

import pandas as pd

from core.vm_common import (
    build_vendor_master_population,
    build_vm_last_invoice_population,
    get_valid_vendor_population,
    load_vm_vendor_postings,
    load_vm_vendors,
    normalize_company,
    safe_text,
    write_vm_control_sheet,
)


CONTROL_ID = "VM_003"
SHEET_NAME = "VM03"

INVOICE_COLUMNS = [
    "Last Invoice Number",
    "Last Transaction Date",
    "Last Inv Amt Doc Currency",
    "Last Inv Amt Doc Currency Indicator",
]

OUTPUT_COLUMNS = [
    "Company",
    "CoCo",
    "Vendor Code",
    "Vendor Name",
    "Phone1",
    *INVOICE_COLUMNS,
    "Group",
]


def _configured_companies(context: dict[str, Any]) -> set[str]:
    """Replicate the VM01/VM02 CONFIG-company normalization exactly."""
    module_config = context.get("module")

    if not isinstance(module_config, dict):
        raise ValueError(
            f"{CONTROL_ID} requires context['module'] configuration."
        )

    raw = module_config.get("companies", "")

    if raw is None:
        return set()

    if isinstance(raw, (list, tuple, set)):
        values = list(raw)
    else:
        text = safe_text(raw)

        if text == "" or text.upper() in {"ALL", "*", "TODAS", "TODOS"}:
            return set()

        values = (
            text.replace(";", ",")
            .replace("|", ",")
            .split(",")
        )

    companies = {
        normalize_company(value)
        for value in values
        if safe_text(value) != ""
    }

    if companies.intersection({"ALL", "*", "TODAS", "TODOS"}):
        return set()

    return companies


def _filter_configured_companies(
    vendor_master: pd.DataFrame,
    context: dict[str, Any],
) -> tuple[pd.DataFrame, int]:
    """Filter normalized Company codes using the approved VM02 pattern."""
    companies = _configured_companies(context)

    if not companies:
        return vendor_master.copy(), 0

    company = vendor_master["Company"].map(normalize_company)
    included = company.isin(companies)

    return (
        vendor_master.loc[included].copy().reset_index(drop=True),
        int((~included).sum()),
    )


def _require_unique(
    dataframe: pd.DataFrame,
    columns: list[str],
    label: str,
) -> None:
    duplicated = dataframe.duplicated(columns, keep=False)

    if duplicated.any():
        examples = (
            dataframe.loc[duplicated, columns]
            .drop_duplicates()
            .head(20)
            .to_dict("records")
        )
        raise ValueError(
            f"{CONTROL_ID}: {label} is not unique by {columns}. "
            f"Examples: {examples}"
        )


def _load_population(
    context: dict[str, Any],
) -> tuple[pd.DataFrame, dict[str, Any]]:
    """
    Read each input once and perform the one validated invoice merge.

    No business exclusion is recreated locally.
    """
    vendor_source = load_vm_vendors(context)
    vendor_master = build_vendor_master_population(vendor_source)

    vendor_master, excluded_company = _filter_configured_companies(
        vendor_master,
        context,
    )

    valid_population, population_metrics = get_valid_vendor_population(
        vendor_master
    )

    if valid_population.empty:
        raise ValueError(
            f"{CONTROL_ID}: valid vendor population is empty after CONFIG "
            "company and common VM exclusion rules."
        )

    _require_unique(
        valid_population,
        ["Company", "Vendor Code"],
        "valid vendor population",
    )

    postings, posting_metadata = load_vm_vendor_postings(context)

    if (
        postings is None
        or not posting_metadata.get("available", False)
    ):
        raise FileNotFoundError(
            "VM_003 requires VPBSIK and VPBSAK to populate "
            "the last-invoice display columns."
        )

    last_invoices = build_vm_last_invoice_population(postings)

    _require_unique(
        last_invoices,
        ["Company", "Vendor Code"],
        "last-invoice population",
    )

    missing_invoice_columns = [
        column
        for column in INVOICE_COLUMNS
        if column not in last_invoices.columns
    ]
    if missing_invoice_columns:
        raise ValueError(
            f"{CONTROL_ID}: common last-invoice population is missing "
            f"columns: {missing_invoice_columns}"
        )

    rows_before_merge = len(valid_population)

    enriched = valid_population.merge(
        last_invoices[
            ["Company", "Vendor Code", *INVOICE_COLUMNS]
        ],
        how="left",
        on=["Company", "Vendor Code"],
        validate="one_to_one",
    )

    if len(enriched) != rows_before_merge:
        raise AssertionError(
            "VM03 last-invoice enrichment changed vendor population."
        )

    # Numeric amount is intentionally retained as numeric. No FX is performed.
    enriched["Last Inv Amt Doc Currency"] = pd.to_numeric(
        enriched["Last Inv Amt Doc Currency"],
        errors="coerce",
    )

    document_type = (
        postings["Document Type"]
        .astype("string")
        .fillna("")
        .str.strip()
        .str.upper()
    )

    metrics = {
        "source_rows": len(vendor_source),
        "master_rows": len(vendor_master),
        "valid_rows": len(valid_population),
        "excluded_company": excluded_company,
        "posting_rows": len(postings),
        "invoice_posting_rows": int(document_type.isin(["RE", "KR"]).sum()),
        "vendors_with_last_invoice": int(
            enriched["Last Invoice Number"]
            .astype("string")
            .fillna("")
            .str.strip()
            .ne("")
            .sum()
        ),
        "warnings": [
            *population_metrics.get("warnings", []),
            *posting_metadata.get("warnings", []),
        ],
    }

    return enriched, metrics


def build_vm_003(
    vendor_population: pd.DataFrame,
) -> tuple[pd.DataFrame, dict[str, int]]:
    """Build exact Phone1 groups using vectorized pandas operations."""
    required = {
        "Company",
        "Company Name",
        "Vendor Code",
        "Vendor Name",
        "Phone1",
        *INVOICE_COLUMNS,
    }
    missing = sorted(required.difference(vendor_population.columns))

    if missing:
        raise ValueError(
            f"{CONTROL_ID}: vendor population is missing columns: {missing}"
        )

    _require_unique(
        vendor_population,
        ["Company", "Vendor Code"],
        "enriched vendor population",
    )

    data = vendor_population.copy()

    # rule: null -> blank; retain digits only. No minimum length is added.
    data["_Normalized Phone"] = (
        data["Phone1"]
        .astype("string")
        .fillna("")
        .str.replace(r"\D+", "", regex=True)
    )

    comparable = data.loc[
        data["_Normalized Phone"].ne("")
    ].copy()

    if comparable.empty:
        return (
            pd.DataFrame(columns=OUTPUT_COLUMNS),
            {
                "nonblank_phone_rows": 0,
                "exception_rows": 0,
                "groups": 0,
                "duplicate_output_rows": 0,
            },
        )

    company_code = (
        comparable["Company"]
        .astype("string")
        .fillna("")
    )

    comparable["_Group Key"] = company_code.str.cat(
        comparable["_Normalized Phone"],
        sep="\u00a6",
    )

    distinct_vendors = (
        comparable.groupby(
            "_Group Key",
            sort=False,
            observed=True,
        )["Vendor Code"]
        .transform("nunique")
    )

    exceptions = comparable.loc[
        distinct_vendors.ge(2)
    ].copy()

    if exceptions.empty:
        return (
            pd.DataFrame(columns=OUTPUT_COLUMNS),
            {
                "nonblank_phone_rows": len(comparable),
                "exception_rows": 0,
                "groups": 0,
                "duplicate_output_rows": 0,
            },
        )

    exceptions["Group"] = (
        pd.factorize(
            exceptions["_Group Key"],
            sort=True,
        )[0]
        + 1
    )

    # Company/CoCo replacement occurs only after the posting merge.
    exceptions["CoCo"] = (
        exceptions["Company"]
        .astype("string")
        .fillna("")
    )
    exceptions["Company"] = (
        exceptions["Company Name"]
        .astype("string")
        .fillna("")
        .str.strip()
    )

    for column in INVOICE_COLUMNS:
        if column not in exceptions.columns:
            exceptions[column] = pd.NA

    for column in (
        "Company",
        "CoCo",
        "Vendor Code",
        "Vendor Name",
        "Phone1",
        "Last Invoice Number",
        "Last Inv Amt Doc Currency Indicator",
    ):
        exceptions[column] = (
            exceptions[column]
            .astype("string")
            .fillna("")
        )

    exceptions["Last Inv Amt Doc Currency"] = pd.to_numeric(
        exceptions["Last Inv Amt Doc Currency"],
        errors="coerce",
    )

    output = (
        exceptions.sort_values(
            ["Group", "CoCo", "Vendor Code"],
            kind="stable",
        )
        .loc[:, OUTPUT_COLUMNS]
        .reset_index(drop=True)
    )

    duplicate_output_rows = int(
        output.duplicated(
            ["CoCo", "Vendor Code"],
            keep=False,
        ).sum()
    )

    if duplicate_output_rows:
        raise AssertionError(
            "VM03 output contains duplicate CoCo/Vendor Code rows."
        )

    # Validate every group in one vectorized aggregation.
    validation = (
        exceptions.groupby(
            "Group",
            sort=False,
            observed=True,
        )
        .agg(
            companies=("CoCo", "nunique"),
            vendors=("Vendor Code", "nunique"),
            phones=("_Normalized Phone", "nunique"),
            blank_phone=(
                "_Normalized Phone",
                lambda series: bool(series.eq("").any()),
            ),
        )
    )

    if validation["vendors"].lt(2).any():
        raise AssertionError(
            "VM03 contains a Group with fewer than two distinct vendors."
        )

    if validation["companies"].ne(1).any():
        raise AssertionError(
            "VM03 contains a Group crossing CoCo."
        )

    if validation["phones"].ne(1).any():
        raise AssertionError(
            "VM03 contains a Group with different normalized phones."
        )

    if validation["blank_phone"].any():
        raise AssertionError(
            "VM03 contains an exception with a blank normalized phone."
        )

    if list(output.columns) != OUTPUT_COLUMNS:
        raise AssertionError(
            "VM03 final columns or column order do not match"
        )

    return (
        output,
        {
            "nonblank_phone_rows": len(comparable),
            "exception_rows": len(output),
            "groups": int(output["Group"].nunique()),
            "duplicate_output_rows": duplicate_output_rows,
        },
    )


def run_vm_003(
    context: dict[str, Any],
) -> dict[str, Any]:
    """Execute VM03 and replace only the VM03 result worksheet."""
    total_started = perf_counter()
    stage_started = total_started

    population, load_metrics = _load_population(context)
    loaded_at = perf_counter()
    load_seconds = loaded_at - stage_started

    stage_started = loaded_at
    # Population preparation is already completed by the common builders.
    prepared_at = perf_counter()
    preparation_seconds = prepared_at - stage_started

    stage_started = prepared_at
    output, analytic_metrics = build_vm_003(population)
    analytic_at = perf_counter()
    analytic_seconds = analytic_at - stage_started

    # Phone normalization is included in analytic_seconds and performs no
    # additional pass solely for timing.
    normalization_seconds = analytic_seconds

    fx_seconds = 0.0

    stage_started = analytic_at
    output_file = write_vm_control_sheet(
        context=context,
        sheet_name=SHEET_NAME,
        dataframe=output,
        date_columns=["Last Transaction Date"],
        amount_columns=["Last Inv Amt Doc Currency"],
        integer_columns=["Group"],
    )
    finished = perf_counter()
    write_seconds = finished - stage_started

    print(f"{CONTROL_ID} source vendor rows: {load_metrics['source_rows']}")
    print(f"{CONTROL_ID} master vendor rows: {load_metrics['master_rows']}")
    print(f"{CONTROL_ID} valid vendor rows: {load_metrics['valid_rows']}")
    print(f"{CONTROL_ID} posting rows: {load_metrics['posting_rows']}")
    print(
        f"{CONTROL_ID} RE/KR invoice posting rows: "
        f"{load_metrics['invoice_posting_rows']}"
    )
    print(
        f"{CONTROL_ID} vendors with last invoice: "
        f"{load_metrics['vendors_with_last_invoice']}"
    )
    print(
        f"{CONTROL_ID} rows with nonblank normalized phone: "
        f"{analytic_metrics['nonblank_phone_rows']}"
    )
    print(f"{CONTROL_ID} exception rows: {analytic_metrics['exception_rows']}")
    print(f"{CONTROL_ID} groups: {analytic_metrics['groups']}")
    print(
        f"{CONTROL_ID} duplicate CoCo/Vendor rows: "
        f"{analytic_metrics['duplicate_output_rows']}"
    )
    print(
        f"{CONTROL_ID} timings: "
        f"load={load_seconds:.2f}s, "
        f"preparation={preparation_seconds:.2f}s, "
        f"phone_normalization_and_analytic={normalization_seconds:.2f}s, "
        f"FX={fx_seconds:.2f}s, "
        f"write={write_seconds:.2f}s, "
        f"total={finished - total_started:.2f}s"
    )

    for warning in load_metrics["warnings"]:
        print(f"WARNING: {warning}")

    return {
        "status": "ERROR" if not output.empty else "OK",
        "output_file": output_file,
        "sheet_name": SHEET_NAME,
        "rows": len(output),
    }
