"""PO01 - potential split purchase orders.

The control identifies different purchase orders with the same company,
vendor, material and creator when their document dates are within PARAM1 days
(default: 7). Lines without material are outside the LHA material population.
Orders without a trustworthy creator are excluded and measured. Creator
coverage below 90% is reported as a scope warning, but it does not prevent the
control from analysing the eligible population.
"""

import re

import pandas as pd

from core.po_common import load_po_lines, normalize_text, write_control_sheet


CONTROL_ID = "PO_001"
CONTROL_NAME = "Split Purchase Orders"
SHEET_NAME = "PO01"
DEFAULT_SPLIT_WINDOW_DAYS = 7
MINIMUM_CREATOR_COVERAGE = 0.90


GROUP_FIELDS = [
    "Company",
    "Vendor Code",
    "Item Code",
    "PO Creator ID",
]


DOCUMENT_FIELDS = GROUP_FIELDS + [
    "PO Number",
    "PO Doc Date",
]


SUMMARY_COLUMNS = [
    "Record Type",
    "Control",
    "Control Name",
    "Control Result",
    "Message",
    "Split Window Days",
    "Minimum Creator Coverage",
    "Rows Read",
    "Residual Rows",
    "Rows After Config Filters",
    "Distinct POs",
    "Rows With Creator",
    "Rows Without Creator",
    "Creator Line Coverage",
    "POs With Creator",
    "POs Without Creator",
    "Creator PO Coverage",
    "Rows Without Material",
    "Rows Without Vendor",
    "Rows In Material Vendor Scope",
    "Rows Eligible For Split Logic",
    "Rows Excluded By Split Logic",
    "Exception Rows",
    "Exception POs",
]


def parse_split_window(context):
    """Return PARAM1 as a non-negative integer number of days."""
    raw_value = context.get("control", {}).get("param1", "")
    text = normalize_text(raw_value)

    if text == "":
        return DEFAULT_SPLIT_WINDOW_DAYS

    if not re.fullmatch(r"\d+(?:\.0+)?", text):
        raise ValueError(
            f"{CONTROL_ID} PARAM1 must be a non-negative whole number of days; "
            f"received {raw_value!r}."
        )

    window = int(float(text))

    if window < 0:
        raise ValueError(f"{CONTROL_ID} PARAM1 cannot be negative.")

    return window


def normalize_creator_series(series):
    """Normalize creator values without inventing missing users."""
    return series.map(normalize_text)


def prepare_creator_population(dataframe):
    """Resolve safe within-PO creator blanks and calculate coverage metrics."""
    prepared = dataframe.copy()
    prepared["PO Creator ID"] = normalize_creator_series(
        prepared["PO Creator ID"]
    )

    po_key = ["Company", "PO Number"]

    known = prepared[
        prepared["PO Creator ID"].ne("")
    ]

    creator_counts = known.groupby(
        po_key,
        dropna=False,
    )["PO Creator ID"].nunique()

    conflicting = creator_counts[
        creator_counts > 1
    ]

    if not conflicting.empty:
        raise ValueError(
            f"{CONTROL_ID} found {len(conflicting)} POs with more than one "
            "nonblank creator. Creator cannot be resolved safely."
        )

    creator_for_group = prepared[
        "PO Creator ID"
    ].replace("", pd.NA)

    resolved_creator = (
        prepared.assign(
            _CREATOR=creator_for_group
        )
        .groupby(
            po_key,
            dropna=False,
        )["_CREATOR"]
        .transform("first")
        .fillna("")
    )

    prepared["PO Creator ID"] = resolved_creator.map(
        normalize_text
    )

    line_has_creator = prepared[
        "PO Creator ID"
    ].ne("")

    po_creator_status = (
        prepared.assign(
            _HAS_CREATOR=line_has_creator
        )
        .groupby(
            po_key,
            dropna=False,
        )["_HAS_CREATOR"]
        .any()
    )

    total_rows = len(prepared)
    total_pos = len(po_creator_status)
    rows_with_creator = int(
        line_has_creator.sum()
    )
    pos_with_creator = int(
        po_creator_status.sum()
    )

    line_coverage = (
        rows_with_creator / total_rows
        if total_rows
        else 0.0
    )

    po_coverage = (
        pos_with_creator / total_pos
        if total_pos
        else 0.0
    )

    metrics = {
        "distinct_pos": total_pos,
        "rows_with_creator": rows_with_creator,
        "rows_without_creator": total_rows - rows_with_creator,
        "creator_line_coverage": line_coverage,
        "pos_with_creator": pos_with_creator,
        "pos_without_creator": total_pos - pos_with_creator,
        "creator_po_coverage": po_coverage,
    }

    return prepared, metrics


def creator_coverage_is_sufficient(metrics):
    """Return True only when line and PO coverage both meet the reference."""
    return (
        metrics["creator_line_coverage"]
        >= MINIMUM_CREATOR_COVERAGE
        and metrics["creator_po_coverage"]
        >= MINIMUM_CREATOR_COVERAGE
    )


def prepare_split_scope(dataframe):
    """Return the LHA material/vendor scope before measuring creator coverage."""
    prepared = dataframe.copy()

    prepared["Vendor Code"] = prepared[
        "Vendor Code"
    ].map(normalize_text)

    prepared["Item Code"] = prepared[
        "Item Code"
    ].map(normalize_text)

    missing_material = prepared[
        "Item Code"
    ].eq("")

    missing_vendor = prepared[
        "Vendor Code"
    ].eq("")

    scope_mask = ~(
        missing_material
        | missing_vendor
    )

    scope = prepared.loc[
        scope_mask
    ].copy()

    metrics = {
        "rows_without_material": int(
            missing_material.sum()
        ),
        "rows_without_vendor": int(
            missing_vendor.sum()
        ),
        "rows_in_material_vendor_scope": len(scope),
        "rows_eligible_for_logic": 0,
        "rows_excluded_by_logic": int(
            (~scope_mask).sum()
        ),
    }

    return scope, metrics


def exclude_missing_creators(
    dataframe,
    logic_metrics,
):
    """Exclude creatorless rows after coverage is measured in the PO01 scope."""
    prepared = dataframe.copy()

    prepared["PO Creator ID"] = prepared[
        "PO Creator ID"
    ].map(normalize_text)

    creator_mask = prepared[
        "PO Creator ID"
    ].ne("")

    eligible = prepared.loc[
        creator_mask
    ].copy()

    metrics = dict(logic_metrics)

    metrics[
        "rows_eligible_for_logic"
    ] = len(eligible)

    metrics[
        "rows_excluded_by_logic"
    ] += int((~creator_mask).sum())

    return eligible, metrics


def find_split_documents(
    dataframe,
    window_days,
):
    """Return group/document keys having another PO within the date window."""
    if dataframe.empty:
        return pd.DataFrame(
            columns=GROUP_FIELDS
            + ["PO Number"]
        )

    documents = (
        dataframe[DOCUMENT_FIELDS]
        .drop_duplicates()
        .sort_values(
            GROUP_FIELDS
            + [
                "PO Doc Date",
                "PO Number",
            ]
        )
        .reset_index(drop=True)
    )

    flagged_keys = []

    grouped_documents = documents.groupby(
        GROUP_FIELDS,
        sort=False,
        dropna=False,
    )

    for group_values, group in grouped_documents:
        group = (
            group.sort_values(
                [
                    "PO Doc Date",
                    "PO Number",
                ]
            )
            .reset_index(drop=True)
        )

        if len(group) < 2:
            continue

        dates = group[
            "PO Doc Date"
        ].tolist()

        flagged_positions = set()

        for index in range(
            1,
            len(group),
        ):
            difference = abs(
                (
                    dates[index]
                    - dates[index - 1]
                ).days
            )

            if difference <= window_days:
                flagged_positions.add(
                    index - 1
                )
                flagged_positions.add(
                    index
                )

        if not flagged_positions:
            continue

        group_tuple = (
            group_values
            if isinstance(
                group_values,
                tuple,
            )
            else (group_values,)
        )

        for index in sorted(
            flagged_positions
        ):
            flagged_keys.append(
                group_tuple
                + (
                    group.at[
                        index,
                        "PO Number",
                    ],
                )
            )

    return pd.DataFrame(
        flagged_keys,
        columns=GROUP_FIELDS
        + ["PO Number"],
    ).drop_duplicates()


def build_po01_exceptions(
    dataframe,
    window_days,
):
    """Return all detail lines belonging to POs flagged by PO01."""
    flagged_documents = find_split_documents(
        dataframe,
        window_days,
    )

    if flagged_documents.empty:
        empty = dataframe.iloc[
            0:0
        ].copy()

        empty["SPLIT_PO_KEY"] = pd.Series(
            dtype="object"
        )

        empty["Split Window Days"] = pd.Series(
            dtype="int64"
        )

        return empty

    exceptions = dataframe.merge(
        flagged_documents,
        on=GROUP_FIELDS
        + ["PO Number"],
        how="inner",
        validate="many_to_one",
    )

    exceptions["SPLIT_PO_KEY"] = (
        exceptions[GROUP_FIELDS]
        .astype(str)
        .agg(
            "|".join,
            axis=1,
        )
    )

    exceptions[
        "Split Window Days"
    ] = window_days

    return (
        exceptions.sort_values(
            GROUP_FIELDS
            + [
                "PO Doc Date",
                "PO Number",
                "PO Line",
            ]
        )
        .reset_index(drop=True)
    )


def build_summary_row(
    control_result,
    message,
    window_days,
    input_metrics,
    creator_metrics,
    logic_metrics,
    exceptions,
):
    """Return one transparent PO01 summary row."""
    exception_pos = (
        exceptions[
            [
                "Company",
                "PO Number",
            ]
        ]
        .drop_duplicates()
        .shape[0]
        if not exceptions.empty
        else 0
    )

    return {
        "Record Type": "SUMMARY",
        "Control": CONTROL_ID,
        "Control Name": CONTROL_NAME,
        "Control Result": control_result,
        "Message": message,
        "Split Window Days": window_days,
        "Minimum Creator Coverage": MINIMUM_CREATOR_COVERAGE,
        "Rows Read": input_metrics["rows_read"],
        "Residual Rows": input_metrics["residual_rows"],
        "Rows After Config Filters": (
            input_metrics[
                "rows_after_config_filters"
            ]
        ),
        "Distinct POs": creator_metrics[
            "distinct_pos"
        ],
        "Rows With Creator": creator_metrics[
            "rows_with_creator"
        ],
        "Rows Without Creator": creator_metrics[
            "rows_without_creator"
        ],
        "Creator Line Coverage": creator_metrics[
            "creator_line_coverage"
        ],
        "POs With Creator": creator_metrics[
            "pos_with_creator"
        ],
        "POs Without Creator": creator_metrics[
            "pos_without_creator"
        ],
        "Creator PO Coverage": creator_metrics[
            "creator_po_coverage"
        ],
        "Rows Without Material": logic_metrics[
            "rows_without_material"
        ],
        "Rows Without Vendor": logic_metrics[
            "rows_without_vendor"
        ],
        "Rows In Material Vendor Scope": logic_metrics[
            "rows_in_material_vendor_scope"
        ],
        "Rows Eligible For Split Logic": logic_metrics[
            "rows_eligible_for_logic"
        ],
        "Rows Excluded By Split Logic": logic_metrics[
            "rows_excluded_by_logic"
        ],
        "Exception Rows": len(exceptions),
        "Exception POs": exception_pos,
    }


def build_output_dataframe(
    summary_row,
    exceptions,
):
    """Combine one summary record with zero or more exception detail records."""
    summary = pd.DataFrame(
        [summary_row],
        columns=SUMMARY_COLUMNS,
    )

    if exceptions.empty:
        return summary

    details = exceptions.copy()

    details.insert(
        0,
        "Record Type",
        "EXCEPTION",
    )

    details.insert(
        1,
        "Control",
        CONTROL_ID,
    )

    details.insert(
        2,
        "Control Name",
        CONTROL_NAME,
    )

    details.insert(
        3,
        "Control Result",
        "ERROR",
    )

    details.insert(
        4,
        "Message",
        "Potential split purchase order",
    )

    return pd.concat(
        [
            summary,
            details,
        ],
        ignore_index=True,
        sort=False,
    )


def print_po01_metrics(
    summary_row,
    input_metrics,
):
    """Print population and coverage diagnostics to the run log."""
    print()
    print(
        f"{CONTROL_ID} - {CONTROL_NAME}"
    )

    print(
        "-"
        * (
            len(CONTROL_ID)
            + len(CONTROL_NAME)
            + 3
        )
    )

    print(
        "Input file: "
        f"{input_metrics['input_file']}"
    )

    print(
        "Input sheet: "
        f"{input_metrics['input_sheet']}"
    )

    print(
        "Header row: "
        f"{input_metrics['header_row']}"
    )

    print(
        "Rows read: "
        f"{summary_row['Rows Read']}"
    )

    print(
        "Residual rows: "
        f"{summary_row['Residual Rows']}"
    )

    print(
        "Excluded by company: "
        f"{input_metrics['excluded_by_company']}"
    )

    print(
        "Excluded by date: "
        f"{input_metrics['excluded_by_date']}"
    )

    print(
        "Rows after CONFIG filters: "
        f"{summary_row['Rows After Config Filters']}"
    )

    print(
        "Rows without creator: "
        f"{summary_row['Rows Without Creator']}"
    )

    print(
        "Creator line coverage: "
        f"{summary_row['Creator Line Coverage']:.2%}"
    )

    print(
        "POs without creator: "
        f"{summary_row['POs Without Creator']}"
    )

    print(
        "Creator PO coverage: "
        f"{summary_row['Creator PO Coverage']:.2%}"
    )

    print(
        "Rows without material: "
        f"{summary_row['Rows Without Material']}"
    )

    print(
        "Rows without vendor: "
        f"{summary_row['Rows Without Vendor']}"
    )

    print(
        "Rows in material/vendor scope: "
        f"{summary_row['Rows In Material Vendor Scope']}"
    )

    print(
        "Rows eligible for logic: "
        f"{summary_row['Rows Eligible For Split Logic']}"
    )

    print(
        "Exception rows: "
        f"{summary_row['Exception Rows']}"
    )

    print(
        "Exception POs: "
        f"{summary_row['Exception POs']}"
    )

    print(
        "Control result: "
        f"{summary_row['Control Result']}"
    )

    print(
        "Message: "
        f"{summary_row['Message']}"
    )

    print()


def run_po_001(context):
    """Execute PO01 independently and replace only the PO01 output sheet."""
    window_days = parse_split_window(
        context
    )

    po_lines, input_metrics = load_po_lines(
        context
    )

    split_scope, logic_metrics = prepare_split_scope(
        po_lines
    )

    creator_population, creator_metrics = (
        prepare_creator_population(
            split_scope
        )
    )

    eligible, logic_metrics = exclude_missing_creators(
        creator_population,
        logic_metrics,
    )

    coverage_sufficient = (
        creator_coverage_is_sufficient(
            creator_metrics
        )
    )

    exceptions = build_po01_exceptions(
        eligible,
        window_days,
    )

    control_result = (
        "ERROR"
        if not exceptions.empty
        else "OK"
    )

    result_message = (
        f"Potential split POs found: "
        f"{len(exceptions)} detail rows."
        if not exceptions.empty
        else (
            "No potential split POs found "
            "in the eligible population."
        )
    )

    coverage_message = (
        " Creator coverage is below the 90% "
        "reference threshold; the control analysed "
        f"{len(eligible)} eligible rows and excluded "
        f"{creator_metrics['rows_without_creator']} "
        "rows without creator."
        if not coverage_sufficient
        else ""
    )

    message = (
        result_message
        + coverage_message
    )

    summary_row = build_summary_row(
        control_result=control_result,
        message=message,
        window_days=window_days,
        input_metrics=input_metrics,
        creator_metrics=creator_metrics,
        logic_metrics=logic_metrics,
        exceptions=exceptions,
    )

    output = build_output_dataframe(
        summary_row,
        exceptions,
    )

    output_file = write_control_sheet(
        context=context,
        sheet_name=SHEET_NAME,
        dataframe=output,
        date_columns=[
            "PO Doc Date",
        ],
        amount_columns=[
            "PO Quantity",
            "PO Unit Price",
            "PO Line Total",
            "Creator Line Coverage",
            "Creator PO Coverage",
            "Minimum Creator Coverage",
        ],
    )

    print_po01_metrics(
        summary_row,
        input_metrics,
    )

    print(
        f"PO01 output file: {output_file}"
    )

    print(
        f"PO01 output sheet: {SHEET_NAME}"
    )

    return {
        "status": control_result,
        "output_file": output_file,
        "sheet_name": SHEET_NAME,
        "rows": len(exceptions),
        "summary": summary_row,
    }
