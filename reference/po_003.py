"""PO03 - purchase order created and approved by the same user.

The control is independent from every other PO control. It reads the original
PO Lines, CDHDR and CDPOS exports, identifies the highest effective ECC release
event for each purchase order (the ECC equivalent of the highest approval step
used by the LHA implementation), and compares technical creator/approver IDs.

An effective release is a change to EKKO-FRGZU where the new release-status
value retains every release mark present in the old value and adds at least one
new mark. Reset/reversal, strategy-change and non-release events are measured
but excluded. Among effective events, the event with the greatest resulting
number of release marks is selected for each PO; timestamp and change-document
number provide deterministic tie breakers. Exception POs are then joined to
PO Lines so all original lines remain available for review.
"""

import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from core.po_common import (
    ALLOWED_INPUT_EXTENSIONS,
    get_period_suffix,
    inspect_input_workbook,
    load_po_lines,
    normalize_identifier,
    normalize_lookup,
    normalize_text,
    resolve_sheet_name,
    write_control_sheet,
)


CONTROL_ID = "PO_003"
CONTROL_NAME = "Purchase Order Created And Approved By The Same User"
SHEET_NAME = "PO03"
CDHDR_INPUT_PREFIX = "LBR PO CDHDR"
CDPOS_INPUT_PREFIX = "LBR PO CDPOS"
INPUT_SHEET = "Sheet1"
HEADER_ROW = 1
CHANGE_OBJECT = "EINKBELEG"
HEADER_TABLE = "EKKO"
RELEASE_STATUS_FIELD = "FRGZU"

PO_REQUIRED_FIELDS = (
    "Company",
    "PO Number",
    "PO Line",
    "Vendor Code",
    "PO Doc Date",
    "PO Creator ID",
    "Item Code",
    "PO Quantity",
    "PO Unit Price",
    "PO Doc Currency",
    "PO Line Total",
    "PO Material Description",
    "PO Line Deleted",
    "PO Delivery Completed",
    "PR Number",
    "PR Line",
)

CDHDR_ALIASES = {
    "Change Object": (
        "Change doc. object",
        "OBJECTCLAS",
        "CDHDR-OBJECTCLAS",
    ),
    "Object Value": (
        "Object value",
        "OBJECTID",
        "CDHDR-OBJECTID",
    ),
    "Change Number": (
        "Document number",
        "CHANGENR",
        "CDHDR-CHANGENR",
    ),
    "Approver ID": (
        "User",
        "USERNAME",
        "CDHDR-USERNAME",
    ),
    "Approval Date": (
        "Date",
        "UDATE",
        "CDHDR-UDATE",
    ),
    "Approval Time": (
        "Time",
        "UTIME",
        "CDHDR-UTIME",
    ),
    "Transaction Code": (
        "Transaction Code",
        "TCODE",
        "CDHDR-TCODE",
    ),
}

CDPOS_ALIASES = {
    "Change Object": (
        "Change doc. object",
        "OBJECTCLAS",
        "CDPOS-OBJECTCLAS",
    ),
    "Object Value": (
        "Object value",
        "OBJECTID",
        "CDPOS-OBJECTID",
    ),
    "Change Number": (
        "Document number",
        "CHANGENR",
        "CDPOS-CHANGENR",
    ),
    "Table Name": (
        "Table Name",
        "TABNAME",
        "CDPOS-TABNAME",
    ),
    "Table Key": (
        "Table Key",
        "TABKEY",
        "CDPOS-TABKEY",
    ),
    "Field Name": (
        "Field Name",
        "FNAME",
        "CDPOS-FNAME",
    ),
    "Change ID": (
        "Change ID",
        "CHNGIND",
        "CDPOS-CHNGIND",
    ),
    "New Value": (
        "New value",
        "VALUE_NEW",
        "CDPOS-VALUE_NEW",
    ),
    "Old Value": (
        "Old value",
        "VALUE_OLD",
        "CDPOS-VALUE_OLD",
    ),
}

CHANGE_KEY = [
    "Change Object",
    "Object Value",
    "Change Number",
]

CDPOS_KEY = CHANGE_KEY + [
    "Table Name",
    "Table Key",
    "Field Name",
]

PO_KEY = [
    "Company",
    "PO Number",
]

SUMMARY_COLUMNS = [
    "Record Type",
    "Control",
    "Control Name",
    "Control Result",
    "Message",
    "PO Lines Rows Read",
    "PO Lines Residual Rows",
    "Rows After Config Filters",
    "Distinct POs Analyzed",
    "Rows Without Creator",
    "POs Without Creator",
    "CDHDR Rows Read",
    "CDHDR Residual Rows",
    "CDPOS Rows Read",
    "CDPOS Residual Rows",
    "Approval Events Read",
    "Approval Events Eligible",
    "Approval Events Excluded",
    "Release Reversal Events",
    "Duplicate Approval Events",
    "Approval Events Outside PO Population",
    "POs Without Approval Evidence",
    "POs With Approval Evidence",
    "Exception Rows",
    "Exception POs",
    "Distinct Matching Users",
]

LHA_REPORT_COLUMNS = [
    "CoCo",
    "Company",
    "PO Number",
    "PO DocEntry",
    "PO Line",
    "Vendor Code",
    "Vendor Name",
    "PO Doc Date",
    "PO Doc Currency",
    "Company Main Currency",
    "PO Canceled",
    "PO Line Status",
    "Item Code",
    "Account Code",
    "PO Material Description",
    "PO Quantity",
    "PO Unit Price",
    "PO Line Total",
    "PO Line Total USD",
    "USD Rate",
    "USD Rate Date",
    "PO Creator ID",
    "PO Creator Name",
    "PO Approval Date",
    "PO Approver ID",
    "PO Approver Name",
    "PO Approval Status",
    "GR Doc Number",
    "GR Doc Date",
    "GR First Posting Date",
    "GR Last Posting Date",
    "GR Quantity",
    "GR Creator ID",
    "GR Creator Name",
    "PO Month",
    "PR DocEntry",
    "PR Line",
    "From PR",
]


def _resolve_columns(dataframe, aliases, source_name):
    """Resolve every required logical field against observed SAP headers."""
    available = {}

    for column in dataframe.columns:
        available.setdefault(
            normalize_lookup(column),
            [],
        ).append(column)

    mapping = {}
    missing = []
    ambiguous = []

    for logical, candidates in aliases.items():
        matches = []

        for candidate in candidates:
            matches.extend(
                available.get(
                    normalize_lookup(candidate),
                    [],
                )
            )

        matches = list(dict.fromkeys(matches))

        if len(matches) == 1:
            mapping[logical] = matches[0]
        elif not matches:
            missing.append(logical)
        else:
            ambiguous.append(
                f"{logical}: {matches}"
            )

    if missing or ambiguous:
        problems = []

        if missing:
            problems.append(
                f"missing required fields: {missing}"
            )

        if ambiguous:
            problems.append(
                f"ambiguous fields: {ambiguous}"
            )

        raise ValueError(
            f"{source_name} header validation failed "
            f"({'; '.join(problems)}). "
            f"Available headers: {list(dataframe.columns)}"
        )

    return mapping


def _find_change_input(context, prefix):
    """Require exactly one period-specific original change-document input."""
    input_folder = Path(
        context["input_folder"]
    )

    if not input_folder.is_dir():
        raise FileNotFoundError(
            f"PO input folder was not found: {input_folder}"
        )

    suffix = get_period_suffix(
        context["module"]
    )

    expected = normalize_lookup(
        f"{prefix} {suffix}"
    )

    matches = sorted(
        path
        for path in input_folder.rglob("*")
        if path.is_file()
        and not path.name.startswith("~$")
        and path.suffix.casefold() in ALLOWED_INPUT_EXTENSIONS
        and normalize_lookup(path.stem) == expected
    )

    if not matches:
        raise FileNotFoundError(
            f"{prefix} input was not found. "
            "Expected exactly one file "
            f"equivalent to '{prefix}_{suffix}.XLSX' "
            f"below {input_folder}."
        )

    if len(matches) > 1:
        raise ValueError(
            f"Multiple {prefix} inputs were found for {suffix}; "
            "expected exactly one:\n"
            + "\n".join(
                f"- {path}"
                for path in matches
            )
        )

    return matches[0]


def _remove_change_residuals(dataframe, physical_key, source_name):
    """Remove fully blank export footer rows and reject partial join keys."""
    blank = pd.DataFrame(
        {
            column: dataframe[column].map(normalize_text).eq("")
            for column in physical_key
        },
        index=dataframe.index,
    )

    residual = blank.all(axis=1)
    partial = blank.any(axis=1) & ~residual

    if partial.any():
        raise ValueError(
            f"{source_name} contains {int(partial.sum())} rows "
            "with a partially blank change-document key."
        )

    return (
        dataframe.loc[~residual].copy(),
        int(residual.sum()),
    )


def _load_change_input(context, prefix, aliases):
    """Read and validate one original CDHDR/CDPOS workbook."""
    path = _find_change_input(
        context,
        prefix,
    )

    sheet = resolve_sheet_name(
        path,
        INPUT_SHEET,
    )

    workbook_metrics = inspect_input_workbook(
        path,
        sheet,
    )

    try:
        raw = pd.read_excel(
            path,
            sheet_name=sheet,
            header=HEADER_ROW - 1,
            dtype=object,
        )
    except PermissionError as error:
        raise PermissionError(
            f"{prefix} input is open or locked: {path}. "
            "Close Excel and retry."
        ) from error

    raw = (
        raw
        .dropna(axis=0, how="all")
        .dropna(axis=1, how="all")
    )

    if raw.empty:
        raise ValueError(
            f"{prefix} sheet '{sheet}' contains no data rows."
        )

    columns = _resolve_columns(
        raw,
        aliases,
        prefix,
    )

    logical_key = (
        CDPOS_KEY
        if "Table Name" in aliases
        else CHANGE_KEY
    )

    physical_key = [
        columns[name]
        for name in logical_key
    ]

    population, residual_rows = _remove_change_residuals(
        raw,
        physical_key,
        prefix,
    )

    if population.empty:
        raise ValueError(
            f"{prefix} contains no valid rows after residual removal."
        )

    renamed = population.rename(
        columns={
            physical: logical
            for logical, physical in columns.items()
        }
    )

    for column in renamed.columns:
        if column in aliases:
            renamed[column] = (
                renamed[column]
                .map(normalize_text)
            )

    for column in (
        "Object Value",
        "Change Number",
    ):
        renamed[column] = (
            renamed[column]
            .map(normalize_identifier)
        )

    metrics = {
        **workbook_metrics,
        "input_file": str(path),
        "input_sheet": sheet,
        "header_row": HEADER_ROW,
        "rows_read": len(raw),
        "residual_rows": residual_rows,
        "rows_after_residuals": len(renamed),
    }

    return renamed, metrics


def _validate_unique(dataframe, key, source_name):
    """Fail when a physical source key is duplicated."""
    duplicated = dataframe.duplicated(
        key,
        keep=False,
    )

    if duplicated.any():
        examples = (
            dataframe.loc[
                duplicated,
                key,
            ]
            .drop_duplicates()
            .head(5)
        )

        raise ValueError(
            f"{source_name} contains {int(duplicated.sum())} rows "
            f"participating in duplicate physical keys {key}. "
            f"Examples: {examples.to_dict('records')}"
        )


def _parse_approval_timestamp(cdhdr):
    """Parse separate SAP date/time fields without silently losing evidence."""

    def parse_date(value):
        if isinstance(
            value,
            (
                pd.Timestamp,
                datetime,
                date,
            ),
        ):
            return pd.Timestamp(value).normalize()

        text = normalize_text(value)

        if text == "":
            return pd.NaT

        formats = (
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%d",
            "%Y%m%d",
            "%d.%m.%Y %H:%M:%S",
            "%d.%m.%Y",
            "%d/%m/%Y %H:%M:%S",
            "%d/%m/%Y",
            "%d-%m-%Y %H:%M:%S",
            "%d-%m-%Y",
        )

        for date_format in formats:
            parsed = pd.to_datetime(
                text,
                format=date_format,
                errors="coerce",
            )

            if not pd.isna(parsed):
                return pd.Timestamp(parsed).normalize()

        return pd.NaT

    date_text = (
        cdhdr["Approval Date"]
        .map(normalize_text)
    )

    parsed_date = (
        cdhdr["Approval Date"]
        .map(parse_date)
    )

    invalid_date = (
        date_text.ne("")
        & parsed_date.isna()
    )

    if invalid_date.any():
        raise ValueError(
            f"PO CDHDR contains {int(invalid_date.sum())} "
            "invalid nonblank dates."
        )

    def normalize_time(value):
        text = re.sub(
            r"\D",
            "",
            normalize_text(value),
        )

        if text == "":
            return "000000"

        if len(text) > 6:
            raise ValueError(
                f"Invalid CDHDR approval time: {value!r}."
            )

        return text.zfill(6)

    time_text = (
        cdhdr["Approval Time"]
        .map(normalize_time)
    )

    invalid_clock = (
        ~time_text.str.fullmatch(
            r"(?:[01]\d|2[0-3])[0-5]\d[0-5]\d"
        )
    )

    if invalid_clock.any():
        value = cdhdr.loc[
            invalid_clock,
            "Approval Time",
        ].iloc[0]

        raise ValueError(
            f"Invalid CDHDR approval time: {value!r}."
        )

    timestamp = pd.to_datetime(
        (
            parsed_date.dt.strftime("%Y-%m-%d")
            + " "
            + time_text
        ),
        format="%Y-%m-%d %H%M%S",
        errors="coerce",
    )

    missing = timestamp.isna()

    if missing.any():
        raise ValueError(
            f"PO CDHDR contains {int(missing.sum())} change events "
            "without a valid date/timestamp."
        )

    return timestamp


def _release_marks(value):
    """Return positions carrying an X release mark while preserving spacing."""
    text = normalize_text(value).upper()

    return frozenset(
        index
        for index, character in enumerate(text)
        if character == "X"
    )


def build_release_events(cdhdr, cdpos):
    """Build joined FRGZU history and select effective, non-reversal events."""
    _validate_unique(
        cdhdr,
        CHANGE_KEY,
        "PO CDHDR",
    )

    _validate_unique(
        cdpos,
        CDPOS_KEY,
        "PO CDPOS",
    )

    cdhdr = cdhdr.copy()
    cdpos = cdpos.copy()

    cdhdr["Change Object"] = (
        cdhdr["Change Object"]
        .str.upper()
    )

    cdpos["Change Object"] = (
        cdpos["Change Object"]
        .str.upper()
    )

    cdpos["Table Name"] = (
        cdpos["Table Name"]
        .str.upper()
    )

    cdpos["Field Name"] = (
        cdpos["Field Name"]
        .str.upper()
    )

    cdpos["Change ID"] = (
        cdpos["Change ID"]
        .str.upper()
    )

    cdhdr_scope = cdhdr.loc[
        cdhdr["Change Object"].eq(CHANGE_OBJECT)
    ].copy()

    cdpos_scope = cdpos.loc[
        cdpos["Change Object"].eq(CHANGE_OBJECT)
        & cdpos["Table Name"].eq(HEADER_TABLE)
        & cdpos["Field Name"].eq(RELEASE_STATUS_FIELD)
    ].copy()

    if cdhdr_scope.empty:
        raise ValueError(
            f"PO CDHDR contains no {CHANGE_OBJECT} events."
        )

    if cdpos_scope.empty:
        raise ValueError(
            f"PO CDPOS contains no {CHANGE_OBJECT}/{HEADER_TABLE}/"
            f"{RELEASE_STATUS_FIELD} events; "
            "PO approver cannot be determined."
        )

    cdhdr_scope["Approval Timestamp"] = (
        _parse_approval_timestamp(
            cdhdr_scope
        )
    )

    events = cdpos_scope.merge(
        cdhdr_scope[
            CHANGE_KEY
            + [
                "Approver ID",
                "Approval Date",
                "Approval Time",
                "Approval Timestamp",
                "Transaction Code",
            ]
        ],
        on=CHANGE_KEY,
        how="inner",
        validate="many_to_one",
    )

    if events.empty:
        raise ValueError(
            "PO CDHDR and PO CDPOS have no matching "
            "FRGZU change-document keys."
        )

    events["PO Number"] = (
        events["Object Value"]
        .map(normalize_identifier)
    )

    events["PO Approver ID"] = (
        events["Approver ID"]
        .map(normalize_text)
    )

    events["Old Release Marks"] = (
        events["Old Value"]
        .map(_release_marks)
    )

    events["New Release Marks"] = (
        events["New Value"]
        .map(_release_marks)
    )

    events["Old Release Mark Count"] = (
        events["Old Release Marks"]
        .map(len)
    )

    events["New Release Mark Count"] = (
        events["New Release Marks"]
        .map(len)
    )

    events["Release Marks Added"] = events.apply(
        lambda row: len(
            row["New Release Marks"]
            - row["Old Release Marks"]
        ),
        axis=1,
    )

    events["Release Marks Removed"] = events.apply(
        lambda row: len(
            row["Old Release Marks"]
            - row["New Release Marks"]
        ),
        axis=1,
    )

    events["Is Reversal"] = (
        events["Release Marks Removed"]
        .gt(0)
    )

    events["Is Effective Release"] = (
        events["Change ID"].eq("U")
        & events["Release Marks Added"].gt(0)
        & events["Release Marks Removed"].eq(0)
        & events["PO Approver ID"].ne("")
        & events["PO Number"].ne("")
    )

    eligible = events.loc[
        events["Is Effective Release"]
    ].copy()

    if eligible.empty:
        raise ValueError(
            "PO CDHDR/CDPOS contain no eligible FRGZU release advances "
            "with a nonblank approver; PO approver cannot be determined."
        )

    duplicate_key = CHANGE_KEY + [
        "Old Value",
        "New Value",
        "PO Approver ID",
    ]

    duplicate_events = int(
        eligible.duplicated(
            duplicate_key,
            keep="first",
        ).sum()
    )

    eligible = eligible.drop_duplicates(
        duplicate_key,
        keep="first",
    )

    metrics = {
        "events_read": len(events),
        "events_eligible": len(eligible),
        "events_excluded": len(events) - len(eligible),
        "reversal_events": int(
            events["Is Reversal"].sum()
        ),
        "duplicate_events": duplicate_events,
    }

    return events, eligible, metrics


def _prepare_po_population(po_lines):
    """Resolve creator consistently at PO level and calculate coverage."""
    prepared = po_lines.copy()

    prepared["PO Creator ID"] = (
        prepared["PO Creator ID"]
        .map(normalize_text)
    )

    prepared["PO Number"] = (
        prepared["PO Number"]
        .map(normalize_identifier)
    )

    known = prepared.loc[
        prepared["PO Creator ID"].ne("")
    ]

    creator_counts = known.groupby(
        PO_KEY,
        dropna=False,
    )["PO Creator ID"].nunique()

    conflicts = creator_counts.loc[
        creator_counts.gt(1)
    ]

    if not conflicts.empty:
        raise ValueError(
            f"{CONTROL_ID} found {len(conflicts)} POs "
            "with more than one nonblank creator; "
            "creator cannot be resolved safely."
        )

    prepared["PO Creator ID"] = (
        prepared.assign(
            _CREATOR=prepared["PO Creator ID"].replace(
                "",
                pd.NA,
            )
        )
        .groupby(
            PO_KEY,
            dropna=False,
        )["_CREATOR"]
        .transform("first")
        .fillna("")
        .map(normalize_text)
    )

    po_headers = (
        prepared[
            PO_KEY
            + [
                "PO Creator ID",
            ]
        ]
        .drop_duplicates()
    )

    rows_without = int(
        prepared["PO Creator ID"]
        .eq("")
        .sum()
    )

    pos_without = int(
        po_headers["PO Creator ID"]
        .eq("")
        .sum()
    )

    metrics = {
        "rows_without_creator": rows_without,
        "pos_without_creator": pos_without,
        "distinct_pos": len(po_headers),
    }

    return prepared, po_headers, metrics


def _select_lha_approval_event(eligible, po_headers):
    """Select the highest effective release step for each in-population PO."""
    po_number_company = (
        po_headers[
            [
                "Company",
                "PO Number",
            ]
        ]
        .drop_duplicates()
    )

    ambiguous = (
        po_number_company
        .groupby("PO Number")["Company"]
        .nunique()
    )

    ambiguous = ambiguous.loc[
        ambiguous.gt(1)
    ]

    if not ambiguous.empty:
        raise ValueError(
            f"{CONTROL_ID} cannot safely assign CDHDR/CDPOS evidence: "
            f"{len(ambiguous)} PO numbers occur in more than one Company."
        )

    in_population = eligible.merge(
        po_number_company,
        on="PO Number",
        how="inner",
        validate="many_to_one",
    )

    outside = (
        len(eligible)
        - len(in_population)
    )

    if in_population.empty:
        raise ValueError(
            "No eligible approval events correspond to the "
            "CONFIG-filtered PO population."
        )

    in_population["_CHANGE_NUMBER_SORT"] = (
        pd.to_numeric(
            in_population["Change Number"],
            errors="coerce",
        )
        .fillna(-1)
    )

    ranked = in_population.sort_values(
        PO_KEY
        + [
            "New Release Mark Count",
            "Approval Timestamp",
            "_CHANGE_NUMBER_SORT",
        ],
        ascending=[
            True,
            True,
            False,
            False,
            False,
        ],
        kind="stable",
    )

    selected = (
        ranked
        .drop_duplicates(
            PO_KEY,
            keep="first",
        )
        .drop(
            columns=[
                "_CHANGE_NUMBER_SORT",
            ]
        )
    )

    return selected, outside


def _build_exceptions(po_lines, po_headers, selected):
    """Compare normalized technical IDs, then attach all lines of flagged POs."""
    selected = selected.copy()

    selected["PO Approval Date"] = (
        selected["Approval Timestamp"]
    )

    selected["PO Approval Status"] = (
        "EFFECTIVE RELEASE"
    )

    selected["Approval Event"] = (
        "FRGZU ADVANCE"
    )

    selected["Approval Step"] = (
        selected["New Release Mark Count"]
    )

    evidence_columns = PO_KEY + [
        "PO Approver ID",
        "PO Approval Date",
        "PO Approval Status",
        "Approval Event",
        "Approval Step",
        "Approval Date",
        "Approval Time",
        "Approval Timestamp",
        "Transaction Code",
        "Change Number",
        "Field Name",
        "Change ID",
        "Old Value",
        "New Value",
        "Old Release Mark Count",
        "New Release Mark Count",
        "Release Marks Added",
        "Release Marks Removed",
    ]

    comparison = po_headers.merge(
        selected[evidence_columns],
        on=PO_KEY,
        how="left",
        validate="one_to_one",
    )

    comparison["PO Approver ID"] = (
        comparison["PO Approver ID"]
        .map(normalize_text)
    )

    creator_compare = (
        comparison["PO Creator ID"]
        .str.upper()
    )

    approver_compare = (
        comparison["PO Approver ID"]
        .str.upper()
    )

    matching = comparison.loc[
        creator_compare.ne("")
        & approver_compare.ne("")
        & creator_compare.eq(approver_compare)
    ].copy()

    if matching.empty:
        exceptions = (
            po_lines.iloc[
                0:0
            ]
            .copy()
        )

        for column in (
            evidence_columns[2:]
            + [
                "PO03_KEY",
                "Comparison Evidence",
            ]
        ):
            exceptions[column] = pd.Series(
                dtype="object"
            )

        return comparison, exceptions

    matching["PO03_KEY"] = (
        matching[PO_KEY]
        .astype(str)
        .agg(
            "|".join,
            axis=1,
        )
    )

    matching["Comparison Evidence"] = (
        "PO Creator ID="
        + matching["PO Creator ID"]
        + "; PO Approver ID="
        + matching["PO Approver ID"]
    )

    detail_evidence = [
        column
        for column in matching.columns
        if column not in PO_KEY + ["PO Creator ID"]
    ]

    exceptions = po_lines.merge(
        matching[
            PO_KEY
            + detail_evidence
        ],
        on=PO_KEY,
        how="inner",
        validate="many_to_one",
    )

    exceptions = (
        exceptions
        .sort_values(
            PO_KEY
            + [
                "PO Line",
            ],
            kind="stable",
        )
        .reset_index(
            drop=True,
        )
    )

    return comparison, exceptions


def _build_summary(
    input_metrics,
    cdhdr_metrics,
    cdpos_metrics,
    population_metrics,
    event_metrics,
    comparison,
    exceptions,
    outside_events,
):
    """Build the PO03 execution summary."""
    with_evidence = int(
        comparison["PO Approver ID"]
        .map(normalize_text)
        .ne("")
        .sum()
    )

    without_evidence = (
        len(comparison)
        - with_evidence
    )

    exception_pos = (
        exceptions[PO_KEY]
        .drop_duplicates()
        .shape[0]
        if not exceptions.empty
        else 0
    )

    matching_users = (
        exceptions["PO Creator ID"]
        .map(normalize_text)
        .str.upper()
        .nunique()
        if not exceptions.empty
        else 0
    )

    result = (
        "ERROR"
        if exception_pos
        else "OK"
    )

    if exception_pos:
        message = (
            f"Found {exception_pos} POs "
            f"({len(exceptions)} detail lines) where "
            "the creator equals the highest effective "
            "release approver."
        )
    else:
        message = (
            "No creator matched the highest effective release approver. "
            f"Analyzed {len(comparison)} POs; "
            f"{with_evidence} had eligible approval evidence and "
            f"{without_evidence} did not."
        )

    return {
        "Record Type": "SUMMARY",
        "Control": CONTROL_ID,
        "Control Name": CONTROL_NAME,
        "Control Result": result,
        "Message": message,
        "PO Lines Rows Read": input_metrics["rows_read"],
        "PO Lines Residual Rows": input_metrics["residual_rows"],
        "Rows After Config Filters": input_metrics[
            "rows_after_config_filters"
        ],
        "Distinct POs Analyzed": population_metrics[
            "distinct_pos"
        ],
        "Rows Without Creator": population_metrics[
            "rows_without_creator"
        ],
        "POs Without Creator": population_metrics[
            "pos_without_creator"
        ],
        "CDHDR Rows Read": cdhdr_metrics["rows_read"],
        "CDHDR Residual Rows": cdhdr_metrics[
            "residual_rows"
        ],
        "CDPOS Rows Read": cdpos_metrics["rows_read"],
        "CDPOS Residual Rows": cdpos_metrics[
            "residual_rows"
        ],
        "Approval Events Read": event_metrics[
            "events_read"
        ],
        "Approval Events Eligible": event_metrics[
            "events_eligible"
        ],
        "Approval Events Excluded": event_metrics[
            "events_excluded"
        ],
        "Release Reversal Events": event_metrics[
            "reversal_events"
        ],
        "Duplicate Approval Events": event_metrics[
            "duplicate_events"
        ],
        "Approval Events Outside PO Population": outside_events,
        "POs Without Approval Evidence": without_evidence,
        "POs With Approval Evidence": with_evidence,
        "Exception Rows": len(exceptions),
        "Exception POs": exception_pos,
        "Distinct Matching Users": matching_users,
    }


def _build_output(summary_row, exceptions):
    """Return PO03 exceptions using exactly the established LHA columns."""
    del summary_row

    output = pd.DataFrame(
        index=exceptions.index
    )

    def source(column):
        if column in exceptions.columns:
            return exceptions[column]

        return pd.Series(
            "",
            index=exceptions.index,
            dtype="object",
        )

    company = source("Company")

    po_date = pd.to_datetime(
        source("PO Doc Date"),
        errors="coerce",
    )

    pr_number = (
        source("PR Number")
        .map(normalize_identifier)
    )

    deleted = (
        source("PO Line Deleted")
        .map(normalize_text)
    )

    completed = (
        source("PO Delivery Completed")
        .map(normalize_text)
    )

    output["CoCo"] = company
    output["Company"] = company
    output["PO Number"] = source("PO Number")
    output["PO DocEntry"] = ""
    output["PO Line"] = source("PO Line")
    output["Vendor Code"] = source("Vendor Code")
    output["Vendor Name"] = ""
    output["PO Doc Date"] = source("PO Doc Date")
    output["PO Doc Currency"] = source(
        "PO Doc Currency"
    )
    output["Company Main Currency"] = ""
    output["PO Canceled"] = ""

    output["PO Line Status"] = pd.Series(
        "",
        index=exceptions.index,
        dtype="object",
    )

    output.loc[
        deleted.ne(""),
        "PO Line Status",
    ] = "DELETED"

    output.loc[
        deleted.eq("")
        & completed.ne(""),
        "PO Line Status",
    ] = "COMPLETED"

    output["Item Code"] = source("Item Code")
    output["Account Code"] = ""
    output["PO Material Description"] = source(
        "PO Material Description"
    )
    output["PO Quantity"] = source("PO Quantity")
    output["PO Unit Price"] = source(
        "PO Unit Price"
    )
    output["PO Line Total"] = source("PO Line Total")
    output["PO Line Total USD"] = ""
    output["USD Rate"] = ""
    output["USD Rate Date"] = ""
    output["PO Creator ID"] = source("PO Creator ID")
    output["PO Creator Name"] = ""
    output["PO Approval Date"] = source(
        "PO Approval Date"
    )
    output["PO Approver ID"] = source(
        "PO Approver ID"
    )
    output["PO Approver Name"] = ""
    output["PO Approval Status"] = source(
        "PO Approval Status"
    )
    output["GR Doc Number"] = ""
    output["GR Doc Date"] = ""
    output["GR First Posting Date"] = ""
    output["GR Last Posting Date"] = ""
    output["GR Quantity"] = ""
    output["GR Creator ID"] = ""
    output["GR Creator Name"] = ""

    output["PO Month"] = (
        po_date.dt.strftime("%Y-%m")
        .fillna("")
    )

    output["PR DocEntry"] = pr_number
    output["PR Line"] = source("PR Line")

    output["From PR"] = (
        pr_number.ne("")
        .map(
            {
                True: "Y",
                False: "N",
            }
        )
    )

    return (
        output
        .reindex(
            columns=LHA_REPORT_COLUMNS
        )
        .reset_index(
            drop=True
        )
    )


def run_po_003(context):
    """Execute PO03 independently and replace only the PO03 output sheet."""
    po_lines, input_metrics = load_po_lines(
        context,
        required_fields=PO_REQUIRED_FIELDS,
    )

    cdhdr, cdhdr_metrics = _load_change_input(
        context,
        CDHDR_INPUT_PREFIX,
        CDHDR_ALIASES,
    )

    cdpos, cdpos_metrics = _load_change_input(
        context,
        CDPOS_INPUT_PREFIX,
        CDPOS_ALIASES,
    )

    (
        prepared_lines,
        po_headers,
        population_metrics,
    ) = _prepare_po_population(
        po_lines
    )

    (
        _,
        eligible,
        event_metrics,
    ) = build_release_events(
        cdhdr,
        cdpos,
    )

    (
        selected,
        outside_events,
    ) = _select_lha_approval_event(
        eligible,
        po_headers,
    )

    (
        comparison,
        exceptions,
    ) = _build_exceptions(
        prepared_lines,
        po_headers,
        selected,
    )

    summary_row = _build_summary(
        input_metrics,
        cdhdr_metrics,
        cdpos_metrics,
        population_metrics,
        event_metrics,
        comparison,
        exceptions,
        outside_events,
    )

    output = _build_output(
        summary_row,
        exceptions,
    )

    output_file = write_control_sheet(
        context=context,
        sheet_name=SHEET_NAME,
        dataframe=output,
        date_columns=[
            "PO Doc Date",
            "PO Approval Date",
            "USD Rate Date",
            "GR Doc Date",
            "GR First Posting Date",
            "GR Last Posting Date",
        ],
        amount_columns=[
            "PO Quantity",
            "PO Unit Price",
            "PO Line Total",
            "PO Line Total USD",
            "USD Rate",
            "GR Quantity",
        ],
    )

    print()
    print(
        f"{CONTROL_ID} - "
        f"{CONTROL_NAME}"
    )
    print(
        "-"
        * (
            len(CONTROL_ID)
            + len(CONTROL_NAME)
            + 3
        )
    )

    for label in SUMMARY_COLUMNS[5:]:
        print(
            f"{label}: "
            f"{summary_row[label]}"
        )

    print(
        "Control result: "
        f"{summary_row['Control Result']}"
    )

    print(
        "Message: "
        f"{summary_row['Message']}"
    )

    print(
        "PO03 output file: "
        f"{output_file}"
    )

    print(
        "PO03 output sheet: "
        f"{SHEET_NAME}"
    )

    print()

    return {
        "status": summary_row["Control Result"],
        "output_file": output_file,
        "sheet_name": SHEET_NAME,
        "rows": len(exceptions),
        "summary": summary_row,
    }


run = run_po_003
