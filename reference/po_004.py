"""PO04 - goods receipts posted before their purchase order date.

The control independently reads the original PO Lines and PO GR SAP ECC
exports, associates goods-receipt movements with their purchase-order line,
and reports lines whose first GR posting date predates the PO document date.
Its output follows the LHA PO04 detail contract and replaces only sheet PO04.
"""

from pathlib import Path

import pandas as pd

from core.gl_common import (
    load_gl_fx_rates_data,
    normalize_fx_rates,
    select_fx_rate_to_usd,
)
from core.po_common import (
    ALLOWED_INPUT_EXTENSIONS,
    get_period_suffix,
    inspect_input_workbook,
    load_po_lines,
    normalize_company,
    normalize_identifier,
    normalize_lookup,
    normalize_text,
    resolve_sheet_name,
    write_control_sheet,
)


CONTROL_ID = "PO_004"
CONTROL_NAME = "Goods Receipt Before Purchase Order"
SHEET_NAME = "PO04"
INPUT_SHEET = "Sheet1"
HEADER_ROW = 1

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

GR_ALIASES = {
    "PO Number": (
        "Purch.Doc.",
        "Purchasing Document",
        "EBELN",
        "EKBE-EBELN",
    ),
    "PO Line": (
        "Item",
        "PO Line",
        "EBELP",
        "EKBE-EBELP",
    ),
    "GR Doc Number": (
        "Mat. Doc.",
        "Material Document",
        "BELNR",
        "EKBE-BELNR",
    ),
    "Fiscal Year": (
        "MatYr",
        "Fiscal Year",
        "GJAHR",
        "EKBE-GJAHR",
    ),
    "GR Document Line": (
        "Item.1",
        "Document Line",
        "BUZEI",
        "EKBE-BUZEI",
    ),
    "GR Posting Date": (
        "Pstng Date",
        "Posting Date",
        "BUDAT",
        "EKBE-BUDAT",
    ),
    "Transaction Event Type": (
        "Tr./ev.type",
        "VGABE",
        "EKBE-VGABE",
    ),
    "GR Quantity": (
        "Quantity",
        "MENGE",
        "EKBE-MENGE",
    ),
    "Debit Credit Indicator": (
        "D/C",
        "SHKZG",
        "EKBE-SHKZG",
    ),
    "GR Doc Date": (
        "Doc. Date",
        "Document Date",
        "BLDAT",
        "MKPF-BLDAT",
    ),
    "GR Creator ID": (
        "User name",
        "USNAM",
        "MKPF-USNAM",
    ),
    "Company": (
        "CoCd",
        "Company",
        "BUKRS",
        "EKKO-BUKRS",
    ),
    "PO Date In GR": (
        "Doc. Date.1",
        "PO Document Date",
        "BEDAT",
        "EKKO-BEDAT",
    ),
}

GR_KEY = [
    "Company",
    "PO Number",
    "PO Line",
    "GR Doc Number",
    "Fiscal Year",
    "GR Document Line",
]

PO_LINE_KEY = [
    "Company",
    "PO Number",
    "PO Line",
]

LHA_OUTPUT_COLUMNS = [
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
    "Days GR Before PO",
]


def _resolve_columns(dataframe, aliases, source_name):
    """Resolve the documented SAP headers without positional assumptions."""
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

        matches = list(
            dict.fromkeys(matches)
        )

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


def _find_gr_input(context):
    """Return the period-specific PO goods-receipt input."""
    return find_period_input_file(
        context=context,
        input_prefix=PR_GR_INPUT_PREFIX,
        source_name="PO GR",
    )


def load_gr_movements(context):
    """Read, validate and normalize the independent ECC PO GR population."""
    path = _find_gr_input(context)

    sheet = resolve_sheet_name(
        path,
        INPUT_SHEET,
    )

    workbook_metrics = inspect_input_workbook(
        path,
        sheet,
    )

    raw = (
        pd.read_excel(
            path,
            sheet_name=sheet,
            header=HEADER_ROW - 1,
            dtype=object,
        )
        .dropna(
            axis=0,
            how="all",
        )
        .dropna(
            axis=1,
            how="all",
        )
    )

    if raw.empty:
        raise ValueError(
            f"PO GR sheet '{sheet}' contains no data rows."
        )

    columns = _resolve_columns(
        raw,
        GR_ALIASES,
        "PO GR",
    )

    physical_key = [
        columns[column]
        for column in GR_KEY
    ]

    blank_key = pd.DataFrame(
        {
            column: (
                raw[column]
                .map(normalize_text)
                .eq("")
            )
            for column in physical_key
        },
        index=raw.index,
    )

    residual = blank_key.all(
        axis=1
    )

    partial = (
        blank_key.any(axis=1)
        & ~residual
    )

    if partial.any():
        raise ValueError(
            f"PO GR contains {int(partial.sum())} rows "
            "with a partially blank physical movement key."
        )

    population = (
        raw.loc[~residual]
        .rename(
            columns={
                physical: logical
                for logical, physical in columns.items()
            }
        )
        .copy()
    )

    for column in GR_KEY:
        population[column] = (
            population[column]
            .map(normalize_identifier)
        )

    population["Company"] = (
        population["Company"]
        .map(normalize_company)
    )

    duplicates = population.duplicated(
        GR_KEY,
        keep=False,
    )

    if duplicates.any():
        raise ValueError(
            f"PO GR contains {int(duplicates.sum())} rows "
            "participating in duplicate physical keys "
            f"{GR_KEY}."
        )

    event_type = (
        population["Transaction Event Type"]
        .map(normalize_identifier)
    )

    gr_population = population.loc[
        event_type.eq("1")
    ].copy()

    if gr_population.empty:
        raise ValueError(
            "PO GR contains no goods-receipt events "
            "with Tr./ev.type = 1."
        )

    for column in (
        "GR Posting Date",
        "GR Doc Date",
        "PO Date In GR",
    ):
        raw_dates = gr_population[column]

        parsed = pd.to_datetime(
            raw_dates,
            errors="coerce",
            dayfirst=True,
        ).dt.normalize()

        invalid = (
            raw_dates.map(normalize_text).ne("")
            & parsed.isna()
        )

        if invalid.any():
            raise ValueError(
                f"PO GR contains {int(invalid.sum())} "
                "invalid nonblank values "
                f"in {column}."
            )

        gr_population[column] = parsed

    gr_population["GR Quantity"] = pd.to_numeric(
        gr_population["GR Quantity"],
        errors="coerce",
    )

    credit = (
        gr_population["Debit Credit Indicator"]
        .map(normalize_text)
        .str.upper()
        .eq("H")
    )

    gr_population.loc[
        credit,
        "GR Quantity",
    ] = (
        gr_population.loc[
            credit,
            "GR Quantity",
        ]
        .abs()
        .mul(-1)
    )

    metrics = {
        **workbook_metrics,
        "input_file": str(path),
        "input_sheet": sheet,
        "rows_read": len(raw),
        "residual_rows": int(
            residual.sum()
        ),
        "non_gr_rows": (
            len(population)
            - len(gr_population)
        ),
        "gr_rows": len(gr_population),
    }

    return gr_population, metrics


def aggregate_gr_movements(gr_movements):
    """Build the LHA one-row-per-PO-line GR fields from movement history."""
    ordered = gr_movements.sort_values(
        PO_LINE_KEY
        + [
            "GR Posting Date",
            "GR Doc Number",
            "GR Document Line",
        ],
        kind="stable",
    )

    latest = ordered.drop_duplicates(
        PO_LINE_KEY,
        keep="last",
    )

    dates_and_quantity = ordered.groupby(
        PO_LINE_KEY,
        as_index=False,
        dropna=False,
    ).agg(
        **{
            "GR First Posting Date": (
                "GR Posting Date",
                "min",
            ),
            "GR Last Posting Date": (
                "GR Posting Date",
                "max",
            ),
            "GR Quantity": (
                "GR Quantity",
                lambda values: values.sum(
                    min_count=1
                ),
            ),
        }
    )

    latest_fields = latest[
        PO_LINE_KEY
        + [
            "GR Doc Number",
            "GR Doc Date",
            "GR Creator ID",
        ]
    ]

    return dates_and_quantity.merge(
        latest_fields,
        on=PO_LINE_KEY,
        how="left",
        validate="one_to_one",
    )


def prepare_unique_po_lines(po_lines):
    """Return one stable PO detail row per company, PO and line."""
    unique = po_lines.drop_duplicates(
        PO_LINE_KEY,
        keep="first",
    ).copy()

    duplicate_rows_removed = (
        len(po_lines)
        - len(unique)
    )

    return (
        unique,
        duplicate_rows_removed,
    )


def build_po04_exceptions(
    po_lines,
    gr_summary,
):
    """Join PO and GR independently and apply the exact LHA date test."""
    unique_po_lines, _ = prepare_unique_po_lines(
        po_lines
    )

    joined = unique_po_lines.merge(
        gr_summary,
        on=PO_LINE_KEY,
        how="inner",
        validate="one_to_one",
    )

    po_date = pd.to_datetime(
        joined["PO Doc Date"],
        errors="coerce",
    ).dt.normalize()

    gr_date = pd.to_datetime(
        joined["GR First Posting Date"],
        errors="coerce",
    ).dt.normalize()

    exception_mask = (
        gr_date.notna()
        & po_date.notna()
        & gr_date.lt(po_date)
    )

    exceptions = joined.loc[
        exception_mask
    ].copy()

    exceptions["Days GR Before PO"] = (
        po_date.loc[exceptions.index]
        - gr_date.loc[exceptions.index]
    ).dt.days

    return (
        exceptions
        .sort_values(
            [
                "Company",
                "PO Number",
                "PO Line",
            ],
            kind="stable",
        )
        .reset_index(drop=True)
    )


def add_document_date_usd_fields(
    exceptions,
    fx_dataframe,
):
    """Convert exception line totals to USD at the PO document date."""
    result = exceptions.copy()

    result["PO Line Total USD"] = pd.NA
    result["USD Rate"] = pd.NA
    result["USD Rate Date"] = pd.NaT

    if result.empty:
        return result

    if fx_dataframe.empty:
        normalized_fx = pd.DataFrame()
    else:
        normalized_fx = normalize_fx_rates(
            fx_dataframe
        )

    cache = {}

    for index, row in result.iterrows():
        currency = normalize_text(
            row.get(
                "PO Doc Currency",
                "",
            )
        ).upper()

        requested_date = pd.to_datetime(
            row.get(
                "PO Doc Date",
            ),
            errors="coerce",
        )

        if pd.isna(requested_date):
            normalized_date = pd.NaT
        else:
            normalized_date = (
                requested_date.normalize()
            )

        key = (
            currency,
            normalized_date,
        )

        if key not in cache:
            cache[key] = select_fx_rate_to_usd(
                normalized_fx,
                currency,
                requested_date,
            )

        details = cache[key]

        if details is None:
            continue

        amount = pd.to_numeric(
            pd.Series(
                [
                    row.get(
                        "PO Line Total",
                    )
                ]
            ),
            errors="coerce",
        ).iloc[0]

        if not pd.isna(amount):
            result.at[
                index,
                "PO Line Total USD",
            ] = (
                amount
                * details["fx_to_usd"]
            )

        result.at[
            index,
            "USD Rate",
        ] = details["usd_rate"]

        result.at[
            index,
            "USD Rate Date",
        ] = details["rate_date"]

    return result


def build_lha_output(exceptions):
    """Return PO04 detail with the exact LHA column names and order."""
    output = exceptions.copy()

    output["CoCo"] = output.get(
        "Company",
        "",
    )

    po_dates = pd.to_datetime(
        output.get(
            "PO Doc Date",
        ),
        errors="coerce",
    )

    output["PO Month"] = (
        po_dates
        .dt.strftime("%Y-%m")
        .fillna("")
    )

    pr_number = output.get(
        "PR Number",
        pd.Series(
            "",
            index=output.index,
            dtype="object",
        ),
    ).map(normalize_text)

    output["From PR"] = (
        pr_number
        .ne("")
        .map(
            {
                True: "Y",
                False: "N",
            }
        )
    )

    for column in LHA_OUTPUT_COLUMNS:
        if column not in output.columns:
            output[column] = ""

    output = output.loc[
        :,
        LHA_OUTPUT_COLUMNS,
    ]

    if list(output.columns) != LHA_OUTPUT_COLUMNS:
        raise AssertionError(
            f"{CONTROL_ID} output columns do not match "
            "the LHA PO04 contract."
        )

    return output


def run_po_004(context):
    """Execute PO04 independently and replace only the PO04 output sheet."""
    po_lines, po_metrics = load_po_lines(
        context,
        required_fields=PO_REQUIRED_FIELDS,
    )

    gr_movements, gr_metrics = load_gr_movements(
        context
    )

    gr_summary = aggregate_gr_movements(
        gr_movements
    )

    exceptions = build_po04_exceptions(
        po_lines,
        gr_summary,
    )

    exceptions = add_document_date_usd_fields(
        exceptions,
        load_gl_fx_rates_data(context),
    )

    output = build_lha_output(
        exceptions
    )

    output_file = write_control_sheet(
        context=context,
        sheet_name=SHEET_NAME,
        dataframe=output,
        date_columns=[
            "PO Doc Date",
            "USD Rate Date",
            "PO Approval Date",
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
            "Days GR Before PO",
        ],
    )

    status = (
        "ERROR"
        if not exceptions.empty
        else "OK"
    )

    summary = {
        "Control": CONTROL_ID,
        "Control Name": CONTROL_NAME,
        "Control Result": status,
        "PO Rows Read": (
            po_metrics["rows_read"]
        ),
        "PO Rows After Config Filters": (
            po_metrics[
                "rows_after_config_filters"
            ]
        ),
        "GR Rows Read": (
            gr_metrics["rows_read"]
        ),
        "GR Rows": (
            gr_metrics["gr_rows"]
        ),
        "Exception Rows": len(
            exceptions
        ),
        "Exception POs": (
            int(
                exceptions[
                    "PO Number"
                ].nunique()
            )
            if not exceptions.empty
            else 0
        ),
    }

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
    print(
        "PO Lines input: "
        f"{po_metrics['input_file']}"
    )
    print(
        "PO GR input: "
        f"{gr_metrics['input_file']}"
    )
    print(
        "PO rows after CONFIG filters: "
        f"{summary['PO Rows After Config Filters']}"
    )
    print(
        "GR movements analyzed: "
        f"{summary['GR Rows']}"
    )
    print(
        "Exception rows: "
        f"{summary['Exception Rows']}"
    )
    print(
        "Exception POs: "
        f"{summary['Exception POs']}"
    )
    print(
        "Control result: "
        f"{status}"
    )
    print(
        "PO04 output file: "
        f"{output_file}"
    )
    print(
        "PO04 output sheet: "
        f"{SHEET_NAME}"
    )
    print()

    return {
        "status": status,
        "output_file": output_file,
        "sheet_name": SHEET_NAME,
        "rows": len(exceptions),
        "summary": summary,
    }
