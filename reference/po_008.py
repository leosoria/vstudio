"""PO08 - POs por item por mes.

Genera un resumen por Company + PO Month + Item Code con:
- cantidad de lineas de PO;
- cantidad de POs distintas;
- cantidad total;
- importe total;
- importe total convertido a USD.

Optimizaciones:
- Evita el segundo recorrido completo del Excel para buscar formulas.
- Resuelve FX una sola vez por moneda y fecha.
- Utiliza agregaciones vectorizadas de pandas.
"""

from unittest.mock import patch

import pandas as pd

from core import po_common
from core.gl_common import (
    load_gl_fx_rates_data,
    normalize_fx_rates,
    select_fx_rate_to_usd,
)
from core.po_common import write_control_sheet


CONTROL_ID = "PO_008"
CONTROL_NAME = "POs By Item By Month"
SHEET_NAME = "PO08"

REQUIRED_FIELDS = (
    "Company",
    "PO Number",
    "PO Line",
    "PO Doc Date",
    "Item Code",
    "PO Quantity",
    "PO Doc Currency",
    "PO Line Total",
)

OUTPUT_COLUMNS = [
    "Company",
    "PO Month",
    "Item Code",
    "PO Lines",
    "Distinct POs",
    "Total Quantity",
    "Total Line Amount",
    "Total Line Amount USD",
]


def _load_po_lines_fast(context):
    """Carga PO Lines sin recorrer previamente todas las celdas con openpyxl."""
    skipped_inspection = {
        "physical_rows_including_header": 0,
        "physical_columns": 0,
        "formula_cells": 0,
    }

    with patch.object(
        po_common,
        "inspect_input_workbook",
        return_value=skipped_inspection,
    ):
        dataframe, metrics = po_common.load_po_lines(
            context,
            required_fields=REQUIRED_FIELDS,
        )

    metrics["physical_rows_including_header"] = (
        metrics["rows_read"] + po_common.PO_HEADER_ROW
    )
    metrics["physical_columns"] = len(dataframe.columns)

    return dataframe, metrics


def _add_usd_amount(dataframe, context):
    """Resuelve el FX una vez por cada combinacion moneda-fecha."""
    result = dataframe.copy()

    result["__currency"] = (
        result["PO Doc Currency"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    result["__fx_date"] = (
        pd.to_datetime(
            result["PO Doc Date"],
            errors="coerce",
        )
        .dt.normalize()
    )

    normalized_fx = normalize_fx_rates(
        load_gl_fx_rates_data(context)
    )

    fx_keys = (
        result[
            [
                "__currency",
                "__fx_date",
            ]
        ]
        .drop_duplicates()
        .reset_index(drop=True)
    )

    fx_factors = []

    for currency, document_date in fx_keys.itertuples(
        index=False,
        name=None,
    ):
        details = select_fx_rate_to_usd(
            normalized_fx_dataframe=normalized_fx,
            currency=currency,
            requested_date=document_date,
        )

        fx_factors.append(
            pd.NA
            if details is None
            else details["fx_to_usd"]
        )

    fx_keys["__fx_to_usd"] = fx_factors

    result = result.merge(
        fx_keys,
        on=[
            "__currency",
            "__fx_date",
        ],
        how="left",
        validate="many_to_one",
        sort=False,
    )

    result["PO Line Total USD"] = (
        result["PO Line Total"]
        * result["__fx_to_usd"]
    )

    return result.drop(
        columns=[
            "__currency",
            "__fx_date",
            "__fx_to_usd",
        ]
    )


def build_po08(dataframe, context):
    """Agrupa las lineas de PO por compañia, mes e item."""
    data = dataframe.copy()

    data["PO Doc Date"] = pd.to_datetime(
        data["PO Doc Date"],
        errors="coerce",
    )

    data["PO Month"] = (
        data["PO Doc Date"]
        .dt.strftime("%Y-%m")
    )

    numeric_columns = [
        "PO Quantity",
        "PO Line Total",
    ]

    data[numeric_columns] = data[numeric_columns].apply(
        pd.to_numeric,
        errors="coerce",
    )

    data = _add_usd_amount(
        data,
        context,
    )

    group_columns = [
        "Company",
        "PO Month",
        "Item Code",
    ]

    grouped = data.groupby(
        group_columns,
        dropna=False,
        sort=False,
    )

    counts = grouped["PO Number"].agg(
        **{
            "PO Lines": "size",
            "Distinct POs": "nunique",
        }
    )

    totals = grouped[
        [
            "PO Quantity",
            "PO Line Total",
            "PO Line Total USD",
        ]
    ].sum(
        min_count=1,
    )

    output = (
        counts
        .join(totals)
        .rename(
            columns={
                "PO Quantity": "Total Quantity",
                "PO Line Total": "Total Line Amount",
                "PO Line Total USD": "Total Line Amount USD",
            }
        )
        .reset_index()
        .sort_values(
            group_columns,
            kind="stable",
        )
        .reset_index(drop=True)
    )

    return output.loc[
        :,
        OUTPUT_COLUMNS,
    ]


def run_po_008(context):
    """Ejecuta PO08 y reemplaza solamente la hoja PO08."""
    po_lines, metrics = _load_po_lines_fast(
        context
    )

    output = build_po08(
        po_lines,
        context,
    )

    output_file = write_control_sheet(
        context=context,
        sheet_name=SHEET_NAME,
        dataframe=output,
        amount_columns=[
            "Total Quantity",
            "Total Line Amount",
            "Total Line Amount USD",
        ],
    )

    print(f"{CONTROL_ID} - {CONTROL_NAME}")
    print(f"Rows read: {metrics['rows_read']}")
    print(
        "Formula inspection: skipped "
        "(PO08 fast mode)"
    )
    print(f"Summary rows: {len(output)}")
    print(f"Output file: {output_file}")
    print(f"Output sheet: {SHEET_NAME}")

    return {
        "status": "OK",
        "output_file": output_file,
        "sheet_name": SHEET_NAME,
        "rows": len(output),
    }


run = run_po_008
