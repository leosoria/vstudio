"""
VM_006 - Business Number duplicate control.

Functional disposition
----------------------
VM06 follows the approved behavior: it is functionally equivalent to VM05
and therefore does not independently recompute or duplicate the VM05
exceptions.

The control writes an informational VM06 worksheet that directs the reviewer
to the VM05 results.

This control deliberately does not:

- load the vendor master;
- resolve a Tax/Business Number;
- define a priority among SAP ECC tax-number fields;
- load vendor postings;
- calculate last-invoice information;
- perform foreign-exchange conversion; or
- reproduce VM05 analytical logic.
"""

from time import perf_counter
from typing import Any

import pandas as pd

from core.vm_common import write_vm_control_sheet


CONTROL_ID = "VM_006"
SHEET_NAME = "VM06"

OUTPUT_COLUMNS = [
    "Resultado",
]

RESULT_MESSAGE = (
    "Control duplicado de VM05: Business Number y Tax Number se evalúan "
    "mediante el mismo control funcional aprobado. Ver resultados de VM05."
)


def _print_timing(
    stage_name: str,
    started: float,
) -> float:
    """Print one lightweight stage timing and return its completion time."""
    finished = perf_counter()

    print(
        f"{CONTROL_ID} {stage_name}: "
        f"{finished - started:.2f} seconds"
    )

    return finished


def build_vm_006() -> pd.DataFrame:
    """
    Build the approved informational VM06 result.

    Returns
    -------
    pandas.DataFrame
        Exactly one informational row with the single approved column
        ``Resultado``.
    """
    output = pd.DataFrame(
        [
            {
                "Resultado": RESULT_MESSAGE,
            }
        ],
        columns=OUTPUT_COLUMNS,
    )

    if list(output.columns) != OUTPUT_COLUMNS:
        raise AssertionError(
            f"{CONTROL_ID}: unexpected output columns. "
            f"Expected {OUTPUT_COLUMNS}, received {list(output.columns)}."
        )

    if len(output) != 1:
        raise AssertionError(
            f"{CONTROL_ID}: the informational output must contain "
            "exactly one row."
        )

    if output["Resultado"].isna().any():
        raise AssertionError(
            f"{CONTROL_ID}: Resultado cannot be null."
        )

    if output["Resultado"].astype("string").str.strip().eq("").any():
        raise AssertionError(
            f"{CONTROL_ID}: Resultado cannot be blank."
        )

    return output


def run_vm_006(
    context: dict[str, Any],
) -> dict[str, Any]:
    """
    Execute VM06 and replace only the VM06 result worksheet.

    VM06 is informational. It does not treat its single note row as an
    analytical exception.
    """
    total_started = perf_counter()
    stage_started = total_started

    print(f"{CONTROL_ID} source vendor rows: 0 (not applicable)")
    print(f"{CONTROL_ID} master vendor rows: 0 (not applicable)")
    print(f"{CONTROL_ID} rows excluded by CONFIG: 0 (not applicable)")
    print(f"{CONTROL_ID} valid vendor rows: 0 (not applicable)")
    print(f"{CONTROL_ID} posting rows: 0 (not applicable)")
    print(f"{CONTROL_ID} vendors with last invoice: 0 (not applicable)")

    stage_started = _print_timing(
        "vendor load (not applicable)",
        stage_started,
    )

    stage_started = _print_timing(
        "population preparation (not applicable)",
        stage_started,
    )

    stage_started = _print_timing(
        "posting load (not applicable)",
        stage_started,
    )

    stage_started = _print_timing(
        "last-invoice construction (not applicable)",
        stage_started,
    )

    stage_started = _print_timing(
        "analytical normalization (not applicable)",
        stage_started,
    )

    output = build_vm_006()

    stage_started = _print_timing(
        "informational control logic",
        stage_started,
    )

    if list(output.columns) != OUTPUT_COLUMNS:
        raise AssertionError(
            f"{CONTROL_ID}: output column validation failed."
        )

    duplicate_output_key_rows = int(
        output.duplicated(
            subset=OUTPUT_COLUMNS,
            keep=False,
        ).sum()
    )

    if duplicate_output_key_rows != 0:
        raise AssertionError(
            f"{CONTROL_ID}: duplicate informational rows detected."
        )

    stage_started = _print_timing(
        "validations",
        stage_started,
    )

    stage_started = _print_timing(
        "FX conversion (not applicable)",
        stage_started,
    )

    write_started = stage_started

    output_file = write_vm_control_sheet(
        context=context,
        sheet_name=SHEET_NAME,
        dataframe=output,
        date_columns=[],
        amount_columns=[],
        integer_columns=[],
    )

    finished = _print_timing(
        "workbook write",
        write_started,
    )

    print(
        f"{CONTROL_ID} eligible Business Number rows: "
        "0 (not applicable)"
    )
    print(f"{CONTROL_ID} exception rows: 0")
    print(f"{CONTROL_ID} groups: 0")
    print(
        f"{CONTROL_ID} duplicate output-key rows: "
        f"{duplicate_output_key_rows}"
    )
    print(
        f"WARNING: {CONTROL_ID} is informational and does not "
        "independently recompute VM05."
    )
    print(
        f"{CONTROL_ID} total: "
        f"{finished - total_started:.2f} seconds"
    )

    return {
        "status": "OK",
        "output_file": output_file,
        "sheet_name": SHEET_NAME,
        "rows": len(output),
    }
