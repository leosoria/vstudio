"""
VM_018 - Vendors sharing the same email address.

Uses the shared configurable VM matching engine.
Only active eligible suppliers are evaluated.

Enabled conditions are read from the VM sheet in config.xlsx.
Email comparison ignores surrounding whitespace and letter case.
Multiple email rows per supplier are preserved.

Output:
    VM18, one row per qualifying supplier pair.
"""

CONTROL_ID = "VM_018"
SHEET_NAME = "VM18"


def run_vm_018(context):
    """Execute VM18 using the shared VM matching engine."""
    from core.vm_common import run_configured_vm

    return run_configured_vm(
        context,
        sheet_name=SHEET_NAME,
    )
