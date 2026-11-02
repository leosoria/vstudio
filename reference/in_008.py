"""IN08 - Preliminary billing activity; classification pending."""
from core.in_common import run_in_billing

CONTROL_ID = "IN_008"

def run_in_008(context):
    return run_in_billing(context, "V")
