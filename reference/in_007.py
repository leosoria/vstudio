"""IN07 - Preliminary billing activity; classification pending."""
from core.in_common import run_in_billing

CONTROL_ID = "IN_007"

def run_in_007(context):
    return run_in_billing(context, "Q")
