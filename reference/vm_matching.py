"""Shared configurable vendor matching for VM02-VM05 and VM18."""
import re
import pandas as pd

from core.vm_common import (
    load_vm_vendors,
    build_vendor_master_population,
    get_valid_vendor_population,
    normalize_company,
    safe_text,
    resolve_tax_business_number,
    load_vm_vendor_postings,
    build_vm_last_invoice_population,
    write_vm_control_sheet,
)
from core.vm_rules import (
    parse_vm_rules,
    filter_active_records,
    find_vendor_pairs,
)

TAX_PRIORITY = (
    "Tax Number 1",
    "Tax Number 2",
    "Tax Number 3",
    "Tax Number 4",
    "Tax Number 5",
    "VAT Registration Number",
)

IDENTITY = ["Company", "Vendor Code"]


def load_matching_population(context, rules):
    source = load_vm_vendors(context)
    available = set(source.attrs.get("vm_source_columns", ()))

    if not available:
        raise ValueError(
            "VM: update vm_common.py with source-column tracking."
        )

    required = {
        "Company",
        "Vendor Code",
        "Vendor Name",
        "Central Deletion Flag",
        "Company Deletion Flag",
        "Central Posting Block",
        "Purchasing Block",
        "Central Payment Block",
    }

    fields = {
        "Same Address": {"Street", "City", "ZipCode", "Country"},
        "Same Phone Number": {"Phone1"},
        "Same Bank Account Number": {"Bank Account"},
        "Same Email": {"Email"},
    }

    for header, names in fields.items():
        if rules[header]:
            required.update(names)

    missing = sorted(required.difference(available))
    if missing:
        raise ValueError(
            f"VM: fields absent from the source extract: {missing}."
        )

    if (
        rules["Same Tax Registration Number"]
        and not available.intersection(TAX_PRIORITY)
    ):
        raise ValueError(
            "VM: no recognized tax-number field in the source."
        )

    master = build_vendor_master_population(source)
    configured = context["module"].get("companies", "")

    if isinstance(configured, (list, tuple, set)):
        values = configured
    else:
        values = re.split(r"[,;|]", safe_text(configured))

    values = [
        safe_text(value)
        for value in values
        if safe_text(value)
    ]

    if values and not any(
        value.upper() in {"ALL", "*", "TODAS", "TODOS"}
        for value in values
    ):
        companies = {
            normalize_company(value)
            for value in values
        }
        master = master.loc[
            master["Company"].isin(companies)
        ].copy()

    vendors, metrics = get_valid_vendor_population(master)

    # F-code records belong to the employee side, not supplier matching.
    is_employee = vendors["Vendor Code"].str.startswith(
        "F", na=False
    )
    vendors = vendors.loc[~is_employee].copy()

    vendors, active_metrics = filter_active_records(vendors)

    print(f"VM active supplier identities: {len(vendors)}")
    print(
        "VM supplier identities excluded by blocking: "
        f"{active_metrics['excluded_rows']}"
    )

    for warning in metrics.get("warnings", []):
        print(f"WARNING: {warning}")

    # Preserve all bank/email rows for eligible company/vendor identities.
    population = source.merge(
        vendors[IDENTITY],
        on=IDENTITY,
        how="inner",
        validate="many_to_one",
    )

    if rules["Same Tax Registration Number"]:
        population = resolve_tax_business_number(
            population,
            TAX_PRIORITY,
            output_column="Tax/Business Number",
        )

    if rules["Same Email"]:
        multiple_emails = population["Email"].str.contains(
            r"[;,\r\n]",
            regex=True,
            na=False,
        )
        if multiple_emails.any():
            raise ValueError(
                "VM: multiple emails in one cell; "
                "provide one email per source row."
            )

    return population


def run_configured_vm(context, sheet_name):
    control_id = context["control"]["id_control"]
    rules = parse_vm_rules(context)

    print(
        f"{control_id} rules: "
        + "; ".join(
            f"{header}={'T' if enabled else 'F'}"
            for header, enabled in rules.items()
        )
    )

    population = load_matching_population(context, rules)
    pairs, metrics = find_vendor_pairs(population, rules)

    postings, posting_metadata = load_vm_vendor_postings(context)

    if (
        postings is None
        or not posting_metadata.get("available", False)
    ):
        raise FileNotFoundError(
            f"{control_id}: VPBSIK and VPBSAK are required "
            "for last-invoice columns."
        )

    invoices = build_vm_last_invoice_population(postings)

    names = population[
        IDENTITY + ["Company Name"]
    ].drop_duplicates(IDENTITY)

    details = names.merge(
        invoices,
        on=IDENTITY,
        how="left",
        validate="one_to_one",
    )

    for side in ("Left", "Right"):
        renamed = details.rename(
            columns={
                column: column + " " + side
                for column in details.columns
            }
        )

        pairs = pairs.merge(
            renamed,
            on=["Company " + side, "Vendor Code " + side],
            how="left",
            validate="many_to_one",
        )

    pairs.insert(0, "Group", range(1, len(pairs) + 1))
    pairs = pairs.rename(
        columns={
            "Company Left": "CoCo Left",
            "Company Right": "CoCo Right",
        }
    )

    output_file = write_vm_control_sheet(
        context=context,
        sheet_name=sheet_name,
        dataframe=pairs,
        integer_columns=["Group"],
        date_columns=[
            "Last Transaction Date Left",
            "Last Transaction Date Right",
        ],
        amount_columns=[
            "Last Inv Amt Doc Currency Left",
            "Last Inv Amt Doc Currency Right",
        ],
    )

    print(f"{control_id} matching vendor pairs: {len(pairs)}")

    for warning in posting_metadata.get("warnings", []):
        print(f"WARNING: {warning}")

    return {
        "status": "ERROR" if not pairs.empty else "OK",
        "output_file": output_file,
        "sheet_name": sheet_name,
        "rows": len(pairs),
    }
