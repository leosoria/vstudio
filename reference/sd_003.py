"""SD03: same user creates an order and posts a debit goods receipt."""
from pathlib import Path
import os
import tempfile
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

from core.sd_common import (
    _sdp_load_po_lines,
    parse_sap_date,
    parse_sap_amount,
    normalize_identifier,
    normalize_text,
    normalize_company,
    normalize_lookup,
    get_period_suffix,
    parse_config_date,
    parse_config_companies,
    write_control_sheet,
)
from core.sd_common import enrich_sd_names

CONTROL_ID = "SD_003"
SHEET_NAME = "SD03"
KEY = ["Company", "Order Number", "Order Item"]
DOC = KEY[:2]
GRKEY = KEY + ["Document Number", "Fiscal Year", "Document Item"]

COLUMNS = KEY + ['Activity Type', 'Trigger', 'Vendor Code', 'Vendor Name', 'Order Created By', 'Order Created By Name', 'Receipt Posted By', 'Receipt Posted By Name', 'Order Date', 'Entry Date', 'Document Number', 'Fiscal Year', 'Document Item', 'Movement Type', 'Debit Credit', 'Quantity', 'Amount', 'Currency', 'Item Text', 'Source File', 'Source Row']


def _unique(frame, key, label):
    business = [
        c for c in frame
        if c not in {"Source File", "Source Row"}
    ]
    clean = frame.drop_duplicates(business)
    if clean.duplicated(key, keep=False).any():
        raise ValueError(
            f"{label}: contradictory records for the same key"
        )
    return clean.drop_duplicates(key).copy()


def load_orders(context):
    raw, metrics = _sdp_load_po_lines(
        context,
        required_fields=(
            "Company", "PO Number", "PO Line", "PO Creator ID",
            "PO Doc Date", "Vendor Code", "PO Quantity",
            "PO Line Total", "PO Doc Currency",
        ),
    )
    mapping = {
        "Company": "Company",
        "PO Number": "Order Number",
        "PO Line": "Order Item",
        "PO Creator ID": "Order Created By",
        "PO Doc Date": "Order Date",
        "Vendor Code": "Vendor Code",
        "PO Quantity": "Quantity",
        "PO Line Total": "Amount",
        "PO Doc Currency": "Currency",
        "PO Material Description": "Item Text",
    }
    data = pd.DataFrame(index=raw.index)
    for source, target in mapping.items():
        data[target] = raw[source] if source in raw else ""

    data["Source File"] = Path(metrics["input_file"]).name
    data["Source Row"] = raw.index + 2
    data["Company"] = data["Company"].map(normalize_company)

    for col, width in [
        ("Order Number", 10), ("Order Item", 5), ("Vendor Code", 10)
    ]:
        data[col] = data[col].map(normalize_identifier).map(
            lambda v: v.zfill(width) if v.isdigit() else v
        )

    data["Order Created By"] = (
        data["Order Created By"].map(normalize_text).str.upper()
    )
    data["Order Date"] = data["Order Date"].map(parse_sap_date)
    for col in ["Quantity", "Amount"]:
        data[col] = data[col].map(parse_sap_amount)

    data = _unique(data, KEY, "SD03 ORDERS")
    known = data.loc[data["Order Created By"].ne("")]
    if known.groupby(DOC)["Order Created By"].nunique().gt(1).any():
        raise ValueError(
            "SD03: more than one creator for the same order"
        )

    # The creator is a header attribute of the same order.
    data["Order Created By"] = (
        data.groupby(DOC)["Order Created By"].transform(
            lambda s: next((v for v in s if v), "")
        )
    )
    return data


def load_receipts(context):
    expected = normalize_lookup(
        "LBR SD RECEIPTS " + get_period_suffix(context["module"])
    )
    files = [
        p for p in Path(context["input_folder"]).rglob("*")
        if p.is_file()
        and not p.name.startswith("~$")
        and p.suffix.lower() in {".xlsx", ".xlsm"}
        and normalize_lookup(p.stem) == expected
    ]
    if len(files) != 1:
        raise ValueError(
            f"SD03: expected one SD_RECEIPTS file; found {len(files)}"
        )

    raw = pd.read_excel(files[0], sheet_name="Sheet1", dtype=object)
    mapping = {
        "CoCd": "Company",
        "Purch.Doc.": "Order Number",
        "Item": "Order Item",
        "Mat. Doc.": "Document Number",
        "MatYr": "Fiscal Year",
        "Item.1": "Document Item",
        "Tr./ev.type": "History Event",
        "MvT": "Movement Type",
        "D/C": "Debit Credit",
        "Quantity": "Quantity",
        "Amount": "Amount",
        "Crcy": "Currency",
        "User name": "Receipt Posted By",
        "Entry Dte": "Entry Date",
    }
    missing = set(mapping) - set(raw)
    if missing:
        raise ValueError(
            f"SD03 RECEIPTS missing columns: {sorted(missing)}"
        )

    data = raw[list(mapping)].rename(columns=mapping).copy()
    data["Source File"] = files[0].name
    data["Source Row"] = data.index + 2

    for col in GRKEY + [
        "History Event", "Movement Type",
        "Debit Credit", "Receipt Posted By"
    ]:
        data[col] = (
            data[col].map(normalize_identifier).str.upper()
        )

    no_key = data[GRKEY].eq("").all(axis=1)
    evidence_fields = [
        "History Event", "Movement Type",
        "Debit Credit", "Receipt Posted By",
    ]
    totals = (
        no_key
        & data[evidence_fields].eq("").all(axis=1)
        & raw["Entry Dte"].map(normalize_text).eq("")
    )
    data = data.loc[~totals].copy()

    patterns = {
        "Company": r"\d{1,4}",
        "Order Number": r"\d{1,10}",
        "Order Item": r"\d{1,5}",
        "Document Number": r"\d{1,10}",
        "Fiscal Year": r"\d{4}",
        "Document Item": r"\d{1,4}",
    }
    for col, pattern in patterns.items():
        if not data[col].str.fullmatch(pattern).all():
            raise ValueError(
                f"SD03 RECEIPTS: incomplete/invalid {col}"
            )

    data["Company"] = data["Company"].map(normalize_company)
    for col, width in [
        ("Order Number", 10), ("Order Item", 5),
        ("Document Number", 10), ("Document Item", 4),
    ]:
        data[col] = data[col].str.zfill(width)

    data["Entry Date"] = data["Entry Date"].map(parse_sap_date)
    if data["Entry Date"].isna().any():
        raise ValueError(
            "SD03 RECEIPTS: missing or invalid Entry Dte"
        )

    for col in ["Quantity", "Amount"]:
        data[col] = data[col].map(parse_sap_amount)

    if not data["Debit Credit"].isin(["S", "H"]).all():
        raise ValueError(
            "SD03 RECEIPTS: invalid debit/credit indicator"
        )
    if data["Quantity"].isna().any():
        raise ValueError("SD03 RECEIPTS: missing quantity")

    data = _unique(data, GRKEY, "SD03 RECEIPTS")

    companies = parse_config_companies(
        context["module"].get("companies", "")
    )
    if companies is not None:
        data = data.loc[data["Company"].isin(companies)].copy()

    start = parse_config_date(context["module"]["from"], "FROM")
    end = parse_config_date(context["module"]["to"], "TO")
    data = data.loc[
        data["Entry Date"].between(start, end)
    ].copy()

    return data, int(totals.sum())


def analyze(orders, receipts):
    attached = receipts.merge(
        orders[KEY + [
            "Order Created By", "Order Date",
            "Vendor Code", "Item Text",
        ]],
        on=KEY,
        how="left",
        validate="many_to_one",
        indicator=True,
    )

    linked = attached["_merge"].eq("both")
    candidate = (
        linked
        & attached["History Event"].eq("1")
        & attached["Debit Credit"].eq("S")
        & attached["Quantity"].gt(0)
        & attached["Order Date"].le(attached["Entry Date"])
    )
    evidence = (
        attached["Order Created By"].fillna("").ne("")
        & attached["Receipt Posted By"].ne("")
    )
    conflict = (
        candidate
        & evidence
        & attached["Order Created By"].eq(
            attached["Receipt Posted By"]
        )
    )

    affected = attached.loc[conflict, DOC].drop_duplicates()
    detail = attached.merge(
        affected, on=DOC, how="inner", validate="many_to_one"
    )
    trigger_keys = set(
        attached.loc[conflict, GRKEY].itertuples(
            index=False, name=None
        )
    )
    detail["Trigger"] = [
        "YES" if key in trigger_keys else "NO"
        for key in detail[GRKEY].itertuples(
            index=False, name=None
        )
    ]
    detail["Activity Type"] = "RECEIPT"

    base = orders.merge(
        affected, on=DOC, how="inner", validate="many_to_one"
    )
    base["Trigger"] = "NO"
    base["Activity Type"] = "ORDER"

    output = pd.concat(
        [base, detail], ignore_index=True
    ).reindex(columns=COLUMNS)
    output = output.sort_values(
        DOC + ["Activity Type", "Order Item", "Document Number"],
        kind="stable",
    ).reset_index(drop=True)

    headers = orders[
        DOC + ["Order Created By"]
    ].drop_duplicates()

    stats = {
        "Orders in supplied scope": len(headers),
        "Orders without creator": int(
            headers["Order Created By"].eq("").sum()
        ),
        "Receipt positions in entry-date scope": len(receipts),
        "Receipt positions without matching supplied order item": int(
            (~linked).sum()
        ),
        "Candidate debit receipt positions": int(candidate.sum()),
        "Candidate positions without creator or receiver": int(
            (candidate & ~evidence).sum()
        ),
        "Evaluated receipt positions with user evidence": int(
            (candidate & evidence).sum()
        ),
        "Receipt positions before order date": int(
            (
                linked
                & attached["Order Date"].gt(attached["Entry Date"])
            ).sum()
        ),
        "Exception orders": len(affected),
        "Exception receipt positions": int(conflict.sum()),
        "Output positions": len(output),
    }
    return output, stats


def _note(path, output, metrics):
    book = load_workbook(path)
    temporary = None
    try:
        sheet = book[SHEET_NAME]
        if output.empty:
            sheet.merge_cells("A3:P3")
            sheet["A3"] = (
                "SD03 - No exceptions; coverage limitations"
            )
            sheet["A3"].font = Font(bold=True)

            sheet.merge_cells("A4:P7")
            count = metrics[
                "Evaluated receipt positions with user evidence"
            ]
            sheet["A4"] = (
                f"Evaluated {count:,} debit receipt positions "
                "with user evidence. "
                "No user matched the order creator. "
                "Scope: supplied period-dated orders and VGABE=1 "
                "receipts entered in the period. "
                f"Orders without creator: "
                f"{metrics['Orders without creator']:,}. "
                f"Unmatched receipt positions: "
                f"{metrics['Receipt positions without matching supplied order item']:,}. "
                "Service coverage outside VGABE=1 is not demonstrated. "
                "Receipt reversals are supporting history, "
                "not trigger positions."
            )
            sheet["A4"].alignment = Alignment(
                wrap_text=True, vertical="top"
            )
            for row in range(4, 8):
                sheet.row_dimensions[row].height = 24
            sheet.auto_filter.ref = None
        else:
            sheet.freeze_panes = "F2"

        with tempfile.NamedTemporaryFile(
            dir=Path(path).parent, suffix=".xlsx", delete=False
        ) as handle:
            temporary = Path(handle.name)

        book.save(temporary)
        book.close()
        os.replace(temporary, path)
    finally:
        book.close()
        if temporary is not None and temporary.exists():
            temporary.unlink()


def run_sd_003(context):
    orders = load_orders(context)
    receipts, totals = load_receipts(context)
    output, metrics = analyze(orders, receipts)
    metrics["Receipt total/blank rows removed"] = totals

    for name, value in metrics.items():
        print(f"SD_003 - {name}: {value}")

    print(
        "SD_003 - Scope: period-dated supplied orders; "
        "VGABE=1 debit receipts entered in period; "
        "reversals retained as support; "
        "service coverage not demonstrated."
    )

    output = enrich_sd_names(
        context, output,
        user_columns={'Order Created By': 'Order Created By Name', 'Receipt Posted By': 'Receipt Posted By Name'},
    )
    path = write_control_sheet(
        context,
        SHEET_NAME,
        output,
        date_columns=["Order Date", "Entry Date"],
        amount_columns=["Quantity", "Amount"],
    )
    _note(path, output, metrics)
    print(f"SD_003 - output: {path}")

    return {
        "status": "OK",
        "output_file": path,
        "sheet_name": SHEET_NAME,
        "rows": len(output),
        "summary": metrics,
    }


run = run_sd_003
