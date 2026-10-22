"""SD04: same user posts customer credit memos and refunds.

Matches company/customer/user within the configured entry-date period.
The link identifies shared activity, not settlement of a particular memo.
Defaults: DG/RV with 11/H; refunds 05/S; positive amounts; no reversals.
"""
from pathlib import Path
from collections import defaultdict
import re
import os
import tempfile
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from core.sd_common import (
    load_sd_ar_items, load_sd_ar_journals,
    load_sd_customer_master, normalize_text,
    write_control_sheet,
)

DOC = ["Company", "Fiscal Year", "Document Number"]
KEY = DOC + ["Document Item"]
GROUP = ["Company", "Customer Code", "User ID"]

VISIBLE = [
    "Company", "Document Number", "Fiscal Year", "Document Item",
    "Trigger", "Activity Type", "Customer Code", "Customer Name",
    "Journal Posted By", "Entry Date", "Debit Amount",
    "Credit Amount", "Document Currency", "Item Text",
    "Related Credit Memo Documents", "Related Refund Documents",
]
EXTRA = [
    "Journal Key", "Conflict Customer Codes", "Document Type",
    "Transaction Code", "Account Type", "Posting Key", "GL Account",
    "Vendor Code", "Debit Credit", "Amount in Document Currency",
    "Amount in Local Currency", "Local Currency",
    "AR Source File", "AR Source Row",
    "Journal Source File", "Journal Source Row", "Link Basis",
]


def parameters(context):
    control = context.get("control", {})
    types = {
        value for value in re.split(
            r"[,;|\s]+",
            normalize_text(control.get("param1", "")).upper(),
        ) if value
    } or {"DG", "RV"}
    keys = {
        value for value in re.split(
            r"[,;|\s]+",
            normalize_text(control.get("param2", "")),
        ) if value
    } or {"05"}

    if any(not re.fullmatch(r"[A-Z0-9]{2}", v) for v in types):
        raise ValueError(
            "SD04 PARAM1: expected credit memo types, e.g. DG,RV"
        )
    if any(not re.fullmatch(r"\d{1,2}", v) for v in keys):
        raise ValueError(
            "SD04 PARAM2: expected refund posting keys, e.g. 05"
        )
    return types, {value.zfill(2) for value in keys}


def select(ar, types, refund_keys):
    data = ar.copy()
    data["User ID"] = (
        data["User ID"].map(normalize_text).str.upper()
    )
    reversed_docs = set()

    for _, row in data.loc[
        data["Reversal Document"].ne("")
    ].iterrows():
        reversed_docs.add(tuple(row[name] for name in DOC))
        if row["Reversal Year"]:
            reversed_docs.add((
                row["Company"], row["Reversal Year"],
                row["Reversal Document"],
            ))

    reversal = pd.Series(
        [
            tuple(row) in reversed_docs
            for row in data[DOC].itertuples(
                index=False, name=None
            )
        ],
        index=data.index,
        dtype=bool,
    )
    reversal |= data["Transaction Code"].isin({
        "VF11", "FB08", "F.80"
    })

    credit_base = (
        data["Document Type"].isin(types)
        & data["Posting Key"].eq("11")
        & data["Debit Credit"].eq("H")
    )
    refund_base = (
        data["Posting Key"].isin(refund_keys)
        & data["Debit Credit"].eq("S")
    )
    positive = data["Amount in Document Currency"].gt(0)

    credits = data.loc[
        credit_base & ~reversal & positive
    ].copy()
    refunds = data.loc[
        refund_base & ~reversal & positive
    ].copy()

    credit_groups = credits.loc[
        credits["User ID"].ne(""), GROUP
    ].drop_duplicates()
    refund_groups = refunds.loc[
        refunds["User ID"].ne(""), GROUP
    ].drop_duplicates()

    groups = credit_groups.merge(
        refund_groups, on=GROUP, how="inner",
        validate="one_to_one",
    )
    linked_credit = credits.merge(
        groups, on=GROUP, how="inner",
        validate="many_to_one",
    )
    linked_refund = refunds.merge(
        groups, on=GROUP, how="inner",
        validate="many_to_one",
    )

    credit_keys = set(
        linked_credit[KEY].itertuples(
            index=False, name=None
        )
    )
    linked_refund_keys = set(
        linked_refund[KEY].itertuples(
            index=False, name=None
        )
    )

    credit_docs = defaultdict(set)
    refund_docs = defaultdict(set)

    for _, row in linked_credit.iterrows():
        group = tuple(row[name] for name in GROUP)
        credit_docs[group].add(
            "|".join(row[name] for name in DOC)
        )
    for _, row in linked_refund.iterrows():
        group = tuple(row[name] for name in GROUP)
        refund_docs[group].add(
            "|".join(row[name] for name in DOC)
        )

    linked = pd.concat(
        [linked_credit, linked_refund], ignore_index=True
    )
    relations = {}

    for _, row in linked.iterrows():
        doc = tuple(row[name] for name in DOC)
        group = tuple(row[name] for name in GROUP)
        relation = relations.setdefault(doc, {
            "customers": set(), "credits": set(),
            "refunds": set(), "roles": set(),
        })
        relation["customers"].add(row["Customer Code"])
        relation["credits"].update(credit_docs[group])
        relation["refunds"].update(refund_docs[group])
        relation["roles"].add(
            "CREDIT MEMO"
            if tuple(row[name] for name in KEY) in credit_keys
            else "REFUND"
        )

    stats = {
        "AR positions loaded": len(data),
        "Credit memo positions before exclusions":
            int(credit_base.sum()),
        "Credit memo positions excluded as reversals":
            int((credit_base & reversal).sum()),
        "Eligible credit memo positions": len(credits),
        "Refund positions before exclusions":
            int(refund_base.sum()),
        "Refund positions excluded as reversals":
            int((refund_base & reversal).sum()),
        "Eligible refund positions": len(refunds),
        "Eligible positions without posting user": int(
            credits["User ID"].eq("").sum()
            + refunds["User ID"].eq("").sum()
        ),
        "Conflicting company/customer/user groups": len(groups),
        "Exception journals": len(relations),
        "Exception trigger positions":
            len(credit_keys | linked_refund_keys),
    }
    return (
        linked, credit_keys, linked_refund_keys,
        relations, stats,
    )


def full_detail(
    ar, journals, linked, credit_keys,
    refund_keys, relations, customers,
):
    affected = pd.DataFrame(list(relations), columns=DOC)
    required = ar.merge(
        affected, on=DOC, how="inner",
        validate="many_to_one",
    )
    header_fields = DOC + [
        "User ID", "Entry Date", "Document Currency",
        "Document Type", "Transaction Code",
    ]
    headers = required[header_fields].drop_duplicates()
    if headers.duplicated(DOC).any():
        raise ValueError(
            "SD04: conflicting document header evidence"
        )

    detail = journals.merge(
        headers, on=DOC, how="inner",
        validate="many_to_one",
    )
    coverage = headers[DOC].merge(
        detail[DOC].drop_duplicates(),
        on=DOC, how="left", indicator=True,
    )
    if coverage["_merge"].eq("left_only").any():
        raise ValueError(
            "SD04: missing journal detail in supplied SD AR inputs. "
            "No partial result published; review extraction coverage."
        )

    comparison = required.merge(
        detail, on=KEY, how="left",
        suffixes=("_AR", "_BSEG"),
        indicator=True, validate="many_to_one",
    )
    bad = comparison["_merge"].ne("both")
    for field in [
        "Customer Code", "Posting Key", "Debit Credit"
    ]:
        bad |= comparison[field + "_AR"].ne(
            comparison[field + "_BSEG"]
        )
    for field in [
        "Amount in Document Currency",
        "Amount in Local Currency",
    ]:
        bad |= (
            comparison[field + "_AR"]
            - comparison[field + "_BSEG"]
        ).abs().gt(.01)

    if bad.any():
        raise ValueError(
            "SD04: missing or contradictory AR/BSEG positions. "
            "No partial result published."
        )

    if "BSEG Document Currency" in detail:
        mismatch = (
            detail["BSEG Document Currency"].ne("")
            & detail["BSEG Document Currency"].ne(
                detail["Document Currency"]
            )
        )
        if mismatch.any():
            raise ValueError(
                "SD04: BSEG and header currencies disagree"
            )

    sign = detail["Debit Credit"].map({"S": 1, "H": -1})
    balances = detail[DOC].copy()
    for field in [
        "Amount in Document Currency",
        "Amount in Local Currency",
    ]:
        balances[field] = detail[field] * sign

    if balances.groupby(DOC).sum().abs().gt(.01).any(axis=None):
        raise ValueError(
            "SD04: candidate journals do not balance. "
            "No partial result published."
        )

    names = customers[
        ["Company", "Customer Code", "Customer Name"]
    ].drop_duplicates()
    if names.duplicated(["Company", "Customer Code"]).any():
        raise ValueError("SD04: contradictory customer names")

    detail = detail.merge(
        names, on=["Company", "Customer Code"],
        how="left", validate="many_to_one",
    ).rename(columns={"User ID": "Journal Posted By"})

    sources = linked[
        KEY + ["Source File", "Source Row"]
    ].drop_duplicates(KEY).rename(columns={
        "Source File": "AR Source File",
        "Source Row": "AR Source Row",
    })
    detail = detail.merge(
        sources, on=KEY, how="left",
        validate="one_to_one",
    )

    keys = list(
        detail[KEY].itertuples(index=False, name=None)
    )
    detail["Trigger"] = [
        "YES" if key in credit_keys or key in refund_keys
        else "NO"
        for key in keys
    ]
    detail["Activity Type"] = [
        " / ".join(sorted(relations[key[:3]]["roles"]))
        for key in keys
    ]
    detail["Journal Key"] = [
        "|".join(key[:3]) for key in keys
    ]
    detail["Conflict Customer Codes"] = [
        ", ".join(sorted(relations[key[:3]]["customers"]))
        for key in keys
    ]
    detail["Related Credit Memo Documents"] = [
        ", ".join(sorted(relations[key[:3]]["credits"]))
        for key in keys
    ]
    detail["Related Refund Documents"] = [
        ", ".join(sorted(relations[key[:3]]["refunds"]))
        for key in keys
    ]
    detail["Debit Amount"] = (
        detail["Amount in Document Currency"].where(
            detail["Debit Credit"].eq("S"), 0.
        )
    )
    detail["Credit Amount"] = (
        detail["Amount in Document Currency"].where(
            detail["Debit Credit"].eq("H"), 0.
        )
    )
    detail["Link Basis"] = (
        "Same company, customer and posting user in configured "
        "period; no specific settlement link asserted"
    )
    return (
        detail.reindex(columns=VISIBLE + EXTRA)
        .sort_values(KEY).reset_index(drop=True)
    )


def present(path, output, stats, types, refund_keys):
    book = load_workbook(path)
    temporary = None
    try:
        sheet = book["SD04"]
        if output.empty:
            credit = stats["Eligible credit memo positions"]
            refund = stats["Eligible refund positions"]

            if not refund:
                title = "SD04 - No eligible customer refunds"
                message = (
                    f"Evaluated {credit:,} credit memo positions. "
                    "No eligible customer refunds were found."
                )
            elif not credit:
                title = "SD04 - No eligible credit memos"
                message = (
                    f"Evaluated {refund:,} refund positions. "
                    "No eligible credit memos were found."
                )
            else:
                title = "SD04 - No exceptions"
                message = (
                    f"Evaluated {credit:,} credit memo and "
                    f"{refund:,} refund positions. "
                    "No company/customer/user group performed "
                    "both activities."
                )

            if stats["Eligible positions without posting user"]:
                title += "; missing user evidence"

            texts = {
                3: title,
                4: message,
                5: (
                    f"Credit memos: {','.join(sorted(types))}, 11/H. "
                    f"Refunds: {','.join(sorted(refund_keys))}/S. "
                    "Positive amounts; reversals excluded."
                ),
                6: (
                    "Both activities within configured entry-date "
                    "period. Matching is by company, customer "
                    "and posting user."
                ),
            }
            for row, value in texts.items():
                sheet.merge_cells(
                    start_row=row, start_column=1,
                    end_row=row, end_column=16,
                )
                cell = sheet.cell(row, 1, value)
                cell.alignment = Alignment(
                    wrap_text=True, vertical="top"
                )
                sheet.row_dimensions[row].height = 30
                if row == 3:
                    cell.font = Font(bold=True)
            sheet.auto_filter.ref = None
        else:
            for index in range(
                len(VISIBLE) + 1, len(output.columns) + 1
            ):
                dimension = sheet.column_dimensions[
                    get_column_letter(index)
                ]
                dimension.hidden = True
                dimension.outlineLevel = 1
            sheet.sheet_properties.outlinePr.summaryRight = False
            sheet.sheet_view.showOutlineSymbols = True
            sheet.sheet_format.outlineLevelCol = 1
            sheet.column_dimensions[
                get_column_letter(len(VISIBLE))
            ].collapsed = True

        with tempfile.NamedTemporaryFile(
            dir=Path(path).parent,
            suffix=".xlsx", delete=False,
        ) as handle:
            temporary = Path(handle.name)

        book.save(temporary)
        book.close()
        os.replace(temporary, path)
    finally:
        book.close()
        if temporary is not None and temporary.exists():
            temporary.unlink()


def run_sd_004(context):
    types, refund_keys = parameters(context)
    ar, _ = load_sd_ar_items(context)
    (
        linked, credit_keys, linked_refund_keys,
        relations, stats,
    ) = select(ar, types, refund_keys)

    print(
        "SD_004 - Credit memo types:",
        ",".join(sorted(types)),
    )
    print("SD_004 - Credit memo posting key / D-C: 11 / H")
    print(
        "SD_004 - Refund posting keys / D-C:",
        ",".join(sorted(refund_keys)), "/ S",
    )
    for name, value in stats.items():
        print(f"SD_004 - {name}: {value}")

    output = pd.DataFrame(columns=VISIBLE + EXTRA)
    if relations:
        journals, _ = load_sd_ar_journals(context)
        customers, _ = load_sd_customer_master(context)
        output = full_detail(
            ar, journals, linked, credit_keys,
            linked_refund_keys, relations, customers,
        )

    path = write_control_sheet(
        context, "SD04", output,
        date_columns=["Entry Date"],
        amount_columns=[
            "Debit Amount", "Credit Amount",
            "Amount in Document Currency",
            "Amount in Local Currency",
        ],
    )
    present(path, output, stats, types, refund_keys)

    print(
        "SD_004 - Link basis: same company/customer/posting user "
        "within configured period; no specific settlement link asserted."
    )
    print(
        "SD_004 - Scope: supplied AR extracts; configured "
        "document/posting keys; positive amounts; reversals excluded."
    )
    print(f"SD_004 - Output positions: {len(output)}")
    print(f"SD_004 - output: {path}")

    return {
        "status": "OK", "output_file": path,
        "sheet_name": "SD04", "rows": len(output),
        "summary": stats,
    }


run = run_sd_004
