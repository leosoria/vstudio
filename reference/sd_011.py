"""SD11: preliminary receipt / possible write-off conflict.

Independent runtime input: SD_INVENTORY for the configured cutoff.
Matches company, posting user and material during the configured period.
976/Z68 are possible write-offs/outbound adjustments, not confirmed losses.
Original activity is retained; reversing/reference positions are not triggers.
"""
from pathlib import Path
from collections import defaultdict, deque
import os
import re
import math
import tempfile
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from core.sd_common import (
    normalize_text, normalize_identifier, parse_config_date,
    parse_config_companies, parse_sap_date, parse_sap_amount,
    write_control_sheet,
)


DOC = ["Company", "Fiscal Year", "Material Document"]
KEY = DOC + ["Document Item"]
PHYSICAL = ["Fiscal Year", "Material Document", "Document Item"]
GROUP = ["Company", "User ID", "Material"]

VISIBLE = [
    "Company", "Material Document", "Fiscal Year", "Document Item",
    "Activity Type", "Trigger", "User ID", "Entry Date",
    "Movement Type", "Material", "Plant", "Quantity", "Base Unit",
    "Debit Amount LC", "Credit Amount LC", "Local Currency", "Item Text",
]

EXTRA = [
    "Document Key", "Conflict Groups",
    "Matched Receipt Documents", "Matched Possible Write-off Documents",
    "Matched Receipt Positions", "Matched Possible Write-off Positions",
    "Classification Status", "Link Basis",
    "Storage Location", "Batch", "Debit Credit", "Amount LC",
    "Document Type", "Posting Date (Source)", "Entry Time",
    "Transaction Code", "Order Number", "Order Item", "Vendor Code",
    "Reason", "Original Document", "Original Year", "Original Item",
    "Original Available In Inputs", "Referenced By Reversing Position",
    "Source File", "Source Row",
]

def ident(value, width=0, optional=False):
    value = normalize_identifier(value)
    value = value.zfill(width) if value.isdigit() else value.upper()
    return "" if optional and value and set(value) == {"0"} else value

def label(value):
    value = re.sub(r"\.\d+$", "", normalize_text(value)).lower()
    return re.sub(r"[^a-z0-9]", "", value)

def load_inventory(context):
    suffix = parse_config_date(
        context["module"]["to"], "TO"
    ).strftime("%Y%m%d")
    prefix = "lbrsdinventory" + suffix

    paths = sorted(
        p for p in Path(context["input_folder"]).iterdir()
        if p.is_file()
        and not p.name.startswith("~$")
        and p.suffix.lower() in {".xlsx", ".xlsm"}
        and re.fullmatch(re.escape(prefix) + r"\d*", label(p.stem))
    )

    if not paths:
        raise FileNotFoundError(
            "SD11: missing SD_INVENTORY input for configured TO"
        )

    mapping = {
        "Company": ("CoCd", 1),
        "Material Document": ("Mat. Doc.", 1),
        "Fiscal Year": ("MatYr", 1),
        "Document Item": ("Item", 1),
        "Document Type": ("Doc. Type", 1),
        "Posting Date (Source)": ("Pstng Date", 1),
        "Entry Date": ("Entry Dte", 1),
        "Entry Time": ("Time", 1),
        "User ID": ("User name", 1),
        "Movement Type": ("MvT", 1),
        "Material": ("Material", 1),
        "Plant": ("Plnt", 1),
        "Storage Location": ("SLoc", 1),
        "Batch": ("Batch", 1),
        "Debit Credit": ("D/C", 1),
        "Quantity": ("Quantity", 1),
        "Base Unit": ("BUn", 1),
        "Amount LC": ("Amount in LC", 1),
        "Local Currency": ("Crcy", 1),
        "Item Text": ("Text", 1),
        "Order Number": ("PO", 1),
        "Order Item": ("Item", 2),
        "Vendor Code": ("Vendor", 1),
        "Reason": ("Reas.", 1),
        "Original Document": ("Mat. Doc.", 2),
        "Original Year": ("MatYr", 2),
        "Original Item": ("Item", 3),
    }

    frames = []
    totals_removed = 0
    rows_read = 0

    for path in paths:
        print("SD_011 - Reading: " + path.name, flush=True)

        raw = pd.read_excel(
            path, sheet_name="Sheet1", header=None,
            dtype=object, keep_default_na=False,
        )
        if raw.empty:
            raise ValueError("SD11: empty input " + path.name)

        headings = [label(v) for v in raw.iloc[0]]
        body = raw.iloc[1:]
        frame = pd.DataFrame(index=body.index)

        for target, (source, occurrence) in mapping.items():
            found = [
                i for i, heading in enumerate(headings)
                if heading == label(source)
            ]
            if len(found) < occurrence:
                raise ValueError(
                    f"{path.name}: missing {source}, "
                    f"occurrence {occurrence}"
                )
            frame[target] = body.iloc[:, found[occurrence - 1]]

        found = [
            i for i, heading in enumerate(headings)
            if heading in {"mkpftcode", "tcode", "transactioncode"}
        ]
        frame["Transaction Code"] = (
            body.iloc[:, found[0]] if found else ""
        )

        texts = frame.apply(lambda column: column.map(normalize_text))
        blank = texts.eq("").all(axis=1)
        allowed = {"Quantity", "Base Unit", "Amount LC", "Local Currency"}

        totals = (
            texts[KEY].eq("").all(axis=1)
            & texts[
                [column for column in texts if column not in allowed]
            ].eq("").all(axis=1)
            & ~blank
        )

        bad = texts[KEY].eq("").any(axis=1) & ~blank & ~totals
        if bad.any():
            raise ValueError(
                f"{path.name}: incomplete keys at Excel rows "
                f"{(frame.index[bad] + 1).tolist()[:10]}"
            )

        rows_read += len(frame)
        totals_removed += int(totals.sum())
        frame = frame.loc[~blank & ~totals].copy()
        frame["Source File"] = path.name
        frame["Source Row"] = frame.index + 1
        frames.append(frame)

    data = pd.concat(frames, ignore_index=True)

    for column in data:
        if column not in {"Entry Date", "Quantity", "Amount LC", "Source Row"}:
            data[column] = data[column].map(normalize_text)

    for column, width in [
        ("Company", 4), ("Fiscal Year", 4),
        ("Material Document", 10), ("Document Item", 4),
        ("Movement Type", 3),
    ]:
        data[column] = data[column].map(
            lambda value, w=width: ident(value, w)
        )

    for column, width in [
        ("Original Document", 10), ("Original Year", 4),
        ("Original Item", 4), ("Order Number", 10),
        ("Order Item", 5), ("Vendor Code", 10),
    ]:
        data[column] = data[column].map(
            lambda value, w=width: ident(value, w, True)
        )

    for column in [
        "User ID", "Debit Credit", "Plant", "Transaction Code",
        "Document Type", "Local Currency",
    ]:
        data[column] = data[column].str.upper()

    data["Material"] = data["Material"].map(
        lambda value: ident(value, 18)
        if normalize_identifier(value).isdigit()
        else normalize_text(value)
    )

    for column, width in [
        ("Fiscal Year", 4), ("Material Document", 10),
        ("Document Item", 4),
    ]:
        if not data[column].str.fullmatch(
            r"\d{" + str(width) + r"}"
        ).all():
            raise ValueError("SD11: invalid " + column)

    if not data["Movement Type"].str.fullmatch(r"[A-Z0-9]{3}").all():
        raise ValueError("SD11: invalid movement code")

    if not data["Debit Credit"].isin({"S", "H"}).all():
        raise ValueError("SD11: invalid D/C")

    for column in ["Quantity", "Amount LC"]:
        data[column] = data[column].map(parse_sap_amount)
        if not data[column].map(
            lambda value: pd.notna(value) and math.isfinite(value)
        ).all():
            raise ValueError("SD11: invalid " + column)

    data["Entry Date"] = pd.to_datetime(
        data["Entry Date"].map(parse_sap_date), errors="coerce"
    )
    if data["Entry Date"].isna().any():
        raise ValueError("SD11: invalid entry dates")

    business = [
        column for column in data
        if column not in {"Source File", "Source Row"}
    ]
    distinct = data.drop_duplicates(business)

    if distinct.duplicated(PHYSICAL).any():
        raise ValueError(
            "SD11: contradictory material-document positions"
        )

    duplicate_count = len(data) - len(distinct)
    data = distinct.copy()

    start = parse_config_date(context["module"]["from"], "FROM")
    end = parse_config_date(context["module"]["to"], "TO")
    companies = parse_config_companies(
        context["module"].get("companies", "")
    )

    if companies is not None:
        data = data.loc[data["Company"].isin(companies)].copy()

    data = data.loc[data["Entry Date"].between(start, end)].copy()

    header_fields = DOC + [
        "User ID", "Entry Date", "Entry Time", "Document Type",
        "Posting Date (Source)", "Transaction Code",
    ]
    if data[header_fields].drop_duplicates().duplicated(DOC).any():
        raise ValueError(
            "SD11: contradictory material-document headers"
        )

    references = data[
        ["Original Document", "Original Year", "Original Item"]
    ].ne("")
    if (references.any(axis=1) & ~references.all(axis=1)).any():
        raise ValueError("SD11: incomplete reversal references")

    data["Document Key"] = [
        "|".join(row)
        for row in data[DOC].itertuples(index=False, name=None)
    ]

    stats = {
        "Files": len(paths),
        "Rows read": rows_read,
        "Export totals removed": totals_removed,
        "Identical duplicates removed": duplicate_count,
        "Scoped positions": len(data),
        "Material/plant/positive-quantity positions": int(
            (
                data["Material"].ne("")
                & data["Plant"].ne("")
                & data["Quantity"].gt(0)
            ).sum()
        ),
        "Negative-quantity positions outside trigger scope": int(
            data["Quantity"].lt(0).sum()
        ),
        "Positions without transaction code": int(
            data["Transaction Code"].eq("").sum()
        ),
    }
    return data, stats


def parameters(context):
    control = context.get("control", {})

    def codes(value, default):
        result = {
            token for token in re.split(
                r"[,;|\s]+", normalize_text(value).upper()
            ) if token
        } or set(default)
        if any(not re.fullmatch(r"[A-Z0-9]{3}", token) for token in result):
            raise ValueError("SD11: movement types must have 3 characters")
        return result

    return (
        codes(control.get("param1", ""), {"101"}),
        codes(control.get("param2", ""), {"976", "Z68"}),
    )


def select_cases(data, receipt_types, outbound_types, stats):
    required = set(
        KEY + GROUP + [
            "Plant", "Quantity", "Movement Type", "Debit Credit",
            "Original Document", "Original Year", "Original Item",
            "Amount LC", "Source File", "Source Row",
        ]
    )
    missing = required - set(data.columns)
    if missing:
        raise ValueError("SD11: unexpected inventory schema: " + repr(sorted(missing)))

    if data.duplicated(KEY).any():
        raise ValueError("SD11: repeated inventory position keys")

    stock = (
        data["Material"].ne("")
        & data["Plant"].ne("")
        & data["Quantity"].gt(0)
    )
    user = data["User ID"].ne("")
    reference = data["Original Document"].map(
        lambda value: bool(normalize_text(value).strip("0"))
    )

    receipt_base = (
        stock
        & data["Movement Type"].isin(receipt_types)
        & data["Debit Credit"].eq("S")
    )
    outbound_base = (
        stock
        & data["Movement Type"].isin(outbound_types)
        & data["Debit Credit"].eq("H")
    )

    receipts = data.loc[receipt_base & ~reference & user].copy()
    outbound = data.loc[outbound_base & ~reference & user].copy()

    groups = receipts[GROUP].drop_duplicates().merge(
        outbound[GROUP].drop_duplicates(),
        on=GROUP, how="inner", validate="one_to_one",
    )
    matched_receipts = receipts.merge(
        groups, on=GROUP, how="inner", validate="many_to_one",
    )
    matched_outbound = outbound.merge(
        groups, on=GROUP, how="inner", validate="many_to_one",
    )

    stats.update({
        "Receipt positions before exclusions": int(receipt_base.sum()),
        "Reversing/referenced receipt positions excluded":
            int((receipt_base & reference).sum()),
        "Eligible receipt positions": len(receipts),
        "Possible write-off positions before exclusions": int(outbound_base.sum()),
        "Reversing/referenced possible write-off positions excluded":
            int((outbound_base & reference).sum()),
        "Eligible possible write-off positions": len(outbound),
        "Candidate positions without user": int(
            ((receipt_base | outbound_base) & ~reference & ~user).sum()
        ),
        "Conflicting company/user/material groups": len(groups),
        "Exception possible write-off documents":
            len(matched_outbound[DOC].drop_duplicates()),
        "Related receipt documents":
            len(matched_receipts[DOC].drop_duplicates()),
        "Trigger positions": len(matched_outbound),
    })
    return matched_receipts, matched_outbound


def build_output(data, receipts, outbound, stats):
    if outbound.empty:
        stats["Output material documents"] = 0
        stats["Additional reference support documents"] = 0
        stats["Output positions"] = 0
        return pd.DataFrame(columns=VISIBLE + EXTRA)

    data = data.copy()
    row_keys = list(data[KEY].itertuples(index=False, name=None))
    docs = list(data[DOC].itertuples(index=False, name=None))
    receipt_keys = set(receipts[KEY].itertuples(index=False, name=None))
    outbound_keys = set(outbound[KEY].itertuples(index=False, name=None))

    receipt_by_group = defaultdict(set)
    outbound_by_group = defaultdict(set)
    receipt_count = defaultdict(int)
    outbound_count = defaultdict(int)
    groups_by_doc = defaultdict(set)

    for frame, destination, counts in [
        (receipts, receipt_by_group, receipt_count),
        (outbound, outbound_by_group, outbound_count),
    ]:
        for _, row in frame.iterrows():
            group = tuple(row[column] for column in GROUP)
            doc = tuple(row[column] for column in DOC)
            destination[group].add(doc)
            counts[group] += 1
            groups_by_doc[doc].add(group)

    initial = set(groups_by_doc)
    all_docs = set(docs)
    position_lookup = defaultdict(set)
    for key in row_keys:
        position_lookup[(key[1], key[2], key[3])].add(key[:3])

    # References provide support in either direction; never create triggers.
    neighbours = defaultdict(set)
    original_available = {}
    referenced_positions = set()

    for key, (_, row) in zip(row_keys, data.iterrows()):
        original = normalize_text(row["Original Document"])
        year = normalize_text(row["Original Year"])
        item = normalize_text(row["Original Item"])
        physical = (
            ident(year, 4), ident(original, 10), ident(item, 4)
        )
        available = (
            position_lookup.get(physical, set())
            if original.strip("0") and year.strip("0") and item.strip("0")
            else set()
        )
        original_available[key] = bool(available)

        for original_doc in available:
            neighbours[key[:3]].add(original_doc)
            neighbours[original_doc].add(key[:3])
            referenced_positions.add(original_doc + (physical[2],))

    selected = set(initial)
    queue = deque(initial)
    while queue:
        doc = queue.popleft()
        for neighbour in neighbours.get(doc, set()):
            if neighbour in all_docs and neighbour not in selected:
                selected.add(neighbour)
                queue.append(neighbour)

    mask = pd.Series([doc in selected for doc in docs], index=data.index)
    output = data.loc[mask].copy()
    output_keys = list(output[KEY].itertuples(index=False, name=None))

    def join_docs(values):
        return "; ".join("|".join(doc) for doc in sorted(values))

    roles = []
    conflict_values = []
    receipt_documents = []
    outbound_documents = []
    receipt_counts = []
    outbound_counts = []

    for key in output_keys:
        groups = groups_by_doc.get(key[:3], set())

        if key in outbound_keys:
            roles.append("POSSIBLE WRITE-OFF / OUTBOUND ADJUSTMENT")
        elif key in receipt_keys:
            roles.append("RECEIPT")
        else:
            roles.append("DOCUMENT / REFERENCE SUPPORT")

        conflict_values.append(
            "; ".join("|".join(group) for group in sorted(groups))
        )
        related_receipts = set()
        related_outbound = set()
        for group in groups:
            related_receipts.update(receipt_by_group[group])
            related_outbound.update(outbound_by_group[group])

        receipt_documents.append(join_docs(related_receipts))
        outbound_documents.append(join_docs(related_outbound))
        receipt_counts.append(sum(receipt_count[group] for group in groups))
        outbound_counts.append(sum(outbound_count[group] for group in groups))

    output["Activity Type"] = roles
    output["Trigger"] = [
        "YES" if key in outbound_keys else "NO" for key in output_keys
    ]
    output["Document Key"] = [
        "|".join(key[:3]) for key in output_keys
    ]
    output["Conflict Groups"] = conflict_values
    output["Matched Receipt Documents"] = receipt_documents
    output["Matched Possible Write-off Documents"] = outbound_documents
    output["Matched Receipt Positions"] = receipt_counts
    output["Matched Possible Write-off Positions"] = outbound_counts
    output["Original Available In Inputs"] = [
        "YES" if original_available[key] else "NO" for key in output_keys
    ]
    output["Referenced By Reversing Position"] = [
        "YES" if key in referenced_positions else "NO" for key in output_keys
    ]
    output["Classification Status"] = (
        "PRELIMINARY: possible write-off/outbound adjustment; "
        "loss, damage or scrap not confirmed"
    )
    output["Link Basis"] = (
        "Same company, posting user and material within configured period; "
        "no same-plant or sequence requirement"
    )

    if not output["Debit Credit"].isin({"S", "H"}).all():
        raise ValueError("SD11: invalid D/C in selected document support")
    if not output["Amount LC"].map(
        lambda value: pd.notna(value) and math.isfinite(value)
    ).all():
        raise ValueError("SD11: invalid local amount")

    output["Debit Amount LC"] = output["Amount LC"].where(
        output["Debit Credit"].eq("S"), 0.
    )
    output["Credit Amount LC"] = output["Amount LC"].where(
        output["Debit Credit"].eq("H"), 0.
    )
    output = output.reindex(columns=VISIBLE + EXTRA).sort_values(KEY)
    output = output.reset_index(drop=True)

    if output.duplicated(KEY).any():
        raise ValueError("SD11: repeated output positions")
    for column in output:
        if output[column].map(
            lambda value: isinstance(value, str) and len(value) > 32767
        ).any():
            raise ValueError("SD11: Excel text limit exceeded")

    stats["Output material documents"] = len(selected)
    stats["Additional reference support documents"] = len(selected - initial)
    stats["Output positions"] = len(output)
    return output


def present(path, output, stats, receipt_types, outbound_types):
    book = load_workbook(path)
    temporary = None
    try:
        sheet = book["SD11"]
        if not set(VISIBLE).issubset({
            cell.value for cell in sheet[1]
        }):
            raise ValueError("SD11: unexpected output header layout")

        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = (
            f"A1:{get_column_letter(len(output.columns))}{len(output) + 1}"
            if not output.empty else None
        )
        for index in range(len(VISIBLE) + 1, len(output.columns) + 1):
            dimension = sheet.column_dimensions[get_column_letter(index)]
            dimension.hidden = True
            dimension.outlineLevel = 1

        sheet.sheet_properties.outlinePr.summaryRight = False
        sheet.sheet_view.showOutlineSymbols = True
        sheet.sheet_format.outlineLevelCol = 1
        sheet.column_dimensions[
            get_column_letter(len(VISIBLE))
        ].collapsed = True

        notes = [
            "SD11 - Preliminary exceptions; write-off classification unconfirmed",
            (
                f"Detected {stats['Exception possible write-off documents']:,} "
                f"possible write-off/outbound adjustment documents; "
                f"{stats['Trigger positions']:,} trigger positions; "
                f"{len(output):,} output positions."
            ),
            (
                f"Receipts: {','.join(sorted(receipt_types))}/S. "
                f"Possible write-offs/outbound adjustments: "
                f"{','.join(sorted(outbound_types))}/H. "
                "Same company, posting user and material within configured period."
            ),
            (
                "Loss, damage or scrap is not confirmed. "
                "No same-plant or activity-sequence requirement. "
                "Original activity retained without netting; "
                "referenced/reversing entries are support, not triggers."
            ),
            (
                "All supplied positions of selected material documents and "
                "available linked reference documents are included. "
                "Material documents are not FI journals; no journal-balance "
                "requirement or SAP-wide completeness certification applies."
            ),
        ]

        start = len(output) + 4
        for offset, value in enumerate(notes):
            row = start + offset
            sheet.merge_cells(
                start_row=row, start_column=1,
                end_row=row, end_column=len(VISIBLE),
            )
            cell = sheet.cell(row, 1, value)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            sheet.row_dimensions[row].height = 32
            if offset == 0:
                cell.font = Font(bold=True)

        with tempfile.NamedTemporaryFile(
            dir=Path(path).parent, suffix=".xlsx", delete=False
        ) as handle:
            temporary = Path(handle.name)
        book.save(temporary)
        book.close()
        os.replace(temporary, path)
        temporary = None
    finally:
        book.close()
        if temporary is not None and temporary.exists():
            temporary.unlink()


def run_sd_011(context):
    receipt_types, outbound_types = parameters(context)
    data, stats = load_inventory(context)

    receipts, outbound = select_cases(
        data, receipt_types, outbound_types, stats
    )
    output = build_output(data, receipts, outbound, stats)

    print("SD_011 - Receipt movements:", ",".join(sorted(receipt_types)))
    print(
        "SD_011 - Possible write-off/outbound movements:",
        ",".join(sorted(outbound_types)),
    )
    for name, value in stats.items():
        print(f"SD_011 - {name}: {value}")

    path = write_control_sheet(
        context, "SD11", output,
        date_columns=["Entry Date"],
        amount_columns=["Debit Amount LC", "Credit Amount LC", "Amount LC"],
    )
    present(path, output, stats, receipt_types, outbound_types)

    print(
        "SD_011 - Match: same company/user/material during period; "
        "no same-plant or activity-sequence condition."
    )
    print(
        "SD_011 - Classification: PRELIMINARY. "
        "Loss, damage or scrap is not confirmed."
    )
    print(
        "SD_011 - Runtime input: SD_INVENTORY only. "
        "No other control or movement-type catalog is loaded."
    )
    print("SD_011 - output:", path)

    return {
        "status": "OK",
        "result_status": "PRELIMINARY",
        "output_file": path,
        "sheet_name": "SD11",
        "rows": len(output),
        "summary": stats,
    }


run = run_sd_011
