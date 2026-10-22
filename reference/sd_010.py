"""SD10: same company/user posts goods receipts and inventory adjustments.

Independent input: LBR SD_INVENTORY_<TO>[_part].xlsx.
Defaults: receipts 101/S; adjustments 702,975,976,Z68,Z69.
Matching is by company/user during the period.
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
GROUP = ["Company", "User ID"]

VISIBLE = [
    "Company", "Material Document", "Fiscal Year", "Document Item",
    "Activity Type", "Trigger", "User ID", "Entry Date",
    "Movement Type", "Material", "Plant", "Quantity", "Base Unit",
    "Debit Amount LC", "Credit Amount LC", "Local Currency", "Item Text",
]

EXTRA = [
    "Document Key", "Conflict Groups", "Matched Receipt Positions",
    "Matched Adjustment Positions", "Storage Location", "Batch",
    "Debit Credit", "Amount LC", "Document Type", "Posting Date (Source)",
    "Entry Time", "Transaction Code", "Order Number", "Order Item",
    "Vendor Code", "Reason", "Original Document", "Original Year",
    "Original Item", "Original Available In Inputs",
    "Referenced By Reversing Position", "Source File", "Source Row",
]


def ident(value, width=0, optional=False):
    value = normalize_identifier(value)
    value = value.zfill(width) if value.isdigit() else value.upper()
    return "" if optional and value and set(value) == {"0"} else value


def label(value):
    value = re.sub(r"\.\d+$", "", normalize_text(value)).lower()
    return re.sub(r"[^a-z0-9]", "", value)


def parameters(context):
    def codes(parameter, default):
        value = normalize_text(
            context.get("control", {}).get(parameter)
        ).upper()
        result = {
            v for v in re.split(r"[,;|\s]+", value) if v
        } or set(default)

        if any(not re.fullmatch(r"[A-Z0-9]{3}", v) for v in result):
            raise ValueError(
                f"SD10 {parameter}: expected 3-character movement codes"
            )
        return result

    receipts = codes("param1", {"101"})
    adjustments = codes("param2", {"702", "975", "976", "Z68", "Z69"})

    if receipts & adjustments:
        raise ValueError(
            "SD10: receipt and adjustment movement lists overlap"
        )

    return receipts, adjustments


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
            "SD10: missing SD_INVENTORY input for configured TO"
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
        print("SD_010 - Reading: " + path.name, flush=True)

        raw = pd.read_excel(
            path, sheet_name="Sheet1", header=None,
            dtype=object, keep_default_na=False,
        )
        if raw.empty:
            raise ValueError("SD10: empty input " + path.name)

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
            raise ValueError("SD10: invalid " + column)

    if not data["Movement Type"].str.fullmatch(r"[A-Z0-9]{3}").all():
        raise ValueError("SD10: invalid movement code")

    if not data["Debit Credit"].isin({"S", "H"}).all():
        raise ValueError("SD10: invalid D/C")

    for column in ["Quantity", "Amount LC"]:
        data[column] = data[column].map(parse_sap_amount)
        if not data[column].map(
            lambda value: pd.notna(value) and math.isfinite(value)
        ).all():
            raise ValueError("SD10: invalid " + column)

    data["Entry Date"] = pd.to_datetime(
        data["Entry Date"].map(parse_sap_date), errors="coerce"
    )
    if data["Entry Date"].isna().any():
        raise ValueError("SD10: invalid entry dates")

    business = [
        column for column in data
        if column not in {"Source File", "Source Row"}
    ]
    distinct = data.drop_duplicates(business)

    if distinct.duplicated(PHYSICAL).any():
        raise ValueError(
            "SD10: contradictory material-document positions"
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
            "SD10: contradictory material-document headers"
        )

    references = data[
        ["Original Document", "Original Year", "Original Item"]
    ].ne("")
    if (references.any(axis=1) & ~references.all(axis=1)).any():
        raise ValueError("SD10: incomplete reversal references")

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


def select(data, receipt_types, adjustment_types, stats):
    stock = (
        data["Material"].ne("")
        & data["Plant"].ne("")
        & data["Quantity"].gt(0)
    )
    reversing = data["Original Document"].ne("")

    receipt_base = (
        stock
        & data["Movement Type"].isin(receipt_types)
        & data["Debit Credit"].eq("S")
    )
    adjustment_base = stock & data["Movement Type"].isin(adjustment_types)

    eligible = (receipt_base | adjustment_base) & ~reversing
    if data.loc[eligible, "User ID"].eq("").any():
        raise ValueError("SD10: qualifying activity without user")

    receipts = data.loc[receipt_base & ~reversing].copy()
    adjustments = data.loc[adjustment_base & ~reversing].copy()

    receipt_groups = set(
        receipts[GROUP].itertuples(index=False, name=None)
    )
    adjustment_groups = set(
        adjustments[GROUP].itertuples(index=False, name=None)
    )
    conflicts = receipt_groups & adjustment_groups

    receipts = receipts.loc[
        [
            group in conflicts
            for group in receipts[GROUP].itertuples(index=False, name=None)
        ]
    ].copy()
    adjustments = adjustments.loc[
        [
            group in conflicts
            for group in adjustments[GROUP].itertuples(index=False, name=None)
        ]
    ].copy()

    stats.update({
        "Receipt positions before reversal exclusion":
            int(receipt_base.sum()),
        "Reversing receipt positions excluded":
            int((receipt_base & reversing).sum()),
        "Adjustment positions before reversal exclusion":
            int(adjustment_base.sum()),
        "Reversing adjustment positions excluded":
            int((adjustment_base & reversing).sum()),
        "Eligible receipt positions":
            int((receipt_base & ~reversing).sum()),
        "Eligible adjustment positions":
            int((adjustment_base & ~reversing).sum()),
        "Conflicting company/user groups": len(conflicts),
        "Distinct conflicting users": len({group[1] for group in conflicts}),
        "Exception adjustment documents":
            len(adjustments[DOC].drop_duplicates()),
        "Trigger adjustment positions": len(adjustments),
        "Related receipt documents": len(receipts[DOC].drop_duplicates()),
    })

    return receipts, adjustments


def chunks(values):
    result = []
    current = ""

    for value in sorted(values):
        candidate = current + "; " + value if current else value
        if len(candidate) > 30000:
            result.append(current)
            current = value
        else:
            current = candidate

    if current:
        result.append(current)

    return result


def build_output(data, receipts, adjustments, stats):
    if adjustments.empty:
        stats["Output positions"] = 0
        return pd.DataFrame(columns=VISIBLE + EXTRA)

    receipt_keys = set(
        receipts[KEY].itertuples(index=False, name=None)
    )
    adjustment_keys = set(
        adjustments[KEY].itertuples(index=False, name=None)
    )

    groups = defaultdict(lambda: {
        "receipts": set(), "adjustments": set(),
        "receipt_positions": 0, "adjustment_positions": 0,
    })
    seeds = defaultdict(set)

    for frame, role in [
        (receipts, "receipts"), (adjustments, "adjustments")
    ]:
        for row in frame.to_dict("records"):
            group = (row["Company"], row["User ID"])
            document = tuple(row[column] for column in DOC)

            groups[group][role].add("|".join(document))
            count_field = (
                "receipt_positions"
                if role == "receipts"
                else "adjustment_positions"
            )
            groups[group][count_field] += 1
            seeds[document].add(group)

    physical_to_doc = {
        tuple(row[column] for column in PHYSICAL):
        tuple(row[column] for column in DOC)
        for row in data.to_dict("records")
    }

    graph = defaultdict(set)
    referenced = set()
    original_available = {}

    for row in data.to_dict("records"):
        key = tuple(row[column] for column in KEY)
        document = key[:3]
        original = (
            row["Original Year"],
            row["Original Document"],
            row["Original Item"],
        )
        target = (
            physical_to_doc.get(original)
            if row["Original Document"]
            else None
        )

        original_available[key] = (
            "YES" if target
            else "NO" if row["Original Document"]
            else ""
        )

        if target:
            graph[document].add(target)
            graph[target].add(document)
            referenced.add(original)

    related = {
        document: set(value) for document, value in seeds.items()
    }
    queue = deque(related)

    while queue:
        document = queue.popleft()
        for neighbour in graph[document]:
            before = len(related.get(neighbour, set()))
            related.setdefault(neighbour, set()).update(related[document])
            if len(related[neighbour]) > before:
                queue.append(neighbour)

    selected = set(related)
    mask = [
        document in selected
        for document in data[DOC].itertuples(index=False, name=None)
    ]
    output = data.loc[mask].copy()

    metadata = []
    max_receipts = 0
    max_adjustments = 0

    for document, conflicts in related.items():
        receipt_documents = set().union(
            *(groups[group]["receipts"] for group in conflicts)
        )
        adjustment_documents = set().union(
            *(groups[group]["adjustments"] for group in conflicts)
        )

        record = dict(zip(DOC, document))
        record["Conflict Groups"] = "; ".join(
            "|".join(group) for group in sorted(conflicts)
        )
        record["Matched Receipt Positions"] = sum(
            groups[group]["receipt_positions"] for group in conflicts
        )
        record["Matched Adjustment Positions"] = sum(
            groups[group]["adjustment_positions"] for group in conflicts
        )

        receipt_chunks = chunks(receipt_documents)
        adjustment_chunks = chunks(adjustment_documents)
        max_receipts = max(max_receipts, len(receipt_chunks))
        max_adjustments = max(max_adjustments, len(adjustment_chunks))

        for index, value in enumerate(receipt_chunks, 1):
            record[f"Related Receipt Documents {index}"] = value
        for index, value in enumerate(adjustment_chunks, 1):
            record[f"Related Adjustment Documents {index}"] = value

        metadata.append(record)

    output = output.merge(
        pd.DataFrame(metadata), on=DOC,
        how="left", validate="many_to_one",
    )

    row_keys = list(
        output[KEY].itertuples(index=False, name=None)
    )
    output["Activity Type"] = [
        "ADJUSTMENT" if key in adjustment_keys
        else "RECEIPT" if key in receipt_keys
        else "SUPPORT"
        for key in row_keys
    ]
    output["Trigger"] = [
        "YES" if key in adjustment_keys else "NO"
        for key in row_keys
    ]
    output["Original Available In Inputs"] = [
        original_available[key] for key in row_keys
    ]
    output["Referenced By Reversing Position"] = [
        "YES" if (key[1], key[2], key[3]) in referenced else "NO"
        for key in row_keys
    ]

    output["Debit Amount LC"] = output["Amount LC"].where(
        output["Debit Credit"].eq("S"), 0.
    )
    output["Credit Amount LC"] = output["Amount LC"].where(
        output["Debit Credit"].eq("H"), 0.
    )

    columns = (
        VISIBLE + EXTRA
        + [
            f"Related Receipt Documents {index}"
            for index in range(1, max_receipts + 1)
        ]
        + [
            f"Related Adjustment Documents {index}"
            for index in range(1, max_adjustments + 1)
        ]
    )
    output = (
        output.reindex(columns=columns)
        .sort_values(KEY).reset_index(drop=True)
    )

    if output.duplicated(KEY).any():
        raise ValueError("SD10: duplicate output positions")

    expected = set(
        data.loc[mask, KEY].itertuples(index=False, name=None)
    )
    actual = set(
        output[KEY].itertuples(index=False, name=None)
    )
    if actual != expected:
        raise ValueError("SD10: source/output position mismatch")

    for column in output:
        too_long = output[column].map(
            lambda value: len(value) > 32767
            if isinstance(value, str) else False
        )
        if too_long.any():
            raise ValueError("SD10: Excel text limit exceeded")

    stats["Output material documents"] = len(selected)
    stats["Additional reversal/original support documents"] = len(
        selected - set(seeds)
    )
    stats["Output positions"] = len(output)

    return output


def present(path, output, stats, receipt_types, adjustment_types):
    book = load_workbook(path)
    temporary = None

    try:
        sheet = book["SD10"]

        if output.empty:
            title = "SD10 - No exceptions"
            if not stats["Eligible receipt positions"]:
                title = "SD10 - No eligible goods receipts"
            elif not stats["Eligible adjustment positions"]:
                title = "SD10 - No eligible inventory adjustments"

            texts = {
                3: title,
                4: (
                    f"Evaluated {stats['Eligible receipt positions']:,} "
                    f"receipt and {stats['Eligible adjustment positions']:,} "
                    "adjustment positions. No company/user group "
                    "performed both activities."
                ),
                5: (
                    f"Receipts: {','.join(sorted(receipt_types))}/S. "
                    f"Adjustments: {','.join(sorted(adjustment_types))}. "
                    "Material, plant and positive quantity required; "
                    "reversing entries excluded from triggers."
                ),
                6: (
                    "Both activities in configured entry-date period. "
                    "Matching key: company/user. Material, plant and "
                    "time order are context, not matching conditions."
                ),
            }

            for row, value in texts.items():
                sheet.merge_cells(
                    start_row=row, start_column=1,
                    end_row=row, end_column=len(VISIBLE),
                )
                cell = sheet.cell(row, 1, value)
                cell.alignment = Alignment(
                    wrap_text=True, vertical="top"
                )
                if row == 3:
                    cell.font = Font(bold=True)
                sheet.row_dimensions[row].height = 30

            sheet.auto_filter.ref = None
        else:
            sheet.column_dimensions.group(
                get_column_letter(len(VISIBLE) + 1),
                get_column_letter(len(output.columns)),
                outline_level=1, hidden=True,
            )
            sheet.sheet_properties.outlinePr.summaryRight = False
            sheet.sheet_view.showOutlineSymbols = True
            sheet.sheet_format.outlineLevelCol = 1
            sheet.column_dimensions[
                get_column_letter(len(VISIBLE))
            ].collapsed = True

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


def run_sd_010(context):
    receipt_types, adjustment_types = parameters(context)
    data, stats = load_inventory(context)

    receipts, adjustments = select(
        data, receipt_types, adjustment_types, stats
    )
    output = build_output(data, receipts, adjustments, stats)

    print(
        "SD_010 - Receipt movements: "
        + ",".join(sorted(receipt_types))
    )
    print(
        "SD_010 - Adjustment movements: "
        + ",".join(sorted(adjustment_types))
    )
    for name, value in stats.items():
        print(f"SD_010 - {name}: {value}")

    path = write_control_sheet(
        context, "SD10", output,
        date_columns=["Entry Date"],
        amount_columns=[
            "Debit Amount LC", "Credit Amount LC", "Amount LC",
        ],
    )
    present(path, output, stats, receipt_types, adjustment_types)

    print(
        "SD_010 - Match: same company/user during period. "
        "Materials, plants and timestamps provide context; "
        "no same-material or sequence condition applied."
    )
    print(
        "SD_010 - Output includes all supplied positions of selected "
        "material documents and available linked reversal/original documents."
    )
    print(
        "SD_010 - Reversing entries are support, not triggers. "
        "Original posted activity is retained, without netting inventory effects."
    )
    print(
        "SD_010 - Amounts are material-movement local amounts. "
        "Material documents are not required to balance as FI journals."
    )
    print(
        "SD_010 - Runtime input: SD_INVENTORY only. "
        "Movement-type catalog is not loaded."
    )
    print(f"SD_010 - output: {path}")

    return {
        "status": "OK", "output_file": path,
        "sheet_name": "SD10", "rows": len(output),
        "summary": stats,
    }


run = run_sd_010
