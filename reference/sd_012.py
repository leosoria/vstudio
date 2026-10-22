"""SD12: same user posts a receipt and its associated supplier invoice.

Independent inputs: SD_INVOICES, SD_RECEIPTS, SD_HEADERS, SD_JOURNAL.
Exact receipt references and PO-item associations are labelled separately.
"""
from pathlib import Path
import json
import math
import os
import re
import tempfile

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

from core.sd_common import (
    load_sd_journals, normalize_text, normalize_identifier,
    parse_config_date, parse_config_companies, parse_sap_date,
    parse_sap_amount, write_control_sheet,
)

DOC = ["Company", "Fiscal Year", "Document Number"]
KEY = DOC + ["Document Item"]
MM_DOC = ["Company", "MM Year", "MM Document"]
MM_KEY = MM_DOC + ["MM Item"]
PO = ["Company", "Order Number", "Order Item"]
GR_KEY = PO + [
    "Receipt Document", "Receipt Year", "Receipt Item"
]

VISIBLE = [
    "Company", "Document Number", "Fiscal Year", "Document Item",
    "Trigger", "Vendor Code", "Invoice Posted By",
    "Receipt Posted By", "Entry Date", "MM Invoice Number",
    "Receipt Documents", "GL Account", "Debit Amount",
    "Credit Amount", "Document Currency", "Item Text", "Link Basis",
]
EXTRA = [
    "Journal Key", "Order Items", "Matched Invoice Items",
    "Matched Receipt Positions", "Receipt Entry Dates",
    "Sequence Basis", "MM Users", "Document Type",
    "Transaction Code", "Document Date", "Posting Date",
    "Reference", "Reference Key", "Account Type", "Posting Key",
    "Debit Credit", "Amount in Document Currency",
    "Amount in Local Currency", "Local Currency",
    "BSEG Document Currency", "Header Source File",
    "Header Source Row", "Journal Source File", "Journal Source Row",
]


def ident(value, width):
    value = normalize_identifier(value)
    return value.zfill(width) if value.isdigit() else value.upper()


def optional(value, width):
    value = ident(value, width)
    return "" if value and set(value) == {"0"} else value


def label(value):
    value = re.sub(r"\.\d+$", "", normalize_text(value))
    return re.sub(r"[^a-z0-9]", "", value.lower())


def read_input(context, kind, mapping, keys, allowed):
    suffix = parse_config_date(
        context["module"]["to"], "TO"
    ).strftime("%Y%m%d")
    expected = label("LBR SD " + kind + suffix)

    files = sorted(
        p for p in Path(context["input_folder"]).iterdir()
        if p.is_file()
        and not p.name.startswith("~$")
        and p.suffix.lower() in {".xlsx", ".xlsm"}
        and re.fullmatch(
            re.escape(expected) + r"\d*", label(p.stem)
        )
    )
    if not files:
        raise FileNotFoundError(
            f"SD12: missing LBR SD_{kind}_{suffix}.xlsx"
        )

    frames = []
    removed = 0

    for path in files:
        raw = pd.read_excel(
            path, sheet_name="Sheet1", header=None,
            dtype=object, keep_default_na=False,
        )
        if raw.empty:
            raise ValueError("SD12: empty input " + path.name)

        headings = [label(v) for v in raw.iloc[0]]
        body = raw.iloc[1:]
        data = pd.DataFrame(index=body.index)

        for target, (source, occurrence) in mapping.items():
            found = [
                i for i, heading in enumerate(headings)
                if heading == label(source)
            ]
            if len(found) < occurrence:
                raise ValueError(
                    f"SD12 {path.name}: missing {source}, "
                    f"occurrence {occurrence}"
                )
            data[target] = body.iloc[:, found[occurrence - 1]]

        texts = data.apply(
            lambda column: column.map(normalize_text)
        )
        blank = texts.eq("").all(axis=1)
        totals = (
            texts[keys].eq("").all(axis=1)
            & texts[
                [c for c in texts if c not in allowed]
            ].eq("").all(axis=1)
        )
        bad = (
            texts[keys].eq("").any(axis=1)
            & ~blank & ~totals
        )
        if bad.any():
            rows = (data.index[bad] + 1).tolist()[:10]
            raise ValueError(
                f"SD12 {path.name}: incomplete keys at rows {rows}"
            )

        removed += int((totals & ~blank).sum())
        data = data.loc[~blank & ~totals].copy()
        data["Source File"] = path.name
        data["Source Row"] = data.index + 1
        frames.append(data)

    result = pd.concat(frames, ignore_index=True)
    result["Company"] = result["Company"].map(
        lambda value: ident(value, 4)
    )
    companies = parse_config_companies(
        context["module"].get("companies", "")
    )
    if companies is not None:
        result = result.loc[
            result["Company"].isin(companies)
        ].copy()

    print(
        f"SD_012 - {kind}: {len(result)} supplied rows; "
        f"totals removed: {removed}",
        flush=True,
    )
    return result


def unique(frame, key, name):
    business = [
        c for c in frame
        if c not in {"Source File", "Source Row"}
    ]
    distinct = frame.drop_duplicates(business)
    if distinct.duplicated(key).any():
        raise ValueError(
            "SD12: contradictory duplicate " + name
        )
    return frame.drop_duplicates(key).copy()


def reverse_keys(frame, doc, number, year):
    result = set()
    for row in frame.loc[
        frame[number].ne("")
    ].to_dict("records"):
        result.add(tuple(row[c] for c in doc))
        if row[year]:
            result.add((
                row["Company"], row[year], row[number]
            ))
    return result


def population(context):
    inv = read_input(context, "INVOICES", {
        "Company": ("CoCd", 1),
        "MM Document": ("Doc. No.", 1),
        "MM Year": ("Year", 1),
        "MM Item": ("InvItem", 1),
        "Invoice Vendor": ("Inv. Pty", 1),
        "MM Type": ("Doc. Type", 1),
        "MM Entry Date": ("Entry Dte", 1),
        "MM User": ("User name", 1),
        "Invoice Currency": ("Crcy", 1),
        "Invoice Indicator": ("I", 1),
        "Status": ("St", 1),
        "MM Reversal": ("Rvrsd by", 1),
        "MM Reversal Year": ("Year", 2),
        "Order Number": ("Purch.Doc.", 1),
        "Order Item": ("Item", 1),
        "Invoice DC": ("D/C", 1),
        "Invoice Amount": ("Amount", 1),
        "Invoice Quantity": ("Quantity", 1),
        "Invoice UOM": ("OUn", 1),
        "Ref Document": ("Ref. Doc.", 1),
        "Ref Year": ("Year", 3),
        "Ref Item": ("RfIt", 1),
    }, MM_KEY, {
        "Invoice Currency", "Invoice Amount",
        "Invoice Quantity", "Invoice UOM",
    })

    gr = read_input(context, "RECEIPTS", {
        "Company": ("CoCd", 1),
        "Order Number": ("Purch.Doc.", 1),
        "Order Item": ("Item", 1),
        "Receipt Document": ("Mat. Doc.", 1),
        "Receipt Year": ("MatYr", 1),
        "Receipt Item": ("Item", 2),
        "Receipt Entry Date": ("Entry Dte", 1),
        "Receipt User": ("User name", 1),
        "History Event": ("Tr./ev.type", 1),
        "Movement Type": ("MvT", 1),
        "Receipt DC": ("D/C", 1),
        "Receipt Quantity": ("Quantity", 1),
        "Receipt Amount": ("Amount", 1),
        "Receipt Currency": ("Crcy", 1),
        "Receipt UOM": ("OUn", 1),
    }, GR_KEY, {
        "Receipt Quantity", "Receipt Amount",
        "Receipt Currency", "Receipt UOM",
    })

    fi = read_input(context, "HEADERS", {
        "Company": ("CoCd", 1),
        "Document Number": ("DocumentNo", 1),
        "Fiscal Year": ("Year", 1),
        "Document Type": ("Doc. Type", 1),
        "Document Date": ("Doc. Date", 1),
        "Posting Date": ("Pstng Date", 1),
        "Entry Date": ("Entry Dte", 1),
        "Entry Time": ("Time", 1),
        "Invoice Posted By": ("User name", 1),
        "Transaction Code": ("TCode", 1),
        "Document Currency": ("Crcy", 1),
        "Reference": ("Reference", 1),
        "Reference Transaction": ("Ref. Tran.", 1),
        "Reference Key": ("Ref. Key", 1),
        "Reversal Document": ("Rev. with", 1),
        "Reversal Year": ("Year", 2),
    }, DOC, {"Document Currency"})

    for frame in (inv, gr):
        for column, width in [
            ("Order Number", 10), ("Order Item", 5)
        ]:
            frame[column] = frame[column].map(
                lambda value, width=width: optional(value, width)
            )

    widths_by_frame = [
        (inv, [
            ("MM Document", 10), ("MM Year", 4),
            ("MM Item", 6), ("Invoice Vendor", 10),
        ]),
        (gr, [
            ("Receipt Document", 10), ("Receipt Year", 4),
            ("Receipt Item", 4),
        ]),
        (fi, [
            ("Document Number", 10), ("Fiscal Year", 4),
        ]),
    ]
    for frame, widths in widths_by_frame:
        for column, width in widths:
            frame[column] = frame[column].map(
                lambda value, width=width: ident(value, width)
            )

    optional_widths = [
        (inv, [
            ("Ref Document", 10), ("Ref Year", 4), ("Ref Item", 4),
            ("MM Reversal", 10), ("MM Reversal Year", 4),
        ]),
        (fi, [
            ("Reversal Document", 10), ("Reversal Year", 4),
        ]),
    ]
    for frame, widths in optional_widths:
        for column, width in widths:
            frame[column] = frame[column].map(
                lambda value, width=width: optional(value, width)
            )

    for frame, columns in [
        (inv, ["MM Entry Date"]),
        (gr, ["Receipt Entry Date"]),
        (fi, ["Entry Date", "Document Date", "Posting Date"]),
    ]:
        for column in columns:
            frame[column] = _sd12_parse_dates(frame, column)
            if frame[column].isna().any():
                raise ValueError(
                    "SD12: missing/invalid " + column
                )

    for frame, columns in [
        (inv, [
            "MM User", "MM Type", "Invoice DC",
            "Invoice Indicator", "Invoice Currency",
        ]),
        (gr, [
            "Receipt User", "Receipt DC",
            "Movement Type", "Receipt Currency",
        ]),
        (fi, [
            "Invoice Posted By", "Transaction Code",
            "Document Type", "Document Currency",
            "Reference Transaction",
        ]),
    ]:
        for column in columns:
            frame[column] = (
                frame[column].map(normalize_text).str.upper()
            )

    inv["Invoice UOM"] = inv["Invoice UOM"].map(normalize_text)
    gr["Receipt UOM"] = gr["Receipt UOM"].map(normalize_text)
    for column in ["Entry Time", "Reference"]:
        fi[column] = fi[column].map(normalize_text)

    fi["Reference Key"] = fi["Reference Key"].map(
        normalize_identifier
    )
    inv["Status"] = inv["Status"].map(
        lambda value: ident(value, 1)
    )
    gr["History Event"] = gr["History Event"].map(
        lambda value: ident(value, 1)
    )

    for frame, columns in [
        (inv, ["Invoice Amount", "Invoice Quantity"]),
        (gr, ["Receipt Amount", "Receipt Quantity"]),
    ]:
        for column in columns:
            frame[column] = frame[column].map(parse_sap_amount)
            valid = frame[column].map(
                lambda value: pd.notna(value)
                and math.isfinite(value)
            )
            if not valid.all():
                raise ValueError(
                    "SD12: invalid amount/quantity " + column
                )

    inv = unique(inv, MM_KEY, "MM item")
    gr = unique(gr, GR_KEY, "receipt")
    fi = unique(fi, DOC, "FI header")

    mm_reversed = reverse_keys(
        inv, MM_DOC, "MM Reversal", "MM Reversal Year"
    )
    fi_reversed = reverse_keys(
        fi, DOC, "Reversal Document", "Reversal Year"
    )

    start = parse_config_date(
        context["module"]["from"], "FROM"
    )
    end = parse_config_date(
        context["module"]["to"], "TO"
    )
    inv = inv.loc[
        inv["MM Entry Date"].between(start, end)
    ].copy()
    gr = gr.loc[
        gr["Receipt Entry Date"].between(start, end)
    ].copy()

    mm_base = (
        inv["Status"].eq("5")
        & inv["Invoice Indicator"].eq("X")
        & inv["Invoice DC"].eq("S")
        & inv["Invoice Amount"].gt(0)
    )
    not_reversed = pd.Series(
        [
            key not in mm_reversed
            for key in inv[MM_DOC].itertuples(
                index=False, name=None
            )
        ],
        index=inv.index, dtype=bool,
    )
    inv = inv.loc[
        mm_base
        & inv["Order Number"].ne("")
        & inv["Order Item"].ne("")
        & not_reversed
    ].copy()

    gr = gr.loc[
        gr["History Event"].eq("1")
        & gr["Receipt DC"].eq("S")
        & gr["Receipt Quantity"].gt(0)
    ].copy()

    fi = fi.loc[
        fi["Reference Transaction"].eq("RMRP")
    ].copy()
    if not fi["Reference Key"].str.fullmatch(r"\d{14}").all():
        raise ValueError(
            "SD12: unsupported RMRP reference key"
        )
    if fi.duplicated(["Company", "Reference Key"]).any():
        raise ValueError("SD12: ambiguous MM/FI link")

    fi["FI Excluded"] = [
        key in fi_reversed
        for key in fi[DOC].itertuples(index=False, name=None)
    ]
    fi["FI Excluded"] |= (
        fi["Transaction Code"].isin({"MR8M", "FB08", "F.80"})
        | ~fi["Entry Date"].between(start, end)
    )

    inv["MM Key"] = inv["MM Document"] + inv["MM Year"]
    inv = inv.rename(columns={
        "Source File": "Invoice Source File",
        "Source Row": "Invoice Source Row",
    })
    fi = fi.rename(columns={
        "Source File": "Header Source File",
        "Source Row": "Header Source Row",
    })
    gr = gr.rename(columns={
        "Source File": "Receipt Source File",
        "Source Row": "Receipt Source Row",
    })

    inv = inv.merge(
        fi,
        left_on=["Company", "MM Key"],
        right_on=["Company", "Reference Key"],
        how="left", validate="many_to_one",
    )

    stats = {
        "Eligible MM invoice documents":
            len(inv[MM_DOC].drop_duplicates()),
        "MM invoice documents without FI link":
            len(inv.loc[
                inv["Document Number"].isna(), MM_DOC
            ].drop_duplicates()),
        "Eligible receipt positions": len(gr),
        "Receipt positions without user":
            int(gr["Receipt User"].eq("").sum()),
    }
    return inv, gr, fi, stats


def match_activities(inv, gr, stats):
    if stats["MM invoice documents without FI link"]:
        raise ValueError(
            "SD12: eligible MM invoices lack FI headers; "
            "no partial result published"
        )

    refs = inv[
        ["Ref Document", "Ref Year", "Ref Item"]
    ].ne("")
    explicit = refs.all(axis=1)
    no_ref = ~refs.any(axis=1)
    stats["Invoice positions with incomplete reference"] = int(
        (~explicit & ~no_ref).sum()
    )

    direct = inv.loc[explicit].merge(
        gr,
        left_on=PO + ["Ref Document", "Ref Year", "Ref Item"],
        right_on=PO + [
            "Receipt Document", "Receipt Year", "Receipt Item"
        ],
        how="inner", validate="many_to_many",
    )
    direct["Link Basis"] = "EXACT_REFERENCE"

    fallback = inv.loc[no_ref].merge(
        gr, on=PO, how="inner", validate="many_to_many"
    )
    fallback["Link Basis"] = "PO_ITEM_ONLY"

    linked = pd.concat(
        [direct, fallback], ignore_index=True
    )
    stats["Invoice positions with supplied receipt"] = len(
        linked[MM_KEY].drop_duplicates()
    )

    mm_match = (
        linked["MM User"].ne("")
        & linked["Receipt User"].eq(linked["MM User"])
        & linked["Receipt Entry Date"].le(
            linked["MM Entry Date"]
        )
    )
    pending = linked.loc[
        mm_match & linked["Document Number"].isna()
    ]
    if not pending.empty:
        raise ValueError(
            "SD12: MM candidate invoices lack FI headers; "
            "no partial result published"
        )

    known = linked["Document Number"].notna()
    if linked.loc[known, "Invoice Posted By"].eq("").any():
        raise ValueError(
            "SD12: linked FI invoice without posting user"
        )

    matches = linked.loc[
        known
        & linked["FI Excluded"].eq(False)
        & linked["Receipt User"].ne("")
        & linked["Receipt User"].eq(
            linked["Invoice Posted By"]
        )
        & linked["Receipt Entry Date"].le(
            linked["Entry Date"]
        )
    ].copy()

    matches = matches.drop_duplicates(
        MM_KEY + [
            "Receipt Document", "Receipt Year",
            "Receipt Item", "Link Basis",
        ]
    )
    if matches["Invoice Currency"].ne(
        matches["Document Currency"]
    ).any():
        raise ValueError("SD12: MM/FI currencies disagree")

    stats["Exception journals"] = len(
        matches[DOC].drop_duplicates()
    )
    stats["Matched invoice positions"] = len(
        matches[MM_KEY].drop_duplicates()
    )
    stats["Matched receipt positions"] = len(
        matches[GR_KEY].drop_duplicates()
    )
    stats["Exact-reference journals"] = len(
        matches.loc[
            matches["Link Basis"].eq("EXACT_REFERENCE"), DOC
        ].drop_duplicates()
    )
    stats["PO-item-only journals"] = len(
        matches.loc[
            matches["Link Basis"].eq("PO_ITEM_ONLY"), DOC
        ].drop_duplicates()
    )
    stats["Same-day matched invoice positions"] = len(
        matches.loc[
            matches["Receipt Entry Date"].eq(
                matches["Entry Date"]
            ), MM_KEY
        ].drop_duplicates()
    )
    return matches


def evidence_chunks(group):
    fields = MM_KEY + [
        "Invoice Vendor", "MM User", "MM Entry Date",
        "Invoice Amount", "Invoice Currency",
        "Invoice Quantity", "Invoice UOM",
        "Order Number", "Order Item", "Receipt Document",
        "Receipt Year", "Receipt Item", "Receipt User",
        "Receipt Entry Date", "Receipt Quantity", "Receipt UOM",
        "Receipt Amount", "Receipt Currency",
        "Movement Type", "Link Basis",
        "Invoice Source File", "Invoice Source Row",
        "Receipt Source File", "Receipt Source Row",
    ]
    records = group[fields].copy()
    for column in ["MM Entry Date", "Receipt Entry Date"]:
        records[column] = records[column].dt.strftime(
            "%Y-%m-%d"
        )

    chunks = []
    current = []
    for record in records.to_dict("records"):
        value = json.dumps(
            record, ensure_ascii=False, default=str
        )
        if len(value) > 29000:
            raise ValueError(
                "SD12: oversized activity evidence record"
            )
        proposed = "[" + ",".join(current + [value]) + "]"
        if current and len(proposed) > 30000:
            chunks.append("[" + ",".join(current) + "]")
            current = []
        current.append(value)

    if current:
        chunks.append("[" + ",".join(current) + "]")
    return chunks


def build_output(context, matches, headers, stats):
    if matches.empty:
        stats["Output positions"] = 0
        return pd.DataFrame(columns=VISIBLE + EXTRA)

    selected = matches[DOC].drop_duplicates()
    journal, _ = load_sd_journals(context)

    for column, width in [
        ("Company", 4), ("Fiscal Year", 4),
        ("Document Number", 10), ("Document Item", 3),
        ("GL Account", 10), ("Vendor Code", 10),
    ]:
        journal[column] = journal[column].map(
            lambda value, width=width: ident(value, width)
        )

    detail = journal.merge(
        selected, on=DOC, how="inner",
        validate="many_to_one",
    )
    if detail.duplicated(KEY).any():
        raise ValueError("SD12: duplicate journal positions")

    present = set(
        detail[DOC].itertuples(index=False, name=None)
    )
    needed = set(
        selected.itertuples(index=False, name=None)
    )
    if needed - present:
        raise ValueError(
            f"SD12: {len(needed - present)} candidate journals "
            "lack SD journal detail; no partial result published"
        )

    amount_fields = [
        "Amount in Document Currency",
        "Amount in Local Currency",
    ]
    for field in amount_fields:
        valid = detail[field].map(
            lambda value: pd.notna(value)
            and math.isfinite(value)
        )
        if not valid.all():
            raise ValueError("SD12: invalid journal amounts")

    if detail["GL Account"].eq("").any():
        raise ValueError("SD12: missing journal GL account")

    sign = detail["Debit Credit"].map({"S": 1, "H": -1})
    if sign.isna().any():
        raise ValueError("SD12: invalid journal D/C")

    balances = detail[DOC].copy()
    for field in amount_fields:
        balances[field] = detail[field] * sign

    balance = balances.groupby(DOC)[amount_fields].sum()
    counts = detail.groupby(DOC).size()
    if (
        balance.abs().gt(.01).any(axis=None)
        or counts.lt(2).any()
    ):
        raise ValueError(
            "SD12: unbalanced/inconsistent journals; "
            "no partial result published"
        )

    head = headers.merge(
        selected, on=DOC, how="inner",
        validate="one_to_one",
    )
    head_fields = DOC + [
        "Invoice Posted By", "Entry Date",
        "Document Currency", "Document Type",
        "Transaction Code", "Document Date", "Posting Date",
        "Reference", "Reference Key",
        "Header Source File", "Header Source Row",
    ]

    overlaps = (
        set(head_fields) & set(detail.columns)
    ) - set(DOC)
    for column in overlaps:
        if column == "Document Currency":
            nonblank = detail[column].map(
                normalize_text
            ).ne("")
            check = detail.loc[
                nonblank, DOC + [column]
            ].merge(
                head[DOC + [column]], on=DOC,
                suffixes=("_J", "_H"),
            )
            if check[column + "_J"].ne(
                check[column + "_H"]
            ).any():
                raise ValueError(
                    "SD12: journal/header currency disagreement"
                )
        detail = detail.drop(columns=column)

    detail = detail.merge(
        head[head_fields], on=DOC, how="inner",
        validate="many_to_one",
    )
    if "BSEG Document Currency" in detail:
        bad = (
            detail["BSEG Document Currency"].ne("")
            & detail["BSEG Document Currency"].ne(
                detail["Document Currency"]
            )
        )
        if bad.any():
            raise ValueError(
                "SD12: BSEG/header currencies disagree"
            )

    metadata = []
    expected_vendors = {}
    max_chunks = 0

    for document, group in matches.groupby(DOC, sort=False):
        expected_vendors[document] = set(
            group["Invoice Vendor"]
        )
        row = dict(zip(DOC, document))
        row.update({
            "Receipt Posted By": ", ".join(
                sorted(set(group["Receipt User"]))
            ),
            "MM Invoice Number": ", ".join(
                sorted(set(
                    group["MM Document"] + "/" + group["MM Year"]
                ))
            ),
            "Receipt Documents": ", ".join(
                sorted(set(
                    group["Receipt Document"]
                    + "/" + group["Receipt Year"]
                ))
            ),
            "Link Basis": "; ".join(
                sorted(set(group["Link Basis"]))
            ),
            "Order Items": "; ".join(
                sorted(set(
                    group["Order Number"]
                    + "/" + group["Order Item"]
                ))
            ),
            "Matched Invoice Items":
                len(group[MM_KEY].drop_duplicates()),
            "Matched Receipt Positions":
                len(group[GR_KEY].drop_duplicates()),
            "Receipt Entry Dates": ", ".join(
                sorted(set(
                    group["Receipt Entry Date"]
                    .dt.strftime("%d/%m/%Y")
                ))
            ),
            "Sequence Basis": (
                "INCLUDES_SAME_DAY_SEQUENCE_UNCONFIRMED"
                if group["Receipt Entry Date"].eq(
                    group["Entry Date"]
                ).any()
                else "RECEIPT_DATE_BEFORE_INVOICE"
            ),
            "MM Users": ", ".join(
                sorted(set(group["MM User"]))
            ),
        })

        chunks = evidence_chunks(group)
        max_chunks = max(max_chunks, len(chunks))
        for index, chunk in enumerate(chunks, 1):
            row[f"Match Evidence {index}"] = chunk

        metadata.append(row)

    detail = detail.merge(
        pd.DataFrame(metadata), on=DOC, how="left",
        validate="many_to_one",
    )

    triggers = []
    represented = set()

    for row in detail.to_dict("records"):
        document = tuple(row[c] for c in DOC)
        trigger = (
            row["Vendor Code"] in expected_vendors[document]
            and row["Account Type"] == "K"
            and row["Debit Credit"] == "H"
            and row["Amount in Document Currency"] > 0
        )
        triggers.append("YES" if trigger else "NO")
        if trigger:
            represented.add((
                document, row["Vendor Code"]
            ))

    for document, vendors in expected_vendors.items():
        if any(
            (document, vendor) not in represented
            for vendor in vendors
        ):
            raise ValueError(
                "SD12: matched invoice vendor lacks supporting "
                "K/H journal position; no partial result published"
            )

    detail["Trigger"] = triggers
    detail["Journal Key"] = detail[DOC].agg("|".join, axis=1)
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

    columns = VISIBLE + EXTRA + [
        f"Match Evidence {i}"
        for i in range(1, max_chunks + 1)
    ]
    output = (
        detail.reindex(columns=columns)
        .sort_values(KEY).reset_index(drop=True)
    )

    for column in output:
        too_long = output[column].map(
            lambda value:
                len(value) > 32767
                if isinstance(value, str) else False
        )
        if too_long.any():
            raise ValueError(
                "SD12: Excel text limit exceeded in " + column
            )

    stats["Trigger positions"] = triggers.count("YES")
    stats["Output positions"] = len(output)
    return output


def present(path, output, stats):
    book = load_workbook(path)
    temporary = None

    try:
        sheet = book["SD12"]

        if output.empty:
            texts = {
                3: "SD12 - No exceptions",
                4: (
                    f"Evaluated {stats['Eligible MM invoice documents']:,} "
                    "MM invoices and "
                    f"{stats['Eligible receipt positions']:,} "
                    "receipt positions. No qualifying "
                    "receipt/invoice user matches were found."
                ),
                5: (
                    "Scope: supplied VGABE=1 debit receipts and "
                    "posted invoice items; both entry dates in "
                    "period; exact references or PO-item association."
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


def run_sd_012(context):
    if any(
        normalize_text(context.get("control", {}).get(parameter))
        for parameter in ["param1", "param2"]
    ):
        raise ValueError(
            "SD12: leave PARAM1 and PARAM2 empty"
        )

    inv, receipts, headers, stats = population(context)
    matches = match_activities(inv, receipts, stats)
    output = build_output(context, matches, headers, stats)

    for name, value in stats.items():
        print(f"SD_012 - {name}: {value}")

    path = write_control_sheet(
        context, "SD12", output,
        date_columns=[
            "Entry Date", "Document Date", "Posting Date"
        ],
        amount_columns=[
            "Debit Amount", "Credit Amount",
            "Amount in Document Currency",
            "Amount in Local Currency",
        ],
    )
    present(path, output, stats)

    print(
        "SD_012 - Trigger marks supporting vendor K/H positions "
        "in flagged journals; MM item numbers are not FI line numbers."
    )
    print(
        "SD_012 - PO_ITEM_ONLY does not assert a specific receipt "
        "settlement. Same-day matches do not establish time sequence."
    )
    print(
        "SD_012 - Scope: supplied VGABE=1 receipts; original "
        "receipt activity is not netted against reversals; "
        "comprehensive service coverage is not demonstrated."
    )
    print(
        "SD_012 - FI detail keys, vendor evidence, currencies and "
        "DC/LC balances checked; SAP-wide completeness is not "
        "independently certified."
    )
    print(f"SD_012 - output: {path}")

    return {
        "status": "OK",
        "output_file": path,
        "sheet_name": "SD12",
        "rows": len(output),
        "summary": stats,
    }


run = run_sd_012


def _sd12_parse_dates(frame, column):
    from datetime import datetime, date, timedelta
    from numbers import Real

    support = column in {"Document Date", "Posting Date"}
    values = []
    unusual = []

    for index, raw in frame[column].items():
        row = frame.loc[index]
        location = (
            f"{row.get('Source File', 'unknown input')}, "
            f"Excel row {row.get('Source Row', '?')}, "
            f"column {column}"
        )

        try:
            if raw is None or pd.isna(raw) or not normalize_text(raw):
                raise ValueError("blank date")

            if isinstance(raw, pd.Timestamp):
                parsed = raw.to_pydatetime()
            elif isinstance(raw, datetime):
                parsed = raw
            elif isinstance(raw, date):
                parsed = datetime.combine(
                    raw, datetime.min.time()
                )
            else:
                text = normalize_text(raw)
                parsed = None
                formats = (
                    ("%Y%m%d",)
                    if re.fullmatch(r"\d{8}", text)
                    else ("%d/%m/%Y", "%d.%m.%Y", "%d-%m-%Y")
                )

                for fmt in formats:
                    try:
                        parsed = datetime.strptime(text, fmt)
                        break
                    except ValueError:
                        pass

                if (
                    parsed is None
                    and re.match(r"^\d{4}-\d{2}-\d{2}", text)
                ):
                    parsed = datetime.fromisoformat(text)

                if (
                    parsed is None
                    and isinstance(raw, Real)
                    and 0 < raw < 2958466
                ):
                    parsed = (
                        datetime(1899, 12, 30)
                        + timedelta(days=float(raw))
                    )

                if parsed is None:
                    fallback = parse_sap_date(raw)
                    if pd.isna(fallback):
                        raise ValueError("unrecognized date")
                    parsed = pd.Timestamp(
                        fallback
                    ).to_pydatetime()

            parsed = parsed.replace(
                hour=0, minute=0, second=0,
                microsecond=0, tzinfo=None,
            )

            try:
                # Comprueba explicitamente el rango de nanosegundos.
                pd.Timestamp(parsed).value
            except (OverflowError, ValueError):
                if not support:
                    raise ValueError(
                        "entry date outside pandas nanosecond range; "
                        "scope cannot be determined safely"
                    )
                unusual.append((
                    location, parsed.strftime("%Y-%m-%d")
                ))

            values.append(parsed)

        except Exception as error:
            raise ValueError(
                f"SD12: invalid date at {location}; "
                f"raw={raw!r}; {error}"
            ) from error

    if unusual:
        print(
            f"SD_012 - {column}: {len(unusual)} date(s) "
            "outside pandas nanosecond range retained "
            "unchanged as support dates.",
            flush=True,
        )
        for location, value in unusual[:5]:
            print(
                f"SD_012 - Date notice: {location}: {value}",
                flush=True,
            )

    if support:
        # Conserva fechas Python sin forzarlas a datetime64[ns].
        return pd.Series(
            values, index=frame.index, dtype=object
        )

    return pd.Series(
        pd.to_datetime(values), index=frame.index
    )
