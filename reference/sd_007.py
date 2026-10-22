"""SD07: same user changes a vendor bank account and posts its payment.

Scope: same-event LFBK replacements/direct BANKN updates; AP 25/S.
Both activities in period; same-day sequence remains unconfirmed.
"""
from pathlib import Path
import json
import os
import re
import tempfile

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

from core.sd_common import (
    load_ap_items,
    load_sd_journals,
    normalize_text,
    normalize_identifier,
    parse_config_date,
    parse_sap_date,
    write_control_sheet,
)

DOC = ["Company", "Fiscal Year", "Document Number"]
KEY = DOC + ["Document Item"]

VISIBLE = [
    "Company", "Document Number", "Fiscal Year", "Document Item",
    "Trigger", "Vendor Code", "Bank Changed By", "Bank Change Date",
    "Old Bank Account", "New Bank Account", "Payment Posted By",
    "Entry Date", "GL Account", "Debit Amount", "Credit Amount",
    "Document Currency", "Item Text",
]
EXTRA = [
    "Journal Key", "Matched Change Events", "Sequence Basis",
    "Bank Change Evidence", "Document Type", "Transaction Code",
    "Account Type", "Posting Key", "Debit Credit",
    "Amount in Document Currency", "Amount in Local Currency",
    "Local Currency", "BSEG Document Currency",
    "AP Source File", "AP Source Row",
    "Journal Source File", "Journal Source Row",
]


def ident(value, width):
    value = normalize_identifier(value)
    return value.zfill(width) if value.isdigit() else value.upper()


def read_changes(context, kind, mapping):
    suffix = parse_config_date(
        context["module"]["to"], "TO"
    ).strftime("%Y%m%d")
    wanted = f"lbrsdbank{kind.lower()}{suffix}"

    files = [
        p for p in Path(context["input_folder"]).iterdir()
        if p.is_file()
        and not p.name.startswith("~$")
        and p.suffix.lower() in {".xlsx", ".xlsm"}
        and re.sub(r"[^a-z0-9]", "", p.stem.lower()) == wanted
    ]
    if len(files) != 1:
        raise ValueError(
            f"SD07: expected one SD_BANK_{kind} file; "
            f"found {len(files)}"
        )

    raw = pd.read_excel(
        files[0], sheet_name="Sheet1", header=None,
        dtype=object, keep_default_na=False,
    )
    if raw.empty:
        raise ValueError("SD07: empty change extract")

    headings = [
        re.sub(r"\.\d+$", "", normalize_text(v)).casefold()
        for v in raw.iloc[0]
    ]
    body = raw.iloc[1:]
    result = pd.DataFrame(index=body.index)

    for target, source in mapping.items():
        hits = [
            i for i, heading in enumerate(headings)
            if heading == source.casefold()
        ]
        if not hits:
            raise ValueError(
                f"SD07 {files[0].name}: missing {source}"
            )
        result[target] = body.iloc[:, hits[0]]

    result = result.loc[
        ~result.apply(
            lambda column: column.map(normalize_text)
        ).eq("").all(axis=1)
    ].copy()

    result["Source Row"] = result.index + 1
    result["Source File"] = files[0].name

    for column in mapping:
        if column != "Change Date":
            result[column] = result[column].map(normalize_text)

    for column in ["Class", "Vendor Code", "Change Document"]:
        if result[column].eq("").any():
            raise ValueError("SD07: incomplete change key")

    result["Class"] = result["Class"].str.upper()
    result["Vendor Code"] = result["Vendor Code"].map(
        lambda value: ident(value, 10)
    )
    result["Change Document"] = result["Change Document"].map(
        lambda value: ident(value, 10)
    )
    return result


def decode(key, vendor):
    # LFBK: MANDT(3), LIFNR(10), BANKS(3), BANKL(15), BANKN(18).
    if (
        not 32 <= len(key) <= 49
        or not key[:3].isdigit()
        or key[3:13] != vendor
    ):
        raise ValueError(
            "SD07: unsupported LFBK key or vendor mismatch"
        )

    key = key.ljust(49)
    country = key[13:16].strip()
    bank = key[16:31].strip()
    account = key[31:49].strip()

    if (
        not re.fullmatch(r"[A-Za-z]{2,3}", country)
        or not bank
        or not account
    ):
        raise ValueError("SD07: cannot interpret LFBK bank key")

    return country, bank, account


def bank_changes(context):
    hdr = read_changes(context, "CDHDR", {
        "Class": "Change doc. object",
        "Vendor Code": "Object value",
        "Change Document": "Document number",
        "Bank Changed By": "User",
        "Change Date": "Date",
        "Change Time": "Time",
        "Change TCode": "Transaction Code",
    })
    pos = read_changes(context, "CDPOS", {
        "Class": "Change doc. object",
        "Vendor Code": "Object value",
        "Change Document": "Document number",
        "Table": "Table Name",
        "Table Key": "Table Key",
        "Field": "Field Name",
        "Indicator": "Change ID",
        "New Value": "New value",
        "Old Value": "Old value",
    })
    event = ["Class", "Vendor Code", "Change Document"]

    hdr["Change Date"] = pd.to_datetime(
        hdr["Change Date"].map(parse_sap_date)
    )
    if hdr["Change Date"].isna().any():
        raise ValueError("SD07: invalid change dates")

    hdr["Bank Changed By"] = hdr["Bank Changed By"].str.upper()
    header_fields = event + [
        "Bank Changed By", "Change Date",
        "Change Time", "Change TCode",
    ]
    if (
        hdr[header_fields].drop_duplicates()
        .duplicated(event).any()
    ):
        raise ValueError("SD07: contradictory change headers")

    hdr = hdr.drop_duplicates(event).rename(columns={
        "Source File": "Header Source File",
        "Source Row": "Header Source Row",
    })

    for column in ["Table", "Field", "Indicator"]:
        pos[column] = pos[column].str.upper()

    physical = event + ["Table", "Table Key", "Field"]
    business = physical + ["Indicator", "New Value", "Old Value"]
    if (
        pos[business].drop_duplicates()
        .duplicated(physical).any()
    ):
        raise ValueError("SD07: contradictory change positions")

    pos = pos.drop_duplicates(business)
    data = pos.merge(
        hdr, on=event, how="left",
        validate="many_to_one", indicator=True,
    )
    if data["_merge"].ne("both").any():
        raise ValueError("SD07: change positions without headers")

    start = parse_config_date(context["module"]["from"], "FROM")
    end = parse_config_date(context["module"]["to"], "TO")
    data = data.loc[
        data["Class"].eq("KRED")
        & data["Table"].eq("LFBK")
        & data["Change Date"].between(start, end)
    ].copy()

    stats = {
        "Bank change positions": len(data),
        "Paired replacement events": 0,
        "Account number change events": 0,
        "Bank/country-only events": 0,
        "Unpaired insertion events excluded": 0,
        "Unpaired deletion events excluded": 0,
    }
    records = []

    for _, group in data.groupby(event, sort=False):
        added = set(group.loc[
            group["Indicator"].eq("I")
            & group["Field"].eq("KEY"),
            "Table Key",
        ])
        deleted = set(group.loc[
            group["Indicator"].isin({"D", "E"}),
            "Table Key",
        ])
        account_changes = []

        if added and deleted:
            stats["Paired replacement events"] += 1
            if len(added) != 1 or len(deleted) != 1:
                raise ValueError(
                    "SD07: ambiguous multiple-record bank replacement"
                )

            old_key = next(iter(deleted))
            new_key = next(iter(added))
            vendor = group.iloc[0]["Vendor Code"]
            old = decode(old_key, vendor)
            new = decode(new_key, vendor)

            if old[2] != new[2]:
                account_changes.append(
                    (old[2], new[2], old_key, new_key)
                )
            elif old[:2] != new[:2]:
                stats["Bank/country-only events"] += 1
        else:
            stats["Unpaired insertion events excluded"] += int(
                bool(added)
            )
            stats["Unpaired deletion events excluded"] += int(
                bool(deleted)
            )

        direct_updates = group.loc[
            group["Indicator"].eq("U")
            & group["Field"].eq("BANKN")
        ]
        for _, row in direct_updates.iterrows():
            if not row["Old Value"] or not row["New Value"]:
                raise ValueError(
                    "SD07: BANKN update without old/new values"
                )
            if row["Old Value"] != row["New Value"]:
                account_changes.append((
                    row["Old Value"], row["New Value"],
                    row["Table Key"], row["Table Key"],
                ))

        if len(account_changes) > 1:
            raise ValueError(
                "SD07: multiple account changes in one event need review"
            )

        if account_changes:
            first = group.iloc[0]
            old, new, old_key, new_key = account_changes[0]

            if not first["Bank Changed By"]:
                raise ValueError(
                    "SD07: qualifying bank change without user"
                )

            sources = "; ".join(
                f"{row['Source File']}:row {int(row['Source Row'])}"
                for _, row in group.iterrows()
            )
            records.append({
                "Vendor Code": first["Vendor Code"],
                "Bank Changed By": first["Bank Changed By"],
                "Bank Change Date": first["Change Date"],
                "Old Bank Account": old,
                "New Bank Account": new,
                "Change Document": first["Change Document"],
                "Change Time": first["Change Time"],
                "Change TCode": first["Change TCode"],
                "Old Record Key": old_key,
                "New Record Key": new_key,
                "Header Source File": first["Header Source File"],
                "Header Source Row": int(first["Header Source Row"]),
                "Position Sources": sources,
            })

    stats["Account number change events"] = len(records)
    columns = [
        "Vendor Code", "Bank Changed By", "Bank Change Date",
        "Old Bank Account", "New Bank Account", "Change Document",
        "Change Time", "Change TCode", "Old Record Key",
        "New Record Key", "Header Source File",
        "Header Source Row", "Position Sources",
    ]
    return pd.DataFrame(records, columns=columns), stats


def select_payments(ap, changes):
    ap = ap.copy()

    for column, width in [
        ("Company", 4), ("Fiscal Year", 4),
        ("Document Number", 10), ("Document Item", 3),
        ("Vendor Code", 10), ("Posting Key", 2),
        ("Reversal Document", 10), ("Reversal Year", 4),
    ]:
        ap[column] = ap[column].map(
            lambda value, width=width: ident(value, width)
        )

    for column in ["User ID", "Transaction Code", "Debit Credit"]:
        ap[column] = ap[column].map(normalize_text).str.upper()

    for column in ["Reversal Document", "Reversal Year"]:
        ap.loc[ap[column].str.fullmatch(r"0+"), column] = ""

    reversed_docs = set()
    for _, row in ap.loc[
        ap["Reversal Document"].ne("")
    ].iterrows():
        reversed_docs.add(tuple(row[column] for column in DOC))
        if row["Reversal Year"]:
            reversed_docs.add((
                row["Company"], row["Reversal Year"],
                row["Reversal Document"],
            ))

    reversal = pd.Series(
        [
            key in reversed_docs
            for key in ap[DOC].itertuples(index=False, name=None)
        ],
        index=ap.index, dtype=bool,
    )
    reversal |= ap["Transaction Code"].isin({
        "FB08", "F.80", "MR8M"
    })

    base = (
        ap["Posting Key"].eq("25")
        & ap["Debit Credit"].eq("S")
    )
    payments = ap.loc[
        base & ~reversal
        & ap["Amount in Document Currency"].gt(0)
    ].copy()

    if payments["User ID"].eq("").any():
        raise ValueError(
            "SD07: eligible payments without posting user"
        )
    if payments.duplicated(KEY).any():
        raise ValueError("SD07: duplicate payment position")

    pairs = payments.merge(
        changes,
        left_on=["Vendor Code", "User ID"],
        right_on=["Vendor Code", "Bank Changed By"],
        how="inner", validate="many_to_many",
    )
    linked = pairs.loc[
        pairs["Entry Date"] >= pairs["Bank Change Date"]
    ].copy()

    same_day = linked.loc[
        linked["Entry Date"].eq(linked["Bank Change Date"])
    ]

    stats = {
        "AP positions": len(ap),
        "Payment positions before exclusions": int(base.sum()),
        "Payment positions excluded as reversals":
            int((base & reversal).sum()),
        "Eligible payment positions": len(payments),
        "Eligible payment journals":
            len(payments[DOC].drop_duplicates()),
        "Matching payment journals before date check":
            len(pairs[DOC].drop_duplicates()),
        "Exception journals": len(linked[DOC].drop_duplicates()),
        "Exception trigger positions":
            len(linked[KEY].drop_duplicates()),
        "Same-day linked payment positions":
            len(same_day[KEY].drop_duplicates()),
    }
    return ap, linked, stats


def full_detail(context, ap, linked):
    documents = linked[DOC].drop_duplicates()
    required = ap.merge(
        documents, on=DOC, how="inner",
        validate="many_to_one",
    )
    fields = DOC + [
        "User ID", "Entry Date", "Document Currency",
        "Document Type", "Transaction Code",
    ]
    headers = required[fields].drop_duplicates()
    if headers.duplicated(DOC).any():
        raise ValueError("SD07: conflicting payment headers")

    journals, _ = load_sd_journals(context)
    if journals.duplicated(KEY).any():
        raise ValueError("SD07: duplicate journal position")

    detail = journals.merge(
        headers, on=DOC, how="inner",
        validate="many_to_one",
    )
    comparison = required.merge(
        detail, on=KEY, how="left",
        suffixes=("_AP", "_BSEG"),
        indicator=True, validate="many_to_one",
    )

    bad = comparison["_merge"].ne("both")
    for field in ["Vendor Code", "Posting Key", "Debit Credit"]:
        bad |= comparison[field + "_AP"].ne(
            comparison[field + "_BSEG"]
        )
    for field in [
        "Amount in Document Currency",
        "Amount in Local Currency",
    ]:
        bad |= (
            comparison[field + "_AP"]
            - comparison[field + "_BSEG"]
        ).abs().gt(.01)

    if bad.any():
        raise ValueError(
            "SD07: missing/contradictory AP journal detail; "
            "no partial result published"
        )

    sign = detail["Debit Credit"].map({"S": 1, "H": -1})
    if sign.isna().any():
        raise ValueError("SD07: invalid journal debit/credit")

    balances = detail[DOC].copy()
    for field in [
        "Amount in Document Currency",
        "Amount in Local Currency",
    ]:
        if detail[field].isna().any():
            raise ValueError("SD07: blank journal amounts")
        balances[field] = detail[field] * sign

    if balances.groupby(DOC).sum().abs().gt(.01).any(axis=None):
        raise ValueError(
            "SD07: unbalanced journal detail; "
            "no partial result published"
        )

    if "BSEG Document Currency" in detail:
        disagreement = (
            detail["BSEG Document Currency"].ne("")
            & detail["BSEG Document Currency"].ne(
                detail["Document Currency"]
            )
        )
        if disagreement.any():
            raise ValueError(
                "SD07: document currency disagreement"
            )

    evidence_fields = [
        "Change Document", "Bank Change Date", "Change Time",
        "Bank Changed By", "Old Bank Account", "New Bank Account",
        "Change TCode", "Old Record Key", "New Record Key",
        "Header Source File", "Header Source Row", "Position Sources",
    ]
    rows = []

    for _, group in linked.groupby(KEY, sort=False):
        group = group.sort_values([
            "Bank Change Date", "Change Time", "Change Document",
        ])
        last = group.iloc[-1]
        row = {
            column: last[column]
            for column in KEY + [
                "Bank Changed By", "Bank Change Date",
                "Old Bank Account", "New Bank Account",
            ]
        }
        row["Matched Change Events"] = len(group)
        row["Sequence Basis"] = (
            "SAME_DAY_SEQUENCE_UNCONFIRMED"
            if last["Entry Date"] == last["Bank Change Date"]
            else "PAYMENT_DATE_AFTER_CHANGE"
        )

        evidence = group[evidence_fields].copy()
        evidence["Bank Change Date"] = (
            evidence["Bank Change Date"].dt.strftime("%Y-%m-%d")
        )
        row["Bank Change Evidence"] = json.dumps(
            evidence.to_dict("records"),
            ensure_ascii=False, default=str,
        )
        if len(row["Bank Change Evidence"]) > 32000:
            raise ValueError(
                "SD07: change evidence exceeds Excel cell limit"
            )

        row["AP Source File"] = last["Source File"]
        row["AP Source Row"] = last["Source Row"]
        row["Trigger"] = "YES"
        rows.append(row)

    detail = detail.merge(
        pd.DataFrame(rows), on=KEY, how="left",
        validate="one_to_one",
    )
    detail["Trigger"] = detail["Trigger"].fillna("NO")
    detail = detail.rename(columns={
        "User ID": "Payment Posted By",
    })
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
    return (
        detail.reindex(columns=VISIBLE + EXTRA)
        .sort_values(KEY).reset_index(drop=True)
    )


def present(path, output, stats):
    book = load_workbook(path)
    temporary = None

    try:
        sheet = book["SD07"]

        if output.empty:
            message = (
                f"Evaluated {stats['Account number change events']:,} "
                "bank account changes and "
                f"{stats['Eligible payment journals']:,} "
                "payment journals. "
                "No qualifying change/payment pairs were found."
            )
            texts = {
                3: "SD07 - No exceptions",
                4: message,
                5: (
                    "Scope: account-number replacements/direct updates; "
                    "vendor payments 25/S; both activities in "
                    "configured period."
                ),
                6: (
                    "Same vendor and posting user; payment date "
                    "on/after change date. "
                    "Same-day sequence is unconfirmed."
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
                get_column_letter(len(VISIBLE + EXTRA)),
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


def run_sd_007(context):
    control = context.get("control", {})
    if any(
        normalize_text(control.get(parameter))
        for parameter in ["param1", "param2"]
    ):
        raise ValueError(
            "SD07: leave PARAM1 and PARAM2 empty; scope is fixed "
            "to account changes and 25/S payments"
        )

    changes, stats = bank_changes(context)
    ap, _ = load_ap_items(context)
    ap, linked, payment_stats = select_payments(ap, changes)
    stats.update(payment_stats)

    for name, value in stats.items():
        print(f"SD_007 - {name}: {value}")

    output = (
        full_detail(context, ap, linked)
        if not linked.empty
        else pd.DataFrame(columns=VISIBLE + EXTRA)
    )

    path = write_control_sheet(
        context, "SD07", output,
        date_columns=["Bank Change Date", "Entry Date"],
        amount_columns=[
            "Debit Amount", "Credit Amount",
            "Amount in Document Currency",
            "Amount in Local Currency",
        ],
    )
    present(path, output, stats)

    print(
        "SD_007 - Scope: same-period same-user/vendor bank-account "
        "changes and payment postings; "
        "initial/unpaired additions excluded."
    )
    print(
        "SD_007 - Payment bank account used and bank "
        "execution/approval are not established by these inputs."
    )
    print(
        "SD_007 - Same-day matches require sequence review. "
        "Journal balance does not independently certify "
        "SAP completeness."
    )
    print(f"SD_007 - Output positions: {len(output)}")
    print(f"SD_007 - output: {path}")

    return {
        "status": "OK",
        "output_file": path,
        "sheet_name": "SD07",
        "rows": len(output),
        "summary": stats,
    }


run = run_sd_007
