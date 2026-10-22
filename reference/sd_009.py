from pathlib import Path
import re
"""SD09 preliminary analysis of supplied sales-condition changes and billing.
Always reports partial coverage. No case numbers are hardcoded.
"""
from pathlib import Path
from datetime import datetime, time, timedelta
import json
import math
import os
import re
import tempfile
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries
from core.sd_common import (
    normalize_text, normalize_identifier, parse_config_date,
    parse_config_companies, parse_sap_date, parse_sap_amount,
    load_sd_ar_journals, write_control_sheet,
)

DOC = ["Company", "Fiscal Year", "Document Number"]
KEY = DOC + ["Document Item"]
EVENT = ["Object Class", "Sales Document", "Change Document"]

VISIBLE = [
    "Company", "Document Number", "Fiscal Year", "Document Item",
    "Trigger", "Billing Document", "Customer Code",
    "Price Changed By", "Invoice Created By",
    "Latest Price Change At", "Invoice Created At", "Sequence Status",
    "GL Account", "Debit Amount", "Credit Amount",
    "Document Currency", "Item Text",
]

EXTRA = [
    "Journal Key", "Posting User", "Entry Date", "Posting Key",
    "Account Type", "Vendor Code", "Debit Credit",
    "Amount in Document Currency", "Amount in Local Currency",
    "Local Currency", "BSEG Document Currency",
    "Journal Source File", "Journal Source Row",
    "FI Header Source File", "FI Header Source Row",
    "FI Reference Type", "FI Reference Key",
    "Matched Sales Items", "Matched Billing Items",
    "Changed Condition Types", "Matched Change Events",
    "Qualifying Activity Links", "Customer Trigger Basis",
    "Coverage Status", "Trigger Basis",
]


def ident(value, width=0, optional=False):
    value = normalize_identifier(value)
    value = value.zfill(width) if value.isdigit() else value.upper()
    return (
        "" if optional and value and set(value) == {"0"}
        else value
    )


def label(value):
    return re.sub(
        r"[^a-z0-9]", "",
        re.sub(r"\.\d+$", "", normalize_text(value)).lower(),
    )


def simple(fields):
    return {
        target: [(source, 1)]
        for target, source in fields.items()
    }


def read(
    context, prefix, mapping, keys, name,
    amounts=(), samples=False,
):
    suffix = parse_config_date(
        context["module"]["to"], "TO"
    ).strftime("%Y%m%d")

    paths = sorted(
        path for path in Path(context["input_folder"]).iterdir()
        if path.is_file()
        and not path.name.startswith("~$")
        and path.suffix.lower() in {".xlsx", ".xlsm"}
        and (
            label(path.stem).startswith(label(prefix))
            and suffix in label(path.stem)
            if samples else
            re.fullmatch(
                re.escape(label(prefix) + suffix) + r"\d*",
                label(path.stem),
            )
        )
    )
    if not paths:
        raise FileNotFoundError(
            "SD09: missing " + prefix + " input"
        )

    frames = []
    removed = 0
    rows_read = 0

    for path in paths:
        print("SD_009 - Reading: " + path.name, flush=True)
        raw = pd.read_excel(
            path, sheet_name="Sheet1", header=None,
            dtype=object, keep_default_na=False,
        )
        if raw.empty:
            raise ValueError("SD09: empty input " + path.name)

        headings = [label(value) for value in raw.iloc[0]]
        body = raw.iloc[1:]
        frame = pd.DataFrame(index=body.index)

        for target, options in mapping.items():
            selected = None
            for source, occurrence in options:
                found = [
                    index for index, heading in enumerate(headings)
                    if heading == label(source)
                ]
                if len(found) >= occurrence:
                    selected = found[occurrence - 1]
                    break

            if selected is None:
                raise ValueError(path.name + ": missing " + target)
            frame[target] = body.iloc[:, selected]

        frame = frame.apply(
            lambda column: column.map(normalize_text)
        )
        blank = frame.eq("").all(axis=1)
        totals = (
            frame[keys].eq("").all(axis=1)
            & frame[
                [column for column in frame if column not in amounts]
            ].eq("").all(axis=1)
        )
        bad = frame[keys].eq("").any(axis=1) & ~blank & ~totals

        if bad.any():
            raise ValueError(
                f"{path.name}: incomplete keys at rows "
                f"{(frame.index[bad] + 1).tolist()[:10]}"
            )

        rows_read += len(frame)
        removed += int((blank | totals).sum())
        frame = frame.loc[~blank & ~totals].copy()
        frame[name + " File"] = path.name
        frame[name + " Row"] = frame.index + 1
        frames.append(frame)

    print(
        f"SD_009 - {name}: {len(paths)} files; "
        f"{rows_read} rows; {removed} blank/total rows removed"
    )
    return pd.concat(frames, ignore_index=True)


def unique(data, keys, name):
    business = [
        column for column in data
        if not column.endswith(" File")
        and not column.endswith(" Row")
    ]
    result = data.drop_duplicates(business).copy()
    print(
        f"SD_009 - {name} identical duplicates removed: "
        f"{len(data) - len(result)}"
    )
    if result.duplicated(keys).any():
        raise ValueError("SD09: contradictory " + name + " keys")
    return result


def dates(data, column):
    data[column] = pd.to_datetime(
        data[column].map(parse_sap_date), errors="coerce"
    )
    if data[column].isna().any():
        raise ValueError("SD09: invalid " + column)


def clock(value):
    if isinstance(value, datetime):
        return value.time()
    if isinstance(value, time):
        return value

    value = re.sub(
        r"\s+", " ", normalize_text(value).lower()
    ).strip()
    value = re.sub(
        r"([ap])\s*\.?\s*m\s*\.?",
        lambda match: match.group(1).upper() + "M",
        value,
    )
    value = re.sub(
        r"\s*(AM|PM)\s*$", r" \1", value
    ).strip()

    for fmt in [
        "%I:%M:%S %p", "%I:%M %p",
        "%H:%M:%S", "%H:%M", "%H:%M:%S.%f",
    ]:
        try:
            return datetime.strptime(value, fmt).time()
        except ValueError:
            pass

    try:
        number = float(value)
        if 0 <= number < 1:
            return (
                datetime(2000, 1, 1)
                + timedelta(seconds=round(number * 86400))
            ).time()
    except ValueError:
        pass

    return None


def stamp(day, value):
    parsed = clock(value)
    if parsed is None:
        return pd.NaT
    return pd.Timestamp(datetime.combine(day.date(), parsed))


def number(value):
    if not normalize_text(value):
        raise ValueError("SD09: missing condition amount")
    result = parse_sap_amount(value)
    if pd.isna(result) or not math.isfinite(result):
        raise ValueError("SD09: invalid condition amount")
    return result


def population(context):
    h = read(
        context, "LBR SD_SALES_CDHDR",
        simple({
            "Object Class": "Change doc. object",
            "Sales Document": "Object value",
            "Change Document": "Document number",
            "Changed By": "User",
            "Change Date": "Date",
            "Change Time": "Time",
            "Transaction": "Transaction Code",
        }),
        EVENT, "Header",
    )

    p = read(
        context, "LBR SD_SALES_CDPOS",
        simple({
            "Object Class": "Change doc. object",
            "Sales Document": "Object value",
            "Change Document": "Document number",
            "Table": "Table Name",
            "Table Key": "Table Key",
            "Field": "Field Name",
            "Change ID": "Change ID",
            "Old Raw": "Old value",
            "New Raw": "New value",
        }),
        EVENT + ["Table", "Table Key", "Field", "Change ID"],
        "CDPOS", samples=True,
    )

    t = read(
        context, "LBR SD_PRICE_TYPES",
        simple({
            "Application": "Application",
            "Condition Type": "Condition type",
            "Condition Class": "Condition class",
        }),
        ["Application", "Condition Type"], "Types",
    )

    columns = simple({
        "Company": "CoCd",
        "Billing Document": "Bill.Doc.",
        "Billing Type": "BillT",
        "Category": "DocCa",
        "Billing User": "Created by",
        "Billing Created": "Created on",
        "Billing Time": "Time",
        "Currency": "Curr.",
        "Payer": "Payer",
        "Sold To": "Sold-to pt",
        "Cancelled": "Can",
        "Cancellation Document": "CancBillDc",
        "Posting Status": "PsSt",
        "Sales Document": "Sales Doc.",
    })
    columns["Sales Item"] = [("Sales Item", 1), ("Item", 1)]
    columns["Billing Item"] = [("Billing Item", 1), ("Item", 2)]

    b = read(
        context, "LBR SD_BILLING", columns,
        ["Company", "Billing Document", "Billing Item"],
        "Billing", amounts={"Currency"},
    )

    for data in [h, p]:
        data["Object Class"] = data["Object Class"].str.upper()
        for column in ["Sales Document", "Change Document"]:
            data[column] = data[column].map(
                lambda value: ident(value, 10)
            )

    h["Changed By"] = h["Changed By"].str.upper()
    for column in ["Table", "Field", "Change ID"]:
        p[column] = p[column].str.upper()
    for column in ["Application", "Condition Type", "Condition Class"]:
        t[column] = t[column].str.upper()

    h = unique(h, EVENT, "CDHDR")
    p = unique(
        p, EVENT + ["Table", "Table Key", "Field", "Change ID"],
        "CDPOS",
    )
    t = unique(t, ["Application", "Condition Type"], "Types")

    for column, width, optional in [
        ("Company", 4, False),
        ("Billing Document", 10, False),
        ("Billing Item", 6, False),
        ("Sales Document", 10, True),
        ("Sales Item", 6, True),
        ("Payer", 10, True),
        ("Sold To", 10, True),
        ("Cancellation Document", 10, True),
    ]:
        b[column] = b[column].map(
            lambda value, w=width, opt=optional:
            ident(value, w, opt)
        )

    for column in [
        "Billing User", "Category", "Cancelled", "Posting Status"
    ]:
        b[column] = b[column].str.upper()

    dates(h, "Change Date")
    dates(b, "Billing Created")
    b = unique(
        b, ["Company", "Billing Document", "Billing Item"], "Billing"
    )

    start = parse_config_date(context["module"]["from"], "FROM")
    end = parse_config_date(context["module"]["to"], "TO")
    companies = parse_config_companies(
        context["module"].get("companies", "")
    )

    h = h.loc[
        h["Object Class"].eq("VERKBELEG")
        & h["Change Date"].between(start, end)
    ].copy()
    b = b.loc[b["Billing Created"].between(start, end)].copy()

    if companies is not None:
        b = b.loc[b["Company"].isin(companies)].copy()

    header_fields = [
        "Company", "Billing Document", "Billing User",
        "Billing Created", "Billing Time", "Currency",
        "Category", "Billing Type", "Cancelled",
        "Cancellation Document", "Posting Status", "Payer", "Sold To",
    ]
    if b[header_fields].drop_duplicates().duplicated(
        ["Company", "Billing Document"]
    ).any():
        raise ValueError("SD09: contradictory billing headers")

    eligible = b.loc[
        b["Category"].eq("M")
        & b["Cancelled"].eq("")
        & b["Cancellation Document"].eq("")
    ].copy()

    if (
        eligible["Billing User"].eq("").any()
        or h["Changed By"].eq("").any()
    ):
        raise ValueError(
            "SD09: missing activity user; cannot silently exclude"
        )

    b["Billing At"] = pd.to_datetime([
        stamp(day, value)
        for day, value in zip(b["Billing Created"], b["Billing Time"])
    ])
    eligible = b.loc[
        b["Category"].eq("M")
        & b["Cancelled"].eq("")
        & b["Cancellation Document"].eq("")
    ].copy()

    h["Change At"] = pd.to_datetime([
        stamp(day, value)
        for day, value in zip(h["Change Date"], h["Change Time"])
    ])

    orders = eligible[[
        "Company", "Billing Document", "Sales Document",
        "Billing User", "Billing Created",
    ]].drop_duplicates()

    possible = h.merge(
        orders,
        left_on=["Sales Document", "Changed By"],
        right_on=["Sales Document", "Billing User"],
        how="inner",
    )
    possible = possible.loc[
        possible["Billing Created"].ge(possible["Change Date"])
    ]

    potential = set(
        possible[EVENT].itertuples(index=False, name=None)
    )
    represented = set(
        p[EVENT].itertuples(index=False, name=None)
    )
    hkeys = set(
        h[EVENT].itertuples(index=False, name=None)
    )

    u = p.loc[
        p["Object Class"].eq("VERKBELEG")
        & p["Table"].eq("KONVC")
        & p["Field"].eq("KBETR")
        & p["Change ID"].eq("U")
    ].copy()

    parsed = u["Table Key"].str.extract(
        r"^(\d{6})(\d{3})(\d{2})([A-Z0-9]{4})$"
    )
    if parsed[0].isna().any():
        raise ValueError(
            "SD09: unrecognized KONVC update key; review format"
        )

    u["Sales Item"] = parsed[0]
    u["Condition Type"] = parsed[3]

    u = u.merge(
        h, on=EVENT, how="left",
        validate="many_to_one", indicator="Header Link",
    )
    if u["Header Link"].ne("both").any():
        raise ValueError(
            "SD09: condition updates without in-scope change header"
        )

    u = u.merge(
        t.loc[
            t["Application"].eq("V"),
            ["Condition Type", "Condition Class"],
        ],
        on="Condition Type", how="left", validate="many_to_one",
    )
    if u["Condition Class"].isna().any():
        raise ValueError(
            "SD09: update condition absent from V catalog"
        )

    u["Old Numeric"] = u["Old Raw"].map(number)
    u["New Numeric"] = u["New Raw"].map(number)
    u = u.loc[
        u["Condition Class"].eq("B")
        & (u["Old Numeric"] - u["New Numeric"]).abs().gt(1e-9)
    ].copy()

    header_condition_count = int(
        u["Sales Item"].eq("000000").sum()
    )
    u = u.loc[u["Sales Item"].ne("000000")].copy()

    links = u.merge(
        eligible,
        left_on=["Sales Document", "Sales Item", "Changed By"],
        right_on=["Sales Document", "Sales Item", "Billing User"],
        how="inner", validate="many_to_many",
    )
    links = links.loc[
        links["Billing Created"].ge(links["Change Date"])
    ].copy()

    known = links["Change At"].notna() & links["Billing At"].notna()
    before = known & links["Billing At"].lt(links["Change At"])
    excluded_before = int(before.sum())
    links = links.loc[~before].copy()

    links["Sequence"] = "AFTER CHANGE DATE; TIME INCOMPLETE"
    same_day = links["Billing Created"].eq(links["Change Date"])
    links.loc[same_day, "Sequence"] = "SAME DAY; TIME UNCONFIRMED"

    known = links["Change At"].notna() & links["Billing At"].notna()
    links.loc[
        known & links["Billing At"].gt(links["Change At"]),
        "Sequence",
    ] = "BILLING AFTER CHANGE"
    links.loc[
        known & links["Billing At"].eq(links["Change At"]),
        "Sequence",
    ] = "SAME TIMESTAMP; ORDER UNCONFIRMED"

    stats = {
        "Change header events": len(h),
        "CDPOS positions after deduplication": len(p),
        "Change events represented in CDPOS": len(represented & hkeys),
        "Potential order/user/date events": len(potential),
        "Potential events with no CDPOS rows":
            len(potential - represented),
        "Class B header-condition updates outside detailed scope":
            header_condition_count,
        "Activity links excluded because billing timestamp precedes change":
            excluded_before,
        "Qualifying activity links": len(links),
        "Preliminary billing documents": len(
            links[["Company", "Billing Document"]].drop_duplicates()
        ),
    }
    return links, stats


def fi_headers(context, links):
    f = read(
        context, "LBR SD_BILLING_HEADERS",
        {
            "Company": [
                ("Company Code", 1), ("CoCd", 1)
            ],
            "Document Number": [
                ("Document Number", 1), ("DocumentNo", 1)
            ],
            "Fiscal Year": [
                ("Fiscal Year", 1), ("Year", 1)
            ],
            "Posting User": [("User name", 1)],
            "Entry Date": [
                ("Entry Date", 1), ("Entry Dte", 1)
            ],
            "Document Currency": [
                ("Currency", 1), ("Crcy", 1)
            ],
            "Local Currency": [("Local Currency", 1)],
            "FI Reference Type": [
                ("Reference Transact.", 1), ("Ref. Tran.", 1)
            ],
            "FI Reference Key": [
                ("Reference Key", 1), ("Ref. Key", 1)
            ],
        },
        DOC, "FI Header",
    )

    for column, width in [
        ("Company", 4), ("Document Number", 10), ("Fiscal Year", 4)
    ]:
        f[column] = f[column].map(
            lambda value, w=width: ident(value, w)
        )

    f["FI Reference Type"] = f["FI Reference Type"].str.upper()
    f["Posting User"] = f["Posting User"].str.upper()
    f = f.loc[f["FI Reference Type"].eq("VBRK")].copy()

    def billing_ref(value):
        value = normalize_identifier(value).strip()
        return (
            value.zfill(10)
            if value.isdigit() and len(value) <= 10
            else value[:10]
        )

    f["Billing Document"] = f["FI Reference Key"].map(billing_ref)
    f = f.merge(
        links[["Company", "Billing Document"]].drop_duplicates(),
        on=["Company", "Billing Document"],
        how="inner", validate="many_to_one",
    )
    f = unique(f, DOC, "FI Header")
    dates(f, "Entry Date")

    required = links[
        ["Company", "Billing Document"]
    ].drop_duplicates()
    coverage = required.merge(
        f[["Company", "Billing Document"]].drop_duplicates(),
        on=["Company", "Billing Document"],
        how="left", indicator=True,
    )
    missing = coverage.loc[
        coverage["_merge"].ne("both"),
        ["Company", "Billing Document"],
    ]
    if not missing.empty:
        raise ValueError(
            "SD09: missing explicit FI reference for "
            + missing.to_string(index=False)
        )

    if f.duplicated(["Company", "Billing Document"]).any():
        raise ValueError(
            "SD09: multiple FI keys per billing; review required"
        )
    if f[["Document Currency", "Local Currency"]].eq("").any(axis=None):
        raise ValueError("SD09: missing FI header currencies")

    compare = links[[
        "Company", "Billing Document", "Billing User", "Currency"
    ]].drop_duplicates().merge(
        f, on=["Company", "Billing Document"],
        validate="one_to_one",
    )
    if compare["Billing User"].ne(compare["Posting User"]).any():
        raise ValueError(
            "SD09: FI posting user differs from billing creator"
        )
    if compare["Currency"].str.upper().ne(
        compare["Document Currency"].str.upper()
    ).any():
        raise ValueError("SD09: billing/FI currency disagreement")

    return f


def activity_json(rows):
    fields = [
        "Sales Document", "Sales Item", "Condition Type", "Table Key",
        "Change Document", "Changed By", "Change Date", "Change Time",
        "Transaction", "Old Raw", "New Raw",
        "Billing Document", "Billing Item", "Billing Created",
        "Billing Time", "Sequence",
        "CDPOS File", "CDPOS Row", "Header File", "Header Row",
        "Billing File", "Billing Row",
    ]

    result = []
    for row in rows.to_dict("records"):
        record = {
            column: (
                row[column].isoformat()
                if isinstance(row[column], (datetime, pd.Timestamp))
                else row[column]
            )
            for column in fields
        }
        result.append(json.dumps(
            record, ensure_ascii=False,
            default=str, separators=(",", ":"),
        ))

    chunks = []
    current = []
    size = 2

    for entry in result:
        if current and size + len(entry) + 1 > 30000:
            chunks.append("[" + ",".join(current) + "]")
            current = []
            size = 2
        current.append(entry)
        size += len(entry) + 1

    if current:
        chunks.append("[" + ",".join(current) + "]")

    return chunks



def supplement_sd09_journals(context, journals, headers):
    """Add complete missing journals from explicitly named AR supplements."""
    present = set(journals[DOC].itertuples(index=False, name=None))
    missing = set(headers[DOC].itertuples(index=False, name=None)) - present
    if not missing:
        return journals

    folder = Path(context["input_folder"])
    files = sorted(
        p for p in folder.iterdir()
        if p.is_file() and re.fullmatch(
            r"LBR SD_JOURNAL_AR_20260731_\d+\.xlsx",
            p.name, flags=re.IGNORECASE,
        )
    )

    expected_headers = [
        "CoCd", "DocumentNo", "Year", "Itm", "AccTy", "PK",
        "G/L", "Vendor", "Customer", "D/C", "Amount", "Crcy",
        "Amount in LC", "Crcy", "Text",
    ]
    records = []

    for source in files:
        print("SD_009 - Supplemental journal:", source.name)
        book = load_workbook(source, read_only=True, data_only=True)
        try:
            rows = book.worksheets[0].iter_rows(values_only=True)
            raw_headers = next(rows, None)
            if raw_headers is None:
                raise ValueError(
                    f"SD09: empty supplemental journal: {source.name}"
                )
            actual = [normalize_text(v) for v in raw_headers]
            while actual and actual[-1] == "":
                actual.pop()
            if actual != expected_headers:
                raise ValueError(
                    "SD09: unexpected supplemental journal columns in "
                    + source.name + ": " + repr(actual)
                )

            for excel_row, raw in enumerate(rows, start=2):
                values = list(raw[:15])
                if len(values) < 15:
                    values.extend([None] * (15 - len(values)))

                # Ignore unrelated documents and SAP total rows.
                try:
                    key = (
                        ident(values[0], 4),
                        ident(values[2], 4),
                        ident(values[1], 10),
                    )
                except (ValueError, TypeError):
                    continue
                if key not in missing:
                    continue

                records.append({
                    "Company": key[0],
                    "Fiscal Year": key[1],
                    "Document Number": key[2],
                    "Document Item": ident(values[3], 3),
                    "Account Type": normalize_text(values[4]).upper(),
                    "Posting Key": ident(values[5], 2),
                    "GL Account": ident(values[6], 10, True),
                    "Vendor Code": ident(values[7], 10, True),
                    "Customer Code": ident(values[8], 10, True),
                    "Debit Credit": normalize_text(values[9]).upper(),
                    "Amount in Document Currency":
                        parse_sap_amount(values[10]),
                    "BSEG Document Currency":
                        normalize_text(values[11]).upper(),
                    "Amount in Local Currency":
                        parse_sap_amount(values[12]),
                    "Local Currency":
                        normalize_text(values[13]).upper(),
                    "Item Text": normalize_text(values[14]),
                    "Journal Source File": source.name,
                    "Journal Source Row": excel_row,
                })
        finally:
            book.close()

    if not records:
        return journals

    extra = pd.DataFrame(records)
    compare_columns = [
        column for column in extra.columns
        if column not in {"Journal Source File", "Journal Source Row"}
    ]
    extra = extra.drop_duplicates(compare_columns)
    if extra.duplicated(KEY).any():
        raise ValueError(
            "SD09: contradictory supplemental journal positions"
        )

    if (
        extra["Account Type"].eq("D")
        & extra["Customer Code"].eq("")
    ).any():
        raise ValueError(
            "SD09: supplemental customer position without customer ID"
        )

    print(
        "SD_009 - Supplemental missing journals:",
        len(extra[DOC].drop_duplicates()),
    )
    print("SD_009 - Supplemental positions:", len(extra))
    return pd.concat([journals, extra], ignore_index=True)


def output(context, links, stats):
    if links.empty:
        return pd.DataFrame(columns=VISIBLE + EXTRA)

    f = fi_headers(context, links)
    journals, _ = load_sd_ar_journals(context)
    journals = journals.copy()
    journals = supplement_sd09_journals(context, journals, f)

    overlaps = {
        column: "Journal " + column
        for column in journals
        if column in f and column not in DOC
    }
    journals = journals.rename(columns=overlaps)

    lines = journals.merge(
        f.rename(columns={
            "FI Header File": "FI Header Source File",
            "FI Header Row": "FI Header Source Row",
        }),
        on=DOC, how="inner", validate="many_to_one",
    )

    required = [
        "Document Item", "Account Type", "GL Account",
        "Customer Code", "Vendor Code", "Posting Key", "Debit Credit",
        "Amount in Document Currency", "Amount in Local Currency",
        "Item Text", "Journal Source File", "Journal Source Row",
    ]
    if any(column not in lines for column in required):
        raise ValueError("SD09: unexpected shared journal schema")

    lines["Document Item"] = lines["Document Item"].map(
        lambda value: ident(value, 3)
    )
    if lines.duplicated(KEY).any():
        raise ValueError("SD09: repeated journal positions")

    for column in ["Customer Code", "Vendor Code"]:
        lines[column] = lines[column].map(
            lambda value: ident(value, 10, True)
        )

    if not lines["Debit Credit"].isin({"S", "H"}).all():
        raise ValueError("SD09: invalid journal D/C")

    for column in [
        "Amount in Document Currency", "Amount in Local Currency"
    ]:
        if not lines[column].map(
            lambda value: pd.notna(value) and math.isfinite(value)
        ).all():
            raise ValueError("SD09: invalid journal amounts")

    for source, target in [
        ("BSEG Document Currency", "Document Currency"),
        ("Journal Document Currency", "Document Currency"),
        ("BSEG Local Currency", "Local Currency"),
        ("Journal Local Currency", "Local Currency"),
    ]:
        if source in lines:
            value = lines[source].map(normalize_text).str.upper()
            if (
                value.ne("")
                & value.ne(lines[target].str.upper())
            ).any():
                raise ValueError(
                    "SD09: source/header currency disagreement"
                )

    seen = set(lines[DOC].itertuples(index=False, name=None))
    expected = set(f[DOC].itertuples(index=False, name=None))
    if seen != expected:
        missing = expected - seen
        print("\nSD09 - ASIENTOS SIN DETALLE EN JOURNAL")
        print("====================================")
        print(pd.DataFrame(
            sorted(missing), columns=DOC
        ).to_string(index=False))

        print("\nClaves faltantes, mostrando formato exacto:")
        for key in sorted(missing):
            print(repr(key))

        def diagnostic_number(value):
            return normalize_text(value).lstrip("0") or "0"

        missing_numbers = {
            diagnostic_number(key[DOC.index("Document Number")])
            for key in missing
        }
        nearby = journals.loc[
            journals["Document Number"].map(
                diagnostic_number
            ).isin(missing_numbers),
            DOC,
        ].drop_duplicates()

        print("\nMismos numeros de documento presentes en Journal:")
        if nearby.empty:
            print("Ninguno.")
        else:
            print(nearby.to_string(index=False))
            print("\nFormato exacto de las claves encontradas:")
            for key in nearby.itertuples(index=False, name=None):
                print(repr(key))

        raise ValueError(
            "SD09: missing supplied journal detail for "
            + str(len(missing))
            + " journal(s). See diagnostic keys above."
        )

    sign = lines["Debit Credit"].map({"S": 1, "H": -1})
    balances = lines[DOC].copy()
    for column in [
        "Amount in Document Currency", "Amount in Local Currency"
    ]:
        balances[column] = lines[column] * sign

    if balances.groupby(DOC).sum().abs().gt(
        .010000001
    ).any(axis=None):
        raise ValueError("SD09: journal does not balance")

    metadata = []
    max_chunks = 0

    for (company, bill), group in links.groupby(
        ["Company", "Billing Document"], sort=True
    ):
        sequence = "BILLING AFTER CHANGE"
        if group["Sequence"].str.contains(
            "UNCONFIRMED|INCOMPLETE"
        ).any():
            sequence = "SEQUENCE REVIEW REQUIRED"

        latest = (
            group["Change At"].max()
            if group["Change At"].notna().all()
            else pd.NaT
        )
        parties = set(group["Payer"]) | set(group["Sold To"])
        parties.discard("")
        if not parties:
            raise ValueError("SD09: missing billing customer evidence")

        record = {
            "Company": company,
            "Billing Document": bill,
            "Price Changed By": group["Changed By"].iloc[0],
            "Invoice Created By": group["Billing User"].iloc[0],
            "Latest Price Change At": latest,
            "Invoice Created At": group["Billing At"].iloc[0],
            "Sequence Status": sequence,
            "Matched Sales Items": "; ".join(sorted(set(
                group["Sales Document"] + "|" + group["Sales Item"]
            ))),
            "Matched Billing Items": "; ".join(sorted(set(
                group["Billing Item"]
            ))),
            "Changed Condition Types": "; ".join(sorted(set(
                group["Condition Type"]
            ))),
            "Matched Change Events":
                len(group[EVENT].drop_duplicates()),
            "Qualifying Activity Links": len(group),
            "Customer Trigger Basis": "; ".join(sorted(parties)),
        }

        chunks = activity_json(group)
        max_chunks = max(max_chunks, len(chunks))
        for index, value in enumerate(chunks, 1):
            record[f"Activity Evidence JSON {index}"] = value

        metadata.append(record)

    lines = lines.merge(
        pd.DataFrame(metadata),
        on=["Company", "Billing Document"],
        how="left", validate="many_to_one",
    )

    customer_match = pd.Series(
        [
            customer in basis.split("; ")
            for customer, basis in zip(
                lines["Customer Code"],
                lines["Customer Trigger Basis"],
            )
        ],
        index=lines.index,
    )
    trigger = (
        lines["Account Type"].eq("D")
        & lines["Debit Credit"].eq("S")
        & customer_match
    )
    customer_trigger = trigger.copy()
    journal_keys = pd.Series(
        list(lines[DOC].itertuples(index=False, name=None)),
        index=lines.index,
    )
    with_customer = set(
        lines.loc[
            lines["Account Type"].eq("D"), DOC
        ].itertuples(index=False, name=None)
    )
    with_customer_trigger = set(
        lines.loc[customer_trigger, DOC].itertuples(
            index=False, name=None
        )
    )
    contradictory_customer = with_customer - with_customer_trigger
    if contradictory_customer:
        raise ValueError(
            "SD09: customer positions exist but no matching customer "
            "debit supports flagged billing: "
            + repr(sorted(contradictory_customer))
        )

    no_customer = expected - with_customer
    journal_review = journal_keys.isin(no_customer)
    trigger = customer_trigger | journal_review
    lines["Trigger"] = trigger.map({True: "YES", False: "NO"})

    stats["Journals without customer positions"] = len(no_customer)
    stats["Journal-level review positions"] = int(journal_review.sum())
    stats["Total trigger positions"] = int(trigger.sum())

    if no_customer:
        print(
            "SD_009 - Journals without customer positions; "
            "all supplied positions marked for journal-level review:"
        )
        print(pd.DataFrame(
            sorted(no_customer), columns=DOC
        ).to_string(index=False))

    lines["Debit Amount"] = lines[
        "Amount in Document Currency"
    ].where(lines["Debit Credit"].eq("S"), 0.)
    lines["Credit Amount"] = lines[
        "Amount in Document Currency"
    ].where(lines["Debit Credit"].eq("H"), 0.)

    lines["Journal Key"] = [
        "|".join(row)
        for row in lines[DOC].itertuples(index=False, name=None)
    ]
    lines["Coverage Status"] = "PRELIMINARY - PARTIAL COVERAGE"
    lines["Trigger Basis"] = [
        (
            "JOURNAL-LEVEL REVIEW: explicit VBRK billing/FI link; "
            "no customer positions in supplied journal. All supplied "
            "journal positions marked; no customer-line match asserted."
        ) if review else (
            "Customer debit position supporting flagged billing; "
            "commercial item numbers are not FI line numbers"
        )
        for review in journal_review
    ]

    stats["Output FI journals"] = len(f)
    stats["Output positions"] = len(lines)
    stats["Customer trigger positions"] = int(customer_trigger.sum())

    result = (
        lines.reindex(columns=VISIBLE + EXTRA + [
            f"Activity Evidence JSON {index}"
            for index in range(1, max_chunks + 1)
        ])
        .sort_values(KEY)
        .reset_index(drop=True)
    )

    for column in result:
        if result[column].map(
            lambda value: len(value) > 32767
            if isinstance(value, str) else False
        ).any():
            raise ValueError("SD09: Excel text limit exceeded")

    return result


def present(path, data, stats):
    book = load_workbook(path)
    temporary = None

    try:
        sheet = book["SD09"]
        header_row = None

        for row in range(1, min(sheet.max_row, 30) + 1):
            if set(VISIBLE).issubset({
                normalize_text(cell.value) for cell in sheet[row]
            }):
                header_row = row
                break

        if header_row is None:
            raise ValueError("SD09: output header not found")

        tables = [
            (table, range_boundaries(table.ref))
            for table in sheet.tables.values()
        ]
        sheet.insert_rows(header_row, 4)
        new_header = header_row + 4

        legends = [
            (
                "SD09 - Preliminary exceptions; Partial coverage"
                if not data.empty else
                "SD09 - No identified exceptions in supplied detail; "
                "Partial coverage"
            ),
            (
                f"Detected {stats['Preliminary billing documents']:,} "
                f"billing documents; output "
                f"{stats.get('Output FI journals', 0):,} FI journals "
                f"and {len(data):,} positions."
            ),
            (
                f"Supplied CDPOS represents "
                f"{stats['Change events represented in CDPOS']:,} "
                f"of {stats['Change header events']:,} header events. "
                f"{stats['Potential events with no CDPOS rows']:,} "
                "potential order/user/date events have no downloaded detail."
            ),
            (
                "Scope: supplied KBETR/U class B item changes and "
                "noncancelled M billings; same order/item/user. "
                "Equal or missing timestamps require review. "
                "Insertions/deletions, header conditions and "
                "master-data changes are not resolved."
            ),
        ]

        for offset, value in enumerate(legends):
            row = header_row + offset
            sheet.merge_cells(
                start_row=row, start_column=1,
                end_row=row, end_column=len(VISIBLE),
            )
            cell = sheet.cell(row, 1, value)
            cell.alignment = Alignment(
                wrap_text=True, vertical="top"
            )
            cell.font = Font(bold=offset == 0)
            sheet.row_dimensions[row].height = (
                32 if offset < 3 else 45
            )

        sheet.row_dimensions[new_header].height = 30

        if data.empty:
            sheet.auto_filter.ref = None
        else:
            sheet.auto_filter.ref = (
                f"A{new_header}:"
                f"{get_column_letter(len(data.columns))}"
                f"{new_header + len(data)}"
            )

        for table, (left, top, right, bottom) in tables:
            table.ref = (
                f"{get_column_letter(left)}{top + 4}:"
                f"{get_column_letter(right)}{bottom + 4}"
            )
            if table.autoFilter is not None:
                table.autoFilter.ref = table.ref

        sheet.freeze_panes = f"F{new_header + 1}"
        sheet.column_dimensions.group(
            get_column_letter(len(VISIBLE) + 1),
            get_column_letter(len(data.columns)),
            outline_level=1, hidden=True,
        )
        sheet.sheet_properties.outlinePr.summaryRight = False
        sheet.sheet_view.showOutlineSymbols = True
        sheet.sheet_format.outlineLevelCol = 1
        sheet.column_dimensions[
            get_column_letter(len(VISIBLE))
        ].collapsed = True

        for name in [
            "Latest Price Change At", "Invoice Created At"
        ]:
            index = list(data.columns).index(name) + 1
            for row in range(
                new_header + 1, new_header + len(data) + 1
            ):
                sheet.cell(row, index).number_format = (
                    "dd/mm/yyyy hh:mm:ss"
                )

        sheet.column_dimensions["J"].width = 23
        sheet.column_dimensions["K"].width = 23
        sheet.column_dimensions["L"].width = 31

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



def _sd09_notes_at_end(path):
    from pathlib import Path
    from copy import copy
    import os
    import tempfile
    from openpyxl import load_workbook
    from openpyxl.utils.cell import (
        range_boundaries, get_column_letter, coordinate_to_tuple
    )

    book = load_workbook(path)
    temporary = None
    try:
        sheet = book["SD09"]
        required = {"Company", "Billing Document", "Trigger"}
        header_row = None

        for row in range(1, min(sheet.max_row, 30) + 1):
            values = {cell.value for cell in sheet[row]}
            if required.issubset(values):
                header_row = row
                break

        if header_row == 1:
            return

        if (
            header_row != 5
            or not str(sheet.cell(1, 1).value or "").startswith("SD09 - ")
        ):
            raise ValueError(
                "SD09: unexpected layout. No presentation changes saved."
            )

        note_count = header_row - 1
        notes = []
        heights = []

        for row in range(1, note_count + 1):
            heights.append(sheet.row_dimensions[row].height)
            notes.append([
                (
                    sheet.cell(row, column).value,
                    copy(sheet.cell(row, column)._style),
                )
                for column in range(1, 18)
            ])

        merges = []
        for merged in list(sheet.merged_cells.ranges):
            bounds = (
                merged.min_col, merged.min_row,
                merged.max_col, merged.max_row,
            )
            if merged.min_row <= note_count < merged.max_row:
                raise ValueError("SD09: unexpected merged range.")
            merges.append(bounds)
            sheet.unmerge_cells(str(merged))

        def shift_reference(reference):
            left, top, right, bottom = range_boundaries(reference)
            if top <= note_count:
                raise ValueError(
                    "SD09: unexpected table/filter reference."
                )
            return (
                f"{get_column_letter(left)}{top - note_count}:"
                f"{get_column_letter(right)}{bottom - note_count}"
            )

        filter_ref = sheet.auto_filter.ref
        shifted_filter = (
            shift_reference(filter_ref) if filter_ref else None
        )

        tables = []
        for table in sheet.tables.values():
            tables.append((
                table,
                shift_reference(table.ref),
                shift_reference(table.autoFilter.ref)
                if table.autoFilter is not None
                and table.autoFilter.ref else None,
            ))

        freeze = sheet.freeze_panes
        if freeze:
            freeze_row, freeze_col = coordinate_to_tuple(str(freeze))
            shifted_freeze = (
                f"{get_column_letter(freeze_col)}"
                f"{max(1, freeze_row - note_count)}"
            )
        else:
            shifted_freeze = None

        last_data_row = max(
            row for row in range(header_row + 1, sheet.max_row + 1)
            if sheet.cell(row, 1).value not in (None, "")
        ) if sheet.max_row > header_row else header_row

        sheet.delete_rows(1, note_count)

        # Row dimensions do not move automatically with delete_rows.
        for row in range(1, note_count + 1):
            sheet.row_dimensions[row].height = None

        sheet.auto_filter.ref = shifted_filter
        sheet.freeze_panes = "A2"

        for table, table_ref, table_filter in tables:
            table.ref = table_ref
            if table.autoFilter is not None:
                table.autoFilter.ref = table_filter

        footer_start = last_data_row - note_count + 3

        for offset, cells in enumerate(notes):
            target_row = footer_start + offset
            for column, (value, style) in enumerate(cells, 1):
                cell = sheet.cell(target_row, column, value)
                cell._style = style
            sheet.row_dimensions[target_row].height = heights[offset]

        for left, top, right, bottom in merges:
            if bottom <= note_count:
                new_top = footer_start + top - 1
                new_bottom = footer_start + bottom - 1
            else:
                new_top = top - note_count
                new_bottom = bottom - note_count

            sheet.merge_cells(
                start_row=new_top, start_column=left,
                end_row=new_bottom, end_column=right,
            )

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


_sd09_present_before_footer = present


def present(path, data, stats):
    _sd09_present_before_footer(path, data, stats)
    _sd09_notes_at_end(path)


def run_sd_009(context):
    control = context.get("control", {})
    if (
        normalize_text(control.get("param1"))
        or normalize_text(control.get("param2"))
    ):
        raise ValueError(
            "SD09 preliminary version: leave PARAM1 and PARAM2 blank"
        )

    links, stats = population(context)
    data = output(context, links, stats)

    for name, value in stats.items():
        print(f"SD_009 - {name}: {value}")

    print(
        "SD_009 - Result classification: "
        "PRELIMINARY / PARTIAL COVERAGE"
    )

    path = write_control_sheet(
        context, "SD09", data,
        date_columns=[
            "Latest Price Change At", "Invoice Created At", "Entry Date"
        ],
        amount_columns=[
            "Debit Amount", "Credit Amount",
            "Amount in Document Currency", "Amount in Local Currency",
        ],
    )
    present(path, data, stats)

    print(
        "SD_009 - Customer debit positions are trigger support; "
        "commercial pricing items are not equated to FI line items."
    )
    print(
        "SD_009 - Same timestamp or unrecognized time is retained "
        "for sequence review. FI posting time does not replace "
        "billing creation time."
    )
    print(
        "SD_009 - Supplied journals reconciled and DC/LC balances "
        "checked. SAP extraction completeness and use of the changed "
        "price in billing are not independently certified."
    )
    print(
        "SD_009 - Runtime inputs: SD_SALES_CDHDR, all supplied "
        "SD_SALES_CDPOS samples/parts, SD_PRICE_TYPES, SD_BILLING; "
        "SD_BILLING_HEADERS and SD_JOURNAL_AR when cases exist."
    )
    print(f"SD_009 - output: {path}")

    return {
        "status": "OK",
        "result_status": "PRELIMINARY_PARTIAL",
        "output_file": path,
        "sheet_name": "SD09",
        "rows": len(data),
        "summary": stats,
    }


run = run_sd_009
