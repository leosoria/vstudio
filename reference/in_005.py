"""IN05: period receipts and issues recorded by the same SAP user.

PARAM1: receipts;issues;reversal:original pairs. Empty uses documented defaults.
Source: SD_INVENTORY directly, independently of SD or other IN results.
Scope is entry date; originals are retained even when reversal references exist.
"""
from collections import Counter, defaultdict
from datetime import date, datetime
from decimal import Decimal
import json
import math
import os
from pathlib import Path
import re
import tempfile
from time import perf_counter
from xml.etree import ElementTree as ET
import zipfile

from openpyxl import Workbook, load_workbook
from openpyxl.utils.datetime import from_excel
from core.in_common import (
    REL, PKG, CT, _tag, _styles, _xml_preserve_namespaces, write_sheet,
)

CONTROL_ID = "IN_005"
DEFAULT_PARAM1 = "101;601;102:101,602:601"
SHEETS = ("IN05", "IN05_Detail")
SUMMARY = [
    "Company", "Result", "Receipt Records", "Issue Records", "Cases",
    "Trigger Records", "Reversal Reference Gaps", "From", "To",
]
VISIBLE = [
    "Case", "Company", "Entry Date", "User", "Material", "Plant",
    "Role", "Trigger", "Movement", "Quantity", "Unit",
    "Material Document", "Year", "Item", "Reversal Evidence",
    "Match Pattern",
]
TECH = [
    "Case Key", "Case Shared Entry Dates", "Case User", "Case Material", "Case Plant",
    "Storage Location", "Batch", "Posting Date", "Entry Time", "D/C",
    "Amount in LC", "Currency", "PO", "PO Item", "Vendor", "Reason", "Text",
    "Original Material Document", "Original Year", "Original Item",
    "Mapped Reversal Positions", "Source File", "Source Sheet", "Source Row",
    "All Source Origins", "Record Key", "Coverage",
]


def txt(value):
    return "" if value is None else str(value).strip()


def identifier(value, width=0, optional=False):
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if not math.isfinite(value) or int(value) != value:
            raise ValueError(f"Invalid identifier: {value!r}")
        value = str(int(value))
    result = txt(value)
    if optional and (not result or set(result) == {"0"}):
        return ""
    return result.zfill(width) if result.isdigit() else result


def material(value):
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if math.isfinite(value) and int(value) == value:
            return str(int(value))
    return "" if value is None else str(value)


def number(value):
    if value is None or isinstance(value, bool) or not txt(value):
        raise ValueError(f"Missing/invalid number: {value!r}")
    result = Decimal(str(value))
    if not result.is_finite():
        raise ValueError(f"Non-finite number: {value!r}")
    return result


def day(value, epoch=None):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool) and epoch is not None:
        converted = from_excel(value, epoch)
        if isinstance(converted, datetime):
            return converted.date()
        raise ValueError(f"Invalid Excel date: {value!r}")
    for pattern in ("%Y-%m-%d", "%d/%m/%Y", "%d.%m.%Y", "%Y%m%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(txt(value), pattern).date()
        except ValueError:
            pass
    raise ValueError(f"Missing/invalid date: {value!r}")


def parameters(raw):
    blank = raw is None or (isinstance(raw, float) and math.isnan(raw)) or not txt(raw)
    effective = DEFAULT_PARAM1 if blank else txt(raw).upper()
    parts = [p.strip() for p in effective.split(";")]
    if len(parts) != 3:
        raise ValueError("IN05 PARAM1: receipts;issues;reversal:original. Example: " + DEFAULT_PARAM1)

    def codes(part):
        values = [v.strip() for v in part.split(",")]
        if not values or any(not re.fullmatch(r"[A-Z0-9]{3}", v) for v in values):
            raise ValueError("IN05 PARAM1: movement codes must have three letters/digits; no empty codes.")
        if len(values) != len(set(values)):
            raise ValueError("IN05 PARAM1: duplicate movement codes.")
        return set(values)

    receipts, issues = codes(parts[0]), codes(parts[1])
    if receipts & issues:
        raise ValueError("IN05 PARAM1: receipt/issue sets overlap: " + str(sorted(receipts & issues)))
    reversals = {}
    if parts[2] != "NONE":
        for pair in parts[2].split(","):
            match = re.fullmatch(r"\s*([A-Z0-9]{3})\s*:\s*([A-Z0-9]{3})\s*", pair)
            if not match:
                raise ValueError("IN05 PARAM1: invalid reversal:original pair: " + repr(pair))
            reversal, original = match.groups()
            if reversal in reversals or reversal in receipts | issues or original not in receipts | issues:
                raise ValueError("IN05 PARAM1: repeated/overlapping reversal or original outside configured movements: " + pair)
            reversals[reversal] = original
    canonical = ",".join(sorted(receipts)) + ";" + ",".join(sorted(issues)) + ";" + (
        ",".join(k + ":" + v for k, v in sorted(reversals.items())) if reversals else "NONE"
    )
    return receipts, issues, reversals, canonical, blank


def read_sources(folder, cutoff, scope, metrics):
    paths = sorted(
        p for p in Path(folder).rglob("*")
        if p.is_file() and not p.name.startswith("~$")
        and p.suffix.lower() in {".xlsx", ".xlsm"}
        and re.fullmatch(
            "LBRSDINVENTORY" + cutoff.strftime("%Y%m%d") + r"(?:PART\d+|\d+)?",
            re.sub(r"[ _-]+", "", p.stem).upper(),
        )
    )
    if not paths:
        raise FileNotFoundError(f"IN05 missing SD_INVENTORY for cutoff {cutoff:%Y%m%d}; folder={folder}")
    records = {}
    unique = [
        "CoCd", "Doc. Type", "Pstng Date", "Entry Dte", "Time", "User name",
        "MKPF-TCODE", "MvT", "Material", "Plnt", "SLoc", "Batch", "D/C",
        "Quantity", "Amount in LC", "Text", "PO", "Vendor", "Reas.",
    ]
    for path in paths:
        print("IN05 reading:", path, flush=True)
        book = load_workbook(path, read_only=True, data_only=True)
        try:
            rows = book["Sheet1"].iter_rows(values_only=True)
            headers = [txt(v) for v in next(rows, ())]
            columns = defaultdict(list)
            for i, heading in enumerate(headers):
                columns[heading].append(i)
            wrong = [h for h in unique if len(columns[h]) != 1]
            wrong += [h for h, n in (("Mat. Doc.", 2), ("MatYr", 2), ("Item", 3)) if len(columns[h]) != n]
            wrong += [h for h in ("BUn", "Crcy") if not columns[h]]
            if wrong:
                raise ValueError(f"{path}: missing/ambiguous layout columns: {wrong}; headers={headers}")

            def get(row, heading, occurrence=0):
                index = columns[heading][occurrence]
                return row[index] if index < len(row) else None

            for rownum, values in enumerate(rows, 2):
                metrics["source rows read"] += 1
                origin = f"{path.relative_to(Path(folder))}|Sheet1|{rownum}"
                try:
                    if not any(txt(v) for v in values):
                        metrics["blank rows"] += 1
                        continue
                    identity = ("Mat. Doc.", "MatYr", "Item", "MvT", "Material", "User name", "Plnt")
                    if all(not txt(get(values, h)) for h in identity) and all(
                        not txt(v) or (i < len(headers) and headers[i] in {"CoCd", "Quantity", "Amount in LC", "BUn", "Crcy"})
                        for i, v in enumerate(values)
                    ):
                        metrics["export totals"] += 1
                        continue
                    company = identifier(get(values, "CoCd"), 4)
                    document = identifier(get(values, "Mat. Doc."), 10)
                    year = identifier(get(values, "MatYr"), 4)
                    item = identifier(get(values, "Item"), 4)
                    movement = identifier(get(values, "MvT")).upper()
                    if not all((company, document, year, item, movement)) or any(
                        not v.isdigit() or int(v) <= 0 for v in (document, year, item)
                    ):
                        raise ValueError("Incomplete/invalid material document key or movement.")
                    if scope is not None and company not in scope:
                        metrics["outside configured companies"] += 1
                        continue
                    unit_values = {txt(get(values, "BUn", i)) for i in range(len(columns["BUn"]))} - {""}
                    currency_values = {txt(get(values, "Crcy", i)) for i in range(len(columns["Crcy"]))} - {""}
                    if len(unit_values) > 1 or len(currency_values) > 1:
                        raise ValueError("Contradictory unit/currency occurrences.")
                    ref = (
                        identifier(get(values, "Mat. Doc.", 1), 10, True),
                        identifier(get(values, "MatYr", 1), 4, True),
                        identifier(get(values, "Item", 2), 4, True),
                    )
                    if any(ref) and not all(ref):
                        raise ValueError(f"Incomplete original material document reference: {ref}")
                    data = {
                        "company": company, "document": document, "year": year,
                        "item": item, "movement": movement,
                        "material": material(get(values, "Material")),
                        "plant": txt(get(values, "Plnt")), "storage": txt(get(values, "SLoc")),
                        "batch": txt(get(values, "Batch")), "user": txt(get(values, "User name")),
                        "entry": day(get(values, "Entry Dte"), book.epoch).isoformat(),
                        "posting": day(get(values, "Pstng Date"), book.epoch).isoformat(),
                        "time": txt(get(values, "Time")), "quantity": number(get(values, "Quantity")),
                        "unit": next(iter(unit_values), ""), "amount": number(get(values, "Amount in LC")),
                        "currency": next(iter(currency_values), ""), "dc": txt(get(values, "D/C")),
                        "po": identifier(get(values, "PO"), 10, True),
                        "po_item": identifier(get(values, "Item", 1), 5, True),
                        "vendor": identifier(get(values, "Vendor"), 10, True),
                        "reason": txt(get(values, "Reas.")), "text": txt(get(values, "Text")),
                        "type": txt(get(values, "Doc. Type")), "tcode": txt(get(values, "MKPF-TCODE")),
                        "ref": ref,
                    }
                    key = (company, year, document, item)
                    if key in records:
                        if records[key]["data"] != data:
                            raise ValueError(f"Contradictory key {key}; previous origins={records[key]['origins']}")
                        records[key]["origins"].append(origin)
                        metrics["identical duplicates"] += 1
                    else:
                        records[key] = {"data": data, "origins": [origin]}
                        metrics["distinct source positions"] += 1
                except Exception as error:
                    raise ValueError(f"{origin}: {error}") from error
        finally:
            book.close()
    return records, paths



# IN05_PERIOD_MATCH_V2
def case_dates(records, roles):
    receipt_dates = {
        records[key]["data"]["entry"] for key in roles["RECEIPT"]
    }
    issue_dates = {
        records[key]["data"]["entry"] for key in roles["ISSUE"]
    }
    shared = receipt_dates & issue_dates
    if not shared:
        pattern = "DIFFERENT DAYS"
    elif len(receipt_dates | issue_dates) == 1:
        pattern = "SAME DAY"
    else:
        pattern = "SAME DAY AND DIFFERENT DAYS"
    return pattern, sorted(shared)


def evaluate(records, start, cutoff, scope, receipts, issues, reversals, metrics):
    groups = defaultdict(lambda: {"RECEIPT": [], "ISSUE": []})
    counts = defaultdict(Counter)
    documents = defaultdict(list)
    linked = defaultdict(list)
    gaps = []
    inside = lambda d: start.isoformat() <= d["entry"] <= cutoff.isoformat()
    for key, record in records.items():
        row = record["data"]
        documents[key[:3]].append(key)
        counts[row["company"]]
        if not inside(row):
            metrics["outside entry period"] += 1
            continue
        metrics["within entry period"] += 1
        if row["movement"] not in receipts | issues | set(reversals):
            metrics["positions outside configured movement sets"] += 1
        if row["movement"] in receipts | issues:
            if not row["material"].strip():
                metrics["configured originals without material"] += 1
                continue
            if row["quantity"] <= 0:
                metrics["configured originals with nonpositive quantity"] += 1
                continue
            if not row["plant"] or not row["user"]:
                raise ValueError(f"IN05 cannot evaluate selected position {key}: missing plant/user; origins={record['origins']}")
            role = "RECEIPT" if row["movement"] in receipts else "ISSUE"
            counts[row["company"]][role] += 1
            group = (row["company"], row["plant"], row["material"], row["user"])
            groups[group][role].append(key)
            if not row["unit"]:
                metrics["eligible originals missing unit"] += 1
    for key, record in records.items():
        row = record["data"]
        if row["movement"] not in reversals or not row["material"].strip() or row["quantity"] <= 0:
            continue
        doc, year, item = row["ref"]
        original = (row["company"], year, doc, item) if doc else None
        expected = reversals[row["movement"]]
        if original in records and records[original]["data"]["movement"] == expected:
            linked[original].append(key)
        elif inside(row):
            reason = (
                "MISSING ORIGINAL REFERENCE" if original is None else
                "ORIGINAL POSITION NOT SUPPLIED" if original not in records else
                "ORIGINAL MOVEMENT DOES NOT MATCH PARAM1"
            )
            gaps.append((key, reason))
            counts[row["company"]]["gaps"] += 1
            print(f"IN05 reversal reference gap: {reason}; position={key}; original={original}; origins={record['origins']}", flush=True)
    cases = []
    for group, roles in sorted(groups.items()):
        if roles["RECEIPT"] and roles["ISSUE"]:
            case = f"IN05-{len(cases) + 1:06d}"
            counts[group[0]]["cases"] += 1
            counts[group[0]]["triggers"] += sum(map(len, roles.values()))
            cases.append((case, group, roles))
    for company in scope or sorted(counts) or ["ALL"]:
        counts[company]
    summary = []
    for company, count in sorted(counts.items()):
        result = (
            "CASES / REVERSAL REFERENCE GAPS" if count["cases"] and count["gaps"] else
            "CASES FOR REVIEW" if count["cases"] else
            "PARTIAL EVIDENCE - SEE DETAIL" if count["gaps"] else
            "NO COMPARABLE RECEIPTS/ISSUES" if not count["RECEIPT"] or not count["ISSUE"] else
            "NO MATCHES IN PERIOD"
        )
        summary.append([
            company, result, count["RECEIPT"], count["ISSUE"], count["cases"],
            count["triggers"], count["gaps"], start.isoformat(), cutoff.isoformat(),
        ])
    metrics["cases"] = len(cases)
    metrics["reversal reference gaps"] = len(gaps)
    metrics["trigger positions across cases"] = sum(c["triggers"] for c in counts.values())
    metrics["configured receipts"] = sum(c["RECEIPT"] for c in counts.values())
    metrics["configured issues"] = sum(c["ISSUE"] for c in counts.values())
    return summary, cases, documents, linked, gaps


def evidence(records, cases, documents, linked, gaps, metrics):
    patterns = {
        case: case_dates(records, roles)
        for case, group, roles in cases
    }

    def output(key, case, group, role, trigger, status):
        row = records[key]["data"]
        origins = records[key]["origins"]
        file, sheet, line = origins[0].rsplit("|", 2)
        doc, year, item = row["ref"]
        return [
            case, row["company"], row["entry"], row["user"], row["material"], row["plant"],
            role, trigger, row["movement"], row["quantity"], row["unit"],
            row["document"], row["year"], row["item"], status,
            patterns[case][0] if group else "",
            json.dumps(group, ensure_ascii=False) if group else "",
            json.dumps(patterns[case][1]) if group else "",
            group[3] if group else "",
            group[2] if group else "", group[1] if group else "",
            row["storage"], row["batch"], row["posting"], row["time"], row["dc"],
            row["amount"], row["currency"], row["po"], row["po_item"], row["vendor"],
            row["reason"], row["text"], doc, year, item,
            json.dumps(linked.get(key, []), ensure_ascii=False), file, sheet, int(line),
            json.dumps(origins, ensure_ascii=False), json.dumps(key, ensure_ascii=False),
            "CONFIGURED MOVEMENTS / SUPPLIED POSITIONS ONLY",
        ]
    count = 0
    for case, group, roles in cases:
        triggers = {key: role for role, keys in roles.items() for key in keys}
        reversal_keys = {k for original in triggers for k in linked.get(original, [])}
        doc_keys = {k[:3] for k in set(triggers) | reversal_keys}
        for key in sorted({k for doc in doc_keys for k in documents[doc]}):
            role = triggers.get(key, "REVERSAL SUPPORT" if key in reversal_keys else "DOCUMENT SUPPORT")
            status = (
                "REFERENCE FOUND" if linked.get(key) else
                "MAPPED REVERSAL" if key in reversal_keys else
                "NOT FOUND IN SUPPLIED SOURCES" if key in triggers else ""
            )
            count += 1
            yield output(key, case, group, role, "YES" if key in triggers else "NO", status)
    for key, reason in sorted(gaps):
        count += 1
        yield output(key, "SOURCE GAP", None, "REVERSAL REFERENCE GAP", "NO", reason)
    metrics["detail evidence rows"] = count
    if count == 0:
        yield ["No receipt/issue matches in period or mapped reversal reference gaps in supplied sources. See IN05 for evaluated counts and scope."]


class PlainLayout:
    def __init__(self, stream, summary):
        self.stream = stream
        self.summary = summary

    def write(self, data):
        # IN_NO_MERGE_V1: no merged cells; only header styles are retained.
        value = data.decode("utf-8")

        def cell(match):
            attributes = match[1]
            reference = re.search(r'\br="[A-Z]+(\d+)"', attributes)
            if reference and int(reference[1]) != 1:
                attributes = re.sub(r'\s+s="\d+"', "", attributes)
            return "<c" + attributes + ">"

        value = re.sub(r"<c(\s[^>]+)>", cell, value)
        if value.startswith("<col "):
            column = ET.fromstring(value)
            index = int(column.get("min"))
            widths = {1: 12, 2: 42} if self.summary else {1: 20, 2: 12, 3: 16, 4: 20, 5: 30, 6: 14, 7: 28, 8: 12, 12: 22, 13: 12, 14: 12}
            if index in widths:
                column.set("width", str(widths[index]))
            value = ET.tostring(column, encoding="unicode")
        return self.stream.write(value.encode("utf-8"))


def write_output(context, cutoff, summary, details, notes):
    output = Path(context["output_folder"]) / f"LBR_Results_IN_{cutoff:%Y%m%d}.xlsx"
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="in05_output_", dir=output.parent) as temporary:
        temporary = Path(temporary)
        base = output
        if not output.exists():
            base = temporary / "base.xlsx"
            book = Workbook(write_only=True)
            for name in SHEETS:
                book.create_sheet(name)
            book.save(base)
        new = temporary / "result.xlsx"
        with zipfile.ZipFile(base) as original:
            workbook = ET.fromstring(original.read("xl/workbook.xml"))
            relationships = ET.fromstring(original.read("xl/_rels/workbook.xml.rels"))
            types = ET.fromstring(original.read("[Content_Types].xml"))
            styles = ET.fromstring(original.read("xl/styles.xml"))
            style_ids = _styles(styles)
            for style_id in style_ids:
                styles.find(_tag("cellXfs"))[style_id].find(_tag("alignment")).set("wrapText", "0")
            sheets = workbook.find(_tag("sheets"))
            paths = {}
            for name in SHEETS:
                existing = next((s for s in sheets if s.get("name") == name), None)
                if existing is not None:
                    relation = next(r for r in relationships if r.get("Id") == existing.get("{" + REL + "}id"))
                    target = relation.get("Target")
                    paths[name] = target.lstrip("/") if target.startswith("/") else "xl/" + target
                else:
                    index = 1
                    while f"xl/worksheets/in05_{index}.xml" in original.namelist() or f"xl/worksheets/in05_{index}.xml" in paths.values():
                        index += 1
                    path = f"xl/worksheets/in05_{index}.xml"
                    rid = 1
                    while any(r.get("Id") == f"rId{rid}" for r in relationships):
                        rid += 1
                    sheetid = max([0] + [int(s.get("sheetId")) for s in sheets]) + 1
                    ET.SubElement(sheets, _tag("sheet"), {"name": name, "sheetId": str(sheetid), "{" + REL + "}id": f"rId{rid}"})
                    ET.SubElement(relationships, "{" + PKG + "}Relationship", {"Id": f"rId{rid}", "Type": REL + "/worksheet", "Target": path[3:]})
                    ET.SubElement(types, "{" + CT + "}Override", {"PartName": "/" + path, "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"})
                    paths[name] = path
            replacements = {
                "xl/workbook.xml": _xml_preserve_namespaces(workbook, original.read("xl/workbook.xml")),
                "xl/_rels/workbook.xml.rels": _xml_preserve_namespaces(relationships, original.read("xl/_rels/workbook.xml.rels")),
                "[Content_Types].xml": _xml_preserve_namespaces(types, original.read("[Content_Types].xml")),
                "xl/styles.xml": _xml_preserve_namespaces(styles, original.read("xl/styles.xml")),
            }
            with zipfile.ZipFile(new, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=3, allowZip64=True) as destination:
                for info in original.infolist():
                    if info.filename in replacements or info.filename in paths.values():
                        continue
                    with original.open(info) as source, destination.open(info, "w", force_zip64=True) as target:
                        while True:
                            chunk = source.read(1024 * 1024)
                            if not chunk:
                                break
                            target.write(chunk)
                for path, content in replacements.items():
                    destination.writestr(path, content)
                for name, headers, rows, visible, footer in (
                    (SHEETS[0], SUMMARY, iter(summary), len(SUMMARY), lambda: notes),
                    (SHEETS[1], VISIBLE + TECH, details, len(VISIBLE), lambda: ["See IN05 for scope. One case is company/plant/material/user across the period. Match Pattern describes the case; support rows inherit that pattern. Trigger identifies qualifying receipt/issue positions, not cases. Document support can repeat across cases; expand technical columns with +."]),
                ):
                    with destination.open(paths[name], "w", force_zip64=True) as stream:
                        write_sheet(PlainLayout(stream, name == SHEETS[0]), headers, rows, visible, style_ids, footer)
        os.replace(new, output)
    return output


def run_in_005(context):
    started = perf_counter()
    start, cutoff = day(context["module"]["from"]), day(context["module"]["to"])
    if start > cutoff:
        raise ValueError("IN05 FROM is after TO.")
    raw_companies = txt(context["module"].get("companies"))
    scope = None if not raw_companies or raw_companies.upper() in {"ALL", "*", "TODAS", "TODOS", "NAN"} else {
        identifier(v, 4) for v in re.split(r"[,;\s]+", raw_companies) if v
    }
    receipts, issues, reversals, effective, default = parameters(context["control"].get("param1"))
    metrics = Counter()
    print("IN05 - Same company / plant / material / user across the entry-date period. Same/different-day patterns. Configured movements only.", flush=True)
    print(f"IN05 PARAM1 ({'DEFAULT' if default else 'CONFIGURED'}): {effective}", flush=True)
    stage = perf_counter()
    records, paths = read_sources(context["input_folder"], cutoff, scope, metrics)
    print(f"IN05 time | load and validate: {perf_counter() - stage:.2f} seconds", flush=True)
    stage = perf_counter()
    summary, cases, documents, linked, gaps = evaluate(records, start, cutoff, scope, receipts, issues, reversals, metrics)
    print(f"IN05 time | evaluate: {perf_counter() - stage:.2f} seconds", flush=True)
    notes = [
        f"Scope: entry dates {start.isoformat()} to {cutoff.isoformat()}; PARAM1={effective}. Other movement types and complete SAP coverage are not evaluated.",
        "Case = same company, plant, material and user with receipt and issue within the entry-date period. Match Pattern distinguishes SAME DAY, DIFFERENT DAYS, and SAME DAY AND DIFFERENT DAYS. Case Shared Entry Dates lists dates with both roles. No receipt-before-issue sequence is required. User is the document recorder; quantities need not match and physical custody is not established.",
        "Originals are retained. Reversal references are shown without assuming full cancellation; missing originals are source gaps. Trigger rows are not case counts; document support includes all supplied positions, not certified complete documents.",
    ]
    stage = perf_counter()
    output = write_output(context, cutoff, summary, evidence(records, cases, documents, linked, gaps, metrics), notes)
    print(f"IN05 time | write report: {perf_counter() - stage:.2f} seconds", flush=True)
    for path in paths:
        print("IN05 input:", path, flush=True)
    for name, value in sorted(metrics.items()):
        print(f"IN05 {name}: {value}", flush=True)
    print("IN05 output:", output, flush=True)
    print("Runner OK does not certify complete SAP coverage or a confirmed segregation-of-duties breach.", flush=True)
    print(f"IN05 time | total: {perf_counter() - started:.2f} seconds", flush=True)
    return output
