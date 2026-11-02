# IN03_PROJECT_STOCK_V1
"""IN03: all supplied valuation keys are tested for zero SAP cost.

Review evidence: zero cost with nonzero valued quantity or inventory value.
Zero-cost records with zero quantity AND value remain in informative counts.
Pending history, unknown price control and nonpositive Price Unit are not zero.
Runtime sources: IN_MBEW and IN_MBEW_H for balances/prices; IN_MARD and
IN_MARD_H for supplied unit metadata. Numeric parts are supported.
No other control, report, material or deletion classification is used.
"""
from collections import Counter
from decimal import Decimal
import json
import os
from pathlib import Path
import re
import tempfile
from xml.etree import ElementTree as ET
import zipfile
from openpyxl import Workbook
from core.in_common import (
    QTY, REL, PKG, CT, choose, meta_get, records_by_key,
    cutoff_from, companies_from, find_sources, database,
    load_source, write_sheet, _tag, _styles, _xml_preserve_namespaces,
)

CONTROL_ID = "IN_003"
SHEETS = ("IN03", "IN03_Detail")
SUMMARY = ['Company', 'Population', 'Result', 'Evaluated', 'Zero Cost With Balance', 'Zero Cost / Zero Balance', 'Pending Evaluation', 'Other Records', 'Cutoff']
VISIBLE = ['Status', 'Company', 'Material', 'Population', 'WBS Element', 'Special Stock', 'Valuation Area', 'Valuation Type', 'Quantity', 'Base Unit', 'Inventory Value', 'Currency', 'SAP Cost per Base Unit', 'Price Control', 'Valuation Class']
TECH = [
    "Cutoff", "Balance Evidence", "Valuation Record Level",
    "Moving Average Price", "Standard Price", "Price Unit",
    "Record Key", "Source File", "Source Sheet", "Source Row",
    "All Source Origins", "Source Fiscal Year", "Source Period",
    "Unit Basis", "Currency Basis",
    "Current Record Flag (not excluded)", "Coverage",
]


def evaluate(connection, cutoff, company_filter, metrics):
    connection.execute(
        "CREATE TEMP TABLE detail(seq INTEGER PRIMARY KEY, payload TEXT)"
    )
    groups = {}
    target = cutoff.year * 100 + cutoff.month
    metric_names = {
        "review": "zero cost with balance",
        "informative": "zero cost zero balance",
        "pending": "pending evaluation",
        "other": "other records",
    }
    for name in (
        "evaluated keys", "zero cost with balance",
        "zero cost zero balance", "pending evaluation",
        "other records", "detail rows",
    ):
        metrics["IN03 " + name] = 0

    for family in ("V", "P"):
        for key, records in records_by_key(connection, family):
            selected, evidence, anchor = choose(records, target)
            source, row, origins = selected or anchor
            company = row["company"]
            count = groups.setdefault((company, family), Counter())
            count["evaluated"] += 1
            metrics["IN03 evaluated keys"] += 1
            metrics[family + " " + evidence] += 1
            population = (
                "Project valuation records" if family == "P"
                else "Valuation records"
            )

            quantity = value = average = standard = price_unit = cost = None
            control = row["price_control"]

            if selected is None:
                status, bucket = "PENDING BALANCE EVIDENCE", "pending"
            else:
                numbers = row["numbers"]
                quantity = Decimal(numbers["Total Stock"])
                value = Decimal(numbers["Total Value"])
                average = Decimal(numbers["MovAvgPrice"])
                standard = Decimal(numbers["Standard price"])
                price_unit = Decimal(numbers["per"])
                issues = []
                if control not in ("S", "V"):
                    issues.append("UNKNOWN PRICE CONTROL")
                if price_unit <= 0:
                    issues.append("INVALID PRICE UNIT")
                if issues:
                    status = "PENDING - " + " / ".join(issues)
                    bucket = "pending"
                else:
                    cost = (
                        standard if control == "S" else average
                    ) / price_unit
                    if cost == 0:
                        bucket = (
                            "review" if quantity != 0 or value != 0
                            else "informative"
                        )
                        status = (
                            "ZERO COST WITH VALUATION BALANCE"
                            if bucket == "review"
                            else "ZERO COST / ZERO BALANCE - INFORMATIVE"
                        )
                    else:
                        status, bucket = "NONZERO COST", "other"

            count[bucket] += 1
            metrics["IN03 " + metric_names[bucket]] += 1
            if bucket not in ("review", "pending"):
                continue

            unit, unit_basis = (
                (row["unit"], "Selected source master attribute")
                if row["unit"]
                else meta_get(connection, "unit", row["material"])
            )
            currency, currency_basis = (
                (row["currency"], "Selected source master attribute")
                if row["currency"]
                else meta_get(connection, "currency", company)
            )
            if not unit:
                status += "; UNIT UNAVAILABLE"
                metrics["IN03 reported keys missing unit"] += 1
            if not currency:
                status += "; CURRENCY UNAVAILABLE"
                metrics["IN03 reported keys missing currency"] += 1

            first = origins[0].rsplit("|", 2)
            values = dict(zip(VISIBLE + TECH, [
                status, company, row["material"],
                population, row.get("wbs", ""), row.get("special_stock", ""),
                row["area"], row["vt"], quantity, unit, value, currency,
                cost, control, row["valuation_class"],
                cutoff.isoformat(), evidence,
                "BLANK VALUATION TYPE" if not row["vt"] else "VALUATION TYPE DETAIL",
                average, standard, price_unit, key,
                first[0], first[1], int(first[2]),
                json.dumps(origins, ensure_ascii=False),
                row["period"] // 100, row["period"] % 100,
                unit_basis, currency_basis, anchor[1]["flag"],
                "SUPPLIED ORDINARY/PROJECT VALUATION TABLES ONLY - PRELIMINARY",
            ]))
            # La lista se verifica tambien mediante las pruebas del instalador.
            payload = [values[h] for h in VISIBLE + TECH]
            connection.execute(
                "INSERT INTO detail(payload) VALUES(?)",
                (json.dumps(payload, ensure_ascii=False, default=str),),
            )
            metrics["IN03 detail rows"] += 1

    companies = (
        company_filter
        or sorted({company for company, family in groups})
        or ["ALL"]
    )
    for company in companies:
        for family in ("V", "P"):
            groups.setdefault((company, family), Counter())

    summaries = []
    for (company, family), count in sorted(groups.items()):
        result = (
            "NO SUPPLIED VALUATION RECORDS" if not count["evaluated"]
            else "ZERO COST CASES / PARTIAL EVIDENCE"
            if count["review"] and count["pending"]
            else "ZERO COST CASES FOUND" if count["review"]
            else "PARTIAL EVIDENCE - SEE DETAIL" if count["pending"]
            else "ZERO COST ONLY ON ZERO BALANCES" if count["informative"]
            else "NO ZERO COST IDENTIFIED"
        )
        summaries.append([
            company,
            "Project valuation records" if family == "P" else "Valuation records",
            result, count["evaluated"], count["review"],
            count["informative"], count["pending"], count["other"],
            cutoff.isoformat(),
        ])
    return summaries


def detail_rows(connection):
    found = False
    headings = VISIBLE + TECH
    numeric = {
        "Quantity", "Inventory Value", "SAP Cost per Base Unit",
        "Moving Average Price", "Standard Price", "Price Unit",
    }
    for (payload,) in connection.execute(
        "SELECT payload FROM detail ORDER BY seq"
    ):
        found = True
        values = json.loads(payload)
        if len(values) != len(headings):
            raise ValueError("IN03 detail/header length mismatch")
        for i, heading in enumerate(headings):
            if heading in numeric and values[i] is not None:
                values[i] = Decimal(str(values[i]))
        yield values
    if not found:
        yield [
            "No zero-cost records with valuation balance or pending evaluations. "
            "Zero-cost/zero-balance records are summarized in IN03."
        ]


class PlainLayout:
    def __init__(self, stream, widths, visible, records):
        self.stream, self.widths = stream, widths

    def write(self, data):
        # IN_NO_MERGE_V1
        text = data.decode("utf-8")

        if text.startswith("<col "):
            column = ET.fromstring(text)
            index = int(column.get("min")) - 1
            if index < len(self.widths):
                column.set("width", str(self.widths[index]))
            text = ET.tostring(column, encoding="unicode")

        def cell(match):
            attributes = match[1]
            reference = re.search(
                r'\br="[A-Z]+(\d+)"', attributes
            )
            if reference and int(reference[1]) != 1:
                attributes = re.sub(
                    r'\s+s="\d+"', "", attributes
                )
            return "<c" + attributes + ">"

        text = re.sub(r"<c(\s[^>]+)>", cell, text)
        return self.stream.write(text.encode("utf-8"))


def write_output(context, connection, cutoff, summaries, metrics):
    output = Path(context['output_folder']) / f'LBR_Results_IN_{cutoff:%Y%m%d}.xlsx'
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='in03_output_', dir=output.parent) as temporary:
        temporary = Path(temporary)
        base = output
        if not output.exists():
            base = temporary / 'base.xlsx'
            workbook = Workbook(write_only=True)
            for name in SHEETS:
                workbook.create_sheet(name)
            workbook.save(base)
        new = temporary / 'result.xlsx'
        with zipfile.ZipFile(base) as original:
            workbook = ET.fromstring(original.read('xl/workbook.xml'))
            relations = ET.fromstring(original.read('xl/_rels/workbook.xml.rels'))
            types = ET.fromstring(original.read('[Content_Types].xml'))
            styles = ET.fromstring(original.read('xl/styles.xml'))
            style_ids = _styles(styles)
            for style_id in style_ids:
                styles.find(_tag('cellXfs'))[style_id].find(_tag('alignment')).set('wrapText', '0')
            sheets = workbook.find(_tag('sheets'))
            paths = {}
            for name in SHEETS:
                existing = next((s for s in sheets if s.get('name') == name), None)
                if existing is not None:
                    relation = next((r for r in relations if r.get('Id') == existing.get('{' + REL + '}id')))
                    target = relation.get('Target')
                    paths[name] = target.lstrip('/') if target.startswith('/') else 'xl/' + target
                else:
                    index = 1
                    while f'xl/worksheets/in03_{index}.xml' in original.namelist() or f'xl/worksheets/in03_{index}.xml' in paths.values():
                        index += 1
                    path = f'xl/worksheets/in03_{index}.xml'
                    rid = 1
                    while any((r.get('Id') == f'rId{rid}' for r in relations)):
                        rid += 1
                    sheetid = max([0] + [int(s.get('sheetId')) for s in sheets]) + 1
                    ET.SubElement(sheets, _tag('sheet'), {'name': name, 'sheetId': str(sheetid), '{' + REL + '}id': f'rId{rid}'})
                    ET.SubElement(relations, '{' + PKG + '}Relationship', {'Id': f'rId{rid}', 'Type': REL + '/worksheet', 'Target': path[3:]})
                    ET.SubElement(types, '{' + CT + '}Override', {'PartName': '/' + path, 'ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'})
                    paths[name] = path
            replacements = {'xl/workbook.xml': _xml_preserve_namespaces(workbook, original.read('xl/workbook.xml')), 'xl/_rels/workbook.xml.rels': _xml_preserve_namespaces(relations, original.read('xl/_rels/workbook.xml.rels')), '[Content_Types].xml': _xml_preserve_namespaces(types, original.read('[Content_Types].xml')), 'xl/styles.xml': _xml_preserve_namespaces(styles, original.read('xl/styles.xml'))}
            notes = [f'Scope: supplied MBEW/MBEWH and project QBEW/QBEWH valuation records at {cutoff.isoformat()}. Ordinary and project populations are separate. Complete SAP coverage, other special stock and lot detail are not validated.', 'Cost = standard price (S) or moving average price (V), divided by Price Unit. Zero cost with zero quantity and value is informational.', 'Review: cost = 0 with quantity or inventory value <> 0. Missing/invalid evidence is pending, never zero. Valuation levels are reviewed separately.']
            with zipfile.ZipFile(new, 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=3, allowZip64=True) as destination:
                for info in original.infolist():
                    if info.filename in replacements or info.filename in paths.values():
                        continue
                    with original.open(info) as source:
                        with destination.open(info, 'w', force_zip64=True) as target:
                            while True:
                                chunk = source.read(1024 * 1024)
                                if not chunk:
                                    break
                                target.write(chunk)
                for path, content in replacements.items():
                    destination.writestr(path, content)
                for name, headers, rows, visible, count, widths, scope in ((SHEETS[0], SUMMARY, iter(summaries), len(SUMMARY), len(summaries), [12, 44, 18, 28, 28, 24, 18, 16], lambda: notes), (SHEETS[1], VISIBLE + TECH, detail_rows(connection), len(VISIBLE), metrics['IN03 detail rows'], [48, 12, 30, 20, 20, 18, 16, 20, 16, 26, 18, 20], lambda: ['See IN03 for scope and counts. Zero-cost cases and pending evaluations have separate statuses.'])):
                    with destination.open(paths[name], 'w', force_zip64=True) as stream:
                        write_sheet(PlainLayout(stream, widths, visible, count), headers, rows, visible, style_ids, scope)
        os.replace(new, output)
    return output


def run_in_003(context):
    from time import perf_counter
    started = perf_counter()
    cutoff = cutoff_from(context['module'])
    companies = companies_from(context['module'].get('companies'))
    sources = find_sources(context['input_folder'], cutoff)
    metrics = Counter()
    metrics['_project_cutoff'] = cutoff.strftime('%Y%m%d')
    metrics['_input'] = str(Path(context['input_folder']))
    metrics['_control'] = 'IN03'

    def elapsed(stage, since):
        print(f'IN03 time | {stage}: {perf_counter() - since:.2f} seconds', flush=True)
    print('IN03 - Zero SAP cost at configured cutoff; independent valuation evaluation.', flush=True)
    with tempfile.TemporaryDirectory(prefix='lbr_in03_') as temporary:
        connection = database(str(Path(temporary) / 'working.sqlite'))
        try:
            connection.execute('PRAGMA journal_mode=OFF')
            connection.execute('PRAGMA synchronous=OFF')
            connection.execute('PRAGMA cache_size=-65536')
            print('IN03 temporary SQLite: recovery journal disabled; cache target 64 MiB.', flush=True)
            for source in ('MARD', 'MARDH', 'MBEW', 'MBEWH', 'QBEW', 'QBEWH'):
                stage = perf_counter()
                metrics[source + ' files'] = len(sources[source])
                load_source(connection, source, sources[source], companies, metrics)
                elapsed('load ' + source, stage)
            stage = perf_counter()
            connection.execute('CREATE INDEX snapshot_family_key ON snapshot(family,k,source,p)')
            connection.commit()
            elapsed('build index', stage)
            stage = perf_counter()
            summaries = evaluate(connection, cutoff, companies, metrics)
            elapsed('evaluate', stage)
            stage = perf_counter()
            output = write_output(context, connection, cutoff, summaries, metrics)
            elapsed('write report', stage)
        finally:
            connection.close()
    for name, paths in sources.items():
        for path in paths:
            print(f'IN03 input {name}: {path}', flush=True)
    for name, value in sorted(metrics.items()):
        if not name.startswith('_'):
            print(f'{name}: {value}', flush=True)
    print('IN03 output:', output, flush=True)
    print('Runner OK does not certify complete SAP inventory coverage.', flush=True)
    elapsed('total control', started)
    return output
