# IN02_PROJECT_STOCK_V1
"""IN02: negative cutoff quantities, independent of all other controls.

Runtime inputs: LBR IN_MARD_<TO>, LBR IN_MARD_H_<TO>,
LBR IN_MBEW_<TO>, LBR IN_MBEW_H_<TO> (.xlsx; numeric parts supported).
The configured cutoff and companies use core.in_common. No price, value,
deletion or material classification is used to exclude quantities.
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
from openpyxl.utils import get_column_letter
from core.in_common import (
    QTY, REL, PKG, CT, choose, meta_get, records_by_key,
    cutoff_from, companies_from, find_sources, database,
    load_source, write_sheet, _tag, _styles, _xml_preserve_namespaces,
)

CONTROL_ID = "IN_002"
SHEETS = ("IN02", "IN02_Detail")
SUMMARY = ["Company", "Population", "Result", "Evaluated", "Negative Records", "Pending Balance", "Cutoff"]
VISIBLE = ['Status', 'Company', 'Material', 'WBS Element', 'Special Stock', 'Location', 'Storage Location', 'Valuation Type', 'Stock Category', 'Quantity', 'Base Unit']

TECH = ["Cutoff", "Balance Evidence", "Population", "Record Key", "Source File", "Source Sheet", "Source Row", "All Source Origins", "Source Fiscal Year", "Source Period", "Unit Basis", "Current Record Flag (not excluded)", "Coverage"]


def evaluate(connection, cutoff, company_filter, metrics):
    connection.execute('CREATE TEMP TABLE detail(seq INTEGER PRIMARY KEY, payload TEXT)')
    groups = {}
    target = cutoff.year * 100 + cutoff.month
    metrics['IN02 negative keys'] = 0
    metrics['IN02 pending keys'] = 0
    metrics['IN02 detail rows'] = 0
    for family in ('Q', 'V', 'P'):
        for key, records in records_by_key(connection, family):
            selected, evidence, anchor = choose(records, target)
            source, row, origins = selected or anchor
            count = groups.setdefault((row['company'], family), Counter())
            count['evaluated'] += 1
            metrics[family + ' evaluated keys'] += 1
            metrics[family + ' ' + evidence] += 1
            if selected is None:
                count['pending'] += 1
                metrics['IN02 pending keys'] += 1
                cases = [('UNAVAILABLE', None)]
                status = 'PENDING BALANCE EVIDENCE'
            else:
                categories = QTY if family == 'Q' else ('Total Stock',)
                cases = [(name, Decimal(row['numbers'][name])) for name in categories if Decimal(row['numbers'][name]) < 0]
                if not cases:
                    continue
                count['negative'] += 1
                metrics['IN02 negative keys'] += 1
                status = 'NEGATIVE QUANTITY'
            unit, basis = (row['unit'], 'Selected source master attribute') if row['unit'] else meta_get(connection, 'unit', row['material'])
            if not unit:
                status += '; UNIT UNAVAILABLE'
                metrics[family + ' reported keys missing unit'] += 1
            population = 'Storage-location quantities' if family == 'Q' else 'Project valuation quantities' if family == 'P' else 'Valuation quantities'
            first = origins[0].rsplit('|', 2)
            for category, quantity in cases:
                values = [status, row['company'], row['material'], row.get('wbs', ''), row.get('special_stock', ''), row['plant'] if family == 'Q' else row['area'], row['sloc'] if family == 'Q' else None, row['vt'] if family != 'Q' else None, category, str(quantity) if quantity is not None else None, unit, cutoff.isoformat(), evidence, population, family + '|' + key, first[0], first[1], int(first[2]), json.dumps(origins, ensure_ascii=False), row['period'] // 100, row['period'] % 100, basis, anchor[1]['flag'], 'SUPPLIED TABLES ONLY - PRELIMINARY']
                connection.execute('INSERT INTO detail(payload) VALUES(?)', (json.dumps(values, ensure_ascii=False),))
                metrics['IN02 detail rows'] += 1
    companies = company_filter or sorted({company for company, family in groups}) or ['ALL']
    for company in companies:
        for family in ('Q', 'V', 'P'):
            groups.setdefault((company, family), Counter())
    summaries = []
    for (company, family), count in sorted(groups.items()):
        result = 'NO SUPPLIED RECORDS' if not count['evaluated'] else 'NEGATIVE QUANTITIES / PARTIAL EVIDENCE' if count['negative'] and count['pending'] else 'NEGATIVE QUANTITIES FOUND' if count['negative'] else 'PARTIAL EVIDENCE - SEE DETAIL' if count['pending'] else 'NO NEGATIVE QUANTITIES'
        summaries.append([company, 'Storage-location quantities' if family == 'Q' else 'Project valuation quantities' if family == 'P' else 'Valuation quantities', result, count['evaluated'], count['negative'], count['pending'], cutoff.isoformat()])
    return summaries



def detail_rows(connection):
    found = False
    for payload, in connection.execute('SELECT payload FROM detail ORDER BY seq'):
        found = True
        row = json.loads(payload)
        if row[9] is not None:
            row[9] = Decimal(row[9])
        yield row
    if not found:
        yield ['No negative quantities or pending cutoff balances in supplied sources. See IN02 for scope.']



class PlainLayout:
    def __init__(self, stream, widths, visible, records):
        self.stream, self.widths, self.visible = stream, widths, visible
        self.last, self.notes = max(1, records) + 1, []

    def write(self, data):
        text = data.decode("utf-8")
        if text.startswith("<col "):
            column = ET.fromstring(text)
            index = int(column.get("min")) - 1
            if index < len(self.widths):
                column.set("width", str(self.widths[index]))
            text = ET.tostring(column, encoding="unicode")
        def row(match):
            number = int(match[1])
            if number > self.last:
                self.notes.append(number)
            return match[0]
        text = re.sub(r'<row r="(\d+)">', row, text)
        def cell(match):
            attrs = match[1]
            ref = re.search(r'\br="[A-Z]+(\d+)"', attrs)
            if ref and int(ref[1]) != 1:
                attrs = re.sub(r'\s+s="\d+"', '', attrs)
            return '<c' + attrs + '>'
        text = re.sub(r'<c(\s[^>]+)>', cell, text)
        # IN_NO_MERGE_V1
        return self.stream.write(text.encode("utf-8"))


def write_output(context, connection, cutoff, summaries, metrics):
    output = Path(context['output_folder']) / f'LBR_Results_IN_{cutoff:%Y%m%d}.xlsx'
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='in02_output_', dir=output.parent) as temporary:
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
                    while f'xl/worksheets/in02_{index}.xml' in original.namelist() or f'xl/worksheets/in02_{index}.xml' in paths.values():
                        index += 1
                    path = f'xl/worksheets/in02_{index}.xml'
                    rid = 1
                    while any((r.get('Id') == f'rId{rid}' for r in relations)):
                        rid += 1
                    sheetid = max([0] + [int(s.get('sheetId')) for s in sheets]) + 1
                    ET.SubElement(sheets, _tag('sheet'), {'name': name, 'sheetId': str(sheetid), '{' + REL + '}id': f'rId{rid}'})
                    ET.SubElement(relations, '{' + PKG + '}Relationship', {'Id': f'rId{rid}', 'Type': REL + '/worksheet', 'Target': path[3:]})
                    ET.SubElement(types, '{' + CT + '}Override', {'PartName': '/' + path, 'ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'})
                    paths[name] = path
            replacements = {'xl/workbook.xml': _xml_preserve_namespaces(workbook, original.read('xl/workbook.xml')), 'xl/_rels/workbook.xml.rels': _xml_preserve_namespaces(relations, original.read('xl/_rels/workbook.xml.rels')), '[Content_Types].xml': _xml_preserve_namespaces(types, original.read('[Content_Types].xml')), 'xl/styles.xml': _xml_preserve_namespaces(styles, original.read('xl/styles.xml'))}
            notes = [f'Scope: supplied MARD/MARDH, MBEW/MBEWH and QBEW/QBEWH balances at {cutoff.isoformat()}. Project valuation is included separately. Complete SAP coverage, other special stock and lot detail are not validated.', 'Negative Records counts source keys with a quantity below zero; categories are not netted. Populations have different grains and must be reviewed separately.', 'Pending Balance means cutoff evidence is unavailable, not zero. Detail may contain several negative categories for one key; technical fields expand with +.']
            with zipfile.ZipFile(new, 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=3, allowZip64=True) as destination:
                for info in original.infolist():
                    if info.filename in replacements or info.filename in paths.values():
                        continue
                    with original.open(info) as src, destination.open(info, 'w', force_zip64=True) as dst:
                        while True:
                            chunk = src.read(1024 * 1024)
                            if not chunk:
                                break
                            dst.write(chunk)
                for path, content in replacements.items():
                    destination.writestr(path, content)
                for name, headers, rows, visible, count, widths, scope in ((SHEETS[0], SUMMARY, iter(summaries), len(SUMMARY), len(summaries), [12, 30, 44, 18, 22, 20, 16], lambda: notes), (SHEETS[1], VISIBLE + TECH, detail_rows(connection), len(VISIBLE), metrics['IN02 detail rows'], [42, 12, 30, 18, 20, 20, 26, 18, 16], lambda: ['See IN02 for scope and counts. Pending evidence is reported separately from negative quantities.'])):
                    with destination.open(paths[name], 'w', force_zip64=True) as stream:
                        write_sheet(PlainLayout(stream, widths, visible, count), headers, rows, visible, style_ids, scope)
        os.replace(new, output)
    return output



def run_in_002(context):
    cutoff = cutoff_from(context['module'])
    companies = companies_from(context['module'].get('companies'))
    sources = find_sources(context['input_folder'], cutoff)
    metrics = Counter()
    metrics['_project_cutoff'] = cutoff.strftime('%Y%m%d')
    metrics['_input'] = str(Path(context['input_folder']))
    metrics['_control'] = 'IN02'
    print('IN02 - Negative quantities at configured cutoff; independent source evaluation.', flush=True)
    with tempfile.TemporaryDirectory(prefix='lbr_in02_') as temporary:
        connection = database(str(Path(temporary) / 'working.sqlite'))
        try:
            for source in ('MARD', 'MARDH', 'MBEW', 'MBEWH', 'QBEW', 'QBEWH'):
                metrics[source + ' files'] = len(sources[source])
                load_source(connection, source, sources[source], companies, metrics)
            connection.execute('CREATE INDEX snapshot_family_key ON snapshot(family,k,source,p)')
            connection.commit()
            summaries = evaluate(connection, cutoff, companies, metrics)
            output = write_output(context, connection, cutoff, summaries, metrics)
        finally:
            connection.close()
    for name, paths in sources.items():
        for path in paths:
            print(f'IN02 input {name}: {path}', flush=True)
    for name, value in sorted(metrics.items()):
        if not name.startswith('_'):
            print(f'{name}: {value}', flush=True)
    print('IN02 output:', output, flush=True)
    print('Runner OK does not certify complete SAP inventory coverage.', flush=True)
    return output

