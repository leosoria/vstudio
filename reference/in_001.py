# IN01_PROJECT_STOCK_V1

"""IN01 - Preliminary inventory listing from supplied sources."""
from collections import Counter
from decimal import Decimal
import json
import os
from pathlib import Path
import tempfile
from xml.etree import ElementTree as ET
import zipfile

from openpyxl import Workbook

from core.in_common import (
    QTY, REL, PKG, CT,
    choose, meta_get, records_by_key,
    cutoff_from, companies_from, find_sources,
    database, load_source, write_sheet,
    _tag, _styles, _xml_preserve_namespaces,
)

CONTROL_ID = "IN_001"
SHEETS = ("IN01", "IN01_Detail")


def report_rows(connection, family, cutoff, metrics):
    target = cutoff.year * 100 + cutoff.month
    for key, records in records_by_key(connection, family):
        selected, status, anchor = choose(records, target)
        row = (selected or anchor)[1]
        origins = (selected or anchor)[2]
        unit, unit_basis = (row['unit'], 'Selected source master attribute') if row['unit'] else meta_get(connection, 'unit', row['material'])
        currency, currency_basis = (row['currency'], 'Selected source master attribute') if row['currency'] else meta_get(connection, 'currency', row['company'])
        metrics[family + ' output records'] += 1
        metrics[family + ' ' + status] += 1
        metrics[family + ' missing unit'] += int(not unit)
        if family != 'Q':
            metrics[family + ' missing currency'] += int(not currency)
        numbers = selected[1]['numbers'] if selected else {}
        if selected:
            measured = QTY if family == 'Q' else ('Total Stock', 'Total Value')
            metrics[family + ' records with supplied nonzero balances'] += int(any((Decimal(numbers[heading]) != 0 for heading in measured)))
        if family == 'Q':
            visible = [row['company'], row['plant'], row['sloc'], row['material'], unit, *[Decimal(numbers[heading]) if selected else None for heading in QTY], cutoff.isoformat(), status]
        else:
            price_unit = Decimal(numbers['per']) if selected else None
            control = selected[1]['price_control'] if selected else row['price_control']
            raw_price = numbers.get('Standard price') if control == 'S' else numbers.get('MovAvgPrice') if control == 'V' else None
            cost = Decimal(raw_price) / price_unit if raw_price is not None and price_unit is not None and (price_unit > 0) else None
            price_status = 'SAP PRICE CONTROL / PRICE UNIT' if cost is not None else 'UNAVAILABLE OR INVALID PRICE CONTROL/UNIT'
            visible = [row['company'], row['area'], row['material'], row['vt'], 'BLANK VALUATION TYPE' if not row['vt'] else 'VALUATION TYPE DETAIL', unit, currency, Decimal(numbers['Total Stock']) if selected else None, Decimal(numbers['Total Value']) if selected else None, control, Decimal(numbers['MovAvgPrice']) if selected else None, Decimal(numbers['Standard price']) if selected else None, price_unit, cost, row['valuation_class'], None, 'NOT SUPPLIED', cutoff.isoformat(), status]
        first = origins[0].rsplit('|', 2)
        technical = [row['area'], (selected or anchor)[0], row['period'] // 100, row['period'] % 100, first[0], first[1], int(first[2]), json.dumps(origins, ensure_ascii=False), unit_basis, currency_basis if family != 'Q' else 'Not applicable', repr(row['material']), int('\xa0' in row['material']), anchor[1]['flag'], 'SUPPLIED TABLES ONLY - PRELIMINARY', key]
        if family != 'Q':
            technical.append(price_status)
        yield (visible + technical + [row.get('wbs', ''), row.get('special_stock', '')])



Q_HEADERS = [
    "Company", "Plant", "Storage Location",
    "Material", "Base Unit",
    *QTY, "Cutoff", "Balance Evidence",
]

V_HEADERS = [
    "Company", "Valuation Area", "Material",
    "Valuation Type", "Valuation Record Level",
    "Base Unit", "Currency",
    "Valuated Quantity", "Inventory Value",
    "Price Control", "Moving Average Price",
    "Standard Price", "Price Unit",
    "SAP Cost per Base Unit", "Valuation Class",
    "Sales Price", "Sales Price Status",
    "Cutoff", "Balance Evidence",
]

TECH_HEADERS = [
    "Source Valuation Area",
    "Selected or Candidate Source",
    "Source Fiscal Year", "Source Period",
    "Source File", "Source Sheet", "Source Row",
    "All Duplicate Origins",
    "Unit Basis", "Currency Basis",
    "Material Original Representation",
    "Material Contains NBSP",
    "Current Record Flag (not excluded)",
    "Coverage Status", "Record Key",
]



# IN01_SUMMARY_DETAIL_V1

SUMMARY_HEADERS = [
    "Company", "Population", "Records Evaluated", "With Balance",
    "Zero Balance", "Pending Balance", "Data Gaps", "Result", "Cutoff",
]

DETAIL_VISIBLE = ['Evidence / Review', 'Company', 'Material', 'WBS Element', 'Special Stock', 'Population', 'Location', 'Storage Location', 'Valuation Type', 'Stock Category', 'Quantity', 'Base Unit', 'Inventory Value', 'Currency', 'SAP Cost per Base Unit']


DETAIL_TECH = [
    "Cutoff", "Balance Evidence", *TECH_HEADERS,
    "SAP Cost Calculation Basis", "Price Control",
    "Moving Average Price", "Standard Price", "Price Unit",
    "Valuation Class", "Valuation Record Level",
]


def _in01_prepare(context, connection, cutoff, metrics):
    connection.execute('DROP TABLE IF EXISTS temp.in01_detail')
    connection.execute('CREATE TEMP TABLE in01_detail(seq INTEGER PRIMARY KEY, payload TEXT)')
    connection.execute('DROP TABLE IF EXISTS temp.in01_support_keys')
    connection.execute('CREATE TEMP TABLE in01_support_keys(company TEXT, material TEXT, area TEXT, PRIMARY KEY(company,material,area)) WITHOUT ROWID')
    summaries = {}
    metrics['IN01 detail evidence rows'] = 0
    for family, headers in (('Q', Q_HEADERS), ('V', V_HEADERS), ('P', V_HEADERS)):
        full = headers + TECH_HEADERS + (['SAP Cost Calculation Basis'] if family != 'Q' else []) + ['WBS Element', 'Special Stock']
        for values in report_rows(connection, family, cutoff, metrics):
            row = dict(zip(full, values))
            count = summaries.setdefault((row['Company'], family), Counter())
            count['evaluated'] += 1
            pending = row['Balance Evidence'].startswith('PENDING')
            fields = QTY if family == 'Q' else ('Valuated Quantity', 'Inventory Value')
            nonzero = not pending and any((row[heading] != 0 for heading in fields))
            count['pending' if pending else 'nonzero' if nonzero else 'zero'] += 1
            count['unit'] += int(not row['Base Unit'])
            count['currency'] += int(family != 'Q' and (not row['Currency']))
            support_key = (row['Company'], row['Material'], row['Source Valuation Area'])
            if family == 'Q' and (pending or nonzero):
                connection.execute('INSERT OR IGNORE INTO in01_support_keys VALUES(?,?,?)', support_key)
            linked = family == 'V' and connection.execute('SELECT 1 FROM in01_support_keys WHERE company=? AND material=? AND area=?', support_key).fetchone() is not None
            if not pending and (not nonzero) and (not linked):
                continue
            reason = 'BALANCE EVIDENCE PENDING' if pending else 'NONZERO BALANCE - LISTING SUPPORT' if nonzero else 'ZERO VALUATION - SUPPORT FOR LISTED QUANTITIES'
            if not row['Base Unit']:
                reason += '; UNIT UNAVAILABLE'
            if family != 'Q' and (not row['Currency']):
                reason += '; CURRENCY UNAVAILABLE'
            if family == 'Q':
                categories = [('UNAVAILABLE', None)] if pending else [(heading, row[heading]) for heading in QTY if row[heading] != 0]
            else:
                categories = [('VALUATION', row['Valuated Quantity'])]
            for category, quantity in categories:
                detail = [reason, row['Company'], row['Material'], row.get('WBS Element', ''), row.get('Special Stock', ''), 'Storage-location quantities' if family == 'Q' else 'Project valuation records' if family == 'P' else 'Valuation records', row['Plant'] if family == 'Q' else row['Valuation Area'], row.get('Storage Location'), row.get('Valuation Type'), category, quantity, row['Base Unit'], row.get('Inventory Value'), row.get('Currency'), row.get('SAP Cost per Base Unit'), *[row.get(heading) for heading in DETAIL_TECH]]
                connection.execute('INSERT INTO in01_detail(payload) VALUES(?)', (json.dumps(detail, ensure_ascii=False, default=str),))
                metrics['IN01 detail evidence rows'] += 1
    companies = companies_from(context['module'].get('companies'))
    population = companies or sorted({company for company, family in summaries}) or ['ALL']
    for company in population:
        for family in ('Q', 'V', 'P'):
            summaries.setdefault((company, family), Counter())
    result = []
    for (company, family), count in sorted(summaries.items()):
        gaps = []
        if count['unit']:
            gaps.append(f"Missing unit: {count['unit']:,}")
        if count['currency']:
            gaps.append(f"Missing currency: {count['currency']:,}")
        status = 'NO SUPPLIED RECORDS' if not count['evaluated'] else 'PENDING BALANCES - SEE DETAIL' if count['pending'] else 'BALANCES PRESENT - SEE DETAIL' if count['nonzero'] else 'ZERO BALANCES IN SUPPLIED SOURCES'
        result.append([company, 'Storage-location quantities' if family == 'Q' else 'Project valuation records' if family == 'P' else 'Valuation records', count['evaluated'], count['nonzero'], count['zero'], count['pending'], '; '.join(gaps) or 'None identified in required unit/currency fields', status, cutoff.isoformat()])
    for family in ('Q', 'V', 'P'):
        metrics[family + ' evaluated records'] = metrics.pop(family + ' output records', 0)
    metrics['IN01 summary rows'] = len(result)
    return result



def _in01_details(connection):
    found = False
    headers = DETAIL_VISIBLE + DETAIL_TECH
    numeric = {
        "Quantity", "Inventory Value", "SAP Cost per Base Unit",
        "Moving Average Price", "Standard Price", "Price Unit",
    }

    for (payload,) in connection.execute(
        "SELECT payload FROM in01_detail ORDER BY seq"
    ):
        found = True
        values = json.loads(payload)

        for index, heading in enumerate(headers):
            if heading in numeric and values[index] is not None:
                values[index] = Decimal(str(values[index]))

        yield values

    if not found:
        yield [
            "No detail required: no nonzero or pending cutoff "
            "balances in supplied sources. See IN01 for scope "
            "and data gaps."
        ]


def _in01_notes(context, cutoff, metrics):
    return [f'Scope: supplied MARD/MARDH, MBEW/MBEWH and project valuation QBEW/QBEWH at {cutoff.isoformat()}. Populations are separate; complete SAP coverage and lot detail are not certified.', 'Project valuation keys include company, material, valuation area, valuation type, special stock and WBS. Do not add historical snapshots to current balances.', 'Base units may be completed from the supplied MARA master at extraction; company currencies use supplied IN_VAL_MAP. Missing or conflicting metadata remains visible.', 'IN01_Detail includes nonzero balances, pending evidence and related ordinary valuation support. Evidence rows are not audit exception counts.']




class _IN01Layout:
    def __init__(
        self, stream, widths, visible, records, text_style
    ):
        self.stream = stream
        self.widths = widths
        self.visible = visible
        self.last = max(1, records) + 1
        self.text_style = text_style
        self.note_rows = []

    def write(self, data):
        # IN01_PLAIN_BODY_V1
        import re
        from openpyxl.utils import get_column_letter

        text = data.decode("utf-8")

        if text.startswith("<col "):
            node = ET.fromstring(text)
            index = int(node.get("min")) - 1
            if index < len(self.widths):
                node.set("width", str(self.widths[index]))
            text = ET.tostring(node, encoding="unicode")

        def row(match):
            number = int(match[1])
            if number > self.last:
                self.note_rows.append(number)
            return f'<row r="{number}">'

        text = re.sub(r'<row r="(\d+)">', row, text)

        def cell(match):
            attributes = match[1]
            reference = re.search(
                r'\br="[A-Z]+(\d+)"', attributes
            )
            if reference and int(reference[1]) != 1:
                attributes = re.sub(
                    r'\s+s="\d+"', '', attributes
                )
            return '<c' + attributes + '>'

        text = re.sub(r'<c(\s[^>]+)>', cell, text)

        # IN_NO_MERGE_V1


        return self.stream.write(text.encode("utf-8"))


def write_output(
    context, connection, cutoff, metrics
):
    summary_rows = _in01_prepare(context, connection, cutoff, metrics)
    output = (
        Path(context["output_folder"])
        / f"LBR_Results_IN_{cutoff:%Y%m%d}.xlsx"
    )

    output.parent.mkdir(
        parents=True, exist_ok=True
    )

    with tempfile.TemporaryDirectory(
        prefix="in01_output_",
        dir=output.parent,
    ) as temporary:
        temporary = Path(temporary)

        if output.exists():
            base = output
        else:
            base = temporary / "base.xlsx"
            blank = Workbook(write_only=True)

            for name in SHEETS:
                blank.create_sheet(name)

            blank.save(base)

        new = temporary / "result.xlsx"

        with zipfile.ZipFile(base) as original:
            workbook = ET.fromstring(
                original.read("xl/workbook.xml")
            )
            relationships = ET.fromstring(
                original.read(
                    "xl/_rels/workbook.xml.rels"
                )
            )
            content_types = ET.fromstring(
                original.read("[Content_Types].xml")
            )
            styles = ET.fromstring(
                original.read("xl/styles.xml")
            )

            style_ids = _styles(styles)
            for style_id in style_ids:
                styles.find(_tag("cellXfs"))[style_id].find(
                    _tag("alignment")
                ).set("wrapText", "0")
            sheets = workbook.find(_tag("sheets"))
            paths = {}

            legacy = next(
                (s for s in sheets if s.get("name") == "IN01_Valuation"),
                None,
            )
            if legacy is not None:
                if any(s.get("name") == "IN01_Detail" for s in sheets):
                    raise ValueError(
                        "Existen IN01_Valuation e IN01_Detail "
                        "simultaneamente. Revisar antes de migrar; "
                        "reporte no reemplazado."
                    )
                legacy.set("name", "IN01_Detail")
            for name in SHEETS:
                existing = next((
                    sheet for sheet in sheets
                    if sheet.get("name") == name
                ), None)

                if existing is not None:
                    relation = next(
                        relation
                        for relation in relationships
                        if relation.get("Id") == existing.get(
                            "{" + REL + "}id"
                        )
                    )

                    target = relation.get("Target")

                    paths[name] = (
                        target.lstrip("/")
                        if target.startswith("/")
                        else "xl/" + target
                    )

                else:
                    index = 1

                    while (
                        f"xl/worksheets/in01_{index}.xml"
                        in original.namelist()
                        or
                        f"xl/worksheets/in01_{index}.xml"
                        in paths.values()
                    ):
                        index += 1

                    path = (
                        f"xl/worksheets/in01_{index}.xml"
                    )

                    relation_id = 1
                    while any(
                        relation.get("Id")
                        == f"rId{relation_id}"
                        for relation in relationships
                    ):
                        relation_id += 1

                    sheet_id = max(
                        [0] + [
                            int(sheet.get("sheetId"))
                            for sheet in sheets
                        ]
                    ) + 1

                    ET.SubElement(
                        sheets,
                        _tag("sheet"),
                        {
                            "name": name,
                            "sheetId": str(sheet_id),
                            "{" + REL + "}id":
                                f"rId{relation_id}",
                        },
                    )

                    ET.SubElement(
                        relationships,
                        "{" + PKG + "}Relationship",
                        {
                            "Id": f"rId{relation_id}",
                            "Type": REL + "/worksheet",
                            "Target": path[3:],
                        },
                    )

                    ET.SubElement(
                        content_types,
                        "{" + CT + "}Override",
                        {
                            "PartName": "/" + path,
                            "ContentType":
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                        },
                    )

                    paths[name] = path

            replacements = {
                "xl/workbook.xml":
                    _xml_preserve_namespaces(
                        workbook,
                        original.read(
                            "xl/workbook.xml"
                        ),
                    ),
                "xl/_rels/workbook.xml.rels":
                    _xml_preserve_namespaces(
                        relationships,
                        original.read(
                            "xl/_rels/workbook.xml.rels"
                        ),
                    ),
                "[Content_Types].xml":
                    _xml_preserve_namespaces(
                        content_types,
                        original.read(
                            "[Content_Types].xml"
                        ),
                    ),
                "xl/styles.xml":
                    _xml_preserve_namespaces(
                        styles,
                        original.read(
                            "xl/styles.xml"
                        ),
                    ),
            }

            scope_notes = _in01_notes(context, cutoff, metrics)

            def notes():
                return scope_notes

            with zipfile.ZipFile(
                new,
                "w",
                compression=zipfile.ZIP_DEFLATED,
                compresslevel=3,
                allowZip64=True,
            ) as destination:

                for info in original.infolist():
                    if (
                        info.filename in replacements
                        or info.filename in paths.values()
                    ):
                        continue

                    with original.open(info) as source:
                        with destination.open(
                            info,
                            "w",
                            force_zip64=True,
                        ) as target:
                            while True:
                                chunk = source.read(
                                    1024 * 1024
                                )
                                if not chunk:
                                    break
                                target.write(chunk)

                for path, content in replacements.items():
                    destination.writestr(
                        path, content
                    )

                for name, headings, rows, visible in (
                    (
                        "IN01",
                        SUMMARY_HEADERS,
                        iter(summary_rows),
                        len(SUMMARY_HEADERS),
                    ),
                    (
                        "IN01_Detail",
                        DETAIL_VISIBLE + DETAIL_TECH,
                        _in01_details(connection),
                        len(DETAIL_VISIBLE),
                    ),
                ):
                    with destination.open(
                        paths[name], "w", force_zip64=True
                    ) as stream:
                        widths = (
                            [12, 30, 20, 16, 16, 18, 34, 44, 16]
                            if name == "IN01"
                            else [
                                46, 12, 28, 30, 16, 18, 18,
                                24, 18, 14, 20, 14, 22,
                            ]
                        )
                        records = (
                            len(summary_rows)
                            if name == "IN01"
                            else metrics["IN01 detail evidence rows"]
                        )
                        layout = _IN01Layout(
                            stream, widths, visible,
                            records, style_ids[2],
                        )
                        detail_notes = lambda: [
                            "See IN01 for scope, inputs, data gaps "
                            "and evaluated counts. Evidence rows "
                            "are not audit exception counts."
                        ]
                        write_sheet(
                            layout, headings, rows, visible,
                            style_ids,
                            notes if name == "IN01" else detail_notes,
                        )

        os.replace(new, output)

    return output


def run_in_001(context):
    cutoff = cutoff_from(context['module'])
    companies = companies_from(context['module'].get('companies'))
    sources = find_sources(context['input_folder'], cutoff)
    metrics = Counter()
    metrics['_project_cutoff'] = cutoff.strftime('%Y%m%d')
    metrics['_input'] = str(Path(context['input_folder']))
    with tempfile.TemporaryDirectory(prefix='lbr_in01_') as temporary:
        connection = database(str(Path(temporary) / 'working.sqlite'))
        try:
            for source in ('MARD', 'MARDH', 'MBEW', 'MBEWH', 'QBEW', 'QBEWH'):
                metrics[source + ' files'] = len(sources[source])
                load_source(connection, source, sources[source], companies, metrics)
            connection.execute('CREATE INDEX snapshot_family_key ON snapshot(family,k,source,p)')
            connection.commit()
            output = write_output(context, connection, cutoff, metrics)
        finally:
            connection.close()
    print('IN01 - PRELIMINARY - supplied tables only', flush=True)
    for key, value in sorted(metrics.items()):
        if not key.startswith('_'):
            print(f'{key}: {value}', flush=True)
    print('IN01 output:', output, flush=True)
    print('Runner OK does not certify complete SAP inventory coverage.', flush=True)
    return output

