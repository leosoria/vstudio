
"""Independent AP infrastructure. No SD/VM dependencies."""
import json
import math
import os
import re
import tempfile
import unicodedata
from collections import Counter
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel

VERSION = 'AP05_V1_20261102'

OPEN = (
    'CoCd|Vendor|DocumentNo|Year|Itm|PK|D/C|Amount|Crcy|'
    'Amount in LC|Crcy|Text|Doc. Type|Doc. Date|Pstng Date|'
    'Entry Dte|Time|User name|TCode|Crcy|Reference|Rev. with|Year'
).split('|')

CLEARED = (
    'CoCd|DocumentNo|Year|Doc. Type|Doc. Date|Pstng Date|Reference|'
    'Ref. Tran.|Ref. Key|Vendor|Clearing|Clrng doc.|Year|DocumentNo|'
    'Pstng Date|Doc. Date|Crcy|Reference|D/C|Amount in LC|Crcy|'
    'Amount|Crcy|G/L Acct|G/L|PayT|Inv. ref.|User name|Profit Ctr|'
    'Entry Dte|TCode|Itm|PK|Rev. with|Year'
).split('|')

JOURNAL = (
    'CoCd|DocumentNo|Year|Itm|AccTy|PK|G/L|Vendor|Customer|'
    'D/C|Amount|Crcy|Amount in LC|Crcy|Text'
).split('|')

MAPS = {
    'OPEN': dict(
        company=0, vendor=1, doc=2, year=3, item=4,
        pk=5, dc=6, amount=7, currency=19, lc=9, local=10,
        type=12, docdate=13, posting=14, entry=15,
        time=16, user=17, tcode=18, reference=20,
        reversal=21, reversalyear=22,
    ),
    'CLEARED': dict(
        company=0, doc=1, year=2, type=3, docdate=4,
        posting=5, reference=6, vendor=9, clearing=10,
        clearingdoc=11, currency=16, dc=18, lc=19,
        local=20, amount=21, user=27, entry=29,
        tcode=30, item=31, pk=32, reversal=33, reversalyear=34,
    ),
}

TOTALS = {
    'OPEN': {7, 8, 9, 10, 19},
    'CLEARED': {16, 19, 20, 21, 22},
    'JOURNAL': {10, 11, 12, 13},
}


def text(value):
    if value is None:
        return ''
    if isinstance(value, float) and math.isnan(value):
        return ''
    return str(value).strip()


def label(value):
    return ''.join(
        c.lower()
        for c in unicodedata.normalize('NFKD', text(value))
        if c.isalnum()
    )


def ident(value, width):
    value = text(value)
    if re.fullmatch(r'\d+\.0', value):
        value = value[:-2]
    return value.zfill(width) if value.isdigit() else value.upper()


def sap_date(value, epoch=None):
    if value is None or text(value) in {'', '0', '00000000'}:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if (
        epoch is not None
        and isinstance(value, (int, float))
        and not isinstance(value, bool)
        and 20000 <= value <= 80000
    ):
        return from_excel(value, epoch).date()
    for fmt in (
        '%Y%m%d', '%d/%m/%Y', '%d.%m.%Y', '%d-%m-%Y',
        '%Y-%m-%d', '%Y-%m-%d %H:%M:%S', '%Y-%m-%dT%H:%M:%S',
    ):
        try:
            return datetime.strptime(text(value), fmt).date()
        except ValueError:
            pass
    raise ValueError('Invalid date: ' + repr(value))


def money(value):
    if value is None or isinstance(value, bool) or text(value) == '':
        raise ValueError('Missing/invalid amount')
    value = text(value).replace('\u00a0', '').replace(' ', '')
    negative = (
        value.endswith('-')
        or (value.startswith('(') and value.endswith(')'))
    )
    if value.endswith('-'):
        value = value[:-1]
    if value.startswith('(') and value.endswith(')'):
        value = value[1:-1]
    if ',' in value and '.' in value:
        if value.rfind(',') > value.rfind('.'):
            value = value.replace('.', '').replace(',', '.')
        else:
            value = value.replace(',', '')
    elif ',' in value:
        value = value.replace(',', '.')
    try:
        result = Decimal(value)
    except InvalidOperation as error:
        raise ValueError('Invalid amount: ' + repr(value)) from error
    if not result.is_finite():
        raise ValueError('Nonfinite amount')
    return -result if negative else result


def settings(context):
    module = context['module']
    control = context['control']

    start = sap_date(module.get('from'))
    end = sap_date(module.get('to'))
    if not start or not end or start > end:
        raise ValueError('Invalid AP CONFIG FROM/TO')

    raw = module.get('companies', '')
    if isinstance(raw, (list, tuple, set)):
        companies = {ident(v, 4) for v in raw if text(v)}
    elif text(raw).upper() in {'ALL', '*', 'TODAS', 'TODOS'}:
        companies = None
    else:
        companies = {
            ident(v, 4)
            for v in re.split(r'[,;|\s]+', text(raw)) if v
        }
    if companies == set():
        raise ValueError('AP COMPANIES is empty')

    raw_options = text(control.get('param1'))
    options = json.loads(raw_options) if raw_options else {}
    allowed = {'types', 'min_days', 'reversals', 'balance_tolerance'}
    if not isinstance(options, dict) or set(options) - allowed:
        raise ValueError(
            'AP05 PARAM1: use JSON with types, min_days, '
            'reversals, balance_tolerance'
        )

    raw_types = options.get('types', 'KR,RE')
    if not isinstance(raw_types, str):
        raise ValueError('AP05 types must be a string, e.g. KR,RE')
    types = {
        value.upper()
        for value in re.split(r'[,;|\s]+', raw_types) if value
    }
    if not types or any(
        not re.fullmatch(r'[A-Z0-9]{2}', value) for value in types
    ):
        raise ValueError('AP05 types: expected document types, e.g. KR,RE')

    days = options.get('min_days', 1)
    if isinstance(days, bool) or not isinstance(days, int) or days < 1:
        raise ValueError('AP05 min_days must be an integer >= 1')

    reversals = text(options.get('reversals', 'KEEP')).upper()
    tolerance = money(options.get('balance_tolerance', '0.01'))
    if reversals not in {'KEEP', 'EXCLUDE_LINKED'} or tolerance < 0:
        raise ValueError('Invalid reversal policy/balance tolerance')

    return dict(
        start=start, end=end, companies=companies, types=types,
        min_days=days, reversals=reversals, tolerance=tolerance,
    )


def discover(context, family, end):
    pattern = re.compile(
        r'lbrap' + family.lower() + end.strftime('%Y%m%d')
        + r'(?:(?:part)?([0-9]+))?$'
    )
    found = []
    for path in Path(context['input_folder']).rglob('*'):
        if (
            path.is_file()
            and not path.name.startswith('~$')
            and path.suffix.lower() == '.xlsx'
        ):
            match = pattern.fullmatch(label(path.stem))
            if match:
                part = None if match[1] is None else int(match[1])
                found.append((part, path))

    if not found:
        raise FileNotFoundError(
            f'Mandatory source missing: '
            f'LBR AP_{family}_{end:%Y%m%d}.xlsx'
        )

    parts = [number for number, path in found]
    if None in parts:
        if len(found) != 1:
            raise ValueError(
                f'{family}: duplicate base or mixed base/parts: {found}'
            )
    elif (
        len(set(parts)) != len(parts)
        or sorted(parts) != list(range(1, max(parts) + 1))
    ):
        raise ValueError(
            f'{family}: duplicate/missing part numbers; expected 1..N'
        )

    if family == 'COMPANY' and (
        len(found) != 1 or parts != [None]
    ):
        raise ValueError('COMPANY requires one base file')

    return [
        path for number, path in sorted(
            found, key=lambda pair: pair[0] or 0
        )
    ]


def rows(path, family):
    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        if family == 'COMPANY':
            if len(workbook.worksheets) != 1:
                raise ValueError('COMPANY must have one sheet')
            sheet = workbook.worksheets[0]
        else:
            if 'Sheet1' not in workbook.sheetnames:
                raise ValueError(f'{path.name}: missing Sheet1')
            sheet = workbook['Sheet1']

        iterator = sheet.iter_rows()
        header = next(iterator, None)
        if header is None:
            raise ValueError(f'{path.name}: empty workbook')

        headings = [text(cell.value) for cell in header]
        while headings and not headings[-1]:
            headings.pop()

        expected = {
            'OPEN': OPEN, 'CLEARED': CLEARED, 'JOURNAL': JOURNAL,
        }.get(family)
        if expected is not None and headings != expected:
            raise ValueError(
                f'{path.name}: changed/ambiguous layout; no output written'
            )

        for row_number, cells in enumerate(iterator, 2):
            if not any(text(cell.value) for cell in cells):
                continue
            if any(cell.data_type == 'f' for cell in cells):
                raise ValueError(
                    f'{path.name}, row {row_number}: formula in source'
                )
            yield (
                headings,
                [cell.value for cell in cells],
                row_number,
                workbook.epoch,
            )
    finally:
        workbook.close()


def trace(path, row_number):
    return f'{Path(path).resolve()}#row={row_number}'


def load_companies(paths):
    result = {}
    for path in paths:
        for headers, values, row_number, epoch in rows(path, 'COMPANY'):
            normalized = [label(value) for value in headers]
            indexes = []
            for aliases in (
                {'cocd', 'bukrs', 'companycode', 't001bukrs'},
                {'crcy', 'currency', 'waers', 't001waers'},
            ):
                matches = [
                    index for index, value in enumerate(normalized)
                    if value in aliases
                ]
                if len(matches) != 1:
                    raise ValueError(
                        'COMPANY: missing/ambiguous company/currency field'
                    )
                indexes.append(matches[0])

            company = ident(values[indexes[0]], 4)
            currency = text(values[indexes[1]]).upper()
            if not company or not re.fullmatch(r'[A-Z0-9]{3,5}', currency):
                raise ValueError(
                    f'COMPANY row {row_number}: invalid company/currency'
                )
            if company in result and result[company][0] != currency:
                raise ValueError(
                    'COMPANY: contradictory currencies for ' + company
                )
            previous = result.get(company, (currency, []))[1]
            result[company] = (
                currency, previous + [trace(path, row_number)]
            )
    if not result:
        raise ValueError('COMPANY: no records')
    return result


def load_ap(paths, config, companies):
    result = {}
    metrics = Counter()

    for family in ('OPEN', 'CLEARED'):
        mapping = MAPS[family]
        for path in paths[family]:
            print('AP - Reading:', path.name, flush=True)
            for headers, values, row_number, epoch in rows(path, family):
                get = lambda name: values[mapping[name]]
                blanks = [
                    not text(get(name))
                    for name in ('company', 'year', 'doc', 'item', 'vendor')
                ]
                if all(blanks) and all(
                    not text(value)
                    for index, value in enumerate(values)
                    if index not in TOTALS[family]
                ):
                    metrics['AP export totals'] += 1
                    continue
                if any(blanks):
                    raise ValueError(
                        f'{path.name}, row {row_number}: incomplete AP key'
                    )

                record = {
                    name: text(get(name)) for name in mapping
                }
                for name, width in (
                    ('company', 4), ('year', 4), ('doc', 10),
                    ('item', 3), ('vendor', 10), ('pk', 2),
                    ('reversal', 10), ('reversalyear', 4),
                ):
                    record[name] = ident(get(name), width)

                for name in ('posting', 'docdate', 'entry'):
                    record[name] = sap_date(get(name), epoch)
                    if record[name] is None:
                        raise ValueError(
                            f'{path.name}, row {row_number}: missing {name}'
                        )

                if family == 'CLEARED':
                    mismatch = (
                        ident(values[12], 4) != record['year']
                        or ident(values[13], 10) != record['doc']
                        or sap_date(values[14], epoch) != record['posting']
                        or sap_date(values[15], epoch) != record['docdate']
                    )
                    if mismatch:
                        raise ValueError(
                            f'{path.name}, row {row_number}: '
                            'BKPF/BSAK key/date mismatch'
                        )

                for name in ('amount', 'lc'):
                    record[name] = money(get(name))
                for name in (
                    'currency', 'local', 'dc', 'type', 'user', 'tcode',
                ):
                    record[name] = record[name].upper()

                if (
                    record['dc'] not in {'S', 'H'}
                    or not record['currency']
                    or not record['type']
                    or not record['user']
                ):
                    raise ValueError(
                        f'{path.name}, row {row_number}: '
                        'missing/invalid D/C, currency, type or user'
                    )

                other_currency = text(
                    values[8] if family == 'OPEN' else values[22]
                ).upper()
                if (
                    other_currency
                    and other_currency != record['currency']
                ):
                    raise ValueError(
                        'AP header/item document currencies disagree'
                    )

                metrics['AP positions read'] += 1
                if (
                    config['companies'] is not None
                    and record['company'] not in config['companies']
                ):
                    metrics['AP outside companies'] += 1
                    continue

                if record['company'] not in companies:
                    raise ValueError(
                        'Missing COMPANY currency for ' + record['company']
                    )
                if (
                    record['local']
                    and record['local'] != companies[record['company']][0]
                ):
                    raise ValueError(
                        'AP local currency disagrees with COMPANY'
                    )
                record['local'] = companies[record['company']][0]

                key = tuple(
                    record[name]
                    for name in ('company', 'year', 'doc', 'item')
                )
                raw = tuple(
                    (type(value).__name__, text(value))
                    for value in values
                )

                if key in result:
                    previous = result[key]
                    if (
                        previous['state'] != family
                        or previous['raw'] != raw
                    ):
                        raise ValueError(
                            f'AP contradictory/OPEN-CLEARED duplicate: {key}'
                        )
                    previous['sources'].append(trace(path, row_number))
                    metrics['AP identical duplicates'] += 1
                else:
                    record.update(
                        state=family, raw=raw,
                        sources=[trace(path, row_number)],
                    )
                    result[key] = record

    if not result:
        raise ValueError(
            'No AP population available for configured companies'
        )
    return result, metrics


def select_candidates(ap, config):
    linked = set()
    for key, record in ap.items():
        if record['reversal']:
            linked.add(key[:3])
            if record['reversalyear']:
                linked.add((
                    key[0], record['reversalyear'], record['reversal'],
                ))

    headers = {}
    candidates = {}
    metrics = Counter()

    for key, record in ap.items():
        doc = key[:3]
        header = tuple(
            record[name]
            for name in (
                'type', 'posting', 'docdate', 'entry', 'currency',
                'local', 'user', 'tcode', 'reference',
                'reversal', 'reversalyear',
            )
        )
        if doc in headers and headers[doc] != header:
            raise ValueError(
                'Conflicting AP header values for ' + repr(doc)
            )
        headers[doc] = header

        if (
            not config['start'] <= record['posting'] <= config['end']
            or record['type'] not in config['types']
        ):
            continue

        metrics['Scoped invoice positions'] += 1
        if (
            record['docdate'] - record['posting']
        ).days < config['min_days']:
            continue

        if (
            config['reversals'] == 'EXCLUDE_LINKED'
            and doc in linked
        ):
            metrics['Candidate positions excluded by reversal links'] += 1
            continue

        candidates[key] = record

    metrics['Exception documents'] = len({
        key[:3] for key in candidates
    })
    metrics['Trigger positions'] = len(candidates)
    return candidates, linked, metrics


def load_journals(paths, candidates, companies, tolerance):
    targets = {key[:3] for key in candidates}
    lines = {}
    metrics = Counter()

    for path in paths:
        print('AP - Reading JOURNAL:', path.name, flush=True)
        for headers, values, row_number, epoch in rows(path, 'JOURNAL'):
            blanks = [
                not text(values[index]) for index in (0, 1, 2, 3)
            ]
            if all(blanks) and all(
                not text(value)
                for index, value in enumerate(values)
                if index not in TOTALS['JOURNAL']
            ):
                metrics['JOURNAL export totals'] += 1
                continue
            if any(blanks):
                raise ValueError(
                    f'{path.name}, row {row_number}: incomplete JOURNAL key'
                )

            metrics['JOURNAL positions scanned'] += 1
            key = (
                ident(values[0], 4), ident(values[2], 4),
                ident(values[1], 10), ident(values[3], 3),
            )
            if key[:3] not in targets:
                continue

            record = dict(
                account_type=text(values[4]).upper(),
                pk=ident(values[5], 2),
                gl=ident(values[6], 10),
                vendor=ident(values[7], 10),
                customer=ident(values[8], 10),
                dc=text(values[9]).upper(),
                amount=money(values[10]),
                currency=text(values[11]).upper(),
                lc=money(values[12]),
                local=text(values[13]).upper(),
                item_text=text(values[14]),
            )
            if (
                record['dc'] not in {'S', 'H'}
                or not record['account_type']
                or not record['pk']
            ):
                raise ValueError(
                    f'{path.name}, row {row_number}: invalid line attributes'
                )
            if (
                record['account_type'] == 'K'
                and not record['vendor']
            ):
                raise ValueError(
                    'JOURNAL vendor position without vendor ID'
                )

            raw = tuple(
                (type(value).__name__, text(value))
                for value in values
            )
            if key in lines:
                if lines[key]['raw'] != raw:
                    raise ValueError(
                        'JOURNAL contradictory duplicate: ' + repr(key)
                    )
                lines[key]['sources'].append(trace(path, row_number))
                metrics['JOURNAL identical duplicates'] += 1
            else:
                record.update(
                    raw=raw, sources=[trace(path, row_number)],
                )
                lines[key] = record

    if not metrics['JOURNAL positions scanned']:
        raise ValueError(
            'Mandatory JOURNAL source contains no keyed positions'
        )

    by_doc = {}
    for key in lines:
        by_doc.setdefault(key[:3], []).append(key)

    missing = targets - by_doc.keys()
    if missing:
        raise ValueError(
            'Candidate journals absent from JOURNAL: '
            + repr(sorted(missing)[:10])
        )

    header_records = {
        key[:3]: record for key, record in candidates.items()
    }

    for doc, keys in by_doc.items():
        header = header_records[doc]
        total_dc = Decimal(0)
        total_lc = Decimal(0)

        for key in keys:
            record = lines[key]
            if (
                record['currency']
                and record['currency'] != header['currency']
            ):
                raise ValueError(
                    'JOURNAL/AP document currency mismatch'
                )
            if (
                record['local']
                and record['local'] != companies[doc[0]][0]
            ):
                raise ValueError(
                    'JOURNAL/COMPANY local currency mismatch'
                )

            record['currency_source'] = (
                'JOURNAL + AP header' if record['currency']
                else 'AP header (JOURNAL blank)'
            )
            record['local_source'] = (
                'JOURNAL + T001' if record['local']
                else 'T001 (JOURNAL blank)'
            )
            record['currency'] = header['currency']
            record['local'] = companies[doc[0]][0]

            sign = 1 if record['dc'] == 'S' else -1
            total_dc += record['amount'] * sign
            total_lc += record['lc'] * sign

        if abs(total_dc) > tolerance or abs(total_lc) > tolerance:
            raise ValueError(
                f'Unbalanced supplied journal {doc}: '
                f'DC={total_dc}, LC={total_lc}'
            )

    for key, expected in candidates.items():
        record = lines.get(key)
        if (
            record is None
            or record['account_type'] != 'K'
            or any(
                record[name] != expected[name]
                for name in ('vendor', 'pk', 'dc', 'amount', 'lc')
            )
        ):
            raise ValueError(
                'AP/JOURNAL trigger missing or inconsistent: ' + repr(key)
            )

    metrics['Validated supplied journals'] = len(by_doc)
    metrics['Output positions'] = len(lines)
    return lines, metrics


VISIBLE = [
    'Company', 'Document Number', 'Fiscal Year', 'Document Item',
    'Trigger', 'Vendor Code', 'Document Type', 'Document Date',
    'Posting Date', 'Days In Future', 'Debit Amount', 'Credit Amount',
    'Document Currency', 'Debit Amount LC', 'Credit Amount LC',
    'Local Currency', 'Reversal Link Status', 'GL Account', 'Item Text',
]

TECHNICAL = [
    'Journal Key', 'Account Type', 'Posting Key', 'Debit Credit',
    'Amount DC', 'Amount LC', 'Entry Date', 'Entry Time', 'Entered By',
    'Transaction Code', 'Reference', 'Reversal Document', 'Reversal Year',
    'Clearing Date', 'Clearing Document', 'AP State',
    'Document Currency Source', 'Local Currency Source',
    'AP Header Sources', 'AP Position Sources',
    'JOURNAL Sources', 'COMPANY Sources',
]


def build_report(candidates, lines, linked, ap, companies):
    headers = {
        key[:3]: record for key, record in candidates.items()
    }
    header_sources = {}
    for key, record in ap.items():
        if key[:3] in headers:
            header_sources.setdefault(key[:3], []).extend(
                record['sources']
            )

    result = []
    for key, record in sorted(lines.items()):
        doc = key[:3]
        header = headers[doc]
        position = ap.get(key)
        output = dict(zip(
            ['Company', 'Fiscal Year', 'Document Number', 'Document Item'],
            key,
        ))

        output.update({
            'Trigger': 'YES' if key in candidates else 'NO',
            'Vendor Code': record['vendor'],
            'Document Type': header['type'],
            'Document Date': header['docdate'],
            'Posting Date': header['posting'],
            'Days In Future': (
                header['docdate'] - header['posting']
            ).days,
            'Debit Amount': (
                record['amount'] if record['dc'] == 'S'
                else Decimal(0)
            ),
            'Credit Amount': (
                record['amount'] if record['dc'] == 'H'
                else Decimal(0)
            ),
            'Document Currency': record['currency'],
            'Debit Amount LC': (
                record['lc'] if record['dc'] == 'S'
                else Decimal(0)
            ),
            'Credit Amount LC': (
                record['lc'] if record['dc'] == 'H'
                else Decimal(0)
            ),
            'Local Currency': record['local'],
            'Reversal Link Status': (
                'REFERENCE_PRESENT' if header['reversal']
                else (
                    'REFERENCED_BY_ANOTHER_DOCUMENT'
                    if doc in linked else 'NO_LINK_OBSERVED'
                )
            ),
            'GL Account': record['gl'],
            'Item Text': record['item_text'],
            'Journal Key': '|'.join(doc),
            'Account Type': record['account_type'],
            'Posting Key': record['pk'],
            'Debit Credit': record['dc'],
            'Amount DC': record['amount'],
            'Amount LC': record['lc'],
            'Entry Date': header['entry'],
            'Entry Time': header.get('time', ''),
            'Entered By': header['user'],
            'Transaction Code': header['tcode'],
            'Reference': header['reference'],
            'Reversal Document': header['reversal'],
            'Reversal Year': header['reversalyear'],
            'Clearing Date': (
                position.get('clearing', '') if position else ''
            ),
            'Clearing Document': (
                position.get('clearingdoc', '') if position else ''
            ),
            'AP State': position['state'] if position else '',
            'Document Currency Source': record['currency_source'],
            'Local Currency Source': record['local_source'],
            'AP Header Sources': ' ; '.join(
                sorted(set(header_sources[doc]))
            ),
            'AP Position Sources': (
                ' ; '.join(position['sources']) if position else ''
            ),
            'JOURNAL Sources': ' ; '.join(record['sources']),
            'COMPANY Sources': ' ; '.join(companies[key[0]][1]),
        })
        result.append(output)
    return result


def write_report(context, records, config, metrics, sheet_name='AP05', visible=None, technical=None, empty_note=None):
    folder = Path(context['output_folder'])
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / f'LBR_Results_AP_{config["end"]:%Y%m%d}.xlsx'
    workbook = load_workbook(path) if path.exists() else Workbook()
    temporary = None

    try:
        if not path.exists():
            workbook.remove(workbook.active)
        if sheet_name in workbook.sheetnames:
            workbook.remove(workbook[sheet_name])

        sheet = workbook.create_sheet(sheet_name)
        columns = (VISIBLE if visible is None else visible) + (TECHNICAL if technical is None else technical)
        sheet.append(columns)

        for record in records:
            if any(len(str(value)) > 32767 for value in record.values()):
                raise ValueError(
                    'Trace/text exceeds Excel cell limit; '
                    'no output replaced'
                )
            sheet.append([
                float(record[name])
                if isinstance(record[name], Decimal)
                else record[name]
                for name in columns
            ])
            for cell in sheet[sheet.max_row]:
                if isinstance(cell.value, str):
                    cell.data_type = 's'

        for cell in sheet[1]:
            cell.fill = PatternFill('solid', fgColor='D9EAF7')
            cell.font = Font(bold=True, color='000000')

        sheet.freeze_panes = 'A2'
        sheet.auto_filter.ref = (
            f'A1:{get_column_letter(len(columns))}'
            f'{max(1, len(records) + 1)}'
        )

        date_columns = {'Document Date', 'Posting Date', 'Entry Date'}
        amount_columns = {
            'Debit Amount', 'Credit Amount',
            'Debit Amount LC', 'Credit Amount LC',
            'Amount DC', 'Amount LC',
        }

        for index, name in enumerate(columns, 1):
            sheet.column_dimensions[
                get_column_letter(index)
            ].width = 40 if name == 'Item Text' else 20

            if name in date_columns or name in amount_columns:
                for row in sheet.iter_rows(
                    min_row=2, max_row=len(records) + 1,
                    min_col=index, max_col=index,
                ):
                    row[0].number_format = (
                        'dd/mm/yyyy' if name in date_columns
                        else '#,##0.00'
                    )

        first = len(VISIBLE if visible is None else visible) + 1
        sheet.column_dimensions.group(
            get_column_letter(first),
            get_column_letter(len(columns)),
            outline_level=1,
            hidden=True,
        )
        sheet.column_dimensions[
            get_column_letter(len(columns) + 1)
        ].collapsed = True

        if not records:
            sheet.cell(2, 1, '')
            note = (
                f'No exceptions in supplied AP inputs for Posting Date '
                f'{config["start"]} to {config["end"]}; '
                f'types={",".join(sorted(config["types"]))}; '
                f'Document Date - Posting Date >= '
                f'{config["min_days"]} days; '
                f'reversal policy={config["reversals"]}. '
                f'Scoped invoice positions='
                f'{metrics["Scoped invoice positions"]}. '
                'This does not certify SAP extraction completeness.'
            )
            sheet.append([note if empty_note is None else empty_note])
            sheet.cell(sheet.max_row, 1).data_type = 's'

        with tempfile.NamedTemporaryFile(
            dir=folder, suffix='.xlsx', delete=False,
        ) as handle:
            temporary = Path(handle.name)

        workbook.save(temporary)
        os.replace(temporary, path)
        temporary = None
    finally:
        workbook.close()
        if temporary is not None and temporary.exists():
            temporary.unlink()

    return path
