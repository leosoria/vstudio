"""SD06: account creator also posts journals. Reconstructs supporting positions from SD inputs. SAP completeness is not independently certified."""
from pathlib import Path
import csv
import math
import re
import os
import tempfile
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from core.sd_common import normalize_text, normalize_identifier, normalize_company, normalize_lookup, parse_sap_date, parse_sap_amount, parse_config_date, parse_config_companies, get_period_suffix, write_control_sheet
from core.sd_common import enrich_sd_names
DOC = ['Company', 'Fiscal Year', 'Document Number']
KEY = DOC + ['Document Item']
VISIBLE = ['Company', 'Document Number', 'Fiscal Year', 'Document Item', 'Trigger', 'GL Account', 'GL Account Name', 'Vendor Code', 'Vendor Name', 'Account Created By', 'Account Created By Name', 'Account Created On', 'Journal Posted By', 'Journal Posted By Name', 'Entry Date', 'Debit Amount', 'Credit Amount', 'Document Currency', 'Item Text']
EXTRA = ['Journal Key', 'Chart of Accounts', 'Account Type', 'Posting Key', 'Customer Code', 'Amount in Local Currency', 'Local Currency', 'Document Type', 'Transaction Code', 'Debit Credit', 'Amount in Document Currency', 'Journal Source File', 'Journal Source Row', 'GL Source File', 'GL Source Row', 'Detail Sources', 'Detail Validation']

def ident(value, width):
    text = normalize_identifier(value).upper()
    return text.zfill(width) if text.isdigit() else text

def columns(headers):
    names = [normalize_lookup(re.sub('\\.\\d+$', '', normalize_text(value))) for value in headers]

    def find(label, occurrence=1):
        positions = [i for i, value in enumerate(names) if value == normalize_lookup(label)]
        if len(positions) < occurrence:
            raise ValueError(f'SD06: missing column {label}, occurrence {occurrence}')
        return positions[occurrence - 1]
    return find

def master(context):
    suffix = get_period_suffix(context['module'])
    files = [p for p in Path(context['input_folder']).rglob('*') if p.is_file() and p.suffix.lower() in {'.xlsx', '.xlsm'} and (normalize_lookup(p.stem) == normalize_lookup('LBR SD ACCOUNTS ' + suffix))]
    if len(files) != 1:
        raise ValueError(f'SD06: expected one SD_ACCOUNTS file; found {len(files)}')
    raw = pd.read_excel(files[0], sheet_name='Sheet1', header=None, dtype=object)
    find = columns(raw.iloc[0])
    positions = {name: find(label, occurrence) for name, label, occurrence in [('Company', 'CoCd', 1), ('Chart of Accounts', 'ChAc', 1), ('Account Chart', 'ChAc', 2), ('GL Account', 'G/L Acct', 1), ('Account Created On', 'Date', 1), ('Account Created By', 'Created by', 1), ('Local Currency', 'Crcy', 1), ('GL Account Name', 'Long Text', 1)]}
    result = {}
    company_chart = {}
    companies = parse_config_companies(context['module']['companies'])
    for _, row in raw.iloc[1:].iterrows():
        rec = {name: normalize_text(row.iloc[index]) for name, index in positions.items()}
        if not rec['Company'] and (not rec['GL Account']):
            if any((rec[name] for name in ['Chart of Accounts', 'Account Created By', 'Account Created On'])):
                raise ValueError('SD06: account master row without company/account keys')
            continue
        rec['Company'] = normalize_company(rec['Company'])
        rec['GL Account'] = ident(rec['GL Account'], 10)
        if not re.fullmatch('\\d{4}', rec['Company']) or not re.fullmatch('\\d{10}', rec['GL Account']):
            raise ValueError('SD06: invalid account master keys')
        if companies is not None and rec['Company'] not in companies:
            continue
        if not rec['Chart of Accounts'] or rec['Chart of Accounts'] != rec.pop('Account Chart'):
            raise ValueError('SD06: company and account charts do not agree')
        company = rec['Company']
        chart = rec['Chart of Accounts']
        if company in company_chart and company_chart[company] != chart:
            raise ValueError('SD06: more than one chart for a company')
        company_chart[company] = chart
        rec['Account Created By'] = rec['Account Created By'].upper()
        rec['Account Created On'] = parse_sap_date(rec['Account Created On'])
        rec['Local Currency'] = rec['Local Currency'].upper()
        key = (company, rec['GL Account'])
        if key in result:
            for field in ['Chart of Accounts', 'Account Created By', 'Account Created On', 'Local Currency']:
                a, b = (result[key][field], rec[field])
                if not (pd.isna(a) and pd.isna(b)) and a != b:
                    raise ValueError(f'SD06: contradictory account master evidence for {key}')
        else:
            result[key] = rec
    if not result:
        raise ValueError('SD06: empty account master population')
    return result

def detect(context, accounts):
    files = sorted((p for p in Path(context['input_folder']).rglob('*') if p.is_file() and p.suffix.lower() in {'.xlsx', '.xlsm'} and re.fullmatch('lbrsdentries(open|cleared)\\d+', normalize_lookup(p.stem))))
    if not files:
        raise ValueError('SD06: no SD_ENTRIES_OPEN/CLEARED input files')
    start = parse_config_date(context['module']['from'], 'FROM')
    end = parse_config_date(context['module']['to'], 'TO')
    companies = parse_config_companies(context['module']['companies'])
    stats = dict(files=len(files), scoped_positions=0, missing_master=0, missing_creation_evidence=0, missing_poster=0)
    candidates = {}
    for path in files:
        book = _sd06_read_input(path, read_only=True, data_only=False)
        count = 0
        try:
            rows = book['Sheet1'].iter_rows(values_only=True)
            find = columns(next(rows))
            indices = {name: find(label) for name, label in {'Company': 'CoCd', 'Fiscal Year': 'Year', 'Document Number': 'DocumentNo', 'Document Item': 'Itm', 'GL Account': 'G/L', 'Journal Posted By': 'User name', 'Entry Date': 'Entry Dte', 'Document Currency': 'Crcy', 'Document Type': 'Doc. Type', 'Transaction Code': 'TCode', 'Debit Credit': 'D/C', 'Amount in Document Currency': 'Amount', 'Amount in Local Currency': 'Amount in LC'}.items()}
            for rownum, row in enumerate(rows, 2):
                rec = {name: row[index] for name, index in indices.items()}
                if all((not normalize_text(rec[name]) for name in KEY)):
                    if any((normalize_text(rec[name]) for name in ['GL Account', 'Journal Posted By', 'Entry Date', 'Debit Credit'])):
                        raise ValueError(f'{path.name}: incomplete keys at row {rownum}')
                    continue
                for name, width in zip(KEY + ['GL Account'], [4, 4, 10, 3, 10]):
                    rec[name] = ident(rec[name], width)
                    if not rec[name].isdigit():
                        raise ValueError(f'{path.name}: invalid {name}, row {rownum}')
                date = parse_sap_date(rec['Entry Date'])
                if pd.isna(date):
                    raise ValueError(f'{path.name}: invalid Entry Dte, row {rownum}')
                if not start <= date <= end or (companies is not None and rec['Company'] not in companies):
                    continue
                stats['scoped_positions'] += 1
                count += 1
                account = accounts.get((rec['Company'], rec['GL Account']))
                if account is None:
                    stats['missing_master'] += 1
                    continue
                poster = normalize_text(rec['Journal Posted By']).upper()
                if not poster:
                    stats['missing_poster'] += 1
                    continue
                if not account['Account Created By'] or pd.isna(account['Account Created On']):
                    stats['missing_creation_evidence'] += 1
                    continue
                if poster != account['Account Created By'] or account['Account Created On'] > date:
                    continue
                rec.update({'Entry Date': date, 'Journal Posted By': poster, 'GL Source File': path.name, 'GL Source Row': rownum})
                for name in ['Document Currency', 'Document Type', 'Transaction Code', 'Debit Credit']:
                    rec[name] = normalize_text(rec[name]).upper()
                for name in ['Amount in Document Currency', 'Amount in Local Currency']:
                    rec[name] = parse_sap_amount(rec[name])
                    if not math.isfinite(rec[name]):
                        raise ValueError(f'{path.name}: invalid amount, row {rownum}')
                if rec['Debit Credit'] not in {'S', 'H'} or not rec['Document Currency']:
                    raise ValueError(f'{path.name}: invalid D/C or document currency')
                key = tuple((rec[name] for name in KEY))
                if key in candidates:
                    business = set(rec) - {'GL Source File', 'GL Source Row'}
                    if any((rec[name] != candidates[key][name] for name in business)):
                        raise ValueError(f'SD06: conflicting duplicate GL position {key}')
                candidates[key] = rec
        finally:
            book.close()
        print(f'SD_006 - {path.name}: {count} scoped positions', flush=True)
    if not stats['scoped_positions']:
        raise ValueError('SD06: no GL positions within CONFIG scope')
    return (candidates, stats)

def details(context, documents):
    files = sorted((p for p in Path(context['input_folder']).rglob('*') if p.is_file() and (not p.name.startswith('~$')) and (p.suffix.lower() in {'.xlsx', '.xlsm'}) and normalize_lookup(p.stem).startswith(('lbrsdjournal', 'lbrsddetail'))))
    result = {}
    schema = [('Company', 'CoCd', 1), ('Fiscal Year', 'Year', 1), ('Document Number', 'DocumentNo', 1), ('Document Item', 'Itm', 1), ('Account Type', 'AccTy', 1), ('Posting Key', 'PK', 1), ('GL Account', 'G/L', 1), ('Vendor Code', 'Vendor', 1), ('Customer Code', 'Customer', 1), ('Debit Credit', 'D/C', 1), ('Amount in Document Currency', 'Amount', 1), ('Amount in Local Currency', 'Amount in LC', 1), ('Item Text', 'Text', 1), ('BSEG Document Currency', 'Crcy', 1), ('BSEG Local Currency', 'Crcy', 2)]
    for path in files:
        book = _sd06_read_input(path, read_only=True, data_only=False)
        count = 0
        try:
            rows = book['Sheet1'].iter_rows(values_only=True)
            find = columns(next(rows))
            indices = {name: find(label, occurrence) for name, label, occurrence in schema}
            for rownum, row in enumerate(rows, 2):
                doc = tuple((ident(row[indices[name]], width) for name, width in zip(DOC, [4, 4, 10])))
                if doc not in documents:
                    continue
                rec = {name: row[index] for name, index in indices.items()}
                rec.update(dict(zip(DOC, doc)))
                for name, width in [('Document Item', 3), ('GL Account', 10), ('Posting Key', 2), ('Vendor Code', 10), ('Customer Code', 10)]:
                    rec[name] = ident(rec[name], width)
                if not rec['Document Item'].isdigit():
                    raise ValueError(f'{path.name}: incomplete BSEG item at row {rownum}')
                for name in ['Account Type', 'Debit Credit', 'Item Text', 'BSEG Document Currency', 'BSEG Local Currency']:
                    rec[name] = normalize_text(rec[name])
                for name in ['Account Type', 'Debit Credit', 'BSEG Document Currency', 'BSEG Local Currency']:
                    rec[name] = rec[name].upper()
                if not rec['GL Account'].isdigit() or not rec['Posting Key'].isdigit() or rec['Account Type'] not in {'S', 'K', 'D', 'A', 'M'}:
                    raise ValueError(f'{path.name}: invalid BSEG account/posting key/type')
                if rec['Account Type'] == 'K' and (not rec['Vendor Code']) or (rec['Account Type'] == 'D' and (not rec['Customer Code'])):
                    raise ValueError(f'{path.name}: BSEG party identifier missing')
                if rec['Debit Credit'] not in {'S', 'H'}:
                    raise ValueError(f'{path.name}: invalid BSEG D/C')
                for name in ['Amount in Document Currency', 'Amount in Local Currency']:
                    rec[name] = parse_sap_amount(rec[name])
                    if not math.isfinite(rec[name]):
                        raise ValueError(f'{path.name}: invalid BSEG amount')
                key = tuple((rec[name] for name in KEY))
                if key in result and any((result[key][name] != rec[name] for name in rec)):
                    raise ValueError(f'SD06: contradictory BSEG position {key}')
                rec['Journal Source File'] = path.name
                rec['Journal Source Row'] = rownum
                result[key] = rec
                count += 1
        finally:
            book.close()
        print(f'SD_006 - BSEG {path.name}: {count} linked positions', flush=True)
    return result

def _sd06_collect_available(context, documents, records, headers, bseg=False):
    files = []
    for path in Path(context['input_folder']).rglob('*'):
        if not path.is_file() or path.name.startswith('~$') or path.suffix.lower() not in {'.xlsx', '.xlsm'}:
            continue
        name = normalize_lookup(path.stem)
        if bseg:
            match = name.startswith(('lbrsdjournal', 'lbrsddetail'))
        else:
            match = bool(re.fullmatch('lbrsdentries(open|cleared)\\d+', name)) or name.startswith(('lbrsdapopen', 'lbrsdapcleared', 'lbrsdapclrd', 'lbrsdaropen', 'lbrsdarcleared'))
        if match:
            files.append(path)
    for path in sorted(files):
        book = None
        count = 0
        try:
            book = _sd06_read_input(path, read_only=True, data_only=False)
            rows = book['Sheet1'].iter_rows(values_only=True)
            labels = [normalize_lookup(re.sub('\\.\\d+$', '', normalize_text(value))) for value in next(rows)]

            def col(label, required=True):
                found = [i for i, value in enumerate(labels) if value == normalize_lookup(label)]
                if not found and required:
                    raise ValueError('SD06: missing column ' + label + ' in ' + path.name)
                return found[0] if found else None
            mandatory = {name: col(label) for name, label in {'Company': 'CoCd', 'Fiscal Year': 'Year', 'Document Number': 'DocumentNo', 'Document Item': 'Itm', 'Debit Credit': 'D/C', 'Amount in Document Currency': 'Amount', 'Amount in Local Currency': 'Amount in LC'}.items()}
            optional = {name: col(label, False) for name, label in {'GL Account': 'G/L', 'Account Type': 'AccTy', 'Posting Key': 'PK', 'Vendor Code': 'Vendor', 'Customer Code': 'Customer', 'Item Text': 'Text', 'Document Currency': 'Crcy', 'Journal Posted By': 'User name', 'Entry Date': 'Entry Dte', 'Document Type': 'Doc. Type', 'Transaction Code': 'TCode'}.items()}
            for rownum, row in enumerate(rows, 2):
                doc = tuple((ident(row[mandatory[name]], width) for name, width in zip(DOC, [4, 4, 10])))
                if doc not in documents:
                    continue
                rec = dict(zip(DOC, doc))
                rec['Document Item'] = ident(row[mandatory['Document Item']], 3)
                if not rec['Document Item'].isdigit():
                    raise ValueError(f'{path.name}: invalid item at row {rownum}')
                rec['Debit Credit'] = normalize_text(row[mandatory['Debit Credit']]).upper()
                for name in ['Amount in Document Currency', 'Amount in Local Currency']:
                    rec[name] = parse_sap_amount(row[mandatory[name]])
                    if not math.isfinite(rec[name]):
                        raise ValueError(f'{path.name}: invalid amount at row {rownum}')
                if rec['Debit Credit'] not in {'S', 'H'}:
                    raise ValueError(f'{path.name}: invalid D/C at row {rownum}')
                for name, index in optional.items():
                    rec[name] = normalize_text(row[index]) if index is not None else ''
                for name, width in [('GL Account', 10), ('Posting Key', 2), ('Vendor Code', 10), ('Customer Code', 10)]:
                    rec[name] = ident(rec[name], width)
                for name in ['Account Type', 'Document Currency', 'Journal Posted By', 'Document Type', 'Transaction Code']:
                    rec[name] = rec[name].upper()
                if rec['Entry Date']:
                    rec['Entry Date'] = parse_sap_date(row[optional['Entry Date']])
                    if pd.isna(rec['Entry Date']):
                        raise ValueError(f'{path.name}: invalid Entry Dte at row {rownum}')
                for name in ['Document Currency', 'Journal Posted By', 'Entry Date', 'Document Type', 'Transaction Code']:
                    if normalize_text(rec[name]) and rec[name] != headers[doc][name]:
                        raise ValueError(f'SD06: inconsistent {name} in {path.name}, row {rownum}')
                if rec['GL Account'] and (not rec['GL Account'].isdigit()):
                    raise ValueError(f'{path.name}: invalid G/L account at row {rownum}')
                key = doc + (rec['Document Item'],)
                location = f'{path.name}:row {rownum}'
                if key in records:
                    old = records[key]
                    if old['Debit Credit'] != rec['Debit Credit'] or any((abs(old[name] - rec[name]) > 0.01 for name in ['Amount in Document Currency', 'Amount in Local Currency'])):
                        raise ValueError(f'SD06: contradictory amounts or D/C for {key}')
                    for name in ['GL Account', 'Account Type', 'Posting Key', 'Vendor Code', 'Customer Code']:
                        if old[name] and rec[name] and (old[name] != rec[name]):
                            raise ValueError(f'SD06: contradictory {name} for {key}')
                        if not old[name]:
                            old[name] = rec[name]
                    if not old['Item Text']:
                        old['Item Text'] = rec['Item Text']
                    old['Detail Sources'] += '; ' + location
                else:
                    rec['Journal Source File'] = path.name
                    rec['Journal Source Row'] = rownum
                    rec['Detail Sources'] = location
                    records[key] = rec
                count += 1
        finally:
            if book is not None:
                book.close()
        print(f'SD_006 - Reconstruction {path.name}: {count} linked positions', flush=True)

def _sd06_unresolved(candidates, lines, documents):
    missing = set()
    balances = {}
    counts = {}
    for key, source in candidates.items():
        line = lines.get(key)
        if line is None or any((line[name] != source[name] for name in ['GL Account', 'Debit Credit'])):
            missing.add(key[:3])
            continue
        if any((abs(line[name] - source[name]) > 0.01 for name in ['Amount in Document Currency', 'Amount in Local Currency'])):
            missing.add(key[:3])
    for key, line in lines.items():
        if not line['GL Account']:
            missing.add(key[:3])
        balance = balances.setdefault(key[:3], [0.0, 0.0])
        counts[key[:3]] = counts.get(key[:3], 0) + 1
        sign = 1 if line['Debit Credit'] == 'S' else -1
        balance[0] += sign * line['Amount in Document Currency']
        balance[1] += sign * line['Amount in Local Currency']
    for doc in documents:
        if counts.get(doc, 0) < 2 or any((abs(value) > 0.01 for value in balances.get(doc, [float('inf')]))):
            missing.add(doc)
    return missing

def complete(context, candidates):
    documents = {key[:3] for key in candidates}
    headers = {}
    lines = {}
    for key, source in candidates.items():
        header = {name: source[name] for name in ['Journal Posted By', 'Entry Date', 'Document Currency', 'Document Type', 'Transaction Code']}
        if key[:3] in headers and headers[key[:3]] != header:
            raise ValueError(f'SD06: inconsistent candidate document header {key[:3]}')
        headers[key[:3]] = header
    _sd06_collect_available(context, documents, lines, headers)
    unresolved = _sd06_unresolved(candidates, lines, documents)
    if unresolved:
        _sd06_collect_available(context, unresolved, lines, headers, bseg=True)
        unresolved = _sd06_unresolved(candidates, lines, documents)
    if unresolved:
        examples = ', '.join(('|'.join(doc) for doc in sorted(unresolved)[:5]))
        raise ValueError(f'SD06: {len(unresolved)} candidate journals remain inconsistent or lack required detail after automatic reconstruction. Examples: {examples}. No partial result published. Input coverage must be completed or extraction automated.')
    for line in lines.values():
        line['Detail Validation'] = 'RECONSTRUCTED_CONSISTENT; SAP completeness not independently certified'
    print(f'SD_006 - Automatically reconstructed journals: {len(documents)}', flush=True)
    print(f'SD_006 - Reconstructed positions: {len(lines)}', flush=True)
    print('SD_006 - Validation: candidate positions reconciled, duplicates checked, balances verified. SAP completeness is not independently certified.', flush=True)
    return (lines, headers)

def present(path, output, stats):
    book = load_workbook(path)
    temporary = None
    try:
        sheet = book['SD06']
        if output.empty:
            sheet.merge_cells('A3:O3')
            sheet['A3'] = 'SD06 - No exceptions within supplied input scope'
            sheet['A3'].font = Font(bold=True)
            sheet.merge_cells('A4:O7')
            sheet['A4'] = f"Evaluated {stats['scoped_positions']:,} G/L positions from the supplied open/cleared item extracts. No account creator matched the journal poster with a valid creation date. Missing master: {stats['missing_master']:,}; missing creation evidence: {stats['missing_creation_evidence']:,}; missing poster: {stats['missing_poster']:,}. SAP-wide completeness is not certified by these extracts."
            sheet['A4'].alignment = Alignment(wrap_text=True, vertical='top')
            for row in range(4, 8):
                sheet.row_dimensions[row].height = 24
            sheet.auto_filter.ref = None
        else:
            for index in range(len(VISIBLE) + 1, len(output.columns) + 1):
                dimension = sheet.column_dimensions[get_column_letter(index)]
                dimension.hidden = True
                dimension.outlineLevel = 1
            sheet.sheet_properties.outlinePr.summaryRight = False
            sheet.sheet_view.showOutlineSymbols = True
            sheet.sheet_format.outlineLevelCol = 1
            sheet.column_dimensions[get_column_letter(len(VISIBLE))].collapsed = True
        with tempfile.NamedTemporaryFile(dir=Path(path).parent, suffix='.xlsx', delete=False) as handle:
            temporary = Path(handle.name)
        book.save(temporary)
        book.close()
        os.replace(temporary, path)
    finally:
        book.close()
        if temporary is not None and temporary.exists():
            temporary.unlink()


# SD06 performance helpers: input reading only.
# Output writing continues using the original openpyxl reader/writer.
from functools import lru_cache as _sd06_lru_cache
from functools import wraps as _sd06_wraps
from time import perf_counter as _sd06_clock
from datetime import date as _sd06_date
from datetime import datetime as _sd06_datetime
from datetime import time as _sd06_time
from itertools import islice as _sd06_islice
from itertools import zip_longest as _sd06_zip_longest
from zipfile import ZipFile as _sd06_ZipFile
from python_calamine import CalamineWorkbook as _sd06_CalamineWorkbook


def _sd06_fast_value(value):
    if isinstance(value, _sd06_date) and not isinstance(
        value, _sd06_datetime
    ):
        return _sd06_datetime.combine(value, _sd06_time())
    if isinstance(value, str) and value == "":
        return None
    return value


class _Sd06FastSheet:
    def __init__(self, sheet):
        self.sheet = sheet

    def iter_rows(self, values_only=False):
        if not values_only:
            raise ValueError("SD06 fast reader supports values_only=True")
        start = self.sheet.start
        end = self.sheet.end
        if start is None or end is None:
            return
        first_row, first_column = start
        width = end[1] + 1

        # Preserve physical Excel row/column coordinates.
        for _ in range(first_row):
            yield (None,) * width

        for values in self.sheet.iter_rows():
            row = [None] * first_column
            row.extend(_sd06_fast_value(value) for value in values)
            if len(row) < width:
                row.extend([None] * (width - len(row)))
            yield tuple(row)


class _Sd06FastBook:
    def __init__(self, path):
        self.path = Path(path)
        self.started = _sd06_clock()
        self.book = _sd06_CalamineWorkbook.from_path(str(path))

    def __getitem__(self, name):
        return _Sd06FastSheet(self.book.get_sheet_by_name(name))

    def close(self):
        self.book.close()
        print(
            f"SD_006 - Input {self.path.name}: "
            f"{_sd06_clock() - self.started:.2f}s "
            "(calamine: read + processing)",
            flush=True,
        )


@_sd06_lru_cache(maxsize=128)
def _sd06_contains_formulas(path_text, size, modified):
    # Calamine reads formula results. Preserve the previous data_only=False
    # behavior by using openpyxl whenever formulas are present.
    pattern = re.compile(
        rb"<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?f(?=[\s/>])"
    )
    with _sd06_ZipFile(path_text) as archive:
        for name in archive.namelist():
            if not (
                name.startswith("xl/worksheets/")
                and name.endswith(".xml")
            ):
                continue
            with archive.open(name) as handle:
                carry = b""
                while True:
                    block = handle.read(1024 * 1024)
                    if not block:
                        break
                    content = carry + block
                    if pattern.search(content):
                        return True
                    carry = content[-256:]
    return False


_sd06_checked_inputs = set()


def _sd06_read_input(path, read_only=True, data_only=False):
    if not read_only or data_only:
        raise ValueError("SD06: unexpected optimized input-reader arguments")

    path = Path(path)
    stat = path.stat()
    signature = (
        str(path.resolve()), stat.st_size, stat.st_mtime_ns
    )

    if _sd06_contains_formulas(*signature):
        print(
            f"SD_006 - {path.name}: formulas detected; "
            "using original openpyxl reader.",
            flush=True,
        )
        return load_workbook(path, read_only=True, data_only=False)

    fast = None
    try:
        fast = _Sd06FastBook(path)

        # Validate a prefix against the original reader once per file/version.
        # A mismatch falls back to the original reader.
        if signature not in _sd06_checked_inputs:
            original = load_workbook(
                path, read_only=True, data_only=False
            )
            try:
                old_rows = list(_sd06_islice(
                    original["Sheet1"].iter_rows(values_only=True), 20
                ))
                new_rows = list(_sd06_islice(
                    fast["Sheet1"].iter_rows(values_only=True), 20
                ))

                if len(old_rows) != len(new_rows):
                    raise ValueError("prefix row count differs")

                for old, new in zip(old_rows, new_rows):
                    for left, right in _sd06_zip_longest(old, new):
                        if _sd06_fast_value(left) != _sd06_fast_value(right):
                            raise ValueError("prefix cell value differs")
            finally:
                original.close()

            _sd06_checked_inputs.add(signature)

        return fast

    except Exception as error:
        if fast is not None:
            fast.close()
        print(
            f"SD_006 - {path.name}: fast-reader check unavailable "
            f"({type(error).__name__}: {error}); "
            "using original openpyxl reader.",
            flush=True,
        )
        return load_workbook(path, read_only=True, data_only=False)


def _sd06_timed(label, function):
    @_sd06_wraps(function)
    def measured(*args, **kwargs):
        started = _sd06_clock()
        try:
            return function(*args, **kwargs)
        finally:
            print(
                f"SD_006 - Stage {label}: "
                f"{_sd06_clock() - started:.2f}s",
                flush=True,
            )
    return measured


# Bounded caches for pure scalar conversions; no persistent data cache.
ident = _sd06_lru_cache(maxsize=16384)(ident)
parse_sap_date = _sd06_lru_cache(maxsize=8192)(parse_sap_date)

master = _sd06_timed("account master", master)
detect = _sd06_timed("candidate detection", detect)
complete = _sd06_timed("journal reconstruction", complete)


def run_sd_006(context):
    accounts = master(context)
    candidates, stats = detect(context, accounts)
    stats['Candidate trigger positions'] = len(candidates)
    stats['Exception journals'] = len({key[:3] for key in candidates})
    for name, value in stats.items():
        print(f'SD_006 - {name}: {value}', flush=True)
    output = []
    if candidates:
        lines, headers = complete(context, candidates)
        for key, line in lines.items():
            rec = {**line, **headers[key[:3]]}
            account = accounts.get((rec['Company'], rec['GL Account']), {})
            for name in ['Chart of Accounts', 'GL Account Name', 'Account Created By', 'Account Created On', 'Local Currency']:
                rec[name] = account.get(name, '')
            source = candidates.get(key, {})
            rec['GL Source File'] = source.get('GL Source File', '')
            rec['GL Source Row'] = source.get('GL Source Row', '')
            rec['Trigger'] = 'YES' if key in candidates else 'NO'
            rec['Journal Key'] = '|'.join(key[:3])
            amount = rec['Amount in Document Currency']
            rec['Debit Amount'] = amount if rec['Debit Credit'] == 'S' else 0.0
            rec['Credit Amount'] = amount if rec['Debit Credit'] == 'H' else 0.0
            output.append(rec)
    frame = pd.DataFrame(output, columns=VISIBLE + EXTRA)
    if not frame.empty:
        frame = frame.sort_values(KEY).reset_index(drop=True)
    frame = enrich_sd_names(
        context, frame,
        user_columns={'Account Created By': 'Account Created By Name', 'Journal Posted By': 'Journal Posted By Name'},
    )
    path = write_control_sheet(context, 'SD06', frame, date_columns=['Account Created On', 'Entry Date'], amount_columns=['Debit Amount', 'Credit Amount', 'Amount in Document Currency', 'Amount in Local Currency'])
    present(path, frame, stats)
    print(f'SD_006 - Output positions: {len(frame)}')
    print(f'SD_006 - output: {path}')
    return {'status': 'OK', 'output_file': path, 'sheet_name': 'SD06', 'rows': len(frame), 'summary': stats}
run = run_sd_006
