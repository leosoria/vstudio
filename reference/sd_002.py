"""SD_002: evaluate complete invoice journals and report every position.

KNA1 general creation is the conflict basis.
KNB1 extension is supporting information.
PARAM1: invoice document types; default DR,RV.
Scope uses AR Entry Date.
Dedicated SD BSEG exports are mandatory.
No GL module dependency.
Coverage and balance checks precede conflict detection.
These checks do not prove that a SAP export was untruncated.
"""
import re
import os
import tempfile
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from core.sd_common import load_sd_customer_master, load_sd_ar_items, load_sd_ar_journals, normalize_text, write_control_sheet
DOC = ['Company', 'Fiscal Year', 'Document Number']
KEY = DOC + ['Document Item']

def get_invoice_types(context):
    text = normalize_text(context.get('control', {}).get('param1', '')).upper()
    if not text:
        return {'DR', 'RV'}
    types = {value for value in re.split('[,;|\\s]+', text) if value}
    if any((not re.fullmatch('[A-Z0-9]{2}', value) for value in types)):
        raise ValueError('SD_002 PARAM1: expected document types, e.g. DR,RV')
    return types

def prepare_ar_items(items):
    """Retained for compatibility with the earlier diagnostic."""
    business = [column for column in items.columns if column not in {'Source File', 'Source Row'}]
    distinct = items.drop_duplicates(business)
    if distinct.duplicated(KEY, keep=False).any():
        raise ValueError('SD_002: conflicting AR records for the same item')
    prepared = items.drop_duplicates(KEY).copy()
    return (prepared, len(items) - len(prepared))

def prepare_customer_master(customers):
    fields = ['Company', 'Customer Code', 'Customer Name', 'Customer Created On', 'Customer Created By', 'Company Customer Created On', 'Company Customer Created By']
    distinct = customers[fields].drop_duplicates()
    if distinct.duplicated(['Company', 'Customer Code']).any():
        raise ValueError('SD_002: conflicting creation evidence in customer master')
    master = customers[fields + ['Source File', 'Source Row']].drop_duplicates(['Company', 'Customer Code']).copy()
    return master.rename(columns={'Source File': 'Customer Source File', 'Source Row': 'Customer Source Row'})

def analyze_full_journals(customers, ar, journals, types):
    header_fields = DOC + ['Document Type', 'Entry Date', 'User ID', 'Transaction Code', 'Document Currency', 'Reversal Document', 'Reversal Year']
    for optional in ['Document Date', 'Posting Date', 'Reference']:
        if optional in ar:
            header_fields.append(optional)
    headers = ar[header_fields].drop_duplicates()
    if headers.duplicated(DOC).any():
        raise ValueError('SD_002: inconsistent AR header values for the same document')
    selected = headers.loc[headers['Document Type'].isin(types)].copy()
    reversed_keys = set()
    linked = ar.loc[ar['Reversal Document'].ne('')]
    reversed_keys.update(linked[DOC].itertuples(index=False, name=None))
    for _, row in linked.iterrows():
        if row['Reversal Year']:
            reversed_keys.add((row['Company'], row['Reversal Year'], row['Reversal Document']))
    reversed_mask = pd.Series([key in reversed_keys for key in selected[DOC].itertuples(index=False, name=None)], index=selected.index, dtype=bool)
    reversed_mask |= selected['Transaction Code'].isin({'VF11', 'FB08', 'F.80'})
    selected = selected.loc[~reversed_mask].copy()
    ar_lines = ar.merge(selected[DOC], on=DOC, how='inner', validate='many_to_one')
    lines = journals.merge(selected, on=DOC, how='inner', validate='many_to_one')
    coverage = selected[DOC].merge(lines[DOC].drop_duplicates(), on=DOC, how='left', indicator=True)
    missing = coverage.loc[coverage['_merge'].eq('left_only'), DOC]
    if not missing.empty:
        raise ValueError(f"SD_002: {len(missing)} scoped journals missing from SD JOURNAL input. Examples: {missing.head(5).to_dict('records')}")
    comparison = ar_lines.merge(lines, on=KEY, how='left', suffixes=('_AR', '_BSEG'), indicator=True, validate='many_to_one')
    bad = comparison['_merge'].ne('both')
    for name in ['Customer Code', 'Posting Key', 'Debit Credit']:
        bad |= comparison[name + '_AR'].ne(comparison[name + '_BSEG'])
    for name in ['Amount in Document Currency', 'Amount in Local Currency']:
        bad |= (comparison[name + '_AR'] - comparison[name + '_BSEG']).abs().gt(0.01)
    if bad.any():
        examples = comparison.loc[bad, KEY].head(5)
        raise ValueError(f"SD_002: AR/BSEG positions missing or inconsistent: {examples.to_dict('records')}")
    if 'BSEG Document Currency' in lines:
        mismatch = lines['BSEG Document Currency'].ne('') & lines['BSEG Document Currency'].ne(lines['Document Currency'])
        if mismatch.any():
            raise ValueError('SD_002: BSEG/AR document currencies disagree')
    sign = lines['Debit Credit'].map({'S': 1, 'H': -1})
    check = lines[DOC].copy()
    check['LC'] = lines['Amount in Local Currency'] * sign
    check['DC'] = lines['Amount in Document Currency'] * sign
    balances = check.groupby(DOC)[['LC', 'DC']].sum()
    bad_balances = balances.loc[balances.abs().gt(0.01).any(axis=1)]
    if not bad_balances.empty:
        examples = bad_balances.head(5).reset_index().to_dict('records')
        raise ValueError(f'SD_002: {len(bad_balances)} scoped journals do not balance; review complete extraction. Examples: {examples}')
    trigger = lines.loc[lines['Account Type'].eq('D') & lines['Posting Key'].eq('01') & lines['Debit Credit'].eq('S')].copy()
    master = prepare_customer_master(customers)
    evaluated = trigger.merge(master, on=['Company', 'Customer Code'], how='left', validate='many_to_one', indicator=True)
    matched = evaluated['_merge'].eq('both')
    evidence = matched & evaluated['Customer Created On'].notna() & evaluated['Customer Created By'].notna()
    evidence &= evaluated['Customer Created By'].ne('') & evaluated['User ID'].ne('')
    conflict = evidence & evaluated['Customer Created By'].eq(evaluated['User ID'])
    conflict &= evaluated['Customer Created On'].le(evaluated['Entry Date'])
    conflicts = evaluated.loc[conflict].copy()
    affected = conflicts[DOC].drop_duplicates()
    output = lines.merge(affected, on=DOC, how='inner', validate='many_to_one')
    output['Journal Key'] = output['Company'].astype(str) + '|' + output['Fiscal Year'].astype(str) + '|' + output['Document Number'].astype(str)
    trigger_keys = set(conflicts[KEY].itertuples(index=False, name=None))
    output['Trigger'] = ['YES' if key in trigger_keys else 'NO' for key in output[KEY].itertuples(index=False, name=None)]
    output = output.merge(master, on=['Company', 'Customer Code'], how='left', validate='many_to_one')
    details = conflicts.groupby(DOC).agg(**{'Conflict Customer Codes': ('Customer Code', lambda values: ', '.join(sorted(set(values)))), 'Conflict User ID': ('User ID', lambda values: ', '.join(sorted(set(values))))}).reset_index()
    output = output.merge(details, on=DOC, how='left', validate='many_to_one')
    output['Same Day Creation And Invoice'] = ''
    mask = output['Trigger'].eq('YES')
    output.loc[mask, 'Same Day Creation And Invoice'] = output.loc[mask, 'Customer Created On'].eq(output.loc[mask, 'Entry Date']).map({True: 'YES', False: 'NO'})
    preferred = ['Journal Key', 'Trigger', 'Company', 'Document Number', 'Fiscal Year', 'Document Item', 'Conflict User ID', 'Conflict Customer Codes', 'Account Type', 'GL Account', 'Customer Code', 'Customer Name', 'Customer Created By', 'Customer Created On', 'User ID', 'Entry Date', 'Posting Key', 'Debit Credit', 'Amount in Document Currency', 'Document Currency', 'Amount in Local Currency', 'Local Currency', 'Item Text']
    columns = [name for name in preferred if name in output]
    output = output[columns + [name for name in output if name not in columns]].sort_values(KEY).reset_index(drop=True)
    metrics = {'Invoice document types': ','.join(sorted(types)), 'AR header documents': len(headers), 'Scoped validated journals': len(selected), 'Invoice journals with PK01/S customer positions': len(trigger[DOC].drop_duplicates()), 'Trigger positions without customer master': int((~matched).sum()), 'Trigger positions without creation/user evidence': int((matched & ~evidence).sum()), 'Exception journals': len(affected), 'Exception trigger positions': len(conflicts), 'Output positions': len(output)}
    return (output, metrics)

def remove_old_qa(output_file):
    workbook = load_workbook(output_file)
    temporary_path = None
    try:
        if 'SD_002_QA' not in workbook.sheetnames:
            return
        del workbook['SD_002_QA']
        with tempfile.NamedTemporaryFile(dir=Path(output_file).parent, suffix='.xlsx', delete=False) as temporary_file:
            temporary_path = Path(temporary_file.name)
        workbook.save(temporary_path)
        os.replace(temporary_path, output_file)
    finally:
        workbook.close()
        if temporary_path is not None and temporary_path.exists():
            temporary_path.unlink()

def run_sd_002(context):
    customers, _ = load_sd_customer_master(context)
    ar, _ = load_sd_ar_items(context)
    journals, journal_metrics = load_sd_ar_journals(context)
    output, metrics = analyze_full_journals(customers, ar, journals, get_invoice_types(context))
    path = write_control_sheet(context, 'SD02', output, date_columns=['Customer Created On', 'Company Customer Created On', 'Entry Date', 'Document Date', 'Posting Date'], amount_columns=['Amount in Document Currency', 'Amount in Local Currency'])
    remove_old_qa(path)
    for name, value in {**journal_metrics, **metrics}.items():
        print(f'SD_002 - {name}: {value}')
    print(f'SD_002 - output: {path}')
    _sd002_format_report(path)
    _sd002_write_empty_explanation(path, output, metrics, get_invoice_types(context))
    return {'status': 'OK', 'output_file': path, 'sheet_name': 'SD02', 'rows': len(output), 'summary': metrics}

def _sd002_format_report(output_file):
    """Standard formatting with debit/credit and expandable detail."""
    import os
    import tempfile
    from pathlib import Path
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill
    from openpyxl.utils import get_column_letter
    output_file = Path(output_file)
    book = load_workbook(output_file)
    temporary = None
    try:
        sheet = book['SD02']
        headers = [cell.value for cell in sheet[1]]
        labels = ['Invoice Posted By' if name == 'User ID' else name for name in headers]
        if len(labels) != len(set(labels)):
            raise ValueError('SD_002: duplicate headers.')
        visible = ['Company', 'Document Number', 'Fiscal Year', 'Document Item', 'Trigger', 'Customer Code', 'Customer Name', 'Customer Created By', 'Customer Created On', 'Invoice Posted By', 'Entry Date', 'GL Account', 'Debit Amount', 'Credit Amount', 'Document Currency', 'Item Text']
        required = set(visible) - {'Debit Amount', 'Credit Amount'} | {'Debit Credit', 'Amount in Document Currency'}
        missing = required - set(labels)
        if missing:
            raise ValueError(f'SD_002: missing columns {sorted(missing)}')
        order = visible + [name for name in labels if name not in visible]
        stored = []
        for row in sheet.iter_rows(min_row=2, max_col=len(headers)):
            if any((cell.data_type == 'f' for cell in row)):
                raise ValueError('SD_002 contains formulas. Automatic reordering requires review.')
            record = {name: (cell.value, cell.number_format, cell.data_type) for name, cell in zip(labels, row)}
            dc = str(record['Debit Credit'][0]).strip().upper()
            amount = record['Amount in Document Currency'][0]
            if dc not in {'S', 'H'} or amount is None or isinstance(amount, bool):
                raise ValueError(f'SD_002: invalid amount or S/H in Excel row {row[0].row}')
            try:
                amount = float(amount)
            except (TypeError, ValueError) as error:
                raise ValueError(f'SD_002: nonnumeric amount in Excel row {row[0].row}') from error
            record['Debit Amount'] = (amount if dc == 'S' else 0.0, '#,##0.00', 'n')
            record['Credit Amount'] = (amount if dc == 'H' else 0.0, '#,##0.00', 'n')
            stored.append(record)
        for column, name in enumerate(order, 1):
            sheet.cell(1, column, name)
        for row_number, record in enumerate(stored, 2):
            for column, name in enumerate(order, 1):
                value, number_format, kind = record[name]
                cell = sheet.cell(row_number, column)
                cell.value = value
                cell.data_type = kind
                cell.number_format = number_format
        sheet.column_dimensions.clear()
        dates = {'Customer Created On', 'Company Customer Created On', 'Entry Date', 'Document Date', 'Posting Date'}
        amounts = {'Debit Amount', 'Credit Amount', 'Amount in Document Currency', 'Amount in Local Currency'}
        for column, name in enumerate(order, 1):
            header = sheet.cell(1, column)
            header.font = Font(bold=True)
            header.fill = PatternFill(fill_type='solid', fgColor='D9EAF7')
            header.border = Border()
            header.alignment = Alignment()
            maximum_length = len(str(name))
            for row_number in range(2, sheet.max_row + 1):
                cell = sheet.cell(row_number, column)
                cell.font = Font()
                cell.fill = PatternFill(fill_type=None)
                cell.border = Border()
                cell.alignment = Alignment()
                if name in dates:
                    cell.number_format = 'dd/mm/yyyy'
                elif name in amounts:
                    cell.number_format = '#,##0.00'
                if cell.value is not None:
                    maximum_length = max(maximum_length, len(str(cell.value)))
            dimension = sheet.column_dimensions[get_column_letter(column)]
            dimension.width = min(maximum_length + 2, 45)
            dimension.hidden = name not in visible
            dimension.outlineLevel = 0 if name in visible else 1
            dimension.collapsed = False
        for dimension in sheet.row_dimensions.values():
            dimension.height = None
            dimension.hidden = False
        sheet.sheet_properties.outlinePr.summaryRight = False
        sheet.sheet_view.showOutlineSymbols = True
        sheet.sheet_format.outlineLevelCol = 1
        sheet.column_dimensions[get_column_letter(len(visible))].collapsed = True
        sheet.freeze_panes = 'A2'
        sheet.auto_filter.ref = sheet.dimensions
        sheet.auto_filter.filterColumn = []
        sheet.auto_filter.sortState = None
        sheet.sheet_view.showGridLines = True
        with tempfile.NamedTemporaryFile(dir=output_file.parent, suffix='.xlsx', delete=False) as handle:
            temporary = Path(handle.name)
        book.save(temporary)
        book.close()
        os.replace(temporary, output_file)
    finally:
        book.close()
        if temporary is not None and temporary.exists():
            temporary.unlink()

def _sd002_write_empty_explanation(output_file, output, metrics, types):
    if not output.empty:
        return

    import os
    import tempfile
    from pathlib import Path
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font

    output_file = Path(output_file)
    count = int(
        metrics["Invoice journals with PK01/S customer positions"]
    )
    text = (
        f"Evaluated {count:,} invoice journals. "
        "No cases were found where the same user created "
        "the customer and posted the invoice."
    )
    if count == 0:
        text = "No eligible invoice journals were found within the configured scope."

    book = load_workbook(output_file)
    temporary = None
    try:
        sheet = book["SD02"]

        for merged in list(sheet.merged_cells.ranges):
            if merged.min_row <= 7 and merged.max_row >= 3:
                sheet.unmerge_cells(str(merged))

        for row in sheet.iter_rows(min_row=3, max_row=7):
            for cell in row:
                cell.value = None

        sheet.merge_cells("A3:P3")
        sheet["A3"] = "SD02 - No exceptions"
        sheet["A3"].font = Font(bold=True)

        sheet.merge_cells("A4:P4")
        sheet["A4"] = text
        sheet["A4"].font = Font()
        sheet["A4"].alignment = Alignment(
            wrap_text=True, vertical="top"
        )
        sheet.row_dimensions[3].height = 22
        sheet.row_dimensions[4].height = 32
        for row in range(5, 8):
            sheet.row_dimensions[row].height = None

        sheet.auto_filter.ref = None

        with tempfile.NamedTemporaryFile(
            dir=output_file.parent, suffix=".xlsx", delete=False
        ) as handle:
            temporary = Path(handle.name)

        book.save(temporary)
        book.close()
        os.replace(temporary, output_file)
    finally:
        book.close()
        if temporary is not None and temporary.exists():
            temporary.unlink()
