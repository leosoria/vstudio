"""Shared infrastructure for LBR Segregation of Duties controls.

Loads the three confirmed SAP exports. Conflict matching, invoice classification,
blocked-vendor exclusions and USD conversion belong in the controls.
Loaders return (dataframe, metrics), using run_analysis.py's context.
Input headers must be on row 1. Repeated headers are resolved by occurrence.
Dependencies: pandas, openpyxl; xlrd only for legacy .xls inputs.
"""

import os
import re
import tempfile
import unicodedata
from collections import Counter
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

INVALID_TEXT_VALUES = {"", "nan", "none", "<na>", "nat"}
HEADER_FILL = "D9EAF7"
DATE_NUMBER_FORMAT = "dd/mm/yyyy"
AMOUNT_NUMBER_FORMAT = "#,##0.00"
ALLOWED_INPUT_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}


def get_sd_output_file(context):
    return Path(context["output_folder"]) / (
        f"LBR_Results_SD_{get_period_suffix(context['module'])}.xlsx"
    )


def normalize_text(value):
    """Return trimmed text without pandas missing-value placeholders."""
    if value is None or pd.isna(value):
        return ""

    text = str(value).strip()

    if text.casefold() in INVALID_TEXT_VALUES:
        return ""

    return text


def normalize_lookup(value):
    """Return an accent-, punctuation- and case-insensitive lookup key."""
    text = unicodedata.normalize(
        "NFKD",
        normalize_text(value),
    )

    text = "".join(
        char
        for char in text
        if not unicodedata.combining(char)
    )

    return re.sub(
        r"[^a-z0-9]+",
        "",
        text.casefold(),
    )


def normalize_identifier(value):
    """Normalize an SAP identifier while preserving leading zeroes in text."""
    text = normalize_text(value)

    if re.fullmatch(
        r"\d+\.0",
        text,
    ):
        return text[:-2]

    return text


def normalize_company(value):
    """Normalize an SAP company code to four digits when it is numeric."""
    text = normalize_identifier(
        value
    )

    if text.isdigit():
        return text.zfill(4)

    return text.upper()


def parse_config_date(
    value,
    field_name,
):
    """Parse one required configuration date."""
    text = normalize_text(
        value
    )

    if re.fullmatch(
        r"\d{4}-\d{2}-\d{2}",
        text,
    ):
        parsed = pd.to_datetime(
            text,
            format="%Y-%m-%d",
            errors="coerce",
        )
    else:
        parsed = pd.to_datetime(
            value,
            errors="coerce",
            dayfirst=True,
        )

    if pd.isna(parsed):
        raise ValueError(
            f"SD CONFIG {field_name} is "
            f"empty or invalid: {value!r}."
        )

    return pd.Timestamp(
        parsed
    ).normalize()


def get_period_suffix(module_config):
    """Return YYYYMMDD from the SD module TO date."""
    parsed_to = parse_config_date(
        module_config.get(
            "to",
            "",
        ),
        "TO",
    )

    return parsed_to.strftime(
        "%Y%m%d"
    )


def parse_config_companies(value):
    """Return normalized company codes or None when COMPANIES is ALL."""
    is_collection = isinstance(
        value,
        (
            list,
            tuple,
            set,
        ),
    )

    if (
        not is_collection
        and normalize_text(value).upper()
        == "ALL"
    ):
        return None

    if is_collection:
        raw_values = list(value)
    else:
        raw_values = re.split(
            r"[,;|\n]+",
            normalize_text(value),
        )

    companies = {
        normalize_company(item)
        for item in raw_values
        if normalize_company(item)
    }

    if not companies:
        raise ValueError(
            "SD CONFIG COMPANIES is empty; "
            "use ALL or provide at least "
            "one company."
        )

    return companies


def excel_safe_value(value):
    """Convert pandas/numpy values to safe openpyxl scalars."""
    if value is None or pd.isna(value):
        return None

    if isinstance(
        value,
        pd.Timestamp,
    ):
        return value.to_pydatetime()

    if isinstance(
        value,
        (
            datetime,
            date,
            str,
            int,
            float,
            bool,
        ),
    ):
        return value

    return str(value)


def write_control_sheet(
    context,
    sheet_name,
    dataframe,
    date_columns=None,
    amount_columns=None,
):
    """Replace only one SD sheet while preserving all other sheets."""
    output_file = get_sd_output_file(
        context
    )

    output_file.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    if output_file.exists():
        try:
            workbook = load_workbook(
                output_file
            )
        except PermissionError as error:
            raise PermissionError(
                "SD output is open or "
                f"locked: {output_file}. "
                "Close Excel and retry."
            ) from error
    else:
        workbook = Workbook()

        workbook.remove(
            workbook.active
        )

    if sheet_name in workbook.sheetnames:
        workbook.remove(
            workbook[sheet_name]
        )

    worksheet = workbook.create_sheet(
        sheet_name
    )

    safe = dataframe.copy()

    for column in safe.columns:
        safe[column] = (
            safe[column]
            .map(excel_safe_value)
        )

    for row in dataframe_to_rows(
        safe,
        index=False,
        header=True,
    ):
        worksheet.append(
            row
        )

    header_fill = PatternFill(
        fill_type="solid",
        fgColor=HEADER_FILL,
    )

    for cell in worksheet[1]:
        cell.font = Font(
            bold=True
        )

        cell.fill = header_fill

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    date_columns = set(
        date_columns or []
    )

    amount_columns = set(
        amount_columns or []
    )

    header_positions = {
        cell.value: cell.column
        for cell in worksheet[1]
    }

    for column in date_columns:
        if column in header_positions:
            column_letter = (
                get_column_letter(
                    header_positions[column]
                )
            )

            for cell in worksheet[
                column_letter
            ][1:]:
                cell.number_format = (
                    DATE_NUMBER_FORMAT
                )

    for column in amount_columns:
        if column in header_positions:
            column_letter = (
                get_column_letter(
                    header_positions[column]
                )
            )

            for cell in worksheet[
                column_letter
            ][1:]:
                cell.number_format = (
                    AMOUNT_NUMBER_FORMAT
                )

    for index, column in enumerate(
        safe.columns,
        start=1,
    ):
        values = [
            str(column),
        ] + [
            normalize_text(value)
            for value in safe[column]
        ]

        worksheet.column_dimensions[
            get_column_letter(index)
        ].width = min(
            max(
                len(value)
                for value in values
            ) + 2,
            45,
        )

    temporary_path = None

    try:
        with tempfile.NamedTemporaryFile(
            prefix=(
                f".{output_file.stem}_"
            ),
            suffix=output_file.suffix,
            dir=output_file.parent,
            delete=False,
        ) as temporary_file:
            temporary_path = Path(
                temporary_file.name
            )

        workbook.save(
            temporary_path
        )

        os.replace(
            temporary_path,
            output_file,
        )

    except PermissionError as error:
        raise PermissionError(
            "SD output cannot be saved "
            "because it is open or locked: "
            f"{output_file}."
        ) from error

    finally:
        workbook.close()

        if (
            temporary_path is not None
            and temporary_path.exists()
        ):
            temporary_path.unlink()

    return output_file


# (SAP export heading, canonical name, technical field).
VENDOR_SCHEMA = [
    ('Vendor', 'Vendor Code', 'LFB1-LIFNR'),
    ('Cty', 'Vendor Country', 'LFA1-LAND1'),
    ('Name 1', 'Vendor Name', 'LFA1-NAME1'),
    ('City', 'Vendor City', 'LFA1-ORT01'),
    ('Date', 'Vendor Created On', 'LFA1-ERDAT'),
    ('Created by', 'Vendor Created By', 'LFA1-ERNAM'),
    ('Central Del LOEVM', 'Central Deletion Flag', 'LFA1-LOEVM'),
    ('Central Posting Block SPERR', 'Central Posting Block', 'LFA1-SPERR'),
    ('Central Purchasing Block SPERM', 'Central Purchasing Block', 'LFA1-SPERM'),
    ('Central Payment Block SPERZ', 'Central Payment Block', 'LFA1-SPERZ'),
    ('CoCd', 'Company', 'LFB1-BUKRS'),
    ('District', 'Vendor District', 'LFA1-ORT02'),
    ('Date', 'Company Vendor Created On', 'LFB1-ERDAT'),
    ('Created by', 'Company Vendor Created By', 'LFB1-ERNAM'),
    ('CC Posting Block SPERR', 'Company Posting Block', 'LFB1-SPERR'),
    ('CC Deletion Flag LOE', 'Company Deletion Flag', 'LFB1-LOEVM'),
    ('Pmt meths', 'Payment Methods', 'LFB1-ZWELS'),
    ('PayT', 'Payment Terms', 'LFB1-ZTERM'),
    ('Block key for Payment ZAHLS', 'Payment Block', 'LFB1-ZAHLS'),
    ('Group', 'Vendor Account Group', 'LFA1-KTOKK'),
    ('E-Mail Address', 'Vendor Email', 'ADR6-SMTP_ADDR'),
]

OPEN_SCHEMA = [
    ('CoCd', 'Company', 'BSIK-BUKRS'),
    ('Vendor', 'Vendor Code', 'BSIK-LIFNR'),
    ('DocumentNo', 'Document Number', 'BSIK-BELNR'),
    ('Year', 'Fiscal Year', 'BSIK-GJAHR'),
    ('Itm', 'Document Item', 'BSIK-BUZEI'),
    ('PK', 'Posting Key', 'BSIK-BSCHL'),
    ('D/C', 'Debit Credit', 'BSIK-SHKZG'),
    ('Amount', 'Amount in Document Currency', 'BSIK-WRBTR'),
    ('Crcy', 'Amount Document Currency', None),
    ('Amount in LC', 'Amount in Local Currency', 'BSIK-DMBTR'),
    ('Crcy', 'Local Currency', None),
    ('Text', 'Item Text', 'BSIK-SGTXT'),
    ('Doc. Type', 'Document Type', 'BKPF-BLART'),
    ('Doc. Date', 'Document Date', 'BKPF-BLDAT'),
    ('Pstng Date', 'Posting Date', 'BKPF-BUDAT'),
    ('Entry Dte', 'Entry Date', 'BKPF-CPUDT'),
    ('Time', 'Entry Time', 'BKPF-CPUTM'),
    ('User name', 'User ID', 'BKPF-USNAM'),
    ('TCode', 'Transaction Code', 'BKPF-TCODE'),
    ('Crcy', 'Document Currency', 'BKPF-WAERS'),
    ('Reference', 'Reference', 'BKPF-XBLNR'),
    ('Rev. with', 'Reversal Document', 'BKPF-STBLG'),
    ('Year', 'Reversal Year', 'BKPF-STJAH'),
]

CLEARED_SCHEMA = [
    ('CoCd', 'Company', 'BKPF-BUKRS'),
    ('DocumentNo', 'Document Number', 'BKPF-BELNR'),
    ('Year', 'Fiscal Year', 'BKPF-GJAHR'),
    ('Doc. Type', 'Document Type', 'BKPF-BLART'),
    ('Doc. Date', 'Document Date', 'BKPF-BLDAT'),
    ('Pstng Date', 'Posting Date', 'BKPF-BUDAT'),
    ('Reference', 'Reference', 'BKPF-XBLNR'),
    ('Ref. Tran.', 'Reference Transaction', 'BKPF-AWTYP'),
    ('Ref. Key', 'Reference Key', 'BKPF-AWKEY'),
    ('Vendor', 'Vendor Code', 'BSAK-LIFNR'),
    ('Clearing', 'Clearing Date', 'BSAK-AUGDT'),
    ('Clrng doc.', 'Clearing Document', 'BSAK-AUGBL'),
    ('Year', 'Item Fiscal Year', 'BSAK-GJAHR'),
    ('DocumentNo', 'Item Document Number', 'BSAK-BELNR'),
    ('Pstng Date', 'Item Posting Date', 'BSAK-BUDAT'),
    ('Doc. Date', 'Item Document Date', 'BSAK-BLDAT'),
    ('Crcy', 'Document Currency', 'BSAK-WAERS'),
    ('Reference', 'Item Reference', 'BSAK-XBLNR'),
    ('D/C', 'Debit Credit', 'BSAK-SHKZG'),
    ('Amount in LC', 'Amount in Local Currency', 'BSAK-DMBTR'),
    ('Crcy', 'Local Currency', None),
    ('Amount', 'Amount in Document Currency', 'BSAK-WRBTR'),
    ('Crcy', 'Amount Document Currency', None),
    ('G/L Acct', 'GL Account', 'BSAK-SAKNR'),
    ('G/L', 'GL Account HKONT', 'BSAK-HKONT'),
    ('PayT', 'Payment Terms', 'BSAK-ZTERM'),
    ('Inv. ref.', 'Invoice Reference', 'BSAK-REBZG'),
    ('User name', 'User ID', 'BKPF-USNAM'),
    ('Profit Ctr', 'Profit Center', None),
    ('Entry Dte', 'Entry Date', 'BKPF-CPUDT'),
    ('TCode', 'Transaction Code', 'BKPF-TCODE'),
    ('Itm', 'Document Item', 'BSAK-BUZEI'),
    ('PK', 'Posting Key', 'BSAK-BSCHL'),
    ('Rev. with', 'Reversal Document', 'BKPF-STBLG'),
    ('Year', 'Reversal Year', 'BKPF-STJAH'),
]

SCHEMAS = {
    'VENDOR': VENDOR_SCHEMA,
    'AP_OPEN': OPEN_SCHEMA,
    'AP_CLEARED': CLEARED_SCHEMA,
}

AP_REQUIRED = {
    'Company', 'Vendor Code', 'Document Number', 'Fiscal Year',
    'Document Item', 'Posting Key', 'Debit Credit', 'Document Type',
    'Entry Date', 'User ID', 'Transaction Code', 'Document Currency',
    'Amount in Document Currency', 'Amount in Local Currency',
    'Reversal Document', 'Reversal Year',
}

VENDOR_REQUIRED = {
    'Company', 'Vendor Code', 'Vendor Name', 'Vendor Created On',
    'Vendor Created By', 'Company Vendor Created On',
    'Company Vendor Created By',
}


def find_sd_input_file(context, source):
    """Find one period-specific input, accepting CLRD as a CLEARED alias."""
    if source not in SCHEMAS:
        raise ValueError(f'Unknown SD source: {source}')

    folder = Path(context['input_folder'])

    if not folder.is_dir():
        raise FileNotFoundError(f'SD input folder not found: {folder}')

    suffix = get_period_suffix(context['module'])
    names = {normalize_lookup(f'LBR_SD_{source}_{suffix}')}

    if source == 'AP_CLEARED':
        names.add(normalize_lookup(f'LBR_SD_AP_CLRD_{suffix}'))

    matches = sorted(
        p for p in folder.rglob('*')
        if p.is_file()
        and not p.name.startswith('~$')
        and p.suffix.lower() in ALLOWED_INPUT_EXTENSIONS
        and normalize_lookup(p.stem) in names
    )

    if len(matches) != 1:
        raise ValueError(
            f'SD {source}: expected exactly one '
            f'LBR_SD_{source}_{suffix}.xlsx in {folder}; found {matches}'
        )

    return matches[0]


def resolve_sd_columns(headers, source):
    """Map duplicate headings BEFORE pandas can suffix or drop empty columns.

    Every repeated human heading must have the confirmed occurrence count.
    Technical or canonical headings are also accepted when uniquely named.
    Unknown extra columns are listed in the loader metrics.
    """
    schema = SCHEMAS[source]
    keys = [normalize_lookup(h) for h in headers]
    counts = Counter(keys)
    expected = Counter(normalize_lookup(h) for h, _, _ in schema)

    for heading, count in expected.items():
        if count > 1 and counts[heading] and counts[heading] != count:
            raise ValueError(
                f'SD {source}: repeated heading {heading!r} has '
                f'{counts[heading]} occurrences; expected {count}. '
                'Keep the confirmed export or use unique technical headers.'
            )

    occurrences = Counter()
    mapping = {}
    used = set()

    for heading, canonical, technical in schema:
        key = normalize_lookup(heading)
        occurrence = occurrences[key]
        occurrences[key] += 1

        unique_aliases = {normalize_lookup(canonical)} - {key}

        if technical:
            unique_aliases.add(normalize_lookup(technical))

        explicit = [
            i for i, k in enumerate(keys)
            if k in unique_aliases
        ]
        human = [
            i for i, k in enumerate(keys)
            if k == key
        ]

        candidates = explicit or human[occurrence:occurrence + 1]

        if len(candidates) > 1:
            raise ValueError(
                f'SD {source}: ambiguous field {canonical}'
            )

        if candidates:
            i = candidates[0]

            if i in used:
                raise ValueError(
                    f'SD {source}: column reused for {canonical}'
                )

            mapping[canonical] = i
            used.add(i)

    required = VENDOR_REQUIRED if source == 'VENDOR' else AP_REQUIRED
    missing = required - mapping.keys()

    if missing:
        raise ValueError(
            f'SD {source}: missing fields {sorted(missing)}. '
            f'Headers: {headers}'
        )

    return mapping


def parse_sap_date(value):
    if not normalize_text(value) or normalize_text(value) in {
        '00000000', '0'
    }:
        return pd.NaT

    if isinstance(value, (datetime, date, pd.Timestamp)):
        return pd.Timestamp(value).normalize()

    if isinstance(value, (int, float)) and 20000 <= value <= 80000:
        return (
            pd.Timestamp('1899-12-30')
            + pd.Timedelta(days=value)
        )

    text = normalize_text(value)
    fmt = '%Y%m%d' if re.fullmatch(r'\d{8}', text) else None

    if re.match(r'^\d{4}-\d{2}-\d{2}', text):
        return pd.to_datetime(text, errors='coerce').normalize()

    return pd.to_datetime(
        text,
        format=fmt,
        dayfirst=True,
        errors='coerce',
    )


def parse_sap_amount(value):
    """Parse Excel numbers or US/BR formatted strings; blank stays missing."""
    text = normalize_text(value)

    if not text:
        return float('nan')

    if isinstance(value, (int, float)):
        return float(value)

    text = text.replace('\u00a0', '').replace(' ', '')
    negative = (
        text.endswith('-')
        or (text.startswith('(') and text.endswith(')'))
    )
    text = text.rstrip('-').strip('()')

    if ',' in text and '.' in text:
        decimal, thousands = (
            (',', '.')
            if text.rfind(',') > text.rfind('.')
            else ('.', ',')
        )
        text = text.replace(thousands, '').replace(decimal, '.')

    elif ',' in text:
        text = (
            text.replace(',', '')
            if re.fullmatch(r'[+-]?\d{1,3}(,\d{3})+', text)
            else text.replace(',', '.')
        )

    amount = float(text)
    return -abs(amount) if negative else amount


def _normalize_sd_data(frame, source):
    original = frame.copy()

    for column in frame:
        frame[column] = frame[column].map(normalize_text)

    frame['Company'] = frame['Company'].map(normalize_company)

    # SAP vendor IDs may be alphanumeric.
    # Only numeric IDs are zero padded.
    widths = {
        'Vendor Code': 10,
        'Document Number': 10,
        'Fiscal Year': 4,
        'Document Item': 3,
        'Posting Key': 2,
        'Reversal Document': 10,
        'Reversal Year': 4,
        'Item Document Number': 10,
        'Item Fiscal Year': 4,
        'Clearing Document': 10,
        'Invoice Reference': 10,
    }

    for column, width in widths.items():
        if column in frame:
            def convert(value):
                text = normalize_identifier(value)

                if (
                    column in {
                        'Reversal Document',
                        'Reversal Year',
                        'Invoice Reference',
                    }
                    and text
                    and set(text) == {'0'}
                ):
                    return ''

                return (
                    text.zfill(width)
                    if text.isdigit()
                    else text.upper()
                )

            frame[column] = frame[column].map(convert)

    for column in (
        'Vendor Created By',
        'Company Vendor Created By',
        'User ID',
        'Transaction Code',
        'Document Type',
        'Debit Credit',
        'Document Currency',
        'Amount Document Currency',
        'Local Currency',
    ):
        if column in frame:
            frame[column] = frame[column].str.upper()

    date_columns = (
        ('Vendor Created On', 'Company Vendor Created On')
        if source == 'VENDOR'
        else (
            'Entry Date',
            'Document Date',
            'Posting Date',
            'Clearing Date',
            'Item Posting Date',
            'Item Document Date',
        )
    )

    for column in date_columns:
        if column in frame:
            raw = original[column]
            parsed = raw.map(parse_sap_date)
            nonblank = raw.map(normalize_text).ne('')

            if (nonblank & parsed.isna()).any():
                raise ValueError(
                    f'SD {source}: invalid dates in {column}'
                )

            frame[column] = pd.to_datetime(parsed)

    for column in (
        'Amount in Document Currency',
        'Amount in Local Currency',
    ):
        if column in frame:
            try:
                frame[column] = frame[column].map(parse_sap_amount)
            except (ValueError, TypeError) as error:
                raise ValueError(
                    f'SD {source}: invalid amounts in {column}'
                ) from error

            if frame[column].isna().any():
                raise ValueError(
                    f'SD {source}: blank amounts in {column}'
                )

    for column in (
        'Fiscal Year',
        'Document Number',
        'Posting Date',
        'Document Date',
    ):
        other = 'Item ' + column

        if column in frame and other in frame:
            if not frame[column].equals(frame[other]):
                raise ValueError(
                    f'SD {source}: header/item mismatch for {column}'
                )

    if 'Amount Document Currency' in frame:
        populated = frame['Amount Document Currency'].ne('')

        if (
            populated
            & frame['Amount Document Currency'].ne(
                frame['Document Currency']
            )
        ).any():
            raise ValueError(
                f'SD {source}: document currencies disagree'
            )

    return frame


def load_sd_input(context, source):
    """Load the first sheet or context['sd_input_sheets'][source].

    All physical columns, including blank reversal/currency columns,
    are preserved until mapping. Only wholly empty rows are removed.
    Vendor dates are NOT period-filtered.
    AP uses Entry Date, never Clearing Date.
    No transaction, invoice, deletion or block filter is applied.
    """
    file = find_sd_input_file(context, source)

    with pd.ExcelFile(file) as book:
        sheet = context.get(
            'sd_input_sheets', {}
        ).get(source, book.sheet_names[0])

        raw = pd.read_excel(
            book,
            sheet_name=sheet,
            header=None,
            dtype=object,
            keep_default_na=False,
        )

    if raw.empty:
        raise ValueError(
            f'SD {source}: missing header row in {file}'
        )

    headers = [
        normalize_text(v)
        for v in raw.iloc[0]
    ]
    mapping = resolve_sd_columns(headers, source)

    body = raw.iloc[1:].copy()
    body = body.loc[
        ~body.apply(
            lambda col: col.map(normalize_text)
        ).eq('').all(axis=1)
    ]

    frame = pd.DataFrame({
        name: body.iloc[:, pos]
        for name, pos in mapping.items()
    })

    metrics = {
        'input_file': str(file),
        'input_sheet': sheet,
        'rows_read': len(frame),
        'unmapped_headers': [
            h for i, h in enumerate(headers)
            if i not in mapping.values()
        ],
        'column_mapping': {
            name: {
                'position': pos + 1,
                'header': headers[pos],
            }
            for name, pos in mapping.items()
        },
    }

    keys = (
        ['Company', 'Vendor Code']
        if source == 'VENDOR'
        else [
            'Company',
            'Document Number',
            'Fiscal Year',
            'Document Item',
            'Vendor Code',
        ]
    )

    frame, removed_totals = _remove_sd_export_amount_totals(frame, keys, source)
    metrics["excluded_export_total_rows"] = removed_totals

    frame = _normalize_sd_data(frame, source)

    companies = parse_config_companies(
        context['module'].get('companies', '')
    )

    company_mask = (
        pd.Series(True, index=frame.index)
        if companies is None
        else frame['Company'].isin(companies)
    )

    metrics['excluded_by_company'] = int(
        (~company_mask).sum()
    )
    frame = frame.loc[company_mask].copy()

    start = parse_config_date(
        context['module'].get('from', ''),
        'FROM',
    )
    end = parse_config_date(
        context['module'].get('to', ''),
        'TO',
    )

    if start > end:
        raise ValueError(
            'SD CONFIG FROM must not be after TO'
        )

    metrics.update(
        date_from=start,
        date_to=end,
        excluded_by_date=0,
    )

    if source != 'VENDOR':
        if frame['Entry Date'].isna().any():
            raise ValueError(
                f'SD {source}: missing Entry Date'
            )

        date_mask = frame['Entry Date'].between(start, end)
        metrics['excluded_by_date'] = int(
            (~date_mask).sum()
        )
        frame = frame.loc[date_mask].copy()

    frame['Source File'] = file.name

    # pandas index 1 corresponds to Excel row 2.
    frame['Source Row'] = frame.index + 1
    frame['Source Type'] = source

    metrics['rows_after_config_filters'] = len(frame)
    metrics['duplicate_key_rows'] = int(
        frame.duplicated(keys, keep=False).sum()
    )

    # Keep supporting detail; each control defines aggregation.
    return frame.reset_index(drop=True), metrics


def load_vendor_master(context):
    return load_sd_input(context, 'VENDOR')


def load_ap_open(context):
    return load_sd_input(context, 'AP_OPEN')


def load_ap_cleared(context):
    return load_sd_input(context, 'AP_CLEARED')


def load_ap_items(context):
    """Combine sources, refusing open/cleared overlap."""
    opened, open_metrics = load_ap_open(context)
    cleared, cleared_metrics = load_ap_cleared(context)

    keys = [
        'Company',
        'Document Number',
        'Fiscal Year',
        'Document Item',
    ]
    overlap = opened[keys].merge(
        cleared[keys],
        on=keys,
    )

    if not overlap.empty:
        raise ValueError(
            'SD AP inputs overlap on company/document/year/item; '
            'refresh both exports from the same extraction run.'
        )

    return pd.concat(
        [opened, cleared],
        ignore_index=True,
    ), {
        'AP_OPEN': open_metrics,
        'AP_CLEARED': cleared_metrics,
    }


def _remove_sd_export_amount_totals(frame, keys, source):
    """Remove AP amount-only totals, never incomplete transactions."""
    blank_keys = frame[keys].apply(
        lambda column: column.map(normalize_text)
    ).eq("")

    allowed_total_columns = {
        "Amount in Document Currency",
        "Amount in Local Currency",
        "Amount Document Currency",
        "Document Currency",
        "Local Currency",
    }

    other_columns = [
        column for column in frame.columns
        if column not in allowed_total_columns
    ]

    other_blank = frame[other_columns].apply(
        lambda column: column.map(normalize_text)
    ).eq("").all(axis=1)

    amount_columns = [
        column for column in (
            "Amount in Document Currency",
            "Amount in Local Currency",
        )
        if column in frame.columns
    ]

    if amount_columns:
        has_amount = frame[amount_columns].apply(
            lambda column: column.map(normalize_text)
        ).ne("").any(axis=1)
    else:
        has_amount = pd.Series(False, index=frame.index)

    totals = (
        blank_keys.all(axis=1)
        & other_blank
        & has_amount
    )

    if source == "VENDOR":
        totals[:] = False

    invalid = blank_keys.any(axis=1) & ~totals

    if invalid.any():
        examples = []

        for index in frame.index[invalid][:10]:
            missing = [
                column for column in keys
                if blank_keys.at[index, column]
            ]

            examples.append(
                f"Excel row {index + 1}: missing {missing}"
            )

        raise ValueError(
            f"SD {source}: incomplete transaction keys "
            "after excluding totals. "
            + "; ".join(examples)
        )

    count = int(totals.sum())

    if count:
        print(
            f"SD {source}: excluded {count} "
            "SAP export amount-total rows."
        )

    return frame.loc[~totals].copy(), count



def load_sd_journals(context):
    """Read dedicated SD BSEG exports without GL module dependencies."""
    folder = Path(context['input_folder'])
    files = sorted((path for path in folder.rglob('*') if path.is_file() and (not path.name.startswith('~$')) and (path.suffix.lower() in ALLOWED_INPUT_EXTENSIONS) and normalize_lookup(path.stem).startswith('lbrsdjournal') if _sd_journal_file_matches(path, context)))
    if not files:
        raise FileNotFoundError('Missing input: LBR SD_JOURNAL_*.xlsx')
    schema = [('CoCd', 'Company', 'BUKRS'), ('DocumentNo', 'Document Number', 'BELNR'), ('Year', 'Fiscal Year', 'GJAHR'), ('Itm', 'Document Item', 'BUZEI'), ('AccTy', 'Account Type', 'KOART'), ('PK', 'Posting Key', 'BSCHL'), ('G/L', 'GL Account', 'HKONT'), ('Vendor', 'Vendor Code', 'LIFNR'), ('Customer', 'Customer Code', 'KUNNR'), ('D/C', 'Debit Credit', 'SHKZG'), ('Amount', 'Amount in Document Currency', 'WRBTR'), ('Crcy', 'BSEG Document Currency', None), ('Amount in LC', 'Amount in Local Currency', 'DMBTR'), ('Crcy', 'Local Currency', None), ('Text', 'Item Text', 'SGTXT')]
    frames = []
    rows_read = 0
    totals_removed = 0
    for path in files:
        with pd.ExcelFile(path) as book:
            sheet = next((name for name in book.sheet_names if name.lower() == 'sheet1'), book.sheet_names[0])
            raw = pd.read_excel(book, sheet_name=sheet, header=None, dtype=object, keep_default_na=False)
        if raw.empty:
            raise ValueError(f'SD JOURNAL: empty file {path.name}')
        headers = [normalize_lookup(re.sub('\\.\\d+$', '', normalize_text(value))) for value in raw.iloc[0]]
        mapping = {}
        occurrences = Counter()
        for heading, name, technical in schema:
            key = normalize_lookup(heading)
            positions = [index for index, value in enumerate(headers) if value == key]
            direct = [index for index, value in enumerate(headers) if technical and value in {normalize_lookup(technical), normalize_lookup('BSEG-' + technical)}]
            occurrence = occurrences[key]
            occurrences[key] += 1
            found = direct or positions[occurrence:occurrence + 1]
            if len(found) != 1:
                if name in {'BSEG Document Currency', 'Local Currency', 'Item Text'} and (not found):
                    continue
                raise ValueError(f'SD JOURNAL {path.name}: missing/ambiguous field {name}')
            mapping[name] = found[0]
        body = raw.iloc[1:].copy()
        body = body.loc[~body.apply(lambda column: column.map(normalize_text)).eq('').all(axis=1)]
        frame = pd.DataFrame({name: body.iloc[:, index].map(normalize_text) for name, index in mapping.items()})
        rows_read += len(frame)
        keys = ['Company', 'Document Number', 'Fiscal Year', 'Document Item']
        blank = frame[keys].eq('')
        allowed = {'Amount in Document Currency', 'Amount in Local Currency', 'BSEG Document Currency', 'Local Currency'}
        totals = blank.all(axis=1) & frame[[column for column in frame if column not in allowed]].eq('').all(axis=1)
        totals_removed += int(totals.sum())
        bad = blank.any(axis=1) & ~totals
        if bad.any():
            rows = (frame.index[bad] + 1).tolist()[:5]
            raise ValueError(f'SD JOURNAL {path.name}: incomplete keys in Excel rows {rows}')
        frame = frame.loc[~totals].copy()
        frame['Company'] = frame['Company'].map(normalize_company)
        widths = {'Document Number': 10, 'Fiscal Year': 4, 'Document Item': 3, 'Vendor Code': 10, 'Customer Code': 10, 'Posting Key': 2, 'GL Account': 10}
        for name, width in widths.items():
            frame[name] = frame[name].map(normalize_identifier).map(lambda value: value.zfill(width) if value.isdigit() else value.upper())
        for name in ['Account Type', 'Debit Credit', 'BSEG Document Currency', 'Local Currency']:
            if name in frame:
                frame[name] = frame[name].str.upper()
        for name in ['Amount in Document Currency', 'Amount in Local Currency']:
            frame[name] = frame[name].map(parse_sap_amount)
            if frame[name].isna().any():
                raise ValueError(f'SD JOURNAL {path.name}: blank amounts in {name}')
        if not frame['Debit Credit'].isin(['S', 'H']).all():
            raise ValueError(f'SD JOURNAL {path.name}: invalid Debit/Credit indicator')
        missing_vendor = frame['Account Type'].eq('K') & frame['Vendor Code'].eq('')
        if missing_vendor.any():
            raise ValueError(f'SD JOURNAL {path.name}: vendor positions without vendor ID')
        frame['Journal Source File'] = path.name
        frame['Journal Source Row'] = frame.index + 1
        frames.append(frame)
        print(f'SD JOURNAL: {path.name}, {len(frame)} positions', flush=True)
    combined = pd.concat(frames, ignore_index=True)
    keys = ['Company', 'Fiscal Year', 'Document Number', 'Document Item']
    business = [column for column in combined if column not in {'Journal Source File', 'Journal Source Row'}]
    distinct = combined.drop_duplicates(business)
    conflicting = distinct.duplicated(keys, keep=False)
    if conflicting.any():
        examples = distinct.loc[conflicting, keys].head(5)
        raise ValueError(f"SD JOURNAL: conflicting duplicate positions: {examples.to_dict('records')}")
    unique = combined.drop_duplicates(keys).copy()
    return (unique, {'files': len(files), 'rows_read': rows_read, 'export_totals_removed': totals_removed, 'duplicate_rows_removed': len(combined) - len(unique)})


def _sd_journal_file_matches(path, context):
    """Separate AP and AR journal files."""
    is_ar = normalize_lookup(path.stem).startswith("lbrsdjournalar")
    scope = str(context.get("sd_journal_scope", "AP")).upper()

    if scope not in {"AP", "AR"}:
        raise ValueError("sd_journal_scope must be AP or AR")

    return is_ar if scope == "AR" else not is_ar


def _sd02_load_input(context, source):
    prefix = {
        "CUSTOMER": "LBR_SD_CUSTOMER",
        "AR_OPEN": "LBR_SD_AR_OPEN",
        "AR_CLEARED": "LBR_SD_AR_CLEARED",
    }[source]

    folder = Path(context["input_folder"])
    suffix = get_period_suffix(context["module"])
    expected = normalize_lookup(prefix + "_" + suffix)

    matches = sorted(
        path for path in folder.rglob("*")
        if path.is_file()
        and path.suffix.lower() in ALLOWED_INPUT_EXTENSIONS
        and not path.name.startswith("~$")
        and normalize_lookup(path.stem) == expected
    )

    if len(matches) != 1:
        raise ValueError(
            f"SD02 {source}: expected exactly one "
            f"{prefix}_{suffix}.xlsx; found {matches}"
        )

    path = matches[0]

    customer_schema = [
        ("CoCd", "Company"),
        ("Customer", "Customer Code"),
        ("Name 1", "Customer Name"),
        ("Cty", "Customer Country"),
        ("City", "Customer City"),
        ("Date", "Customer Created On"),
        ("Created by", "Customer Created By"),
        ("Group", "Customer Account Group"),
        ("DelF", "Central Deletion Flag"),
        ("B", "Central Posting Block"),
        ("Date", "Company Customer Created On"),
        ("Created by", "Company Customer Created By"),
        ("DelF", "Company Deletion Flag"),
        ("B", "Company Posting Block"),
    ]

    ar_schema = [
        ("CoCd", "Company"),
        ("Customer", "Customer Code"),
        ("DocumentNo", "Document Number"),
        ("Year", "Fiscal Year"),
        ("Itm", "Document Item"),
        ("PK", "Posting Key"),
        ("D/C", "Debit Credit"),
        ("Amount", "Amount in Document Currency"),
        ("Crcy", "Amount Document Currency"),
        ("Amount in LC", "Amount in Local Currency"),
        ("Crcy", "Local Currency"),
        ("Text", "Item Text"),
        ("Doc. Type", "Document Type"),
        ("Doc. Date", "Document Date"),
        ("Pstng Date", "Posting Date"),
        ("Entry Dte", "Entry Date"),
        ("Time", "Entry Time"),
        ("User name", "User ID"),
        ("TCode", "Transaction Code"),
        ("Crcy", "Document Currency"),
        ("Reference", "Reference"),
        ("Rev. with", "Reversal Document"),
        ("Year", "Reversal Year"),
    ]

    schema = (
        customer_schema if source == "CUSTOMER" else ar_schema
    )

    if source == "AR_CLEARED":
        schema = schema + [
            ("Clearing", "Clearing Date"),
            ("Clrng doc.", "Clearing Document"),
        ]

    with pd.ExcelFile(path) as book:
        sheet = next(
            (
                name for name in book.sheet_names
                if name.lower() == "sheet1"
            ),
            book.sheet_names[0],
        )

        raw = pd.read_excel(
            book,
            sheet_name=sheet,
            header=None,
            dtype=object,
            keep_default_na=False,
        )

    if raw.empty:
        raise ValueError(f"SD02 {source}: missing header row")

    headers = [
        normalize_lookup(
            re.sub(r"\.\d+$", "", normalize_text(value))
        )
        for value in raw.iloc[0]
    ]

    expected_counts = Counter(
        normalize_lookup(heading)
        for heading, _ in schema
    )

    counts = Counter(headers)

    for key, count in expected_counts.items():
        if counts[key] != count:
            raise ValueError(
                f"SD02 {source}: header {key!r} occurs "
                f"{counts[key]} times; expected {count}"
            )

    used = Counter()
    mapping = {}

    for heading, name in schema:
        key = normalize_lookup(heading)
        positions = [
            index for index, value in enumerate(headers)
            if value == key
        ]
        mapping[name] = positions[used[key]]
        used[key] += 1

    body = raw.iloc[1:].copy()

    body = body.loc[
        ~body.apply(
            lambda column: column.map(normalize_text)
        ).eq("").all(axis=1)
    ]

    original = pd.DataFrame({
        name: body.iloc[:, index]
        for name, index in mapping.items()
    })

    frame = original.apply(
        lambda column: column.map(normalize_text)
    )

    keys = (
        ["Company", "Customer Code"]
        if source == "CUSTOMER"
        else [
            "Company",
            "Fiscal Year",
            "Document Number",
            "Document Item",
            "Customer Code",
        ]
    )

    blank = frame[keys].eq("")

    allowed = {
        "Amount in Document Currency",
        "Amount in Local Currency",
        "Amount Document Currency",
        "Document Currency",
        "Local Currency",
    }

    totals = (
        blank.all(axis=1)
        & frame[
            [column for column in frame if column not in allowed]
        ].eq("").all(axis=1)
    )

    if source == "CUSTOMER":
        totals[:] = False

    invalid = blank.any(axis=1) & ~totals

    if invalid.any():
        rows = (frame.index[invalid] + 1).tolist()[:10]

        raise ValueError(
            f"SD02 {source}: incomplete keys in Excel rows {rows}"
        )

    rows_read = len(frame)
    removed = int(totals.sum())
    frame = frame.loc[~totals].copy()

    frame["Company"] = frame["Company"].map(normalize_company)

    widths = {
        "Customer Code": 10,
        "Document Number": 10,
        "Fiscal Year": 4,
        "Document Item": 3,
        "Posting Key": 2,
        "Reversal Document": 10,
        "Reversal Year": 4,
        "Clearing Document": 10,
    }

    for name, width in widths.items():
        if name in frame:
            def convert(value):
                text = normalize_identifier(value)

                if (
                    name in {"Reversal Document", "Reversal Year"}
                    and text
                    and set(text) == {"0"}
                ):
                    return ""

                return (
                    text.zfill(width)
                    if text.isdigit()
                    else text.upper()
                )

            frame[name] = frame[name].map(convert)

    for name in [
        "Customer Created By",
        "Company Customer Created By",
        "User ID",
        "Document Type",
        "Transaction Code",
        "Debit Credit",
        "Amount Document Currency",
        "Document Currency",
        "Local Currency",
    ]:
        if name in frame:
            frame[name] = frame[name].str.upper()

    for name in [
        "Customer Created On",
        "Company Customer Created On",
        "Entry Date",
        "Document Date",
        "Posting Date",
        "Clearing Date",
    ]:
        if name in frame:
            values = original.loc[frame.index, name]
            parsed = values.map(parse_sap_date)

            if (
                values.map(normalize_text).ne("")
                & parsed.isna()
            ).any():
                raise ValueError(
                    f"SD02 {source}: invalid dates in {name}"
                )

            frame[name] = pd.to_datetime(parsed)

    for name in [
        "Amount in Document Currency",
        "Amount in Local Currency",
    ]:
        if name in frame:
            frame[name] = frame[name].map(parse_sap_amount)

            if frame[name].isna().any():
                raise ValueError(
                    f"SD02 {source}: blank amounts in {name}"
                )

    companies = parse_config_companies(
        context["module"].get("companies", "")
    )

    company_mask = (
        pd.Series(True, index=frame.index)
        if companies is None
        else frame["Company"].isin(companies)
    )

    excluded_company = int((~company_mask).sum())
    frame = frame.loc[company_mask].copy()

    start = parse_config_date(
        context["module"].get("from", ""), "FROM"
    )
    end = parse_config_date(
        context["module"].get("to", ""), "TO"
    )

    if start > end:
        raise ValueError(
            "SD02 CONFIG FROM must not be after TO"
        )

    excluded_date = 0

    if source != "CUSTOMER":
        if frame["Entry Date"].isna().any():
            raise ValueError(
                f"SD02 {source}: missing Entry Date"
            )

        mask = frame["Entry Date"].between(start, end)
        excluded_date = int((~mask).sum())
        frame = frame.loc[mask].copy()

        currency_bad = (
            frame["Amount Document Currency"].ne("")
            & frame["Amount Document Currency"].ne(
                frame["Document Currency"]
            )
        )

        if currency_bad.any():
            raise ValueError(
                f"SD02 {source}: document currencies disagree"
            )

    frame["Source File"] = path.name
    frame["Source Row"] = frame.index + 1
    frame["Source Type"] = source

    metrics = {
        "rows_read": rows_read,
        "removed_export_totals": removed,
        "excluded_company": excluded_company,
        "excluded_date": excluded_date,
        "rows_after_filters": len(frame),
        "input_file": str(path),
    }

    print(
        f"SD02 {source}: {len(frame)} rows; "
        f"export totals removed: {removed}"
    )

    return frame.reset_index(drop=True), metrics


def load_sd_customer_master(context):
    return _sd02_load_input(context, "CUSTOMER")


def load_sd_ar_items(context):
    opened, open_metrics = _sd02_load_input(
        context, "AR_OPEN"
    )
    cleared, cleared_metrics = _sd02_load_input(
        context, "AR_CLEARED"
    )

    keys = [
        "Company",
        "Fiscal Year",
        "Document Number",
        "Document Item",
    ]

    if not opened[keys].merge(
        cleared[keys], on=keys
    ).empty:
        raise ValueError(
            "SD02: AR Open and AR Cleared overlap; "
            "refresh both exports"
        )

    combined = pd.concat(
        [opened, cleared], ignore_index=True
    )

    business = [
        column for column in combined
        if column not in {"Source File", "Source Row"}
    ]

    distinct = combined.drop_duplicates(business)

    if distinct.duplicated(keys, keep=False).any():
        raise ValueError(
            "SD02: conflicting duplicate AR positions"
        )

    combined = combined.drop_duplicates(keys).copy()

    return combined, {
        "AR_OPEN": open_metrics,
        "AR_CLEARED": cleared_metrics,
    }


def load_sd_ar_journals(context):
    frame, metrics = load_sd_journals({
        **context,
        "sd_journal_scope": "AR",
    })

    bad = (
        frame["Account Type"].eq("D")
        & frame["Customer Code"].eq("")
    )

    if bad.any():
        raise ValueError(
            "SD02 JOURNAL: customer positions without customer ID"
        )

    return frame, metrics

# SD_PURCHASING_UTILITIES_BEGIN
"""Common infrastructure for independent LBR Purchase Order controls.

This module contains input discovery, SAP Excel normalization, configuration
filters and output-workbook handling. Control-specific analytics do not belong
here. In particular, PO01 split detection and creator-coverage rules belong in
``modules/PO/po_001.py``.
"""
import os as _sdp_os
import re as _sdp_re
import tempfile as _sdp_tempfile
import unicodedata as _sdp_unicodedata
from datetime import date as _sdp_date, datetime as _sdp_datetime
from pathlib import Path as _sdp_Path
import pandas as _sdp_pd
from openpyxl import Workbook as _sdp_Workbook, load_workbook as _sdp_load_workbook
from openpyxl.styles import Font as _sdp_Font, PatternFill as _sdp_PatternFill
from openpyxl.utils import get_column_letter as _sdp_get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows as _sdp_dataframe_to_rows
_sdp_PO_INPUT_PREFIX = 'LBR SD ORDERS'
_sdp_PR_LINES_INPUT_PREFIX = 'LBR PO PR Lines'
_sdp_PR_GR_INPUT_PREFIX = 'LBR SD RECEIPTS'
_sdp_PO_CDHDR_INPUT_PREFIX = 'LBR SD RELEASE CDHDR'
_sdp_PO_CDPOS_INPUT_PREFIX = 'LBR SD RELEASE CDPOS'
_sdp_PO_INPUT_EXTENSION = '.xlsx'
_sdp_PO_INPUT_SHEET = 'Sheet1'
_sdp_PO_HEADER_ROW = 1
_sdp_PO_OUTPUT_PREFIX = 'LBR_Results_PO'
_sdp_ALLOWED_INPUT_EXTENSIONS = {'.xlsx', '.xlsm', '.xls'}
_sdp_INVALID_TEXT_VALUES = {'', 'nan', 'none', '<na>'}
_sdp_HEADER_FILL = 'D9EAF7'
_sdp_DATE_NUMBER_FORMAT = 'dd/mm/yyyy'
_sdp_AMOUNT_NUMBER_FORMAT = '#,##0.00'
_sdp_PO_LINE_ALIASES = {'Company': ('CoCd', 'Company', 'Company Code', 'BUKRS', 'EKKO-BUKRS'), 'PO Number': ('Purch.Doc.', 'Purchasing Document', 'PO Number', 'EBELN', 'EKKO-EBELN'), 'PO Line': ('Item', 'PO Line', 'EBELP', 'EKPO-EBELP'), 'Vendor Code': ('Vendor', 'Vendor Code', 'LIFNR', 'EKKO-LIFNR'), 'Vendor Name': ('Vendor Name', 'Name 1', 'NAME1', 'LFA1-NAME1'), 'PO Doc Date': ('Doc. Date', 'Document Date', 'PO Document Date', 'BEDAT', 'EKKO-BEDAT'), 'PO Creator ID': ('Created by', 'Created By', 'PO Creator ID', 'ERNAM', 'EKKO-ERNAM'), 'Item Code': ('Material', 'Item Code', 'MATNR', 'EKPO-MATNR'), 'PO Quantity': ('PO Quantity', 'Order Quantity', 'MENGE', 'EKPO-MENGE'), 'PO UOM': ('OUn.1', 'Order Unit', 'Unit of Measure', 'MEINS', 'EKPO-MEINS'), 'PO Unit Price': ('Net Price', 'PO Unit Price', 'NETPR', 'EKPO-NETPR'), 'PO Price Unit': ('Per', 'Price Unit', 'PEINH', 'EKPO-PEINH'), 'PO Doc Currency': ('Crcy.1', 'Currency', 'WAERS', 'EKKO-WAERS'), 'PO Line Total': ('Net Value', 'Net Order Value', 'NETWR', 'EKPO-NETWR'), 'PO Material Description': ('Short Text', 'Description', 'TXZ01', 'EKPO-TXZ01'), 'PO Document Type': ('Type', 'Document Type', 'BSART', 'EKKO-BSART'), 'Purchasing Organization': ('POrg', 'Purchasing Organization', 'EKORG', 'EKKO-EKORG'), 'Purchasing Group': ('PGr', 'Purchasing Group', 'EKGRP', 'EKKO-EKGRP'), 'Plant': ('Plnt', 'Plant', 'WERKS', 'EKPO-WERKS'), 'PO Line Deleted': ('DCI', 'Deletion Indicator', 'LOEKZ', 'EKPO-LOEKZ'), 'PO Delivery Completed': ('FIn', 'Delivery Completed', 'ELIKZ', 'EKPO-ELIKZ'), 'PR Number': ('Purch.Req.', 'Purchase Requisition', 'PR Number', 'BANFN', 'EKPO-BANFN'), 'PR Line': ('Item.1', 'Requisition Item', 'PR Line', 'BNFPO', 'EKPO-BNFPO')}
_sdp_PO01_REQUIRED_FIELDS = ('Company', 'PO Number', 'PO Line', 'Vendor Code', 'PO Doc Date', 'PO Creator ID', 'Item Code')
_sdp_PO02_REQUIRED_FIELDS = ('Company', 'PO Number', 'PO Line', 'Vendor Code', 'PO Doc Date', 'Item Code', 'PO Quantity')

def _sdp_normalize_text(value):
    """Return trimmed text without pandas missing-value placeholders."""
    if value is None or _sdp_pd.isna(value):
        return ''
    text = str(value).strip()
    if text.casefold() in _sdp_INVALID_TEXT_VALUES:
        return ''
    return text

def _sdp_normalize_lookup(value):
    """Return an accent-, punctuation- and case-insensitive lookup key."""
    text = _sdp_unicodedata.normalize('NFKD', _sdp_normalize_text(value))
    text = ''.join((char for char in text if not _sdp_unicodedata.combining(char)))
    return _sdp_re.sub('[^a-z0-9]+', '', text.casefold())

def _sdp_normalize_identifier(value):
    """Normalize an SAP identifier while preserving leading zeroes in text."""
    text = _sdp_normalize_text(value)
    if _sdp_re.fullmatch('\\d+\\.0', text):
        return text[:-2]
    return text

def _sdp_normalize_company(value):
    """Normalize an SAP company code to four digits when it is numeric."""
    text = _sdp_normalize_identifier(value)
    if text.isdigit():
        return text.zfill(4)
    return text.upper()

def _sdp_parse_config_date(value, field_name):
    """Parse one required configuration date."""
    text = _sdp_normalize_text(value)
    if _sdp_re.fullmatch('\\d{4}-\\d{2}-\\d{2}', text):
        parsed = _sdp_pd.to_datetime(text, format='%Y-%m-%d', errors='coerce')
    else:
        parsed = _sdp_pd.to_datetime(value, errors='coerce', dayfirst=True)
    if _sdp_pd.isna(parsed):
        raise ValueError(f'PO CONFIG {field_name} is empty or invalid: {value!r}.')
    return _sdp_pd.Timestamp(parsed).normalize()

def _sdp_get_period_suffix(module_config):
    """Return YYYYMMDD from the PO module TO date."""
    parsed_to = _sdp_parse_config_date(module_config.get('to', ''), 'TO')
    return parsed_to.strftime('%Y%m%d')

def _sdp_parse_config_companies(value):
    """Return normalized company codes or None when COMPANIES is ALL."""
    is_collection = isinstance(value, (list, tuple, set))
    if not is_collection and _sdp_normalize_text(value).upper() == 'ALL':
        return None
    if is_collection:
        raw_values = list(value)
    else:
        raw_values = _sdp_re.split('[,;|\\n]+', _sdp_normalize_text(value))
    companies = {_sdp_normalize_company(item) for item in raw_values if _sdp_normalize_company(item)}
    if not companies:
        raise ValueError('PO CONFIG COMPANIES is empty; use ALL or provide at least one company.')
    return companies

def _sdp_discover_po_input_files(input_folder, module_config):
    """Return period-matching PO Lines files below input_folder."""
    input_folder = _sdp_Path(input_folder)
    if not input_folder.is_dir():
        raise FileNotFoundError(f'PO input folder was not found: {input_folder}')
    period_suffix = _sdp_get_period_suffix(module_config)
    expected_stem = _sdp_normalize_lookup(f'{_sdp_PO_INPUT_PREFIX} {period_suffix}')
    return sorted((path for path in input_folder.rglob('*') if path.is_file() and path.suffix.casefold() in _sdp_ALLOWED_INPUT_EXTENSIONS and (not path.name.startswith('~$')) and (_sdp_normalize_lookup(path.stem) == expected_stem)))

def _sdp_find_period_input_file(context, input_prefix, source_name):
    """
    Encuentra exactamente una bajada del período CONFIG TO.

    La comparación ignora mayúsculas, espacios, guiones y underscores.
    """
    input_folder = _sdp_Path(context['input_folder'])
    if not input_folder.is_dir():
        raise FileNotFoundError(f'{source_name} input folder was not found: {input_folder}')
    period_suffix = _sdp_get_period_suffix(context['module'])
    expected_stem = _sdp_normalize_lookup(f'{input_prefix} {period_suffix}')
    matches = sorted((path for path in input_folder.rglob('*') if path.is_file() and (not path.name.startswith('~$')) and (path.suffix.casefold() in _sdp_ALLOWED_INPUT_EXTENSIONS) and (_sdp_normalize_lookup(path.stem) == expected_stem)))
    expected_name = f"{input_prefix.replace(' ', '_')}_{period_suffix}.xlsx"
    if not matches:
        raise FileNotFoundError(f"{source_name} input was not found. Expected exactly one file equivalent to '{expected_name}' below {input_folder}.")
    if len(matches) > 1:
        details = '\n'.join((f'- {path}' for path in matches))
        raise ValueError(f'Multiple {source_name} inputs were found for {period_suffix}; expected exactly one:\n{details}')
    return matches[0]

def _sdp_find_po_input_file(context):
    """Return the period-specific PO Lines input."""
    return _sdp_find_period_input_file(context=context, input_prefix=_sdp_PO_INPUT_PREFIX, source_name='PO Lines')

def _sdp_resolve_sheet_name(file_path, expected_sheet=_sdp_PO_INPUT_SHEET):
    """Resolve the explicitly configured input sheet case-insensitively."""
    try:
        excel_file = _sdp_pd.ExcelFile(file_path)
    except PermissionError as error:
        raise PermissionError(f'PO input cannot be opened because it is locked: {file_path}. Close Excel and retry.') from error
    expected = _sdp_normalize_lookup(expected_sheet)
    matches = [sheet for sheet in excel_file.sheet_names if _sdp_normalize_lookup(sheet) == expected]
    if len(matches) != 1:
        raise ValueError(f"PO input must contain exactly one sheet named '{expected_sheet}'. Available sheets: {excel_file.sheet_names}")
    return matches[0]

def _sdp_inspect_input_workbook(file_path, sheet_name):
    """Fail on formulas in the selected data sheet and return dimensions."""
    try:
        workbook = _sdp_load_workbook(file_path, read_only=True, data_only=False)
    except PermissionError as error:
        raise PermissionError(f'PO input cannot be read because it is locked: {file_path}. Close Excel and retry.') from error
    except Exception as error:
        raise ValueError(f'PO input workbook is not readable: {file_path}: {error}') from error
    worksheet = workbook[sheet_name]
    formula_count = sum((cell.data_type == 'f' for row in worksheet.iter_rows() for cell in row))
    dimensions = {'physical_rows_including_header': worksheet.max_row, 'physical_columns': worksheet.max_column, 'formula_cells': formula_count}
    workbook.close()
    if formula_count:
        raise ValueError(f"PO input sheet '{sheet_name}' contains {formula_count} formulas. Export stored SAP values instead of formulas.")
    return dimensions

def _sdp_resolve_po_columns(dataframe, required_fields=_sdp_PO01_REQUIRED_FIELDS):
    """Resolve real SAP headers and return logical-to-physical mapping."""
    normalized_headers = {}
    for column in dataframe.columns:
        normalized_headers.setdefault(_sdp_normalize_lookup(column), []).append(column)
    mapping = {}
    missing = []
    ambiguous = []
    for logical_name, aliases in _sdp_PO_LINE_ALIASES.items():
        matches = []
        for alias in aliases:
            matches.extend(normalized_headers.get(_sdp_normalize_lookup(alias), []))
        matches = list(dict.fromkeys(matches))
        if len(matches) == 1:
            mapping[logical_name] = matches[0]
        elif len(matches) > 1:
            ambiguous.append(f'{logical_name}: {matches}')
        elif logical_name in required_fields:
            missing.append(logical_name)
    if missing or ambiguous:
        messages = []
        if missing:
            messages.append(f'missing required fields: {missing}')
        if ambiguous:
            messages.append(f'ambiguous fields: {ambiguous}')
        raise ValueError(f"PO Lines header validation failed ({'; '.join(messages)}). Available headers: {list(dataframe.columns)}")
    return mapping

def _sdp_remove_residual_rows(dataframe, columns):
    """Remove SAP export footer rows and fail on partially populated keys."""
    identity = [columns['Company'], columns['PO Number'], columns['PO Line']]
    blank = _sdp_pd.DataFrame({column: dataframe[column].map(_sdp_normalize_text).eq('') for column in identity}, index=dataframe.index)
    residual_mask = blank.all(axis=1)
    partial_mask = blank.any(axis=1) & ~residual_mask
    if partial_mask.any():
        raise ValueError(f'PO Lines contains {int(partial_mask.sum())} rows with a partially blank Company + PO Number + PO Line key.')
    filtered = dataframe.loc[~residual_mask].copy()
    return (filtered, int(residual_mask.sum()))

def _sdp_filter_companies(dataframe, company_column, companies):
    """Filter rows to CONFIG COMPANIES or retain all when configured ALL."""
    normalized = dataframe[company_column].map(_sdp_normalize_company)
    if companies is None:
        filtered = dataframe.copy()
        filtered[company_column] = normalized
        return (filtered, 0)
    mask = normalized.isin(companies)
    filtered = dataframe.loc[mask].copy()
    filtered[company_column] = normalized.loc[mask]
    return (filtered, int((~mask).sum()))

def _sdp_filter_dates(dataframe, date_column, date_from, date_to):
    """Filter PO document dates inclusively and fail on invalid dates."""
    raw = dataframe[date_column]
    parsed = _sdp_pd.to_datetime(raw, errors='coerce', dayfirst=True).dt.normalize()
    invalid = raw.map(_sdp_normalize_text).ne('') & parsed.isna()
    if invalid.any():
        raise ValueError(f'PO Lines contains {int(invalid.sum())} nonblank invalid PO document dates.')
    mask = parsed.between(date_from, date_to, inclusive='both')
    filtered = dataframe.loc[mask].copy()
    filtered[date_column] = parsed.loc[mask]
    return (filtered, int((~mask).sum()))

def _sdp_normalize_po_dataframe(dataframe, columns):
    """Rename resolved columns and normalize identifiers."""
    renamed = dataframe.rename(columns={physical: logical for logical, physical in columns.items()})
    for column in ('Company', 'PO Number', 'PO Line', 'Vendor Code', 'Vendor Name', 'PO Creator ID', 'Item Code', 'PR Number', 'PR Line'):
        if column in renamed.columns:
            renamed[column] = renamed[column].map(_sdp_normalize_identifier)
    if 'Company' in renamed.columns:
        renamed['Company'] = renamed['Company'].map(_sdp_normalize_company)
    return renamed

def _sdp_load_po_lines(context, required_fields=_sdp_PO01_REQUIRED_FIELDS):
    """Load, validate and CONFIG-filter PO Lines with audit metrics."""
    input_file = _sdp_find_po_input_file(context)
    sheet_name = _sdp_resolve_sheet_name(input_file)
    workbook_metrics = _sdp_inspect_input_workbook(input_file, sheet_name)
    try:
        raw = _sdp_pd.read_excel(input_file, sheet_name=sheet_name, header=_sdp_PO_HEADER_ROW - 1, dtype=object)
    except PermissionError as error:
        raise PermissionError(f'PO input cannot be read because it is locked: {input_file}. Close Excel and retry.') from error
    raw = raw.dropna(axis=0, how='all').dropna(axis=1, how='all')
    if raw.empty:
        raise ValueError(f"PO input sheet '{sheet_name}' contains no data rows.")
    columns = _sdp_resolve_po_columns(raw, required_fields=required_fields)
    population, residual_rows = _sdp_remove_residual_rows(raw, columns)
    if population.empty:
        raise ValueError('PO input contains no valid rows after residual rows were excluded.')
    companies = _sdp_parse_config_companies(context['module'].get('companies', ''))
    date_from = _sdp_parse_config_date(context['module'].get('from', ''), 'FROM')
    date_to = _sdp_parse_config_date(context['module'].get('to', ''), 'TO')
    if date_from > date_to:
        raise ValueError(f'PO CONFIG FROM {date_from.date()} is after TO {date_to.date()}.')
    by_company, excluded_company = _sdp_filter_companies(population, columns['Company'], companies)
    by_date, excluded_date = _sdp_filter_dates(by_company, columns['PO Doc Date'], date_from, date_to)
    normalized = _sdp_normalize_po_dataframe(by_date, columns)
    if companies is None:
        companies_metric = ['ALL']
    else:
        companies_metric = sorted(companies)
    metrics = {**workbook_metrics, 'input_file': str(input_file), 'input_sheet': sheet_name, 'header_row': _sdp_PO_HEADER_ROW, 'rows_read': len(raw), 'residual_rows': residual_rows, 'rows_after_residuals': len(population), 'excluded_by_company': excluded_company, 'excluded_by_date': excluded_date, 'rows_after_config_filters': len(normalized), 'companies': companies_metric, 'date_from': date_from, 'date_to': date_to}
    if normalized.empty:
        raise ValueError(f'PO input has zero rows after CONFIG filters. Rows read={len(raw)}, residual={residual_rows}, excluded company={excluded_company}, excluded date={excluded_date}.')
    return (normalized, metrics)

def _sdp_get_po_output_file(context):
    """Return output/LBR_Results_PO_YYYYMMDD.xlsx."""
    output_folder = _sdp_Path(context['output_folder'])
    period_suffix = _sdp_get_period_suffix(context['module'])
    return output_folder / f'{_sdp_PO_OUTPUT_PREFIX}_{period_suffix}.xlsx'

def _sdp_excel_safe_value(value):
    """Convert pandas/numpy values to safe openpyxl scalars."""
    if value is None or _sdp_pd.isna(value):
        return None
    if isinstance(value, _sdp_pd.Timestamp):
        return value.to_pydatetime()
    if isinstance(value, (_sdp_datetime, _sdp_date, str, int, float, bool)):
        return value
    return str(value)

def _sdp_write_control_sheet(context, sheet_name, dataframe, date_columns=None, amount_columns=None):
    """Replace only one PO sheet while preserving all other sheets."""
    output_file = _sdp_get_po_output_file(context)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    if output_file.exists():
        try:
            workbook = _sdp_load_workbook(output_file)
        except PermissionError as error:
            raise PermissionError(f'PO output is open or locked: {output_file}. Close Excel and retry.') from error
    else:
        workbook = _sdp_Workbook()
        workbook.remove(workbook.active)
    if sheet_name in workbook.sheetnames:
        workbook.remove(workbook[sheet_name])
    worksheet = workbook.create_sheet(sheet_name)
    safe = dataframe.copy()
    for column in safe.columns:
        safe[column] = safe[column].map(_sdp_excel_safe_value)
    for row in _sdp_dataframe_to_rows(safe, index=False, header=True):
        worksheet.append(row)
    header_fill = _sdp_PatternFill(fill_type='solid', fgColor=_sdp_HEADER_FILL)
    for cell in worksheet[1]:
        cell.font = _sdp_Font(bold=True)
        cell.fill = header_fill
    date_columns = set(date_columns or [])
    amount_columns = set(amount_columns or [])
    header_positions = {cell.value: cell.column for cell in worksheet[1]}
    for column in date_columns:
        if column in header_positions:
            column_letter = _sdp_get_column_letter(header_positions[column])
            for cell in worksheet[column_letter][1:]:
                cell.number_format = _sdp_DATE_NUMBER_FORMAT
    for column in amount_columns:
        if column in header_positions:
            column_letter = _sdp_get_column_letter(header_positions[column])
            for cell in worksheet[column_letter][1:]:
                cell.number_format = _sdp_AMOUNT_NUMBER_FORMAT
    for index, column in enumerate(safe.columns, start=1):
        values = [str(column)] + [_sdp_normalize_text(value) for value in safe[column]]
        worksheet.column_dimensions[_sdp_get_column_letter(index)].width = min(max((len(value) for value in values)) + 2, 45)
    temporary_path = None
    try:
        with _sdp_tempfile.NamedTemporaryFile(prefix=f'.{output_file.stem}_', suffix=output_file.suffix, dir=output_file.parent, delete=False) as temporary_file:
            temporary_path = _sdp_Path(temporary_file.name)
        workbook.save(temporary_path)
        _sdp_os.replace(temporary_path, output_file)
    except PermissionError as error:
        raise PermissionError(f'PO output cannot be saved because it is open or locked: {output_file}.') from error
    finally:
        workbook.close()
        if temporary_path is not None and temporary_path.exists():
            temporary_path.unlink()
    return output_file
# SD_PURCHASING_UTILITIES_END


# Shared display-name enrichment for SD controls.
from pathlib import Path as _sdnames_Path
from decimal import Decimal as _sdnames_Decimal, InvalidOperation as _sdnames_InvalidOperation
import hashlib as _sdnames_hashlib
import unicodedata as _sdnames_unicodedata
import pandas as _sdnames_pd
from openpyxl import load_workbook as _sdnames_load_workbook

def _sdnames_text(value):
    if value is None:
        return ''
    try:
        if _sdnames_pd.isna(value):
            return ''
    except (TypeError, ValueError):
        pass
    return ' '.join(str(value).strip().split())

def _sdnames_label(value):
    value = _sdnames_unicodedata.normalize('NFKD', _sdnames_text(value))
    return ''.join((c.lower() for c in value if c.isalnum()))

def _sdnames_identifier(value, width):
    value = _sdnames_text(value)
    if not value:
        return ''
    try:
        number = _sdnames_Decimal(value)
        if number.is_finite() and number == number.to_integral_value():
            value = str(int(number))
    except _sdnames_InvalidOperation:
        pass
    if value.isdigit():
        return '' if int(value) == 0 else value.zfill(width)
    return value.upper()

def _sdnames_hash(path):
    result = _sdnames_hashlib.sha256()
    with path.open('rb') as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b''):
            result.update(block)
    return result.hexdigest()

def _sdnames_find(context, family):
    suffix = get_period_suffix(context['module'])
    expected = _sdnames_label(f'LBR SD {family} {suffix}')
    folder = _sdnames_Path(context['input_folder'])
    matches = sorted((path for path in folder.rglob('*') if path.is_file() and (not path.name.startswith('~$')) and (path.suffix.lower() in {'.xlsx', '.xlsm'}) and (_sdnames_label(path.stem) == expected)))
    if not matches:
        raise FileNotFoundError(f'SD names: missing LBR SD_{family}_{suffix}.xlsx')
    if len(matches) > 1:
        hashes = {_sdnames_hash(path) for path in matches}
        if len(hashes) != 1:
            raise ValueError(f'SD names: multiple different copies of SD_{family}: ' + '; '.join((str(path) for path in matches)))
        print(f'SD names - {family}: {len(matches)} identical copies; one copy used. Keep only one active copy for other loaders.')
    chosen = sorted(matches, key=lambda path: (path.parent.name.casefold() != 'sd', str(path).casefold()))[0]
    print('SD names - Source:', chosen)
    return chosen

def _sdnames_read(path):
    book = _sdnames_load_workbook(path, read_only=True, data_only=True)
    try:
        rows = book.worksheets[0].iter_rows(values_only=True)
        first = next(rows, None)
        if first is None:
            raise ValueError('SD names: empty source ' + path.name)
        headers = [_sdnames_label(value) for value in first]
        for row_number, values in enumerate(rows, 2):
            if any((_sdnames_text(value) for value in values)):
                yield (headers, values, row_number)
    finally:
        book.close()

def _sdnames_get(headers, values, *names, optional=False):
    wanted = {_sdnames_label(name) for name in names}
    indexes = [index for index, name in enumerate(headers) if name in wanted]
    if not indexes:
        if optional:
            return ''
        raise ValueError('SD names: missing source column ' + repr(names))
    index = indexes[0]
    return values[index] if index < len(values) else ''

def _sdnames_vendor_names(context):
    path = _sdnames_find(context, 'VENDOR')
    names = {}
    for headers, values, row_number in _sdnames_read(path):
        company = _sdnames_identifier(_sdnames_get(headers, values, 'CoCd', 'Company Code', 'Company'), 4)
        vendor = _sdnames_identifier(_sdnames_get(headers, values, 'Vendor', 'Vendor Code'), 10)
        name = _sdnames_text(_sdnames_get(headers, values, 'Name 1', 'Vendor Name'))
        if not company and (not vendor):
            continue
        if not company or not vendor:
            raise ValueError(f'SD names: incomplete vendor key in {path.name}, row {row_number}')
        key = (company, vendor)
        previous = names.get(key, '')
        if previous and name and (previous.casefold() != name.casefold()):
            raise ValueError('SD names: contradictory vendor names for ' + repr(key))
        if name:
            names[key] = previous or name
    return names

def _sdnames_user_names(context):
    path = _sdnames_find(context, 'USERS')
    names = {}
    other_versions = 0
    for headers, values, row_number in _sdnames_read(path):
        user = _sdnames_text(_sdnames_get(headers, values, 'User Name', 'BNAME', 'User ID')).upper()
        if not user:
            continue
        version = _sdnames_text(_sdnames_get(headers, values, 'Version', 'NATION', optional=True))
        if version:
            other_versions += 1
            continue
        name = _sdnames_text(_sdnames_get(headers, values, 'Full Name', 'NAME_TEXT', optional=True))
        if not name:
            first = _sdnames_text(_sdnames_get(headers, values, 'First name', 'NAME_FIRST', optional=True))
            last = _sdnames_text(_sdnames_get(headers, values, 'Last name', 'NAME_LAST', optional=True))
            name = ' '.join((part for part in (first, last) if part))
        if not name:
            continue
        previous = names.get(user, '')
        if previous and previous.casefold() != name.casefold():
            raise ValueError('SD names: multiple distinct standard-version names for ' + user + '. Review the source; no name was selected automatically.')
        names[user] = previous or name
    print('SD names - Users with standard-version name:', len(names))
    print('SD names - Nonstandard-version rows excluded:', other_versions)
    print('SD names - Names reflect the supplied user master; historical identity at the transaction date is not established.')
    return names

def enrich_sd_names(context, frame, user_columns):
    """Populate names by keys without merging, duplicating or filtering rows."""
    required = {'Company', 'Vendor Code', *user_columns.keys()}
    missing = required - set(frame.columns)
    if missing:
        raise ValueError('SD names: missing report columns ' + repr(sorted(missing)))
    result = frame.copy()
    added = {'Vendor Name', *user_columns.values()}
    original_columns = [column for column in frame.columns if column not in added]
    if not result.empty:
        vendors = _sdnames_vendor_names(context)
        users = _sdnames_user_names(context)
    else:
        vendors, users = ({}, {})
    vendor_keys = [(_sdnames_identifier(company, 4), _sdnames_identifier(vendor, 10)) for company, vendor in zip(result['Company'], result['Vendor Code'])]
    result['Vendor Name'] = [vendors.get(key, '') if key[1] else '' for key in vendor_keys]
    missing_vendors = sorted({key for key in vendor_keys if key[1] and (not vendors.get(key))})
    print('SD names - Vendor positions with name:', int(result['Vendor Name'].ne('').sum()))
    print('SD names - Vendor keys without name:', len(missing_vendors))
    for key in missing_vendors[:20]:
        print('SD names - Missing vendor:', '|'.join(key))
    for source, target in user_columns.items():
        ids = result[source].map(lambda value: _sdnames_text(value).upper())
        result[target] = ids.map(lambda value: users.get(value, ''))
        missing_users = sorted({value for value in ids if value and (not users.get(value))})
        print(f"SD names - {target}: {int(result[target].ne('').sum())} positions with name; {len(missing_users)} IDs without name")
        if missing_users:
            print('SD names - Missing user IDs:', ', '.join(missing_users[:30]))
    if len(result) != len(frame):
        raise ValueError('SD names: unexpected row count change')
    if not result[original_columns].equals(frame[original_columns]):
        raise ValueError('SD names: original report values changed')
    return result
