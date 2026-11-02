
"""Shared IN source reading, month-end evidence selection and Excel writing."""
import calendar
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from itertools import groupby
import json
from pathlib import Path
import re
import sqlite3
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

QTY = (
    "Unrestricted", "In Qual. Insp.", "Blocked",
    "Stock in tfr", "Restricted", "Returns",
)
VAL = (
    "Total Stock", "Total Value",
    "MovAvgPrice", "Standard price", "per",
)

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"


def txt(value):
    return "" if value is None else str(value).strip()


def material(value):
    # Preserve original spaces, including U+00A0.
    if value is None:
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if int(value) == value:
            return str(int(value))
    return str(value)


def num(value):
    if value is None or isinstance(value, bool) or txt(value) == "":
        raise ValueError(f"Missing/invalid number {value!r}")

    try:
        result = Decimal(str(value).strip())
    except InvalidOperation as exc:
        raise ValueError(
            f"Unrecognized numeric representation {value!r}"
        ) from exc

    if not result.is_finite():
        raise ValueError(f"Non-finite number {value!r}")

    if result == 0:
        return "0"

    formatted = format(result, "f")
    return (
        formatted.rstrip("0").rstrip(".")
        if "." in formatted else formatted
    )


def period(year, month):
    year_value = Decimal(num(year))
    month_value = Decimal(num(month))

    if (
        year_value != int(year_value)
        or month_value != int(month_value)
        or not 1 <= int(year_value) <= 9999
        or not 1 <= int(month_value) <= 12
    ):
        raise ValueError(
            f"Invalid year/period: {year!r}/{month!r}"
        )

    return int(year_value) * 100 + int(month_value)


def cutoff_from(module):
    raw = module.get("to")

    if isinstance(raw, datetime):
        cutoff = raw.date()
    elif isinstance(raw, date):
        cutoff = raw
    else:
        cutoff = datetime.strptime(
            txt(raw), "%Y-%m-%d"
        ).date()

    if cutoff.day != calendar.monthrange(
        cutoff.year, cutoff.month
    )[1]:
        raise ValueError(
            "IN01 currently supports month-end TO only; "
            "no daily balance is inferred."
        )

    return cutoff


def companies_from(value):
    value = txt(value)

    if (
        not value
        or value.lower() == "nan"
        or value.upper() in {"ALL", "*", "TODAS", "TODOS"}
    ):
        return None

    return {
        company.zfill(4)
        for company in re.split(r"[,;\s]+", value)
        if company
    }


def find_sources(folder, cutoff):
    suffix = cutoff.strftime("%Y%m%d")

    prefixes = {
        "MARD": "LBRINMARD",
        "MARDH": "LBRINMARDH",
        "MBEW": "LBRINMBEW",
        "MBEWH": "LBRINMBEWH",
    }

    result = {source: [] for source in prefixes}

    for path in sorted(Path(folder).rglob("*")):
        if (
            not path.is_file()
            or path.name.startswith("~$")
            or path.suffix.lower() not in {".xlsx", ".xlsm"}
        ):
            continue

        name = re.sub(
            r"[ _-]+", "", path.stem
        ).upper()

        for source, prefix in prefixes.items():
            pattern = (
                prefix + suffix + r"(?:PART\d+|\d+)?"
            )
            if re.fullmatch(pattern, name):
                result[source].append(path)

    missing = [
        source
        for source, paths in result.items()
        if not paths
    ]

    if missing:
        raise FileNotFoundError(
            "IN01 missing sources: "
            + ", ".join(missing)
            + f"; cutoff={suffix}; folder={folder}"
        )

    return result


def database(path):
    connection = sqlite3.connect(path)
    connection.execute("PRAGMA temp_store=MEMORY")

    connection.execute(
        "CREATE TABLE snapshot("
        "family TEXT, source TEXT, k TEXT, p INTEGER, "
        "data TEXT, origins TEXT, "
        "PRIMARY KEY(source,k,p)) WITHOUT ROWID"
    )

    connection.execute(
        "CREATE TABLE meta("
        "kind TEXT, k TEXT, v TEXT, origin TEXT, "
        "PRIMARY KEY(kind,k,v)) WITHOUT ROWID"
    )

    connection.execute(
        "CREATE TABLE mapping("
        "kind TEXT, k TEXT PRIMARY KEY, "
        "v TEXT, origin TEXT)"
    )

    return connection


def meta_add(connection, kind, key, value, origin):
    if value:
        connection.execute(
            "INSERT OR IGNORE INTO meta VALUES(?,?,?,?)",
            (kind, key, value, origin),
        )


def meta_get(connection, kind, key):
    rows = connection.execute(
        "SELECT v,origin FROM meta "
        "WHERE kind=? AND k=? ORDER BY v",
        (kind, key),
    ).fetchall()

    if len(rows) == 1:
        return (
            rows[0][0],
            "Supplied master attribute: " + rows[0][1],
        )

    return (
        "",
        "MISSING"
        if not rows
        else "CONFLICTING SUPPLIED MASTER ATTRIBUTES",
    )


def mapping_add(connection, kind, key, value, origin):
    full_key = kind + "|" + key

    previous = connection.execute(
        "SELECT v,origin FROM mapping WHERE k=?",
        (full_key,),
    ).fetchone()

    if previous and previous[0] != value:
        raise ValueError(
            f"IN01 conflicting {kind} mapping {key!r}: "
            f"{previous} versus {value!r} at {origin}"
        )

    connection.execute(
        "INSERT OR IGNORE INTO mapping VALUES(?,?,?,?)",
        (kind, full_key, value, origin),
    )


def load_source(
    connection, source, paths, company_filter, metrics
):
    # IN_SOURCE_LOG_LABEL_V1
    label = str(metrics.get("_control", "IN01"))
    quantity = source.startswith("MARD")
    historical = source.endswith("H")
    family = "Q" if quantity else "V"

    for path in paths:
        print(
            f"{label} reading {source}: {path}",
            flush=True,
        )

        workbook = load_workbook(
            path, read_only=True, data_only=True
        )

        try:
            if "Sheet1" not in workbook.sheetnames:
                raise ValueError(
                    f"{path}: Sheet1 missing; "
                    f"sheets={workbook.sheetnames}"
                )

            rows = workbook["Sheet1"].iter_rows(
                values_only=True
            )

            headers = [
                txt(value) for value in next(rows, ())
            ]

            columns = defaultdict(list)
            for index, heading in enumerate(headers):
                columns[heading].append(index)

            required = [
                "CoCd", "Material", "ValA", "Year", "Pe"
            ]

            required += (
                ["Plnt", "SLoc", *QTY]
                if quantity else
                ["Val. Type", "Pr.", "ValCl", *VAL]
            )

            wrong = [
                f"{heading} "
                f"({len(columns[heading])} occurrences)"
                for heading in required
                if len(columns[heading]) != 1
            ]

            if wrong:
                raise ValueError(
                    f"{path}: missing/ambiguous "
                    f"headers: {wrong}"
                )

            def get(row, heading, last=False):
                indices = columns.get(heading, [])
                if not indices:
                    return None

                index = (
                    indices[-1] if last else indices[0]
                )

                return (
                    row[index]
                    if index < len(row) else None
                )

            def row_error(row_number, message):
                raise ValueError(
                    f"{path.name}, Sheet1, "
                    f"Excel row {row_number}: {message}"
                )

            allowed_totals = (
                set(QTY if quantity else VAL)
                | {"BUn", "Crcy"}
            )

            for row_number, row in enumerate(rows, 2):
                metrics[source + " rows read"] += 1

                if (row_number - 1) % 100000 == 0:
                    print(
                        f"{label} {source}: "
                        f"{row_number - 1:,} rows read",
                        flush=True,
                    )

                if all(txt(value) == "" for value in row):
                    metrics[source + " blank rows"] += 1
                    continue

                key_headers = (
                    ["CoCd", "Material", "ValA"]
                    + (
                        ["Plnt", "SLoc"]
                        if quantity else ["Val. Type"]
                    )
                )

                if historical:
                    key_headers += ["Year", "Pe"]

                empty_keys = all(
                    txt(get(row, heading)) == ""
                    for heading in key_headers
                )

                totals_content = all(
                    txt(value) == ""
                    or (
                        index < len(headers)
                        and headers[index] in allowed_totals
                    )
                    for index, value in enumerate(row)
                )

                if empty_keys and totals_content:
                    metrics[source + " export totals"] += 1
                    continue

                company = txt(get(row, "CoCd"))
                if company.isdigit():
                    company = company.zfill(4)

                material_code = material(
                    get(row, "Material")
                )
                area = txt(get(row, "ValA"))

                plant, storage = (
                    (
                        txt(get(row, "Plnt")),
                        txt(get(row, "SLoc")),
                    )
                    if quantity else ("", "")
                )

                valuation_type = (
                    txt(get(row, "Val. Type"))
                    if not quantity else ""
                )

                if (
                    not company
                    or not material_code.strip()
                    or not area
                    or (
                        quantity
                        and (not plant or not storage)
                    )
                ):
                    row_error(
                        row_number,
                        "Incomplete key; valuation type "
                        "may be blank, material/area/"
                        "company may not.",
                    )

                if (
                    company_filter is not None
                    and company not in company_filter
                ):
                    metrics[
                        source + " excluded company"
                    ] += 1
                    continue

                try:
                    source_period = period(
                        get(row, "Year"), get(row, "Pe")
                    )

                    numbers = {
                        heading: num(get(row, heading))
                        for heading in (
                            QTY if quantity else VAL
                        )
                    }
                except ValueError as error:
                    row_error(row_number, str(error))

                fiscal_variant = txt(get(row, "FV"))
                if (
                    fiscal_variant
                    and fiscal_variant != "K4"
                ):
                    row_error(
                        row_number,
                        "Unsupported fiscal variant "
                        f"{fiscal_variant}; month/calendar "
                        "correspondence not assumed.",
                    )

                unit = txt(get(row, "BUn", True))
                currency = txt(get(row, "Crcy", True))

                for heading in ("BUn", "Crcy"):
                    supplied = {
                        txt(row[index])
                        for index in columns.get(
                            heading, []
                        )
                        if index < len(row)
                        and txt(row[index])
                    }

                    if len(supplied) > 1:
                        row_error(
                            row_number,
                            f"Inconsistent {heading} "
                            f"occurrences: {supplied}",
                        )

                relative_path = path.relative_to(
                    Path(metrics["_input"])
                )

                origin = (
                    f"{relative_path}|Sheet1|{row_number}"
                )

                logical_key = (
                    [
                        company, material_code,
                        plant, storage,
                    ]
                    if quantity else
                    [
                        company, material_code,
                        area, valuation_type,
                    ]
                )

                key = json.dumps(
                    logical_key, ensure_ascii=False
                )

                data = {
                    "company": company,
                    "material": material_code,
                    "area": area,
                    "plant": plant,
                    "sloc": storage,
                    "vt": valuation_type,
                    "period": source_period,
                    "unit": unit,
                    "currency": currency,
                    "fv": fiscal_variant,
                    "numbers": numbers,
                    "price_control": txt(
                        get(row, "Pr.")
                    ),
                    "valuation_class": txt(
                        get(row, "ValCl")
                    ),
                    "flag": (
                        txt(get(row, "SL"))
                        if quantity
                        else txt(get(row, "VTy"))
                    ),
                }

                payload = json.dumps(
                    data,
                    ensure_ascii=False,
                    sort_keys=True,
                )

                stored_period = (
                    source_period if historical else 0
                )

                previous = connection.execute(
                    "SELECT data,origins FROM snapshot "
                    "WHERE source=? AND k=? AND p=?",
                    (source, key, stored_period),
                ).fetchone()

                if previous:
                    if previous[0] != payload:
                        row_error(
                            row_number,
                            f"Contradictory key {key}; "
                            "previous source rows: "
                            f"{previous[1]}",
                        )

                    origins = json.loads(previous[1])
                    origins.append(origin)

                    connection.execute(
                        "UPDATE snapshot SET origins=? "
                        "WHERE source=? AND k=? AND p=?",
                        (
                            json.dumps(
                                origins,
                                ensure_ascii=False,
                            ),
                            source, key, stored_period,
                        ),
                    )

                    metrics[
                        source + " identical duplicates"
                    ] += 1
                else:
                    connection.execute(
                        "INSERT INTO snapshot "
                        "VALUES(?,?,?,?,?,?)",
                        (
                            family, source, key,
                            stored_period, payload,
                            json.dumps(
                                [origin],
                                ensure_ascii=False,
                            ),
                        ),
                    )

                    metrics[
                        source + " distinct records"
                    ] += 1

                meta_add(
                    connection, "unit",
                    material_code, unit, origin,
                )
                meta_add(
                    connection, "currency",
                    company, currency, origin,
                )
                mapping_add(
                    connection, "area-company",
                    area, company, origin,
                )

                if quantity:
                    mapping_add(
                        connection,
                        "plant-area-company",
                        plant,
                        json.dumps([area, company]),
                        origin,
                    )

                metrics[
                    source + " material contains NBSP"
                ] += int("\u00a0" in material_code)

                metrics[source + " scoped rows"] += 1

            connection.commit()

        finally:
            workbook.close()


def choose(records, target):
    """Select evidence for a month-end without filling missing history."""
    current = [
        record for record in records
        if not record[0].endswith("H")
    ]

    history = sorted(
        (
            record for record in records
            if record[0].endswith("H")
        ),
        key=lambda record: record[1]["period"],
    )

    if not current:
        return (
            None,
            "PENDING - HISTORY WITHOUT CURRENT RECORD",
            history[0] if history else None,
        )

    current_record = current[0]
    current_period = current_record[1]["period"]

    if any(
        record[1]["period"] >= current_period
        for record in history
    ):
        return (
            None,
            "PENDING - INCONSISTENT CURRENT/HISTORY PERIODS",
            current_record,
        )

    if current_period <= target:
        return (
            current_record,
            "CURRENT RECORD - NO LATER BALANCE UPDATE RECORDED",
            current_record,
        )

    eligible = [
        record for record in history
        if target <= record[1]["period"] < current_period
    ]

    if eligible:
        return (
            eligible[0],
            "HISTORICAL PERIOD-END SNAPSHOT",
            current_record,
        )

    return (
        None,
        "PENDING - NO HISTORY COVERING CUTOFF",
        current_record,
    )


def records_by_key(connection, family):
    cursor = connection.execute(
        "SELECT k,source,data,origins FROM snapshot "
        "WHERE family=? ORDER BY k,source,p",
        (family,),
    )

    for key, group in groupby(
        cursor, key=lambda row: row[0]
    ):
        records = [
            (
                row[1],
                json.loads(row[2]),
                json.loads(row[3]),
            )
            for row in group
        ]
        yield key, records


def _tag(name):
    return "{" + NS + "}" + name


def _xml_preserve_namespaces(tree, original):
    declared = dict(re.findall(
        r'xmlns(?::([A-Za-z_][\w.-]*))?="([^"]+)"',
        original.decode("utf-8"),
    ))

    used = set()
    for node in tree.iter():
        for name in (node.tag, *node.attrib):
            if name.startswith("{"):
                used.add(
                    name[1:].split("}", 1)[0]
                )

    for prefix, uri in declared.items():
        if not re.fullmatch(r"ns\d+", prefix):
            ET.register_namespace(prefix, uri)

        if prefix and uri not in used:
            tree.set("xmlns:" + prefix, uri)

    return ET.tostring(
        tree,
        encoding="utf-8",
        xml_declaration=True,
    )


def _styles(tree):
    def child(name):
        node = tree.find(_tag(name))
        if node is None:
            node = ET.SubElement(
                tree, _tag(name), {"count": "0"}
            )
        return node

    fonts = child("fonts")
    fills = child("fills")
    formats = child("cellXfs")

    font_id = len(fonts)
    fill_id = len(fills)

    font = ET.SubElement(fonts, _tag("font"))
    ET.SubElement(font, _tag("b"))
    ET.SubElement(
        font, _tag("sz"), {"val": "11"}
    )
    ET.SubElement(
        font, _tag("name"), {"val": "Calibri"}
    )

    fill = ET.SubElement(fills, _tag("fill"))
    pattern = ET.SubElement(
        fill,
        _tag("patternFill"),
        {"patternType": "solid"},
    )

    ET.SubElement(
        pattern,
        _tag("fgColor"),
        {"rgb": "FFD9EAF7"},
    )
    ET.SubElement(
        pattern,
        _tag("bgColor"),
        {"indexed": "64"},
    )

    number_formats = child("numFmts")
    existing = {
        node.get("formatCode"): node.get("numFmtId")
        for node in number_formats
    }

    format_code = "#,##0.000000"
    number_id = existing.get(format_code)

    if number_id is None:
        number_id = str(max(
            [163] + [
                int(node.get("numFmtId"))
                for node in number_formats
            ]
        ) + 1)

        ET.SubElement(
            number_formats,
            _tag("numFmt"),
            {
                "numFmtId": number_id,
                "formatCode": format_code,
            },
        )

    first_style = len(formats)

    for font_value, fill_value, number_value in (
        (font_id, fill_id, "0"),
        (0, 0, number_id),
        (0, 0, "0"),
    ):
        node = ET.SubElement(
            formats,
            _tag("xf"),
            {
                "numFmtId": str(number_value),
                "fontId": str(font_value),
                "fillId": str(fill_value),
                "borderId": "0",
                "xfId": "0",
                "applyAlignment": "1",
            },
        )

        ET.SubElement(
            node,
            _tag("alignment"),
            {"vertical": "top"},
        )

    for node in (
        fonts, fills, formats, number_formats
    ):
        node.set("count", str(len(node)))

    tree.remove(number_formats)
    tree.insert(0, number_formats)

    return (
        first_style,
        first_style + 1,
        first_style + 2,
    )


def write_sheet(
    stream, headers, rows, visible_count,
    styles, notes_factory,
):
    header_style, numeric_style, note_style = styles

    def write(value):
        stream.write(value.encode("utf-8"))

    def cell(value, column, row_number, style=None):
        if value is None:
            return ""

        reference = (
            get_column_letter(column)
            + str(row_number)
        )
        style_attribute = (
            f' s="{style}"'
            if style is not None else ""
        )

        if (
            isinstance(value, (Decimal, int, float))
            and not isinstance(value, bool)
        ):
            return (
                f'<c r="{reference}"'
                f'{style_attribute}>'
                f'<v>{value}</v></c>'
            )

        value = str(value)

        if len(value) > 32767:
            raise ValueError(
                f"Excel text exceeds 32767 characters "
                f"at {reference}; trace was not truncated."
            )

        if re.search(
            r"[\x00-\x08\x0B\x0C\x0E-\x1F]", value
        ):
            raise ValueError(
                "Invalid XML control character at "
                f"{reference}; text was not changed."
            )

        return (
            f'<c r="{reference}" t="inlineStr"'
            f'{style_attribute}><is>'
            '<t xml:space="preserve">'
            f'{escape(value)}</t></is></c>'
        )

    write(
        '<?xml version="1.0" encoding="utf-8"?>'
        f'<worksheet xmlns="{NS}">'
    )

    write(
        '<sheetPr><outlinePr summaryRight="0"/>'
        '</sheetPr>'
        '<sheetViews><sheetView workbookViewId="0">'
        '<pane ySplit="1" topLeftCell="A2" '
        'activePane="bottomLeft" state="frozen"/>'
        '<selection pane="bottomLeft" '
        'activeCell="A2" sqref="A2"/>'
        '</sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15" '
        'outlineLevelCol="1"/><cols>'
    )

    for index, heading in enumerate(headers, 1):
        width = (
            38 if "Evidence" in heading
            else 28 if heading == "Material"
            else min(30, max(15, len(heading) + 2))
        )

        extra = (
            ' hidden="1" outlineLevel="1"'
            if index > visible_count else ""
        )

        if index == visible_count:
            extra += ' collapsed="1"'

        write(
            f'<col min="{index}" max="{index}" '
            f'width="{width}" customWidth="1"'
            f'{extra}/>'
        )

    write('</cols><sheetData><row r="1">')
    write("".join(
        cell(heading, index, 1, header_style)
        for index, heading in enumerate(headers, 1)
    ))
    write("</row>")

    row_number = 1

    for values in rows:
        row_number += 1

        if row_number >= 1048550:
            raise ValueError(
                "IN01 exceeds one worksheet; "
                "output was not truncated or replaced."
            )

        write(f'<row r="{row_number}">')

        write("".join(
            cell(
                value, index, row_number,
                numeric_style
                if isinstance(value, Decimal)
                else None,
            )
            for index, value in enumerate(values, 1)
        ))

        write("</row>")

        if row_number % 100000 == 0:
            print(
                f"IN01 writing: "
                f"{row_number - 1:,} records",
                flush=True,
            )

    last_data = row_number

    if row_number == 1:
        row_number += 1
        message = (
            "No supplied records for configured "
            "companies. This is not proof "
            "of zero inventory."
        )
        write(
            f'<row r="{row_number}">'
            f'{cell(message, 1, row_number)}'
            '</row>'
        )

    for note in notes_factory():
        row_number += (
            2 if row_number == last_data else 1
        )

        if row_number > 1048576:
            raise ValueError(
                "No space for IN01 scope notes."
            )

        write(
            f'<row r="{row_number}">'
            f'{cell(note, 1, row_number, note_style)}'
            '</row>'
        )

    write(
        '</sheetData>'
        f'<autoFilter ref="A1:'
        f'{get_column_letter(len(headers))}'
        f'{max(1, last_data)}"/>'
        '</worksheet>'
    )

    return last_data - 1


# IN_BILLING_IN_COMMON_V1
"""Shared preliminary billing analysis for IN07 and IN08."""
from collections import Counter as _in_billing_Counter, defaultdict as _in_billing_defaultdict
from datetime import date as _in_billing_date, datetime as _in_billing_datetime
from decimal import Decimal as _in_billing_Decimal
import json as _in_billing_json
import math as _in_billing_math
import os as _in_billing_os
from pathlib import Path as _in_billing_Path
import re as _in_billing_re
import tempfile as _in_billing_tempfile
from time import perf_counter as _in_billing_perf_counter
from xml.etree import ElementTree as _in_billing_ET
import zipfile as _in_billing_zipfile
from openpyxl import Workbook as _in_billing_Workbook, load_workbook as _in_billing_load_workbook
from openpyxl.utils.datetime import from_excel as _in_billing_from_excel
_in_billing_CATEGORY = {'M': 'M - Invoice', 'N': 'N - Invoice cancellation', 'O': 'O - Credit memo', 'P': 'P - Debit memo', 'S': 'S - Credit memo cancellation', '5': '5 - Intercompany invoice', '6': '6 - Intercompany credit'}

def _in_billing_text(value):
    return '' if value is None else str(value).strip()

def _in_billing_code(value, width=0, optional=False):
    if isinstance(value, (int, float)) and (not isinstance(value, bool)):
        if not _in_billing_math.isfinite(value) or int(value) != value:
            raise ValueError(f'Invalid identifier: {value!r}')
        value = str(int(value))
    value = _in_billing_text(value)
    if optional and (not value or not value.strip('0')):
        return ''
    return value.zfill(width) if value.isdigit() else value

def _in_billing_material(value):
    if isinstance(value, (int, float)) and (not isinstance(value, bool)):
        if _in_billing_math.isfinite(value) and int(value) == value:
            return str(int(value))
    return '' if value is None else str(value)

def _in_billing_numeric(value):
    if value is None or isinstance(value, bool) or (not _in_billing_text(value)):
        raise ValueError(f'Missing/invalid number: {value!r}')
    result = _in_billing_Decimal(str(value))
    if not result.is_finite():
        raise ValueError(f'Non-finite number: {value!r}')
    return result

def _in_billing_day(value, epoch=None):
    if isinstance(value, _in_billing_datetime):
        return value.date()
    if isinstance(value, _in_billing_date):
        return value
    if isinstance(value, (int, float)) and (not isinstance(value, bool)) and (epoch is not None):
        result = _in_billing_from_excel(value, epoch)
        if isinstance(result, _in_billing_datetime):
            return result.date()
    for pattern in ('%Y-%m-%d', '%d/%m/%Y', '%d.%m.%Y', '%Y%m%d', '%Y-%m-%d %H:%M:%S'):
        try:
            return _in_billing_datetime.strptime(_in_billing_text(value), pattern).date()
        except ValueError:
            pass
    raise ValueError(f'Missing/invalid date: {value!r}')

def _in_billing_load(context, cutoff, scope, kind, label, metrics):
    folder = _in_billing_Path(context['input_folder'])
    paths = sorted((p for p in folder.rglob('*') if p.is_file() and (not p.name.startswith('~$')) and (p.suffix.lower() in {'.xlsx', '.xlsm'}) and _in_billing_re.fullmatch('LBRSDBILLING' + cutoff.strftime('%Y%m%d') + '(?:PART\\d+|\\d+)?', _in_billing_re.sub('[ _-]+', '', p.stem).upper())))
    if not paths:
        raise FileNotFoundError(f'{label}: missing SD_BILLING for {cutoff:%Y%m%d}; folder={folder}')
    records, documents, signatures = ({}, {}, {})
    required = ['CoCd', 'Bill.Doc.', 'BillT', 'DocCa', 'Created by', 'Created on', 'Time', 'Billing Date', 'Payer', 'Sold-to pt', 'SOrg.', 'Can', 'CancBillDc', 'PsSt', 'DocumentNo', 'Year', 'Doc.cond.', 'Sales Doc.', 'Ref.doc.', 'RefItm', 'Material', 'Description', 'Billed Quantity', 'Net value', 'Tax amount']
    for path in paths:
        print(f'{label} reading: {path}', flush=True)
        book = _in_billing_load_workbook(path, read_only=True, data_only=True)
        try:
            rows = book['Sheet1'].iter_rows(values_only=True)
            headers = [_in_billing_text(v) for v in next(rows, ())]
            columns = _in_billing_defaultdict(list)
            for i, heading in enumerate(headers):
                columns[heading].append(i)
            wrong = [h for h in required if len(columns[h]) != 1]
            wrong += [h for h, n in (('Item', 2), ('SU', 2), ('Curr.', 3)) if len(columns[h]) != n]
            if wrong:
                raise ValueError(f'{path}: missing/ambiguous headers: {wrong}; observed={headers}')

            def value(row, h, occurrence=0):
                i = columns[h][occurrence]
                return row[i] if i < len(row) else None
            for rownum, row in enumerate(rows, 2):
                metrics['source rows read'] += 1
                origin = f'{path.relative_to(folder)}|Sheet1|{rownum}'
                try:
                    if not any((_in_billing_text(v) for v in row)):
                        metrics['blank rows'] += 1
                        continue
                    identity = [value(row, h) for h in ('CoCd', 'Bill.Doc.', 'BillT', 'DocCa', 'Material', 'Billing Date')] + [value(row, 'Item', 1)]
                    if not any((_in_billing_text(v) for v in identity)) and all((not _in_billing_text(v) or (i < len(headers) and headers[i] in {'Billed Quantity', 'SU', 'Net value', 'Curr.', 'Tax amount'}) for i, v in enumerate(row))):
                        metrics['export totals'] += 1
                        continue
                    company = _in_billing_code(value(row, 'CoCd'), 4)
                    document = _in_billing_code(value(row, 'Bill.Doc.'), 10)
                    item = _in_billing_code(value(row, 'Item', 1), 6)
                    if not company or any((not v.isdigit() or int(v) <= 0 for v in (document, item))):
                        raise ValueError('Incomplete/invalid company/billing document/billing item key.')
                    if scope is not None and company not in scope:
                        metrics['outside configured companies'] += 1
                        continue
                    billtype = _in_billing_code(value(row, 'BillT'))
                    category = _in_billing_code(value(row, 'DocCa'))
                    mat = _in_billing_material(value(row, 'Material'))
                    if not billtype or not category or (not mat.strip()):
                        raise ValueError('Missing billing type, category or material; row was not excluded.')
                    units = {_in_billing_text(value(row, 'SU', i)) for i in range(2)} - {''}
                    currencies = {_in_billing_text(value(row, 'Curr.', i)) for i in range(3)} - {''}
                    if len(units) > 1 or len(currencies) > 1:
                        raise ValueError('Contradictory repeated SU/Curr. values.')
                    unit = next(iter(units), '')
                    currency = next(iter(currencies), '')
                    if kind == 'Q' and (not unit) or (kind == 'V' and (not currency)):
                        raise ValueError('Missing aggregation unit/currency; no total was inferred.')
                    quantity, amount, tax = (_in_billing_numeric(value(row, h)) for h in ('Billed Quantity', 'Net value', 'Tax amount'))
                    billed = _in_billing_day(value(row, 'Billing Date'), book.epoch)
                    created = _in_billing_day(value(row, 'Created on'), book.epoch)
                    reference = _in_billing_code(value(row, 'CancBillDc'), 10, True)
                    cancel = _in_billing_text(value(row, 'Can'))
                    status = _in_billing_text(value(row, 'PsSt'))
                    key = (company, document, item)
                    header = (billtype, category, billed, created, reference, cancel, status, currency, _in_billing_text(value(row, 'SOrg.')), _in_billing_text(value(row, 'Payer')), _in_billing_text(value(row, 'Sold-to pt')), _in_billing_text(value(row, 'Created by')), _in_billing_text(value(row, 'Time')))
                    dk = (company, document)
                    if dk in documents and documents[dk]['header'] != header:
                        raise ValueError(f"Contradictory billing header {dk}; previous={documents[dk]['origin']}")
                    documents.setdefault(dk, {'header': header, 'reference': reference, 'origin': origin})
                    if key in records:
                        if signatures[key] != tuple(row):
                            raise ValueError(f"Contradictory billing item {key}; previous={records[key]['origins']}")
                        records[key]['origins'].append(origin)
                        metrics['identical duplicates'] += 1
                        continue
                    records[key] = {'company': company, 'document': document, 'item': item, 'billtype': billtype, 'category': category, 'material': mat, 'description': _in_billing_text(value(row, 'Description')), 'unit': unit, 'currency': currency, 'quantity': quantity, 'amount': amount, 'billed': billed, 'created': created, 'reference': reference, 'cancel': cancel, 'status': status, 'origins': [origin]}
                    signatures[key] = tuple(row)
                    metrics['distinct source positions'] += 1
                except Exception as error:
                    raise ValueError(f'{origin}: {error}') from error
        finally:
            book.close()
    return (records, documents, paths)

def _in_billing_accumulator():
    return {'value': _in_billing_Decimal(0), 'positions': 0, 'docs': set(), 'materials': set(), 'types': set(), 'descriptions': set(), 'units': set(), 'currencies': set(), 'statuses': set(), 'cancel': 0, 'references': 0, 'gaps': set(), 'rows': []}

def _in_billing_aggregate(records, documents, start, cutoff, scope, kind, metrics, label):
    summaries, details, missing_refs = ({}, {}, set())
    for key, row in sorted(records.items()):
        if not start <= row['billed'] <= cutoff:
            metrics['outside configured billing dates'] += 1
            continue
        metrics['evaluated source positions'] += 1
        metrics['zero net-value positions'] += int(row['amount'] == 0)
        metrics['negative quantity positions'] += int(row['quantity'] < 0)
        metrics['negative net-value positions'] += int(row['amount'] < 0)
        axis = row['unit'] if kind == 'Q' else row['currency']
        measure = row['quantity'] if kind == 'Q' else row['amount']
        sk = (row['company'], axis, row['category'])
        dk = (row['company'], row['billed'].strftime('%Y-%m'), row['material'], axis, row['category'], row['billtype'])
        has_gap = bool(row['reference'] and (row['company'], row['reference']) not in documents)
        if has_gap:
            missing_refs.add((row['company'], row['document'], row['reference']))
        for mapping, group, trace in ((summaries, sk, False), (details, dk, True)):
            acc = mapping.setdefault(group, _in_billing_accumulator())
            acc['value'] += measure
            acc['positions'] += 1
            acc['docs'].add((row['company'], row['document']))
            acc['materials'].add(row['material'])
            acc['types'].add(row['billtype'])
            acc['descriptions'].add(row['description'])
            acc['units'].add(row['unit'])
            acc['currencies'].add(row['currency'])
            acc['statuses'].add(row['status'])
            acc['cancel'] += int(bool(row['cancel']))
            acc['references'] += int(bool(row['reference']))
            if has_gap:
                acc['gaps'].add((row['company'], row['document'], row['reference']))
            if trace:
                acc['rows'].append([row['document'], row['item'], row['billed'].isoformat(), row['created'].isoformat(), str(row['quantity']), str(row['amount']), row['cancel'], row['reference'], row['status'], row['origins']])
    for company, document, reference in sorted(missing_refs):
        print(f"{label} unmatched cancellation reference: company={company}; billing document={document}; referenced document={reference}; origin={documents[company, document]['origin']}", flush=True)
    metrics['unmatched reference documents'] = len(missing_refs)
    metrics['summary groups'] = len(summaries)
    metrics['material-month-type groups'] = len(details)
    trace_parts = {}
    for key, acc in details.items():
        raw = _in_billing_json.dumps(acc['rows'], ensure_ascii=True, separators=(',', ':'))
        trace_parts[key] = [raw[i:i + 30000] for i in range(0, len(raw), 30000)]
    chunks = max([1] + [len(parts) for parts in trace_parts.values()])
    return (summaries, details, trace_parts, chunks)

def _in_billing_category_label(value):
    return _in_billing_CATEGORY.get(value, value + ' - Other category (not classified)')

def _in_billing_summary_rows(groups, start, cutoff, scope, sources):
    emitted = set()
    for (company, axis, category), acc in sorted(groups.items()):
        emitted.add(company)
        yield [company, axis, _in_billing_category_label(category), acc['value'], len(acc['docs']), acc['positions'], len(acc['materials']), start.isoformat(), cutoff.isoformat(), _in_billing_json.dumps(sorted(acc['types'])), len(acc['gaps']), _in_billing_json.dumps(sources), 'PRELIMINARY BILLING ACTIVITY - INVENTORY SALES CLASSIFICATION PENDING']
    for company in sorted((scope or ({'ALL'} if not groups else set())) - emitted):
        yield [company, '', 'NO SUPPLIED BILLING POSITIONS IN SCOPE', None, 0, 0, 0, start.isoformat(), cutoff.isoformat()]

def _in_billing_detail_rows(groups, traces, chunks):
    if not groups:
        yield ['No supplied billing positions meet the configured companies and Billing Date period. This does not establish absence of inventory sales.']
    for key, acc in sorted(groups.items()):
        company, month, mat, axis, category, billtype = key
        descriptions = sorted(acc['descriptions'])
        description = descriptions[0] if len(descriptions) == 1 else 'MULTIPLE SOURCE DESCRIPTIONS - EXPAND TECHNICAL FIELDS'
        origin = acc['rows'][0][-1][0]
        file, sheet, row = origin.rsplit('|', 2)
        parts = traces[key]
        yield [company, month, mat, description, axis, _in_billing_category_label(category), billtype, acc['value'], len(acc['docs']), acc['positions'], _in_billing_json.dumps(key, ensure_ascii=False), _in_billing_json.dumps(descriptions, ensure_ascii=False), _in_billing_json.dumps(sorted(acc['units'])), _in_billing_json.dumps(sorted(acc['currencies'])), acc['cancel'], acc['references'], _in_billing_json.dumps(sorted(acc['gaps'])), _in_billing_json.dumps(sorted(acc['statuses'])), file, sheet, int(row), len(parts), *parts, *[None] * (chunks - len(parts))]

class _in_billing_PlainLayout:

    def __init__(self, stream, summary):
        self.stream, self.summary = (stream, summary)

    def write(self, data):
        value = data.decode('utf-8')

        def cell(match):
            attrs = match[1]
            ref = _in_billing_re.search('\\br="[A-Z]+(\\d+)"', attrs)
            if ref and int(ref[1]) != 1:
                attrs = _in_billing_re.sub('\\s+s="\\d+"', '', attrs)
            return '<c' + attrs + '>'
        value = _in_billing_re.sub('<c(\\s[^>]+)>', cell, value)
        if value.startswith('<col '):
            col = _in_billing_ET.fromstring(value)
            i = int(col.get('min'))
            widths = {1: 12, 2: 16, 3: 36, 4: 26, 5: 22, 6: 20, 7: 20} if self.summary else {1: 12, 2: 14, 3: 30, 4: 44, 5: 16, 6: 36, 7: 16, 8: 26, 9: 22, 10: 20}
            if i in widths:
                col.set('width', str(widths[i]))
            value = _in_billing_ET.tostring(col, encoding='unicode')
        return self.stream.write(value.encode('utf-8'))

def _in_billing__render(context, cutoff, tables):
    SHEETS = tuple((table[0] for table in tables))
    output = _in_billing_Path(context['output_folder']) / f'LBR_Results_IN_{cutoff:%Y%m%d}.xlsx'
    output.parent.mkdir(parents=True, exist_ok=True)
    with _in_billing_tempfile.TemporaryDirectory(prefix='in_billing_output_', dir=output.parent) as temporary:
        temporary = _in_billing_Path(temporary)
        base = output
        if not output.exists():
            base = temporary / 'base.xlsx'
            book = _in_billing_Workbook(write_only=True)
            for name in SHEETS:
                book.create_sheet(name)
            book.save(base)
        new = temporary / 'result.xlsx'
        with _in_billing_zipfile.ZipFile(base) as original:
            workbook = _in_billing_ET.fromstring(original.read('xl/workbook.xml'))
            relationships = _in_billing_ET.fromstring(original.read('xl/_rels/workbook.xml.rels'))
            types = _in_billing_ET.fromstring(original.read('[Content_Types].xml'))
            styles = _in_billing_ET.fromstring(original.read('xl/styles.xml'))
            style_ids = _styles(styles)
            for style_id in style_ids:
                styles.find(_tag('cellXfs'))[style_id].find(_tag('alignment')).set('wrapText', '0')
            sheets = workbook.find(_tag('sheets'))
            paths = {}
            for name in SHEETS:
                existing = next((s for s in sheets if s.get('name') == name), None)
                if existing is not None:
                    relation = next((r for r in relationships if r.get('Id') == existing.get('{' + REL + '}id')))
                    target = relation.get('Target')
                    paths[name] = target.lstrip('/') if target.startswith('/') else 'xl/' + target
                else:
                    index = 1
                    while f'xl/worksheets/in_billing_{index}.xml' in original.namelist() or f'xl/worksheets/in_billing_{index}.xml' in paths.values():
                        index += 1
                    path = f'xl/worksheets/in_billing_{index}.xml'
                    rid = 1
                    while any((r.get('Id') == f'rId{rid}' for r in relationships)):
                        rid += 1
                    sheetid = max([0] + [int(s.get('sheetId')) for s in sheets]) + 1
                    _in_billing_ET.SubElement(sheets, _tag('sheet'), {'name': name, 'sheetId': str(sheetid), '{' + REL + '}id': f'rId{rid}'})
                    _in_billing_ET.SubElement(relationships, '{' + PKG + '}Relationship', {'Id': f'rId{rid}', 'Type': REL + '/worksheet', 'Target': path[3:]})
                    _in_billing_ET.SubElement(types, '{' + CT + '}Override', {'PartName': '/' + path, 'ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'})
                    paths[name] = path
            replacements = {'xl/workbook.xml': _xml_preserve_namespaces(workbook, original.read('xl/workbook.xml')), 'xl/_rels/workbook.xml.rels': _xml_preserve_namespaces(relationships, original.read('xl/_rels/workbook.xml.rels')), '[Content_Types].xml': _xml_preserve_namespaces(types, original.read('[Content_Types].xml')), 'xl/styles.xml': _xml_preserve_namespaces(styles, original.read('xl/styles.xml'))}
            with _in_billing_zipfile.ZipFile(new, 'w', compression=_in_billing_zipfile.ZIP_DEFLATED, compresslevel=3, allowZip64=True) as destination:
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
                for name, headers, rows, visible, footer in tables:
                    with destination.open(paths[name], 'w', force_zip64=True) as stream:
                        write_sheet(_in_billing_PlainLayout(stream, name == SHEETS[0]), headers, rows, visible, style_ids, footer)
        _in_billing_os.replace(new, output)
    return output

def run_in_billing(context, kind):
    if kind not in {'Q', 'V'}:
        raise ValueError('Unsupported billing analysis kind.')
    label = 'IN07' if kind == 'Q' else 'IN08'
    started = _in_billing_perf_counter()
    start = _in_billing_day(context['module']['from'])
    cutoff = _in_billing_day(context['module']['to'])
    if start > cutoff:
        raise ValueError(label + ': FROM is after TO.')
    raw = _in_billing_text(context['module'].get('companies'))
    scope = None if not raw or raw.upper() in {'ALL', '*', 'TODAS', 'TODOS', 'NAN'} else {_in_billing_code(v, 4) for v in _in_billing_re.split('[,;\\s]+', raw) if v}
    metrics = _in_billing_Counter()
    print(label + ' - PRELIMINARY BILLING ACTIVITY; inventory-sales classification pending.', flush=True)
    stage = _in_billing_perf_counter()
    records, documents, paths = _in_billing_load(context, cutoff, scope, kind, label, metrics)
    print(f'{label} time | read/validate: {_in_billing_perf_counter() - stage:.2f} seconds', flush=True)
    stage = _in_billing_perf_counter()
    summaries, details, traces, chunks = _in_billing_aggregate(records, documents, start, cutoff, scope, kind, metrics, label)
    print(f'{label} time | aggregate: {_in_billing_perf_counter() - stage:.2f} seconds', flush=True)
    axis = 'Unit' if kind == 'Q' else 'Currency'
    measure = 'Billed Quantity (as supplied)' if kind == 'Q' else 'Billing Net Value (as supplied)'
    summary_headers = ['Company', axis, 'Document Category', measure, 'Billing Documents', 'Source Positions', 'Materials']
    summary_tech = ['From', 'To', 'Billing Types', 'Unmatched Reference Documents', 'Source Files', 'Scope']
    detail_headers = ['Company', 'Month', 'Material', 'Description', axis, 'Document Category', 'Billing Type', measure, 'Billing Documents', 'Source Positions']
    detail_tech = ['Group Key', 'Source Descriptions', 'Source Units', 'Source Currencies', 'Canceled Source Positions', 'Positions With Cancellation Reference', 'Unmatched Cancellation References', 'Posting Statuses', 'Source File', 'Source Sheet', 'Source Row', 'Trace Parts'] + [f'Trace {i + 1}' for i in range(chunks)]
    if len(detail_headers) + len(detail_tech) > 16384:
        raise ValueError(label + ': trace exceeds Excel column limit; output was not replaced.')
    note = [f'PRELIMINARY BILLING ACTIVITY, not confirmed inventory sales. Billing Date: {start.isoformat()} to {cutoff.isoformat()}. Complete SAP coverage is not certified; the documented extraction is selected by Created on.', 'Values remain as supplied, separated by category/type and unit/currency. No cancellation netting or automatic exclusions of billing types or SV units. Commercial classification is pending.', f"Unmatched cancellation references: {metrics['unmatched reference documents']} documents; see technical traces. Document counts are not additive across groups. Trigger is not used."]
    source_names = [str(p.relative_to(_in_billing_Path(context['input_folder']))) for p in paths]
    population = scope if scope is not None else {r['company'] for r in records.values()} or {'ALL'}
    tables = [(label, summary_headers + summary_tech, _in_billing_summary_rows(summaries, start, cutoff, population, source_names), len(summary_headers), lambda: note), (label + '_Detail', detail_headers + detail_tech, _in_billing_detail_rows(details, traces, chunks), len(detail_headers), lambda: ['See summary for scope. Concatenate Trace parts to JSON; each entry is [BillDoc, Item, BillingDate, CreatedOn, Quantity, NetValue, Can, CancBillDc, PsSt, Origins]. Origins are file|sheet|physical Excel row. Trigger is not used.'])]
    stage = _in_billing_perf_counter()
    output = _in_billing__render(context, cutoff, tables)
    print(f'{label} time | write report: {_in_billing_perf_counter() - stage:.2f} seconds', flush=True)
    for name, value in sorted(metrics.items()):
        print(f'{label} {name}: {value}', flush=True)
    for path in paths:
        print(label + ' input:', path, flush=True)
    print(label + ' output:', output, flush=True)
    print(label + ' Runner OK is not certification of complete inventory-sales coverage.', flush=True)
    print(f'{label} time | total: {_in_billing_perf_counter() - started:.2f} seconds', flush=True)
    return output



# IN10_SOURCE_READERS_V1

def in10_read_sources(context, start, cutoff, companies, metrics):
    """Read PO quantities once per PO item and material movements independently."""
    from collections import defaultdict
    from datetime import date, datetime
    from decimal import Decimal
    import math
    from pathlib import Path
    import re
    from openpyxl import load_workbook
    from openpyxl.utils.datetime import from_excel

    def text(v):
        return '' if v is None else str(v).strip()

    def code(v, width=0):
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            if not math.isfinite(v) or v != int(v):
                raise ValueError(f'Invalid identifier {v!r}')
            v = str(int(v))
        s = text(v)
        return s.zfill(width) if s.isdigit() else s

    def mat(v):
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return code(v)
        return '' if v is None else str(v)

    def numeric(v):
        if v is None or isinstance(v, bool) or not text(v):
            raise ValueError(f'Missing/invalid quantity {v!r}')
        n = Decimal(str(v))
        if not n.is_finite():
            raise ValueError(f'Nonfinite quantity {v!r}')
        return n

    def day(v, epoch):
        if isinstance(v, datetime):
            return v.date()
        if isinstance(v, date):
            return v
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            v = from_excel(v, epoch)
            if isinstance(v, datetime):
                return v.date()
        for f in (
            '%Y-%m-%d', '%d/%m/%Y', '%d.%m.%Y',
            '%Y%m%d', '%Y-%m-%d %H:%M:%S',
        ):
            try:
                return datetime.strptime(text(v), f).date()
            except ValueError:
                pass
        raise ValueError(f'Missing/invalid date {v!r}')

    folder = Path(context['input_folder'])
    files = {}
    for kind, prefix in (
        ('PO', 'LBRSDORDERS'),
        ('MOV', 'LBRSDINVENTORY'),
    ):
        files[kind] = sorted(
            p for p in folder.rglob('*')
            if p.is_file()
            and not p.name.startswith('~$')
            and p.suffix.lower() in {'.xlsx', '.xlsm'}
            and re.fullmatch(
                prefix + cutoff.strftime('%Y%m%d')
                + r'(?:PART\d+|\d+)?',
                re.sub(r'[ _-]+', '', p.stem).upper(),
            )
        )
        if not files[kind]:
            raise FileNotFoundError(
                f'IN10 missing {kind} source for '
                f'{cutoff:%Y%m%d}; folder={folder}'
            )

    result = {'PO': {}, 'MOV': {}}
    signatures = {}
    layouts = {}

    for kind, paths in files.items():
        for path in paths:
            print(f'IN10 reading {kind}: {path}', flush=True)
            book = load_workbook(
                path, read_only=True, data_only=True
            )
            try:
                rows = book['Sheet1'].iter_rows(values_only=True)
                headers = [text(v) for v in next(rows, ())]

                if kind in layouts and layouts[kind] != headers:
                    raise ValueError(
                        f'{path}: inconsistent part layout'
                    )
                layouts[kind] = headers

                c = defaultdict(list)
                for i, h in enumerate(headers):
                    c[h].append(i)

                required = (
                    [
                        'CoCd', 'Material', 'PO Quantity',
                        'Plnt', 'Doc. Date', 'D', 'DCI', 'FIn',
                    ]
                    if kind == 'PO'
                    else [
                        'CoCd', 'Material', 'Plnt', 'Entry Dte',
                        'Pstng Date', 'MvT', 'Quantity', 'D/C',
                    ]
                )
                repeated = (
                    (('Purch.Doc.', 3), ('Item', 3), ('OUn', 2))
                    if kind == 'PO'
                    else (
                        ('Mat. Doc.', 2), ('MatYr', 2),
                        ('Item', 3), ('BUn', 2),
                    )
                )
                wrong = [
                    h for h in required if len(c[h]) != 1
                ]
                wrong += [
                    h for h, n in repeated if len(c[h]) != n
                ]
                if wrong:
                    raise ValueError(
                        f'{path}: missing/ambiguous headers '
                        f'{wrong}; observed={headers}'
                    )

                def get(r, h, occ=0):
                    i = c[h][occ]
                    return r[i] if i < len(r) else None

                for line, r in enumerate(rows, 2):
                    metrics[kind + ' rows read'] += 1
                    origin = (
                        f'{path.relative_to(folder)}'
                        f'|Sheet1|{line}'
                    )
                    try:
                        if not any(text(v) for v in r):
                            metrics[kind + ' blank rows'] += 1
                            continue

                        key_headers = (
                            ['CoCd', 'Purch.Doc.', 'Item']
                            if kind == 'PO'
                            else ['Mat. Doc.', 'MatYr', 'Item']
                        )
                        allowed = (
                            {
                                'PO Quantity', 'OUn', 'Net Price',
                                'Crcy', 'Per', 'Net Value',
                            }
                            if kind == 'PO'
                            else {
                                'CoCd', 'Quantity', 'Amount in LC',
                                'BUn', 'Crcy',
                            }
                        )
                        if (
                            all(
                                not text(get(r, h))
                                for h in key_headers
                            )
                            and all(
                                not text(v)
                                or (
                                    i < len(headers)
                                    and headers[i] in allowed
                                )
                                for i, v in enumerate(r)
                            )
                        ):
                            metrics[kind + ' export totals'] += 1
                            continue

                        co = code(get(r, 'CoCd'), 4)
                        doc = code(
                            get(
                                r,
                                'Purch.Doc.'
                                if kind == 'PO'
                                else 'Mat. Doc.',
                            ),
                            10,
                        )
                        item = code(
                            get(r, 'Item'),
                            5 if kind == 'PO' else 4,
                        )
                        year = (
                            ''
                            if kind == 'PO'
                            else code(get(r, 'MatYr'), 4)
                        )

                        if (
                            not co
                            or not doc.isdigit()
                            or int(doc) <= 0
                            or not item.isdigit()
                            or int(item) <= 0
                            or (
                                kind == 'MOV'
                                and (
                                    not year.isdigit()
                                    or int(year) <= 0
                                )
                            )
                        ):
                            raise ValueError(
                                'Incomplete/invalid source key'
                            )

                        if companies is not None and co not in companies:
                            metrics[kind + ' outside companies'] += 1
                            continue

                        when = day(
                            get(
                                r,
                                'Doc. Date'
                                if kind == 'PO'
                                else 'Entry Dte',
                            ),
                            book.epoch,
                        )
                        posted = (
                            when
                            if kind == 'PO'
                            else day(
                                get(r, 'Pstng Date'), book.epoch
                            )
                        )
                        qty = numeric(
                            get(
                                r,
                                'PO Quantity'
                                if kind == 'PO'
                                else 'Quantity',
                            )
                        )

                        h = 'OUn' if kind == 'PO' else 'BUn'
                        units = {
                            text(get(r, h, i))
                            for i in range(2)
                        } - {''}

                        if len(units) > 1:
                            raise ValueError(
                                f'Contradictory {h} occurrences: '
                                f'{units}'
                            )

                        unit = (
                            text(get(r, h, 1))
                            if kind == 'PO'
                            else next(iter(units), '')
                        )
                        material = mat(get(r, 'Material'))
                        plant = text(get(r, 'Plnt'))
                        key = (
                            (co, doc, item)
                            if kind == 'PO'
                            else (co, year, doc, item)
                        )

                        record = dict(
                            company=co, doc=doc, item=item,
                            year=year, date=when, posting=posted,
                            quantity=qty, unit=unit,
                            material=material, plant=plant,
                            origins=[origin],
                        )

                        if kind == 'PO':
                            record['flags'] = [[
                                text(get(r, h))
                                for h in ('D', 'DCI', 'FIn')
                            ]]
                            signature = (
                                material, plant, qty, unit, when
                            )
                        else:
                            record.update(
                                movement=code(get(r, 'MvT')),
                                dc=text(get(r, 'D/C')),
                            )
                            refdoc = code(
                                get(r, 'Mat. Doc.', 1), 10
                            )
                            if not refdoc.strip('0'):
                                refdoc = ''

                            refyear = code(
                                get(r, 'MatYr', 1), 4
                            )
                            refitem = code(
                                get(r, 'Item', 2), 4
                            )
                            if refdoc and (
                                not refdoc.isdigit()
                                or not refyear.isdigit()
                                or not refitem.isdigit()
                                or int(refyear) <= 0
                                or int(refitem) <= 0
                            ):
                                raise ValueError(
                                    'Incomplete/invalid '
                                    'original-position reference'
                                )

                            record['reference'] = (
                                (co, refyear, refdoc, refitem)
                                if refdoc else None
                            )
                            signature = tuple(r)

                        if key in result[kind]:
                            old = result[kind][key]
                            if signatures[(kind, key)] != signature:
                                raise ValueError(
                                    f'Contradictory {kind} key '
                                    f'{key}; prior={old["origins"]}'
                                )

                            old['origins'].append(origin)
                            if (
                                kind == 'PO'
                                and record['flags'][0]
                                not in old['flags']
                            ):
                                old['flags'] += record['flags']

                            metrics[
                                kind + ' additional rows consolidated'
                            ] += 1
                        else:
                            result[kind][key] = record
                            signatures[(kind, key)] = signature

                    except Exception as error:
                        raise ValueError(
                            f'{origin}: {error}'
                        ) from error
            finally:
                book.close()

    for kind in result:
        metrics[kind + ' distinct keys'] = len(result[kind])

    return result['PO'], result['MOV'], files




# IN09_SOURCE_READERS_V1

def _in09_text(value):
    return '' if value is None else str(value).strip()


def _in09_code(value, width=0, optional=False):
    import math
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if not math.isfinite(value) or int(value) != value:
            raise ValueError(f'Invalid identifier {value!r}')
        value = str(int(value))
    value = _in09_text(value)
    if optional and not value.strip('0'):
        return ''
    return value.zfill(width) if value.isdigit() else value


def _in09_day(value, epoch=None):
    from datetime import date, datetime
    from openpyxl.utils.datetime import from_excel
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if (
        isinstance(value, (int, float))
        and not isinstance(value, bool)
        and epoch is not None
    ):
        value = from_excel(value, epoch)
        if isinstance(value, datetime):
            return value.date()
    for fmt in (
        '%Y-%m-%d', '%d/%m/%Y', '%d.%m.%Y',
        '%Y%m%d', '%Y-%m-%d %H:%M:%S',
    ):
        try:
            return datetime.strptime(
                _in09_text(value), fmt
            ).date()
        except ValueError:
            pass
    raise ValueError(f'Missing/invalid date {value!r}')


def _in09_number(value):
    from decimal import Decimal
    if value is None or isinstance(value, bool) or not _in09_text(value):
        raise ValueError(f'Missing/invalid amount {value!r}')
    value = Decimal(str(value))
    if not value.is_finite():
        raise ValueError(f'Nonfinite amount {value!r}')
    return value


def in09_find_sources(folder, cutoff):
    from pathlib import Path
    import re
    suffix = cutoff.strftime('%Y%m%d')
    patterns = {
        'MASTER': r'LBRSDACCOUNTS' + suffix + r'(?:PART\d+|\d+)?',
        'MAP': r'LBRINVALMAP' + suffix + r'(?:PART\d+|\d+)?',
        'T030': r'LBRINT030' + suffix + r'(?:PART\d+|\d+)?',
        'OPEN': r'LBRSDENTRIESOPEN(?:\d+|' + suffix + r'(?:PART\d+|\d+)?)?',
        'CLEARED': r'LBRSDENTRIESCLEARED(?:\d+|' + suffix + r'(?:PART\d+|\d+)?)?',
        'JOURNAL': r'LBRSDJOURNAL(?:AR)?' + suffix + r'(?:PART\d+|\d+)?',
    }
    found = {kind: [] for kind in patterns}
    for path in sorted(Path(folder).rglob('*')):
        if (
            not path.is_file()
            or path.name.startswith('~$')
            or path.suffix.lower() not in {'.xlsx', '.xlsm'}
        ):
            continue
        name = re.sub(r'[ _-]+', '', path.stem).upper()
        for kind, pattern in patterns.items():
            if re.fullmatch(pattern, name):
                found[kind].append(path)
    missing = [
        kind for kind, paths in found.items()
        if not paths and kind != 'JOURNAL'
    ]
    if missing:
        raise FileNotFoundError(
            f'IN09 missing sources {missing}; '
            f'cutoff={suffix}; folder={folder}'
        )
    return found


def _in09_table(paths, headers, folder, label, metrics):
    from pathlib import Path
    from openpyxl import load_workbook
    for path in paths:
        print(f'IN09 reading {label}: {path}', flush=True)
        book = load_workbook(path, read_only=True, data_only=True)
        try:
            rows = book['Sheet1'].iter_rows(values_only=True)
            observed = [_in09_text(v) for v in next(rows, ())]
            if observed != headers:
                raise ValueError(
                    f'{path}: layout differs; observed={observed}'
                )
            for n, row in enumerate(rows, 2):
                metrics[label + ' rows read'] += 1
                if (n - 1) % 100000 == 0:
                    print(
                        f'IN09 {path.name}: {n - 1:,} rows',
                        flush=True,
                    )
                origin = (
                    f'{path.relative_to(Path(folder))}|Sheet1|{n}'
                )
                if (
                    len(row) > len(headers)
                    and any(
                        _in09_text(v)
                        for v in row[len(headers):]
                    )
                ):
                    raise ValueError(
                        f'{origin}: extra populated columns'
                    )
                values = (
                    list(row) + [None] * len(headers)
                )[:len(headers)]
                if not any(_in09_text(v) for v in values):
                    metrics[label + ' blank rows'] += 1
                    continue
                yield values, origin, book.epoch
        finally:
            book.close()


# IN09_T001_LOCAL_CURRENCY_V1
def in09_inventory_accounts(sources, folder, companies, metrics):
    local_currency = {}
    from collections import defaultdict
    t, c = (_in09_text, _in09_code)
    plans = defaultdict(set)
    names = defaultdict(set)
    masters = defaultdict(set)
    mh = ['CoCd', 'Company Name', 'ChAc', 'Crcy', 'ChAc', 'G/L Acct', 'Date', 'Created by', 'B/S acct', 'Group acct', 'ChAc', 'G/L Acct', 'Language', 'Short Text', 'Long Text']
    for r, origin, epoch in _in09_table(sources['MASTER'], mh, folder, 'MASTER', metrics):
        co = c(r[0], 4)
        if companies is not None and co not in companies:
            continue
        charts = {c(r[i], 4) for i in (2, 4, 10)} - {''}
        accounts = {c(r[i], 10) for i in (5, 11)} - {''}
        if not co or len(charts) != 1 or len(accounts) != 1:
            raise ValueError(f'{origin}: missing/contradictory master company/chart/account')
        chart = next(iter(charts))
        account = next(iter(accounts))
        plans[co].add(chart)
        masters[co, chart].add(account)
        names[co, account].add((t(r[12]), t(r[13]), t(r[14]), origin))
    for co, ch in plans.items():
        if len(ch) != 1:
            raise ValueError(f'IN09 company {co}: multiple supplied charts {ch}')
    areas = {}
    groups = defaultdict(set)
    for r, origin, epoch in _in09_table(sources['MAP'], ['CoCd', 'ValA', 'VGCd', 'Crcy'], folder, 'MAP', metrics):
        co, area, group = (c(r[0], 4), t(r[1]), c(r[2], 4))
        if companies is not None and co not in companies:
            continue
        if not co or not area:
            raise ValueError(f'{origin}: missing company/valuation area')
        if area in areas and areas[area] != (co, group):
            raise ValueError(f'{origin}: conflicting area {area}')
        areas[area] = (co, group)
        groups[co, group].add(area)
        currency = t(r[3])
        if not currency:
            raise ValueError(f'{origin}: T001-WAERS missing for company {co}')
        previous = local_currency.get(co)
        if previous and previous['currency'] != currency:
            raise ValueError(f'{origin}: conflicting T001-WAERS for company {co}; prior={previous}')
        if previous is None:
            local_currency[co] = {'currency': currency, 'origins': []}
        if origin not in local_currency[co]['origins']:
            local_currency[co]['origins'].append(origin)
    config = {}
    origins = defaultdict(list)
    for r, origin, epoch in _in09_table(sources['T030'], ['ChAc', 'Trs', 'VGCd', 'AM', 'ValCl', 'G/L Acct', 'G/L Acct'], folder, 'T030', metrics):
        if t(r[1]) != 'BSX':
            raise ValueError(f'{origin}: unexpected transaction {r[1]!r}')
        key = (c(r[0], 4), c(r[2], 4), t(r[3]), c(r[4], 4))
        accounts = (c(r[5], 10), c(r[6], 10))
        if key in config and config[key] != accounts:
            raise ValueError(f'{origin}: contradictory T030 key {key}; prior={origins[key]}')
        config[key] = accounts
        origins[key].append(origin)
    result = defaultdict(set)
    evidence = defaultdict(list)
    population = companies if companies is not None else set(plans)
    for co in population:
        if len(plans.get(co, set())) != 1:
            raise ValueError(f'IN09 no unique supplied chart for company {co}')
        if co not in {v[0] for v in areas.values()}:
            raise ValueError(f'IN09 no valuation areas supplied for company {co}')
    for (co, group), area_set in sorted(groups.items()):
        if co not in plans:
            raise ValueError(f'IN09 no chart for company {co}; areas={area_set}')
        chart = next(iter(plans[co]))
        matched = [(k, a) for k, a in config.items() if k[0] == chart and k[1] == group]
        if not matched or not any((any(a) for k, a in matched)):
            raise ValueError(f'IN09 no supplied BSX accounts: company={co}, chart={chart}, group={group!r}, areas={sorted(area_set)}')
        for key, accounts in matched:
            if key[2]:
                raise ValueError(f'IN09 nonblank account modifier requires functional definition: {key}; origins={origins[key]}')
            for account in set(accounts) - {''}:
                result[co].add(account)
                evidence[co, account].append([chart, group, key[3], sorted(area_set), origins[key]])
        print(f'IN09 scope: company={co}; group={group!r}; areas={sorted(area_set)}', flush=True)
    display = {}
    for (co, account), options in names.items():
        languages = sorted({v[0] for v in options}, key=lambda v: (0 if v.upper() in {'EN', 'E'} else 1 if v.upper() in {'PT', 'P'} else 2, v))
        selected = {short or long for lang, short, long, origin in options if languages and lang == languages[0]}
        display[co, account] = next(iter(selected)) if len(selected) == 1 else 'MULTIPLE SOURCE DESCRIPTIONS' if selected else 'ACCOUNT NOT SUPPLIED IN MASTER'
    return (dict(result), display, dict(evidence), local_currency)


def in09_gl_rows(
    sources, folder, companies, accounts,
    tcodes, metrics, force_docs=None,
):
    t, c, d, n = (
        _in09_text, _in09_code, _in09_day, _in09_number
    )
    base = [
        'CoCd', 'DocumentNo', 'Year', 'Doc. Type',
        'Doc. Date', 'Pstng Date', 'Entry Dte', 'Time',
        'User name', 'TCode', 'Document Header Text',
        'Crcy', 'Reference', 'CoCd', 'DocumentNo',
        'Year', 'Itm', 'G/L', 'D/C', 'Amount in LC',
        'Crcy', 'Amount', 'Crcy', 'Text',
        'Cost Ctr', 'Profit Ctr',
    ]
    for family in ('OPEN', 'CLEARED'):
        h = (
            base
            + (['Order'] if family == 'CLEARED' else [])
            + [
                'Assignment', 'Clrng doc.', 'Clearing',
                'Rev. with', 'Year', 'Reason', 'Parked by',
            ]
        )
        label = (
            family if force_docs is None
            else family + ' RELATED PASS'
        )
        for r, origin, epoch in _in09_table(
            sources[family], h, folder, label, metrics
        ):
            if (
                all(
                    not t(r[i])
                    for i in (1, 2, 14, 15, 16, 17)
                )
                and all(
                    not t(v) or i in {0, 11, 13, 19, 20, 21, 22}
                    for i, v in enumerate(r)
                )
            ):
                metrics[label + ' export totals'] += 1
                continue
            try:
                co = c(r[0], 4)
                if not co:
                    raise ValueError('Missing company')
                if companies is not None and co not in companies:
                    continue
                if co not in accounts:
                    raise ValueError(
                        f'No account mapping for supplied company {co}'
                    )
                doc, year, item = (
                    c(r[1], 10), c(r[2], 4), c(r[16], 3)
                )
                dk = (co, year, doc)
                account = c(r[17], 10)
                tcode = t(r[9])
                if force_docs is not None:
                    if dk not in force_docs:
                        continue
                elif account not in accounts[co] and tcode not in tcodes:
                    continue

                if any(
                    not v.isdigit() or int(v) <= 0
                    for v in (doc, year, item)
                ):
                    raise ValueError('Incomplete/invalid FI key')
                if (
                    co != c(r[13], 4)
                    or doc != c(r[14], 10)
                    or year != c(r[15], 4)
                ):
                    raise ValueError('Conflicting header/item keys')
                if not account:
                    raise ValueError('Missing G/L account')
                if t(r[11]) and t(r[22]) and t(r[11]) != t(r[22]):
                    raise ValueError(
                        'Conflicting header/item document currency'
                    )

                ri = 30 if family == 'CLEARED' else 29
                header = dict(
                    company=co, year=year, doc=doc,
                    posting=d(r[5], epoch),
                    entry=d(r[6], epoch),
                    docdate=d(r[4], epoch),
                    user=t(r[8]), tcode=tcode,
                    doctype=t(r[3]), dc_currency=t(r[11]),
                    reference=t(r[12]), text=t(r[10]),
                    revdoc=c(r[ri], 10, True),
                    revyear=c(r[ri + 1], 4, True),
                    origin=origin,
                )
                line = dict(
                    company=co, year=year, doc=doc, item=item,
                    account=account, dc=t(r[18]),
                    amount=n(r[21]), lc_amount=n(r[19]),
                    currency=t(r[22]), lc_currency=t(r[20]),
                    text=t(r[23]), account_type='', pk='',
                    vendor='', customer='', origins=[origin],
                    source_types={'GL'},
                )
                yield dk, header, line
            except Exception as error:
                raise ValueError(
                    f'{origin}: {error}'
                ) from error


def in09_journal_rows(paths, folder, docs, companies, metrics):
    t, c, n = _in09_text, _in09_code, _in09_number
    headers = [
        'CoCd', 'DocumentNo', 'Year', 'Itm',
        'AccTy', 'PK', 'G/L', 'Vendor', 'Customer',
        'D/C', 'Amount', 'Crcy', 'Amount in LC',
        'Crcy', 'Text',
    ]
    for r, origin, epoch in _in09_table(
        paths, headers, folder, 'JOURNAL', metrics
    ):
        if (
            all(not t(r[i]) for i in (0, 1, 2, 3, 6))
            and all(
                not t(v) or i in {10, 11, 12, 13}
                for i, v in enumerate(r)
            )
        ):
            metrics['JOURNAL export totals'] += 1
            continue
        try:
            co, doc, year = c(r[0], 4), c(r[1], 10), c(r[2], 4)
            dk = (co, year, doc)
            if companies is not None and co not in companies:
                continue
            if dk not in docs:
                continue
            item, account = c(r[3], 3), c(r[6], 10)
            if not account or not item.isdigit() or int(item) <= 0:
                raise ValueError('Missing account/invalid FI item')
            yield dict(
                company=co, year=year, doc=doc, item=item,
                account=account, dc=t(r[9]),
                amount=n(r[10]), lc_amount=n(r[12]),
                currency=t(r[11]), lc_currency=t(r[13]),
                text=t(r[14]), account_type=t(r[4]), pk=t(r[5]),
                vendor=c(r[7], 10, True),
                customer=c(r[8], 10, True),
                origins=[origin], source_types={'BSEG'},
            )
        except Exception as error:
            raise ValueError(
                f'{origin}: {error}'
            ) from error






# IN_COMMON_HELPERS_UNIFIED_V1
# Project-stock source support
"""Shared source support for project-stock valuation."""
from collections import defaultdict as _in_project_internal_defaultdict
import json as _in_project_internal_json
from pathlib import Path as _in_project_internal_Path
import re as _in_project_internal_re
from openpyxl import load_workbook as _in_project_internal_load_workbook

def _in_project_internal_files(folder, prefix, cutoff):
    result = sorted((p for p in _in_project_internal_Path(folder).rglob('*') if p.is_file() and (not p.name.startswith('~$')) and (p.suffix.lower() in {'.xlsx', '.xlsm'}) and _in_project_internal_re.fullmatch(prefix + cutoff.strftime('%Y%m%d') + '(?:PART\\d+|\\d+)?', _in_project_internal_re.sub('[ _-]+', '', p.stem).upper())))
    if not result:
        raise FileNotFoundError(f'Missing source {prefix} for {cutoff:%Y%m%d}')
    return result

def _in_project_internal_rows(paths, expected, folder):
    for path in paths:
        print('Project stock reading:', path.name, flush=True)
        book = _in_project_internal_load_workbook(path, read_only=True, data_only=True)
        try:
            iterator = book['Sheet1'].iter_rows(values_only=True)
            headers = ['' if v is None else str(v).strip() for v in next(iterator, ())]
            if headers != expected:
                raise ValueError(f'{path}: unexpected headers: {headers}')
            for line, values in enumerate(iterator, 2):
                row = (list(values) + [None] * len(headers))[:len(headers)]
                if not any((v is not None and str(v).strip() for v in row)):
                    continue
                origin = f'{path.relative_to(_in_project_internal_Path(folder))}|Sheet1|{line}'
                yield (row, origin)
        finally:
            book.close()

def _in_project_internal_extend_sources(existing, folder, cutoff):
    result = dict(existing)
    result['QBEW'] = _in_project_internal_files(folder, 'LBRINQBEW', cutoff)
    result['QBEWH'] = _in_project_internal_files(folder, 'LBRINQBEWH', cutoff)
    return result

def _in_project_internal_prepare_metadata(connection, companies, metrics):
    from core import in_common as c
    if connection.execute('SELECT 1 FROM meta WHERE kind=? AND k=?', ('project-metadata-ready', 'ALL')).fetchone():
        return
    folder = _in_project_internal_Path(metrics['_input'])
    suffix = metrics.get('_project_cutoff')
    if not suffix:
        raise ValueError('Project metadata: cutoff was not initialized.')
    from datetime import datetime
    cutoff = datetime.strptime(suffix, '%Y%m%d').date()
    for row, origin in _in_project_internal_rows(_in_project_internal_files(folder, 'LBRINMATERIALUNITS', cutoff), ['Material', 'BUn'], folder):
        material, unit = (c.material(row[0]), c.txt(row[1]))
        if not material.strip():
            raise ValueError(f'{origin}: missing material')
        c.meta_add(connection, 'unit', material, unit, origin)
        metrics['MARA unit metadata rows'] += 1
    for row, origin in _in_project_internal_rows(_in_project_internal_files(folder, 'LBRINVALMAP', cutoff), ['CoCd', 'ValA', 'VGCd', 'Crcy'], folder):
        company = c.txt(row[0])
        if company.isdigit():
            company = company.zfill(4)
        if companies is not None and company not in companies:
            continue
        area, currency = (c.txt(row[1]), c.txt(row[3]))
        if not company or not area or (not currency):
            raise ValueError(f'{origin}: incomplete company/area/currency')
        c.mapping_add(connection, 'area-company', area, company, origin)
        existing = connection.execute('SELECT DISTINCT v FROM meta WHERE kind=? AND k=?', ('currency', company)).fetchall()
        if existing and any((v[0] != currency for v in existing)):
            raise ValueError(f'{origin}: conflicting company currency')
        c.meta_add(connection, 'currency', company, currency, origin)
    c.meta_add(connection, 'project-metadata-ready', 'ALL', '1', 'loader')
    connection.commit()

def _in_project_internal_load_project(connection, source, paths, companies, metrics):
    from core import in_common as c
    _in_project_internal_prepare_metadata(connection, companies, metrics)
    historical = source == 'QBEWH'
    expected = ['Material', 'ValA', 'Val. Type', 'S', 'WBS Element', 'Year', 'Pe', 'Total Stock', 'BUn', 'Total Value', 'Crcy', 'Pr.', 'MovAvgPrice', 'Crcy', 'Standard price', 'Crcy', 'per', 'ValCl']
    for row, origin in _in_project_internal_rows(paths, expected, metrics['_input']):
        metrics[source + ' rows read'] += 1
        allowed_totals = {7, 8, 9, 10, 12, 13, 14, 15, 16}
        if all((not c.txt(row[i]) for i in range(7))) and all((not c.txt(v) or i in allowed_totals for i, v in enumerate(row))):
            metrics[source + ' export totals'] += 1
            continue
        try:
            material = c.material(row[0])
            area, valuation, special, wbs = (c.txt(row[1]), c.txt(row[2]), c.txt(row[3]), c.txt(row[4]))
            if not material.strip() or not area or (not wbs) or (special != 'Q'):
                raise ValueError('Invalid material/area/WBS/special-stock key')
            mapping = connection.execute('SELECT v FROM mapping WHERE k=?', ('area-company|' + area,)).fetchone()
            if not mapping:
                if companies is not None:
                    metrics[source + ' outside mapped company scope'] += 1
                    continue
                raise ValueError(f'No company mapping for area {area}')
            company = mapping[0]
            if companies is not None and company not in companies:
                continue
            period = c.period(row[5], row[6])
            amounts = {'Total Stock': c.num(row[7]), 'Total Value': c.num(row[9]), 'MovAvgPrice': c.num(row[12]), 'Standard price': c.num(row[14]), 'per': c.num(row[16])}
            currencies = {c.txt(row[i]) for i in (10, 13, 15)} - {''}
            mapped_currency, _ = c.meta_get(connection, 'currency', company)
            if not mapped_currency:
                raise ValueError('Missing/conflicting company currency')
            if currencies and currencies != {mapped_currency}:
                raise ValueError('Source currency conflicts with company currency')
            data = dict(company=company, material=material, area=area, plant='', sloc='', vt=valuation, period=period, unit=c.txt(row[8]), currency=next(iter(currencies), ''), fv='', numbers=amounts, price_control=c.txt(row[11]), valuation_class=c.txt(row[17]), flag='', wbs=wbs, special_stock=special)
            key = _in_project_internal_json.dumps([company, material, area, valuation, special, wbs], ensure_ascii=False)
            payload = _in_project_internal_json.dumps(data, ensure_ascii=False, sort_keys=True)
            stored_period = period if historical else 0
            previous = connection.execute('SELECT data,origins FROM snapshot WHERE source=? AND k=? AND p=?', (source, key, stored_period)).fetchone()
            if previous:
                if previous[0] != payload:
                    raise ValueError(f'Contradictory project-stock key {key}')
                origins = _in_project_internal_json.loads(previous[1]) + [origin]
                connection.execute('UPDATE snapshot SET origins=? WHERE source=? AND k=? AND p=?', (_in_project_internal_json.dumps(origins, ensure_ascii=False), source, key, stored_period))
                metrics[source + ' identical duplicates'] += 1
            else:
                connection.execute('INSERT INTO snapshot VALUES(?,?,?,?,?,?)', ('P', source, key, stored_period, payload, _in_project_internal_json.dumps([origin], ensure_ascii=False)))
                metrics[source + ' distinct records'] += 1
        except Exception as error:
            raise ValueError(f'{origin}: {error}') from error
    connection.commit()

# Report footnotes redirected to the runner log
"""Send explanatory IN report footnotes to the runner console/log."""
import functools as _in_notes_internal_functools
import inspect as _in_notes_internal_inspect
_in_notes_internal_ENABLED = {'IN_001', 'IN_002', 'IN_003', 'IN_005', 'IN_007', 'IN_008', 'IN_009', 'IN_010'}

def _in_notes_internal_emit(label, callback):
    if callback is None:
        return
    values = callback() if callable(callback) else callback
    if values is None:
        return
    if isinstance(values, str):
        values = [values]
    for value in values:
        if value is not None and str(value).strip():
            print(f'{label} | REPORT NOTE | {value}', flush=True)

def _in_notes_internal_install(namespace):
    writer = namespace.get('write_sheet')
    renderer = namespace.get('_in_billing__render')
    if callable(writer) and (not getattr(writer, '_notes_to_log', False)):
        writer_signature = _in_notes_internal_inspect.signature(writer)
        parameters = list(writer_signature.parameters.values())
        if not parameters:
            raise ValueError('write_sheet signature not supported')
        footer_name = parameters[-1].name
        if not any((word in footer_name.lower() for word in ('note', 'footer'))):
            raise ValueError(f'Unsupported write_sheet footer parameter: {footer_name}')

        @_in_notes_internal_functools.wraps(writer)
        def write_without_notes(*args, **kwargs):
            frame = _in_notes_internal_inspect.currentframe()
            try:
                control = frame.f_back.f_globals.get('CONTROL_ID', '')
            finally:
                del frame
            if control not in _in_notes_internal_ENABLED:
                return writer(*args, **kwargs)
            bound = writer_signature.bind(*args, **kwargs)
            bound.apply_defaults()
            _in_notes_internal_emit(control, bound.arguments.get(footer_name))
            bound.arguments[footer_name] = lambda: []
            return writer(*bound.args, **bound.kwargs)
        write_without_notes._notes_to_log = True
        namespace['write_sheet'] = write_without_notes
    if callable(renderer) and (not getattr(renderer, '_notes_to_log', False)):
        renderer_signature = _in_notes_internal_inspect.signature(renderer)
        parameters = list(renderer_signature.parameters.values())
        if len(parameters) < 3:
            raise ValueError('Renderer signature not supported')
        specs_name = parameters[2].name

        @_in_notes_internal_functools.wraps(renderer)
        def render_without_notes(*args, **kwargs):
            bound = renderer_signature.bind(*args, **kwargs)
            bound.apply_defaults()
            cleaned = []
            for spec in bound.arguments[specs_name]:
                if len(spec) != 5:
                    raise ValueError('Unsupported IN report sheet specification')
                name, headers, rows, visible, footer = spec
                base = str(name).split('_', 1)[0]
                control = 'IN_' + base[2:].zfill(3) if base.startswith('IN') else ''
                if control in _in_notes_internal_ENABLED:
                    _in_notes_internal_emit(name, footer)
                    footer = lambda: []
                cleaned.append((name, headers, rows, visible, footer))
            bound.arguments[specs_name] = cleaned
            return renderer(*bound.args, **bound.kwargs)
        render_without_notes._notes_to_log = True
        namespace['_in_billing__render'] = render_without_notes

# IN_PROJECT_SOURCE_SUPPORT_V1
_in_project_original_find_sources = find_sources
_in_project_original_load_source = load_source


def find_sources(folder, cutoff):
    extend_sources = _in_project_internal_extend_sources
    return extend_sources(
        _in_project_original_find_sources(folder, cutoff),
        folder, cutoff,
    )


def load_source(connection, source, paths, company_filter, metrics):
    prepare_metadata = _in_project_internal_prepare_metadata
    load_project = _in_project_internal_load_project
    # IN01 sets this before loading sources.
    # Other controls retain their existing loading until separately upgraded.
    if metrics.get("_project_cutoff"):
        prepare_metadata(connection, company_filter, metrics)
    if source in ("QBEW", "QBEWH"):
        return load_project(
            connection, source, paths, company_filter, metrics
        )
    return _in_project_original_load_source(
        connection, source, paths, company_filter, metrics
    )


# IN_REPORT_NOTES_TO_LOG_V1
_install_in_notes_log = _in_notes_internal_install
_install_in_notes_log(globals())


# IN_REPORT_NOTES_POLICY_READY_V1
# Report footnotes go to the runner log; body messages remain.
_in_notes_internal_install(globals())
