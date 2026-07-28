"""GL_009 - Potential duplicate journals: same GL account and same USD amount.

The control follows the LHA GL09 rule at accounting-line level. It groups signed
USD amounts rounded to two decimals by company and GL account, and reports all
lines only when the group contains more than one distinct journal.
"""

import posixpath
import shutil
import tempfile
import zipfile
from pathlib import Path
from time import perf_counter
from xml.etree import ElementTree

import pandas as pd
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from core.gl_common import (
    build_company_name_map,
    build_gl_account_name_map,
    filter_by_company,
    get_gl_output_file,
    get_optional_column,
    load_gl_bsas_data,
    load_gl_bsis_data,
    load_gl_fx_rates_data,
    load_gl_master_data,
    normalize_company_output,
    normalize_fx_rates,
    open_or_create_gl_output_workbook,
    recreate_gl_sheet,
    require_columns,
    save_gl_output_workbook,
    select_fx_rate_to_usd,
    to_datetime_value,
    write_dataframe_to_sheet,
    write_single_sheet_workbook_fast,
)


SHEET_NAME = "GL09"
REPORT_CURRENCY = "USD"


REQUIRED_COLUMNS = {
    "company_code": [
        "Empr",
        "Empr.1",
        "BUKRS",
        "BKPF-BUKRS",
        "BSIS-BUKRS",
        "BSAS-BUKRS",
    ],
    "fiscal_year": [
        "Ano",
        "GJAHR",
        "BKPF-GJAHR",
        "BSIS-GJAHR",
        "BSAS-GJAHR",
    ],
    "document_number": [
        "Nº doc.",
        "Nº doc..1",
        "Nº doc",
        "BELNR",
        "BKPF-BELNR",
        "BSIS-BELNR",
        "BSAS-BELNR",
    ],
    "line_number": [
        "Itm",
        "BUZEI",
        "BSIS-BUZEI",
        "BSAS-BUZEI",
    ],
    "gl_account": [
        "Razão",
        "Razao",
        "Cta.Razão",
        "Cta.Razão.1",
        "HKONT",
        "BSIS-HKONT",
        "BSAS-HKONT",
    ],
    "debit_credit": [
        "D/C",
        "SHKZG",
        "BSIS-SHKZG",
        "BSAS-SHKZG",
    ],
    "amount_local": [
        "Montante em MI",
        "DMBTR",
        "BSIS-DMBTR",
        "BSAS-DMBTR",
    ],
    "amount_document": [
        "Montante",
        "WRBTR",
        "BSIS-WRBTR",
        "BSAS-WRBTR",
    ],
    "document_currency": [
        "Moeda",
        "Moeda.1",
        "WAERS",
        "BKPF-WAERS",
    ],
    "document_date": [
        "Data doc.",
        "Data doc",
        "BLDAT",
        "BKPF-BLDAT",
    ],
    "entry_date": [
        "Dt.entr.",
        "Dt.entr",
        "CPUDT",
        "BKPF-CPUDT",
    ],
    "posting_date": [
        "Dt.lçto.",
        "Dt.lçto",
        "Dt.lcto.",
        "Dt.lcto",
        "BUDAT",
        "BKPF-BUDAT",
    ],
}


OPTIONAL_COLUMNS = {
    "document_type": [
        "Tp.doc.",
        "Tp.doc",
        "BLART",
        "BKPF-BLART",
    ],
    "document_text": [
        "Texto cabeçalho documento",
        "Texto cabecalho documento",
        "BKTXT",
        "BKPF-BKTXT",
    ],
    "line_text": [
        "Texto",
        "SGTXT",
        "BSIS-SGTXT",
        "BSAS-SGTXT",
    ],
    "create_user": [
        "Pré-edição",
        "Pre-edição",
        "Pré-edicao",
        "Pre-edicao",
        "PPNAM",
        "BKPF-PPNAM",
    ],
    "approver_user": [
        "Nome do usuário",
        "Nome do usuario",
        "USNAM",
        "BKPF-USNAM",
    ],
    "transaction_code": [
        "CódT",
        "CodT",
        "TCODE",
        "BKPF-TCODE",
    ],
    "clearing_document": [
        "DocCompens",
        "AUGBL",
        "BSIS-AUGBL",
        "BSAS-AUGBL",
    ],
    "clearing_date": [
        "Compensaç.",
        "Compensac.",
        "Compensação",
        "Compensacao",
        "AUGDT",
        "BSIS-AUGDT",
        "BSAS-AUGDT",
    ],
}


OUTPUT_COLUMNS = [
    "Company Code",
    "Company Name",
    "GL Account",
    "GL Account Description",
    "Document Type",
    "Document Number",
    "Document Text",
    "Line Number",
    "Debit",
    "Credit",
    "Amount in Reporting Currency",
    "Report Currency",
    "Amount in Document Currency",
    "Document Currency",
    "Document Date",
    "Date Entered",
    "Posting Date",
    "Create User ID",
    "Create User Name",
    "Approver User ID",
    "Approver User Name",
    "Fiscal Period",
    "Transaction Code",
    "Transaction Code Description",
    "Fiscal Year",
    "Amount in Reporting Currency Rounded",
    "Duplicate Journal Count",
    "Duplicate Line Count",
    "Duplicate Key",
    "FX Method",
    "FX Rate",
    "FX Rate Date",
    "Source",
]


DATE_COLUMNS = [
    "Document Date",
    "Date Entered",
    "Posting Date",
    "FX Rate Date",
]


AMOUNT_COLUMNS = [
    "Debit",
    "Credit",
    "Amount in Reporting Currency",
    "Amount in Document Currency",
    "Amount in Reporting Currency Rounded",
    "FX Rate",
]


INTEGER_COLUMNS = [
    "Duplicate Journal Count",
    "Duplicate Line Count",
]


SPREADSHEET_NAMESPACE = (
    "http://schemas.openxmlformats.org/"
    "spreadsheetml/2006/main"
)

DOCUMENT_RELATIONSHIP_NAMESPACE = (
    "http://schemas.openxmlformats.org/"
    "officeDocument/2006/relationships"
)

PACKAGE_RELATIONSHIP_NAMESPACE = (
    "http://schemas.openxmlformats.org/"
    "package/2006/relationships"
)

EMPTY_WORKSHEET_XML = (
    '<?xml version="1.0" '
    'encoding="UTF-8" '
    'standalone="yes"?>'
    f'<worksheet xmlns="{SPREADSHEET_NAMESPACE}">'
    "<sheetData/>"
    "</worksheet>"
).encode("utf-8")


GL09_COLUMN_WIDTHS = {
    "Company Code": 14,
    "Company Name": 28,
    "GL Account": 18,
    "GL Account Description": 34,
    "Document Type": 15,
    "Document Number": 20,
    "Document Text": 36,
    "Line Number": 14,
    "Debit": 18,
    "Credit": 18,
    "Amount in Reporting Currency": 28,
    "Report Currency": 16,
    "Amount in Document Currency": 28,
    "Document Currency": 18,
    "Document Date": 15,
    "Date Entered": 15,
    "Posting Date": 15,
    "Create User ID": 18,
    "Create User Name": 24,
    "Approver User ID": 18,
    "Approver User Name": 24,
    "Fiscal Period": 15,
    "Transaction Code": 18,
    "Transaction Code Description": 32,
    "Fiscal Year": 14,
    "Amount in Reporting Currency Rounded": 34,
    "Duplicate Journal Count": 24,
    "Duplicate Line Count": 21,
    "Duplicate Key": 48,
    "FX Method": 34,
    "FX Rate": 18,
    "FX Rate Date": 15,
    "Source": 14,
}


def _blank(index):
    return pd.Series(
        "",
        index=index,
        dtype="object",
    )


def _optional(
    dataframe,
    resolved_columns,
    logical_name,
):
    column_name = resolved_columns.get(
        logical_name
    )

    if column_name is None:
        return _blank(
            dataframe.index
        )

    return dataframe[
        column_name
    ]


def _clean_text(series):
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .replace(
            {
                "nan": "",
                "None": "",
                "<NA>": "",
            }
        )
    )


def _normalize_code(series):
    return (
        _clean_text(series)
        .str.replace(
            r"\.0$",
            "",
            regex=True,
        )
    )


def _normalize_company(series):
    return _normalize_code(
        series
    ).map(
        normalize_company_output
    )


def _parse_number(series):
    text = _clean_text(
        series
    )

    negative_parentheses = (
        text.str.match(
            r"^\(.*\)$"
        )
    )

    text = (
        text.str.replace(
            r"[()]",
            "",
            regex=True,
        )
        .str.replace(
            " ",
            "",
            regex=False,
        )
    )

    both_separators = (
        text.str.contains(
            ",",
            regex=False,
        )
        & text.str.contains(
            ".",
            regex=False,
        )
    )

    comma_decimal = (
        text.str.contains(
            r",\d{1,6}$",
            regex=True,
        )
    )

    text = text.where(
        ~both_separators,
        (
            text.str.replace(
                ".",
                "",
                regex=False,
            )
            .str.replace(
                ",",
                ".",
                regex=False,
            )
        ),
    )

    text = text.where(
        (
            both_separators
            | ~comma_decimal
        ),
        (
            text.str.replace(
                ".",
                "",
                regex=False,
            )
            .str.replace(
                ",",
                ".",
                regex=False,
            )
        ),
    )

    text = text.where(
        (
            both_separators
            | comma_decimal
        ),
        text.str.replace(
            ",",
            "",
            regex=False,
        ),
    )

    values = pd.to_numeric(
        text,
        errors="coerce",
    )

    return values.where(
        ~negative_parentheses,
        -values.abs(),
    )


def _parse_date(series):
    parsed_dates = pd.to_datetime(
        series,
        errors="coerce",
        dayfirst=True,
    )

    numeric_dates = pd.to_numeric(
        series,
        errors="coerce",
    )

    excel_dates = pd.to_datetime(
        numeric_dates,
        unit="D",
        origin="1899-12-30",
        errors="coerce",
    )

    return parsed_dates.where(
        parsed_dates.notna(),
        excel_dates,
    )


def _resolve_optional_columns(
    dataframe,
):
    return {
        logical_name: get_optional_column(
            dataframe,
            possible_names,
        )
        for logical_name, possible_names
        in OPTIONAL_COLUMNS.items()
    }


def _apply_gl09_fast_formatting(
    worksheet,
    dataframe,
):
    """
    Apply GL09 formatting without scanning every cell to calculate widths.
    """
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = (
        worksheet.dimensions
    )

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="D9EAF7",
    )

    header_font = Font(
        bold=True
    )

    for cell in worksheet[1]:
        cell.font = header_font
        cell.fill = header_fill

    column_positions = {
        column_name: column_index
        for column_index, column_name
        in enumerate(
            dataframe.columns,
            start=1,
        )
    }

    for (
        column_name,
        column_width,
    ) in GL09_COLUMN_WIDTHS.items():
        column_index = (
            column_positions.get(
                column_name
            )
        )

        if column_index is None:
            continue

        column_letter = (
            get_column_letter(
                column_index
            )
        )

        worksheet.column_dimensions[
            column_letter
        ].width = column_width

    format_by_column = {
        **{
            column_name: "dd/mm/yyyy"
            for column_name
            in DATE_COLUMNS
        },
        **{
            column_name: (
                "#,##0.00;[Red]-#,##0.00"
            )
            for column_name
            in AMOUNT_COLUMNS
        },
        **{
            column_name: "0"
            for column_name
            in INTEGER_COLUMNS
        },
    }

    maximum_row = (
        worksheet.max_row
    )

    for (
        column_name,
        number_format,
    ) in format_by_column.items():
        column_index = (
            column_positions.get(
                column_name
            )
        )

        if column_index is None:
            continue

        for row_index in range(
            2,
            maximum_row + 1,
        ):
            worksheet.cell(
                row=row_index,
                column=column_index,
            ).number_format = (
                number_format
            )


def _find_sheet_xml_path(
    workbook_archive,
    sheet_name,
):
    workbook_xml = (
        ElementTree.fromstring(
            workbook_archive.read(
                "xl/workbook.xml"
            )
        )
    )

    relationship_id = None

    for sheet in workbook_xml.findall(
        f".//{{{SPREADSHEET_NAMESPACE}}}sheet"
    ):
        current_sheet_name = str(
            sheet.attrib.get(
                "name",
                "",
            )
        ).strip()

        if (
            current_sheet_name.casefold()
            != sheet_name.casefold()
        ):
            continue

        relationship_id = (
            sheet.attrib.get(
                (
                    "{"
                    f"{DOCUMENT_RELATIONSHIP_NAMESPACE}"
                    "}id"
                )
            )
        )
        break

    if relationship_id is None:
        return None

    relationships_xml = (
        ElementTree.fromstring(
            workbook_archive.read(
                "xl/_rels/workbook.xml.rels"
            )
        )
    )

    for relationship in (
        relationships_xml.findall(
            (
                "{"
                f"{PACKAGE_RELATIONSHIP_NAMESPACE}"
                "}Relationship"
            )
        )
    ):
        if (
            relationship.attrib.get(
                "Id"
            )
            != relationship_id
        ):
            continue

        target = relationship.attrib.get(
            "Target",
            "",
        )

        if target.startswith("/"):
            return target.lstrip("/")

        return posixpath.normpath(
            posixpath.join(
                "xl",
                target,
            )
        )

    return None


def _create_lightweight_local_workbook(
    source_file,
    local_file,
):
    """
    Copy the XLSX locally while replacing the previous GL09 XML with an
    empty worksheet.

    GL09 is replaced immediately after open. Avoiding the parse of the
    previous 100k+ row sheet substantially reduces openpyxl load time.
    Every other workbook member and worksheet remains unchanged.
    """
    source_file = Path(
        source_file
    )

    local_file = Path(
        local_file
    )

    with zipfile.ZipFile(
        source_file,
        "r",
    ) as source_archive:
        gl09_xml_path = (
            _find_sheet_xml_path(
                source_archive,
                SHEET_NAME,
            )
        )

        with zipfile.ZipFile(
            local_file,
            "w",
            compression=zipfile.ZIP_DEFLATED,
            compresslevel=1,
            allowZip64=True,
        ) as local_archive:
            for member in (
                source_archive.infolist()
            ):
                if member.is_dir():
                    local_archive.writestr(
                        member,
                        b"",
                    )
                    continue

                if (
                    member.filename
                    == gl09_xml_path
                ):
                    local_archive.writestr(
                        member,
                        EMPTY_WORKSHEET_XML,
                    )
                    continue

                with source_archive.open(
                    member,
                    "r",
                ) as source_member:
                    with local_archive.open(
                        member,
                        "w",
                    ) as local_member:
                        shutil.copyfileobj(
                            source_member,
                            local_member,
                            length=1024 * 1024,
                        )


def _company_currency_map(
    master_dataframe,
):
    if master_dataframe.empty:
        return {}

    company_column = (
        get_optional_column(
            master_dataframe,
            [
                "BUKRS",
                "Empr",
                "Company Code",
            ],
        )
    )

    currency_column = (
        get_optional_column(
            master_dataframe,
            [
                "WAERS",
                "Moeda",
                "Currency",
                "Company Main Currency",
            ],
        )
    )

    if (
        company_column is None
        or currency_column is None
    ):
        return {}

    currency_dataframe = pd.DataFrame(
        {
            "company": _normalize_company(
                master_dataframe[
                    company_column
                ]
            ),
            "currency": _clean_text(
                master_dataframe[
                    currency_column
                ]
            ).str.upper(),
        }
    )

    currency_dataframe = (
        currency_dataframe[
            (
                currency_dataframe[
                    "company"
                ] != ""
            )
            & (
                currency_dataframe[
                    "currency"
                ] != ""
            )
        ]
        .drop_duplicates(
            subset=["company"],
            keep="first",
        )
    )

    return (
        currency_dataframe
        .set_index(
            "company"
        )["currency"]
        .to_dict()
    )


def _prepare_source(
    source_dataframe,
    context,
    source_name,
):
    if source_dataframe.empty:
        return pd.DataFrame()

    required_columns = require_columns(
        dataframe=source_dataframe,
        required_columns=REQUIRED_COLUMNS,
        source_name=(
            f"GL09 {source_name}"
        ),
    )

    optional_columns = (
        _resolve_optional_columns(
            source_dataframe
        )
    )

    module_config = (
        context["module"]
    )

    filtered_dataframe = (
        filter_by_company(
            dataframe=source_dataframe,
            company_column=(
                required_columns[
                    "company_code"
                ]
            ),
            companies_filter=(
                module_config.get(
                    "companies",
                    "",
                )
            ),
        )
    )

    posting_date = _parse_date(
        filtered_dataframe[
            required_columns[
                "posting_date"
            ]
        ]
    )

    from_date = to_datetime_value(
        module_config.get(
            "from",
            "",
        )
    )

    to_date = to_datetime_value(
        module_config.get(
            "to",
            "",
        )
    )

    posting_date_filter = (
        posting_date.between(
            from_date,
            to_date,
            inclusive="both",
        )
    )

    filtered_dataframe = (
        filtered_dataframe.loc[
            posting_date_filter
        ].copy()
    )

    posting_date = posting_date.loc[
        filtered_dataframe.index
    ]

    debit_credit_indicator = (
        _clean_text(
            filtered_dataframe[
                required_columns[
                    "debit_credit"
                ]
            ]
        ).str.upper()
    )

    amount_local_absolute = (
        _parse_number(
            filtered_dataframe[
                required_columns[
                    "amount_local"
                ]
            ]
        ).abs()
    )

    amount_document_absolute = (
        _parse_number(
            filtered_dataframe[
                required_columns[
                    "amount_document"
                ]
            ]
        ).abs()
    )

    amount_local_signed = (
        amount_local_absolute.where(
            (
                debit_credit_indicator
                != "H"
            ),
            -amount_local_absolute,
        )
    )

    amount_document_signed = (
        amount_document_absolute.where(
            (
                debit_credit_indicator
                != "H"
            ),
            -amount_document_absolute,
        )
    )

    result = pd.DataFrame(
        index=filtered_dataframe.index
    )

    result["Company Code"] = (
        _normalize_company(
            filtered_dataframe[
                required_columns[
                    "company_code"
                ]
            ]
        )
    )

    result["Fiscal Year"] = (
        _normalize_code(
            filtered_dataframe[
                required_columns[
                    "fiscal_year"
                ]
            ]
        )
    )

    result["Document Number"] = (
        _normalize_code(
            filtered_dataframe[
                required_columns[
                    "document_number"
                ]
            ]
        )
    )

    result["Line Number"] = (
        _normalize_code(
            filtered_dataframe[
                required_columns[
                    "line_number"
                ]
            ]
        )
    )

    result["GL Account"] = (
        _normalize_code(
            filtered_dataframe[
                required_columns[
                    "gl_account"
                ]
            ]
        )
    )

    result["Document Type"] = (
        _clean_text(
            _optional(
                filtered_dataframe,
                optional_columns,
                "document_type",
            )
        )
    )

    result["Document Text"] = (
        _clean_text(
            _optional(
                filtered_dataframe,
                optional_columns,
                "document_text",
            )
        )
    )

    result["Line Text"] = (
        _clean_text(
            _optional(
                filtered_dataframe,
                optional_columns,
                "line_text",
            )
        )
    )

    result["Debit"] = (
        amount_local_absolute.where(
            (
                debit_credit_indicator
                == "S"
            ),
            pd.NA,
        )
    )

    result["Credit"] = (
        amount_local_absolute.where(
            (
                debit_credit_indicator
                == "H"
            ),
            pd.NA,
        )
    )

    result[
        "_LOCAL_SIGNED"
    ] = amount_local_signed

    result[
        "Amount in Document Currency"
    ] = amount_document_signed

    result["Document Currency"] = (
        _clean_text(
            filtered_dataframe[
                required_columns[
                    "document_currency"
                ]
            ]
        ).str.upper()
    )

    result["Document Date"] = (
        _parse_date(
            filtered_dataframe[
                required_columns[
                    "document_date"
                ]
            ]
        )
    )

    result["Date Entered"] = (
        _parse_date(
            filtered_dataframe[
                required_columns[
                    "entry_date"
                ]
            ]
        )
    )

    result["Posting Date"] = (
        posting_date
    )

    result["Create User ID"] = (
        _clean_text(
            _optional(
                filtered_dataframe,
                optional_columns,
                "create_user",
            )
        )
    )

    result["Approver User ID"] = (
        _clean_text(
            _optional(
                filtered_dataframe,
                optional_columns,
                "approver_user",
            )
        )
    )

    result[
        "Transaction Code"
    ] = _clean_text(
        _optional(
            filtered_dataframe,
            optional_columns,
            "transaction_code",
        )
    )

    result[
        "_CLEARING_DOCUMENT"
    ] = _normalize_code(
        _optional(
            filtered_dataframe,
            optional_columns,
            "clearing_document",
        )
    )

    result[
        "_CLEARING_DATE"
    ] = _parse_date(
        _optional(
            filtered_dataframe,
            optional_columns,
            "clearing_date",
        )
    )

    result["Source"] = (
        source_name
    )

    print(
        f"GL09 {source_name} "
        "rows read: "
        f"{len(source_dataframe)}"
    )

    print(
        f"GL09 {source_name} "
        "rows after CONFIG filters: "
        f"{len(result)}"
    )

    return result.reset_index(
        drop=True
    )


def _deduplicate_sources(
    bsis_dataframe,
    bsas_dataframe,
):
    line_key = [
        "Company Code",
        "Fiscal Year",
        "Document Number",
        "Line Number",
    ]

    combined_dataframe = pd.concat(
        [
            bsis_dataframe,
            bsas_dataframe,
        ],
        ignore_index=True,
    )

    if combined_dataframe.empty:
        return (
            combined_dataframe,
            0,
        )

    combined_dataframe[
        "_SOURCE_PRIORITY"
    ] = (
        combined_dataframe[
            "Source"
        ]
        .map(
            {
                "BSIS": 0,
                "BSAS": 1,
            }
        )
        .fillna(0)
    )

    combined_dataframe[
        "_ORIGINAL_ORDER"
    ] = range(
        len(combined_dataframe)
    )

    overlap_mask = (
        combined_dataframe
        .duplicated(
            subset=line_key,
            keep=False,
        )
    )

    overlapping_sources = (
        combined_dataframe.loc[
            overlap_mask
        ]
        .groupby(
            line_key,
            dropna=False,
        )["Source"]
        .agg(
            lambda values: "/".join(
                sorted(
                    set(values)
                )
            )
        )
    )

    combined_dataframe = (
        combined_dataframe
        .sort_values(
            by=[
                "_SOURCE_PRIORITY",
                "_ORIGINAL_ORDER",
            ],
            kind="stable",
        )
        .drop_duplicates(
            subset=line_key,
            keep="last",
        )
    )

    if not overlapping_sources.empty:
        combined_dataframe = (
            combined_dataframe.merge(
                overlapping_sources.rename(
                    "_MERGED_SOURCE"
                ),
                on=line_key,
                how="left",
            )
        )

        combined_dataframe[
            "Source"
        ] = (
            combined_dataframe[
                "_MERGED_SOURCE"
            ]
            .fillna(
                combined_dataframe[
                    "Source"
                ]
            )
        )

        combined_dataframe = (
            combined_dataframe.drop(
                columns=[
                    "_MERGED_SOURCE"
                ]
            )
        )

    removed_rows = (
        len(bsis_dataframe)
        + len(bsas_dataframe)
        - len(combined_dataframe)
    )

    combined_dataframe = (
        combined_dataframe.drop(
            columns=[
                "_SOURCE_PRIORITY",
                "_ORIGINAL_ORDER",
            ]
        )
    )

    return (
        combined_dataframe,
        removed_rows,
    )


def _add_usd_amount(
    dataframe,
    master_dataframe,
    fx_dataframe,
):
    result = dataframe.copy()

    result[
        "Amount in Reporting Currency"
    ] = pd.NA

    result["FX Method"] = ""
    result["FX Rate"] = pd.NA
    result["FX Rate Date"] = pd.NaT

    if result.empty:
        return result

    company_currencies = (
        result["Company Code"]
        .map(
            _company_currency_map(
                master_dataframe
            )
        )
        .fillna("")
        .str.upper()
    )

    document_currencies = (
        result["Document Currency"]
        .fillna("")
        .str.upper()
    )

    company_usd = (
        company_currencies.isin(
            [
                "USD",
                "$",
            ]
        )
    )

    document_usd = (
        ~company_usd
        & document_currencies.isin(
            [
                "USD",
                "$",
            ]
        )
    )

    direct_usd = (
        company_usd
        | document_usd
    )

    result.loc[
        company_usd,
        "Amount in Reporting Currency",
    ] = result.loc[
        company_usd,
        "_LOCAL_SIGNED",
    ]

    result.loc[
        company_usd,
        "FX Method",
    ] = (
        "Local amount "
        "(company currency USD)"
    )

    result.loc[
        document_usd,
        "Amount in Reporting Currency",
    ] = result.loc[
        document_usd,
        "Amount in Document Currency",
    ]

    result.loc[
        document_usd,
        "FX Method",
    ] = (
        "Document amount "
        "(document currency USD)"
    )

    result.loc[
        direct_usd,
        "FX Rate",
    ] = 1.0

    result.loc[
        direct_usd,
        "FX Rate Date",
    ] = result.loc[
        direct_usd,
        "Document Date",
    ]

    pending_conversion = (
        ~direct_usd
    )

    if (
        pending_conversion.any()
        and not fx_dataframe.empty
    ):
        normalized_fx_dataframe = (
            normalize_fx_rates(
                fx_dataframe
            )
        )

        conversion_currencies = (
            company_currencies.where(
                (
                    company_currencies
                    != ""
                ),
                document_currencies,
            )
        )

        unique_rate_keys = (
            pd.DataFrame(
                {
                    "currency": (
                        conversion_currencies.loc[
                            pending_conversion
                        ]
                    ),
                    "date": result.loc[
                        pending_conversion,
                        "Document Date",
                    ],
                }
            )
            .drop_duplicates()
        )

        rate_rows = []

        for rate_key in (
            unique_rate_keys.itertuples(
                index=False
            )
        ):
            fx_details = (
                select_fx_rate_to_usd(
                    normalized_fx_dataframe=(
                        normalized_fx_dataframe
                    ),
                    currency=(
                        rate_key.currency
                    ),
                    requested_date=(
                        rate_key.date
                    ),
                )
            )

            rate_rows.append(
                {
                    "currency": (
                        rate_key.currency
                    ),
                    "date": (
                        rate_key.date
                    ),
                    "factor": (
                        pd.NA
                        if fx_details is None
                        else fx_details[
                            "fx_to_usd"
                        ]
                    ),
                    "method": (
                        ""
                        if fx_details is None
                        else fx_details[
                            "method"
                        ]
                    ),
                    "rate": (
                        pd.NA
                        if fx_details is None
                        else fx_details[
                            "usd_rate"
                        ]
                    ),
                    "rate_date": (
                        pd.NaT
                        if fx_details is None
                        else fx_details[
                            "rate_date"
                        ]
                    ),
                }
            )

        rate_dataframe = pd.DataFrame(
            rate_rows
        )

        pending_dataframe = (
            pd.DataFrame(
                {
                    "_index": (
                        result.index[
                            pending_conversion
                        ]
                    ),
                    "currency": (
                        conversion_currencies.loc[
                            pending_conversion
                        ]
                    ),
                    "date": result.loc[
                        pending_conversion,
                        "Document Date",
                    ],
                }
            )
        )

        pending_dataframe = (
            pending_dataframe
            .merge(
                rate_dataframe,
                on=[
                    "currency",
                    "date",
                ],
                how="left",
            )
            .set_index(
                "_index"
            )
        )

        fx_to_usd = pd.to_numeric(
            pending_dataframe[
                "factor"
            ],
            errors="coerce",
        )

        result.loc[
            pending_dataframe.index,
            "Amount in Reporting Currency",
        ] = (
            result.loc[
                pending_dataframe.index,
                "_LOCAL_SIGNED",
            ]
            * fx_to_usd
        )

        result.loc[
            pending_dataframe.index,
            "FX Method",
        ] = pending_dataframe[
            "method"
        ]

        result.loc[
            pending_dataframe.index,
            "FX Rate",
        ] = pending_dataframe[
            "rate"
        ]

        result.loc[
            pending_dataframe.index,
            "FX Rate Date",
        ] = pending_dataframe[
            "rate_date"
        ]

    result[
        "Amount in Reporting Currency"
    ] = pd.to_numeric(
        result[
            "Amount in Reporting Currency"
        ],
        errors="coerce",
    )

    return result


def create_gl09_duplicate_journals(
    bsis_dataframe,
    bsas_dataframe,
    master_dataframe,
    fx_dataframe,
    context,
):
    stage_started = (
        perf_counter()
    )

    bsis_prepared = _prepare_source(
        bsis_dataframe,
        context,
        "BSIS",
    )

    bsas_prepared = _prepare_source(
        bsas_dataframe,
        context,
        "BSAS",
    )

    print(
        "GL09 preparation seconds: "
        f"{perf_counter() - stage_started:.2f}"
    )

    stage_started = (
        perf_counter()
    )

    (
        combined_dataframe,
        removed_rows,
    ) = _deduplicate_sources(
        bsis_prepared,
        bsas_prepared,
    )

    print(
        "GL09 combined rows: "
        f"{len(bsis_prepared) + len(bsas_prepared)}"
    )

    print(
        "GL09 technical duplicates removed: "
        f"{removed_rows}"
    )

    print(
        "GL09 technical deduplication seconds: "
        f"{perf_counter() - stage_started:.2f}"
    )

    invalid_company_rows = (
        combined_dataframe[
            "Company Code"
        ].eq("").sum()
        if not combined_dataframe.empty
        else 0
    )

    invalid_account_rows = (
        combined_dataframe[
            "GL Account"
        ].eq("").sum()
        if not combined_dataframe.empty
        else 0
    )

    basic_valid_mask = (
        (
            combined_dataframe[
                "Company Code"
            ] != ""
        )
        & (
            combined_dataframe[
                "Fiscal Year"
            ] != ""
        )
        & (
            combined_dataframe[
                "Document Number"
            ] != ""
        )
        & (
            combined_dataframe[
                "Line Number"
            ] != ""
        )
        & (
            combined_dataframe[
                "GL Account"
            ] != ""
        )
    )

    stage_started = (
        perf_counter()
    )

    working_dataframe = (
        _add_usd_amount(
            combined_dataframe.loc[
                basic_valid_mask
            ].copy(),
            master_dataframe,
            fx_dataframe,
        )
    )

    print(
        "GL09 reporting currency seconds: "
        f"{perf_counter() - stage_started:.2f}"
    )

    working_dataframe[
        "Amount in Reporting Currency Rounded"
    ] = working_dataframe[
        "Amount in Reporting Currency"
    ].round(2)

    working_dataframe[
        "Amount in Reporting Currency Rounded"
    ] = working_dataframe[
        "Amount in Reporting Currency Rounded"
    ].mask(
        working_dataframe[
            "Amount in Reporting Currency Rounded"
        ].eq(0),
        0.0,
    )

    reporting_amount_available = (
        working_dataframe[
            "Amount in Reporting Currency Rounded"
        ].notna()
    )

    group_key = [
        "Company Code",
        "GL Account",
        "Amount in Reporting Currency Rounded",
    ]

    journal_key = [
        "Fiscal Year",
        "Document Number",
    ]

    candidate_dataframe = (
        working_dataframe.loc[
            reporting_amount_available
        ].copy()
    )

    stage_started = (
        perf_counter()
    )

    journal_index = (
        pd.MultiIndex.from_frame(
            candidate_dataframe[
                journal_key
            ]
        )
    )

    candidate_dataframe[
        "_JOURNAL_ID"
    ] = pd.factorize(
        journal_index,
        sort=False,
    )[0]

    group_statistics = (
        candidate_dataframe
        .groupby(
            group_key,
            sort=False,
            observed=True,
            dropna=False,
        )
        .agg(
            **{
                "Duplicate Journal Count": (
                    "_JOURNAL_ID",
                    "nunique",
                ),
                "Duplicate Line Count": (
                    "_JOURNAL_ID",
                    "size",
                ),
            }
        )
        .reset_index()
    )

    duplicate_groups = (
        group_statistics[
            (
                group_statistics[
                    "Duplicate Journal Count"
                ] > 1
            )
        ].copy()
    )

    if duplicate_groups.empty:
        output_dataframe = (
            pd.DataFrame(
                columns=OUTPUT_COLUMNS
            )
        )
    else:
        output_dataframe = (
            candidate_dataframe.merge(
                duplicate_groups,
                on=group_key,
                how="inner",
            )
        )

        company_name_map = (
            build_company_name_map(
                master_dataframe
            )
        )

        account_name_map = (
            build_gl_account_name_map(
                master_dataframe
            )
        )

        output_dataframe[
            "Company Name"
        ] = (
            output_dataframe[
                "Company Code"
            ]
            .map(
                company_name_map
            )
            .fillna("")
        )

        output_dataframe[
            "GL Account Description"
        ] = (
            output_dataframe[
                "GL Account"
            ]
            .map(
                account_name_map
            )
            .fillna("")
        )

        output_dataframe[
            "Report Currency"
        ] = REPORT_CURRENCY

        output_dataframe[
            "Create User Name"
        ] = ""

        output_dataframe[
            "Approver User Name"
        ] = ""

        output_dataframe[
            "Fiscal Period"
        ] = (
            output_dataframe[
                "Posting Date"
            ]
            .dt.strftime(
                "%Y-%m"
            )
            .fillna("")
        )

        output_dataframe[
            "Transaction Code Description"
        ] = ""

        rounded_amount_key = (
            output_dataframe[
                "Amount in Reporting Currency Rounded"
            ].map(
                lambda value: (
                    f"{value:.2f}"
                )
            )
        )

        output_dataframe[
            "Duplicate Key"
        ] = (
            output_dataframe[
                "Company Code"
            ]
            + " | "
            + output_dataframe[
                "GL Account"
            ]
            + " | "
            + rounded_amount_key
        )

        output_dataframe = (
            output_dataframe
            .sort_values(
                by=(
                    group_key
                    + [
                        "Fiscal Year",
                        "Document Number",
                        "Line Number",
                    ]
                ),
                kind="stable",
            )
            .reset_index(
                drop=True
            )
        )

        output_dataframe = (
            output_dataframe[
                OUTPUT_COLUMNS
            ]
        )

    print(
        "GL09 duplicate detection seconds: "
        f"{perf_counter() - stage_started:.2f}"
    )

    print(
        "GL09 rows without Company Code: "
        f"{invalid_company_rows}"
    )

    print(
        "GL09 rows without GL Account: "
        f"{invalid_account_rows}"
    )

    print(
        "GL09 rows with reporting amount: "
        f"{reporting_amount_available.sum()}"
    )

    print(
        "GL09 rows without reporting amount: "
        f"{(~reporting_amount_available).sum()}"
    )

    rounded_zero_rows = (
        working_dataframe[
            "Amount in Reporting Currency Rounded"
        ].eq(0).sum()
    )

    print(
        "GL09 rows with rounded zero amount: "
        f"{rounded_zero_rows}"
    )

    print(
        "GL09 duplicate keys: "
        f"{len(duplicate_groups)}"
    )

    print(
        "GL09 output rows: "
        f"{len(output_dataframe)}"
    )

    return output_dataframe


def write_gl09_output(
    output_dataframe,
    context,
):
    write_started = (
        perf_counter()
    )

    output_file = get_gl_output_file(
        context
    )

    if not output_file.exists():
        fast_written = (
            write_single_sheet_workbook_fast(
                output_file=output_file,
                sheet_name=SHEET_NAME,
                dataframe=output_dataframe,
                date_columns=DATE_COLUMNS,
                amount_columns=AMOUNT_COLUMNS,
                integer_columns=INTEGER_COLUMNS,
            )
        )

        if fast_written:
            print(
                f"GL09 output workbook: "
                f"{output_file}"
            )

            print(
                "GL09 sheet written: GL09"
            )

            print(
                "GL09 workbook write seconds: "
                f"{perf_counter() - write_started:.2f}"
            )

            return output_file

    with tempfile.TemporaryDirectory(
        prefix="lbr_gl09_"
    ) as temporary_folder:
        local_output_file = (
            Path(temporary_folder)
            / output_file.name
        )

        stage_started = (
            perf_counter()
        )

        _create_lightweight_local_workbook(
            output_file,
            local_output_file,
        )

        print(
            "GL09 local staging seconds: "
            f"{perf_counter() - stage_started:.2f}"
        )

        stage_started = (
            perf_counter()
        )

        workbook = (
            open_or_create_gl_output_workbook(
                local_output_file
            )
        )

        print(
            "GL09 workbook open seconds: "
            f"{perf_counter() - stage_started:.2f}"
        )

        worksheet = recreate_gl_sheet(
            workbook,
            SHEET_NAME,
        )

        stage_started = (
            perf_counter()
        )

        write_dataframe_to_sheet(
            worksheet=worksheet,
            dataframe=output_dataframe,
        )

        print(
            "GL09 worksheet data write seconds: "
            f"{perf_counter() - stage_started:.2f}"
        )

        stage_started = (
            perf_counter()
        )

        _apply_gl09_fast_formatting(
            worksheet,
            output_dataframe,
        )

        print(
            "GL09 worksheet formatting seconds: "
            f"{perf_counter() - stage_started:.2f}"
        )

        stage_started = (
            perf_counter()
        )

        save_gl_output_workbook(
            workbook,
            local_output_file,
        )

        print(
            "GL09 local workbook save seconds: "
            f"{perf_counter() - stage_started:.2f}"
        )

        workbook.close()

        stage_started = (
            perf_counter()
        )

        try:
            shutil.copy2(
                local_output_file,
                output_file,
            )
        except PermissionError as error:
            raise PermissionError(
                "Could not replace GL output workbook: "
                f"{output_file}. "
                "Close the workbook and run again."
            ) from error

        print(
            "GL09 final workbook copy seconds: "
            f"{perf_counter() - stage_started:.2f}"
        )

    print(
        f"GL09 output workbook: "
        f"{output_file}"
    )

    print(
        "GL09 sheet replaced: GL09"
    )

    print(
        "GL09 workbook write seconds: "
        f"{perf_counter() - write_started:.2f}"
    )

    return output_file


def run_gl_009(context):
    """
    Execute GL_009 and write or replace only the GL09 worksheet.
    """
    started_at = (
        perf_counter()
    )

    print("=" * 80)

    print(
        "Running GL_009 - Potential Duplicate General Journals: "
        "Same GL Account & Same Amount"
    )

    print(
        "LHA logic: Company + GL Account + signed USD amount "
        "rounded to 2 decimals; distinct journals > 1"
    )

    bsis_dataframe = (
        load_gl_bsis_data(
            context
        )
    )

    bsas_dataframe = (
        load_gl_bsas_data(
            context
        )
    )

    master_dataframe = (
        load_gl_master_data(
            context
        )
    )

    fx_dataframe = (
        load_gl_fx_rates_data(
            context
        )
    )

    if (
        bsis_dataframe.empty
        and bsas_dataframe.empty
    ):
        raise FileNotFoundError(
            "GL09 requires BSIS and/or "
            "BSAS journal input."
        )

    output_dataframe = (
        create_gl09_duplicate_journals(
            bsis_dataframe=bsis_dataframe,
            bsas_dataframe=bsas_dataframe,
            master_dataframe=master_dataframe,
            fx_dataframe=fx_dataframe,
            context=context,
        )
    )

    write_gl09_output(
        output_dataframe=output_dataframe,
        context=context,
    )

    companies_reported = (
        output_dataframe[
            "Company Code"
        ].nunique()
        if not output_dataframe.empty
        else 0
    )

    accounts_reported = (
        output_dataframe[
            "GL Account"
        ].nunique()
        if not output_dataframe.empty
        else 0
    )

    print(
        "GL09 companies reported: "
        f"{companies_reported}"
    )

    print(
        "GL09 accounts reported: "
        f"{accounts_reported}"
    )

    print(
        "GL09 elapsed seconds: "
        f"{perf_counter() - started_at:.2f}"
    )

    return output_dataframe
