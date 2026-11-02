"""
Common utilities for CD01-CD04.

Responsibilities:
- Locate and read the shared LBR CA_YYYYMMDD input using the CD module TO date.
- Normalize Portuguese or English LBR CA headers to one canonical schema.
- Apply common company, intercompany, parsing, output, and formatting helpers.

Design rules:
- LBR CA is a module-level input shared by CD01-CD04.
- Every CD control must remain independently executable.
- Each control writes or replaces only its own output sheet.
- Controls must preserve sheets produced by other CD controls.
- The shared output workbook is:
    output/LBR_Results_CD_YYYYMMDD.xlsx

Shared responsibilities for CD01-CD04:
- Locate the period-based LBR CA input.
- Read and clean Excel, CSV, or TXT extracts.
- Normalize Portuguese or English LBR CA headers.
- Apply company and intercompany filters.
- Provide shared output workbook and formatting helpers.

The common loader returns canonical Portuguese headers so existing CD controls
remain independent and do not need language-specific input logic.
"""

import warnings
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


warnings.filterwarnings(
    "ignore",
    message="Workbook contains no default style, apply openpyxl's default",
    category=UserWarning,
)


ALLOWED_EXTENSIONS = [
    ".xlsx",
    ".xls",
    ".csv",
    ".txt",
]

CD_DEFAULT_INPUT_KEYWORD = "LBR CA"
CD_HEADER_FILL = "D9EAF7"
CD_DATE_FORMAT = "dd/mm/yyyy"
CD_AMOUNT_FORMAT = "#,##0.00"
CD_INTEGER_FORMAT = "#,##0"


CD_BASE_COLUMN_ALIASES = {
    "Empr": [
        "Empr",
        "CoCd",
    ],
    "Nº doc.": [
        "Nº doc.",
        "Nº doc",
        "DocumentNo",
    ],
    "Fornecedor": [
        "Fornecedor",
        "Vendor",
    ],
    "Compensaç.": [
        "Compensaç.",
        "Compensaç",
        "Compensac.",
        "Compensac",
        "Clearing",
    ],
    "DocCompens": [
        "DocCompens",
        "Clrng doc.",
    ],
    "D/C": [
        "D/C",
    ],
    "Montante": [
        "Montante",
        "Amount",
    ],
    "Nome da firma": [
        "Nome da firma",
        "Company Name",
    ],
    "Nome 1": [
        "Nome 1",
        "Name 1",
    ],
    "Data doc.": [
        "Data doc.",
        "Data doc",
        "Doc. Date",
    ],
    "Data doc..1": [
        "Data doc..1",
        "Data doc.1",
        "Data doc1",
        "Doc. Date.1",
    ],
    "Moeda": [
        "Moeda",
        "Crcy",
    ],
    "Moeda.4": [
        "Moeda.4",
        "Crcy.4",
    ],
    "Referência": [
        "Referência",
        "Referencia",
        "Reference",
    ],
    "Referência.1": [
        "Referência.1",
        "Referencia.1",
        "Reference.1",
    ],
    "Chv.ref.": [
        "Chv.ref.",
        "Chv.ref",
        "Ref. Key",
    ],
    "CódT": [
        "CódT",
        "CodT",
        "TCode",
    ],
    "MP": [
        "MP",
        "PM",
    ],
    "Nº doc..1": [
        "Nº doc..1",
        "Nº doc.1",
        "DocumentNo.1",
    ],
    "Estorno c/": [
        "Estorno c/",
        "Estorno c",
        "Rev. with",
    ],
    "Motiv.est": [
        "Motiv.est",
        "Motiv est",
        "Reason",
    ],
    "Estorno": [
        "Estorno",
        "Reversal",
    ],
}


CD_BASE_REQUIRED_CANONICAL_COLUMNS = [
    "Empr",
    "Nº doc.",
    "Fornecedor",
    "Compensaç.",
    "DocCompens",
    "D/C",
    "Montante",
]


def has_allowed_extension(file_path):
    """
    Return True if the file extension is supported.
    """
    file_path = Path(file_path)

    return (
        file_path.suffix.lower()
        in ALLOWED_EXTENSIONS
    )


def normalize_text(value):
    """
    Normalize text values and pandas-like blanks.
    """
    if value is None:
        return ""

    value_text = str(value).strip()

    if value_text.lower() in [
        "nan",
        "none",
        "nat",
        "<na>",
    ]:
        return ""

    return value_text


def normalize_header(value):
    """
    Normalize a header for case-insensitive matching.
    """
    return normalize_text(value).lower()


def normalize_company_output(value):
    """
    Normalize company code for output display.

    Examples:
    - 0030 -> 30
    - 0034 -> 34
    - 0052 -> 52
    - 0093 -> 93
    """
    value_text = normalize_text(value)

    if value_text == "":
        return ""

    if value_text.endswith(".0"):
        value_text = value_text[:-2]

    if value_text.isdigit():
        return str(int(value_text))

    return value_text


def parse_companies_filter(value):
    """
    Parse module COMPANIES filter into normalized values.

    Accepted examples:
    - ALL
    - 0030,0034,0052
    - 0030;0034;0052
    - 0030|0034|0052
    - 0030 0034 0052
    - 30 34 52
    """
    value_text = normalize_text(value)

    if value_text == "":
        return []

    if value_text.upper() in [
        "ALL",
        "TODAS",
        "TODOS",
    ]:
        return []

    for separator in [
        ";",
        "|",
        "\n",
        "\r",
        "\t",
    ]:
        value_text = value_text.replace(
            separator,
            ",",
        )

    if (
        "," not in value_text
        and " " in value_text
    ):
        value_text = ",".join(
            value_text.split()
        )

    companies = []

    for item in value_text.split(","):
        item = normalize_text(item)

        if item == "":
            continue

        if item.upper() in [
            "ALL",
            "TODAS",
            "TODOS",
        ]:
            return []

        companies.append(
            normalize_company_output(item)
        )

    return companies


def to_datetime_value(value):
    """
    Convert Excel, ISO, English, or local dates to pandas datetime.
    """
    if pd.isna(value):
        return pd.NaT

    if isinstance(value, pd.Timestamp):
        return value

    value_text = normalize_text(value)

    if value_text == "":
        return pd.NaT

    if (
        len(value_text) >= 10
        and value_text[4:5] == "-"
        and value_text[7:8] == "-"
    ):
        return pd.to_datetime(
            value_text,
            errors="coerce",
            dayfirst=False,
        )

    return pd.to_datetime(
        value_text,
        errors="coerce",
        dayfirst=True,
    )


def parse_number(value):
    """
    Parse SAP/Excel numbers using US or Brazilian separators.
    """
    if pd.isna(value):
        return pd.NA

    if isinstance(value, (int, float)):
        return value

    value_text = normalize_text(value)

    if value_text == "":
        return pd.NA

    value_text = value_text.replace(
        " ",
        "",
    )

    is_negative = (
        value_text.startswith("(")
        and value_text.endswith(")")
    )

    if is_negative:
        value_text = value_text[1:-1]

    if (
        "," in value_text
        and "." in value_text
    ):
        last_comma = value_text.rfind(",")
        last_dot = value_text.rfind(".")

        if last_comma > last_dot:
            value_text = value_text.replace(
                ".",
                "",
            )
            value_text = value_text.replace(
                ",",
                ".",
            )

        else:
            value_text = value_text.replace(
                ",",
                "",
            )

    elif (
        "," in value_text
        and "." not in value_text
    ):
        value_text = value_text.replace(
            ",",
            ".",
        )

    try:
        parsed_number = float(value_text)

    except ValueError:
        return pd.NA

    if is_negative:
        parsed_number = parsed_number * -1

    return parsed_number


def get_period_suffix(module_config):
    """
    Return YYYYMMDD using the CD module TO date.
    """
    to_date = module_config.get(
        "to",
        "",
    )

    parsed_date = to_datetime_value(
        to_date
    )

    if pd.isna(parsed_date):
        raise ValueError(
            "Could not determine CD output period "
            "because module TO date is empty or invalid."
        )

    return parsed_date.strftime("%Y%m%d")


def get_cd_output_file(context):
    """
    Return output/LBR_Results_CD_YYYYMMDD.xlsx.
    """
    output_folder = Path(
        context["output_folder"]
    )

    module_config = context["module"]

    period_suffix = get_period_suffix(
        module_config
    )

    return (
        output_folder
        / f"LBR_Results_CD_{period_suffix}.xlsx"
    )


def find_all_files_containing(
    base_folder,
    text_to_find,
):
    """
    Find all supported files whose filename contains the requested text.
    """
    base_folder = Path(base_folder)

    text_to_find = normalize_text(
        text_to_find
    ).lower()

    matched_files = []

    for file_path in base_folder.rglob("*"):
        if not file_path.is_file():
            continue

        if file_path.name.startswith("~$"):
            continue

        if not has_allowed_extension(
            file_path
        ):
            continue

        if (
            text_to_find
            in file_path.name.lower()
        ):
            matched_files.append(file_path)

    return sorted(matched_files)


def get_cd_base_input_keyword(context):
    """
    Return LBR CA_YYYYMMDD using the CD module TO date.
    """
    period_suffix = get_period_suffix(
        context["module"]
    )

    return (
        f"{CD_DEFAULT_INPUT_KEYWORD}_"
        f"{period_suffix}"
    )


def find_cd_base_input_file(context):
    """
    Find exactly one LBR CA input for the configured CD period.
    """
    input_folder = Path(
        context["input_folder"]
    )

    keyword = get_cd_base_input_keyword(
        context
    )

    matched_files = find_all_files_containing(
        base_folder=input_folder,
        text_to_find=keyword,
    )

    if len(matched_files) == 0:
        raise FileNotFoundError(
            "CD input file not found. "
            "Expected one file in input/ containing "
            f"'{keyword}' in the filename. "
            f"Input folder: {input_folder}"
        )

    if len(matched_files) > 1:
        matched_files_text = "\n".join(
            f"- {file_path}"
            for file_path in matched_files
        )

        raise ValueError(
            "Multiple CD input files found using "
            f"keyword '{keyword}'. Expected only "
            "one LBR CA extract for the CD module "
            f"TO date.\n{matched_files_text}"
        )

    return matched_files[0]


def read_input_file(
    file_path,
    sheet_name=0,
):
    """
    Read an xlsx, xls, csv, or tab-separated txt input.
    """
    file_path = Path(file_path)

    if not file_path.exists():
        raise FileNotFoundError(
            f"Input file not found: {file_path}"
        )

    extension = file_path.suffix.lower()

    if extension in [
        ".xlsx",
        ".xls",
    ]:
        return pd.read_excel(
            file_path,
            sheet_name=sheet_name,
        )

    if extension == ".csv":
        return pd.read_csv(file_path)

    if extension == ".txt":
        return pd.read_csv(
            file_path,
            sep="\t",
        )

    raise ValueError(
        "Unsupported input file extension: "
        f"{file_path.suffix}"
    )


def clean_dataframe(dataframe):
    """
    Remove fully empty rows/columns and trim headers.
    """
    result = dataframe.copy()

    result = result.dropna(
        axis=0,
        how="all",
    )

    result = result.dropna(
        axis=1,
        how="all",
    )

    result.columns = [
        normalize_text(column)
        for column in result.columns
    ]

    return result


def normalize_cd_base_columns(dataframe):
    """
    Normalize Portuguese or English LBR CA headers to the canonical names
    already used by CD01-CD04.

    Pandas suffixes on duplicated headers are significant. For example,
    DocumentNo and DocumentNo.1 represent different source fields.
    """
    result = dataframe.copy()
    rename_map = {}

    for canonical_name, aliases in (
        CD_BASE_COLUMN_ALIASES.items()
    ):
        normalized_aliases = {
            normalize_header(alias)
            for alias in aliases
        }

        matching_columns = [
            column
            for column in result.columns
            if normalize_header(column)
            in normalized_aliases
        ]

        if len(matching_columns) > 1:
            raise ValueError(
                "LBR CA contains ambiguous columns "
                f"for '{canonical_name}': "
                f"{matching_columns}. Keep only one "
                "accepted alias for this field."
            )

        if len(matching_columns) == 1:
            source_column = matching_columns[0]

            if source_column != canonical_name:
                rename_map[source_column] = (
                    canonical_name
                )

    normalized_dataframe = (
        result.rename(
            columns=rename_map,
        ).copy()
    )

    if normalized_dataframe.empty:
        raise ValueError(
            "LBR CA contains no data rows."
        )

    missing_columns = [
        column
        for column
        in CD_BASE_REQUIRED_CANONICAL_COLUMNS
        if column
        not in normalized_dataframe.columns
    ]

    if missing_columns:
        raise ValueError(
            "LBR CA is missing required columns "
            "after header normalization. "
            f"Missing: {missing_columns}. "
            f"Available columns: "
            f"{list(result.columns)}"
        )

    return normalized_dataframe


def find_column(
    dataframe,
    possible_names,
):
    """
    Find a DataFrame column using exact case-insensitive matching.
    """
    normalized_lookup = {
        normalize_header(column): column
        for column in dataframe.columns
    }

    for possible_name in possible_names:
        normalized_name = normalize_header(
            possible_name
        )

        if normalized_name in normalized_lookup:
            return normalized_lookup[
                normalized_name
            ]

    return None


def require_columns(
    dataframe,
    required_columns,
):
    """
    Resolve and validate a dictionary of required logical columns.
    """
    resolved_columns = {}
    missing_columns = []

    for (
        logical_name,
        possible_names,
    ) in required_columns.items():
        column_name = find_column(
            dataframe,
            possible_names,
        )

        if column_name is None:
            missing_columns.append(
                f"{logical_name}: expected "
                f"one of {possible_names}"
            )

        else:
            resolved_columns[
                logical_name
            ] = column_name

    if missing_columns:
        raise ValueError(
            "Missing required CD input columns:\n- "
            + "\n- ".join(missing_columns)
        )

    return resolved_columns


def get_optional_column(
    dataframe,
    possible_names,
):
    """
    Return a matching column or None.
    """
    return find_column(
        dataframe,
        possible_names,
    )


def get_series_or_blank(
    dataframe,
    column_name,
):
    """
    Return a DataFrame column or a blank value.
    """
    if column_name is None:
        return ""

    return dataframe[column_name]


def filter_by_company(
    dataframe,
    company_column,
    companies_filter,
):
    """
    Apply the configured company filter.
    """
    result = dataframe.copy()

    companies = parse_companies_filter(
        companies_filter
    )

    if not companies:
        return result

    normalized_company = (
        result[company_column].apply(
            normalize_company_output
        )
    )

    return result[
        normalized_company.isin(companies)
    ].copy()


def apply_signed_amount(
    amount_value,
    debit_credit_indicator,
):
    """
    Apply the CD sign convention: H positive and S negative.
    """
    amount = parse_number(amount_value)

    indicator = normalize_text(
        debit_credit_indicator
    ).upper()

    if pd.isna(amount):
        return pd.NA

    amount = abs(float(amount))

    if indicator == "S":
        return amount * -1

    return amount


def describe_debit_credit_indicator(value):
    """
    Return a readable D/C description.
    """
    indicator = normalize_text(
        value
    ).upper()

    if indicator == "H":
        return "H - Payment / Positive"

    if indicator == "S":
        return "S - Reversal / Negative"

    return "Other / Review"


def build_key(*values):
    """
    Build a pipe-separated key.
    """
    cleaned_values = [
        normalize_text(value)
        for value in values
    ]

    return "|".join(cleaned_values)


def load_cd_base_data(context):
    """
    Load and normalize the common CD input used by CD01-CD04.
    """
    input_file = find_cd_base_input_file(
        context
    )

    print(
        f"CD base input file: {input_file}"
    )

    dataframe = read_input_file(
        input_file
    )

    dataframe = clean_dataframe(
        dataframe
    )

    dataframe = normalize_cd_base_columns(
        dataframe
    )

    return dataframe


def normalize_intercompany_match_value(value):
    """
    Normalize vendor/customer codes and names for intercompany matching.
    """
    value_text = normalize_text(
        value
    ).upper()

    if value_text == "":
        return ""

    if value_text.endswith(".0"):
        value_text = value_text[:-2]

    if value_text.isdigit():
        return str(int(value_text))

    return " ".join(
        value_text.split()
    )


def get_cd_intercompany_reference_sets():
    """
    Return customer-code and company-name intercompany reference sets.
    """
    from core.intercompanies import INTERCOMPANIES

    customer_values = set()
    excluded_name_values = set()

    for intercompany in INTERCOMPANIES:
        customer_value = (
            normalize_intercompany_match_value(
                intercompany.get(
                    "customer",
                    "",
                )
            )
        )

        excluded_name_value = (
            normalize_intercompany_match_value(
                intercompany.get(
                    "excluded_intercompany",
                    "",
                )
            )
        )

        if customer_value != "":
            customer_values.add(
                customer_value
            )

        if excluded_name_value != "":
            excluded_name_values.add(
                excluded_name_value
            )

    return {
        "customers": customer_values,
        "excluded_names": excluded_name_values,
    }


def is_cd_intercompany_vendor(
    vendor_code,
    vendor_name,
    reference_sets=None,
):
    """
    Return True when a CD vendor is identified as intercompany.
    """
    if reference_sets is None:
        reference_sets = (
            get_cd_intercompany_reference_sets()
        )

    normalized_vendor_code = (
        normalize_intercompany_match_value(
            vendor_code
        )
    )

    normalized_vendor_name = (
        normalize_intercompany_match_value(
            vendor_name
        )
    )

    if (
        normalized_vendor_code
        in reference_sets["customers"]
    ):
        return True

    if (
        normalized_vendor_name
        in reference_sets["excluded_names"]
    ):
        return True

    return False


def add_cd_intercompany_flag(dataframe):
    """
    Add the Is Intercompany flag to a CD payment detail DataFrame.
    """
    result = dataframe.copy()

    if result.empty:
        result["Is Intercompany"] = (
            pd.Series(dtype=bool)
        )

        return result

    reference_sets = (
        get_cd_intercompany_reference_sets()
    )

    result["Is Intercompany"] = result.apply(
        lambda row: is_cd_intercompany_vendor(
            vendor_code=row.get(
                "Vendor Code",
                "",
            ),
            vendor_name=row.get(
                "Vendor Name",
                "",
            ),
            reference_sets=reference_sets,
        ),
        axis=1,
    )

    return result


def exclude_cd_intercompany_payments(
    dataframe,
):
    """
    Exclude identified intercompany vendors.
    """
    flagged_dataframe = (
        add_cd_intercompany_flag(
            dataframe
        )
    )

    return flagged_dataframe[
        flagged_dataframe[
            "Is Intercompany"
        ]
        != True
    ].drop(
        columns=["Is Intercompany"]
    ).copy()


def open_or_create_cd_output_workbook(
    output_file,
):
    """
    Open the shared CD output workbook or create an empty one.
    """
    output_file = Path(output_file)

    if output_file.exists():
        return load_workbook(output_file)

    workbook = Workbook()
    default_sheet = workbook.active

    workbook.remove(default_sheet)

    return workbook


def recreate_cd_sheet(
    workbook,
    sheet_name,
):
    """
    Recreate only the requested CD output sheet.
    """
    if sheet_name in workbook.sheetnames:
        del workbook[sheet_name]

    return workbook.create_sheet(
        sheet_name
    )


def save_cd_output_workbook(
    workbook,
    output_file,
):
    """
    Save the shared CD workbook.
    """
    output_file = Path(output_file)

    output_file.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    try:
        workbook.save(output_file)

    except PermissionError as error:
        raise PermissionError(
            f"Could not save output workbook: "
            f"{output_file}. The file may be open "
            "in Excel or locked by OneDrive. "
            "Close the workbook and run again."
        ) from error


def write_dataframe_to_sheet(
    worksheet,
    dataframe,
):
    """
    Write a DataFrame into an openpyxl worksheet.
    """
    for (
        column_index,
        column_name,
    ) in enumerate(
        dataframe.columns,
        start=1,
    ):
        worksheet.cell(
            row=1,
            column=column_index,
            value=column_name,
        )

    for (
        row_index,
        row,
    ) in enumerate(
        dataframe.itertuples(index=False),
        start=2,
    ):
        for (
            column_index,
            value,
        ) in enumerate(
            row,
            start=1,
        ):
            if pd.isna(value):
                value = None

            if isinstance(
                value,
                pd.Timestamp,
            ):
                value = (
                    value.to_pydatetime()
                )

            worksheet.cell(
                row=row_index,
                column=column_index,
                value=value,
            )


def apply_standard_cd_formatting(
    worksheet,
    dataframe,
    date_columns=None,
    amount_columns=None,
    integer_columns=None,
):
    """
    Apply common CD header, filter, width, date, amount, and integer formats.
    """
    date_columns = set(
        date_columns or []
    )

    amount_columns = set(
        amount_columns or []
    )

    integer_columns = set(
        integer_columns or []
    )

    header_fill = PatternFill(
        fill_type="solid",
        fgColor=CD_HEADER_FILL,
    )

    for cell in worksheet[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill

    worksheet.freeze_panes = "A2"

    if (
        worksheet.max_row >= 1
        and worksheet.max_column >= 1
    ):
        worksheet.auto_filter.ref = (
            worksheet.dimensions
        )

    column_name_by_index = {
        index + 1: column_name
        for index, column_name
        in enumerate(dataframe.columns)
    }

    for column_index in range(
        1,
        worksheet.max_column + 1,
    ):
        column_letter = get_column_letter(
            column_index
        )

        column_name = (
            column_name_by_index.get(
                column_index,
                "",
            )
        )

        max_length = len(
            str(column_name)
        )

        for cell in worksheet[column_letter]:
            if cell.value is None:
                continue

            max_length = max(
                max_length,
                len(str(cell.value)),
            )

            if cell.row == 1:
                continue

            if column_name in date_columns:
                cell.number_format = (
                    CD_DATE_FORMAT
                )

            if column_name in amount_columns:
                cell.number_format = (
                    CD_AMOUNT_FORMAT
                )

            if column_name in integer_columns:
                cell.number_format = (
                    CD_INTEGER_FORMAT
                )

        worksheet.column_dimensions[
            column_letter
        ].width = min(
            max_length + 2,
            45,
        )
