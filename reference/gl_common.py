"""
Common GL utilities.

Centralized utilities shared by the independent GL controls.

Supported current-period inputs:
- LBR GL_MD_YYYYMMDD.xlsx
- FxRates_YYYYMMDD.xlsx
- LBR GL_04_YYYYMMDD.xlsx
- LBR GL_JE_BSIS_1.xlsx
- LBR GL_JE_BSIS_2.xlsx
- LBR GL_JE_BSIS_3.xlsx
- LBR GL_JE_BSAS_1.xlsx
- LBR GL_JE_BSAS_2.xlsx
- LBR GL_JE_BSAS_3.xlsx

All matching BSIS and BSAS files are read and concatenated in natural
numeric order. For example, files 1, 2 and 10 are processed in that order.

The loaders accept previous Portuguese/SAP headers and normalize current
English SAP export headers to the canonical names expected by GL controls.

Pandas duplicate-column suffixes such as .1 and .2 are preserved.
Historic journal files are intentionally handled by the controls that
explicitly require historical populations.
"""

import re
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


GL_BSIS_INPUT_KEYWORDS = [
    "LBR GL_JE_BSIS",
    "LBR_GL_JE_BSIS",
]


GL_BSAS_INPUT_KEYWORDS = [
    "LBR GL_JE_BSAS",
    "LBR_GL_JE_BSAS",
]


GL_MASTER_INPUT_KEYWORDS = [
    "LBR GL_MD",
    "LBR_GL_MD",
]


GL_FX_INPUT_KEYWORDS = [
    "FxRates",
    "FX Rates",
    "Fx Rates",
]


GL04_WORKFLOW_INPUT_KEYWORDS = [
    "LBR GL_04",
    "LBR_GL_04",
]


GL_JOURNAL_HEADER_ALIASES = {
    "cocd": "Empr",
    "documentno": "Nº doc.",
    "year": "Ano",
    "doc. type": "Tp.doc.",
    "doc. date": "Data doc.",
    "pstng date": "Dt.lçto.",
    "entry dte": "Dt.entr.",
    "time": "Hora",
    "user name": "Nome do usuário",
    "tcode": "CódT",
    "document header text": "Texto cabeçalho documento",
    "crcy": "Moeda",
    "reference": "Referência",
    "ref. tran.": "Trans.ref.",
    "ref. key": "Chave ref.",
    "itm": "Itm",
    "g/l": "Razão",
    "d/c": "D/C",
    "amount in lc": "Montante em MI",
    "amount": "Montante",
    "text": "Texto",
    "cost ctr": "Centro cst",
    "profit ctr": "Cen.lucro",
    "order": "Ordem",
    "assignment": "Atribuição",
    "clrng doc.": "DocCompens",
    "clearing": "Compensaç.",
    "rev. with": "Estorno c/",
    "reason": "Motiv.est",
    "reversal": "Estorno",
    "parked by": "Pré-edição",
}


GL_MASTER_HEADER_ALIASES = {
    "cocd": "BUKRS",
    "company name": "BUTXT",
    "chac": "KTOPL",
    "crcy": "WAERS",
    "g/l acct": "SAKNR",
    "date": "ERDAT",
    "created by": "ERNAM",
    "b/s acct": "XBILK",
    "group acct": "BILKT",
    "language": "SPRAS",
    "short text": "TXT20",
    "long text": "TXT50",
}


GL_HEADER_FILL = "D9EAF7"
GL_DATE_FORMAT = "dd/mm/yyyy"
GL_AMOUNT_FORMAT = "#,##0.00"
GL_INTEGER_FORMAT = "#,##0"


def has_allowed_extension(file_path):
    file_path = Path(file_path)

    return (
        file_path.suffix.lower()
        in ALLOWED_EXTENSIONS
    )


def normalize_text(value):
    if value is None:
        return ""

    value_text = str(value).strip()

    if value_text.lower() in {
        "nan",
        "none",
        "<na>",
    }:
        return ""

    return value_text


def normalize_header(value):
    return normalize_text(
        value
    ).lower()


def normalize_company_output(value):
    value_text = normalize_text(
        value
    )

    if value_text == "":
        return ""

    if value_text.endswith(".0"):
        value_text = value_text[:-2]

    if value_text.isdigit():
        return str(
            int(value_text)
        )

    return value_text


def normalize_company_match(value):
    return normalize_company_output(
        value
    )


def normalize_code_keep_leading_zeroes(
    value,
):
    value_text = normalize_text(
        value
    )

    if value_text.endswith(".0"):
        value_text = value_text[:-2]

    return value_text


def parse_companies_filter(value):
    value_text = normalize_text(
        value
    )

    if value_text == "":
        return []

    if value_text.upper() in {
        "ALL",
        "TODAS",
        "TODOS",
    }:
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
        item = normalize_text(
            item
        )

        if item == "":
            continue

        if item.upper() in {
            "ALL",
            "TODAS",
            "TODOS",
        }:
            return []

        companies.append(
            normalize_company_match(
                item
            )
        )

    return companies


def to_datetime_value(value):
    if pd.isna(value):
        return pd.NaT

    if isinstance(
        value,
        pd.Timestamp,
    ):
        return value

    value_text = normalize_text(
        value
    )

    if value_text == "":
        return pd.NaT

    if value_text.endswith(".0"):
        value_text = value_text[:-2]

    if (
        len(value_text) == 8
        and value_text.isdigit()
    ):
        return pd.to_datetime(
            value_text,
            format="%Y%m%d",
            errors="coerce",
        )

    if (
        len(value_text) >= 10
        and value_text[4:5] == "-"
        and value_text[7:8] == "-"
    ):
        return pd.to_datetime(
            value_text[:10],
            format="%Y-%m-%d",
            errors="coerce",
        )

    return pd.to_datetime(
        value_text,
        errors="coerce",
        dayfirst=True,
    )


def parse_number(value):
    if pd.isna(value):
        return pd.NA

    if isinstance(
        value,
        (
            int,
            float,
        ),
    ):
        return value

    value_text = normalize_text(
        value
    )

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
        last_comma = value_text.rfind(
            ","
        )

        last_dot = value_text.rfind(
            "."
        )

        if last_comma > last_dot:
            value_text = (
                value_text.replace(
                    ".",
                    "",
                )
            )

            value_text = (
                value_text.replace(
                    ",",
                    ".",
                )
            )
        else:
            value_text = (
                value_text.replace(
                    ",",
                    "",
                )
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
        parsed_number = float(
            value_text
        )
    except ValueError:
        return pd.NA

    if is_negative:
        parsed_number = (
            parsed_number * -1
        )

    return parsed_number


def get_period_suffix(module_config):
    to_date = module_config.get(
        "to",
        "",
    )

    parsed_date = to_datetime_value(
        to_date
    )

    if pd.isna(parsed_date):
        raise ValueError(
            "Could not determine GL output "
            "period because module TO date "
            "is empty or invalid."
        )

    return parsed_date.strftime(
        "%Y%m%d"
    )


def get_gl_output_file(context):
    """
    Return one independent GL output workbook per control.

    Examples:
    - GL_001 -> LBR_Results_GL01_20260731.xlsx
    - GL_009 -> LBR_Results_GL09_20260731.xlsx
    - GL_015 -> LBR_Results_GL15_20260731.xlsx

    The period suffix continues to come from the GL module TO date.
    """
    output_folder = Path(
        context["output_folder"]
    )

    module_config = context["module"]

    control_config = context.get(
        "control",
        {},
    )

    period_suffix = get_period_suffix(
        module_config
    )

    control_id = normalize_text(
        control_config.get(
            "id_control",
            "",
        )
    ).upper()

    if not control_id.startswith("GL"):
        raise ValueError(
            "Could not determine the GL output "
            "file because "
            "context['control']['id_control'] "
            f"is invalid: {control_id!r}"
        )

    control_number = "".join(
        character
        for character in control_id
        if character.isdigit()
    )

    if control_number == "":
        raise ValueError(
            "Could not determine the GL output "
            "file because control ID does not "
            f"contain a number: {control_id!r}"
        )

    control_code = (
        f"GL{int(control_number):02d}"
    )

    return (
        output_folder
        / (
            f"LBR_Results_{control_code}_"
            f"{period_suffix}.xlsx"
        )
    )


def find_all_files_containing(
    base_folder,
    text_to_find,
):
    base_folder = Path(
        base_folder
    )

    text_to_find = normalize_text(
        text_to_find
    ).lower()

    matched_files = []

    for file_path in base_folder.rglob(
        "*"
    ):
        if not file_path.is_file():
            continue

        if file_path.name.startswith(
            "~$"
        ):
            continue

        if not has_allowed_extension(
            file_path
        ):
            continue

        if (
            text_to_find
            in file_path.name.lower()
        ):
            matched_files.append(
                file_path
            )

    return sorted(
        matched_files
    )


def find_optional_period_file(
    context,
    keyword_prefix,
):
    input_folder = Path(
        context["input_folder"]
    )

    period_suffix = get_period_suffix(
        context["module"]
    )

    keyword = (
        f"{keyword_prefix}_"
        f"{period_suffix}"
    )

    matched_files = (
        find_all_files_containing(
            base_folder=input_folder,
            text_to_find=keyword,
        )
    )

    if len(matched_files) == 0:
        return None

    if len(matched_files) > 1:
        matched_files_text = "\n".join(
            f"- {file_path}"
            for file_path
            in matched_files
        )

        raise ValueError(
            "Multiple GL input files found "
            f"using keyword '{keyword}'. "
            "Expected zero or one file.\n"
            f"{matched_files_text}"
        )

    return matched_files[0]


def find_period_file_using_keywords(
    context,
    keyword_prefixes,
):
    for keyword_prefix in keyword_prefixes:
        input_file = (
            find_optional_period_file(
                context=context,
                keyword_prefix=(
                    keyword_prefix
                ),
            )
        )

        if input_file is not None:
            return input_file

    return None


def _natural_path_sort_key(file_path):
    """
    Return a natural sorting key for numbered GL extracts.

    Example:
    - LBR GL_JE_BSIS_1.xlsx
    - LBR GL_JE_BSIS_2.xlsx
    - LBR GL_JE_BSIS_10.xlsx

    The resulting order is 1, 2, 10 instead of 1, 10, 2.
    """
    file_path = Path(
        file_path
    )

    name_parts = re.split(
        r"(\d+)",
        file_path.name.casefold(),
    )

    natural_name_key = tuple(
        (
            0,
            int(part),
        )
        if part.isdigit()
        else (
            1,
            part,
        )
        for part in name_parts
    )

    return (
        str(
            file_path.parent
        ).casefold(),
        natural_name_key,
    )


def find_gl_source_input_files(
    context,
    keyword_prefixes,
):
    """
    Find every current monthly input for a BSIS or BSAS source.

    Temporary Excel files and files containing 'Historic' are excluded.
    """
    input_folder = Path(
        context["input_folder"]
    )

    normalized_prefixes = [
        normalize_text(
            keyword_prefix
        ).casefold()
        for keyword_prefix
        in keyword_prefixes
    ]

    matched_files = []

    for file_path in input_folder.rglob(
        "*"
    ):
        if not file_path.is_file():
            continue

        if file_path.name.startswith(
            "~$"
        ):
            continue

        if not has_allowed_extension(
            file_path
        ):
            continue

        normalized_stem = (
            file_path.stem.casefold()
        )

        if "historic" in normalized_stem:
            continue

        matches_prefix = any(
            (
                normalized_stem
                == keyword_prefix
            )
            or normalized_stem.startswith(
                f"{keyword_prefix}_"
            )
            for keyword_prefix
            in normalized_prefixes
        )

        if matches_prefix:
            matched_files.append(
                file_path
            )

    return sorted(
        set(matched_files),
        key=_natural_path_sort_key,
    )


def find_gl_bsis_input_files(context):
    """
    Find all current monthly BSIS extracts.
    """
    return find_gl_source_input_files(
        context=context,
        keyword_prefixes=(
            GL_BSIS_INPUT_KEYWORDS
        ),
    )


def find_gl_bsas_input_files(context):
    """
    Find all current monthly BSAS extracts.
    """
    return find_gl_source_input_files(
        context=context,
        keyword_prefixes=(
            GL_BSAS_INPUT_KEYWORDS
        ),
    )


def find_gl_bsis_input_file(context):
    """
    Retained for compatibility with controls that still request the
    period-specific BSIS input.
    """
    return find_period_file_using_keywords(
        context=context,
        keyword_prefixes=(
            GL_BSIS_INPUT_KEYWORDS
        ),
    )


def find_gl_bsas_input_file(context):
    """
    Retained for compatibility with controls that still request the
    period-specific BSAS input.
    """
    return find_period_file_using_keywords(
        context=context,
        keyword_prefixes=(
            GL_BSAS_INPUT_KEYWORDS
        ),
    )


def find_gl_master_input_file(context):
    return find_period_file_using_keywords(
        context=context,
        keyword_prefixes=(
            GL_MASTER_INPUT_KEYWORDS
        ),
    )


def find_gl_fx_rates_input_file(context):
    return find_period_file_using_keywords(
        context=context,
        keyword_prefixes=(
            GL_FX_INPUT_KEYWORDS
        ),
    )


def find_gl04_workflow_input_file(context):
    return find_period_file_using_keywords(
        context=context,
        keyword_prefixes=(
            GL04_WORKFLOW_INPUT_KEYWORDS
        ),
    )


def read_input_file(
    file_path,
    sheet_name=0,
    usecols=None,
):
    """
    Read one supported GL input file.

    The optional usecols argument is passed directly to pandas so controls
    can avoid loading columns that are not required.
    """
    file_path = Path(
        file_path
    )

    if not file_path.exists():
        raise FileNotFoundError(
            "Input file not found: "
            f"{file_path}"
        )

    extension = (
        file_path.suffix.lower()
    )

    if extension in {
        ".xlsx",
        ".xls",
    }:
        return pd.read_excel(
            file_path,
            sheet_name=sheet_name,
            dtype=object,
            usecols=usecols,
        )

    if extension == ".csv":
        return pd.read_csv(
            file_path,
            dtype=object,
            usecols=usecols,
        )

    if extension == ".txt":
        return pd.read_csv(
            file_path,
            sep="\t",
            dtype=object,
            usecols=usecols,
        )

    raise ValueError(
        "Unsupported input file extension: "
        f"{file_path.suffix}"
    )


def _split_pandas_duplicate_suffix(
    column_name,
):
    column_text = normalize_text(
        column_name
    )

    base, separator, suffix = (
        column_text.rpartition(".")
    )

    if (
        separator
        and suffix.isdigit()
    ):
        return base, suffix

    return column_text, None


def normalize_gl_input_columns(
    dataframe,
    source_type="journal",
):
    """
    Normalize current English SAP headers.

    Duplicate suffixes generated by pandas are preserved.

    Examples:
    - CoCd -> Empr
    - CoCd.1 -> Empr.1
    - DocumentNo -> Nº doc.
    - DocumentNo.1 -> Nº doc..1
    - Crcy -> Moeda
    - Crcy.2 -> Moeda.2
    """
    if source_type == "master":
        aliases = (
            GL_MASTER_HEADER_ALIASES
        )
    else:
        aliases = (
            GL_JOURNAL_HEADER_ALIASES
        )

    normalized_columns = []
    used_columns = set()

    for column in dataframe.columns:
        (
            base,
            duplicate_suffix,
        ) = (
            _split_pandas_duplicate_suffix(
                column
            )
        )

        normalized_base = aliases.get(
            normalize_header(base),
            normalize_text(base),
        )

        if duplicate_suffix is None:
            normalized_column = (
                normalized_base
            )
        else:
            normalized_column = (
                f"{normalized_base}."
                f"{duplicate_suffix}"
            )

        if normalized_column in used_columns:
            suffix_number = 1

            candidate = (
                f"{normalized_column}."
                f"{suffix_number}"
            )

            while candidate in used_columns:
                suffix_number += 1

                candidate = (
                    f"{normalized_column}."
                    f"{suffix_number}"
                )

            normalized_column = candidate

        normalized_columns.append(
            normalized_column
        )

        used_columns.add(
            normalized_column
        )

    result = dataframe.copy()
    result.columns = normalized_columns

    return result


def clean_dataframe(dataframe):
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
        for column
        in result.columns
    ]

    return result

def find_column(
    dataframe,
    possible_names,
):
    normalized_lookup = {
        normalize_header(column): column
        for column
        in dataframe.columns
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
    source_name="GL input",
):
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
                f"{logical_name}: expected one "
                f"of {possible_names}"
            )
        else:
            resolved_columns[
                logical_name
            ] = column_name

    if missing_columns:
        raise ValueError(
            f"Missing required columns in "
            f"{source_name}:\n- "
            + "\n- ".join(
                missing_columns
            )
        )

    return resolved_columns


def get_optional_column(
    dataframe,
    possible_names,
):
    return find_column(
        dataframe,
        possible_names,
    )


def get_value(
    row,
    column_name,
    default="",
):
    if column_name is None:
        return default

    return row.get(
        column_name,
        default,
    )


def filter_by_company(
    dataframe,
    company_column,
    companies_filter,
):
    result = dataframe.copy()

    companies = parse_companies_filter(
        companies_filter
    )

    if not companies:
        return result

    normalized_company = result[
        company_column
    ].apply(
        normalize_company_match
    )

    return result[
        normalized_company.isin(
            companies
        )
    ].copy()


def apply_sap_debit_credit_sign(
    amount_value,
    debit_credit_indicator,
):
    amount = parse_number(
        amount_value
    )

    indicator = normalize_text(
        debit_credit_indicator
    ).upper()

    if pd.isna(amount):
        return pd.NA

    amount = abs(
        float(amount)
    )

    if indicator in {
        "H",
        "C",
    }:
        return amount * -1

    return amount


def build_document_key(
    company_code,
    fiscal_year,
    document_number,
    line_item,
):
    return "|".join(
        [
            normalize_company_output(
                company_code
            ),
            normalize_code_keep_leading_zeroes(
                fiscal_year
            ),
            normalize_code_keep_leading_zeroes(
                document_number
            ),
            normalize_code_keep_leading_zeroes(
                line_item
            ),
        ]
    )



GL_JOURNAL_IDENTITY_COLUMN_ALIASES = {
    "company_code": [
        "Empr",
        "CoCd",
        "BUKRS",
        "BKPF-BUKRS",
        "BSIS-BUKRS",
        "BSAS-BUKRS",
        "Company Code",
    ],
    "fiscal_year": [
        "Ano",
        "Year",
        "GJAHR",
        "BKPF-GJAHR",
        "BSIS-GJAHR",
        "BSAS-GJAHR",
        "Fiscal Year",
    ],
    "document_number": [
        "N? doc.",
        "N? doc",
        "DocumentNo",
        "Document Number",
        "BELNR",
        "BKPF-BELNR",
        "BSIS-BELNR",
        "BSAS-BELNR",
    ],
    "line_item": [
        "Itm",
        "Line Item",
        "Line Number",
        "BUZEI",
        "BSIS-BUZEI",
        "BSAS-BUZEI",
    ],
}


def remove_non_record_gl_journal_rows(
    dataframe,
):
    """
    Remove rows that do not contain a GL journal identity.

    SAP exports can contain visually empty footer rows that are not removed
    by dropna(how="all") because one or more cells contain empty strings,
    whitespace or invisible Unicode formatting characters.

    A row is considered a non-record only when every available journal
    identity field is empty:
    - company code;
    - fiscal year;
    - document number;
    - line item.

    The function deliberately does not remove duplicate accounting rows.
    """
    if dataframe.empty:
        return dataframe.copy(), 0

    resolved_identity_columns = []

    for possible_names in (
        GL_JOURNAL_IDENTITY_COLUMN_ALIASES.values()
    ):
        column_name = find_column(
            dataframe,
            possible_names,
        )

        if (
            column_name is not None
            and column_name
            not in resolved_identity_columns
        ):
            resolved_identity_columns.append(
                column_name
            )

    if not resolved_identity_columns:
        return dataframe.copy(), 0

    # Require enough identity information to avoid discarding rows from an
    # unexpected file layout based on one coincidentally named column.
    if len(resolved_identity_columns) < 2:
        print(
            "GL non-record row cleanup skipped "
            "because fewer than two journal "
            "identity columns were found. "
            "Columns found: "
            + ", ".join(
                resolved_identity_columns
            )
        )

        return dataframe.copy(), 0

    empty_identity_mask = pd.Series(
        True,
        index=dataframe.index,
        dtype=bool,
    )

    for column_name in resolved_identity_columns:
        normalized_values = (
            dataframe[column_name]
            .astype("string")
            .fillna("")
            .str.replace(
                r"[\s\u00a0\u200b\u200c\u200d\u2060\ufeff]+",
                "",
                regex=True,
            )
            .str.casefold()
        )

        empty_values = normalized_values.isin(
            {
                "",
                "nan",
                "none",
                "<na>",
            }
        )

        empty_identity_mask = (
            empty_identity_mask
            & empty_values
        )

    removed_count = int(
        empty_identity_mask.sum()
    )

    if removed_count == 0:
        return dataframe.copy(), 0

    cleaned_dataframe = (
        dataframe.loc[
            ~empty_identity_mask
        ]
        .copy()
        .reset_index(drop=True)
    )

    return (
        cleaned_dataframe,
        removed_count,
    )


def load_gl_source_data(
    context,
    source_name,
    input_file,
):
    """
    Read, clean and normalize one BSIS or BSAS input file.
    """
    if input_file is None:
        print(
            f"GL {source_name} input file: "
            "not found."
        )

        return pd.DataFrame()

    input_file = Path(
        input_file
    )

    print(
        f"GL {source_name} input file: "
        f"{input_file}"
    )

    dataframe = read_input_file(
        input_file
    )

    dataframe = clean_dataframe(
        dataframe
    )

    dataframe = normalize_gl_input_columns(
        dataframe,
        source_type="journal",
    )

    (
        dataframe,
        non_record_rows_removed,
    ) = remove_non_record_gl_journal_rows(
        dataframe
    )

    if non_record_rows_removed:
        print(
            f"GL {source_name} non-record rows "
            f"removed from {input_file.name}: "
            f"{non_record_rows_removed}"
        )

    dataframe["Source"] = (
        source_name
    )

    dataframe["_INPUT_FILE"] = (
        input_file.name
    )

    print(
        f"GL {source_name} rows loaded "
        f"from {input_file.name}: "
        f"{len(dataframe)}"
    )

    return dataframe


def validate_no_duplicate_gl_rows(
    dataframe,
    source_name,
):
    """
    Stop processing if identical journal rows occur in monthly inputs.

    Source and _INPUT_FILE are excluded from the comparison because they
    are technical traceability columns. Therefore, an accounting row
    present in two different files is still identified as duplicated.

    The function does not remove rows automatically.
    """
    if dataframe.empty:
        return dataframe

    technical_columns = {
        "Source",
        "_INPUT_FILE",
    }

    comparison_columns = [
        column
        for column
        in dataframe.columns
        if column
        not in technical_columns
    ]

    if not comparison_columns:
        print(
            f"GL {source_name} duplicate "
            "validation skipped because no "
            "accounting columns were found."
        )

        return dataframe

    duplicate_mask = dataframe.duplicated(
        subset=comparison_columns,
        keep=False,
    )

    if not duplicate_mask.any():
        print(
            f"GL {source_name} duplicate "
            "validation: OK. "
            f"{len(dataframe)} rows checked."
        )

        return dataframe

    duplicate_rows = dataframe.loc[
        duplicate_mask
    ].copy()

    duplicate_row_count = len(
        duplicate_rows
    )

    duplicate_record_count = (
        duplicate_rows.drop_duplicates(
            subset=comparison_columns
        ).shape[0]
    )

    if "_INPUT_FILE" in duplicate_rows.columns:
        affected_files = sorted(
            duplicate_rows[
                "_INPUT_FILE"
            ]
            .dropna()
            .astype(str)
            .unique()
        )
    else:
        affected_files = []

    if affected_files:
        affected_files_text = (
            ", ".join(
                affected_files
            )
        )
    else:
        affected_files_text = (
            "unknown input files"
        )

    raise ValueError(
        f"GL {source_name} contains "
        f"{duplicate_record_count} duplicated "
        "records represented by "
        f"{duplicate_row_count} rows. "
        "No rows were removed. "
        "Review these input files: "
        f"{affected_files_text}"
    )


def load_gl_source_files(
    context,
    source_name,
    input_files,
):
    """
    Read, normalize and concatenate all monthly files for one GL source.

    Each monthly file is cleaned and normalized before concatenation.
    Identical accounting rows are validated after concatenation.

    No records are removed automatically. If duplicates are found, the
    execution stops with a ValueError identifying the affected files.
    """
    input_files = list(
        input_files or []
    )

    if not input_files:
        print(
            f"GL {source_name} input files: "
            "not found."
        )

        return pd.DataFrame()

    dataframes = []

    for input_file in input_files:
        dataframe = load_gl_source_data(
            context=context,
            source_name=source_name,
            input_file=input_file,
        )

        if dataframe.empty:
            continue

        dataframes.append(
            dataframe
        )

    if not dataframes:
        print(
            f"GL {source_name} input files "
            "were found, but they contained "
            "no rows."
        )

        return pd.DataFrame()

    combined_dataframe = pd.concat(
        dataframes,
        ignore_index=True,
        sort=False,
    )

    print(
        f"GL {source_name} files read: "
        f"{len(input_files)}"
    )

    print(
        f"GL {source_name} total rows "
        "loaded before duplicate "
        f"validation: {len(combined_dataframe)}"
    )

    combined_dataframe = (
        validate_no_duplicate_gl_rows(
            dataframe=combined_dataframe,
            source_name=source_name,
        )
    )

    return combined_dataframe


# GL_PERIOD_PART_SELECTION_V1
def _get_gl_config_inclusive_month_count(context):
    """Return the inclusive number of configured calendar months."""
    module_config = context.get("module", {})

    from_date = pd.to_datetime(
        module_config.get("from", ""),
        errors="coerce",
    )
    to_date = pd.to_datetime(
        module_config.get("to", ""),
        errors="coerce",
    )

    if pd.isna(from_date) or pd.isna(to_date):
        raise ValueError(
            "GL input-file selection requires valid CONFIG FROM and TO dates."
        )

    from_date = pd.Timestamp(from_date).normalize()
    to_date = pd.Timestamp(to_date).normalize()

    if from_date > to_date:
        raise ValueError(
            "GL input-file selection requires CONFIG FROM to be earlier than "
            "or equal to CONFIG TO. "
            f"FROM={from_date:%Y-%m-%d}; TO={to_date:%Y-%m-%d}."
        )

    month_count = (
        (to_date.year - from_date.year) * 12
        + to_date.month
        - from_date.month
        + 1
    )

    return month_count, from_date, to_date


def _get_gl_numbered_part(path_value, source_name):
    """Return the exact trailing numeric part for a BSIS/BSAS filename."""
    path = Path(path_value)
    prefix = f"LBR GL_JE_{source_name}_"
    stem = path.stem

    if not stem.casefold().startswith(prefix.casefold()):
        return None

    suffix = stem[len(prefix):]

    if not suffix.isdigit():
        return None

    part_number = int(suffix)

    if part_number < 1:
        raise ValueError(
            f"GL {source_name} input file has an invalid numeric part: "
            f"{path.name}. Part numbers must start at 1."
        )

    return part_number


def select_gl_period_input_files(
    context,
    source_name,
    input_files,
    include_historical_parts=False,
):
    """
    Select numbered GL input parts according to CONFIG FROM and TO.

    Supported numbered layouts:

    1. Current-period only:
       - Number of parts equals the inclusive CONFIG month count.
       - Part 1 corresponds to CONFIG FROM.
       - The final part corresponds to CONFIG TO.

    2. Six historical months plus current period:
       - Parts 1 through 6 are historical.
       - Part 7 corresponds to CONFIG FROM.
       - The final part corresponds to CONFIG TO.

    Selection behavior:
    - GL08 passes include_historical_parts=True and receives all historical
      and current-period parts.
    - Other controls receive current-period parts only.

    Numbered parts must be consecutive and start at 1.
    """
    source_name = (
        str(source_name)
        .strip()
        .upper()
    )

    input_files = [
        Path(path)
        for path in input_files
    ]

    if not input_files:
        return []

    historical_month_count = 6
    numbered_files = {}
    unnumbered_files = []

    for input_file in input_files:
        part_number = (
            _get_gl_numbered_part(
                input_file,
                source_name,
            )
        )

        if part_number is None:
            unnumbered_files.append(
                input_file
            )

            continue

        if part_number in numbered_files:
            raise ValueError(
                f"GL {source_name} has more "
                "than one file for numbered "
                f"part _{part_number}: "
                f"{numbered_files[part_number].name}, "
                f"{input_file.name}. "
                "No files were read."
            )

        numbered_files[
            part_number
        ] = input_file

    if (
        numbered_files
        and unnumbered_files
    ):
        raise ValueError(
            f"GL {source_name} input "
            "discovery returned a mixture "
            "of numbered and unnumbered "
            "files. No files were read. "
            "Numbered files: "
            + ", ".join(
                numbered_files[number].name
                for number
                in sorted(numbered_files)
            )
            + ". Unnumbered files: "
            + ", ".join(
                path.name
                for path
                in unnumbered_files
            )
        )

    if not numbered_files:
        print(
            f"GL {source_name} "
            "input-file selection: "
            "legacy unnumbered mode"
        )

        print(
            f"GL {source_name} "
            "unnumbered files selected: "
            f"{len(unnumbered_files)}"
        )

        return input_files

    available_parts = sorted(
        numbered_files
    )

    maximum_part = max(
        available_parts
    )

    expected_consecutive_parts = list(
        range(
            1,
            maximum_part + 1,
        )
    )

    missing_sequence_parts = [
        part_number
        for part_number
        in expected_consecutive_parts
        if part_number
        not in numbered_files
    ]

    if missing_sequence_parts:
        raise FileNotFoundError(
            f"GL {source_name} numbered "
            "parts must be consecutive and "
            "start at 1. Missing parts: "
            + ", ".join(
                map(
                    str,
                    missing_sequence_parts,
                )
            )
            + ". Available parts: "
            + ", ".join(
                map(
                    str,
                    available_parts,
                )
            )
            + ". No GL source files "
            "were read."
        )

    (
        active_month_count,
        from_date,
        to_date,
    ) = (
        _get_gl_config_inclusive_month_count(
            context
        )
    )

    current_only_part_count = (
        active_month_count
    )

    historical_layout_part_count = (
        historical_month_count
        + active_month_count
    )

    available_part_count = len(
        available_parts
    )

    if (
        available_part_count
        == current_only_part_count
    ):
        detected_historical_month_count = 0

    elif (
        available_part_count
        == historical_layout_part_count
    ):
        detected_historical_month_count = (
            historical_month_count
        )

    else:
        raise ValueError(
            f"GL {source_name} found "
            f"{available_part_count} numbered "
            "parts for CONFIG "
            f"{from_date:%Y-%m-%d} to "
            f"{to_date:%Y-%m-%d}. "
            "Expected either "
            f"{current_only_part_count} parts "
            "for the current period only, or "
            f"{historical_layout_part_count} "
            "parts for six historical months "
            "plus the current period. "
            "Available parts: "
            + ", ".join(
                map(
                    str,
                    available_parts,
                )
            )
            + ". No GL source files "
            "were read."
        )

    first_active_part = (
        detected_historical_month_count
        + 1
    )

    last_active_part = (
        detected_historical_month_count
        + active_month_count
    )

    active_parts = list(
        range(
            first_active_part,
            last_active_part + 1,
        )
    )

    historical_parts = list(
        range(
            1,
            detected_historical_month_count
            + 1,
        )
    )

    if include_historical_parts:
        if (
            detected_historical_month_count
            != historical_month_count
        ):
            raise FileNotFoundError(
                f"GL {source_name} historical "
                "lookback was requested, but "
                "the six required historical "
                "monthly parts were not found. "
                f"CONFIG requires "
                f"{active_month_count} current "
                "months plus 6 historical "
                f"months, for a total of "
                f"{historical_layout_part_count} "
                "parts. Found "
                f"{available_part_count} parts. "
                "No GL source files were read."
            )

        selected_parts = (
            historical_parts
            + active_parts
        )

        print(
            f"GL {source_name} "
            "loading mode: six historical "
            "months plus configured period"
        )

    else:
        selected_parts = (
            active_parts
        )

        if historical_parts:
            print(
                f"GL {source_name} "
                "loading mode: configured "
                "period after six historical "
                "parts"
            )
        else:
            print(
                f"GL {source_name} "
                "loading mode: configured-"
                "period numbered parts"
            )

    print(
        "GL configured inclusive months: "
        f"{active_month_count} "
        f"({from_date:%Y-%m} through "
        f"{to_date:%Y-%m})"
    )

    if historical_parts:
        print(
            f"GL {source_name} historical "
            "parts detected: "
            + ", ".join(
                map(
                    str,
                    historical_parts,
                )
            )
        )
    else:
        print(
            f"GL {source_name} historical "
            "parts detected: none"
        )

    print(
        f"GL {source_name} active-period "
        "parts: "
        + ", ".join(
            map(
                str,
                active_parts,
            )
        )
    )

    print(
        f"GL {source_name} parts selected: "
        + ", ".join(
            map(
                str,
                selected_parts,
            )
        )
    )

    skipped_parts = [
        part_number
        for part_number
        in available_parts
        if part_number
        not in selected_parts
    ]

    print(
        f"GL {source_name} parts skipped: "
        + (
            ", ".join(
                map(
                    str,
                    skipped_parts,
                )
            )
            if skipped_parts
            else "none"
        )
    )

    return [
        numbered_files[
            part_number
        ]
        for part_number
        in selected_parts
    ]


def load_gl_bsis_data(
    context,
    include_historical_parts=False,
):
    input_files = find_gl_bsis_input_files(context)

    selected_files = select_gl_period_input_files(
        context=context,
        source_name="BSIS",
        input_files=input_files,
        include_historical_parts=include_historical_parts,
    )

    return load_gl_source_files(
        context=context,
        source_name="BSIS",
        input_files=selected_files,
    )


def load_gl_bsas_data(
    context,
    include_historical_parts=False,
):
    input_files = find_gl_bsas_input_files(context)

    selected_files = select_gl_period_input_files(
        context=context,
        source_name="BSAS",
        input_files=input_files,
        include_historical_parts=include_historical_parts,
    )

    return load_gl_source_files(
        context=context,
        source_name="BSAS",
        input_files=selected_files,
    )


def load_gl_master_data(context):
    input_file = (
        find_gl_master_input_file(
            context
        )
    )

    if input_file is None:
        print(
            "GL master input file: not found. "
            "Company and Account Name will "
            "be blank."
        )

        return pd.DataFrame()

    print(
        "GL master input file: "
        f"{input_file}"
    )

    dataframe = read_input_file(
        input_file
    )

    dataframe = clean_dataframe(
        dataframe
    )

    dataframe = normalize_gl_input_columns(
        dataframe,
        source_type="master",
    )

    print(
        "GL master rows loaded: "
        f"{len(dataframe)}"
    )

    return dataframe


def load_gl_fx_rates_data(context):
    input_file = (
        find_gl_fx_rates_input_file(
            context
        )
    )

    if input_file is None:
        print(
            "GL FxRates input file: not found. "
            "USD columns will be blank."
        )

        return pd.DataFrame()

    print(
        "GL FxRates input file: "
        f"{input_file}"
    )

    dataframe = read_input_file(
        input_file
    )

    dataframe = clean_dataframe(
        dataframe
    )

    print(
        "GL FxRates rows loaded: "
        f"{len(dataframe)}"
    )

    return dataframe


def load_gl04_workflow_data(context):
    """
    Load the period-specific GL04 workflow approval extract.

    Expected filename examples:
    - LBR GL_04_20260228.xlsx
    - LBR_GL_04_20260228.xlsx

    The suffix is obtained from the configured GL module TO date.
    Workflow headers are preserved as exported by SAP/SQVI.
    """
    input_file = (
        find_gl04_workflow_input_file(
            context
        )
    )

    if input_file is None:
        print(
            "GL04 Workflow input file: "
            "not found."
        )

        return pd.DataFrame()

    input_file = Path(
        input_file
    )

    print(
        "GL04 Workflow input file: "
        f"{input_file}"
    )

    dataframe = read_input_file(
        input_file
    )

    dataframe = clean_dataframe(
        dataframe
    )

    dataframe["Source"] = (
        "GL04 Workflow"
    )

    dataframe["_INPUT_FILE"] = (
        input_file.name
    )

    print(
        "GL04 Workflow rows loaded "
        f"from {input_file.name}: "
        f"{len(dataframe)}"
    )

    return dataframe


def build_company_name_map(
    master_dataframe,
):
    if master_dataframe.empty:
        return {}

    bukrs_column = get_optional_column(
        master_dataframe,
        [
            "BUKRS",
            "Empr",
            "Company Code",
            "CoCd",
        ],
    )

    company_name_column = (
        get_optional_column(
            master_dataframe,
            [
                "BUTXT",
                "Nome da firma",
                "Nome da empresa",
                "Empresa",
                "Company",
                "Company Name",
            ],
        )
    )

    if (
        bukrs_column is None
        or company_name_column is None
    ):
        return {}

    result = {}

    for _, row in (
        master_dataframe.iterrows()
    ):
        company_code = (
            normalize_company_output(
                row.get(
                    bukrs_column,
                    "",
                )
            )
        )

        company_name = normalize_text(
            row.get(
                company_name_column,
                "",
            )
        )

        if company_code == "":
            continue

        if company_code not in result:
            result[
                company_code
            ] = company_name

    return result


def build_gl_account_name_map(
    master_dataframe,
):
    if master_dataframe.empty:
        return {}

    account_column = get_optional_column(
        master_dataframe,
        [
            "SAKNR",
            "SAKNR.1",
            "HKONT",
            "Cta.Razão",
            "Cta.Razao",
            "Cta.Razão.1",
            "Cta.Razao.1",
            "Razão",
            "Razao",
            "Conta",
            "Account",
            "G/L Acct",
            "G/L Acct.1",
        ],
    )

    language_column = get_optional_column(
        master_dataframe,
        [
            "SPRAS",
            "Idioma",
            "Language",
        ],
    )

    text_column = get_optional_column(
        master_dataframe,
        [
            "TXT50",
            "TxtDescr",
            "TXT20",
            "Texto breve",
            "Texto",
            "Account Name",
            "Description",
            "Long Text",
            "Short Text",
        ],
    )

    if (
        account_column is None
        or text_column is None
    ):
        return {}

    language_priority = {
        "PT": 1,
        "P": 1,
        "ES": 2,
        "S": 2,
        "EN": 3,
        "E": 3,
    }

    candidate_by_account = {}

    for _, row in (
        master_dataframe.iterrows()
    ):
        account = (
            normalize_code_keep_leading_zeroes(
                row.get(
                    account_column,
                    "",
                )
            )
        )

        text = normalize_text(
            row.get(
                text_column,
                "",
            )
        )

        if (
            account == ""
            or text == ""
        ):
            continue

        if language_column:
            language = normalize_text(
                row.get(
                    language_column,
                    "",
                )
            ).upper()
        else:
            language = ""

        priority = language_priority.get(
            language,
            9,
        )

        current = candidate_by_account.get(
            account
        )

        if (
            current is None
            or priority
            < current["priority"]
        ):
            candidate_by_account[
                account
            ] = {
                "priority": priority,
                "text": text,
            }

    return {
        account: value["text"]
        for account, value
        in candidate_by_account.items()
    }


def resolve_fx_columns(
    fx_dataframe,
):
    return require_columns(
        dataframe=fx_dataframe,
        required_columns={
            "rate_type": [
                "CgCâ",
                "CgCa",
                "KURST",
                "TCot",
                "Tipo cotización",
                "Tipo cotizacion",
                "ExRt",
            ],
            "from_currency": [
                "De",
                "FCURR",
                "Moeda de procedência",
                "Moeda de procedencia",
                "From Currency",
                "From",
            ],
            "to_currency": [
                "Para",
                "TCURR",
                "Moeda destino",
                "To Currency",
                "To",
            ],
            "valid_from": [
                "Vál.desde",
                "Val.desde",
                "GDATU",
                "Data válida desde",
                "Data valida desde",
                "Valid From",
                "Valid from",
            ],
            "rate": [
                "Taxa câmbio",
                "Taxa cambio",
                "UKURS",
                "Cotação",
                "Cotacao",
                "Exchange Rate",
                "Exch. Rate",
                "Rate",
            ],
        },
        source_name="GL FxRates",
    )


def get_optional_fx_factor_columns(
    fx_dataframe,
):
    return {
        "factor_from": get_optional_column(
            fx_dataframe,
            [
                "Fator (origem)",
                "Fator origem",
                "FFACT",
                "From Factor",
                "Factor From",
                "Ratio (from)",
            ],
        ),
        "factor_to": get_optional_column(
            fx_dataframe,
            [
                "Fator (para)",
                "Fator para",
                "TFACT",
                "To Factor",
                "Factor To",
                "Ratio (to)",
            ],
        ),
    }


def normalize_fx_rates(
    fx_dataframe,
):
    if fx_dataframe.empty:
        return pd.DataFrame(
            columns=[
                "Rate Type",
                "From Currency",
                "To Currency",
                "Valid From",
                "Rate",
                "Factor From",
                "Factor To",
            ]
        )

    required_columns = (
        resolve_fx_columns(
            fx_dataframe
        )
    )

    optional_columns = (
        get_optional_fx_factor_columns(
            fx_dataframe
        )
    )

    result = pd.DataFrame(
        index=fx_dataframe.index
    )

    result["Rate Type"] = (
        fx_dataframe[
            required_columns[
                "rate_type"
            ]
        ]
        .apply(normalize_text)
        .str.upper()
    )

    result["From Currency"] = (
        fx_dataframe[
            required_columns[
                "from_currency"
            ]
        ]
        .apply(normalize_text)
        .str.upper()
    )

    result["To Currency"] = (
        fx_dataframe[
            required_columns[
                "to_currency"
            ]
        ]
        .apply(normalize_text)
        .str.upper()
    )

    result["Valid From"] = (
        fx_dataframe[
            required_columns[
                "valid_from"
            ]
        ].apply(
            to_datetime_value
        )
    )

    result["Rate"] = (
        fx_dataframe[
            required_columns[
                "rate"
            ]
        ].apply(
            parse_number
        )
    )

    factor_from_column = (
        optional_columns[
            "factor_from"
        ]
    )

    factor_to_column = (
        optional_columns[
            "factor_to"
        ]
    )

    if factor_from_column is None:
        result["Factor From"] = 1.0
    else:
        result["Factor From"] = (
            fx_dataframe[
                factor_from_column
            ]
            .apply(parse_number)
            .fillna(1.0)
        )

    if factor_to_column is None:
        result["Factor To"] = 1.0
    else:
        result["Factor To"] = (
            fx_dataframe[
                factor_to_column
            ]
            .apply(parse_number)
            .fillna(1.0)
        )

    result = result.dropna(
        subset=[
            "Valid From",
            "Rate",
        ]
    ).copy()

    result["Valid From"] = (
        pd.to_datetime(
            result[
                "Valid From"
            ],
            errors="coerce",
        )
        .dt.normalize()
    )

    result = result[
        (
            result[
                "Rate Type"
            ]
            != ""
        )
        & (
            result[
                "From Currency"
            ]
            != ""
        )
        & (
            result[
                "To Currency"
            ]
            != ""
        )
    ].copy()

    return result.reset_index(
        drop=True
    )

def calculate_adjusted_fx_rate(
    rate,
    factor_from,
    factor_to,
):
    if pd.isna(rate):
        return pd.NA

    rate = float(rate)

    if (
        pd.isna(factor_from)
        or float(factor_from) == 0
    ):
        factor_from = 1.0

    if (
        pd.isna(factor_to)
        or float(factor_to) == 0
    ):
        factor_to = 1.0

    return (
        rate
        * float(factor_to)
        / float(factor_from)
    )


def build_fx_date_candidates(
    requested_date,
    max_previous_days=10,
):
    requested_date = to_datetime_value(
        requested_date
    )

    if pd.isna(requested_date):
        return []

    requested_date = pd.Timestamp(
        requested_date
    ).normalize()

    return [
        requested_date
        - pd.Timedelta(
            days=days_back
        )
        for days_back
        in range(
            0,
            max_previous_days + 1,
        )
    ]


def select_fx_rate_to_usd(
    normalized_fx_dataframe,
    currency,
    requested_date,
):
    currency = normalize_text(
        currency
    ).upper()

    requested_date = to_datetime_value(
        requested_date
    )

    if currency == "":
        return None

    if pd.isna(requested_date):
        return None

    requested_date = pd.Timestamp(
        requested_date
    ).normalize()

    if currency in {
        "USD",
        "$",
    }:
        return {
            "fx_to_usd": 1.0,
            "usd_rate": 1.0,
            "rate_date": requested_date,
            "method": (
                "FxRates: USD currency, "
                "rate 1.000000"
            ),
        }

    if normalized_fx_dataframe.empty:
        return None

    rate_type_priority = [
        "EN",
        "M",
        "ZGB",
    ]

    for rate_type in (
        rate_type_priority
    ):
        date_candidates = (
            build_fx_date_candidates(
                requested_date
            )
        )

        for candidate_date in (
            date_candidates
        ):
            candidate_rates = (
                normalized_fx_dataframe[
                    (
                        normalized_fx_dataframe[
                            "Rate Type"
                        ]
                        == rate_type
                    )
                    & (
                        normalized_fx_dataframe[
                            "Valid From"
                        ]
                        == candidate_date
                    )
                ]
            )

            if candidate_rates.empty:
                continue

            direct_rates = (
                candidate_rates[
                    (
                        candidate_rates[
                            "From Currency"
                        ]
                        == currency
                    )
                    & (
                        candidate_rates[
                            "To Currency"
                        ]
                        == "USD"
                    )
                ]
            )

            if not direct_rates.empty:
                row = direct_rates.iloc[0]

                adjusted_rate = (
                    calculate_adjusted_fx_rate(
                        rate=row["Rate"],
                        factor_from=row[
                            "Factor From"
                        ],
                        factor_to=row[
                            "Factor To"
                        ],
                    )
                )

                if (
                    pd.isna(adjusted_rate)
                    or float(adjusted_rate)
                    == 0
                ):
                    continue

                fx_to_usd = float(
                    adjusted_rate
                )

                usd_rate = (
                    1 / fx_to_usd
                )

                return {
                    "fx_to_usd": (
                        fx_to_usd
                    ),
                    "usd_rate": (
                        usd_rate
                    ),
                    "rate_date": (
                        candidate_date
                    ),
                    "method": (
                        f"FxRates {rate_type}: "
                        f"{currency}->USD; "
                        "USD = amount * FX"
                    ),
                }

            inverse_rates = (
                candidate_rates[
                    (
                        candidate_rates[
                            "From Currency"
                        ]
                        == "USD"
                    )
                    & (
                        candidate_rates[
                            "To Currency"
                        ]
                        == currency
                    )
                ]
            )

            if not inverse_rates.empty:
                row = inverse_rates.iloc[0]

                adjusted_rate = (
                    calculate_adjusted_fx_rate(
                        rate=row["Rate"],
                        factor_from=row[
                            "Factor From"
                        ],
                        factor_to=row[
                            "Factor To"
                        ],
                    )
                )

                if (
                    pd.isna(adjusted_rate)
                    or float(adjusted_rate)
                    == 0
                ):
                    continue

                usd_rate = float(
                    adjusted_rate
                )

                fx_to_usd = (
                    1 / usd_rate
                )

                return {
                    "fx_to_usd": (
                        fx_to_usd
                    ),
                    "usd_rate": (
                        usd_rate
                    ),
                    "rate_date": (
                        candidate_date
                    ),
                    "method": (
                        f"FxRates {rate_type}: "
                        f"USD->{currency}; "
                        "USD = amount / "
                        "USD Rate"
                    ),
                }

    return None


def open_or_create_gl_output_workbook(
    output_file,
):
    output_file = Path(
        output_file
    )

    if output_file.exists():
        return load_workbook(
            output_file
        )

    workbook = Workbook()

    default_sheet = (
        workbook.active
    )

    workbook.remove(
        default_sheet
    )

    return workbook


def recreate_gl_sheet(
    workbook,
    sheet_name,
):
    if sheet_name in workbook.sheetnames:
        del workbook[sheet_name]

    return workbook.create_sheet(
        sheet_name
    )


def save_gl_output_workbook(
    workbook,
    output_file,
):
    output_file = Path(
        output_file
    )

    output_file.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    try:
        workbook.save(
            output_file
        )
    except PermissionError as error:
        raise PermissionError(
            "Could not save output workbook: "
            f"{output_file}. "
            "The file may be open in Excel "
            "or locked by OneDrive. Close "
            "the workbook and run again."
        ) from error


def write_dataframe_to_sheet(
    worksheet,
    dataframe,
):
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

    for row_index, row in enumerate(
        dataframe.itertuples(
            index=False
        ),
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


def apply_standard_gl_formatting(
    worksheet,
    dataframe,
    date_columns=None,
    amount_columns=None,
    integer_columns=None,
):
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
        fgColor=GL_HEADER_FILL,
    )

    for cell in worksheet[1]:
        cell.font = Font(
            bold=True
        )

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
        in enumerate(
            dataframe.columns
        )
    }

    for column_index in range(
        1,
        worksheet.max_column + 1,
    ):
        column_letter = (
            get_column_letter(
                column_index
            )
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

        for cell in worksheet[
            column_letter
        ]:
            if cell.value is None:
                continue

            max_length = max(
                max_length,
                len(
                    str(cell.value)
                ),
            )

            if cell.row == 1:
                continue

            if column_name in date_columns:
                cell.number_format = (
                    GL_DATE_FORMAT
                )

            if (
                column_name
                in amount_columns
            ):
                cell.number_format = (
                    GL_AMOUNT_FORMAT
                )

            if (
                column_name
                in integer_columns
            ):
                cell.number_format = (
                    GL_INTEGER_FORMAT
                )

        worksheet.column_dimensions[
            column_letter
        ].width = min(
            max_length + 2,
            45,
        )



# Maximum number of dataframe rows written to each Excel worksheet.
#
# This value can be changed without modifying the writer logic.
# Valid range: 1 through 1,048,575.
#
# Examples:
# - 250_000: smaller worksheets, more sheets;
# - 500_000: default balance between size and sheet count;
# - 1_000_000: large worksheets, fewer sheets;
# - 1_048_575: maximum supported by Excel with one header row.
DEFAULT_EXCEL_DATA_ROWS_PER_SHEET = 500_000


def write_single_sheet_workbook_fast(
    output_file,
    sheet_name,
    dataframe,
    date_columns=None,
    amount_columns=None,
    integer_columns=None,
    max_data_rows_per_sheet=None,
):
    """
    Write a dataframe to an XLSX workbook using XlsxWriter.

    Results are automatically split across worksheets according to
    max_data_rows_per_sheet.

    When max_data_rows_per_sheet is omitted, the value comes from
    DEFAULT_EXCEL_DATA_ROWS_PER_SHEET.

    The first worksheet keeps the requested name. Subsequent worksheets
    use suffixes such as _2, _3, etc.

    Excel permits 1,048,576 total rows per worksheet. Since row 1 contains
    headers, no worksheet may contain more than 1,048,575 dataframe rows.
    """
    output_file = Path(
        output_file
    )

    output_file.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    date_columns = set(
        date_columns or []
    )

    amount_columns = set(
        amount_columns or []
    )

    integer_columns = set(
        integer_columns or []
    )

    if output_file.exists():
        return False

    excel_max_total_rows = 1_048_576
    header_row_count = 1

    excel_max_data_rows = (
        excel_max_total_rows
        - header_row_count
    )

    if max_data_rows_per_sheet is None:
        max_data_rows_per_sheet = (
            DEFAULT_EXCEL_DATA_ROWS_PER_SHEET
        )

    if isinstance(
        max_data_rows_per_sheet,
        bool,
    ):
        raise TypeError(
            "max_data_rows_per_sheet must be "
            "an integer, not a boolean."
        )

    try:
        max_data_rows_per_sheet = int(
            max_data_rows_per_sheet
        )
    except (
        TypeError,
        ValueError,
    ) as error:
        raise TypeError(
            "max_data_rows_per_sheet must be "
            "a positive integer."
        ) from error

    if max_data_rows_per_sheet < 1:
        raise ValueError(
            "max_data_rows_per_sheet must be "
            "greater than zero."
        )

    if (
        max_data_rows_per_sheet
        > excel_max_data_rows
    ):
        raise ValueError(
            "max_data_rows_per_sheet cannot "
            f"exceed {excel_max_data_rows:,}. "
            "Excel reserves one worksheet row "
            "for the header."
        )

    dataframe_row_count = len(
        dataframe
    )

    sheet_count = max(
        1,
        (
            dataframe_row_count
            + max_data_rows_per_sheet
            - 1
        )
        // max_data_rows_per_sheet,
    )

    def build_sheet_name(
        base_name,
        sheet_number,
    ):
        base_name = str(
            base_name
        )

        if sheet_number == 1:
            suffix = ""
        else:
            suffix = (
                f"_{sheet_number}"
            )

        maximum_base_length = (
            31 - len(suffix)
        )

        return (
            base_name[:maximum_base_length]
            + suffix
        )

    print(
        "Excel writer configuration: "
        f"{max_data_rows_per_sheet:,} "
        "data rows per worksheet."
    )

    print(
        "Excel worksheets required: "
        f"{sheet_count}."
    )

    try:
        with pd.ExcelWriter(
            output_file,
            engine="xlsxwriter",
            datetime_format=(
                "dd/mm/yyyy"
            ),
            date_format=(
                "dd/mm/yyyy"
            ),
        ) as writer:
            workbook = writer.book

            header_format = (
                workbook.add_format(
                    {
                        "bold": True,
                        "bg_color": (
                            "#D9EAF7"
                        ),
                        "border": 0,
                    }
                )
            )

            date_format = (
                workbook.add_format(
                    {
                        "num_format": (
                            "dd/mm/yyyy"
                        ),
                    }
                )
            )

            amount_format = (
                workbook.add_format(
                    {
                        "num_format": (
                            "#,##0.00"
                        ),
                    }
                )
            )

            integer_format = (
                workbook.add_format(
                    {
                        "num_format": (
                            "#,##0"
                        ),
                    }
                )
            )

            for sheet_index in range(
                sheet_count
            ):
                start_row = (
                    sheet_index
                    * max_data_rows_per_sheet
                )

                stop_row = min(
                    start_row
                    + max_data_rows_per_sheet,
                    dataframe_row_count,
                )

                current_dataframe = (
                    dataframe.iloc[
                        start_row:stop_row
                    ]
                )

                current_sheet_name = (
                    build_sheet_name(
                        sheet_name,
                        sheet_index + 1,
                    )
                )

                current_dataframe.to_excel(
                    writer,
                    sheet_name=(
                        current_sheet_name
                    ),
                    index=False,
                )

                worksheet = (
                    writer.sheets[
                        current_sheet_name
                    ]
                )

                for (
                    column_index,
                    column_name,
                ) in enumerate(
                    dataframe.columns
                ):
                    worksheet.write(
                        0,
                        column_index,
                        column_name,
                        header_format,
                    )

                    column_width = min(
                        max(
                            len(
                                str(column_name)
                            ) + 2,
                            12,
                        ),
                        45,
                    )

                    if (
                        column_name
                        in date_columns
                    ):
                        cell_format = (
                            date_format
                        )
                    elif (
                        column_name
                        in amount_columns
                    ):
                        cell_format = (
                            amount_format
                        )
                    elif (
                        column_name
                        in integer_columns
                    ):
                        cell_format = (
                            integer_format
                        )
                    else:
                        cell_format = None

                    worksheet.set_column(
                        column_index,
                        column_index,
                        column_width,
                        cell_format,
                    )

                last_data_row = max(
                    len(
                        current_dataframe
                    ),
                    1,
                )

                last_column = max(
                    len(
                        dataframe.columns
                    ) - 1,
                    0,
                )

                worksheet.autofilter(
                    0,
                    0,
                    last_data_row,
                    last_column,
                )

                worksheet.freeze_panes(
                    1,
                    0,
                )

                if current_dataframe.empty:
                    source_range_text = (
                        "no data rows"
                    )
                else:
                    source_range_text = (
                        f"source rows "
                        f"{start_row + 1:,}-"
                        f"{stop_row:,}"
                    )

                print(
                    f"Worksheet "
                    f"{current_sheet_name} written: "
                    f"{len(current_dataframe):,} rows "
                    f"({source_range_text})."
                )

            if sheet_count > 1:
                print(
                    "Excel output split across "
                    f"{sheet_count} worksheets."
                )

            print(
                "Total dataframe rows written: "
                f"{dataframe_row_count:,}."
            )

    except (
        ModuleNotFoundError,
        ImportError,
    ):
        print(
            "Fast writer is not available "
            "because XlsxWriter is not "
            "installed. Falling back to "
            "openpyxl writer."
        )

        return False

    except PermissionError as error:
        raise PermissionError(
            "Could not save output workbook: "
            f"{output_file}. "
            "The file may be open in Excel "
            "or locked by OneDrive. Close "
            "the workbook and run again."
        ) from error

    return True


