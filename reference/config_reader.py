"""
Configuration reader.

This module reads Config.xlsx.

Workbook structure expected:
1. Sheet CONFIG:
   - First column contains module names such as AR, CD, FAM.
   - Columns include FROM, TO, COMPANIES, PARAM1, PARAM2, Execute.
   - Execute = T means the module is active.

2. One sheet per module:
   - Example: AR
   - Contains ID Control, Execute, PARAM1, PARAM2, NAME, Description.
   - Execute = T means that control should be executed.

This module does not execute controls.
It only returns which modules and controls are active.
"""

import warnings

import pandas as pd


warnings.filterwarnings(
    "ignore",
    message="Workbook contains no default style, apply openpyxl's default",
    category=UserWarning,
)


CONFIG_SHEET = "CONFIG"


def normalize_header(value):
    """
    Normalize column headers.
    """
    if value is None:
        return ""

    return str(value).strip().lower()


def normalize_execute_value(value):
    """
    Normalize Execute values.

    True values:
    - T
    - TRUE
    - YES
    - Y
    - 1
    """
    if value is None:
        return False

    value_text = str(value).strip().upper()

    return value_text in ["T", "TRUE", "YES", "Y", "1"]


def format_date(value):
    """
    Format dates as YYYY-MM-DD when possible.
    """
    if pd.isna(value):
        return ""

    parsed_date = pd.to_datetime(value, errors="coerce")

    if pd.isna(parsed_date):
        return value

    return parsed_date.strftime("%Y-%m-%d")


def read_excel_sheet(file_path, sheet_name):
    """
    Read an Excel sheet and clean empty rows/columns.
    """
    dataframe = pd.read_excel(file_path, sheet_name=sheet_name)

    dataframe = dataframe.dropna(axis=0, how="all")
    dataframe = dataframe.dropna(axis=1, how="all")

    dataframe.columns = [
        str(column).strip()
        for column in dataframe.columns
    ]

    return dataframe


def find_column(dataframe, possible_names):
    """
    Find a column using one or more possible names.
    """
    normalized_lookup = {
        normalize_header(column): column
        for column in dataframe.columns
    }

    for possible_name in possible_names:
        normalized_name = normalize_header(possible_name)

        if normalized_name in normalized_lookup:
            return normalized_lookup[normalized_name]

    return None


def get_active_modules(config_df):
    """
    Return modules from CONFIG sheet where Execute = T.
    """
    module_column = config_df.columns[0]

    from_column = find_column(config_df, ["FROM"])
    to_column = find_column(config_df, ["TO"])
    companies_column = find_column(config_df, ["COMPANIES"])
    param1_column = find_column(config_df, ["PARAM1"])
    param2_column = find_column(config_df, ["PARAM2"])
    execute_column = find_column(config_df, ["Execute", "Exetute"])

    if execute_column is None:
        raise ValueError("Could not find Execute column in CONFIG sheet.")

    active_modules = []

    for _, row in config_df.iterrows():
        module_name = str(row.get(module_column, "")).strip()

        if module_name == "" or module_name.lower() == "nan":
            continue

        execute = normalize_execute_value(row.get(execute_column, ""))

        if not execute:
            continue

        active_modules.append(
            {
                "module": module_name,
                "from": format_date(row.get(from_column, "")) if from_column else "",
                "to": format_date(row.get(to_column, "")) if to_column else "",
                "companies": row.get(companies_column, "") if companies_column else "",
                "param1": row.get(param1_column, "") if param1_column else "",
                "param2": row.get(param2_column, "") if param2_column else "",
                "execute": execute,
            }
        )

    return active_modules


def get_active_controls(config_path, module_name):
    """
    Return active controls from a module sheet.
    """
    module_df = read_excel_sheet(config_path, module_name)

    id_control_column = find_column(module_df, ["ID Control"])
    execute_column = find_column(module_df, ["Execute", "Exetute"])
    param1_column = find_column(module_df, ["PARAM1"])
    param2_column = find_column(module_df, ["PARAM2"])
    name_column = find_column(module_df, ["NAME"])
    description_column = find_column(
        module_df,
        [
            "Descripcion | Riks | Action | Procedure",
            "Description | Risk | Action | Procedure",
            "Description",
        ],
    )

    if id_control_column is None:
        raise ValueError(f"Could not find ID Control column in sheet {module_name}.")

    if execute_column is None:
        raise ValueError(f"Could not find Execute column in sheet {module_name}.")

    active_controls = []

    for _, row in module_df.iterrows():
        id_control = str(row.get(id_control_column, "")).strip()

        if id_control == "" or id_control.lower() == "nan":
            continue

        execute = normalize_execute_value(row.get(execute_column, ""))

        if not execute:
            continue

        active_controls.append(
            {
                "module": module_name,
                "id_control": id_control,
                "name": row.get(name_column, "") if name_column else "",
                "param1": row.get(param1_column, "") if param1_column else "",
                "param2": row.get(param2_column, "") if param2_column else "",
                "description": row.get(description_column, "") if description_column else "",
                "execute": execute,
            }
        )

    return active_controls


def read_active_configuration(config_path):
    """
    Read Config.xlsx and return active modules and active controls.
    """
    config_df = read_excel_sheet(config_path, CONFIG_SHEET)

    active_modules = get_active_modules(config_df)

    modules_by_name = {
        module["module"]: module
        for module in active_modules
    }

    active_controls = []

    for module in active_modules:
        module_name = module["module"]
        controls = get_active_controls(config_path, module_name)
        active_controls.extend(controls)

    return {
        "active_modules": active_modules,
        "active_controls": active_controls,
        "modules_by_name": modules_by_name,
    }
