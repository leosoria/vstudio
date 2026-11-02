"""
VM_004 - Vendors sharing the same bank account.

Identifica proveedores distintos que comparten la misma clave bancaria exacta
dentro de una compañía.

Fuentes:
- LBR VM_VENDORS: obligatoria.
- LBR VM_VPBSIK y LBR VM_VPBSAK: obligatorias para última factura.

Granularidad de salida:
    CoCo + Vendor Code + Bank Key

Clave analítica:
    Company + Bank Country + Bank Code + Bank Account

No aplica conversión FX. El importe de última factura permanece en moneda
documental.
"""

from time import perf_counter
from typing import Any

import pandas as pd

from core.vm_common import (
    build_vendor_bank_population,
    build_vendor_master_population,
    build_vm_last_invoice_population,
    get_valid_vendor_population,
    load_vm_vendor_postings,
    load_vm_vendors,
    normalize_company,
    safe_text,
    write_vm_control_sheet,
)


CONTROL_ID = "VM_004"
SHEET_NAME = "VM04"

VENDOR_KEY = [
    "Company",
    "Vendor Code",
]

BANK_KEY_COLUMNS = [
    "Bank Country",
    "Bank Code",
    "Bank Account",
]

OUTPUT_KEY = [
    "CoCo",
    "Vendor Code",
    "Bank Key",
]

OUTPUT_COLUMNS = [
    "CoCo",
    "Vendor Code",
    "Bank Key",
    "Source",
    "Bank Detail",
    "Group",
    "Company",
    "Vendor Name",
    "Last Invoice Number",
    "Last Transaction Date",
    "Last Inv Amt Doc Currency",
    "Last Inv Amt Doc Currency Indicator",
]

def _configured_companies(
    context: dict[str, Any],
) -> set[str]:
    """
    Devuelve los códigos de compañía configurados para VM.

    Un conjunto vacío representa ALL, *, TODAS, TODOS o configuración vacía.
    """
    module_config = context.get("module")

    if not isinstance(module_config, dict):
        raise ValueError(
            f"{CONTROL_ID} requires context['module'] configuration."
        )

    raw_companies = module_config.get("companies", "")

    if raw_companies is None:
        return set()

    if isinstance(raw_companies, (list, tuple, set)):
        raw_values = list(raw_companies)
    else:
        text = safe_text(raw_companies)

        if (
            text == ""
            or text.upper() in {
                "ALL",
                "*",
                "TODAS",
                "TODOS",
            }
        ):
            return set()

        raw_values = (
            text.replace(";", ",")
            .replace("|", ",")
            .split(",")
        )

    companies = {
        normalize_company(value)
        for value in raw_values
        if safe_text(value) != ""
    }

    if companies.intersection(
        {
            "ALL",
            "*",
            "TODAS",
            "TODOS",
        }
    ):
        return set()

    return companies


def _filter_configured_companies(
    vendor_master: pd.DataFrame,
    context: dict[str, Any],
) -> tuple[pd.DataFrame, int]:
    """
    Filtra Company después de la canonicalización realizada por vm_common.
    """
    companies = _configured_companies(context)

    if not companies:
        return vendor_master.copy(), 0

    included = (
        vendor_master["Company"]
        .map(normalize_company)
        .isin(companies)
    )

    excluded_rows = int((~included).sum())

    return (
        vendor_master.loc[included]
        .copy()
        .reset_index(drop=True),
        excluded_rows,
    )


def _timing(label: str, started: float) -> float:
    finished = perf_counter()
    print(f"{CONTROL_ID} {label}: {finished - started:.2f}s")
    return finished


def _require_unique(
    dataframe: pd.DataFrame,
    key: list[str],
    population_name: str,
) -> None:
    duplicated = dataframe.duplicated(key, keep=False)

    if duplicated.any():
        examples = (
            dataframe.loc[duplicated, key]
            .head(20)
            .to_dict("records")
        )
        raise ValueError(
            f"{CONTROL_ID}: {population_name} is not unique by "
            f"{key}. Examples: {examples}"
        )


def build_vm_004(
    valid_vendors: pd.DataFrame,
    vendor_banks: pd.DataFrame,
    last_invoices: pd.DataFrame,
) -> pd.DataFrame:
    """
    Construye las excepciones VM04 sin leer ni escribir archivos.

    Una excepción existe cuando al menos dos Vendor Codes distintos comparten
    una cuenta bancaria exacta dentro de la misma Company.
    """
    required_vendor_columns = {
        "Company",
        "Company Name",
        "Vendor Code",
        "Vendor Name",
    }
    required_bank_columns = {
        "Company",
        "Vendor Code",
        "Bank Country",
        "Bank Code",
        "Bank Account",
    }
    required_invoice_columns = {
        "Company",
        "Vendor Code",
        "Last Invoice Number",
        "Last Transaction Date",
        "Last Inv Amt Doc Currency",
        "Last Inv Amt Doc Currency Indicator",
    }

    for name, dataframe, required in (
        ("valid vendor population", valid_vendors, required_vendor_columns),
        ("vendor bank population", vendor_banks, required_bank_columns),
        ("last-invoice population", last_invoices, required_invoice_columns),
    ):
        missing = sorted(required.difference(dataframe.columns))
        if missing:
            raise ValueError(
                f"{CONTROL_ID}: {name} is missing columns: {missing}"
            )

    _require_unique(
        valid_vendors,
        VENDOR_KEY,
        "valid vendor population",
    )
    _require_unique(
        last_invoices,
        VENDOR_KEY,
        "last-invoice population",
    )
    _require_unique(
        vendor_banks,
        VENDOR_KEY + BANK_KEY_COLUMNS,
        "vendor bank population",
    )

    # Mantiene exclusivamente bancos pertenecientes a la población válida.
    banks = vendor_banks.merge(
        valid_vendors[VENDOR_KEY],
        how="inner",
        on=VENDOR_KEY,
        validate="many_to_one",
    )

    country = (
        banks["Bank Country"]
        .astype("string")
        .fillna("")
        .str.strip()
    )
    bank_code = (
        banks["Bank Code"]
        .astype("string")
        .fillna("")
        .str.strip()
    )
    account = (
        banks["Bank Account"]
        .astype("string")
        .fillna("")
        .str.strip()
    )

    # VM04: una clave bancaria requiere una cuenta no vacía.
    eligible = banks.loc[account.ne("")].copy()

    if eligible.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    eligible["Bank Key"] = (
        country.loc[eligible.index]
        .str.cat(bank_code.loc[eligible.index], sep="|")
        .str.cat(account.loc[eligible.index], sep="|")
    )
    eligible["_Group Key"] = (
        eligible["Company"]
        .astype("string")
        .str.cat(eligible["Bank Key"], sep="\u00a6")
    )

    vendor_count = (
        eligible.groupby(
            "_Group Key",
            sort=False,
            observed=True,
        )["Vendor Code"]
        .transform("nunique")
    )

    exceptions = eligible.loc[vendor_count.ge(2)].copy()

    if exceptions.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    # Numeración estable e independiente del orden de lectura.
    exceptions["Group"] = (
        pd.factorize(
            exceptions["_Group Key"],
            sort=True,
        )[0]
        + 1
    )
    exceptions["Source"] = "LFBK"
    exceptions["Bank Detail"] = (
        exceptions["Bank Account"]
        .astype("string")
        .fillna("")
    )

    display = valid_vendors[
        VENDOR_KEY
        + [
            "Company Name",
            "Vendor Name",
        ]
    ].merge(
        last_invoices[
            VENDOR_KEY
            + [
                "Last Invoice Number",
                "Last Transaction Date",
                "Last Inv Amt Doc Currency",
                "Last Inv Amt Doc Currency Indicator",
            ]
        ],
        how="left",
        on=VENDOR_KEY,
        validate="one_to_one",
    )

    rows_before_enrichment = len(exceptions)

    output = exceptions.merge(
        display,
        how="left",
        on=VENDOR_KEY,
        validate="many_to_one",
    )

    if len(output) != rows_before_enrichment:
        raise AssertionError(
            f"{CONTROL_ID}: display enrichment changed exception rows."
        )

    # Company conserva el código hasta terminar todos los merges.
    output["CoCo"] = (
        output["Company"]
        .astype("string")
        .fillna("")
    )
    output["Company"] = (
        output["Company Name"]
        .astype("string")
        .fillna("")
    )

    output["Last Invoice Number"] = (
        output["Last Invoice Number"]
        .astype("string")
        .fillna("")
    )
    output["Last Transaction Date"] = pd.to_datetime(
        output["Last Transaction Date"],
        errors="coerce",
    )
    output["Last Inv Amt Doc Currency"] = pd.to_numeric(
        output["Last Inv Amt Doc Currency"],
        errors="coerce",
    )
    output["Last Inv Amt Doc Currency Indicator"] = (
        output["Last Inv Amt Doc Currency Indicator"]
        .astype("string")
        .fillna("")
    )

    output = (
        output.sort_values(
            [
                "Group",
                "CoCo",
                "Vendor Code",
                "Bank Key",
            ],
            kind="stable",
        )
        .loc[:, OUTPUT_COLUMNS]
        .reset_index(drop=True)
    )

    # Granularidad final.
    _require_unique(
        output,
        OUTPUT_KEY,
        "VM04 output",
    )

    # Validaciones vectorizadas de Group.
    group_validation = (
        output.groupby(
            "Group",
            observed=True,
        )
        .agg(
            vendors=("Vendor Code", "nunique"),
            companies=("CoCo", "nunique"),
            bank_keys=("Bank Key", "nunique"),
        )
    )

    invalid_group = (
        group_validation["vendors"].lt(2)
        | group_validation["companies"].ne(1)
        | group_validation["bank_keys"].ne(1)
    )

    if invalid_group.any():
        raise AssertionError(
            f"{CONTROL_ID}: invalid Group membership: "
            f"{group_validation.loc[invalid_group].head(20).to_dict('index')}"
        )

    if output["Bank Detail"].eq("").any():
        raise AssertionError(
            f"{CONTROL_ID}: output contains blank bank accounts."
        )

    if list(output.columns) != OUTPUT_COLUMNS:
        raise AssertionError(
            f"{CONTROL_ID}: invalid output columns or order."
        )

    return output


def run_vm_004(context):
    """Execute VM04 through the common configurable matching engine."""
    # VM04_SHARED_MATCHING_V1
    from core.vm_common import run_configured_vm
    return run_configured_vm(context, sheet_name="VM04")
