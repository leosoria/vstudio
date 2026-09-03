"""
VM_001 - Vendors with similar names.

Population:
- Valid LBR vendors loaded through core.vm_common.
- Common VM population exclusions are applied by
  get_valid_vendor_population().
- Comparison is performed within the same CoCo.
- One row per CoCo + Vendor Code.

Exception:
- Clean vendor name has at least 3 characters.
- Names share the same initial.
- Levenshtein distance <= PARAM1; default 2.
- Groups are connected components of matching pairs.

Last-invoice enrichment:
- VPBSIK and VPBSAK are loaded through core.vm_common.
- Only RE and KR documents are considered by the common builder.
- The latest invoice is calculated by
  build_vm_last_invoice_population().
- The enrichment does not change the VM01 population or grouping.

Output:
- output/LBR_Results_VM_YYYYMMDD.xlsx
- Sheet VM01
"""

from time import perf_counter

import numpy as np
import pandas as pd
from rapidfuzz import process
from rapidfuzz.distance import Levenshtein

from core.vm_common import (
    build_vendor_master_population,
    build_vm_last_invoice_population,
    clean_vendor_name,
    get_valid_vendor_population,
    load_vm_vendor_postings,
    load_vm_vendors,
    normalize_company,
    safe_text,
    write_vm_control_sheet,
)


SHEET_NAME = "VM01"


KEY_COLUMNS = [
    "CoCo",
    "Vendor Code",
]


REQUIRED_COLUMNS = [
    "Company",
    "CoCo",
    "Vendor Code",
    "Vendor Name",
]


OUTPUT_COLUMNS = [
    "Company",
    "CoCo",
    "Vendor Code",
    "Vendor Name",
    "Clean Name",
    "Group",
    "Last Invoice Number",
    "Last Transaction Date",
    "Last Inv Amt Doc Currency",
    "Last Inv Amt Doc Currency Indicator",
]


OPTIONAL_INVOICE_COLUMNS = [
    "Last Invoice Number",
    "Last Transaction Date",
    "Last Inv Amt Doc Currency",
    "Last Inv Amt Doc Currency Indicator",
]


class DisjointSet:
    """Union-find over unique cleaned names."""

    def __init__(self, size):
        self.parent = np.arange(
            size,
            dtype=np.int64,
        )

        self.rank = np.zeros(
            size,
            dtype=np.int8,
        )

    def find(self, item):
        parent = self.parent

        while parent[item] != item:
            parent[item] = parent[
                parent[item]
            ]

            item = parent[item]

        return item

    def union(
        self,
        left,
        right,
    ):
        left_root = self.find(
            left
        )

        right_root = self.find(
            right
        )

        if left_root == right_root:
            return

        if (
            self.rank[left_root]
            < self.rank[right_root]
        ):
            left_root, right_root = (
                right_root,
                left_root,
            )

        self.parent[right_root] = (
            left_root
        )

        if (
            self.rank[left_root]
            == self.rank[right_root]
        ):
            self.rank[left_root] += 1


def get_max_distance(context):
    """Return PARAM1 Levenshtein distance, defaulting to 2."""
    value = context[
        "control"
    ].get(
        "param1",
        "",
    )

    if value is None or pd.isna(value):
        return 2

    value_text = str(
        value
    ).strip()

    if (
        value_text == ""
        or value_text.casefold()
        in {
            "nan",
            "none",
            "<na>",
        }
    ):
        return 2

    try:
        numeric_value = float(
            value_text
        )
    except (
        TypeError,
        ValueError,
    ) as error:
        raise ValueError(
            "VM_001 PARAM1 must be a whole number zero or greater. "
            f"Received: {value!r}."
        ) from error

    if not numeric_value.is_integer():
        raise ValueError(
            "VM_001 PARAM1 must be a whole number zero or greater. "
            f"Received: {value!r}."
        )

    distance = int(
        numeric_value
    )

    if distance < 0:
        raise ValueError(
            "VM_001 PARAM1 must be zero or greater. "
            f"Received: {value!r}."
        )

    return distance


def validate_population(vendors):
    """Validate the canonical one-row-per-vendor VM01 population."""
    missing_columns = [
        column
        for column in REQUIRED_COLUMNS
        if column not in vendors.columns
    ]

    if missing_columns:
        raise ValueError(
            "VM_001 missing required canonical columns from "
            "core.vm_common: "
            + ", ".join(
                missing_columns
            )
        )

    duplicate_key = vendors.duplicated(
        subset=KEY_COLUMNS,
        keep=False,
    )

    if duplicate_key.any():
        sample = (
            vendors.loc[
                duplicate_key,
                KEY_COLUMNS,
            ]
            .drop_duplicates()
            .head(10)
            .to_dict("records")
        )

        raise ValueError(
            "VM_001 requires one row per CoCo + Vendor Code. "
            "The common vendor population contains duplicate keys. "
            f"Sample duplicate keys: {sample}"
        )


def add_clean_names(vendors):
    """Normalize each distinct vendor name only once."""
    unique_names = (
        vendors[
            "Vendor Name"
        ]
        .drop_duplicates()
    )

    clean_lookup = pd.Series(
        unique_names.map(
            clean_vendor_name
        ).to_numpy(),
        index=unique_names,
    )

    result = vendors.copy()

    result[
        "Clean Name"
    ] = (
        result[
            "Vendor Name"
        ]
        .map(
            clean_lookup
        )
        .fillna("")
    )

    return result


def assign_similarity_groups(
    vendors,
    max_distance,
):
    """Assign fuzzy-name connected components within CoCo and initial.

    This is the existing VM01 fuzzy implementation. It compares unique names,
    avoids a dense N x N distance matrix and preserves the current connected
    component and Group-numbering behavior.
    """
    result = (
        vendors.reset_index(
            drop=True
        )
        .copy()
    )

    result[
        "_input_order"
    ] = np.arange(
        len(result),
        dtype=np.int64,
    )

    nodes = (
        result.loc[
            result[
                "Clean Name"
            ].str.len().ge(3),
            [
                "CoCo",
                "Clean Name",
            ],
        ]
        .drop_duplicates()
        .reset_index(
            drop=True
        )
    )

    result["Group"] = 0

    if nodes.empty:
        return result.drop(
            columns=[
                "_input_order",
            ]
        )

    nodes[
        "_initial"
    ] = nodes[
        "Clean Name"
    ].str[0]

    nodes[
        "_node"
    ] = np.arange(
        len(nodes),
        dtype=np.int64,
    )

    node_index = (
        pd.MultiIndex.from_frame(
            nodes[
                [
                    "CoCo",
                    "Clean Name",
                ]
            ]
        )
    )

    row_index = (
        pd.MultiIndex.from_frame(
            result[
                [
                    "CoCo",
                    "Clean Name",
                ]
            ]
        )
    )

    result[
        "_node"
    ] = node_index.get_indexer(
        row_index
    )

    node_counts = np.bincount(
        result.loc[
            result["_node"].ge(0),
            "_node",
        ].to_numpy(
            dtype=np.int64
        ),
        minlength=len(nodes),
    )

    components = DisjointSet(
        len(nodes)
    )

    for _, block in nodes.groupby(
        [
            "CoCo",
            "_initial",
        ],
        sort=False,
        observed=True,
    ):
        node_ids = block[
            "_node"
        ].to_numpy(
            dtype=np.int64
        )

        names = block[
            "Clean Name"
        ].tolist()

        if len(names) < 2:
            continue

        for local_left, name in enumerate(
            names
        ):
            matches = process.extract(
                name,
                names,
                scorer=Levenshtein.distance,
                score_cutoff=max_distance,
                score_hint=max_distance,
                limit=None,
            )

            for (
                _,
                _,
                local_right,
            ) in matches:
                if local_right <= local_left:
                    continue

                components.union(
                    int(
                        node_ids[
                            local_left
                        ]
                    ),
                    int(
                        node_ids[
                            local_right
                        ]
                    ),
                )

    roots = np.fromiter(
        (
            components.find(
                node
            )
            for node in range(
                len(nodes)
            )
        ),
        dtype=np.int64,
        count=len(nodes),
    )

    population_by_root = (
        pd.Series(
            node_counts
        )
        .groupby(
            roots,
            sort=False,
        )
        .sum()
    )

    valid_roots = (
        population_by_root[
            population_by_root.ge(2)
        ].index
    )

    nodes[
        "_root"
    ] = roots

    node_to_root = (
        nodes.set_index(
            "_node"
        )[
            "_root"
        ]
    )

    result[
        "_root"
    ] = result[
        "_node"
    ].map(
        node_to_root
    )

    matched = result[
        "_root"
    ].isin(
        valid_roots
    )

    root_order = (
        result.loc[
            matched
        ]
        .groupby(
            "_root",
            sort=False,
        )[
            "_input_order"
        ]
        .min()
        .sort_values(
            kind="stable"
        )
    )

    group_by_root = pd.Series(
        np.arange(
            1,
            len(root_order) + 1,
            dtype=np.int64,
        ),
        index=root_order.index,
    )

    result.loc[
        matched,
        "Group",
    ] = (
        result.loc[
            matched,
            "_root",
        ]
        .map(
            group_by_root
        )
        .astype(
            np.int64
        )
    )

    return result.drop(
        columns=[
            "_input_order",
            "_node",
            "_root",
        ]
    )


def build_vm_001(
    vendors,
    max_distance,
):
    """Build the VM01 exception output from an enriched vendor population."""
    validate_population(
        vendors
    )

    prepared = add_clean_names(
        vendors
    )

    grouped = assign_similarity_groups(
        prepared,
        max_distance,
    )

    output = (
        grouped.loc[
            grouped["Group"].gt(0)
        ]
        .sort_values(
            [
                "Group",
                "CoCo",
                "Vendor Code",
            ],
            kind="stable",
        )
        .copy()
    )

    for column in OPTIONAL_INVOICE_COLUMNS:
        if column not in output.columns:
            output[
                column
            ] = pd.NA

    return (
        output[
            OUTPUT_COLUMNS
        ]
        .reset_index(
            drop=True
        )
    )


def _filter_configured_companies(
    vendor_master,
    context,
):
    """Filter the vendor master using VM CONFIG COMPANIES."""
    configured_value = safe_text(
        context[
            "module"
        ].get(
            "companies",
            "",
        )
    )

    if (
        configured_value == ""
        or configured_value.upper()
        in {
            "ALL",
            "TODAS",
            "TODOS",
        }
    ):
        return (
            vendor_master.copy().reset_index(
                drop=True
            ),
            0,
        )

    normalized_value = (
        configured_value
    )

    for separator in (
        ";",
        "|",
        "\n",
        "\r",
        "\t",
    ):
        normalized_value = (
            normalized_value.replace(
                separator,
                ",",
            )
        )

    if (
        "," not in normalized_value
        and " " in normalized_value
    ):
        normalized_value = ",".join(
            normalized_value.split()
        )

    configured_companies = {
        normalize_company(
            item
        )
        for item in normalized_value.split(
            ","
        )
        if safe_text(item) != ""
    }

    included = (
        vendor_master[
            "Company"
        ]
        .map(
            normalize_company
        )
        .isin(
            configured_companies
        )
    )

    excluded_rows = int(
        (~included).sum()
    )

    return (
        vendor_master.loc[
            included
        ]
        .copy()
        .reset_index(
            drop=True
        ),
        excluded_rows,
    )


def _add_vm01_company_display_columns(
    vendors,
):
    """Convert the common Company key to the official VM01 display columns.

    Before this function:
        Company = normalized company code used by common merges.

    After this function:
        Company = descriptive company name.
        CoCo = normalized company code.
    """
    if "Company" not in vendors.columns:
        raise ValueError(
            "VM_001 requires the canonical 'Company' company-code column."
        )

    if "Company Name" not in vendors.columns:
        raise ValueError(
            "VM_001 requires the canonical 'Company Name' descriptive "
            "column from build_vendor_master_population()."
        )

    result = vendors.copy()

    result[
        "CoCo"
    ] = (
        result[
            "Company"
        ]
        .astype("string")
        .fillna("")
    )

    result[
        "Company"
    ] = (
        result[
            "Company Name"
        ]
        .astype("string")
        .fillna("")
    )

    return result


def run_vm_001(context):
    """Run VM01 and write the result sheet."""
    max_distance = get_max_distance(
        context
    )

    # Load VM VENDORS, VPBSIK and VPBSAK once.
    load_start = perf_counter()

    vendor_source = load_vm_vendors(
        context
    )

    postings, posting_metadata = (
        load_vm_vendor_postings(
            context
        )
    )

    load_seconds = (
        perf_counter()
        - load_start
    )

    if (
        postings is None
        or not posting_metadata.get(
            "available",
            False,
        )
    ):
        raise FileNotFoundError(
            "VM_001 requires VPBSIK and VPBSAK to populate "
            "the last-invoice display columns."
        )

    # Build the approved common population and last-invoice enrichment.
    preparation_start = perf_counter()

    vendor_master = (
        build_vendor_master_population(
            vendor_source
        )
    )

    vendor_master, excluded_company = (
        _filter_configured_companies(
            vendor_master,
            context,
        )
    )

    vendors, population_metrics = (
        get_valid_vendor_population(
            vendor_master
        )
    )

    if vendors.empty:
        raise ValueError(
            "VM_001: valid vendor population is empty after "
            "CONFIG company and common VM exclusion rules."
        )

    last_invoices = (
        build_vm_last_invoice_population(
            postings
        )
    )

    rows_before_invoice_merge = len(
        vendors
    )

    vendors = vendors.merge(
        last_invoices,
        how="left",
        on=[
            "Company",
            "Vendor Code",
        ],
        validate="one_to_one",
    )

    if len(vendors) != rows_before_invoice_merge:
        raise AssertionError(
            "VM01 last-invoice enrichment changed vendor population."
        )

    for column in (
        "Last Invoice Number",
        "Last Transaction Date",
        "Last Inv Amt Doc Currency Indicator",
    ):
        if column not in vendors.columns:
            vendors[
                column
            ] = ""

        vendors[
            column
        ] = (
            vendors[
                column
            ]
            .astype("string")
            .fillna("")
        )

    if (
        "Last Inv Amt Doc Currency"
        not in vendors.columns
    ):
        vendors[
            "Last Inv Amt Doc Currency"
        ] = pd.NA

    vendors[
        "Last Inv Amt Doc Currency"
    ] = pd.to_numeric(
        vendors[
            "Last Inv Amt Doc Currency"
        ],
        errors="coerce",
    )

    # The common merge key is Company. VM01 displays that code as CoCo and
    # places the descriptive company name in Company.
    vendors = (
        _add_vm01_company_display_columns(
            vendors
        )
    )

    validate_population(
        vendors
    )

    prepared = add_clean_names(
        vendors
    )

    preparation_seconds = (
        perf_counter()
        - preparation_start
    )

    # Existing fuzzy logic remains unchanged.
    analytic_start = perf_counter()

    output = assign_similarity_groups(
        prepared,
        max_distance,
    )

    output = (
        output.loc[
            output[
                "Group"
            ].gt(0)
        ]
        .sort_values(
            [
                "Group",
                "CoCo",
                "Vendor Code",
            ],
            kind="stable",
        )
        .copy()
    )

    for column in OPTIONAL_INVOICE_COLUMNS:
        if column not in output.columns:
            output[
                column
            ] = pd.NA

    output = (
        output[
            OUTPUT_COLUMNS
        ]
        .reset_index(
            drop=True
        )
    )

    analytic_seconds = (
        perf_counter()
        - analytic_start
    )

    # VM01 does not perform FX conversion.
    fx_seconds = 0.0

    posting_rows = len(
        postings
    )

    invoice_posting_rows = int(
        postings[
            "Document Type"
        ]
        .astype("string")
        .str.strip()
        .str.upper()
        .isin(
            {
                "RE",
                "KR",
            }
        )
        .sum()
    )

    vendors_with_last_invoice = len(
        last_invoices
    )

    audit = {
        "source_rows": len(
            vendor_source
        ),
        "master_rows": len(
            vendor_master
        ),
        "valid_vendor_rows": len(
            vendors
        ),
        "excluded_company": (
            excluded_company
        ),
        "posting_rows": (
            posting_rows
        ),
        "invoice_posting_rows": (
            invoice_posting_rows
        ),
        "vendors_with_last_invoice": (
            vendors_with_last_invoice
        ),
        **population_metrics,
    }

    write_start = perf_counter()

    output_file = write_vm_control_sheet(
        context=context,
        sheet_name=SHEET_NAME,
        dataframe=output,
        date_columns={
            "Last Transaction Date",
        },
        amount_columns={
            "Last Inv Amt Doc Currency",
        },
    )

    write_seconds = (
        perf_counter()
        - write_start
    )

    print(
        "VM_001 timings | "
        f"load={load_seconds:.2f}s | "
        f"prepare={preparation_seconds:.2f}s | "
        f"analytic={analytic_seconds:.2f}s | "
        f"fx={fx_seconds:.2f}s (not applicable) | "
        f"write={write_seconds:.2f}s"
    )

    print(
        "VM_001 population | "
        f"loaded={len(vendors)} | "
        f"exceptions={len(output)} | "
        f"groups={output['Group'].nunique()}"
    )

    print(
        f"VM_001 posting rows: "
        f"{posting_rows}"
    )

    print(
        f"VM_001 RE/KR invoice posting rows: "
        f"{invoice_posting_rows}"
    )

    print(
        f"VM_001 vendors with last invoice: "
        f"{vendors_with_last_invoice}"
    )

    print(
        f"VM_001 audit metrics: "
        f"{audit}"
    )

    print(
        f"VM_001 output: "
        f"{output_file}"
    )
