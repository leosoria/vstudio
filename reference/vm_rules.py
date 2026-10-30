"""Shared configuration and filtering helpers for VM controls.

Active status follows the agreed onsite criteria:
- exclude central or company deletion flags;
- exclude SPERR, SPERM or SPERZ equal to X.

Control-specific population exclusions remain separate.
Importing this module does not change any control automatically.
"""

import re
import unicodedata

import pandas as pd


# VM_REVIEW_GROUPS_ADDRESS_V1
RULE_HEADERS = (
    "Exclude Same Address",
    "Same Company Code",
    "Different Vendor Code",
    "Same Address",
    "Different Vendor Name",
    "Ignore Name Punctuation",
    "Same Phone Number",
    "Same Bank Account Number",
    "Same Tax Registration Number",
    "Same Email",
)

MATCH_HEADERS = (
    "Same Address",
    "Same Phone Number",
    "Same Bank Account Number",
    "Same Tax Registration Number",
    "Same Email",
)


def _text(value):
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def normalize_rule_header(value):
    """Match headers regardless of case or repeated whitespace."""
    return " ".join(_text(value).casefold().split())


def extract_vm_rule_values(dataframe, row):
    """Read recognized headers by name, independently of column position.

    Missing columns remain absent, so the analytical validator can report
    them explicitly instead of silently treating them as false.
    """
    recognized = {
        normalize_rule_header(header): header
        for header in RULE_HEADERS
    }
    result = {}

    for column in dataframe.columns:
        canonical = recognized.get(normalize_rule_header(column))
        if canonical is None:
            continue
        if canonical in result:
            raise ValueError(
                f"VM: duplicate rule header: {canonical!r}."
            )
        result[canonical] = row[column]

    return result


def parse_vm_rules(context):
    """Validate the T/F settings for a configurable matching control.

    Blank values and unknown settings raise an error.
    F means that the condition is not required.
    """
    control = context.get("control", {})
    control_id = control.get("id_control", "VM")
    raw = control.get("vm_rules_raw")

    if not isinstance(raw, dict):
        raise ValueError(
            f"{control_id}: VM rules are unavailable. "
            "Check config_reader.py and config.xlsx."
        )

    unknown = sorted(set(raw).difference(RULE_HEADERS))
    missing = [header for header in RULE_HEADERS if header not in raw]

    if missing or unknown:
        raise ValueError(
            f"{control_id}: missing rule headers: {missing}; "
            f"unknown rule headers: {unknown}."
        )

    rules = {}
    for header in RULE_HEADERS:
        value = _text(raw[header]).upper()
        if value not in {"T", "F"}:
            raise ValueError(
                f"{control_id}: {header!r} must contain T or F; "
                f"received {raw[header]!r}."
            )
        rules[header] = value == "T"

    if rules["Same Address"] and rules["Exclude Same Address"]:
        raise ValueError(f"{control_id}: Same Address and Exclude Same Address cannot both be T.")

    if not any(rules[header] for header in MATCH_HEADERS):
        raise ValueError(
            f"{control_id}: enable at least one matching condition "
            "(address, phone, bank account, tax number or email)."
        )

    return rules


def filter_active_records(dataframe):
    """Return active records and exclusion counts.

    Works for suppliers and F-code employee records. Does not remove
    employees, intercompanies or vendor prefixes.
    """
    deletion_fields = (
        "Central Deletion Flag",
        "Company Deletion Flag",
    )
    block_fields = (
        "Central Posting Block",
        "Purchasing Block",
        "Central Payment Block",
    )
    required = set(deletion_fields + block_fields)
    missing = sorted(required.difference(dataframe.columns))

    if missing:
        raise ValueError(
            f"VM active filter: missing status fields: {missing}."
        )

    excluded = pd.Series(False, index=dataframe.index)
    metrics = {"input_rows": len(dataframe)}

    for field in deletion_fields + block_fields:
        values = dataframe[field].map(_text).str.upper()
        flagged = (
            values.ne("")
            if field in deletion_fields
            else values.eq("X")
        )
        metrics[field] = int(flagged.sum())
        excluded |= flagged

    result = dataframe.loc[~excluded].copy().reset_index(drop=True)
    metrics["excluded_rows"] = int(excluded.sum())
    metrics["output_rows"] = len(result)

    return result, metrics


def normalize_comparison_name(value, ignore_punctuation=True):
    """Normalize case, accents and whitespace for name comparisons."""
    text = unicodedata.normalize("NFKD", _text(value).upper())
    text = "".join(
        character
        for character in text
        if not unicodedata.combining(character)
    )

    if ignore_punctuation:
        text = "".join(
            character if character.isalnum() or character.isspace() else " "
            for character in text
        )

    return re.sub(r"\s+", " ", text).strip()


def filter_candidate_pairs(pairs, rules, columns):
    """Apply enabled conditions to prepared candidate pairs.

    columns maps a rule header to (left column, right column).
    Address, phone, bank, tax and email keys must be prepared by the
    control before calling this function.

    Matching conditions require both values to be populated.
    Different Vendor Name also requires both names to be populated.
    All enabled conditions are combined with AND.
    """
    condition_headers = [
        header for header in RULE_HEADERS
        if header != "Ignore Name Punctuation" and rules[header]
    ]

    keep = pd.Series(True, index=pairs.index)

    for header in condition_headers:
        if header == "Exclude Same Address" and header not in columns:
            columns = dict(columns)
            if "Same Address" in columns:
                columns[header] = columns["Same Address"]
        if header not in columns:
            raise ValueError(
                f"VM: enabled condition {header!r} is not supported "
                "by this control yet."
            )

        left_column, right_column = columns[header]
        missing = [
            column for column in (left_column, right_column)
            if column not in pairs.columns
        ]
        if missing:
            raise ValueError(
                f"VM: missing comparison columns for {header!r}: {missing}."
            )

        if header == "Different Vendor Name":
            ignore = rules["Ignore Name Punctuation"]
            left = pairs[left_column].map(
                lambda value: normalize_comparison_name(value, ignore)
            )
            right = pairs[right_column].map(
                lambda value: normalize_comparison_name(value, ignore)
            )
        else:
            left = pairs[left_column].map(_text)
            right = pairs[right_column].map(_text)

        populated = left.ne("") & right.ne("")
        different = header in {
            "Different Vendor Code",
            "Different Vendor Name",
        }
        comparison = left.ne(right) if different else left.eq(right)
        if header == "Exclude Same Address":
            keep &= ~(populated & left.eq(right))
        else:
            keep &= populated & comparison

    return pairs.loc[keep].copy().reset_index(drop=True)


# VM_RULES_ENGINE_V2
# VM_BANK_EXCLUSIONS_V1
VM_BANK_ACCOUNT_EXCLUSIONS = frozenset({
    "BOLETO",
    "INTERNACIONAL",
    "CARTAO",
    "IMPOSTO",
})


# VM_TAX_EXCLUSIONS_V1
VM_TAX_EXCLUDED_VALUES = frozenset({"DISPENSADO"})


def find_vendor_pairs(
    population, rules, *, columns=None, max_pairs=1000000
):
    """Apply every enabled condition to the same pair of vendors.

    Input: canonical rows preserving multiple bank accounts and emails.
    Output: one row per pair of Company + Vendor Code identities.
    All enabled matching conditions are combined with AND.
    """
    from core.vm_common import (
        normalize_company,
        normalize_vendor_code,
    )

    specs = {
        "Same Address": ("Street", "City", "ZipCode", "Country"),
        "Same Phone Number": ("Phone1",),
        "Same Bank Account Number": ("Bank Account",),
        "Same Tax Registration Number": ("Tax/Business Number",),
        "Same Email": ("Email",),
    }

    if columns:
        unknown = set(columns).difference(specs)
        if unknown:
            raise ValueError(
                f"VM: unknown comparison mappings: {sorted(unknown)}"
            )
        specs.update(columns)

    if set(rules) != set(RULE_HEADERS) or any(
        type(value) is not bool for value in rules.values()
    ):
        raise ValueError(
            "VM: use parse_vm_rules(context) before matching."
        )

    enabled = [header for header in MATCH_HEADERS if rules[header]]

    if not enabled:
        raise ValueError(
            "VM: enable at least one matching condition."
        )

    identity = ["Company", "Vendor Code"]
    required = set(identity + ["Vendor Name"])

    for header in enabled:
        if (
            not isinstance(specs[header], (tuple, list))
            or not specs[header]
        ):
            raise ValueError(
                f"VM: invalid column mapping for {header}."
            )
        required.update(specs[header])

    if rules["Exclude Same Address"]:
        required.update(specs["Same Address"])

    missing = sorted(required.difference(population.columns))
    if missing:
        raise ValueError(
            f"VM: missing data for enabled conditions: {missing}."
        )

    data = population.copy()
    data["Company"] = data["Company"].map(normalize_company)
    data["Vendor Code"] = data["Vendor Code"].map(
        normalize_vendor_code
    )

    data = data.loc[
        data["Company"].ne("")
        & data["Vendor Code"].ne("")
    ].copy()

    status_fields = [
        "Central Deletion Flag",
        "Company Deletion Flag",
        "Central Posting Block",
        "Purchasing Block",
        "Central Payment Block",
    ]

    missing_status = sorted(
        set(status_fields).difference(data.columns)
    )
    if missing_status:
        raise ValueError(
            f"VM: missing status fields: {missing_status}."
        )

    # Avoid accepting an active row that hides a blocked row for
    # the same company/vendor identity.
    checked = data[
        identity + ["Vendor Name"] + status_fields
    ].copy()

    for field in ["Vendor Name"] + status_fields:
        checked[field] = checked[field].map(_text).str.upper()

    conflicts = checked.groupby(identity)[
        ["Vendor Name"] + status_fields
    ].nunique()

    if conflicts.gt(1).any().any():
        raise ValueError(
            "VM: conflicting names/statuses for a company/vendor."
        )

    data, metrics = filter_active_records(data)

    entities = (
        data[identity + ["Vendor Name"]]
        .drop_duplicates(identity)
        .sort_values(identity)
        .reset_index(drop=True)
    )
    entities["_Entity"] = entities.index
    entities["_Name"] = entities["Vendor Name"].map(
        lambda value: normalize_comparison_name(
            value, rules["Ignore Name Punctuation"]
        )
    )

    data = data.merge(
        entities[identity + ["_Entity"]],
        on=identity,
        how="inner",
        validate="many_to_one",
    )

    out_columns = [
        "Company Left",
        "Vendor Code Left",
        "Vendor Name Left",
        "Company Right",
        "Vendor Code Right",
        "Vendor Name Right",
        *[
            "Matching " + header.removeprefix("Same ")
            for header in enabled
        ],
    ]

    if len(entities) < 2:
        return pd.DataFrame(columns=out_columns), metrics

    def normal(value, header):
        text = _text(value)

        if header == "Same Email":
            return text.casefold()

        if header == "Same Phone Number":
            return re.sub(r"\D+", "", text)

        text = normalize_comparison_name(text, True)
        return "".join(
            character
            for character in text
            if character.isalnum()
        )

    maps = {}
    estimates = {}

    for header in enabled:
        fields = list(specs[header])
        parts = [
            data[field].map(
                lambda value: normal(value, header)
            )
            for field in fields
        ]
        keys = list(zip(*(part.tolist() for part in parts)))

        mapping = data[["_Entity"]].copy()
        mapping["_Key"] = keys

        populated = [
            any(key)
            if header == "Same Address"
            else all(key)
            for key in keys
        ]

        if header == "Same Bank Account Number":
            excluded_account = [
                any(
                    part in VM_BANK_ACCOUNT_EXCLUSIONS
                    or (bool(part) and set(part) == {"0"})
                    for part in key
                )
                for key in keys
            ]

            metrics["excluded_bank_source_rows"] = sum(
                excluded_account
            )

            populated = [
                valid and not excluded
                for valid, excluded in zip(
                    populated, excluded_account
                )
            ]

        if header == "Same Tax Registration Number":
            excluded_tax = [
                any(
                    part in VM_TAX_EXCLUDED_VALUES
                    or (
                        bool(part)
                        and all(character == "0" for character in part)
                    )
                    for part in key
                )
                for key in keys
            ]

            metrics["excluded_tax_source_rows"] = sum(excluded_tax)

            populated = [
                valid and not excluded
                for valid, excluded in zip(populated, excluded_tax)
            ]

        mapping = mapping.loc[populated].drop_duplicates()

        if rules["Same Company Code"]:
            mapping = mapping.merge(
                entities[["_Entity", "Company"]],
                on="_Entity",
            )
            mapping["_Key"] = list(
                zip(mapping["Company"], mapping["_Key"])
            )
            mapping = mapping[["_Entity", "_Key"]]

        maps[header] = mapping
        sizes = mapping.groupby("_Key").size()
        estimates[header] = int(
            ((sizes * (sizes - 1)) // 2).sum()
        )

    # Start with the most selective matching condition.
    ordered = sorted(
        enabled, key=lambda header: estimates[header]
    )
    seed = ordered[0]

    if estimates[seed] == 0:
        return pd.DataFrame(columns=out_columns), metrics

    if estimates[seed] > max_pairs:
        raise ValueError(
            f"VM: estimated candidate pairs exceed {max_pairs:,}; "
            "review shared placeholder values or add "
            "a matching condition."
        )

    seed_map = maps[seed]
    pairs = seed_map.merge(
        seed_map,
        on="_Key",
        suffixes=(" Left", " Right"),
    )

    # One direction per pair; never match an identity with itself.
    pairs = pairs.loc[
        pairs["_Entity Left"].lt(pairs["_Entity Right"])
    ].copy()

    pair_ids = ["_Entity Left", "_Entity Right"]
    pairs = pairs[pair_ids].drop_duplicates()

    for side in ("Left", "Right"):
        info = entities.rename(
            columns={
                column: column + " " + side
                for column in entities.columns
            }
        )
        pairs = pairs.merge(
            info,
            on="_Entity " + side,
            validate="many_to_one",
        )

    if rules["Different Vendor Code"]:
        pairs = pairs.loc[
            pairs["Vendor Code Left"].ne(
                pairs["Vendor Code Right"]
            )
        ]

    if rules["Different Vendor Name"]:
        pairs = pairs.loc[
            pairs["_Name Left"].ne("")
            & pairs["_Name Right"].ne("")
            & pairs["_Name Left"].ne(pairs["_Name Right"])
        ]

    if rules["Exclude Same Address"]:
        fields = list(specs["Same Address"])
        addresses = {}
        for values in data[["_Entity"] + fields].itertuples(index=False, name=None):
            key = tuple(normal(value, "Same Address") for value in values[1:])
            # Country/city alone is not an actual street address.
            if key and key[0]:
                addresses.setdefault(values[0], set()).add(key)
        shared = [
            bool(addresses.get(left, set()) & addresses.get(right, set()))
            for left, right in pairs[pair_ids].itertuples(index=False, name=None)
        ]
        metrics["pairs_excluded_same_address"] = sum(shared)
        pairs = pairs.loc[[not value for value in shared]].copy()

    for header in ordered:
        mapping = maps[header]
        left = mapping.rename(
            columns={"_Entity": "_Entity Left"}
        )
        right = mapping.rename(
            columns={"_Entity": "_Entity Right"}
        )

        matched = pairs[pair_ids].merge(
            left, on="_Entity Left"
        )
        matched = matched.merge(
            right, on=["_Entity Right", "_Key"]
        )

        label = "Matching " + header.removeprefix("Same ")

        if matched.empty:
            return pd.DataFrame(columns=out_columns), metrics

        def display(key):
            if rules["Same Company Code"]:
                key = key[1]
            return " | ".join(key)

        matched[label] = matched["_Key"].map(display)

        evidence = matched.groupby(
            pair_ids, as_index=False
        )[label].agg(
            lambda values: "; ".join(sorted(set(values)))
        )

        pairs = pairs.merge(
            evidence,
            on=pair_ids,
            validate="one_to_one",
        )

    metrics["matching_pairs"] = len(pairs)

    return (
        pairs.loc[:, out_columns].reset_index(drop=True),
        metrics,
    )
