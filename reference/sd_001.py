
"""SD_001: evaluate complete invoice journals and report every position.

LFA1 general creation is the conflict basis.
LFB1 extension is supporting information.
PARAM1: invoice document types; default KR,RE.
Scope uses AP Entry Date.
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

from core.sd_common import (
    load_vendor_master,
    load_ap_items,
    load_sd_journals,
    normalize_text,
    write_control_sheet,
)

DOC = ["Company", "Fiscal Year", "Document Number"]
KEY = DOC + ["Document Item"]


def get_invoice_types(context):
    text = normalize_text(
        context.get("control", {}).get("param1", "")
    ).upper()

    if not text:
        return {"KR", "RE"}

    types = {
        value
        for value in re.split(r"[,;|\s]+", text)
        if value
    }

    if any(
        not re.fullmatch(r"[A-Z0-9]{2}", value)
        for value in types
    ):
        raise ValueError(
            "SD_001 PARAM1: expected document types, e.g. KR,RE"
        )

    return types


def prepare_ap_items(items):
    """Retained for compatibility with the earlier diagnostic."""
    business = [
        column for column in items.columns
        if column not in {"Source File", "Source Row"}
    ]

    distinct = items.drop_duplicates(business)

    if distinct.duplicated(KEY, keep=False).any():
        raise ValueError(
            "SD_001: conflicting AP records for the same item"
        )

    prepared = items.drop_duplicates(KEY).copy()
    return prepared, len(items) - len(prepared)


def prepare_vendor_master(vendors):
    fields = [
        "Company",
        "Vendor Code",
        "Vendor Name",
        "Vendor Created On",
        "Vendor Created By",
        "Company Vendor Created On",
        "Company Vendor Created By",
    ]

    distinct = vendors[fields].drop_duplicates()

    if distinct.duplicated(["Company", "Vendor Code"]).any():
        raise ValueError(
            "SD_001: conflicting creation evidence in vendor master"
        )

    master = vendors[
        fields + ["Source File", "Source Row"]
    ].drop_duplicates(["Company", "Vendor Code"]).copy()

    return master.rename(columns={
        "Source File": "Vendor Source File",
        "Source Row": "Vendor Source Row",
    })


def analyze_full_journals(vendors, ap, journals, types):
    header_fields = DOC + [
        "Document Type",
        "Entry Date",
        "User ID",
        "Transaction Code",
        "Document Currency",
        "Reversal Document",
        "Reversal Year",
    ]

    for optional in [
        "Document Date",
        "Posting Date",
        "Reference",
    ]:
        if optional in ap:
            header_fields.append(optional)

    headers = ap[header_fields].drop_duplicates()

    if headers.duplicated(DOC).any():
        raise ValueError(
            "SD_001: inconsistent AP header values "
            "for the same document"
        )

    selected = headers.loc[
        headers["Document Type"].isin(types)
    ].copy()

    reversed_keys = set()
    linked = ap.loc[ap["Reversal Document"].ne("")]

    reversed_keys.update(
        linked[DOC].itertuples(index=False, name=None)
    )

    for _, row in linked.iterrows():
        if row["Reversal Year"]:
            reversed_keys.add((
                row["Company"],
                row["Reversal Year"],
                row["Reversal Document"],
            ))

    reversed_mask = pd.Series(
        [
            key in reversed_keys
            for key in selected[DOC].itertuples(
                index=False,
                name=None,
            )
        ],
        index=selected.index,
        dtype=bool,
    )

    reversed_mask |= selected["Transaction Code"].isin({
        "MR8M",
        "FB08",
        "F.80",
    })

    selected = selected.loc[~reversed_mask].copy()

    ap_lines = ap.merge(
        selected[DOC],
        on=DOC,
        how="inner",
        validate="many_to_one",
    )

    lines = journals.merge(
        selected,
        on=DOC,
        how="inner",
        validate="many_to_one",
    )

    coverage = selected[DOC].merge(
        lines[DOC].drop_duplicates(),
        on=DOC,
        how="left",
        indicator=True,
    )

    missing = coverage.loc[
        coverage["_merge"].eq("left_only"),
        DOC,
    ]

    if not missing.empty:
        raise ValueError(
            f"SD_001: {len(missing)} scoped journals missing "
            "from SD JOURNAL input. "
            f"Examples: {missing.head(5).to_dict('records')}"
        )

    comparison = ap_lines.merge(
        lines,
        on=KEY,
        how="left",
        suffixes=("_AP", "_BSEG"),
        indicator=True,
        validate="many_to_one",
    )

    bad = comparison["_merge"].ne("both")

    for name in [
        "Vendor Code",
        "Posting Key",
        "Debit Credit",
    ]:
        bad |= comparison[name + "_AP"].ne(
            comparison[name + "_BSEG"]
        )

    for name in [
        "Amount in Document Currency",
        "Amount in Local Currency",
    ]:
        bad |= (
            comparison[name + "_AP"]
            - comparison[name + "_BSEG"]
        ).abs().gt(0.01)

    if bad.any():
        examples = comparison.loc[bad, KEY].head(5)

        raise ValueError(
            "SD_001: AP/BSEG positions missing or inconsistent: "
            f"{examples.to_dict('records')}"
        )

    if "BSEG Document Currency" in lines:
        mismatch = (
            lines["BSEG Document Currency"].ne("")
            & lines["BSEG Document Currency"].ne(
                lines["Document Currency"]
            )
        )

        if mismatch.any():
            raise ValueError(
                "SD_001: BSEG/AP document currencies disagree"
            )

    sign = lines["Debit Credit"].map({
        "S": 1,
        "H": -1,
    })

    check = lines[DOC].copy()

    check["LC"] = (
        lines["Amount in Local Currency"] * sign
    )
    check["DC"] = (
        lines["Amount in Document Currency"] * sign
    )

    balances = check.groupby(DOC)[["LC", "DC"]].sum()

    bad_balances = balances.loc[
        balances.abs().gt(0.01).any(axis=1)
    ]

    if not bad_balances.empty:
        examples = (
            bad_balances.head(5)
            .reset_index()
            .to_dict("records")
        )

        raise ValueError(
            f"SD_001: {len(bad_balances)} scoped journals "
            "do not balance; review complete extraction. "
            f"Examples: {examples}"
        )

    # Evaluate vendor positions after validating the complete journals.
    trigger = lines.loc[
        lines["Account Type"].eq("K")
        & lines["Posting Key"].eq("31")
        & lines["Debit Credit"].eq("H")
    ].copy()

    master = prepare_vendor_master(vendors)

    evaluated = trigger.merge(
        master,
        on=["Company", "Vendor Code"],
        how="left",
        validate="many_to_one",
        indicator=True,
    )

    matched = evaluated["_merge"].eq("both")

    evidence = (
        matched
        & evaluated["Vendor Created On"].notna()
        & evaluated["Vendor Created By"].notna()
    )

    evidence &= (
        evaluated["Vendor Created By"].ne("")
        & evaluated["User ID"].ne("")
    )

    conflict = (
        evidence
        & evaluated["Vendor Created By"].eq(
            evaluated["User ID"]
        )
    )

    conflict &= evaluated["Vendor Created On"].le(
        evaluated["Entry Date"]
    )

    conflicts = evaluated.loc[conflict].copy()
    affected = conflicts[DOC].drop_duplicates()

    # Return every BSEG position of every qualifying journal.
    output = lines.merge(
        affected,
        on=DOC,
        how="inner",
        validate="many_to_one",
    )

    output["Journal Key"] = (
        output["Company"].astype(str)
        + "|"
        + output["Fiscal Year"].astype(str)
        + "|"
        + output["Document Number"].astype(str)
    )

    trigger_keys = set(
        conflicts[KEY].itertuples(index=False, name=None)
    )

    output["Trigger"] = [
        "YES" if key in trigger_keys else "NO"
        for key in output[KEY].itertuples(
            index=False,
            name=None,
        )
    ]

    output = output.merge(
        master,
        on=["Company", "Vendor Code"],
        how="left",
        validate="many_to_one",
    )

    details = conflicts.groupby(DOC).agg(
        **{
            "Conflict Vendor Codes": (
                "Vendor Code",
                lambda values: ", ".join(sorted(set(values))),
            ),
            "Conflict User ID": (
                "User ID",
                lambda values: ", ".join(sorted(set(values))),
            ),
        }
    ).reset_index()

    output = output.merge(
        details,
        on=DOC,
        how="left",
        validate="many_to_one",
    )

    output["Same Day Creation And Invoice"] = ""
    mask = output["Trigger"].eq("YES")

    output.loc[mask, "Same Day Creation And Invoice"] = (
        output.loc[mask, "Vendor Created On"]
        .eq(output.loc[mask, "Entry Date"])
        .map({True: "YES", False: "NO"})
    )

    preferred = [
        "Journal Key",
        "Trigger",
        "Company",
        "Document Number",
        "Fiscal Year",
        "Document Item",
        "Conflict User ID",
        "Conflict Vendor Codes",
        "Account Type",
        "GL Account",
        "Vendor Code",
        "Vendor Name",
        "Vendor Created By",
        "Vendor Created On",
        "User ID",
        "Entry Date",
        "Posting Key",
        "Debit Credit",
        "Amount in Document Currency",
        "Document Currency",
        "Amount in Local Currency",
        "Local Currency",
        "Item Text",
    ]

    columns = [
        name for name in preferred
        if name in output
    ]

    output = (
        output[
            columns
            + [name for name in output if name not in columns]
        ]
        .sort_values(KEY)
        .reset_index(drop=True)
    )

    metrics = {
        "Invoice document types": ",".join(sorted(types)),
        "AP header documents": len(headers),
        "Scoped validated journals": len(selected),
        "Invoice journals with PK31/H vendor positions": len(
            trigger[DOC].drop_duplicates()
        ),
        "Trigger positions without vendor master": int(
            (~matched).sum()
        ),
        "Trigger positions without creation/user evidence": int(
            (matched & ~evidence).sum()
        ),
        "Exception journals": len(affected),
        "Exception trigger positions": len(conflicts),
        "Output positions": len(output),
    }

    return output, metrics


def remove_old_qa(output_file):
    workbook = load_workbook(output_file)
    temporary_path = None

    try:
        if "SD_001_QA" not in workbook.sheetnames:
            return

        del workbook["SD_001_QA"]

        with tempfile.NamedTemporaryFile(
            dir=Path(output_file).parent,
            suffix=".xlsx",
            delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)

        workbook.save(temporary_path)
        os.replace(temporary_path, output_file)

    finally:
        workbook.close()

        if (
            temporary_path is not None
            and temporary_path.exists()
        ):
            temporary_path.unlink()


def run_sd_001(context):
    vendors, _ = load_vendor_master(context)
    ap, _ = load_ap_items(context)

    journals, journal_metrics = load_sd_journals(context)

    output, metrics = analyze_full_journals(
        vendors,
        ap,
        journals,
        get_invoice_types(context),
    )

    path = write_control_sheet(
        context,
        "SD01",
        output,
        date_columns=[
            "Vendor Created On",
            "Company Vendor Created On",
            "Entry Date",
            "Document Date",
            "Posting Date",
        ],
        amount_columns=[
            "Amount in Document Currency",
            "Amount in Local Currency",
        ],
    )

    remove_old_qa(path)

    for name, value in {
        **journal_metrics,
        **metrics,
    }.items():
        print(f"SD_001 - {name}: {value}")

    print(f"SD_001 - output: {path}")

    _sd001_format_report(path)
    return {
        "status": "OK",
        "output_file": path,
        "sheet_name": "SD01",
        "rows": len(output),
        "summary": metrics,
    }




def _sd001_format_report(output_file):
    """Standard formatting with debit/credit and expandable detail."""
    import os
    import tempfile
    from pathlib import Path

    from openpyxl import load_workbook
    from openpyxl.styles import (
        Alignment,
        Border,
        Font,
        PatternFill,
    )
    from openpyxl.utils import get_column_letter

    output_file = Path(output_file)
    book = load_workbook(output_file)
    temporary = None

    try:
        sheet = book["SD01"]
        headers = [cell.value for cell in sheet[1]]

        labels = [
            "Invoice Posted By" if name == "User ID" else name
            for name in headers
        ]

        if len(labels) != len(set(labels)):
            raise ValueError("SD_001: duplicate headers.")

        visible = [
            "Company",
            "Document Number",
            "Fiscal Year",
            "Document Item",
            "Trigger",
            "Vendor Code",
            "Vendor Name",
            "Vendor Created By",
            "Vendor Created On",
            "Invoice Posted By",
            "Entry Date",
            "GL Account",
            "Debit Amount",
            "Credit Amount",
            "Document Currency",
            "Item Text",
        ]

        required = (
            set(visible)
            - {"Debit Amount", "Credit Amount"}
        ) | {
            "Debit Credit",
            "Amount in Document Currency",
        }

        missing = required - set(labels)

        if missing:
            raise ValueError(
                f"SD_001: missing columns {sorted(missing)}"
            )

        order = visible + [
            name for name in labels
            if name not in visible
        ]

        stored = []

        for row in sheet.iter_rows(
            min_row=2,
            max_col=len(headers),
        ):
            if any(cell.data_type == "f" for cell in row):
                raise ValueError(
                    "SD_001 contains formulas. "
                    "Automatic reordering requires review."
                )

            record = {
                name: (
                    cell.value,
                    cell.number_format,
                    cell.data_type,
                )
                for name, cell in zip(labels, row)
            }

            dc = str(
                record["Debit Credit"][0]
            ).strip().upper()

            amount = record["Amount in Document Currency"][0]

            if (
                dc not in {"S", "H"}
                or amount is None
                or isinstance(amount, bool)
            ):
                raise ValueError(
                    "SD_001: invalid amount or S/H "
                    f"in Excel row {row[0].row}"
                )

            try:
                amount = float(amount)
            except (TypeError, ValueError) as error:
                raise ValueError(
                    "SD_001: nonnumeric amount "
                    f"in Excel row {row[0].row}"
                ) from error

            # Preserve signs, including negative postings.
            record["Debit Amount"] = (
                amount if dc == "S" else 0.0,
                "#,##0.00",
                "n",
            )

            record["Credit Amount"] = (
                amount if dc == "H" else 0.0,
                "#,##0.00",
                "n",
            )

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

        dates = {
            "Vendor Created On",
            "Company Vendor Created On",
            "Entry Date",
            "Document Date",
            "Posting Date",
        }

        amounts = {
            "Debit Amount",
            "Credit Amount",
            "Amount in Document Currency",
            "Amount in Local Currency",
        }

        for column, name in enumerate(order, 1):
            header = sheet.cell(1, column)

            header.font = Font(bold=True)
            header.fill = PatternFill(
                fill_type="solid",
                fgColor="D9EAF7",
            )
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
                    cell.number_format = "dd/mm/yyyy"
                elif name in amounts:
                    cell.number_format = "#,##0.00"

                if cell.value is not None:
                    maximum_length = max(
                        maximum_length,
                        len(str(cell.value)),
                    )

            dimension = sheet.column_dimensions[
                get_column_letter(column)
            ]

            dimension.width = min(maximum_length + 2, 45)
            dimension.hidden = name not in visible
            dimension.outlineLevel = (
                0 if name in visible else 1
            )
            dimension.collapsed = False

        for dimension in sheet.row_dimensions.values():
            dimension.height = None
            dimension.hidden = False

        sheet.sheet_properties.outlinePr.summaryRight = False
        sheet.sheet_view.showOutlineSymbols = True
        sheet.sheet_format.outlineLevelCol = 1

        sheet.column_dimensions[
            get_column_letter(len(visible))
        ].collapsed = True

        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        sheet.auto_filter.filterColumn = []
        sheet.auto_filter.sortState = None
        sheet.sheet_view.showGridLines = True

        with tempfile.NamedTemporaryFile(
            dir=output_file.parent,
            suffix=".xlsx",
            delete=False,
        ) as handle:
            temporary = Path(handle.name)

        book.save(temporary)
        book.close()
        os.replace(temporary, output_file)

    finally:
        book.close()

        if temporary is not None and temporary.exists():
            temporary.unlink()



