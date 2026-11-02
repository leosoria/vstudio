"""
FAM_005 - Recorded versus recalculated accumulated depreciation.

Updated auditor criteria:
- Use ANLB-AFABG, not capitalization date.
- Life months = NDJAR * 12 + NDPER; assume 12 periods/year.
- Include the start month and cutoff month. TO must be month-end.
- Used months = min(life months, elapsed calendar months), minimum zero.
- Remaining months = life months minus used months.
- Expected depreciation = min(CAP, CAP / life months * used months),
  rounded to cents. Never exceed CAP.
- FAM05 difference = expected minus abs(recorded depreciation),
  rounded to cents.
- End = last day of the last depreciation month, starting from AFABG.
- FAM06: end <= cutoff and NBV > tolerance.
- FAM07: end > cutoff and abs(NBV) <= tolerance.

Missing AFABG produces DATA_REVIEW; capitalization is not a fallback.
Useful life comes from ANLB per asset.
Residuals, impairments, additions and historical changes may require review.
This is an audit recalculation, not a reproduction of every SAP key.

Inputs, output layout, company filter and parameters are unchanged.
PARAM1 blank = area 01; PARAM2 blank = LC tolerance 0.01.
Each control updates only its own output through its existing run function.
"""

from pathlib import Path
import math
import re
import unicodedata

import pandas as pd
from openpyxl.styles import Alignment

from core.fam_common import (
    apply_standard_fam_formatting,
    build_fx_rate_lookup,
    filter_by_company,
    get_fam_output_file,
    load_fx_rates_data,
    normalize_currency,
    open_or_create_fam_output_workbook,
    parse_number,
    recreate_fam_sheet,
    save_fam_output_workbook,
    to_datetime_value,
    write_dataframe_to_sheet,
)

MODEL = "AFABG_INCLUSIVE_MONTHS_CAPPED_AT_LIFE_AND_CAP"

AR_FIELDS = {
    "company": ["Company Code", "Empresa", "BUKRS", "ANLA-BUKRS"],
    "asset": ["Asset", "Imobilizado", "ANLN1", "ANLA-ANLN1"],
    "sub": ["Subnumber", "Subn\xba", "Subn\xb0", "SNo.", "ANLN2", "ANLA-ANLN2"],
    "description": ["Asset description", "Denomina\xe7\xe3o do imobilizado"],
    "class": ["Asset Class", "Class", "Classe", "Classe imobilizado"],
    "capitalized": ["Capitalized on", "Incorpora\xe7\xe3o em", "Dt.incorp."],
    "cap": ["Acquis.val.", "ValAquis.", "ValAquis"],
    "dep": ["Accum.dep.", "Deprecia\xe7\xe3o ac."],
    "nbv": ["Book val.", "Valor cont\xe1bil"],
    "account": ["Bal.sh.acct APC", "Conta do balan\xe7o CAP"],
    "currency": ["Currency", "Moeda"],
}

ANLB_FIELDS = {
    "company": [
        "ANLB-BUKRS", "BUKRS", "Company Code", "CoCd", "Empresa", "Empr"
    ],
    "asset": [
        "ANLB-ANLN1", "ANLN1", "Asset", "Main Asset Number", "Imobilizado"
    ],
    "sub": [
        "ANLB-ANLN2", "ANLN2", "Subnumber", "SNo.",
        "Subn\xba", "Subn\xb0", "Sbn\xba"
    ],
    "area": [
        "ANLB-AFABE", "AFABE", "Depreciation Area", "Dep. Area",
        "Ar.", "\xc1rea de deprecia\xe7\xe3o"
    ],
    "years": [
        "ANLB-NDJAR", "NDJAR", "Useful life", "Useful life in years",
        "Planned useful life in years", "Vida \xfatil", "Vida \xfatil em anos"
    ],
    "periods": [
        "ANLB-NDPER", "NDPER", "Periods", "Useful life in periods",
        "Planned useful life in periods", "Per\xedodos",
        "Vida \xfatil em per\xedodos"
    ],
    "key": [
        "ANLB-AFASL", "AFASL", "Depreciation key", "Dep. key",
        "DKey", "Chave de deprecia\xe7\xe3o", "Chave depr."
    ],
    "start": [
        "ANLB-AFABG", "AFABG", "Ord.dep.start date", "Ord. dep. start",
        "Depreciation calculation start date", "Depreciation start date",
        "Depr.norm.", "Deprecia\xe7\xe3o normal"
    ],
}


# FAM05 SAP abbreviated ANLB headers
ANLB_FIELDS["years"] += ["Use"]
ANLB_FIELDS["periods"] += ["Per"]
ANLB_FIELDS["key"] += ["DepKy"]
ANLB_FIELDS["start"] += ["ODep.Start"]


KEYS = ["company", "asset", "sub"]

DETAIL_COLUMNS = [
    "Classification", "Trigger",
    "Difference (LC)", "Expected depreciation (LC)",
    "Recorded depreciation magnitude (LC)",
    "Book value (LC)", "CAP (LC)",
    "Accumulated depreciation signed (LC)",
    "Company", "Asset", "Subnumber", "Asset description",
    "Asset Class", "Balance account", "Currency",
    "Capitalization date", "Depreciation start date",
    "Estimated depreciation end",
    "Useful life years", "Useful life periods",
    "Useful life months", "Elapsed months",
    "Depreciation area", "Depreciation key",
    "Calculation basis", "Review notes",
    "USD per LC", "FX date", "FX Status",
    "Difference (USD)", "Expected depreciation (USD)",
    "Recorded depreciation magnitude (USD)",
    "Book value (USD)", "CAP (USD)",
]


def text(value):
    return "" if pd.isna(value) else str(value).strip()


def normalized(value):
    value = unicodedata.normalize("NFKD", text(value))
    return "".join(
        c for c in value if not unicodedata.combining(c)
    ).lower().strip()


def identifier(value):
    value = text(value)
    if re.fullmatch(r"[0-9]+\.0", value):
        value = value[:-2]
    if value.isdigit():
        return value.lstrip("0") or "0"
    return value.upper()


def resolve_columns(frame, fields):
    result = {}
    for name, aliases in fields.items():
        accepted = {normalized(alias) for alias in aliases}
        matches = [
            column for column in frame.columns
            if normalized(column) in accepted
        ]
        if len(matches) != 1:
            raise ValueError(
                f"Missing or ambiguous column: {name}. "
                f"Expected one of: {aliases}. "
                f"Available columns: {list(frame.columns)}"
            )
        result[name] = matches[0]
    return result


def number(value):
    parsed = parse_number(value)
    if pd.isna(parsed):
        return None
    parsed = float(parsed)
    return parsed if math.isfinite(parsed) else None


def settings(context):
    cutoff = to_datetime_value(context["module"].get("to", ""))
    if pd.isna(cutoff):
        raise ValueError("FAM05 requires a valid module TO date.")
    cutoff = cutoff.normalize()
    if cutoff != cutoff + pd.offsets.MonthEnd(0):
        raise ValueError("FAM05 requires a month-end TO date.")

    control = context.get("control", {})

    raw_area = text(control.get("param1", "")) or "01"
    if not re.fullmatch(r"[0-9]{1,2}(?:\.0)?", raw_area):
        raise ValueError(
            "FAM05 PARAM1 must be a valid area or blank for 01."
        )
    area = identifier(raw_area)

    raw_tolerance = text(control.get("param2", ""))
    tolerance = (
        0.01 if not raw_tolerance else number(raw_tolerance)
    )
    if tolerance is None or tolerance <= 0:
        raise ValueError(
            "FAM05 PARAM2 must be positive or blank for 0.01."
        )

    print(f"FAM_005 depreciation area: {area.zfill(2)}")
    return cutoff, area, tolerance


def input_file(context, kind, cutoff):
    expected = {
        f"LBR FAM_{kind}_{cutoff:%Y%m%d}.xlsx".lower(),
        f"LBR FAM {kind}_{cutoff:%Y%m%d}.xlsx".lower(),
    }
    matches = [
        path for path in Path(context["input_folder"]).iterdir()
        if path.is_file() and path.name.lower() in expected
    ]
    if len(matches) != 1:
        raise FileNotFoundError(
            f"Expected ONE input/LBR FAM_{kind}_{cutoff:%Y%m%d}.xlsx; "
            f"found {len(matches)}."
        )
    return matches[0]


def read_sources(context, cutoff, area):
    ar_path = input_file(context, "AR", cutoff)
    anlb_path = input_file(context, "ANLB", cutoff)

    ar = pd.read_excel(ar_path, dtype=object).dropna(how="all")
    anlb = pd.read_excel(anlb_path, dtype=object).dropna(how="all")
    for frame in (ar, anlb):
        frame.columns = [text(column) for column in frame.columns]

    ar_columns = resolve_columns(ar, AR_FIELDS)
    anlb_columns = resolve_columns(anlb, ANLB_FIELDS)

    ar = ar.loc[
        ar[ar_columns["company"]].map(text).ne("")
        & ar[ar_columns["asset"]].map(text).ne("")
        & ar[ar_columns["description"]].map(text).ne("")
    ].copy()

    ar = filter_by_company(
        ar,
        ar_columns["company"],
        context["module"].get("companies", ""),
    )

    ar = ar.rename(columns={
        column: name for name, column in ar_columns.items()
    })
    anlb = anlb.rename(columns={
        column: name for name, column in anlb_columns.items()
    })

    anlb = anlb.loc[
        anlb["area"].map(identifier).eq(area)
        & anlb["company"].map(text).ne("")
        & anlb["asset"].map(text).ne("")
    ].copy()

    if anlb.empty:
        raise ValueError(
            f"ANLB has no rows for selected depreciation area {area}."
        )

    for frame in (ar, anlb):
        for name in KEYS:
            frame[name] = frame[name].map(identifier)
        if frame["sub"].eq("").any():
            raise ValueError(
                "Asset subnumber is blank. Export ANLN2 including subnumber 0."
            )

    if ar.duplicated(KEYS).any():
        raise ValueError(
            "FAM_AR contains duplicate Company+Asset+Subnumber."
        )

    columns = KEYS + ["area", "years", "periods", "key", "start"]
    terms = anlb[columns].copy()
    for name in ("years", "periods"):
        terms[name] = terms[name].map(number)
    terms["key"] = terms["key"].map(text)
    terms["start"] = terms["start"].map(to_datetime_value)
    terms = terms.drop_duplicates(columns)

    ambiguous = terms.duplicated(KEYS, keep=False)
    ambiguous_keys = set(
        map(tuple, terms.loc[ambiguous, KEYS].values.tolist())
    )
    unique = terms.loc[~ambiguous].set_index(KEYS)
    lookup = {
        asset_key: row.to_dict()
        for asset_key, row in unique.iterrows()
    }

    print(f"FAM_005 AR input: {ar_path.name}; asset rows: {len(ar)}")
    print(
        f"FAM_005 ANLB input: {anlb_path.name}; "
        f"selected area: {area}; rows: {len(anlb)}"
    )
    print(f"FAM_005 conflicting ANLB terms: {len(ambiguous_keys)}")
    return ar, lookup, ambiguous_keys


def fam_monthly_values(cap, depreciation, nbv, terms, cutoff, tolerance):
    """Inclusive months from AFABG, capped at useful life and CAP."""
    from decimal import Decimal, ROUND_HALF_UP

    result = {
        "classification": "", "notes": "", "skip": False,
        "expected": None, "actual": None, "difference": None,
        "months": None, "elapsed": None, "remaining": None,
        "life": None, "end": pd.NaT,
    }

    if (
        any(
            value is None or not math.isfinite(float(value))
            for value in (cap, depreciation, nbv)
        )
        or float(cap) < -tolerance
        or float(nbv) < -tolerance
    ):
        result.update(
            classification="DATA_REVIEW",
            notes="Missing/invalid amounts or negative CAP/NBV.",
        )
        return result

    if float(cap) <= tolerance:
        result["skip"] = True
        return result

    years = number(terms.get("years"))
    periods = number(terms.get("periods"))
    start = to_datetime_value(terms.get("start", pd.NaT))

    if (
        years is None or periods is None
        or years < 0 or periods < 0
        or years != int(years) or periods != int(periods)
        or periods >= 12 or years * 12 + periods <= 0
        or pd.isna(start)
    ):
        result.update(
            classification="DATA_REVIEW",
            notes=(
                "Missing/invalid ANLB useful life "
                "or depreciation start (AFABG)."
            ),
        )
        return result

    cutoff = pd.Timestamp(cutoff).normalize()
    if cutoff != cutoff + pd.offsets.MonthEnd(0):
        raise ValueError(
            "FAM05/06/07 monthly method requires a month-end TO date."
        )

    months = int(years * 12 + periods)
    if months > 12000:
        result.update(
            classification="DATA_REVIEW",
            notes="Useful life outside supported range.",
        )
        return result

    start = pd.Timestamp(start).normalize()
    try:
        end = (
            start.replace(day=1)
            + pd.DateOffset(months=months)
            - pd.Timedelta(days=1)
        )
    except (ValueError, OverflowError, pd.errors.OutOfBoundsDatetime):
        result.update(
            classification="DATA_REVIEW",
            notes="Estimated end date outside supported range.",
        )
        return result

    elapsed = (
        0 if start > cutoff else
        min(
            months,
            max(
                0,
                (cutoff.year - start.year) * 12
                + cutoff.month - start.month + 1,
            ),
        )
    )

    cap_value = Decimal(str(cap))
    expected = min(
        cap_value,
        cap_value / Decimal(months) * Decimal(elapsed),
    )
    expected = expected.quantize(
        Decimal("0.01"), rounding=ROUND_HALF_UP
    )
    expected = min(cap_value, expected)

    actual = abs(Decimal(str(depreciation)))
    difference = (expected - actual).quantize(
        Decimal("0.01"), rounding=ROUND_HALF_UP
    )

    result.update(
        expected=float(expected),
        actual=float(actual),
        difference=float(difference),
        months=months,
        elapsed=elapsed,
        remaining=months - elapsed,
        life=months / 12,
        end=end,
    )
    return result


def evaluate(cap, depreciation, nbv, terms, cutoff, tolerance, capitalization_date=None):
    """Use ANLB depreciation start; capitalization is informational."""
    from decimal import Decimal

    result = fam_monthly_values(
        cap, depreciation, nbv, terms, cutoff, tolerance
    )
    result["trigger"] = "NO"

    if result["classification"] == "DATA_REVIEW":
        return result

    if result["skip"]:
        result["classification"] = "NO_CAPITALIZED_VALUE"
        return result

    if not text(terms.get("key")):
        result.update(
            classification="DATA_REVIEW",
            notes="Missing ANLB depreciation key.",
        )
        return result

    if (
        abs(Decimal(str(result["difference"])))
        > Decimal(str(tolerance))
    ):
        result.update(
            classification="DEPRECIATION_DIFFERENCE_ESTIMATE",
            trigger="YES",
        )
    else:
        result["classification"] = "WITHIN_TOLERANCE"

    result["notes"] = (
        "Inclusive months from AFABG, capped at useful life and CAP; "
        "amounts rounded to cents."
    )
    return result


def fx_details(currency, lookup):
    details = lookup.get(currency, {})
    rate = number(details.get("rate"))
    date = details.get("date", pd.NaT)

    if rate is None:
        return None, date, "RATE_NOT_FOUND"
    if rate <= 0 or details.get("status", "OK") != "OK":
        return None, date, "INVALID_RATE"
    return 1.0 / rate, date, "OK"


def compact_fam05_detail(detail, cutoff):
    """Layout similar to Depreciation a) in Reproceso FAM."""
    result = detail[[
        "Company", "Asset", "Subnumber", "Asset description",
        "Capitalization date", "CAP (LC)",
        "Accumulated depreciation signed (LC)", "Book value (LC)",
        "Asset Class", "Depreciation start date", "Currency",
    ]].copy()

    years = pd.to_numeric(
        detail["Useful life years"], errors="coerce"
    )
    periods = pd.to_numeric(
        detail["Useful life periods"], errors="coerce"
    )

    result["Useful life (years)"] = years + periods / 12
    result["Estimated depreciation end"] = (
        detail["Estimated depreciation end"]
    )
    result["Analysis cut-off"] = cutoff
    result["Months"] = detail["Elapsed months"]
    result["Depreciation recalculation (LC)"] = (
        detail["Expected depreciation (LC)"]
    )
    result["Difference (LC)"] = detail["Difference (LC)"]
    result["Observation"] = [
        "DATA_REVIEW: " + str(note)
        if status == "DATA_REVIEW" else ""
        for status, note in zip(
            detail["Classification"], detail["Review notes"]
        )
    ]
    return result


def format_compact_fam05_detail(sheet, frame):
    for index, name in enumerate(frame.columns, start=1):
        if name in {
            "Analysis cut-off", "Capitalization date",
            "Depreciation start date", "Estimated depreciation end"
        }:
            number_format = "dd/mm/yyyy"
        elif name == "Months":
            number_format = "0"
        elif name == "Useful life (years)":
            number_format = "0.00"
        else:
            continue

        for cells in sheet.iter_rows(
            min_row=2, min_col=index, max_col=index
        ):
            cells[0].number_format = number_format


def run_fam_005(context):
    cutoff, area, tolerance = settings(context)
    ar, lookup, ambiguous = read_sources(context, cutoff, area)
    fx_lookup = build_fx_rate_lookup(load_fx_rates_data(context))

    records = []
    counts = {}

    for _, row in ar.iterrows():
        asset_key = tuple(row[name] for name in KEYS)
        terms = lookup.get(asset_key)
        cap, dep, nbv = [
            number(row[name]) for name in ("cap", "dep", "nbv")
        ]

        if terms is None:
            result = {
                "classification": "DATA_REVIEW",
                "trigger": "NO",
                "expected": None,
                "actual": abs(dep) if dep is not None else None,
                "difference": None,
                "months": None,
                "elapsed": None,
                "end": pd.NaT,
                "notes": (
                    "Conflicting ANLB terms."
                    if asset_key in ambiguous else
                    "No matching ANLB asset for selected area."
                ),
            }
            terms = {}
        else:
            result = evaluate(cap, dep, nbv, terms, cutoff, tolerance, row['capitalized'])

        classification = result["classification"]
        counts[classification] = counts.get(classification, 0) + 1

        if result["trigger"] != "YES" and classification != "DATA_REVIEW":
            continue

        currency = normalize_currency(row["currency"])
        rate, date, status = fx_details(currency, fx_lookup)

        record = {
            "Classification": classification,
            "Trigger": result["trigger"],
            "Difference (LC)": result["difference"],
            "Expected depreciation (LC)": result["expected"],
            "Recorded depreciation magnitude (LC)": result["actual"],
            "Book value (LC)": nbv,
            "CAP (LC)": cap,
            "Accumulated depreciation signed (LC)": dep,
            "Company": row["company"].zfill(4),
            "Asset": row["asset"],
            "Subnumber": row["sub"],
            "Asset description": text(row["description"]),
            "Asset Class": text(row["class"]),
            "Balance account": text(row["account"]),
            "Currency": currency,
            "Capitalization date": to_datetime_value(row["capitalized"]),
            "Depreciation start date": terms.get("start", pd.NaT),
            "Estimated depreciation end": result["end"],
            "Useful life years": terms.get("years"),
            "Useful life periods": terms.get("periods"),
            "Useful life months": result["months"],
            "Elapsed months": result["elapsed"],
            "Depreciation area": area.zfill(2),
            "Depreciation key": terms.get("key", ""),
            "Calculation basis": MODEL,
            "Review notes": result["notes"],
            "USD per LC": rate,
            "FX date": date,
            "FX Status": status,
        }

        for base in (
            "Difference", "Expected depreciation",
            "Recorded depreciation magnitude", "Book value", "CAP"
        ):
            value = record[base + " (LC)"]
            record[base + " (USD)"] = (
                None if value is None or rate is None else value * rate
            )
        records.append(record)

    detail = pd.DataFrame(records, columns=DETAIL_COLUMNS)
    detail = compact_fam05_detail(detail, cutoff)

    summary_rows = [
        ("Analysis cut-off", cutoff),
        ("Depreciation area (PARAM1)", area.zfill(2)),
        ("LC tolerance (PARAM2)", tolerance),
        ("Assets analyzed", len(ar)),
        (
            "Potential exceptions",
            sum(record["Trigger"] == "YES" for record in records),
        ),
        ("Data review records", counts.get("DATA_REVIEW", 0)),
        ("Detail rows", len(detail)),
        (
            "Detail rows with FX issues",
            sum(record["FX Status"] != "OK" for record in records),
        ),
        ("Method", MODEL),
        (
            "Period convention",
            "AFABG; inclusive calendar months; "
            "used months capped at useful life; depreciation capped at CAP.",
        ),
        (
            "Inputs",
            "AR balances and ANLB extraction snapshot; "
            "historical terms not reconstructed.",
        ),
        (
            "Interpretation",
            "Audit estimates. A trigger is not proof of a SAP error.",
        ),
    ]
    summary_rows += [
        ("Classification: " + name, count)
        for name, count in sorted(counts.items())
    ]
    summary = pd.DataFrame(
        summary_rows, columns=["Metric", "Result"]
    )

    output = get_fam_output_file(context)
    workbook = open_or_create_fam_output_workbook(output)

    for name, frame in (
        ("FAM05", summary), ("FAM05_Detail", detail)
    ):
        sheet = recreate_fam_sheet(workbook, name)
        write_dataframe_to_sheet(sheet, frame)
        apply_standard_fam_formatting(
            sheet,
            frame,
            date_columns={
                "Capitalization date", "Depreciation start date",
                "Estimated depreciation end", "FX date",
            },
            amount_columns={
                column for column in frame.columns
                if column.endswith("(LC)") or column.endswith("(USD)")
            },
            integer_columns={
                "Useful life years", "Useful life periods",
                "Useful life months", "Elapsed months",
            },
        )

        if name == "FAM05":
            sheet.column_dimensions["A"].width = 42
            sheet.column_dimensions["B"].width = 100
            for cell in sheet["B"]:
                cell.alignment = Alignment(
                    wrap_text=True, vertical="top"
                )
        else:
            format_compact_fam05_detail(sheet, frame)

    # FAM05_FINAL_SHEET
    if "FAM05_Detail" in workbook:
        if "FAM05" in workbook:
            del workbook["FAM05"]
        workbook["FAM05_Detail"].title = "FAM05"
    print("FAM05: se guarda solamente la hoja FAM05 para este control.")
    save_fam_output_workbook(workbook, output)

    print(
        f"FAM_005: {len(ar)} analyzed; "
        f"{len(detail)} detail rows; classifications: {counts}"
    )
    print(f"FAM_005 output: {output}")
    print("FAM_005 replaced only FAM05 and FAM05_Detail.")
