"""
FAM07 - Potential exceptions against ANKB asset-class useful life.

Inputs: AR values/class, ANLB-AFABG start, ANKB class useful life.
ANKB is matched by class, selected chart, area and validity at module TO.
The chart must match the companies being analyzed.
PARAM3 is optional: blank selects the sole chart in ANKB for the area;
multiple charts require an explicit value.
PARAM1 blank = area 01; PARAM2 blank = local-currency tolerance 0.01.

Life months = ANKB-NDJAR * 12 + NDPER; assume 12 periods per year.
Start month is included; end is the last day of the final life month.
TO must be month-end. Capitalization date is not a fallback for AFABG.
FAM06: class life ended and NBV > tolerance.
FAM07: NBV zero within tolerance before class life ends, including CAP zero.

Zero class life is counted as NOT_APPLICABLE_CLASS_LIFE_ZERO.
Missing/conflicting class or asset terms are counted as DATA_REVIEW.
Only potential exceptions are exported to this control's own worksheet.
Class-reference exceptions do not establish that SAP asset terms are incorrect.
This control is independent; it does not import or change FAM05.
"""
from pathlib import Path
import math
import re
import unicodedata
import pandas as pd

from core.fam_common import (
    apply_standard_fam_formatting,
    filter_by_company,
    get_fam_output_file,
    open_or_create_fam_output_workbook,
    parse_number,
    recreate_fam_sheet,
    save_fam_output_workbook,
    to_datetime_value,
    write_dataframe_to_sheet,
)

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

ANLB_FIELDS["years"] += ["Use"]

ANLB_FIELDS["periods"] += ["Per"]

ANLB_FIELDS["key"] += ["DepKy"]

ANLB_FIELDS["start"] += ["ODep.Start"]

KEYS = ["company", "asset", "sub"]

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

    print(f"FAM_007 AR input: {ar_path.name}; asset rows: {len(ar)}")
    print(
        f"FAM_007 ANLB input: {anlb_path.name}; "
        f"selected area: {area}; rows: {len(anlb)}"
    )
    print(f"FAM_007 conflicting ANLB terms: {len(ambiguous_keys)}")
    return fam_class_reference_sources(context, cutoff, area, ar, lookup, ambiguous_keys)


FAM07_COLUMNS = [
    "Company", "Asset", "Subnumber", "Asset description",
    "Capitalization date", "CAP (LC)",
    "Accumulated depreciation signed (LC)", "Book value (LC)",
    "Asset Class", "Balance account", "Depreciation start date",
    "Currency", "Useful life (years)", "Estimated depreciation end",
    "Analysis cut-off", "Classification", "Observation",
]


def settings(context):
    cutoff = to_datetime_value(context["module"].get("to", ""))
    if pd.isna(cutoff):
        raise ValueError("FAM07 requires a valid module TO date.")
    cutoff = cutoff.normalize()

    control = context.get("control", {})
    raw_area = text(control.get("param1", "")) or "01"
    if not re.fullmatch(r"[0-9]{1,2}(?:\.0)?", raw_area):
        raise ValueError(
            "FAM07 PARAM1 must be a valid area or blank for 01."
        )
    area = identifier(raw_area)

    raw_tolerance = text(control.get("param2", ""))
    tolerance = (
        0.01 if not raw_tolerance else number(raw_tolerance)
    )
    if tolerance is None or tolerance <= 0:
        raise ValueError(
            "FAM07 PARAM2 must be positive or blank for 0.01."
        )
    return cutoff, area, tolerance


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


    if terms.get("class_error"):
        result.update(
            classification="DATA_REVIEW",
            notes=terms["class_error"],
        )
        return result

    years = number(terms.get("years"))
    periods = number(terms.get("periods"))
    start = to_datetime_value(terms.get("start", pd.NaT))

    if years == 0 and periods == 0:
        result.update(
            classification="NOT_APPLICABLE_CLASS_LIFE_ZERO",
            notes="ANKB class useful life is zero.",
            skip=True,
        )
        return result

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
                "Missing/invalid ANKB class useful life "
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


def evaluate_fam07(cap, depreciation, nbv, terms, cutoff, tolerance, capitalized=None):
    result = fam_monthly_values(
        cap, depreciation, nbv, terms, cutoff, tolerance
    )
    if result["classification"] == "DATA_REVIEW" or result["skip"]:
        return result
    if result['end'] > pd.Timestamp(cutoff).normalize() and abs(nbv) <= tolerance:
        result.update(
            classification="ZERO_NET_VALUE_BEFORE_LIFE_END",
            notes="Net book value is zero before the ANKB class useful-life end; CAP zero is included.",
        )
    return result


def run_fam_007(context):
    cutoff, area, tolerance = settings(context)
    ar, lookup, ambiguous = read_sources(context, cutoff, area)
    records = []
    counts = {}

    for _, row in ar.iterrows():
        asset_key = tuple(row[name] for name in KEYS)
        cap, dep, nbv = [
            number(row[name]) for name in ("cap", "dep", "nbv")
        ]


        terms = lookup.get(asset_key)
        if terms is None:
            result = {
                "classification": "DATA_REVIEW",
                "life": None,
                "end": pd.NaT,
                "notes": (
                    "Conflicting ANLB terms."
                    if asset_key in ambiguous else
                    "No matching ANLB asset for selected area."
                ),
            }
            terms = {}
        else:
            result = evaluate_fam07(
                cap, dep, nbv, terms, cutoff,
                tolerance, row["capitalized"],
            )

        status = result["classification"]
        # FAM06_07_SPLIT
        if status in (
            "DATA_REVIEW",
            "NOT_APPLICABLE_CLASS_LIFE_ZERO",
        ):
            counts[status] = counts.get(status, 0) + 1
            continue
        if status != "ZERO_NET_VALUE_BEFORE_LIFE_END":
            continue
        if not status:
            continue

        counts[status] = counts.get(status, 0) + 1
        records.append({
            "Company": row["company"].zfill(4),
            "Asset": row["asset"],
            "Subnumber": row["sub"],
            "Asset description": text(row["description"]),
            "Capitalization date": to_datetime_value(
                row["capitalized"]
            ),
            "CAP (LC)": cap,
            "Accumulated depreciation signed (LC)": dep,
            "Book value (LC)": nbv,
            "Asset Class": text(row["class"]),
            "Balance account": text(row["account"]),
            "Depreciation start date": terms.get("start", pd.NaT),
            "Currency": text(row["currency"]),
            "Useful life (years)": result["life"],
            "Estimated depreciation end": result["end"],
            "Analysis cut-off": cutoff,
            "Classification": status,
            "Observation": result["notes"],
        })

    detail = pd.DataFrame(records, columns=FAM07_COLUMNS)
    output = get_fam_output_file(context)
    workbook = open_or_create_fam_output_workbook(output)
    sheet = recreate_fam_sheet(workbook, "FAM07")

    write_dataframe_to_sheet(sheet, detail)
    apply_standard_fam_formatting(
        sheet,
        detail,
        date_columns={
            "Capitalization date", "Depreciation start date",
            "Estimated depreciation end", "Analysis cut-off",
        },
        amount_columns={
            name for name in detail.columns
            if name.endswith("(LC)")
        },
    )

    index = FAM07_COLUMNS.index("Useful life (years)") + 1
    for cells in sheet.iter_rows(
        min_row=2, min_col=index, max_col=index
    ):
        cells[0].number_format = "0.00"

    save_fam_output_workbook(workbook, output)

    exceptions = counts.get("ZERO_NET_VALUE_BEFORE_LIFE_END", 0)
    print(
        f"FAM_007 area: {area.zfill(2)}; "
        f"tolerance LC: {tolerance}"
    )
    print(
        f"FAM_007 analyzed: {len(ar)}; "
        f"potential exceptions: {exceptions}; "
        f"data review: {counts.get('DATA_REVIEW', 0)}"
    )
    print(f"FAM_007 classifications: {counts}")
    print(f"FAM_007 output: {output}; worksheet FAM07 only.")


# FAM_CLASS_LIFE_ANKB_V1
def fam_class_reference_sources(context, cutoff, area, ar, lookup, ambiguous):
    """Use ANKB class life; retain ANLB start and asset identifiers."""
    from datetime import datetime, date

    path = input_file(context, "ANKB", cutoff)
    frame = pd.read_excel(path, dtype=object).dropna(how="all")
    frame.columns = [text(c) for c in frame.columns]

    fields = {
        "class": [
            "Class", "Asset Class", "ANLKL", "ANKB-ANLKL",
        ],
        "chart": [
            "ChDep", "Chart of dep.", "Chart of depreciation",
            "AFAPL", "ANKB-AFAPL",
        ],
        "area": [
            "Ar.", "Depreciation area", "AFABE", "ANKB-AFABE",
        ],
        "from": ["Valid from", "ADATU", "ANKB-ADATU"],
        "to": ["Valid to", "BDATU", "ANKB-BDATU"],
        "years": [
            "Use", "Useful life", "NDJAR", "ANKB-NDJAR",
        ],
        "periods": [
            "Per", "Per.", "Usef.life in periods",
            "Useful life in periods", "NDPER", "ANKB-NDPER",
        ],
    }

    columns = resolve_columns(frame, fields)
    frame = frame.rename(columns={
        original: logical for logical, original in columns.items()
    })
    frame = frame.loc[
        frame["area"].map(identifier).eq(area)
    ].copy()
    frame["chart"] = frame["chart"].map(
        lambda value: text(value).upper()
    )

    charts = sorted(set(frame["chart"]) - {""})
    chart = text(
        context.get("control", {}).get("param3", "")
    ).upper()

    if not chart:
        if len(charts) != 1:
            raise ValueError(
                "ANKB must contain one chart for this area, "
                f"or specify PARAM3. Charts: {charts}"
            )
        chart = charts[0]

    frame = frame.loc[frame["chart"].eq(chart)]
    if frame.empty:
        raise ValueError(
            f"ANKB has no rows for chart {chart}, area {area}."
        )

    def validity(value, lower=False):
        raw = text(value)

        if lower and raw in (
            "", "0000-01-01", "00000000",
            "00.00.0000", "00/00/0000",
        ):
            return date.min

        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, date):
            return value

        for fmt in (
            "%Y-%m-%d", "%d/%m/%Y", "%d.%m.%Y", "%Y%m%d",
        ):
            try:
                return datetime.strptime(raw, fmt).date()
            except ValueError:
                pass

        raise ValueError(
            f"Invalid ANKB validity date: {raw!r}"
        )

    terms_by_class = {}

    for _, row in frame.iterrows():
        if not text(row["class"]):
            continue

        lower = validity(row["from"], True)
        upper = validity(row["to"])

        if lower > upper:
            raise ValueError(
                "ANKB validity start is after validity end."
            )

        if not lower <= pd.Timestamp(cutoff).date() <= upper:
            continue

        name = text(row["class"]).upper()
        value = (
            number(row["years"]),
            number(row["periods"]),
        )
        terms_by_class.setdefault(name, set()).add(value)

    result = {}

    for _, row in ar.iterrows():
        asset_key = tuple(row[key] for key in KEYS)
        terms = dict(lookup.get(asset_key, {}))

        terms["asset_years"] = terms.get("years")
        terms["asset_periods"] = terms.get("periods")

        name = text(row["class"]).upper()
        values = terms_by_class.get(name, set())

        terms["chart"] = chart
        terms["class"] = name
        terms["life_source"] = "ANKB"

        if asset_key in ambiguous:
            terms["class_error"] = (
                "Conflicting ANLB asset terms."
            )
        elif asset_key not in lookup:
            terms["class_error"] = (
                "No matching ANLB asset for selected area."
            )
        elif len(values) != 1:
            terms["class_error"] = (
                "Missing or conflicting ANKB class life "
                "at analysis cutoff."
            )
        else:
            terms["years"], terms["periods"] = next(
                iter(values)
            )

        result[asset_key] = terms

    print(
        f"ANKB input: {path.name}; chart: {chart}; "
        f"area: {area}; life basis: asset class."
    )
    print(
        "Company/chart assignment must match "
        "the selected ANKB chart."
    )
    return ar, result, ambiguous

