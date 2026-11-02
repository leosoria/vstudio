"""
FAM_007 - Zero net book value before useful-life end.

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
    return ar, lookup, ambiguous_keys


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


def evaluate_fam07(cap, depreciation, nbv, terms, cutoff, tolerance, capitalized=None):
    result = fam_monthly_values(
        cap, depreciation, nbv, terms, cutoff, tolerance
    )
    if result["classification"] == "DATA_REVIEW" or result["skip"]:
        return result
    if result['end'] > pd.Timestamp(cutoff).normalize() and abs(nbv) <= tolerance:
        result.update(
            classification="ZERO_NET_VALUE_BEFORE_LIFE_END",
            notes="Net book value is zero before the useful-life end.",
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

        if cap is not None and abs(cap) <= tolerance:
            continue

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
        if status == "DATA_REVIEW":
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

    exceptions = sum(
        count for status, count in counts.items()
        if status != "DATA_REVIEW"
    )
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
