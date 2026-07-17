"""
modules/GL/analysis.py  --  General Journal Analysis (GL_ANALYTIC_01..16).

analyze(row, test, params) -> DataFrame, mismo contrato y estilo que FAM/AR/CD/PO.

Bases (output/{scope}_GL_{import}_{AAAAMMDD}.xlsx):
  - gl_journal_lines    primario; una fila por linea de asiento del periodo
  - gl_accounts         maestro de cuentas (OACT, completo)
  - gl_account_activity MAX(RefDate) por cuenta antes del periodo
  - gl_reversals        asientos de reversa hasta DATE_TO

Notas de criterio (acordadas con el equipo):
  01 fin de semana   -> por ENTRY DATE (alta), no posting; a nivel asiento
  02 redondo de mil  -> total del asiento; PARAM1 basis LOCAL|USD (default LOCAL)
  03 palabras        -> lista default (espanol) ampliable/reemplazable por PARAM1 (coma); + narraciones en blanco
  04 creo=aprobo     -> Creator ID == Approver ID, a nivel asiento
  06 reversados      -> misma cuenta + monto absoluto, count > N (PARAM1, default 2 => 3 o mas)
  07 periodo anterior-> posteado a un FY anterior al de alta / antes de apertura; PARAM1 = FY (filtra alta)
  08 inactivas       -> (a) cuenta Frozen +N meses  Y/O  (b) sin movimiento +N meses; PARAM1 = meses (default 6)
  09 dup cuenta+monto-> monto en USD (reporting currency)
  10 dup cuenta+desc -> descripcion = LineMemo o, si vacia, Memo de cabecera
  11 dup desc+monto  -> idem desc + monto USD
  12..15 reportes    -> resumenes (no excepciones)
  16 cuenta elegida  -> PARAM1 = codigo de cuenta
"""
import os
import re
import glob
import unicodedata
import pandas as pd
from core.analysis_base import load_import_df, OUTPUT_DIR

DEFAULT_REVERSAL_MIN = 2        # 06: count > 2  => 3 o mas
DEFAULT_INACTIVE_MONTHS = 6     # 08

# 03: lista default (la del equipo, en espanol). Reemplazable por PARAM1.
DEFAULT_SUSPICIOUS_WORDS = [
    "fraude", "error", "robo", "cancelamiento", "destrucción", "consultoría",
    "donación", "descuento especial", "préstamo", "adelanto", "asiento extraordinario",
    "corrección fuera de cierre", "sin autorización", "asientos sin soporte", "sin soporte",
    "fraude ocupacional", "sin evidencia", "fuera de cierre", "ajuste de último minuto",
    "asiento retroactivo", "gasto omitido", "ingreso ficticio", "provisión inflada",
]


# ----------------------------------------------------------------------------- helpers
def _num(s):
    return pd.to_numeric(s, errors="coerce")


def _date(s):
    return pd.to_datetime(s, errors="coerce")


def _norm(s):
    """minuscula + sin acentos, para comparar texto en espanol de forma robusta."""
    s = unicodedata.normalize("NFKD", str(s).lower())
    return "".join(c for c in s if not unicodedata.combining(c))


def _analytic_no(test) -> int:
    t = str(test).upper()
    m = re.search(r"ANALYTIC[_-]?(\d+)", t)
    if not m:
        m = re.match(r"\s*GL0*(\d+)", t)
    return int(m.group(1)) if m else 0


def _tokens(params) -> list:
    if isinstance(params, dict):
        seq = list(params.values())
    elif isinstance(params, (list, tuple)):
        seq = list(params)
    elif params in (None, ""):
        seq = []
    else:
        seq = [params]
    return [str(t).strip() for t in seq if str(t).strip() != ""]


def _to_float(v):
    try:
        return float(str(v).strip())
    except (TypeError, ValueError):
        return None


def _first_num(params, default):
    for t in _tokens(params):
        f = _to_float(t)
        if f is not None:
            return f
    return default


def _first_str(params):
    toks = _tokens(params)
    return toks[0] if toks else None


def _load_named_import(row, name):
    fname = f"{row.scope}_{row.module}_{name}_{row.date_to.strftime('%Y%m%d')}.xlsx"
    path = os.path.join(OUTPUT_DIR, fname)
    if os.path.exists(path):
        return pd.read_excel(path)
    alts = glob.glob(os.path.join(OUTPUT_DIR, f"*_{row.module}_{name}_*.xlsx"))
    return pd.read_excel(sorted(alts)[-1]) if alts else None


def _missing(name):
    return pd.DataFrame([{"Aviso": f"(falta el import {name} -- corre run_import --only GL)"}])


def _empty_like(df, extra=None):
    out = df.iloc[0:0].copy()
    for c in (extra or []):
        out[c] = pd.Series(dtype="object")
    return out


def _headers(df):
    """Una fila por asiento (TransId): para tests a nivel cabecera."""
    return df.drop_duplicates(subset=["TransId"]).copy()


def _description(df):
    """Descripcion de linea; si vacia, usa el Memo de cabecera."""
    line = df.get("Line Memo")
    head = df.get("Journal Memo")
    line = line.astype(str).where(line.notna(), "")
    head = head.astype(str).where(head.notna(), "")
    desc = line.str.strip()
    desc = desc.where(desc != "", head.str.strip())
    return desc


# ----------------------------------------------------------------------------- 01..04 (cabecera)
def _weekend(df):                                              # 01
    h = _headers(df)
    wd = pd.to_numeric(h.get("Entry Weekday"), errors="coerce")
    ed = _date(h.get("Entry Date"))
    wd = wd.where(wd.notna(), ed.dt.dayofweek)                 # respaldo si falta la col
    out = h[wd.isin([5, 6])].reset_index(drop=True)            # 5=sab, 6=dom
    out.attrs["highlight"] = "Entry Date"
    return out


def _round_thousand(df, basis):                               # 02
    h = _headers(df)
    col = "Header Total USD" if str(basis).upper() == "USD" else "Header Total Local"
    amt = _num(h.get(col)).round(2).abs()
    flag = (amt > 0) & ((amt % 1000).abs() < 0.01)
    out = h[flag].reset_index(drop=True)
    out.attrs["highlight"] = col
    return out


def _suspicious_words(df, params):                            # 03
    words = _tokens(params) or DEFAULT_SUSPICIOUS_WORDS
    norm_words = [_norm(w) for w in words]
    text = (df.get("Journal Memo").astype(str).fillna("") + " " +
            df.get("Line Memo").astype(str).fillna("")).map(_norm)
    raw = (df.get("Journal Memo").astype(str).where(df.get("Journal Memo").notna(), "").str.strip() + " " +
           df.get("Line Memo").astype(str).where(df.get("Line Memo").notna(), "").str.strip()).str.strip()

    def _match(t):
        for w in norm_words:
            if w and w in t:
                return w
        return None

    matched = text.map(_match)
    is_blank = raw == ""
    keep = matched.notna() | is_blank
    out = df[keep].copy()
    out["Matched Word"] = matched[keep]
    out.loc[is_blank[keep], "Matched Word"] = "(narracion en blanco)"
    out = out.reset_index(drop=True)
    out.attrs["highlight"] = "Matched Word"
    return out


def _created_approved(df):                                    # 04
    h = _headers(df)
    cre = h.get("Creator ID")
    app = h.get("Approver ID")
    mask = (cre.notna() & app.notna()
            & (cre.astype(str).str.strip() == app.astype(str).str.strip())
            & (cre.astype(str).str.strip() != ""))
    out = h[mask].reset_index(drop=True)
    out.attrs["highlight"] = "Approver ID"
    return out


# ----------------------------------------------------------------------------- 06 (reversals)
def _frequently_reversed(rev, min_count):                     # 06
    if rev is None:
        return _missing("gl_reversals")
    d = rev.copy()
    d["Line Amount Abs"] = _num(d.get("Line Amount Abs"))
    keys = ["Company", "Account Code", "Line Amount Abs"]
    g = d.groupby(keys, sort=False)["TransId"].transform("nunique")
    out = d[g > min_count].copy()                             # > N  (default 2 => 3 o mas)
    if out.empty:
        return _empty_like(d, ["Reversal Count"])
    cnt = out.groupby(keys, sort=False)["TransId"].transform("nunique")
    out["Reversal Count"] = cnt
    out = out.sort_values(keys + ["Posting Date"]).reset_index(drop=True)
    out.attrs["highlight"] = "Reversal Count"
    return out


# ----------------------------------------------------------------------------- 07 (periodo anterior)
def _prior_period(df, fy):                                    # 07
    h = _headers(df)
    post = _date(h.get("Posting Date"))
    entry = _date(h.get("Entry Date"))
    post_y = post.dt.year
    entry_y = entry.dt.year
    before_open = h.get("Posted Before Period Open").astype(str).str.upper() == "Y"
    # posteado a un anio fiscal anterior al de alta, o antes de la apertura del periodo
    mask = ((post_y.notna() & entry_y.notna() & (post_y < entry_y)) | before_open)
    if fy is not None:
        try:
            mask = mask & (entry_y == int(float(fy)))
        except (TypeError, ValueError):
            pass
    out = h[mask].copy()
    out["Prior Period Reason"] = ""
    rp = _date(out.get("Posting Date")).dt.year
    re_ = _date(out.get("Entry Date")).dt.year
    out.loc[(rp < re_), "Prior Period Reason"] = "Posteado a FY anterior al de alta"
    out.loc[out.get("Posted Before Period Open").astype(str).str.upper() == "Y", "Prior Period Reason"] = \
        out["Prior Period Reason"].str.cat(["; antes de apertura del periodo"] * len(out), na_rep="").str.strip("; ")
    out = out.reset_index(drop=True)
    out.attrs["highlight"] = "Posting Date"
    return out


# ----------------------------------------------------------------------------- 08 (cuentas inactivas)
def _inactive_accounts(row, df, months):                      # 08
    acc = _load_named_import(row, "gl_accounts")
    act = _load_named_import(row, "gl_account_activity")
    if acc is None:
        return _missing("gl_accounts")
    cutoff_days = int(months) * 30

    def _key(d):  # claves de cruce siempre como texto (los codigos de cuenta numericos
        d["Company"] = d["Company"].astype(str).str.strip()      # se releen como int desde Excel)
        d["Account Code"] = d["Account Code"].astype(str).str.strip()
        return d

    # postings del periodo, una fila por (asiento, cuenta) para no inflar
    j = _key(df[["Company", "TransId", "Journal Number", "Account Code", "Posting Date",
                 "Line Amount Local", "Line Amount USD"]].copy())
    j["Posting Date"] = _date(j["Posting Date"])

    # (a) cuenta marcada inactiva/Frozen hace +N meses
    a = _key(acc[["Company", "Account Code", "Inactive", "Inactive From", "Active"]].copy())
    a["Inactive From"] = _date(a["Inactive From"])
    m = j.merge(a, on=["Company", "Account Code"], how="left")
    flag_a = ((m["Inactive"].astype(str).str.upper() == "Y") &
              m["Inactive From"].notna() &
              ((m["Posting Date"] - m["Inactive From"]).dt.days >= cutoff_days))

    # (b) sin movimiento en +N meses antes del posteo
    if act is not None:
        b = _key(act[["Company", "Account Code", "Last Movement Before Period"]].copy())
        b["Last Movement Before Period"] = _date(b["Last Movement Before Period"])
        m = m.merge(b, on=["Company", "Account Code"], how="left")
        gap = (m["Posting Date"] - m["Last Movement Before Period"]).dt.days
        flag_b = m["Last Movement Before Period"].isna() | (gap >= cutoff_days)
        m["Months Since Last Movement"] = (gap / 30).round(1)
    else:
        flag_b = pd.Series(False, index=m.index)
        m["Months Since Last Movement"] = pd.NA

    m["Inactive Flag (a) Frozen"] = flag_a.map({True: "Y", False: "N"})
    m["Inactive Flag (b) No Movement"] = flag_b.map({True: "Y", False: "N"})
    out = m[flag_a | flag_b].reset_index(drop=True)
    out.attrs["highlight"] = "Account Code"
    return out


# ----------------------------------------------------------------------------- 09..11 (duplicados)
def _dup(df, keys, label_amt=None):
    d = df.copy()
    g = d.groupby(keys, sort=False)["TransId"].transform("nunique")
    out = d[g > 1].copy()
    if out.empty:
        return _empty_like(d, ["DUP_KEY"])
    out["DUP_KEY"] = out[keys].astype(str).agg("|".join, axis=1)
    out = out.sort_values(keys + ["TransId"]).reset_index(drop=True)
    if label_amt:
        out.attrs["highlight"] = label_amt
    return out


def _dup_account_amount(df):                                  # 09 (USD)
    d = df.copy()
    d["Amount USD r"] = _num(d.get("Line Amount USD")).round(2)
    return _dup(d, ["Company", "Account Code", "Amount USD r"], "Line Amount USD")


def _dup_account_desc(df):                                    # 10
    d = df.copy()
    d["Desc"] = _description(d)
    d = d[d["Desc"].str.strip() != ""]
    if d.empty:
        return _empty_like(df, ["DUP_KEY"])
    return _dup(d, ["Company", "Account Code", "Desc"], "Desc")


def _dup_desc_amount(df):                                     # 11
    d = df.copy()
    d["Desc"] = _description(d)
    d["Amount USD r"] = _num(d.get("Line Amount USD")).round(2)
    d = d[d["Desc"].str.strip() != ""]
    if d.empty:
        return _empty_like(df, ["DUP_KEY"])
    return _dup(d, ["Company", "Desc", "Amount USD r"], "Line Amount USD")


# ----------------------------------------------------------------------------- 12..15 (reportes)
def _count_per_account(df):                                   # 12
    d = df.copy()
    g = (d.groupby(["Company", "Account Code", "Account Name"], as_index=False, sort=False)
           .agg(**{"Journals": ("TransId", "nunique"),
                   "Lines": ("TransId", "count"),
                   "Total Local": ("Line Amount Local", lambda s: _num(s).sum()),
                   "Total USD": ("Line Amount USD", lambda s: _num(s).sum())}))
    g = g.sort_values(["Company", "Journals"], ascending=[True, False]).reset_index(drop=True)
    g.attrs["highlight"] = "Journals"
    return g


def _summary_account_month(df):                               # 13
    d = df.copy()
    g = (d.groupby(["Company", "Posting Month", "Account Code", "Account Name"], as_index=False, sort=False)
           .agg(**{"Net Local (D-C)": ("Line Amount Local", lambda s: _num(s).sum()),
                   "Net USD (D-C)": ("Line Amount USD", lambda s: _num(s).sum()),
                   "Lines": ("TransId", "count")}))
    g = g.sort_values(["Company", "Posting Month", "Account Code"]).reset_index(drop=True)
    g.attrs["highlight"] = "Net USD (D-C)"
    return g


def _accounts_by_type(df):                                    # 14
    d = df.copy()
    g = (d.groupby(["Company", "Account Code", "Account Name", "Journal Type"], as_index=False, sort=False)
           .agg(**{"Journals": ("TransId", "nunique"),
                   "Total USD": ("Line Amount USD", lambda s: _num(s).sum())}))
    g = g.sort_values(["Company", "Account Code", "Journal Type"]).reset_index(drop=True)
    g.attrs["highlight"] = "Journal Type"
    return g


def _accounts_by_poster(df):                                  # 15
    d = df.copy()
    g = (d.groupby(["Company", "Account Code", "Account Name", "Creator ID", "Creator Name"],
                   as_index=False, sort=False)
           .agg(**{"Journals": ("TransId", "nunique"),
                   "Total USD": ("Line Amount USD", lambda s: _num(s).sum())}))
    g = g.sort_values(["Company", "Account Code", "Creator ID"]).reset_index(drop=True)
    g.attrs["highlight"] = "Creator ID"
    return g


def _selected_account(df, account):                           # 16
    if not account:
        return pd.DataFrame([{"Aviso": "Indicar la cuenta en PARAM1 (GL_ANALYTIC_16)"}])
    d = df[df.get("Account Code").astype(str).str.strip() == str(account).strip()].reset_index(drop=True)
    d.attrs["highlight"] = "Account Code"
    return d


# ----------------------------------------------------------------------------- dispatcher
def _accounts_created(row):                                   # 05
    acc = _load_named_import(row, "gl_accounts")
    if acc is None:
        return _missing("gl_accounts")
    created = pd.to_datetime(acc.get("Creation Date"), errors="coerce")
    out = acc[(created >= pd.Timestamp(row.date_from)) & (created <= pd.Timestamp(row.date_to))].reset_index(drop=True)
    out.attrs["highlight"] = "Creation Date"
    return out


def analyze(row, test, params) -> pd.DataFrame:
    n = _analytic_no(test)

    if n == 5:
        return _accounts_created(row)
    if n == 6:
        return _frequently_reversed(_load_named_import(row, "gl_reversals"), int(_first_num(params, DEFAULT_REVERSAL_MIN)))

    if n in (1, 2, 3, 4, 7, 9, 10, 11, 12, 13, 14, 15, 16):
        df = load_import_df(row)                              # gl_journal_lines
        if df is None or df.empty:
            return _missing("gl_journal_lines")
        if n == 1:
            return _weekend(df)
        if n == 2:
            return _round_thousand(df, _first_str(params) or "LOCAL")
        if n == 3:
            return _suspicious_words(df, params)
        if n == 4:
            return _created_approved(df)
        if n == 7:
            return _prior_period(df, _first_str(params))
        if n == 9:
            return _dup_account_amount(df)
        if n == 10:
            return _dup_account_desc(df)
        if n == 11:
            return _dup_desc_amount(df)
        if n == 12:
            return _count_per_account(df)
        if n == 13:
            return _summary_account_month(df)
        if n == 14:
            return _accounts_by_type(df)
        if n == 15:
            return _accounts_by_poster(df)
        if n == 16:
            return _selected_account(df, _first_str(params))

    if n == 8:
        df = load_import_df(row)
        if df is None or df.empty:
            return _missing("gl_journal_lines")
        return _inactive_accounts(row, df, int(_first_num(params, DEFAULT_INACTIVE_MONTHS)))

    return pd.DataFrame([{"Test": test, "Resultado": "sin logica definida"}])
