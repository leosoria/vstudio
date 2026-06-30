"""
modules/CD/analysis.py  --  Cash Disbursements (pagos a proveedores).

analyze(row, test, params) -> DataFrame, mismo contrato que FAM/AR.
Base (import cd_main): poblacion de pagos a proveedores (OVPM, CardType='S', no anulados)
en el periodo FROM..TO. Los 4 analiticos corren en Python sobre esa base.

  CD_ANALYTIC_01_CDCS101  Cash Disbursements By Vendor       -> poblacion completa
  CD_ANALYTIC_02_CDCS102  Summary By Vendor                  -> conteo + monto por proveedor/moneda
  CD_ANALYTIC_03          Duplicate Cash Disbursements       -> mismo prov+moneda+monto en <= X dias
                          PARAM1 = ventana en dias (vacio = 30)
  CD_ANALYTIC_04          Vendors With Large # Disbursements -> proveedores con conteo alto
                          PARAM1 = metodo: STATISTICAL | FIXED | TOPN (vacio = FIXED)
                          PARAM2 = N      (umbral de conteo si FIXED / nro de proveedores si TOPN)
"""
import re
import pandas as pd
from core.analysis_base import load_import_df

DUP_KEYS = ["Company", "Vendor Code", "Payment Currency", "Payment Amount"]

DEFAULT_DUP_WINDOW = 30        # dias (CD03 si PARAM1 vacio)
DEFAULT_CD04_METHOD = "FIXED"  # CD04 si PARAM1 vacio
DEFAULT_FIXED_COUNT = 50       # N por default para FIXED si no pasan PARAM2
DEFAULT_TOPN = 10              # N por default para TOPN si no pasan PARAM2

_METHOD_ALIASES = {
    "STATISTICAL": "STATISTICAL", "STAT": "STATISTICAL", "STDEV": "STATISTICAL", "SIGMA": "STATISTICAL",
    "FIXED": "FIXED", "THRESHOLD": "FIXED", "ABS": "FIXED", "ABSOLUTE": "FIXED",
    "TOPN": "TOPN", "TOP": "TOPN", "TOP_N": "TOPN", "TOP-N": "TOPN",
}


def _num(s):
    return pd.to_numeric(s, errors="coerce").fillna(0)


def _analytic_no(test) -> int:
    """Numero de analitico desde el nombre del test. Acepta CD001 / CD_ANALYTIC_01 / ..._CDCS101."""
    t = str(test).upper()
    m = re.search(r"ANALYTIC[_-]?(\d+)", t)        # CD_ANALYTIC_01[_CDCS101]
    if not m:
        m = re.match(r"\s*CD0*(\d+)", t)            # CD001, CD01, CD1
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


def _dup_window(params) -> int:
    for t in _tokens(params):
        f = _to_float(t)
        if f is not None:
            return int(f)
    return DEFAULT_DUP_WINDOW


def _cd04_config(params):
    """Devuelve (metodo, X). metodo en {STATISTICAL, FIXED, TOPN}; X float o None."""
    method, x = None, None
    for t in _tokens(params):
        f = _to_float(t)
        if f is not None and x is None:          # PARAM numerico -> N
            x = f
        elif method is None:
            method = _METHOD_ALIASES.get(t.upper())
    return method or DEFAULT_CD04_METHOD, x


def _vendor_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Una fila por proveedor: Qty Payments (conteo) + Total Amount USD (suma en USD)."""
    d = df.copy()
    d["Payment Amount USD"] = _num(d.get("Payment Amount USD"))
    d["__date"] = pd.to_datetime(d.get("Payment Date"), errors="coerce")
    g = (d.groupby(["Company", "Vendor Code", "Vendor Name"], as_index=False, sort=False)
           .agg(**{"Qty Payments": ("Payment DocEntry", "count"),
                   "Total Amount USD": ("Payment Amount USD", "sum"),
                   "First Payment": ("__date", "min"),
                   "Last Payment": ("__date", "max")}))
    return g


def _summary(df: pd.DataFrame) -> pd.DataFrame:        # CD002
    g = _vendor_summary(df)
    g.attrs["highlight"] = "Total Amount USD"
    return g


def _duplicates(df: pd.DataFrame, window: int) -> pd.DataFrame:
    d = df.copy()
    d["Payment Amount"] = _num(d["Payment Amount"])
    d["__date"] = pd.to_datetime(d.get("Payment Date"), errors="coerce")
    flagged = []
    for _, g in d.groupby(DUP_KEYS, sort=False):
        if len(g) < 2:
            continue
        dts = g["__date"].tolist()
        idx = g.index.tolist()
        for i in range(len(g)):
            for j in range(len(g)):
                if i == j:
                    continue
                a, b = dts[i], dts[j]
                if pd.notna(a) and pd.notna(b) and abs((a - b).days) <= window:
                    flagged.append(idx[i])
                    break
    cols = list(df.columns) + ["DUP_PAYMENT_KEY", "Dup Window Days"]
    if not flagged:
        return pd.DataFrame(columns=cols)
    out = d.loc[sorted(set(flagged))].copy()
    out["DUP_PAYMENT_KEY"] = (out["Company"].astype(str) + "|" + out["Vendor Code"].astype(str)
                              + "|" + out["Payment Currency"].astype(str)
                              + "|" + out["Payment Amount"].map(lambda x: f"{x:g}"))
    out["Dup Window Days"] = window
    out = out.drop(columns="__date").sort_values(DUP_KEYS + ["Payment Date"]).reset_index(drop=True)
    out.attrs["highlight"] = "Payment Amount"
    return out


def _select_vendors(summary: pd.DataFrame, method: str, x) -> pd.DataFrame:
    """Elige proveedores por metodo sobre Qty Payments. Devuelve [Company, Vendor Code, Qty Payments]."""
    if method == "STATISTICAL":                       # media + 2 sigma por compania
        parts = []
        for _, gc in summary.groupby("Company", sort=False):
            thr = gc["Qty Payments"].mean() + 2 * gc["Qty Payments"].std(ddof=0)
            parts.append(gc[gc["Qty Payments"] > thr])
        sel = pd.concat(parts) if parts else summary.iloc[0:0]
    elif method == "TOPN":                            # top N por compania
        n = int(x) if x is not None else DEFAULT_TOPN
        parts = [gc.sort_values("Qty Payments", ascending=False).head(n)
                 for _, gc in summary.groupby("Company", sort=False)]
        sel = pd.concat(parts) if parts else summary.iloc[0:0]
    else:                                             # FIXED: conteo > N
        n = x if x is not None else DEFAULT_FIXED_COUNT
        sel = summary[summary["Qty Payments"] > n]
    return sel[["Company", "Vendor Code", "Qty Payments"]]


def _vendor_detail(df: pd.DataFrame, method: str, x) -> pd.DataFrame:
    """CD004: detalle completo de pagos (estilo CD001) de los proveedores seleccionados."""
    summary = _vendor_summary(df)
    sel = _select_vendors(summary, method, x)
    if sel.empty:
        out = df.iloc[0:0].copy(); out["Qty Payments"] = pd.Series(dtype="int64")
        return out
    detail = df.merge(sel, on=["Company", "Vendor Code"], how="inner")
    sort_cols = [c for c in ["Company", "Qty Payments", "Vendor Code", "Payment Date", "Payment DocEntry"]
                 if c in detail.columns]
    asc = [True if c != "Qty Payments" else False for c in sort_cols]
    detail = detail.sort_values(sort_cols, ascending=asc).reset_index(drop=True)
    detail.attrs["highlight"] = "Qty Payments"
    return detail


def analyze(row, test, params) -> pd.DataFrame:
    df = load_import_df(row)                          # base cd_main (poblacion de pagos)
    n = _analytic_no(test)
    if n == 1:
        return df                                     # CD_ANALYTIC_01: poblacion completa
    if n == 2:
        return _summary(df)                           # CD_ANALYTIC_02
    if n == 3:
        return _duplicates(df, _dup_window(params))   # CD_ANALYTIC_03
    if n == 4:
        method, x = _cd04_config(params)              # CD_ANALYTIC_04
        return _vendor_detail(df, method, x)
    return pd.DataFrame([{"Test": test, "Resultado": "sin logica definida"}])
