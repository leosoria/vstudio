"""
modules/PO/analysis.py  --  Purchase Order Management (tests PO_ANALYTIC_01..11).

analyze(row, test, params) -> DataFrame, mismo contrato y estilo que FAM/AR/CD.

Bases (imports que corre run_import; quedan en output/ como {scope}_PO_{import}_{AAAAMMDD}.xlsx):
  - po_lines  -> primario (load_import_df): una fila por linea de PO + GR + aprobacion + link PR.
  - pr_lines  -> extra (por nombre): una fila por linea de Solicitud de Compra (PR) + PO vinculada.

Analiticos:
  01 Split POs            mismo vendor+material+creador en <= X dias (PARAM1=dias, vacio=7)   [po_lines]
  02 Duplicate POs        mismo vendor+material+cantidad (sin ventana, todo el periodo)        [po_lines]
  03 PO creada y aprobada por el mismo usuario                                                 [po_lines]
  04 GR antes de la fecha de PO (excluye GR sin fecha)                                          [po_lines]
  05 GR > N dias de la aprobacion de la PO (PARAM1=dias, vacio=30)                              [po_lines]
  06 Dif. de precio mismo vendor/material (marca cualquier dif; PARAM1=umbral opcional)         [po_lines]
  07 Mismo usuario creo la PO y recepciono el GR                                                [po_lines]
  08 POs por item por mes (resumen)                                                             [po_lines]
  09 Split PRs            mismo material+creador en <= X dias (PARAM1=dias, vacio=7)            [pr_lines]
  10 PO vs PR            diferencias de item/cantidad PO vs su PR                              [pr_lines + po_lines]
  11 POs sin PR          lineas de PO sin Solicitud de Compra (From PR = N)                    [po_lines]

Notas:
  - "Mismo material" = por Item Code (como el ACL). Las lineas sin Item Code (servicios, usan cuenta)
    se excluyen de los tests por material (01, 02, 06, 09). Cuando se confirme, se agrega AcctCode.
  - USD ya viene convertido al rate del dia del documento, con USD Rate / USD Rate Date visibles.
  - Las POs anuladas (PO Canceled = 'Y') NO se filtran; quedan visibles por su columna.
"""
import os
import re
import glob
import pandas as pd
from core.analysis_base import load_import_df, OUTPUT_DIR

DEFAULT_SPLIT_WINDOW = 7      # dias (01 / 09 si PARAM1 vacio)
DEFAULT_GR_DAYS = 30          # dias (05 si PARAM1 vacio)


# ----------------------------------------------------------------------------- helpers
def _num(s):
    return pd.to_numeric(s, errors="coerce")


def _date(s):
    return pd.to_datetime(s, errors="coerce")


def _analytic_no(test) -> int:
    """Numero de analitico desde el nombre. Acepta PO_ANALYTIC_01 / PO001 / PO01 / PO1."""
    t = str(test).upper()
    m = re.search(r"ANALYTIC[_-]?(\d+)", t)
    if not m:
        m = re.match(r"\s*PO0*(\d+)", t)
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


def _first_num_param(params, default):
    for t in _tokens(params):
        f = _to_float(t)
        if f is not None:
            return f
    return default


def _split_window(params) -> int:
    return int(_first_num_param(params, DEFAULT_SPLIT_WINDOW))


def _gr_days(params) -> int:
    return int(_first_num_param(params, DEFAULT_GR_DAYS))


def _price_threshold(params):
    """Umbral opcional para 06. Vacio = marca cualquier diferencia."""
    return _first_num_param(params, None)


def _has_item(df):
    """Mascara de filas con Item Code real (excluye servicios sin codigo)."""
    it = df.get("Item Code")
    if it is None:
        return pd.Series(False, index=df.index)
    return it.notna() & (it.astype(str).str.strip() != "")


def _load_named_import(row, name):
    fname = f"{row.scope}_{row.module}_{name}_{row.date_to.strftime('%Y%m%d')}.xlsx"
    path = os.path.join(OUTPUT_DIR, fname)
    if os.path.exists(path):
        return pd.read_excel(path)
    alts = glob.glob(os.path.join(OUTPUT_DIR, f"*_{row.module}_{name}_*.xlsx"))
    return pd.read_excel(sorted(alts)[-1]) if alts else None


def _empty_like(df, extra_cols=None):
    out = df.iloc[0:0].copy()
    for c in (extra_cols or []):
        out[c] = pd.Series(dtype="object")
    return out


# ----------------------------------------------------------------------------- split (01 / 09)
def _split(df, group_keys, doc_id_col, date_col, window, key_name="SPLIT_KEY"):
    """Marca documentos del mismo grupo emitidos dentro de 'window' dias entre si.
    Devuelve las filas de detalle de los documentos marcados + SPLIT_KEY + ventana."""
    d = df[_has_item(df)].copy()
    if d.empty:
        return _empty_like(df, [key_name, "Split Window Days"])
    d["__date"] = _date(d[date_col])
    docs = (d.dropna(subset=["__date"])
              .drop_duplicates(group_keys + [doc_id_col])[group_keys + [doc_id_col, "__date"]])
    flagged = []
    for keys, g in docs.groupby(group_keys, sort=False):
        if len(g) < 2:
            continue
        dts = g["__date"].tolist()
        ids = g[doc_id_col].tolist()
        for i in range(len(g)):
            for j in range(len(g)):
                if i != j and abs((dts[i] - dts[j]).days) <= window:
                    flagged.append(tuple(keys if isinstance(keys, tuple) else (keys,)) + (ids[i],))
                    break
    if not flagged:
        return _empty_like(df, [key_name, "Split Window Days"])
    flag_df = pd.DataFrame(flagged, columns=group_keys + [doc_id_col])
    out = df.merge(flag_df, on=group_keys + [doc_id_col], how="inner").copy()
    out[key_name] = out[group_keys].astype(str).agg("|".join, axis=1)
    out["Split Window Days"] = window
    sort_cols = [c for c in (group_keys + [date_col, doc_id_col]) if c in out.columns]
    out = out.sort_values(sort_cols).reset_index(drop=True)
    out.attrs["highlight"] = date_col
    return out


# ----------------------------------------------------------------------------- PO analytics
def _split_pos(df, window):                                   # 01
    return _split(df, ["Company", "Vendor Code", "Item Code", "PO Creator ID"],
                  "PO DocEntry", "PO Doc Date", window, key_name="SPLIT_PO_KEY")


def _duplicate_pos(df):                                        # 02
    d = df[_has_item(df)].copy()
    if d.empty:
        return _empty_like(df, ["DUP_PO_KEY"])
    d["PO Quantity"] = _num(d["PO Quantity"])
    keys = ["Company", "Vendor Code", "Item Code", "PO Quantity"]
    ndocs = d.groupby(keys, sort=False)["PO DocEntry"].transform("nunique")
    out = d[ndocs > 1].copy()
    if out.empty:
        return _empty_like(df, ["DUP_PO_KEY"])
    out["DUP_PO_KEY"] = out[keys].astype(str).agg("|".join, axis=1)
    out = out.sort_values(keys + ["PO Number"]).reset_index(drop=True)
    out.attrs["highlight"] = "PO Quantity"
    return out


def _created_and_approved(df):                                 # 03
    cre, app = df.get("PO Creator ID"), df.get("PO Approver ID")
    mask = (cre.notna() & app.notna()
            & (cre.astype(str).str.strip() == app.astype(str).str.strip())
            & (cre.astype(str).str.strip() != ""))
    out = df[mask].reset_index(drop=True)
    out.attrs["highlight"] = "PO Approver ID"
    return out


def _gr_before_po(df):                                         # 04
    gr = _date(df.get("GR First Posting Date"))
    po = _date(df.get("PO Doc Date"))
    mask = gr.notna() & po.notna() & (gr < po)                 # excluye GR sin fecha
    out = df[mask].copy()
    out["Days GR Before PO"] = (po[mask] - gr[mask]).dt.days
    out = out.reset_index(drop=True)
    out.attrs["highlight"] = "GR First Posting Date"
    return out


def _gr_after_approval(df, days):                              # 05
    gr = _date(df.get("GR Last Posting Date"))
    appr = _date(df.get("PO Approval Date"))
    diff = (gr - appr).dt.days
    mask = gr.notna() & appr.notna() & (diff > days)
    out = df[mask].copy()
    out["Days GR After Approval"] = diff[mask]
    out["Threshold Days"] = days
    out = out.reset_index(drop=True)
    out.attrs["highlight"] = "Days GR After Approval"
    return out


def _price_diff(df, threshold):                                # 06
    d = df[_has_item(df)].copy()
    if d.empty:
        return _empty_like(df, ["Min Unit Price", "Max Unit Price", "Price Difference"])
    d["PO Unit Price"] = _num(d["PO Unit Price"])
    keys = ["Company", "Vendor Code", "Item Code", "PO Doc Currency"]
    grp = d.groupby(keys, sort=False)["PO Unit Price"]
    d["Min Unit Price"] = grp.transform("min")
    d["Max Unit Price"] = grp.transform("max")
    d["Price Difference"] = (d["Max Unit Price"] - d["Min Unit Price"]).round(6)
    if threshold is None:
        out = d[d["Price Difference"] > 0].copy()
    else:
        out = d[d["Price Difference"] > float(threshold)].copy()
    if out.empty:
        return _empty_like(df, ["Min Unit Price", "Max Unit Price", "Price Difference"])
    out = out.sort_values(keys + ["PO Unit Price"]).reset_index(drop=True)
    out.attrs["highlight"] = "Price Difference"
    return out


def _same_user_po_gr(df):                                      # 07
    cre, grc = df.get("PO Creator ID"), df.get("GR Creator ID")
    mask = (cre.notna() & grc.notna()
            & (cre.astype(str).str.strip() == grc.astype(str).str.strip())
            & (cre.astype(str).str.strip() != ""))
    out = df[mask].reset_index(drop=True)
    out.attrs["highlight"] = "GR Creator ID"
    return out


def _po_by_month(df):                                          # 08 (resumen)
    d = df.copy()
    d["PO Line Total"] = _num(d.get("PO Line Total"))
    d["PO Line Total USD"] = _num(d.get("PO Line Total USD"))
    d["PO Quantity"] = _num(d.get("PO Quantity"))
    g = (d.groupby(["Company", "PO Month", "Item Code"], as_index=False, sort=False)
           .agg(**{"PO Lines": ("PO DocEntry", "count"),
                   "Distinct POs": ("PO DocEntry", "nunique"),
                   "Total Quantity": ("PO Quantity", "sum"),
                   "Total Line Amount": ("PO Line Total", "sum"),
                   "Total Line Amount USD": ("PO Line Total USD", "sum")}))
    g = g.sort_values(["Company", "PO Month", "Item Code"]).reset_index(drop=True)
    g.attrs["highlight"] = "Total Line Amount USD"
    return g


def _pos_without_pr(df):                                       # 11
    fp = df.get("From PR")
    mask = fp.isna() | (fp.astype(str).str.strip().str.upper() != "Y")
    out = df[mask].reset_index(drop=True)
    out.attrs["highlight"] = "From PR"
    return out


# ----------------------------------------------------------------------------- PR analytics
def _split_prs(pr, window):                                    # 09
    out = _split(pr, ["Company", "Item Code", "PR Creator ID"],
                 "PR DocEntry", "PR Doc Date", window, key_name="SPLIT_PR_KEY")
    return out


def _po_vs_pr(pr, po):                                         # 10
    cols_po = ["Company", "PR DocEntry", "PR Line", "PO Number", "PO Line",
               "Item Code", "PO Quantity", "PO Material Description", "Vendor Name"]
    left = po[po.get("From PR").astype(str).str.upper() == "Y"][
        [c for c in cols_po if c in po.columns]].copy()
    cols_pr = ["Company", "PR DocEntry", "PR Line", "PR Number",
               "Item Code", "PR Quantity"]
    right = pr[[c for c in cols_pr if c in pr.columns]].copy()
    if left.empty or right.empty:
        return pd.DataFrame(columns=["Company", "PR Number", "PR Line", "PO Number", "PO Line",
                                     "PR Item", "PO Item", "Item Match",
                                     "PR Quantity", "PO Quantity", "Qty Difference",
                                     "Vendor Name", "PO Material Description"])
    m = left.merge(right, on=["Company", "PR DocEntry", "PR Line"],
                   how="inner", suffixes=(" PO", " PR"))
    m["PR Item"] = m["Item Code PR"].astype(str).str.strip()
    m["PO Item"] = m["Item Code PO"].astype(str).str.strip()
    m["Item Match"] = (m["PR Item"] == m["PO Item"]).map({True: "Y", False: "N"})
    m["Qty Difference"] = (_num(m["PO Quantity"]) - _num(m["PR Quantity"])).round(6)
    flagged = m[(m["Item Match"] == "N") | (m["Qty Difference"] != 0)].copy()
    out = flagged[["Company", "PR Number", "PR Line", "PO Number", "PO Line",
                   "PR Item", "PO Item", "Item Match",
                   "PR Quantity", "PO Quantity", "Qty Difference",
                   "Vendor Name", "PO Material Description"]].reset_index(drop=True)
    out.attrs["highlight"] = "Qty Difference"
    return out


# ----------------------------------------------------------------------------- dispatcher
def analyze(row, test, params) -> pd.DataFrame:
    n = _analytic_no(test)

    if n == 9:                                            # solo pr_lines
        pr = _load_named_import(row, "pr_lines")
        if pr is None:
            return pd.DataFrame([{"PR Number": "(falta el import pr_lines -- corre run_import --only PO)"}])
        return _split_prs(pr, _split_window(params))

    if n == 10:                                           # pr_lines + po_lines
        pr = _load_named_import(row, "pr_lines")
        po = load_import_df(row)
        if pr is None:
            return pd.DataFrame([{"PR Number": "(falta el import pr_lines -- corre run_import --only PO)"}])
        return _po_vs_pr(pr, po)

    df = load_import_df(row)                               # base po_lines (01..08, 11)

    if n == 1:
        return _split_pos(df, _split_window(params))
    if n == 2:
        return _duplicate_pos(df)
    if n == 3:
        return _created_and_approved(df)
    if n == 4:
        return _gr_before_po(df)
    if n == 5:
        return _gr_after_approval(df, _gr_days(params))
    if n == 6:
        return _price_diff(df, _price_threshold(params))
    if n == 7:
        return _same_user_po_gr(df)
    if n == 8:
        return _po_by_month(df)
    if n == 11:
        return _pos_without_pr(df)

    return pd.DataFrame([{"Test": test, "Resultado": "sin logica definida"}])
