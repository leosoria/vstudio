"""
modules/VM/analysis.py  --  Analizador del modulo VM (Vendor Management).

Contrato comun: analyze(row, test, params) -> DataFrame (una hoja del Results),
igual que modules/AR/analysis.py y modules/FAM/analysis.py.

Los analiticos corren sobre la UNION de todas las companias (poblacion cross-company),
por eso muestran Company/CoCo: un grupo que cruza companias suele ser la misma entidad
Logicalis en otra BU (no un duplicado real) -- queda a criterio del auditor.

Imports del modulo (los baja run_import; quedan en output/ como
{scope}_{modulo}_{import}_{AAAAMMDD}.xlsx):
  - vm_vendors       (primario)  OCRD CardType='S': identidad, tax id, telefonos, banco maestro, CBU
  - vm_addresses     (extra)     CRD1 de proveedores (direcciones B y S)
  - vm_banks         (extra)     OCRB de proveedores (cuentas bancarias)
  - vm_bank_changes  (extra)     historial de cambios bancarios (ACRB diff) -> VM07

Analiticos:
  VM01  nombres similares      (fuzzy = Levenshtein <= max_dist; PARAM1, default 2, criterio ACL)
  VM02  misma direccion        (B y S)
  VM03  mismo telefono
  VM04  misma cuenta bancaria  (clave Country|BankCode|Account; + CBU como clave alternativa)
  VM05  mismo Tax/Business Number (LicTradNum)
  VM06  = VM05 (en SAP B1 tax number y business number son el mismo campo LicTradNum)
  VM07  cambios frecuentes de banco (> N en el periodo; PARAM1 default 2)
  VM08..VM11  vendor vs empleado  -> NO APLICA (sin maestro de empleados OHEM en B1)
  VM12  sin business number     (LicTradNum vacio)
  VM13  sin tax number          (= VM12, mismo campo)
  VM14  telefono en blanco
  VM15  datos bancarios en blanco
  VM16  solo casilla postal (PO Box) y sin direccion fisica
  VM17  single poster           -> PENDIENTE (requiere transacciones)
  VM_HIT_MATRIX  matriz de excepciones por proveedor
"""
import os
import re
import glob
import unicodedata
import pandas as pd
from core.analysis_base import load_import_df, OUTPUT_DIR

try:
    from rapidfuzz import fuzz, process
    from rapidfuzz.distance import Levenshtein as _RFLev
    import numpy as np
    _HAS_RF = True
except ImportError:                       # fallback sin dependencia
    import difflib
    _HAS_RF = False

# --- listas configurables -------------------------------------------------
# Sufijos de razon social a remover del final del nombre antes del fuzzy (VM01).
# Tokens unicos (el nombre ya viene en MAYUSCULAS, sin acentos ni puntuacion);
# las formas compuestas (p.ej. "S A", "S DE R L DE C V", "PTY LTD") se resuelven
# quitando token por token desde el final. Lista multi-pais portada del ACL.
COMPANY_SUFFIX_TOKENS = {
    # genericos / ingles
    "LTD", "LTDA", "LIMITED", "LIMITADA", "LDA", "INC", "INCORPORATED",
    "INCORPORATION", "CORP", "CORPORATION", "CO", "COMPANY", "COMPANIA", "CIA",
    "LLC", "LLP", "LLLP", "LP", "GP", "SP", "PLC", "PC", "PA", "PLLC", "PSC",
    "TRUST", "FUND", "FOUNDATION", "FOUNDATIONS", "HOLDING", "HOLDINGS", "GROUP",
    "GROUPE", "NGO", "NPO", "ASSOCIATION", "INSTITUTE", "INSTITUTES", "SOCIETY",
    "UNION", "SYNDICATE", "COOP", "COOPERATIVE", "COOPERATIVA", "POOL",
    "NATIONAL", "FEDERAL", "INDUSTRIES", "IND", "BANK", "BANKERS", "CLUB",
    "INTERNATIONAL", "INTL", "UNLTD", "ULTD", "NL", "NO", "LIABILITY",
    "PVT", "PTE", "PTY", "BK", "CC", "LC", "SMLLC", "CIC", "CIO", "CCC",
    "PRIVATE", "PROPRIETARY", "PROFESSIONAL", "REGISTERED", "PARTNERSHIP",
    "JOINT", "VENTURE", "VENTURES", "JV", "OF", "ON", "AT",
    # espanol (LatAm)
    "SA", "SAC", "SACI", "SAICF", "SACIF", "SAS", "SCA", "SRL", "SR",
    "EIRL", "SAPI", "SAB", "SAD", "SAL", "SGR", "SC", "SCP", "SCS", "SCCL",
    "SCOP", "SL", "SLL", "SLNE", "SCRA", "SOC", "SOCIEDAD", "SOCIEDADES",
    "ANONIMA", "LIMITADAS", "COLECTIVA", "COMANDITARIA", "COMANDITA",
    "RESPONSABILIDAD", "RECIPROCA", "GARANTIA", "ACCIONES", "SIMPLIFICADA",
    "UNIPERSONAL", "EMPRESA", "EMPRESAS", "SUCESORES", "SUC", "EU", "Y",
    # portugues (BR/PT)
    "SGPS", "EIRELI", "ME", "EPP", "MEI", "ABERTA", "FECHADA", "SF",
    # italiano
    "SPA", "SAPA", "SNC", "SAA",
    # aleman
    "GMBH", "AG", "KG", "KGAA", "OHG", "GBR", "EG", "UG", "MBH",
    # frances
    "SARL", "SARLU", "EURL", "SASU", "SCI", "SCOP", "SEM", "GIE", "SEP", "FCP",
    "SICAV", "SCE", "EEIG",
    # holandes / belga / nordico / otros
    "BV", "NV", "CV", "VOF", "AB", "OY", "AS", "ASA", "APS", "SE", "SARF",
    "SDN", "BHD", "KK", "YK", "ZAO", "OAO", "OOO", "PAO", "TEO", "TEORANTA",
    "OYJ", "PLLC", "GIE",
    # tokens sueltos que arman formas compuestas (S A, C V, S DE R L DE C V, etc.)
    "S", "A", "C", "V", "L", "R", "E", "U", "I", "P", "G", "F", "B", "T",
    "K", "N", "D", "M", "DE", "EN", "RL", "EU", "DEL", "LA", "EL", "DA",
}
# patrones que indican casilla postal (VM16)
POBOX_PATTERNS = [
    r"\bP\.?\s*O\.?\s*BOX\b", r"\bPOST\s*OFFICE\s*BOX\b",
    r"\bCASILLA\b", r"\bAPARTADO\b", r"\bAP\.?\s*POSTAL\b", r"\bBOX\b",
]
_POBOX_RE = re.compile("|".join(POBOX_PATTERNS), re.IGNORECASE)


# --- carga de imports ------------------------------------------------------
def _load_named(row, name):
    fname = f"{row.scope}_{row.module}_{name}_{row.date_to.strftime('%Y%m%d')}.xlsx"
    path = os.path.join(OUTPUT_DIR, fname)
    if os.path.exists(path):
        return pd.read_excel(path)
    alts = glob.glob(os.path.join(OUTPUT_DIR, f"*_{row.module}_{name}_*.xlsx"))
    return pd.read_excel(sorted(alts)[-1]) if alts else None


def _vendors(row):
    """Maestro de proveedores. Las exclusiones (CardType='S', activos,
    empleados E*/T*, intercompany) se aplican en ORIGEN en vm_vendors.sql,
    asi que aca solo se carga el import ya filtrado."""
    try:
        df = load_import_df(row)
    except Exception:
        df = _load_named(row, "vm_vendors")
        if df is None:
            raise
    return df.reset_index(drop=True)


# --- normalizaciones -------------------------------------------------------
def _strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(c))


def _clean_name(s) -> str:
    """Nombre limpio para fuzzy: mayusculas, sin acentos/puntuacion, sin sufijos."""
    if s is None:
        return ""
    s = _strip_accents(str(s)).upper()
    s = re.sub(r"[^A-Z0-9]+", " ", s).strip()
    toks = s.split()
    while len(toks) > 1 and toks[-1] in COMPANY_SUFFIX_TOKENS:   # deja >=1 token
        toks.pop()
    return " ".join(toks)


def _norm(s) -> str:
    """Normaliza una clave exacta: alfanumerico en mayusculas (sin separadores)."""
    if s is None:
        return ""
    return re.sub(r"[^A-Z0-9]+", "", _strip_accents(str(s)).upper())


def _digits(s) -> str:
    if s is None:
        return ""
    return re.sub(r"\D+", "", str(s))


def _blank(series) -> pd.Series:
    v = series.astype("string").str.strip()
    return v.isna() | v.eq("")


def _vid_cols(df):
    return ["CoCo", "Vendor Code"] if "CoCo" in df.columns else ["Company", "Vendor Code"]


def _vid(df):
    cols = _vid_cols(df)
    return df[cols[0]].astype("string").fillna("") + "|" + df[cols[1]].astype("string").fillna("")


# --- agrupacion ------------------------------------------------------------
class _UF:
    def __init__(self, n): self.p = list(range(n))
    def find(self, x):
        while self.p[x] != x:
            self.p[x] = self.p[self.p[x]]; x = self.p[x]
        return x
    def union(self, a, b): self.p[self.find(a)] = self.find(b)


def _groups_from_pairs(n, pairs):
    """Componentes conectados -> lista de group_id (0 = sin grupo)."""
    uf = _UF(n)
    for a, b in pairs:
        uf.union(a, b)
    roots = {}
    size = {}
    for i in range(n):
        r = uf.find(i)
        size[r] = size.get(r, 0) + 1
    gid_of_root = {}
    nxt = 1
    out = [0] * n
    for i in range(n):
        r = uf.find(i)
        if size[r] < 2:
            continue
        if r not in gid_of_root:
            gid_of_root[r] = nxt; nxt += 1
        out[i] = gid_of_root[r]
    return out


def _bounded_lev(a, b, maxd):
    """Distancia de Levenshtein con corte: devuelve >maxd si supera el umbral."""
    la, lb = len(a), len(b)
    if abs(la - lb) > maxd:
        return maxd + 1
    if la == 0:
        return lb
    if lb == 0:
        return la
    prev = list(range(lb + 1))
    for i in range(1, la + 1):
        cur = [i] + [0] * lb
        best = cur[0]
        ca = a[i - 1]
        for j in range(1, lb + 1):
            cost = 0 if ca == b[j - 1] else 1
            cur[j] = min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost)
            if cur[j] < best:
                best = cur[j]
        if best > maxd:            # toda la fila ya supera el umbral -> corte
            return maxd + 1
        prev = cur
    return prev[lb]


def _fuzzy_pairs(names, max_dist=2, min_len=3, block_keys=None):
    """Pares (i,j) cuyos nombres limpios tienen distancia de Levenshtein <= max_dist
    (criterio del ACL: V_LEVDIST=2). Solo compara DENTRO del mismo bloque:
    si block_keys viene dado (p.ej. la compania), agrupa por (compania, inicial),
    de modo que un mismo proveedor replicado en varias entidades NO forma grupo.
    Ignora nombres de menos de min_len."""
    blocks = {}
    for i, nm in enumerate(names):
        if nm and len(nm) >= min_len:
            bk = (block_keys[i] if block_keys is not None else "", nm[0])
            blocks.setdefault(bk, []).append(i)
    pairs = []
    for _, idxs in blocks.items():
        if len(idxs) < 2:
            continue
        sub = [names[i] for i in idxs]
        if _HAS_RF:
            mat = process.cdist(sub, sub, scorer=_RFLev.distance,
                                workers=-1, dtype=np.int32)
            for a in range(len(idxs)):
                rowv = mat[a]
                for b in range(a + 1, len(idxs)):
                    if rowv[b] <= max_dist:        # incluye distancia 0 (idénticos)
                        pairs.append((idxs[a], idxs[b]))
        else:
            for a in range(len(idxs)):
                for b in range(a + 1, len(idxs)):
                    if _bounded_lev(sub[a], sub[b], max_dist) <= max_dist:
                        pairs.append((idxs[a], idxs[b]))
    return pairs


# columnas extra del maestro a arrastrar en cada analitico (si existen en el import)
_VENDOR_DISPLAY = ["Vendor Name", "Last Invoice Number",
                   "Last Transaction Date", "Last Inv Amt Doc Currency",
                   "Last Inv Amt Doc Currency Indicator"]


def _group_shared(df, keyseries, keep_cols, within_company=True):
    """Filas que comparten una clave no vacia entre >=2 proveedores distintos
    DENTRO de la misma compania (within_company=True: la clave se prefija con CoCo,
    asi un mismo proveedor replicado en varias entidades no se marca como grupo).
    Devuelve esas filas con columna 'Group' (desde 1), ordenado por grupo."""
    d = df.copy()
    raw = keyseries.astype("string").fillna("")
    # clave "vacia" = sin contenido real (ej. direccion toda en blanco -> '|||')
    content = raw.str.replace("|", "", regex=False).str.replace(r"\s+", "", regex=True)
    d["_vid"] = _vid(d).values
    if within_company and "CoCo" in d.columns:
        d["_key"] = (d["CoCo"].astype("string").fillna("") + "\u00a6" + raw.values)
    else:
        d["_key"] = raw.values
    d = d[content.values != ""]
    if d.empty:
        return pd.DataFrame(columns=keep_cols + ["Group"])
    distinct = d.groupby("_key")["_vid"].nunique()
    shared = set(distinct[distinct >= 2].index)
    d = d[d["_key"].isin(shared)]
    if d.empty:
        return pd.DataFrame(columns=keep_cols + ["Group"])
    order = {k: i + 1 for i, k in enumerate(sorted(shared))}
    d["Group"] = d["_key"].map(order)
    d = d.sort_values(["Group"] + _vid_cols(d), kind="stable")
    return d[[c for c in keep_cols if c in d.columns] + ["Group"]].reset_index(drop=True)


def _one_address(addr):
    """Una sola direccion por proveedor: prefiere Bill To (B) si tiene datos,
    si no Ship To (S). Evita duplicar cada proveedor (B y S) en VM02."""
    a = addr.copy()
    parts = pd.Series([""] * len(a), index=a.index, dtype="string")
    for c in ("Street", "City", "ZipCode", "Country"):
        if c in a.columns:
            parts = parts + a[c].astype("string").fillna("").str.strip()
    a["_content"] = (parts.str.len() > 0).astype(int)          # 1 = tiene datos
    at = a["Address Type"].astype("string").fillna("").str.upper().str.strip()
    a["_prio"] = at.map({"B": 0, "S": 1}).fillna(2).astype(int)  # B antes que S
    a = a.sort_values(["_content", "_prio"], ascending=[False, True], kind="stable")
    a = a.drop_duplicates(_vid_cols(a), keep="first")            # 1 fila por proveedor
    return a.drop(columns=["_content", "_prio"]).reset_index(drop=True)


def _merge_name(df, vendors):
    """Trae del maestro la etiqueta Company y las columnas de display
    (Vendor Name, ultima factura, moneda, etc.) por proveedor."""
    vc = _vid_cols(vendors)
    on = _vid_cols(df)
    if on != vc:
        return df
    cols = [c for c in (["Company"] + _VENDOR_DISPLAY)
            if c in vendors.columns and c not in df.columns]
    if not cols:
        return df
    name = vendors[vc + cols].drop_duplicates(vc)
    return df.merge(name, on=vc, how="left")


# --- clave bancaria unificada (VM04 / VM15) --------------------------------
def _bank_long(row):
    """Lista larga de cuentas bancarias DEL PROVEEDOR (no de la compania), de tres
    fuentes a nivel proveedor: OCRB (cuentas bancarias del SN), OCRD Dfl* (cuenta
    por defecto; p.ej. Chile usa DflAccount) y CBU (U_ClaveBancaria, Arg/Uy).
    NO usa los campos House* de OCRD: esos son el banco propio de la compania
    (mismo para todos los proveedores), no la cuenta del proveedor."""
    vendors = _vendors(row)            # ya viene solo activos
    banks = _load_named(row, "vm_banks")
    vc = _vid_cols(vendors)
    active_ids = set(_vid(vendors))
    frames = []

    # 1) OCRB: Country | BankCode | Account  (solo proveedores activos)
    if banks is not None and len(banks):
        b = banks.copy()
        acct = b["Account"].astype("string").fillna("")
        key = (b["Country"].map(_norm) + "|" + b["Bank Code"].map(_norm)
               + "|" + b["Account"].map(_norm))
        b = b.assign(**{"Bank Key": key.values, "Source": "OCRB",
                        "Bank Detail": acct.values})
        b = b[acct.str.strip().ne("")]
        b = b[_vid(b).isin(active_ids)]        # respeta el filtro de activos del maestro
        bvc = _vid_cols(b)
        frames.append(b[bvc + ["Bank Key", "Source", "Bank Detail"]])

    # 2) OCRD Dfl* (cuenta por defecto del proveedor; p.ej. Chile usa DflAccount)
    if "Default Account" in vendors.columns:
        dv = vendors.copy()
        dacct = dv["Default Account"].astype("string").fillna("")
        dbank = (dv["Default Bank"] if "Default Bank" in dv.columns
                 else pd.Series([""] * len(dv), index=dv.index))
        dbrn = (dv["Default Branch"] if "Default Branch" in dv.columns
                else pd.Series([""] * len(dv), index=dv.index))
        dkey = (dbank.map(_norm) + "|" + dbrn.map(_norm) + "|" + dacct.map(_norm))
        d = dv.assign(**{"Bank Key": dkey.values, "Source": "OCRD_Dfl",
                         "Bank Detail": dacct.values})
        d = d[dacct.str.strip().ne("")]
        frames.append(d[vc + ["Bank Key", "Source", "Bank Detail"]])

    # 3) CBU del proveedor (U_ClaveBancaria, solo donde exista): solo digitos
    if "CBU" in vendors.columns:
        cv = vendors.copy()
        cbu = cv["CBU"].map(_digits)
        c = cv.assign(**{"Bank Key": cbu.values, "Source": "U_CBU",
                         "Bank Detail": cbu.values})
        c = c[cbu.str.len() > 0]
        frames.append(c[vc + ["Bank Key", "Source", "Bank Detail"]])

    if not frames:
        return pd.DataFrame(columns=vc + ["Bank Key", "Source", "Bank Detail"])
    return pd.concat(frames, ignore_index=True)


# --- params ----------------------------------------------------------------
def _param(params, names, default):
    if isinstance(params, dict):
        for k in names:
            if params.get(k) not in (None, ""):
                return params[k]
    elif isinstance(params, (list, tuple)) and params:
        if params[0] not in (None, ""):
            return params[0]
    elif isinstance(params, str) and params.strip():
        return params.strip()
    return default


def _note(msg):
    return pd.DataFrame([{"Resultado": msg}])


def _canon(test) -> str:
    """Normaliza el codigo de test a la forma canonica 'VMnn' (o 'HIT_MATRIX'),
    aceptando cualquier convencion: VM001, VM01, VM1, VM_ANALYTIC_07_VMCS119,
    VM_HIT_MATRIX, etc. Asi el ruteo no depende de como los liste el config."""
    t = (test or "").upper().strip()
    if "MATRIX" in t or "HIT" in t:
        return "HIT_MATRIX"
    m = re.search(r"ANALYTIC[_\-\s]*0*(\d+)", t)      # VM_ANALYTIC_07_... -> 7
    if not m:
        m = re.search(r"VM[_\-\s]*0*(\d+)", t)         # VM001 / VM01 / VM1 -> n
    if m:
        return "VM%02d" % int(m.group(1))
    return t


# --- analiticos ------------------------------------------------------------
def analyze(row, test, params) -> pd.DataFrame:
    t = _canon(test)

    if t in ("VM_ANALYTIC_08_VMCS122", "VM_ANALYTIC_09_VMCS123",
             "VM_ANALYTIC_10_VMCS124", "VM_ANALYTIC_11_VMCS117") or t in (
             "VM08", "VM09", "VM10", "VM11"):
        return _note("NO APLICA: SAP B1 no tiene maestro de empleados (OHEM) poblado; "
                     "el cruce vendor vs empleado no es ejecutable.")

    if t in ("VM17", "VM_ANALYTIC_17_VMCS612"):
        return _note("PENDIENTE: requiere base de transacciones (pagos/facturas) para "
                     "identificar el usuario que postea. Se arma despues de los de maestro.")

    if t in ("VM06", "VM_ANALYTIC_06_VMCS301"):
        return _note("Control duplicado de VM05: en SAP B1 'tax number' y 'business number' "
                     "son el mismo campo (OCRD.LicTradNum). Ver resultados de VM05.")

    # VM07 corre sobre su propio import (historial)
    if t in ("VM07", "VM_ANALYTIC_07_VMCS119"):
        chg = _load_named(row, "vm_bank_changes")
        if chg is None or chg.empty:
            return _note("Sin cambios de datos bancarios en el periodo (o falta el import vm_bank_changes).")
        minc = int(float(_param(params, ("min_changes", "umbral", "n"), 2)))
        if "Created" in chg.columns:           # excluye la primera carga (NULL->valor)
            chg = chg[pd.to_numeric(chg["Created"], errors="coerce").fillna(0) != 1]
        if chg.empty:
            return _note("Sin MODIFICACIONES de datos bancarios en el periodo (solo cargas iniciales).")
        vc = _vid_cols(chg)
        gcols = vc + (["Vendor Name"] if "Vendor Name" in chg.columns else [])
        if "Company" in chg.columns and "Company" not in gcols:   # etiqueta de la BU
            gcols = ["Company"] + gcols
        if "Change Doc" in chg.columns:        # cuenta EVENTOS (no filas por campo)
            cnt = chg.groupby(gcols).agg(**{"Cambios": ("Change Doc", "nunique")}).reset_index()
        else:
            cnt = chg.groupby(gcols).size().reset_index(name="Cambios")
        out = cnt[cnt["Cambios"] >= minc].sort_values("Cambios", ascending=False)
        out.attrs["highlight"] = "Cambios"
        return out.reset_index(drop=True)

    vendors = _vendors(row)
    vc = _vid_cols(vendors)
    ids = (["Company"] if "Company" in vendors.columns else []) + vc
    extras = [c for c in _VENDOR_DISPLAY
              if c != "Vendor Name" and c in vendors.columns]
    keep = ids + ["Vendor Name"]

    # VM01: nombres similares DENTRO de la misma compania
    #       (fuzzy = distancia de Levenshtein <= max_dist; ACL V_LEVDIST=2)
    if t in ("VM01", "VM_ANALYTIC_01_VMCS101"):
        maxd = int(float(_param(params, ("max_dist", "dist", "umbral", "threshold"), 2)))
        clean = vendors["Vendor Name"].map(_clean_name).tolist()
        bkey = (vendors["CoCo"] if "CoCo" in vendors.columns
                else vendors["Company"]).astype("string").fillna("").tolist()
        gids = _groups_from_pairs(len(clean), _fuzzy_pairs(clean, maxd, block_keys=bkey))
        out = vendors.assign(**{"Clean Name": clean, "Group": gids})
        out = out[out["Group"] > 0].sort_values(["Group"] + vc, kind="stable")
        cols = [c for c in keep + ["Clean Name", "Group"] + extras if c in out.columns]
        return out[cols].reset_index(drop=True)

    # VM02: misma direccion (1 por proveedor: Bill To si tiene datos, si no Ship To)
    if t in ("VM02", "VM_ANALYTIC_02_VMCS103"):
        addr = _load_named(row, "vm_addresses")
        if addr is None or addr.empty:
            return _note("Sin direcciones (falta el import vm_addresses).")
        addr = addr[_vid(addr).isin(set(_vid(vendors)))]   # solo proveedores validos
        if addr.empty:
            return pd.DataFrame(columns=["CoCo", "Vendor Code", "Street", "City", "Group"])
        addr = _one_address(addr)                          # 1 direccion por proveedor
        key = (addr["Street"].map(_norm) + "|" + addr["City"].map(_norm) + "|"
               + addr["ZipCode"].map(_norm) + "|" + addr["Country"].map(_norm))
        kc = ["CoCo", "Vendor Code", "Address Type", "Street", "City",
              "ZipCode", "State", "Country"]
        kc = [c for c in kc if c in addr.columns]
        out = _group_shared(addr, key, kc)
        return _merge_name(out, vendors)

    # VM03: mismo telefono (dentro de la misma compania)
    if t in ("VM03", "VM_ANALYTIC_03_VMCS125"):
        key = vendors["Phone1"].map(_digits)
        return _group_shared(vendors, key, keep + ["Phone1"] + extras)

    # VM04: misma cuenta bancaria (clave unificada)
    if t in ("VM04", "VM_ANALYTIC_04_VMCS118"):
        bl = _bank_long(row)
        if bl.empty:
            return _note("Sin datos bancarios en ninguna fuente (OCRB/House*/CBU).")
        out = _group_shared(bl, bl["Bank Key"],
                            _vid_cols(bl) + ["Bank Key", "Source", "Bank Detail"])
        return _merge_name(out, vendors)

    # VM05: mismo Tax/Business Number (dentro de la misma compania)
    if t in ("VM05", "VM_ANALYTIC_05_VMCS126"):
        key = vendors["Tax/Business Number"].map(_norm)
        return _group_shared(vendors, key, keep + ["Tax/Business Number"] + extras)

    # VM12 / VM13: sin business / tax number (mismo campo)
    if t in ("VM12", "VM13", "VM_ANALYTIC_12_VMCS120", "VM_ANALYTIC_13_VMCS121"):
        out = vendors[_blank(vendors["Tax/Business Number"])]
        return out[[c for c in keep + ["Tax/Business Number"] + extras if c in out.columns]].reset_index(drop=True)

    # VM14: telefono en blanco
    if t in ("VM14", "VM_ANALYTIC_14_VMCS515"):
        out = vendors[_blank(vendors["Phone1"])]
        return out[[c for c in keep + ["Phone1"] + extras if c in out.columns]].reset_index(drop=True)

    # VM15: sin datos bancarios en ninguna fuente
    if t in ("VM15", "VM_ANALYTIC_15_VMCS516"):
        bl = _bank_long(row)
        with_bank = set(_vid(bl)) if not bl.empty else set()
        out = vendors[~_vid(vendors).isin(with_bank)]
        return out[[c for c in keep + extras if c in out.columns]].reset_index(drop=True)

    # VM16: solo PO Box y sin direccion fisica
    if t in ("VM16", "VM_ANALYTIC_16_VMCS102"):
        addr = _load_named(row, "vm_addresses")
        if addr is None or addr.empty:
            return _note("Sin direcciones (falta el import vm_addresses).")
        addr = addr[_vid(addr).isin(set(_vid(vendors)))]   # solo activos
        if addr.empty:
            return _note("Sin direcciones de proveedores activos.")
        a = addr.copy()
        txt = (a["Street"].astype("string").fillna("") + " "
               + a.get("Building", pd.Series([""] * len(a))).astype("string").fillna(""))
        a["_pobox"] = txt.map(lambda s: bool(_POBOX_RE.search(s or "")))
        street = a["Street"].astype("string").str.strip()
        a["_physical"] = street.notna() & street.ne("") & (~a["_pobox"])
        vc2 = _vid_cols(a)
        agg = a.groupby(vc2).agg(pobox=("_pobox", "any"),
                                 physical=("_physical", "any")).reset_index()
        only_box = agg[agg["pobox"] & (~agg["physical"])]
        kc = vc2 + ["Address Type", "Street", "Building", "City", "Country"]
        kc = [c for c in kc if c in a.columns]
        out = (a[a["_pobox"]].merge(only_box[vc2], on=vc2, how="inner")[kc]
               .drop_duplicates().reset_index(drop=True))
        return _merge_name(out, vendors)

    # VM_HIT_MATRIX
    if t in ("VM_HIT_MATRIX", "HIT_MATRIX"):
        tests = {"VM01": "VM01", "VM02": "VM02", "VM03": "VM03", "VM04": "VM04",
                 "VM05": "VM05", "VM07": "VM07", "VM12": "VM12", "VM14": "VM14",
                 "VM15": "VM15", "VM16": "VM16"}
        base = vendors[vc + ["Vendor Name"]].drop_duplicates(vc).copy()
        base["_vid"] = _vid(base).values
        for label in tests:
            try:
                res = analyze(row, label, {})
            except Exception:
                res = None
            hit = set()
            if res is not None and not res.empty and "Vendor Code" in res.columns:
                rc = ["CoCo", "Vendor Code"] if "CoCo" in res.columns else ["Company", "Vendor Code"]
                if all(c in res.columns for c in rc):
                    hit = set(res[rc[0]].astype("string").fillna("") + "|"
                              + res[rc[1]].astype("string").fillna(""))
            base[label] = base["_vid"].isin(hit).astype(int)
        base["Total Hits"] = base[list(tests)].sum(axis=1)
        out = base[base["Total Hits"] > 0].sort_values("Total Hits", ascending=False)
        return out[[c for c in (vc + ["Vendor Name"] + list(tests) + ["Total Hits"])
                    if c in out.columns]].reset_index(drop=True)

    return _note(f"{test}: sin logica definida")
