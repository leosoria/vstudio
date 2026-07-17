"""
core/registry.py
Define, por modulo:
  - "import": la EXTRACCION del modulo (query + params que vienen del config).
              Corre siempre que el modulo este en EXECUTE=T en la hoja config.
  - "imports_extra": import(es) ADICIONAL(es) del modulo (opcional). Mismo formato que
              "import". Sirve para modulos que necesitan mas de una bajada (ej. AR:
              ar_aging para AR001..AR005 y credit_limit_changes para AR006).
              run_import debe correrlos ademas del primario (ver mini-loop).
  - "tests" : los X tests del modulo. Cada uno con sus param_names (que mapean a
              las columnas PARAM de la hoja del modulo) y sus defaults.

Dos niveles de parametros:
  - los del config  -> alimentan el IMPORT      (ej. FAM: PARAM1 -> area)
  - los de la hoja  -> alimentan cada TEST      (por columna, posicional)

Agregar modulo/test = agregar entrada aca. No se tocan los runners.
"""
import os

BASE = os.path.dirname(os.path.dirname(__file__))


def q(module, name):
    return os.path.join(BASE, "modules", module, "queries", f"{name}.sql")


REGISTRY = {
    "FAM": {
        "import": {
            "name": "cuadro_af",
            "sql": q("FAM", "cuadro_af"),
            "param_names": ["area"],        # config PARAM1 -> area
            "defaults": {"area": "100"},
        },
        "tests": {
            "FAM001": {"param_names": [], "defaults": {}},  # Cuadro AF (listado)
            "FAM002": {"param_names": [], "defaults": {}},  # Fixed Assets With No Depreciation
            "FAM003": {"param_names": [], "defaults": {}},  # Fixed Assets With Zero Book Value
            "FAM004": {"param_names": [], "defaults": {}},  # comparacion vs Reporting Pack
        },
    },
    "AR": {
        # import PRIMARIO: alimenta AR001..AR005 (lo toma load_import_df por su 'name').
        "import": {
            "name": "ar_aging",
            "sql": q("AR", "ar_aging"),
            "param_names": [],              # ar_aging usa la fecha de corte (TO); sin PARAM1/2
            "defaults": {},
        },
        # import ADICIONAL: alimenta AR006. Requiere el mini-loop en run_import.
        "imports_extra": [
            {
                "name": "credit_limit_changes",
                "sql": q("AR", "credit_limit_changes"),
                "param_names": [],          # usa FROM/TO
                "defaults": {},
            },
        ],
        "tests": {
            "AR001": {"param_names": [], "defaults": {}},                   # detalle AR (sin BP Currency)
            "AR002": {"param_names": [], "defaults": {}},                   # resumen + Differences
            "AR003": {"param_names": [], "defaults": {}},                   # OUSTANDING BALANCE < 0
            "AR004": {"param_names": ["dpd"], "defaults": {"dpd": "120"}},  # Max DaysPastDue > dpd
            "AR005": {"param_names": [], "defaults": {}},                   # limite de credito 0
            "AR006": {"param_names": [], "defaults": {}},                   # cambios de limite (copia)
        },
    },
    "CD": {
        "import": {
            "name": "cd_main",                 # poblacion de pagos a proveedores (OVPM, CardType=S)
            "sql": q("CD", "cd_main"),
            "param_names": [],                 # solo usa FROM/TO
            "defaults": {},
        },
        "tests": {
            "CD001": {"param_names": [], "defaults": {}},                       # Cash Disbursements By Vendor (poblacion)
            "CD002": {"param_names": [], "defaults": {}},                       # Summary By Vendor
            "CD003": {"param_names": ["dup_window"], "defaults": {}},           # Duplicate (PARAM1 = dias, vacio=30)
            "CD004": {"param_names": ["method", "x"], "defaults": {}},          # PARAM1=STATISTICAL|FIXED|TOPN (vacio=FIXED), PARAM2=N
        },
    },
    "PO": {
        # import PRIMARIO: base de lineas de PO + recepcion (GR) + aprobacion.
        # Alimenta PO_ANALYTIC_01..08 (load_import_df lo toma por su 'name').
        "import": {
            "name": "po_lines",
            "sql": q("PO", "po_lines"),
            "param_names": [],             # usa FROM/TO (filtra OPOR.DocDate)
            "defaults": {},
        },
        # import ADICIONAL: base de Solicitudes de Compra (PR).
        # Alimenta PO_ANALYTIC_09 (Split PR) y aporta el proveedor del PO vinculado (fix 09).
        # El 10 (PO vs PR) cruza pr_lines con po_lines por PR DocEntry/Line; el 11 (PO sin PR) sale de po_lines.
        "imports_extra": [
            {"name": "pr_lines", "sql": q("PO", "pr_lines"), "param_names": [], "defaults": {}},
        ],
        "tests": {
            "PO_ANALYTIC_01": {"param_names": ["split_window"], "defaults": {}},  # Split POs (PARAM1=dias, vacio=7)
            "PO_ANALYTIC_02": {"param_names": [], "defaults": {}},                # Duplicate POs (sin ventana, todo el periodo)
            "PO_ANALYTIC_03": {"param_names": [], "defaults": {}},                # Creo y aprobo el mismo usuario
            "PO_ANALYTIC_04": {"param_names": [], "defaults": {}},                # GR antes de la fecha de PO (excluye GR sin fecha)
            "PO_ANALYTIC_05": {"param_names": ["gr_days"], "defaults": {}},       # GR > N dias de la aprobacion (PARAM1=dias, vacio=30)
            "PO_ANALYTIC_06": {"param_names": ["price_thr"], "defaults": {}},     # Dif. de precio mismo vendor (PARAM1=umbral, vacio=cualquier dif)
            "PO_ANALYTIC_07": {"param_names": [], "defaults": {}},                # Mismo user creo PO y recepciono GR
            "PO_ANALYTIC_08": {"param_names": [], "defaults": {}},                # POs por item por mes (resumen)
            "PO_ANALYTIC_09": {"param_names": ["split_window"], "defaults": {}},  # Split PRs (PARAM1=dias, vacio=7) -- usa pr_lines
            "PO_ANALYTIC_10": {"param_names": [], "defaults": {}},                # PO vs PR (dif. cantidad/item) -- usa pr_lines
            "PO_ANALYTIC_11": {"param_names": [], "defaults": {}},                # POs sin PR -- usa pr_lines
        },
    },
    "GL": {
        # PRIMARIO: base de asientos (linea) -> 01,02,03,04,07,09..16
        "import": {
            "name": "gl_journal_lines",
            "sql": q("GL", "gl_journal_lines"),
            "param_names": [],         # usa FROM/TO (filtra JDT1.RefDate)
            "defaults": {},
        },
        # ADICIONALES:
        "imports_extra": [
            {"name": "gl_accounts",         "sql": q("GL", "gl_accounts"),         "param_names": [], "defaults": {}},  # 05, 08(a) (maestro completo)
            {"name": "gl_account_activity", "sql": q("GL", "gl_account_activity"), "param_names": [], "defaults": {}},  # 08(b) (ultimo mov < FROM)
            {"name": "gl_reversals",        "sql": q("GL", "gl_reversals"),        "param_names": [], "defaults": {}},  # 06 (reversas <= DATE_TO)
        ],
        "tests": {
            "GL_ANALYTIC_01": {"param_names": [], "defaults": {}},               # fin de semana
            "GL_ANALYTIC_02": {"param_names": ["basis"], "defaults": {}},        # monto redondo de mil (PARAM1 opc: LOCAL|USD)
            "GL_ANALYTIC_03": {"param_names": ["words"], "defaults": {}},        # palabras sospechosas (PARAM1 = lista, coma)
            "GL_ANALYTIC_04": {"param_names": [], "defaults": {}},               # creo = aprobo
            "GL_ANALYTIC_05": {"param_names": [], "defaults": {}},               # cuentas creadas en el periodo (FROM..TO)
            "GL_ANALYTIC_06": {"param_names": ["min_count"], "defaults": {}},    # reversados > N veces (PARAM1, vacio = 2)
            "GL_ANALYTIC_07": {"param_names": ["fy"], "defaults": {}},           # posteado a periodo anterior/cerrado (PARAM1 = FY)
            "GL_ANALYTIC_08": {"param_names": ["months"], "defaults": {}},       # cuentas inactivas (PARAM1 = meses, vacio = 6)
            "GL_ANALYTIC_09": {"param_names": [], "defaults": {}},               # dup. misma cuenta + monto
            "GL_ANALYTIC_10": {"param_names": [], "defaults": {}},               # dup. misma cuenta + descripcion
            "GL_ANALYTIC_11": {"param_names": [], "defaults": {}},               # dup. mismo monto + descripcion
            "GL_ANALYTIC_12": {"param_names": [], "defaults": {}},               # conteo de asientos por cuenta
            "GL_ANALYTIC_13": {"param_names": [], "defaults": {}},               # resumen por cuenta/mes/compania
            "GL_ANALYTIC_14": {"param_names": [], "defaults": {}},               # cuentas por tipo de asiento
            "GL_ANALYTIC_15": {"param_names": [], "defaults": {}},               # cuentas por usuario que postea
            "GL_ANALYTIC_16": {"param_names": ["account"], "defaults": {}},      # asientos de una cuenta (PARAM1 = cuenta)
        },
    },
}


def module_import(module: str) -> dict | None:
    return REGISTRY.get(module, {}).get("import")


def module_imports_extra(module: str) -> list:
    """Imports adicionales del modulo (vacio si no tiene). Para el mini-loop de run_import."""
    return REGISTRY.get(module, {}).get("imports_extra", [])


def module_test(module: str, test: str) -> dict | None:
    return REGISTRY.get(module, {}).get("tests", {}).get(test)


def resolve_params(defn: dict, params_positional: list) -> dict:
    """Combina params posicionales (del Excel) con los nombres del registry + defaults."""
    names = (defn or {}).get("param_names", [])
    out = dict((defn or {}).get("defaults", {}))
    for i, name in enumerate(names):
        if i < len(params_positional):
            out[name] = params_positional[i]
    return out
