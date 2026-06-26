# -*- coding: utf-8 -*-
"""
debug_rp.py -- diagnostico de seleccion de Reporting Pack (FAM004).

Uso (desde la raiz del proyecto):
  python debug_rp.py                       # vuelca TODOS los archivos de input/ y resuelve las 13 para dic-2025
  python debug_rp.py LCC 2025-12-31        # detalle de por que (no) matchea una compañia/fecha
  python debug_rp.py --dir "D:\\LHA\\input" # forzar otra carpeta
"""
import os, sys, argparse
from datetime import datetime

# Encontrar la raiz del proyecto (la carpeta que contiene 'core'), suba donde suba el script.
_here = os.path.dirname(os.path.abspath(__file__))
_root = _here
for _ in range(6):
    if os.path.isdir(os.path.join(_root, "core")):
        break
    _up = os.path.dirname(_root)
    if _up == _root:
        break
    _root = _up
sys.path.insert(0, _root)

try:
    import core.rp_lookup as rl
except ModuleNotFoundError:
    sys.exit("No encuentro el paquete 'core'. Corre el script desde el proyecto "
             "(D:\\LHA) o deja core/ accesible. Raiz detectada: %s" % _root)


def get_input_dir(forced=None):
    if forced:
        return forced
    try:
        from core.analysis_base import INPUT_DIR
        return INPUT_DIR
    except Exception:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "input")


def dump_all(folder):
    print("\n== Archivos en %s ==" % folder)
    if not os.path.isdir(folder):
        print("  (la carpeta no existe)"); return
    files = sorted(os.listdir(folder))
    if not files:
        print("  (vacia)"); return
    for f in files:
        info = rl.parse_rp_filename(f)
        if info is None:
            n = rl._norm(f)
            motivo = "NO es RP: falta la frase 'reporting pack'" if 'reporting' in n \
                     else "ignorado: no contiene 'reporting pack'"
            print("  [-] %-55s %s" % (f[:55], motivo))
        else:
            print("  [RP] %-53s company=%s fy=%s mes=%s año=%s" %
                  (f[:53], info['company'], info['fy'], info['month'], info['year']))
            if info['company'] is None:
                print("        ^ compañia NO reconocida -> revisar rp_aliases (normalizado: '%s')"
                      % rl._norm(f))


def diagnose(folder, code, year, month):
    target_fy = rl.fiscal_year(year, month)
    target_pos = rl.fiscal_pos(month)
    print("\n== Resolver %s para %02d/%d  (FY esperado=%s, pos fiscal=%s) ==" %
          (code, month, year, target_fy, target_pos))
    print("   (el mes del archivo = hasta donde llega; sirve si su pos fiscal >= %s)" % target_pos)
    if not os.path.isdir(folder):
        print("  carpeta inexistente:", folder); return
    for f in sorted(os.listdir(folder)):
        info = rl.parse_rp_filename(f)
        if info is None:
            continue
        if not f.lower().endswith((".xlsx", ".xlsm")):
            verdict, why = "descarta", "extension %s (solo .xlsx/.xlsm)" % os.path.splitext(f)[1]
        elif info['company'] != code:
            verdict, why = "descarta", "company=%s (esperaba %s)" % (info['company'], code)
        elif info['fy'] is not None and info['fy'] != target_fy:
            verdict, why = "descarta", "FY=%s (esperaba %s)" % (info['fy'], target_fy)
        elif info['month'] is None:
            verdict, why = "CANDIDATO", "sin mes en nombre (full-year, prioridad baja)"
        elif rl.fiscal_pos(info['month']) < target_pos:
            verdict, why = "descarta", ("mes=%s pos=%s < %s (el reporte termina antes de tu mes)"
                                        % (info['month'], rl.fiscal_pos(info['month']), target_pos))
        else:
            verdict, why = "CANDIDATO", ("mes=%s pos=%s >= %s (cubre)"
                                         % (info['month'], rl.fiscal_pos(info['month']), target_pos))
        print("   %-9s %-52s %s" % (verdict, f[:52], why))
    chosen = rl.resolve_rp(folder, code, year, month)
    print("   --> ELEGIDO:", os.path.basename(chosen) if chosen else "NINGUNO")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("code", nargs="?", default=None)
    ap.add_argument("date_to", nargs="?", default="2025-12-31")
    ap.add_argument("--dir", default=None)
    a = ap.parse_args()
    folder = get_input_dir(a.dir)
    d = datetime.strptime(a.date_to, "%Y-%m-%d").date()

    dump_all(folder)
    if a.code:
        diagnose(folder, a.code.upper(), d.year, d.month)
    else:
        print("\n== Resolucion por compañia (fecha %s) ==" % d)
        for code in rl.COMPANIES:
            p = rl.resolve_rp(folder, code, d.year, d.month)
            print("   %-5s -> %s" % (code, os.path.basename(p) if p else "NO ENCONTRADO"))


if __name__ == "__main__":
    main()
