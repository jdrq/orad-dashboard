#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# =============================================================================
# cierre_mensual.py  —  UN SOLO COMANDO para el cierre de mes del dashboard ORAD
# ORAD - Gobierno Regional de Lambayeque
#
# Reemplaza el procedimiento manual de ~10 pasos (editar arrays en index.html,
# descomentar PERIODOS en dos scripts, correr cada uno, comparar a ojo...).
#
# QUÉ HACE, EN ORDEN (si un paso crítico falla, se detiene y lo dice):
#   0. (opcional, --descargar) Ejecuta scripts/descargar_xls_mef.py.
#   1. MENSUAL  Valida xls/gore_devengado_mes.xls y xls/sede_devengado_mes.xls,
#               cuadra contra el totalizador del MEF y escribe
#               data/mensual-actual.js (Bloques 2D, 6D y 9). Detecta y registra
#               las reclasificaciones del MEF sobre meses ya cerrados y deja el
#               mes como PROVISIONAL (o lo confirma si ya pasó la fecha de revisión).
#   2. HISTÓRICOS  Comprueba que estén los archivos para todos los años y corre
#               actualizar_hist_gores.py (Bloque 2) y actualizar_rb_hist_sc.py
#               (Bloque 6) con --hasta N: no hay que descomentar nada. Verifica
#               que cada año quedó con el corte nuevo.
#               OJO: los históricos van UN MES ADELANTE del cierre mensual, porque
#               se comparan contra el año en curso (mes en curso, parcial, de 2026
#               vs. ese mismo mes ya completo de años anteriores). Ej.: al cerrar
#               OCTUBRE en 2D/6D/9, los históricos se activan hasta NOVIEMBRE.
#               Es el mismo criterio que se vino usando (ver resumen de octubre).
#   3. RESUMEN  Cifras clave, pendientes y los comandos git sugeridos.
#
# USO (SIEMPRE desde la raíz del repo, donde está index.html):
#   python scripts\cierre_mensual.py                 cierra el mes anterior al de hoy
#   python scripts\cierre_mensual.py --mes 10        cierra un mes concreto
#   python scripts\cierre_mensual.py --sin-historicos   solo la parte mensual
#                                                    (úsalo para REVISAR un mes provisional)
#   python scripts\cierre_mensual.py --hist-hasta 11   fuerza el corte de los históricos
#   python scripts\cierre_mensual.py --dry-run       valida todo sin escribir nada
#   python scripts\cierre_mensual.py --descargar     baja primero los XLS del MEF
#
# CÓDIGOS DE SALIDA: 0 = todo bien · 1 = falló la validación mensual (no se
# escribió nada) · 2 = lo mensual quedó bien pero faltan/fallaron históricos.
# =============================================================================

import argparse
import json
import os
import subprocess
import sys
from datetime import date

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)

import generar_mensual as gm                  # noqa: E402
import actualizar_hist_gores as g             # noqa: E402
import actualizar_rb_hist_sc as rb            # noqa: E402

MES_CORTO = {6: "Jun", 7: "Jul", 8: "Ago", 9: "Set", 10: "Oct", 11: "Nov", 12: "Dic"}
CLAVE_RUBRO = {6: "dev_t2", 7: "dev_sem", 8: "dev_ago", 9: "dev_set",
               10: "dev_oct", 11: "dev_nov", 12: "dev_dic"}
PERIODO_GORES = {6: "T2", 7: "JULIO", 8: "AGOSTO", 9: "SETIEMBRE",
                 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"}


def titulo(txt):
    print("\n" + "=" * 72 + f"\n  {txt}\n" + "=" * 72)


def archivos_faltantes(mes):
    """Archivos históricos que faltan para cerrar `mes` (todos los años de cada script)."""
    faltan = []
    for etiqueta, patron, _ in g.PERIODOS_MAESTRO:
        if g.MES_DEL_PERIODO[etiqueta] > mes:
            continue
        for anio in g.AÑOS:
            ruta = g.ubicar_archivo(patron.format(año=anio), etiqueta)
            if not os.path.isfile(ruta):
                faltan.append(ruta.replace("\\", "/"))
    for anio in g.AÑOS:                           # PIM anual del ranking (Bloque 2)
        ruta = g.ubicar_archivo(g.ARCHIVO_ANUAL.format(año=anio), "ANUAL")
        if not os.path.isfile(ruta):
            faltan.append(ruta.replace("\\", "/"))
    for clave, patron, _ in rb.PERIODOS_MAESTRO:
        if rb.MES_DEL_PERIODO[clave] > mes:
            continue
        for anio in rb.AÑOS:
            ruta = rb.ubicar_archivo(patron.format(año=anio), clave)
            if not os.path.isfile(ruta):
                faltan.append(ruta.replace("\\", "/"))
    return faltan


def correr_script(nombre, mes):
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    p = subprocess.run([sys.executable, os.path.join("scripts", nombre), "--hasta", str(mes)],
                       capture_output=True, text=True, encoding="utf-8", errors="replace", env=env)
    return p.returncode, (p.stdout or "") + (p.stderr or "")


def lineas_problema(salida):
    return [l.strip() for l in salida.splitlines()
            if any(x in l for x in ("⚠️", "❌", "🚫", "BLOQUEADO", "Traceback", "Error"))]


def anios_procesados_ok(salida):
    """
    Años que el propio script dio por buenos EN ESTA CORRIDA (línea '... Año N completo').
    Hace falta además del JSON: un año bloqueado conserva sus datos anteriores, que en una
    segunda corrida ya traen el mes y parecerían correctos aunque el archivo nuevo sea malo.
    """
    ok, actual = set(), None
    for l in salida.splitlines():
        if l.startswith("--- Año "):
            actual = int(l.split()[2])
        elif actual is not None and "completo" in l and "Año" in l:
            ok.add(actual)
    return ok


def verificar_gores(mes, salida):
    """Años cuyo JSON quedó realmente con el corte pedido."""
    with open(g.ARCHIVO_JSON, encoding="utf-8") as f:
        data = json.load(f)["semestres"]
    ok, mal = [], []
    buenos = anios_procesados_ok(salida)
    for anio in g.AÑOS:
        e = data.get(str(anio), {})
        (ok if (e.get("label") == f"Ene–{MES_CORTO[mes]} {anio}" and anio in buenos) else mal).append(anio)
    lam = data.get(str(g.AÑOS[-1]), {}).get("lambayeque", {})
    return ok, mal, lam


def verificar_rubro(mes, salida):
    with open(rb.ARCHIVO_JSON, encoding="utf-8") as f:
        data = json.load(f)
    ok, mal = [], []
    buenos = anios_procesados_ok(salida)
    for anio in rb.AÑOS:
        e = data.get(str(anio), {})
        bueno = (e.get(CLAVE_RUBRO[mes]) is not None and e.get("label") == f"Ene–{MES_CORTO[mes]} {anio}"
                 and anio in buenos)
        (ok if bueno else mal).append(anio)
    return ok, mal, data.get(str(rb.AÑOS[-1]), {})


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="Cierre mensual del dashboard ORAD en un solo comando.")
    ap.add_argument("--mes", type=int, help="mes que se cierra (1-12). Por defecto: el anterior al de la consulta.")
    ap.add_argument("--sin-historicos", action="store_true", help="solo la parte mensual (2D/6D/9).")
    ap.add_argument("--hist-hasta", type=int, choices=range(6, 13), metavar="MES",
                    help="corte de los históricos (por defecto: mes del cierre + 1, máx. diciembre).")
    ap.add_argument("--dry-run", action="store_true", help="valida todo sin escribir nada.")
    ap.add_argument("--descargar", action="store_true", help="ejecuta primero scripts/descargar_xls_mef.py.")
    ap.add_argument("--forzar", action="store_true", help="permite retroceder a un mes anterior al ya guardado.")
    args = ap.parse_args()

    if not (os.path.isfile("index.html") and os.path.isdir("xls") and os.path.isdir("scripts")):
        print("[ERROR] Corre este comando desde la RAIZ del repo (donde está index.html):\n"
              "        cd C:\\Users\\juan-\\Desktop\\orad-dashboard\n"
              "        python scripts\\cierre_mensual.py")
        sys.exit(1)

    titulo("CIERRE MENSUAL — Dashboard ORAD" + ("   [DRY-RUN: no se escribe nada]" if args.dry_run else ""))

    # --- estado previo ---
    try:
        previo = gm.cargar_previo()
    except Exception as e:                        # archivo ilegible: lo avisa pero sigue
        previo = None
        print(f"  [aviso] No se pudo leer el cierre anterior: {e}")
    if previo:
        est = previo.get("estado", {})
        txt = "provisional" if est.get("provisional") else "confirmado"
        extra = f", revisar desde {date.fromisoformat(est['revisar_desde']):%d/%m/%Y}" if est.get("revisar_desde") else ""
        print(f"  Último cierre guardado: {gm.MESES_LARGO[previo['ultimo_mes_cerrado']-1]} {previo['anio']} ({txt}{extra})")

    # --- 0) descarga opcional ---
    if args.descargar:
        titulo("0) DESCARGA de los XLS del MEF (scripts/descargar_xls_mef.py)")
        print("  Se abrirá el navegador. Al terminar debe decir '17/17 archivos OK'. Si no, DETENTE y revisa.")
        rc = subprocess.run([sys.executable, os.path.join("scripts", "descargar_xls_mef.py")]).returncode
        if rc != 0:
            print(f"  [ERROR] La descarga terminó con código {rc}. No se continúa.")
            sys.exit(1)

    # --- 1) mensual ---
    titulo("1) VALORES MENSUALES  (Bloques 2D, 6D y 9)")
    try:
        r = gm.generar(mes=args.mes, dry_run=args.dry_run, forzar=args.forzar)
    except gm.ErrorValidacion as e:
        print(f"\n  [ERROR] {e}\n  No se escribió nada. Corrige el archivo y vuelve a correr el comando.")
        sys.exit(1)
    gm.imprimir_resumen(r)
    mes = r["cierre"]
    mes_nombre = gm.MESES_LARGO[mes - 1]
    mes_hist = args.hist_hasta or min(mes + 1, 12)       # históricos: mes en curso (cierre + 1)
    hist_nombre = gm.MESES_LARGO[mes_hist - 1]

    # --- 2) históricos ---
    resultado_hist = {"estado": "omitido", "detalle": []}
    if args.sin_historicos:
        print("\n  (históricos omitidos por --sin-historicos)")
    elif mes_hist < 6:
        print(f"\n  Los históricos (Bloques 2 y 6) solo se actualizan de junio en adelante; "
              f"no aplican para {mes_nombre}.")
    else:
        titulo(f"2) HISTÓRICOS  (Bloque 2 ranking GOREs · Bloque 6 rubro Sede Central) — "
               f"corte Ene–{MES_CORTO[mes_hist]} (mes en curso)")
        faltan = archivos_faltantes(mes_hist)
        if faltan:
            resultado_hist = {"estado": "faltan", "detalle": faltan}
            print(f"  [!!] Faltan {len(faltan)} archivo(s) histórico(s). NO se corrieron los scripts:")
            for f in faltan[:30]:
                print(f"       - {f}")
            if len(faltan) > 30:
                print(f"       ... y {len(faltan) - 30} más")
        elif args.dry_run:
            print("  Archivos históricos completos (dry-run: no se corrieron los scripts).")
            resultado_hist = {"estado": "dry", "detalle": []}
        else:
            problemas = []
            for etiqueta, script, verificar in (("Ranking GOREs (Bloque 2)", "actualizar_hist_gores.py", verificar_gores),
                                               ("Rubro Sede Central (Bloque 6)", "actualizar_rb_hist_sc.py", verificar_rubro)):
                rc, salida = correr_script(script, mes_hist)
                ok, mal, _ = verificar(mes_hist, salida)
                marca = "[OK]" if (rc == 0 and not mal) else "[!!]"
                print(f"  {marca} {etiqueta}: {len(ok)}/{len(ok) + len(mal)} años con corte Ene–{MES_CORTO[mes_hist]}"
                      + (f"  ·  con problema: {', '.join(map(str, mal))}" if mal else ""))
                if rc != 0 or mal:
                    problemas.append(etiqueta)
                    for l in lineas_problema(salida)[:12]:
                        print(f"       {l}")
            resultado_hist = {"estado": "problemas" if problemas else "ok", "detalle": problemas}

    # --- 3) resumen ---
    titulo("RESUMEN")
    est = r["estado"]
    print(f"  Mensual {mes_nombre}: {'PROVISIONAL' if est['provisional'] else 'CONFIRMADO'}"
          + (f" — volver a revisar desde el {date.fromisoformat(est['revisar_desde']):%d/%m/%Y}: "
             f"baja gore_devengado_mes.xls y sede_devengado_mes.xls y corre\n"
             f"       python scripts\\cierre_mensual.py --sin-historicos" if est["provisional"] else ""))
    h = resultado_hist["estado"]
    print({"ok": "  Históricos: OK",
           "omitido": "  Históricos: omitidos",
           "dry": "  Históricos: archivos completos (no ejecutados)",
           "faltan": "  Históricos: PENDIENTES (faltan archivos, ver arriba)",
           "problemas": "  Históricos: CON PROBLEMAS (ver arriba)"}[h])

    if not args.dry_run:
        print("\n  Para publicar:\n"
              "     git status\n"
              "     git add data/mensual-actual.js data/historico_progresivo.json data/rb_hist_sc_progresivo.json xls/\n"
              f"     git commit -m \"data: cierre {mes_nombre.lower()} {r['anio']}\"\n"
              "     git push origin main\n"
              "  Luego recarga la página con Ctrl + F5.")
    sys.exit(0 if h in ("ok", "omitido", "dry") else 2)


if __name__ == "__main__":
    main()
