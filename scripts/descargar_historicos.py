r"""
descargar_historicos.py
-----------------------
Descarga masiva de los archivos HISTÓRICOS de Consulta Amigable (MEF) que
alimentan los Bloques 2 (ranking GOREs) y 6 (Rubro Sede Central):

    xls/historico/        TRIMESTRES/1T_2025.xls  MESES/julio_2025.xls  ANUALES/anual_2025_gores.xls
    xls/historico_rubro/  TRIMESTRE/1T_RUBRO_2025.xls  MES/JULIO_RUBRO_2025.xls  ANUAL/ANUAL_RUBRO_2025.xls

Antes había que exportarlos a mano uno por uno (176 archivos para 11 años).
Este script reutiliza el drill-down ya probado de descargar_xls_mef.py y solo
agrega: el año (parámetro y= de la URL) y el filtro de período.

USO
    python scripts\descargar_historicos.py --plan                  # qué hay y qué falta (no abre navegador)
    python scripts\descargar_historicos.py                         # descarga solo lo que FALTA (o es inválido)
    python scripts\descargar_historicos.py --años 2025 --hasta 11  # acotar
    python scripts\descargar_historicos.py --tipo rubro            # solo Bloque 6 (o: gores)
    python scripts\descargar_historicos.py --forzar                # re-descarga aunque exista

SEGURIDAD
  * Cada archivo se baja a xls/_tmp_historicos/ y se VALIDA con los mismos
    parsers de actualizar_hist_gores.py / actualizar_rb_hist_sc.py (año, mes,
    Gasto vs Ingresos, Sólo Proyectos, UE 001-855, 26 GOREs...). Solo si pasa
    se mueve a su carpeta definitiva. Un archivo inválido NUNCA reemplaza a uno bueno.
  * Si el archivo ya existe, se conserva una copia .bak al reemplazarlo (--forzar).
  * Tras 3 fallos seguidos se detiene (algo cambió en el sitio) en vez de insistir.
  * No toca Git ni los JSON: después corre `python scripts\cierre_mensual.py`.

CALIBRACIÓN: el filtro de período es un paso más de la cadena de clics
   (botón "Trimestre" -> fila "1er Trimestre"), grabado con playwright codegen.
   Lo que sigue sin verificar en vivo está marcado en la sección FILTRO DE PERÍODO.
"""

import argparse
import os
import shutil
import sys
import time
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))

import actualizar_hist_gores as g    # noqa: E402
import actualizar_rb_hist_sc as rb   # noqa: E402

CARPETA_TMP = RAIZ / "xls" / "_tmp_historicos"
AÑOS_DEFECTO = list(range(2015, 2026))     # agregar el año nuevo cada enero
MAX_FALLOS_SEGUIDOS = 3

# Períodos en orden. (etiqueta_gores, clave_rubro, nombre_mes_en_nombre_archivo, tipo_periodo, valor)
#   tipo_periodo: "trimestre" -> valor 1/2 ; "mes" -> valor 1..12 ; "anual" -> sin filtro
PERIODOS = [
    ("T1", "dev_t1", "1T", "trimestre", 1),
    ("T2", "dev_t2", "2T", "trimestre", 2),
    ("JULIO", "dev_sem", "julio", "mes", 7),
    ("AGOSTO", "dev_ago", "agosto", "mes", 8),
    ("SETIEMBRE", "dev_set", "setiembre", "mes", 9),
    ("OCTUBRE", "dev_oct", "octubre", "mes", 10),
    ("NOVIEMBRE", "dev_nov", "noviembre", "mes", 11),
    ("DICIEMBRE", "dev_dic", "diciembre", "mes", 12),
    ("ANUAL", "ANUAL", "anual", "anual", None),
]
MES_DE = {"T1": 3, "T2": 6, "JULIO": 7, "AGOSTO": 8, "SETIEMBRE": 9, "OCTUBRE": 10,
          "NOVIEMBRE": 11, "DICIEMBRE": 12, "ANUAL": 13}


# ---------------------------------------------------------------------------
# NOMBRES Y RUTAS (idénticos a los que esperan los scripts históricos)
# ---------------------------------------------------------------------------
def nombre_archivo(tipo, año, p):
    et_g, clave_rb, etiqueta, _, _ = p
    if tipo == "gores":
        if et_g == "ANUAL":
            return g.ARCHIVO_ANUAL.format(año=año)
        return dict((x[0], x[1]) for x in g.PERIODOS_MAESTRO)[et_g].format(año=año)
    if et_g == "ANUAL":
        return rb.ARCHIVO_ANUAL.format(año=año)
    return dict((x[0], x[1]) for x in rb.PERIODOS_MAESTRO)[clave_rb].format(año=año)


def ruta_final(tipo, año, p):
    et_g, clave_rb, _, _, _ = p
    nombre = nombre_archivo(tipo, año, p)
    if tipo == "gores":
        carpeta = RAIZ / g.CARPETA_XLS / g.SUBCARPETA.get(et_g, "MESES")
        existente = g.ubicar_archivo(nombre, et_g)
    else:
        carpeta = RAIZ / rb.CARPETA_XLS / rb.SUBCARPETA.get(clave_rb, "MES")
        existente = rb.ubicar_archivo(nombre, clave_rb)
    # ubicar_archivo es relativo al cwd; si el archivo ya existe en otra
    # variante (mayúsculas, "septiembre") se respeta esa ruta.
    existente = Path(existente)
    if not existente.is_absolute():
        existente = RAIZ / existente
    return existente if existente.exists() else carpeta / nombre


def validar(tipo, ruta, año, p):
    """True si el archivo pasa los mismos controles que usan los scripts históricos."""
    et_g, clave_rb = p[0], p[1]
    if tipo == "gores":
        return g.parsear_gores(str(ruta), año=año, periodo=None if et_g == "ANUAL" else et_g) is not None
    return rb.parsear_rubros(str(ruta), año, None if et_g == "ANUAL" else clave_rb) is not None


# ---------------------------------------------------------------------------
# PLAN: qué hay y qué falta
# ---------------------------------------------------------------------------
def construir_tareas(tipos, años, hasta, forzar, con_anual=True, anual_rubro=False):
    tareas, ok, malos = [], 0, []
    for año in años:
        for p in PERIODOS:
            if p[3] == "anual" and not con_anual:
                continue
            if p[3] != "anual" and MES_DE[p[0]] > hasta:
                continue
            for tipo in tipos:
                # El anual de Rubro solo se usa para años que aún no están en el JSON
                # (hoy 2015-2021 ya están); por defecto no se pide.
                if tipo == "rubro" and p[3] == "anual" and not anual_rubro:
                    continue
                ruta = ruta_final(tipo, año, p)
                if ruta.exists() and not forzar:
                    # Se valida en silencio: un archivo viejo equivocado también se re-descarga.
                    import contextlib, io
                    with contextlib.redirect_stdout(io.StringIO()):
                        bueno = validar(tipo, ruta, año, p)
                    if bueno:
                        ok += 1
                        continue
                    malos.append(ruta.name)
                tareas.append({"tipo": tipo, "año": año, "p": p, "ruta": ruta,
                               "existe": ruta.exists()})
    return tareas, ok, malos


def imprimir_plan(tareas, ok, malos):
    print(f"\nArchivos válidos ya presentes: {ok}")
    print(f"Por descargar: {len(tareas)}  (de ellos {len(malos)} existen pero son inválidos)")
    por_año = {}
    for t in tareas:
        por_año.setdefault(t["año"], []).append(f"{t['tipo']}:{t['p'][2]}")
    for año in sorted(por_año):
        print(f"  {año}: {len(por_año[año])} -> {', '.join(por_año[año])}")
    if tareas:
        print(f"Tiempo estimado: ~{len(tareas) * 45 // 60} min "
              f"(≈45 s por archivo; cada enero solo hay ~{9*2} del año nuevo).")


# ---------------------------------------------------------------------------
# FILTRO DE PERÍODO (calibrado con playwright codegen real, 02/11/2026)
# ---------------------------------------------------------------------------
# En Consulta Amigable el período es un paso MÁS de la cadena de clics, igual
# que Nivel de Gobierno / Sector: botón "Trimestre" -> fila "1er Trimestre".
# Se inserta justo después de "Sector 99", tal como lo grabó el codegen:
#   Nivel de Gobierno -> Sector -> Trimestre -> (Pliego ...)
# Verificado: trimestres 1 y 2 (descarga real, datos idénticos a los manuales) y
# la cadena del mes (codegen): Sector -> botón "Mes" -> fila ": 'Julio" -> Pliego.
# Verificado en vivo: julio, agosto, octubre y noviembre; setiembre se corrigió tras el fallo.
# Sin verificar: diciembre (mismo formato). La validación de cada archivo atrapa un período equivocado.
BOTON_PERIODO = {"trimestre": "Trimestre", "mes": "Mes"}
FILA_TRIMESTRE = {1: "1er Trimestre", 2: "2do Trimestre"}
# Las filas de mes llevan un apóstrofo en el HTML del MEF (": 'Julio"), como mostró
# el codegen del 02/11/2026. Se busca por ese texto parcial (NO exacto).
# Setiembre: la fila dice "9: 'Setiembre" (verificado en vivo 02/11/2026). Buscar solo
# "tiembre" fallaba: también coincidía con la celda del encabezado.
FILA_MES = {7: ": 'Julio", 8: ": 'Agosto", 9: ": 'Setiembre", 10: ": 'Octubre",
            11: ": 'Noviembre", 12: ": 'Diciembre"}


def paso_periodo(p):
    """Devuelve el paso (botón, fila) del período, o None si es anual."""
    tipo_periodo, valor = p[3], p[4]
    if tipo_periodo == "anual":
        return None
    fila = FILA_TRIMESTRE[valor] if tipo_periodo == "trimestre" else FILA_MES[valor]
    return (BOTON_PERIODO[tipo_periodo], fila)


# ---------------------------------------------------------------------------
# DESCARGA
# ---------------------------------------------------------------------------
def config_para(tipo, año, p, destino):
    import descargar_xls_mef as d
    base_nombre = "nacional_gores.xls" if tipo == "gores" else "rubro_sede_central.xls"
    base = next(a for a in d.ARCHIVOS if a["nombre"] == base_nombre)
    cfg = dict(base)
    cfg["url"] = f"https://apps5.mineco.gob.pe/transparencia/Navegador/default.aspx?y={año}&ap=Proyecto"
    cfg["destino"] = str(destino)
    pasos = list(base["pasos"])
    periodo = paso_periodo(p)
    if periodo:
        # El checkpoint de breadcrumb sigue apuntando al último nivel de la
        # jerarquía (no al período, cuyo texto en el breadcrumb no está verificado).
        cfg["texto_breadcrumb"] = pasos[-1][1]
        pasos.insert(2, periodo)          # después de Nivel de Gobierno y Sector
        if p[3] == "trimestre":
            cfg["filas_exactas"] = (periodo[1],)   # codegen usó exact=True para "1er Trimestre"
    cfg["pasos"] = pasos
    return d, cfg


def descargar(tareas):
    from playwright.sync_api import sync_playwright
    CARPETA_TMP.mkdir(parents=True, exist_ok=True)
    fallos_seguidos, hechos, fallidos = 0, 0, []
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False, slow_mo=200)
        page = browser.new_page()
        for i, t in enumerate(tareas, 1):
            tmp = CARPETA_TMP / t["ruta"].name
            print(f"\n[{i}/{len(tareas)}] {t['tipo']} {t['año']} {t['p'][2]} -> {t['ruta'].name}")
            try:
                d, cfg = config_para(t["tipo"], t["año"], t["p"], tmp)
                d.procesar_archivo(page, cfg)
                if not validar(t["tipo"], tmp, t["año"], t["p"]):
                    raise RuntimeError("el archivo descargado NO pasó la validación (ver ❌ arriba)")
                t["ruta"].parent.mkdir(parents=True, exist_ok=True)
                if t["ruta"].exists():
                    shutil.copy(t["ruta"], str(t["ruta"]) + ".bak")
                shutil.move(str(tmp), str(t["ruta"]))
                print(f"  ✅ guardado en {t['ruta']}")
                hechos += 1
                fallos_seguidos = 0
            except RuntimeError as e:
                if "calibrar" in str(e):
                    print(f"  ⛔ {e}")
                    fallidos.append(t["ruta"].name)
                    break
                print(f"  ❌ {e}")
                fallidos.append(t["ruta"].name)
                fallos_seguidos += 1
            except Exception as e:
                print(f"  ❌ {e.__class__.__name__}: {e}")
                fallidos.append(t["ruta"].name)
                fallos_seguidos += 1
            if tmp.exists():
                tmp.unlink()
            if fallos_seguidos >= MAX_FALLOS_SEGUIDOS:
                print(f"\n⛔ {MAX_FALLOS_SEGUIDOS} fallos seguidos: me detengo (¿cambió el sitio?).")
                break
            time.sleep(1)
        input("\nTerminé. ENTER para cerrar el navegador...")
        browser.close()
    return hechos, fallidos


def main():
    ap = argparse.ArgumentParser(description="Descarga masiva de históricos (Bloques 2 y 6).")
    ap.add_argument("--plan", action="store_true", help="solo mostrar qué falta")
    ap.add_argument("--años", type=int, nargs="+", default=AÑOS_DEFECTO)
    ap.add_argument("--hasta", type=int, default=11, choices=range(6, 13), metavar="MES",
                    help="último mes a considerar (defecto 11)")
    ap.add_argument("--tipo", choices=["gores", "rubro", "ambos"], default="ambos")
    ap.add_argument("--forzar", action="store_true", help="re-descargar aunque exista y sea válido")
    ap.add_argument("--sin-anual", action="store_true", help="omitir los archivos anuales (PIM)")
    ap.add_argument("--anual-rubro", action="store_true",
                    help="incluir ANUAL_RUBRO_{año} (solo hace falta para un año nuevo)")
    args = ap.parse_args()

    os.chdir(RAIZ)   # los scripts históricos usan rutas relativas
    tipos = ["gores", "rubro"] if args.tipo == "ambos" else [args.tipo]
    tareas, ok, malos = construir_tareas(tipos, args.años, args.hasta, args.forzar,
                                         con_anual=not args.sin_anual, anual_rubro=args.anual_rubro)
    imprimir_plan(tareas, ok, malos)
    if args.plan or not tareas:
        if not tareas:
            print("\n✅ No falta nada.")
        return 0

    hechos, fallidos = descargar(tareas)
    print(f"\nDescargados y validados: {hechos}/{len(tareas)}")
    if fallidos:
        print("Pendientes/fallidos:", ", ".join(fallidos))
        return 1
    print("Siguiente paso: python scripts\\cierre_mensual.py --dry-run")
    return 0


if __name__ == "__main__":
    sys.exit(main())
