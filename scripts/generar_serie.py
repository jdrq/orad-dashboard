#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# =============================================================================
# generar_serie.py  —  CAPA DE DATOS: tablas limpias y diarias en CSV (listas para Power BI)
# ORAD - Gobierno Regional de Lambayeque
#
# POR QUÉ EXISTE
#   Los XLS del MEF se SOBRESCRIBEN cada día: el de hoy reemplaza al de ayer. Aquí se
#   extraen las cifras de cada archivo y se guardan como una serie histórica ordenada
#   ("una foto por día"), con el formato que Power BI entiende sin transformaciones.
#
# CAPAS (arquitectura "medallón", simple)
#   BRONCE  xls/*.xls                    datos crudos del MEF, tal cual (se sobrescriben)
#   PLATA   data/serie/fact_*.csv        datos limpios y acumulados día a día  <- ESTE SCRIPT
#   ORO     dashboard / reportes         lo que ve tu jefe
#
# TABLAS DE HECHOS (una fila por día y por elemento). Medidas: pia, pim, certificacion,
# compromiso_anual, atencion_mensual, devengado, girado, avance_pct
#   fact_pliego_diario    día × (Pliego 452)
#   fact_ue_diario        día × unidad ejecutora            (ue_pliego.xls)
#   fact_rubro_diario     día × ámbito × rubro              (rubro_pliego.xls + rubro_<UE>.xls)
#   fact_funcion_diario   día × función                     (funciones_pliego.xls)
#   fact_gore_diario      día × gobierno regional (nacional) (nacional_gores.xls)
#   fact_proyecto_diario  día × proyecto de la Sede Central (proyectos_sede_central.xls)
# DIMENSIONES (se reconstruyen solas): dim_ue, dim_rubro, dim_funcion, dim_gore, dim_proyecto
#
# USO (desde la raíz del repo)
#   python scripts\generar_serie.py                 agrega/actualiza el día de xls/ (idempotente)
#   python scripts\generar_serie.py --desde-git     reconstruye TODO el historial leyendo Git
#   python scripts\generar_serie.py --desde-git --desde 2026-09-01
# Es idempotente: correrlo dos veces el mismo día no duplica filas (reemplaza la fecha).
# =============================================================================

import argparse
import csv
import re
import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import controles_calidad as cc            # noqa: E402  (reutiliza el lector de XLS)

RAIZ = Path(__file__).resolve().parent.parent
CARPETA_XLS = RAIZ / "xls"
CARPETA_SERIE = RAIZ / "data" / "serie"

MEDIDAS = ["pia", "pim", "certificacion", "compromiso_anual", "atencion_mensual", "devengado", "girado", "avance_pct"]

# tabla -> (columnas clave sin la fecha, columnas descriptivas que NO son clave)
TABLAS = {
    "fact_pliego_diario":   (["pliego_codigo"], ["pliego_nombre"]),
    "fact_ue_diario":       (["ue_codigo"], ["ue_nombre"]),
    "fact_rubro_diario":    (["ambito", "rubro_codigo"], ["rubro_nombre"]),
    "fact_funcion_diario":  (["funcion_codigo"], ["funcion_nombre"]),
    "fact_gore_diario":     (["gore_codigo"], ["gore_nombre"]),
    "fact_proyecto_diario": (["ue_codigo", "proyecto_codigo"], ["proyecto_nombre"]),
}
# dimensión -> (tabla origen, columnas)
DIMENSIONES = {
    "dim_ue":       ("fact_ue_diario", ["ue_codigo", "ue_nombre"]),
    "dim_rubro":    ("fact_rubro_diario", ["rubro_codigo", "rubro_nombre"]),
    "dim_funcion":  ("fact_funcion_diario", ["funcion_codigo", "funcion_nombre"]),
    "dim_gore":     ("fact_gore_diario", ["gore_codigo", "gore_nombre"]),
    "dim_proyecto": ("fact_proyecto_diario", ["ue_codigo", "proyecto_codigo", "proyecto_nombre"]),
}
# Archivos "rubros de una unidad ejecutora": el ámbito sale del encabezado "Unidad Ejecutora XXX-YYY"
ARCHIVOS_RUBRO_UE = ["rubro_sede_central.xls", "rubro_peot.xls", "rubro_salud.xls", "rubro_agricultura.xls",
                     "rubro_h_belen.xls", "rubro_h_mercedes.xls", "rubro_h_regional.xls", "rubro_transportes.xls"]
RE_ETIQUETA = re.compile(r"^\s*([^:]+?)\s*:\s*(.*)$")


def _medidas(n):
    return dict(zip(MEDIDAS, n))


def _partir(etiqueta):
    m = RE_ETIQUETA.match(etiqueta)
    return (m.group(1).strip(), m.group(2).strip()) if m else ("", etiqueta.strip())


def _hijas(info, desde_encabezado=None):
    """Filas hijas (no encabezado). Si se pide, solo las posteriores al encabezado que empiece así."""
    filas, activo = [], desde_encabezado is None
    for f in info["filas"]:
        if f["encabezado"]:
            if desde_encabezado and f["etiqueta"].startswith(desde_encabezado):
                activo = True
            continue
        if activo:
            filas.append(f)
    return filas


def extraer_dia(archivos):
    """archivos: {nombre: info de controles_calidad.leer}. Devuelve (fecha, {tabla: [filas]})."""
    ue_p = archivos.get("ue_pliego.xls")
    if not ue_p or not ue_p["fecha"]:
        return None, {}
    fecha = ue_p["fecha"].isoformat()
    t = {k: [] for k in TABLAS}

    def mismo_dia(nombre):
        i = archivos.get(nombre)
        return i if i and i["fecha"] and i["fecha"].isoformat() == fecha else None

    pl = next((f for f in ue_p["filas"] if f["etiqueta"].startswith("Pliego 452")), None)
    if pl:
        cod, nom = _partir(pl["etiqueta"].replace("Pliego ", "", 1))
        t["fact_pliego_diario"].append({"fecha": fecha, "pliego_codigo": cod, "pliego_nombre": nom, **_medidas(pl["n"])})
    for f in _hijas(ue_p):
        cod, nom = _partir(f["etiqueta"])
        t["fact_ue_diario"].append({"fecha": fecha, "ue_codigo": cod, "ue_nombre": nom, **_medidas(f["n"])})

    if mismo_dia("rubro_pliego.xls"):
        for f in _hijas(archivos["rubro_pliego.xls"]):
            cod, nom = _partir(f["etiqueta"])
            t["fact_rubro_diario"].append({"fecha": fecha, "ambito": "PLIEGO", "rubro_codigo": cod, "rubro_nombre": nom, **_medidas(f["n"])})
    for nombre in ARCHIVOS_RUBRO_UE:
        info = mismo_dia(nombre)
        if not info:
            continue
        enc = next((f for f in info["filas"] if f["etiqueta"].startswith("Unidad Ejecutora")), None)
        if not enc:
            continue
        ue_cod = cc.RE_CODIGO_UE.search(enc["etiqueta"]).group(1)
        for f in _hijas(info, "Unidad Ejecutora"):
            cod, nom = _partir(f["etiqueta"])
            t["fact_rubro_diario"].append({"fecha": fecha, "ambito": ue_cod, "rubro_codigo": cod, "rubro_nombre": nom, **_medidas(f["n"])})

    if mismo_dia("funciones_pliego.xls"):
        for f in _hijas(archivos["funciones_pliego.xls"]):
            cod, nom = _partir(f["etiqueta"])
            t["fact_funcion_diario"].append({"fecha": fecha, "funcion_codigo": cod, "funcion_nombre": nom, **_medidas(f["n"])})
    if mismo_dia("nacional_gores.xls"):
        for f in _hijas(archivos["nacional_gores.xls"]):
            cod, nom = _partir(f["etiqueta"])
            t["fact_gore_diario"].append({"fecha": fecha, "gore_codigo": cod, "gore_nombre": nom, **_medidas(f["n"])})
    if mismo_dia("proyectos_sede_central.xls"):
        info = archivos["proyectos_sede_central.xls"]
        enc = next((f for f in info["filas"] if f["etiqueta"].startswith("Unidad Ejecutora")), None)
        ue_cod = cc.RE_CODIGO_UE.search(enc["etiqueta"]).group(1) if enc else ""
        for f in _hijas(info, "Unidad Ejecutora"):
            cod, nom = _partir(f["etiqueta"])
            t["fact_proyecto_diario"].append({"fecha": fecha, "ue_codigo": ue_cod, "proyecto_codigo": cod, "proyecto_nombre": nom, **_medidas(f["n"])})
    return fecha, t


# ----------------------------------------------------------------------------
# Escritura de CSV (idempotente: reemplaza las fechas que llegan)
# ----------------------------------------------------------------------------
def _fmt(v):
    if isinstance(v, float):
        return str(int(v)) if v == int(v) else f"{v:.2f}".rstrip("0").rstrip(".")
    return v


def _columnas(tabla):
    claves, desc = TABLAS[tabla]
    return ["fecha"] + claves + desc + MEDIDAS


def upsert(tabla, filas_por_fecha, carpeta=CARPETA_SERIE):
    """filas_por_fecha: {fecha: [filas]}. Reemplaza esas fechas en data/serie/<tabla>.csv."""
    cols = _columnas(tabla)
    claves = ["fecha"] + TABLAS[tabla][0]
    ruta = Path(carpeta) / f"{tabla}.csv"
    existentes = []
    if ruta.exists():
        with ruta.open(encoding="utf-8", newline="") as fh:
            existentes = [r for r in csv.DictReader(fh) if r["fecha"] not in filas_por_fecha]
    nuevas = [{c: _fmt(f.get(c, "")) for c in cols} for filas in filas_por_fecha.values() for f in filas]
    todas = existentes + nuevas
    todas.sort(key=lambda r: tuple(str(r[c]) for c in claves))
    Path(carpeta).mkdir(parents=True, exist_ok=True)
    with ruta.open("w", encoding="utf-8", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=cols, lineterminator="\n")
        w.writeheader()
        w.writerows(todas)
    return len(nuevas), len(todas)


def reconstruir_dimensiones(carpeta=CARPETA_SERIE):
    for dim, (origen, cols) in DIMENSIONES.items():
        ruta = Path(carpeta) / f"{origen}.csv"
        if not ruta.exists():
            continue
        ultimo = {}
        with ruta.open(encoding="utf-8", newline="") as fh:
            for r in csv.DictReader(fh):            # ordenado por fecha: la última descripción gana
                ultimo[tuple(r[c] for c in cols[:-1])] = r[cols[-1]]
        with (Path(carpeta) / f"{dim}.csv").open("w", encoding="utf-8", newline="") as fh:
            w = csv.writer(fh, lineterminator="\n")
            w.writerow(cols)
            for clave in sorted(ultimo):
                w.writerow(list(clave) + [ultimo[clave]])


def escribir(por_fecha, dims=True, carpeta=CARPETA_SERIE):
    """por_fecha: {fecha: {tabla: [filas]}}."""
    resumen = {}
    for tabla in TABLAS:
        nuevas, total = upsert(tabla, {f: t[tabla] for f, t in por_fecha.items() if t.get(tabla)}, carpeta)
        resumen[tabla] = {"filas_escritas": nuevas, "filas_totales": total}
    if dims:
        reconstruir_dimensiones(carpeta)
    return resumen


# ----------------------------------------------------------------------------
# Entradas
# ----------------------------------------------------------------------------
def ejecutar(carpeta_xls=CARPETA_XLS, carpeta_serie=CARPETA_SERIE):
    """Día actual (lo que hay en xls/). Lo usa generar_estado.py en el workflow diario."""
    fecha, tablas = extraer_dia(cc.leer_todos(carpeta_xls))
    if not fecha:
        raise RuntimeError("No se pudo leer ue_pliego.xls (sin fecha de consulta).")
    resumen = escribir({fecha: tablas}, carpeta=carpeta_serie)
    return {"fecha": fecha, "tablas": resumen}


def _git(*args):
    return subprocess.run(["git", *args], cwd=RAIZ, capture_output=True)


def desde_git(desde=None, carpeta_serie=CARPETA_SERIE):
    """Reconstruye la serie leyendo las versiones antiguas de los XLS en el historial de Git."""
    nombres = sorted(p.name for p in CARPETA_XLS.glob("*.xls"))
    r = _git("log", "--reverse", "--format=%H", "--", "xls/ue_pliego.xls")
    commits = r.stdout.decode().split()
    if not commits:
        raise RuntimeError("Git no devolvió historial de xls/ue_pliego.xls (¿clon superficial?). Usa 'git fetch --unshallow'.")
    por_fecha, omitidos = {}, []
    with tempfile.TemporaryDirectory() as tmp:
        for i, h in enumerate(commits, 1):
            for nombre in nombres:
                b = _git("show", f"{h}:xls/{nombre}")
                destino = Path(tmp) / nombre
                if b.returncode == 0:
                    destino.write_bytes(b.stdout)
                elif destino.exists():
                    destino.unlink()
            try:
                fecha, tablas = extraer_dia(cc.leer_todos(tmp))
            except Exception as e:
                omitidos.append((h[:7], str(e)[:80]))
                continue
            if not fecha or (desde and fecha < desde):
                continue
            por_fecha[fecha] = tablas            # un commit posterior de la misma fecha reemplaza al anterior
            if i % 20 == 0:
                print(f"  ... {i}/{len(commits)} versiones leídas")
    resumen = escribir(por_fecha, carpeta=carpeta_serie)
    return {"fechas": sorted(por_fecha), "tablas": resumen, "omitidos": omitidos, "commits_leidos": len(commits)}


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="Serie diaria en CSV (capa de datos).")
    ap.add_argument("--desde-git", action="store_true", help="reconstruye el historial leyendo Git")
    ap.add_argument("--desde", help="con --desde-git: primera fecha a incluir (AAAA-MM-DD)")
    args = ap.parse_args()
    if args.desde_git:
        res = desde_git(args.desde)
        f = res["fechas"]
        print(f"Historial reconstruido: {len(f)} días ({f[0]} a {f[-1]}) a partir de {res['commits_leidos']} versiones en Git.")
        for h, e in res["omitidos"]:
            print(f"  [omitido] {h}: {e}")
    else:
        res = ejecutar()
        print(f"Serie actualizada para la consulta del {res['fecha']}.")
    for t, r in res["tablas"].items():
        print(f"  {t:22s} +{r['filas_escritas']:6d} filas · total {r['filas_totales']:7d}")


if __name__ == "__main__":
    main()
