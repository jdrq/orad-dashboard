#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# =============================================================================
# generar_cierres.py  —  data/cierres.json: cómo cerró cada mes (Pliego y Sede Central)
# ORAD - Gobierno Regional de Lambayeque
#
# FUENTE: xls/gore_devengado_mes.xls y xls/sede_devengado_mes.xls ("por Mes"). Traen, por
# cada mes, el movimiento NO acumulado de certificación, compromiso, atención, devengado y
# girado (puede ser negativo: rebajas o anulaciones). Se publican solo los meses CERRADOS
# (data/mensual-actual.js), según la consulta más reciente (incluye reclasificaciones del MEF).
#
# PIM: esos archivos NO traen el PIM por mes. Se completa así, y se indica el origen:
#   1) data/pim_cierre.json (opcional, manual): {"pliego": {"1": 800000000}, "sede_central": {}}
#   2) la última consulta del mes en data/serie (referencial; existe desde junio 2026)
#   3) si no hay ninguno, queda vacío (no se inventa).
#
# USO:  python scripts/generar_cierres.py
# =============================================================================

import csv
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import controles_calidad as cc            # noqa: E402  (reutiliza la lectura de XLS)
import generar_mensual as gm              # noqa: E402

from bs4 import BeautifulSoup             # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent
RUTA_SALIDA = RAIZ / "data" / "cierres.json"
CAMPOS = ["certificacion", "compromiso_anual", "atencion_mensual", "devengado", "girado"]
TOLERANCIA = 10      # soles: redondeos al sumar 12 meses
FUENTES = {
    "pliego": ("gore_devengado_mes.xls", "Pliego 452", "fact_pliego_diario", None),
    "sede_central": ("sede_devengado_mes.xls", "Unidad Ejecutora 001-855", "fact_ue_diario", "001-855"),
}


def meses_del_archivo(ruta, prefijo_total):
    """({mes: {campo: monto}}, fecha_consulta). Valida que los meses sumen el total del MEF."""
    sopa = BeautifulSoup(Path(ruta).read_bytes().decode("cp1252", errors="replace"), "html.parser")
    meses = {}
    for tr in sopa.find_all("tr"):
        c = [" ".join(td.get_text(" ", strip=True).split()) for td in tr.find_all(["td", "th"])]
        m = gm._mes_de_texto(c[0]) if len(c) >= 8 else None
        if m:
            meses[m] = {k: gm._num(c[3 + i]) or 0 for i, k in enumerate(CAMPOS)}
    info = cc.leer(ruta)
    total = next(f for f in info["filas"] if f["etiqueta"].startswith(prefijo_total))
    for i, campo in enumerate(CAMPOS):
        suma, esperado = sum(v[campo] for v in meses.values()), total["n"][2 + i]
        if abs(suma - esperado) > TOLERANCIA:
            raise ValueError(f"{Path(ruta).name}: la suma mensual de {campo} ({suma:,}) no cuadra con {esperado:,.0f}")
    return meses, info["fecha"]


def pim_referencial(anio, tabla, ue):
    """{mes: (pim, fecha)} con la última consulta de cada mes en la serie diaria."""
    ruta = RAIZ / "data" / "serie" / f"{tabla}.csv"
    salida = {}
    if ruta.exists():
        with open(ruta, encoding="utf-8-sig", newline="") as f:
            for fila in sorted(csv.DictReader(f), key=lambda x: x["fecha"]):
                if fila["fecha"].startswith(str(anio)) and (ue is None or fila["ue_codigo"] == ue):
                    salida[int(fila["fecha"][5:7])] = (float(fila["pim"]), fila["fecha"])
    return salida


def calcular():
    previo = gm.cargar_previo(str(RAIZ / "data" / "mensual-actual.js")) or {}
    anio, ultimo = previo.get("anio"), previo.get("ultimo_mes_cerrado") or 0
    ruta_manual = RAIZ / "data" / "pim_cierre.json"
    manual = json.loads(ruta_manual.read_text(encoding="utf-8-sig")) if ruta_manual.exists() else {}
    res = {"version": 1, "anio": anio, "ultimo_mes_cerrado": ultimo}
    for clave, (archivo, prefijo, tabla, ue) in FUENTES.items():
        meses, consulta = meses_del_archivo(RAIZ / "xls" / archivo, prefijo)
        if not consulta or consulta.year != anio:
            raise ValueError(f"{archivo}: la consulta ({consulta}) no es del año {anio}")
        ref = pim_referencial(anio, tabla, ue)
        lista = []
        for m in range(1, ultimo + 1):
            if m not in meses:
                raise ValueError(f"{archivo}: falta el mes {m}")
            dato = {"mes": m, **meses[m], "pim": None, "pim_origen": None}
            if str(m) in manual.get(clave, {}):
                dato["pim"], dato["pim_origen"] = float(manual[clave][str(m)]), "registrado manualmente"
            elif m in ref:
                fecha = ref[m][1]
                # solo vale como referencia de cierre si la consulta es de fin de mes
                if int(fecha[8:10]) >= 25:
                    dato["pim"] = ref[m][0]
                    dato["pim_origen"] = f"consulta del {fecha[8:10]}/{fecha[5:7]}/{fecha[:4]} (referencial)"
            lista.append(dato)
        res[clave] = lista
        res["fecha_consulta"] = consulta.isoformat()
    return res


def acumulado(cierres, ambito, mes):
    """Totales de enero al mes indicado."""
    return {c: sum(d[c] for d in cierres[ambito] if d["mes"] <= mes) for c in CAMPOS}


def ejecutar(destino=RUTA_SALIDA):
    res = calcular()
    Path(destino).write_text(json.dumps(res, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    return res


if __name__ == "__main__":
    r = ejecutar()
    for d in r["pliego"]:
        a = acumulado(r, "pliego", d["mes"])
        print(f"mes {d['mes']:2d}  devengado {d['devengado']:>13,}  acumulado {a['devengado']:>13,}  PIM {d['pim'] or '-'}")
