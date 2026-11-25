#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# =============================================================================
# alertas.py  —  alertas de EJECUCIÓN presupuestal a partir de data/serie/*.csv
# ORAD - Gobierno Regional de Lambayeque
#
# Distinto de controles_calidad.py: aquel pregunta "¿el dato está bien?";
# este pregunta "¿la ejecución va bien?". Escribe:
#   data/alertas.json   para el dashboard / reporte quincenal (Semana 5)
#   data/alertas.md     versión legible para revisar con el jefe
#
# REGLAS (umbrales PROPUESTOS: validarlos con la jefatura y editarlos aquí)
#   A1 unidad ejecutora con avance por debajo del Pliego en BRECHA_PUNTOS o más
#   A2 unidad ejecutora sin nuevo devengado en DIAS_SIN_MOVIMIENTO días
#   A3 desde MES_CONTROL_CERT, unidad ejecutora con certificación < MIN_CERT_PCT del PIM
#   A4 proyectos (Sede Central) con PIM >= PIM_PROYECTO_GRANDE y devengado 0
#   A5 proyectos grandes con avance < AVANCE_BAJO y sin nuevo devengado en DIAS_PROYECTO días
# CONTEXTO (no es alerta): puesto del GORE Lambayeque entre los gobiernos regionales.
#
# USO:  python scripts/alertas.py            (imprime y escribe los archivos)
# =============================================================================

import csv
import json
import sys
from datetime import date, timedelta
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
SERIE = RAIZ / "data" / "serie"

BRECHA_PUNTOS = 5.0
DIAS_SIN_MOVIMIENTO = 7
MES_CONTROL_CERT = 10
MIN_CERT_PCT = 85.0
PIM_PROYECTO_GRANDE = 1_000_000
AVANCE_BAJO = 50.0
DIAS_PROYECTO = 30


def leer(nombre, carpeta):
    with open(Path(carpeta) / f"{nombre}.csv", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def num(fila, campo):
    return float(fila[campo] or 0)


def foto(filas, fecha):
    return [f for f in filas if f["fecha"] == fecha]


def fecha_base(filas, corte, dias):
    """Última fecha disponible que tenga al menos `dias` de antigüedad (None si no hay historia)."""
    limite = (date.fromisoformat(corte) - timedelta(days=dias)).isoformat()
    previas = [f["fecha"] for f in filas if f["fecha"] <= limite]
    return max(previas) if previas else None


def f_fecha(iso):
    return date.fromisoformat(iso).strftime("%d/%m/%Y")


def soles(v):
    return f"S/ {v:,.0f}"


def calcular(carpeta=SERIE):
    ue = leer("fact_ue_diario", carpeta)
    pliego = leer("fact_pliego_diario", carpeta)
    proy = leer("fact_proyecto_diario", carpeta)
    gore = leer("fact_gore_diario", carpeta)

    corte = max(f["fecha"] for f in ue)
    hoy_ue = foto(ue, corte)
    avance_pliego = num(foto(pliego, corte)[0], "avance_pct")
    alertas = []

    def alerta(regla, nivel, ambito, mensaje, saldo):
        alertas.append({"regla": regla, "nivel": nivel, "ambito": ambito,
                        "mensaje": mensaje, "saldo_por_devengar": round(saldo)})

    base7 = fecha_base(ue, corte, DIAS_SIN_MOVIMIENTO)
    antes = {f["ue_codigo"]: f for f in foto(ue, base7)} if base7 else {}
    for u in hoy_ue:
        nombre = f'{u["ue_codigo"]} {u["ue_nombre"]}'
        pim, dev, av = num(u, "pim"), num(u, "devengado"), num(u, "avance_pct")
        if pim <= 0:
            continue
        saldo = pim - dev
        if avance_pliego - av >= BRECHA_PUNTOS:
            alerta("A1", "alta" if avance_pliego - av >= 2 * BRECHA_PUNTOS else "media", nombre,
                   f"Avance {av:.1f}%, {avance_pliego - av:.1f} puntos por debajo del Pliego ({avance_pliego:.1f}%).", saldo)
        a = antes.get(u["ue_codigo"])
        if a and dev <= num(a, "devengado") and saldo > 0:
            alerta("A2", "media", nombre,
                   f"Sin nuevo devengado desde el {f_fecha(base7)}; quedan {soles(saldo)} por devengar.", saldo)
        cert_pct = num(u, "certificacion") / pim * 100
        if date.fromisoformat(corte).month >= MES_CONTROL_CERT and cert_pct < MIN_CERT_PCT:
            alerta("A3", "alta", nombre,
                   f"Solo {cert_pct:.1f}% del PIM certificado a esta altura del año.", saldo)

    hoy_p = foto(proy, corte)
    base30 = fecha_base(proy, corte, DIAS_PROYECTO)
    antes_p = {(f["ue_codigo"], f["proyecto_codigo"]): f for f in foto(proy, base30)} if base30 else {}
    for p in hoy_p:
        pim, dev, av = num(p, "pim"), num(p, "devengado"), num(p, "avance_pct")
        if pim < PIM_PROYECTO_GRANDE:
            continue
        nombre = f'{p["proyecto_codigo"]} {p["proyecto_nombre"]}'
        a = antes_p.get((p["ue_codigo"], p["proyecto_codigo"]))
        if dev <= 0:
            alerta("A4", "alta", nombre, f"PIM {soles(pim)} sin ningún devengado.", pim)
        elif a and av < AVANCE_BAJO and dev <= num(a, "devengado"):
            alerta("A5", "media", nombre,
                   f"Avance {av:.1f}% y sin nuevo devengado desde el {f_fecha(base30)}.", pim - dev)

    def puesto(fecha):
        orden = sorted(foto(gore, fecha), key=lambda g: num(g, "avance_pct"), reverse=True)
        for i, g in enumerate(orden, 1):
            if g["gore_codigo"] == "452":
                return {"fecha": fecha, "puesto": i, "total": len(orden), "avance_pct": num(g, "avance_pct")}
        return None

    base_g = fecha_base(gore, corte, DIAS_PROYECTO)
    alertas.sort(key=lambda x: (x["nivel"] != "alta", x["regla"], -x["saldo_por_devengar"]))
    return {
        "version": 1,
        "fecha_consulta_mef": corte,
        "avance_pliego_pct": avance_pliego,
        "ranking_gores": {"actual": puesto(corte), "anterior": puesto(base_g) if base_g else None},
        "total": len(alertas),
        "altas": sum(a["nivel"] == "alta" for a in alertas),
        "alertas": alertas,
    }


REGLAS = {
    "A1": "Unidades ejecutoras por debajo del avance del Pliego",
    "A2": "Unidades ejecutoras sin nuevo devengado",
    "A3": "Certificación baja",
    "A4": "Proyectos grandes sin devengado",
    "A5": "Proyectos grandes detenidos",
}


def a_markdown(r):
    l = [f"# Alertas de ejecución — consulta MEF {f_fecha(r['fecha_consulta_mef'])}", "",
         f"Avance del Pliego 452 (solo proyectos): **{r['avance_pliego_pct']:.1f}%** · "
         f"{r['total']} alertas ({r['altas']} de nivel alto)."]
    act, ant = r["ranking_gores"]["actual"], r["ranking_gores"]["anterior"]
    if act:
        t = f"Puesto entre gobiernos regionales por avance: **{act['puesto']} de {act['total']}**"
        if ant:
            t += f" (el {f_fecha(ant['fecha'])}: {ant['puesto']})"
        l += ["", t + "."]
    for cod, titulo in REGLAS.items():
        grupo = [a for a in r["alertas"] if a["regla"] == cod]
        if not grupo:
            continue
        l += ["", f"## {cod} · {titulo} ({len(grupo)})", "", "| Nivel | Ámbito | Detalle |", "|---|---|---|"]
        l += [f"| {a['nivel']} | {a['ambito'].replace('|', '/')} | {a['mensaje']} |" for a in grupo]
    l += ["", "_Umbrales propuestos, pendientes de validar con la jefatura (ver scripts/alertas.py)._", ""]
    return "\n".join(l)


def ejecutar(carpeta=SERIE, destino=RAIZ / "data"):
    r = calcular(carpeta)
    (Path(destino) / "alertas.json").write_text(json.dumps(r, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    (Path(destino) / "alertas.md").write_text(a_markdown(r), encoding="utf-8")
    return r


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    print(a_markdown(ejecutar()))
