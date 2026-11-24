#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# =============================================================================
# controles_calidad.py  —  controles de calidad de los 17 archivos del MEF
# ORAD - Gobierno Regional de Lambayeque
#
# QUÉ HACE
#   Revisa los archivos de xls/ con reglas que un analista haría "a ojo" y deja
#   el resultado (alertas / avisos) en data/estado.json (banner del dashboard) y
#   un registro diario en data/calidad_historial.json (serie de datos para el informe).
#   NO bloquea la actualización: las anomalías se INFORMAN. Lo que sí bloquea
#   (archivo ausente, equivocado o de otro día) ya lo hace descargar_xls_mef.py.
#
# CONTROLES (id · severidad)
#   C01 jerarquía de montos en cada fila: Certificación <= PIM, Compromiso <= Certificación,
#       Atención <= Compromiso, Devengado <= Atención, Girado <= Devengado ........... ALERTA
#   C02 el "Avance %" coincide con Devengado / PIM ................................. ALERTA
#   C03 la suma de las filas hijas cuadra con su total (UE, rubros, funciones...) ..... ALERTA
#   C04 el Pliego 452 tiene el mismo PIM y Devengado en todos los archivos .......... ALERTA
#   C05 cada archivo "por unidad ejecutora" coincide con su fila en ue_pliego ....... ALERTA
#   C06 el devengado del Pliego NO disminuye respecto al día anterior ............... ALERTA
#   C07 el tamaño de cada archivo no cambia más de 50% respecto al día anterior ..... ALERTA
#   C08 el devengado no sube de forma anormal en un día (> 3% del PIM por día) ....... AVISO
#   C09 el PIM del Pliego cambió (modificación presupuestal) ........................ AVISO
#   C10 faltan días entre una consulta y la anterior ................................ AVISO
#   C11 cambió la cantidad de unidades ejecutoras ................................... AVISO
#   (C06 a C11 comparan contra el día anterior guardado en calidad_historial.json.
#    El primer día no hay referencia: se omiten y queda anotado.)
#
# USO (desde la raíz del repo)
#   python scripts\controles_calidad.py              muestra el informe (no escribe nada)
#   python scripts\controles_calidad.py --guardar    además actualiza data/calidad_historial.json
# El workflow diario lo ejecuta a través de generar_estado.py.
# =============================================================================

import argparse
import json
import re
import sys
from datetime import date, datetime
from pathlib import Path

from bs4 import BeautifulSoup

RAIZ = Path(__file__).resolve().parent.parent
CARPETA_XLS = RAIZ / "xls"
RUTA_HISTORIAL = RAIZ / "data" / "calidad_historial.json"

TOL_FILA = 2            # soles: redondeos dentro de una fila
TOL_CUADRE = 5          # soles: redondeos al sumar filas (igual que el cuadre mensual)
TOL_AVANCE = 0.15       # puntos porcentuales
UMBRAL_SALTO_DIARIO = 0.03    # 3% del PIM por día
UMBRAL_TAMANIO = 0.50         # 50% de cambio de tamaño
MAX_DIAS_HISTORIAL = 400

# Columnas numéricas de cada fila: PIA, PIM, Certificación, Compromiso Anual,
# Atención de Compromiso Mensual, Devengado, Girado, Avance %
PIM, CERT, COMP, ATEN, DEV, GIR, AVAN = 1, 2, 3, 4, 5, 6, 7

MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7,
         "agosto": 8, "setiembre": 9, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}
RE_ENCABEZADO = re.compile(r"^(TOTAL|Nivel de Gobierno|Sector \d|Pliego \d|Unidad Ejecutora \d|Funci[oó]n|Departamento)")
RE_CODIGO_UE = re.compile(r"(\d{3}-\d{3,4})")


# ----------------------------------------------------------------------------
# Lectura de un archivo XLS del MEF (en realidad HTML en cp1252)
# ----------------------------------------------------------------------------
def _num(s):
    try:
        return float(s.replace(",", "").strip())
    except ValueError:
        return None


def leer(ruta):
    texto = Path(ruta).read_bytes().decode("cp1252", errors="replace")
    sopa = BeautifulSoup(texto, "html.parser")
    info = {"archivo": Path(ruta).name, "bytes": Path(ruta).stat().st_size,
            "fecha": None, "incluye": None, "filas": []}
    for tr in sopa.find_all("tr"):
        celdas = [re.sub(r"\s+", " ", td.get_text(" ", strip=True)) for td in tr.find_all(["td", "th"])]
        if not celdas:
            continue
        if info["fecha"] is None and "Fecha de la Consulta" in celdas[0]:
            m = re.search(r"Consulta:\s*(\d{1,2})-([A-Za-z]+)-(\d{4})", celdas[0])
            if m and m.group(2).lower() in MESES:
                info["fecha"] = date(int(m.group(3)), MESES[m.group(2).lower()], int(m.group(1)))
            m = re.search(r"Incluye:\s*(.+)$", celdas[0])
            info["incluye"] = m.group(1).strip() if m else None
            continue
        if len(celdas) >= 9:
            n = [_num(x) for x in celdas[1:9]]
            if all(v is not None for v in n):
                info["filas"].append({"etiqueta": celdas[0], "n": n,
                                      "encabezado": bool(RE_ENCABEZADO.match(celdas[0]))})
    return info


def leer_todos(carpeta=CARPETA_XLS):
    return {p.name: leer(p) for p in sorted(Path(carpeta).glob("*.xls"))}


# ----------------------------------------------------------------------------
# Controles
# ----------------------------------------------------------------------------
class Resultado:
    def __init__(self):
        self.controles = {}      # id -> {"nombre", "severidad", "revisados", "fallos"}
        self.detalle = []        # [{"id","severidad","mensaje"}]
        self.notas = []

    def definir(self, cid, nombre, severidad):
        self.controles[cid] = {"nombre": nombre, "severidad": severidad, "revisados": 0, "fallos": 0}

    def revisar(self, cid, ok, mensaje=""):
        c = self.controles[cid]
        c["revisados"] += 1
        if not ok:
            c["fallos"] += 1
            self.detalle.append({"id": cid, "severidad": c["severidad"], "mensaje": mensaje})

    @property
    def alertas(self):
        return sum(1 for d in self.detalle if d["severidad"] == "alerta")

    @property
    def avisos(self):
        return sum(1 for d in self.detalle if d["severidad"] == "aviso")

    @property
    def revisados(self):
        return sum(c["revisados"] for c in self.controles.values())


def _s(v):
    return f"S/ {v:,.0f}"


def _fila_pliego(info):
    for f in info["filas"]:
        if f["etiqueta"].startswith("Pliego 452"):
            return f
    return None


def _definir_controles(r):
    r.definir("C01", "Jerarquía de montos por fila", "alerta")
    r.definir("C02", "Avance % coherente con Devengado/PIM", "alerta")
    r.definir("C03", "Suma de filas hijas = total", "alerta")
    r.definir("C04", "Pliego 452 idéntico entre archivos", "alerta")
    r.definir("C05", "Unidades ejecutoras coinciden con ue_pliego", "alerta")
    r.definir("C06", "Devengado no disminuye vs. día anterior", "alerta")
    r.definir("C07", "Tamaño de archivo estable vs. día anterior", "alerta")
    r.definir("C08", "Salto diario de devengado normal", "aviso")
    r.definir("C09", "PIM sin modificaciones", "aviso")
    r.definir("C10", "Sin días faltantes entre consultas", "aviso")
    r.definir("C11", "Cantidad de unidades ejecutoras estable", "aviso")


def controles_del_dia(archivos):
    """C01-C05: usan solo los archivos de hoy."""
    r = Resultado()
    _definir_controles(r)

    # C01 y C02: cada fila con 8 cifras
    for nombre, info in archivos.items():
        for f in info["filas"]:
            n, et = f["n"], f["etiqueta"][:55]
            reglas = (("Certificación > PIM", n[CERT] > n[PIM] + TOL_FILA),
                      ("Compromiso > Certificación", n[COMP] > n[CERT] + TOL_FILA),
                      ("Atención > Compromiso", n[ATEN] > n[COMP] + TOL_FILA),
                      ("Devengado > Atención", n[DEV] > n[ATEN] + TOL_FILA),
                      ("Girado > Devengado", n[GIR] > n[DEV] + TOL_FILA))
            mal = [txt for txt, falla in reglas if falla]
            r.revisar("C01", not mal, f"{nombre} · {et}: {', '.join(mal)}")
            if n[PIM] > 0:
                calc = n[DEV] / n[PIM] * 100
                r.revisar("C02", abs(n[AVAN] - calc) <= TOL_AVANCE,
                          f"{nombre} · {et}: avance {n[AVAN]:.1f}% pero Devengado/PIM = {calc:.1f}%")

    solo_proy = {k: v for k, v in archivos.items() if (v["incluye"] or "").lower().startswith("sólo proyectos")
                 or (v["incluye"] or "").lower().startswith("solo proyectos")}

    # C03: suma de hijas = último encabezado (se omiten archivos sin filas hijas o con otro filtro)
    for nombre, info in solo_proy.items():
        enc = [f for f in info["filas"] if f["encabezado"]]
        hijas = [f for f in info["filas"] if not f["encabezado"]]
        if not enc or not hijas:
            continue
        total = enc[-1]
        for col, txt in ((PIM, "PIM"), (DEV, "Devengado")):
            suma = sum(h["n"][col] for h in hijas)
            dif = suma - total["n"][col]
            r.revisar("C03", abs(dif) <= TOL_CUADRE,
                      f"{nombre}: la suma de {len(hijas)} filas ({txt} {_s(suma)}) no cuadra con "
                      f"«{total['etiqueta'][:35]}» ({_s(total['n'][col])}); diferencia {dif:+,.0f}")

    # C04: Pliego 452 idéntico entre archivos "Sólo Proyectos"
    ref = _fila_pliego(solo_proy.get("ue_pliego.xls", {"filas": []}))
    if ref:
        for nombre, info in solo_proy.items():
            if nombre == "ue_pliego.xls":
                continue
            f = _fila_pliego(info)
            if not f:
                for g in info["filas"]:               # nacional_gores lo trae como "452: ..."
                    if g["etiqueta"].startswith("452:"):
                        f = g
                        break
            if not f:
                continue
            for col, txt in ((PIM, "PIM"), (DEV, "Devengado")):
                dif = f["n"][col] - ref["n"][col]
                r.revisar("C04", abs(dif) <= TOL_CUADRE,
                          f"{nombre}: Pliego 452 {txt} {_s(f['n'][col])} ≠ ue_pliego {_s(ref['n'][col])} ({dif:+,.0f})")

    # C05: archivos "por unidad ejecutora" vs su fila en ue_pliego
    ues = {}
    for f in solo_proy.get("ue_pliego.xls", {"filas": []})["filas"]:
        m = re.match(r"(\d{3}-\d{3,4}):", f["etiqueta"])
        if m:
            ues[m.group(1)] = f
    for nombre, info in solo_proy.items():
        if nombre == "ue_pliego.xls":
            continue
        for f in info["filas"]:
            if f["etiqueta"].startswith("Unidad Ejecutora"):
                m = RE_CODIGO_UE.search(f["etiqueta"])
                if m and m.group(1) in ues:
                    for col, txt in ((PIM, "PIM"), (DEV, "Devengado")):
                        dif = f["n"][col] - ues[m.group(1)]["n"][col]
                        r.revisar("C05", abs(dif) <= TOL_CUADRE,
                                  f"{nombre}: UE {m.group(1)} {txt} {_s(f['n'][col])} ≠ ue_pliego "
                                  f"{_s(ues[m.group(1)]['n'][col])} ({dif:+,.0f})")
    return r, ues


def resumen_del_dia(archivos, ues):
    """Cifras que se guardan en el historial (y contra las que se compara mañana)."""
    ue_p = archivos.get("ue_pliego.xls")
    pl = _fila_pliego(ue_p) if ue_p else None
    fecha = ue_p["fecha"] if ue_p else None
    return {
        "fecha": fecha.isoformat() if fecha else None,
        "pliego": ({"pim": pl["n"][PIM], "certificacion": pl["n"][CERT], "devengado": pl["n"][DEV],
                    "girado": pl["n"][GIR], "avance": pl["n"][AVAN]} if pl else None),
        "ue_cantidad": len(ues),
        "tamanios": {k: v["bytes"] for k, v in archivos.items()},
    }


def controles_vs_anterior(r, hoy, historial):
    """C06-C11: comparan con el último día guardado anterior a la fecha de hoy."""
    if not hoy["fecha"] or not hoy["pliego"]:
        r.notas.append("Sin ue_pliego.xls válido: no se pudieron hacer las comparaciones diarias.")
        return None
    previos = [d for d in historial.get("dias", []) if d.get("fecha") and d["fecha"] < hoy["fecha"] and d.get("pliego")]
    if not previos:
        r.notas.append("Primer día registrado: aún no hay referencia anterior para C06-C11.")
        return None
    ant = max(previos, key=lambda d: d["fecha"])
    dias = (date.fromisoformat(hoy["fecha"]) - date.fromisoformat(ant["fecha"])).days
    p, q = hoy["pliego"], ant["pliego"]

    # C06 / C08: devengado
    dif = p["devengado"] - q["devengado"]
    r.revisar("C06", dif >= -TOL_CUADRE,
              f"El devengado del Pliego bajó de {_s(q['devengado'])} ({ant['fecha']}) a {_s(p['devengado'])}: "
              f"{dif:+,.0f}. El MEF puede reclasificar, pero conviene revisarlo.")
    if p["pim"] > 0:
        limite = UMBRAL_SALTO_DIARIO * p["pim"] * max(dias, 1)
        r.revisar("C08", dif <= limite,
                  f"El devengado subió {_s(dif)} desde el {ant['fecha']} ({dif / p['pim'] * 100:.1f}% del PIM en {dias} día(s)): "
                  f"más de lo habitual.")
    # C09: PIM
    dpim = p["pim"] - q["pim"]
    r.revisar("C09", abs(dpim) <= TOL_CUADRE,
              f"El PIM del Pliego cambió {dpim:+,.0f} desde el {ant['fecha']} ({_s(q['pim'])} → {_s(p['pim'])}).")
    # C10: días faltantes
    r.revisar("C10", dias <= 1,
              f"Pasaron {dias} días entre la consulta del {ant['fecha']} y la del {hoy['fecha']}: faltan registros intermedios.")
    # C11: unidades ejecutoras
    if ant.get("ue_cantidad") is not None:
        r.revisar("C11", hoy["ue_cantidad"] == ant["ue_cantidad"],
                  f"Cambió la cantidad de unidades ejecutoras: {ant['ue_cantidad']} → {hoy['ue_cantidad']}.")
    # C07: tamaño de archivos
    for nombre, tam in hoy["tamanios"].items():
        previo = (ant.get("tamanios") or {}).get(nombre)
        if previo:
            cambio = abs(tam - previo) / previo
            r.revisar("C07", cambio <= UMBRAL_TAMANIO,
                      f"{nombre}: el tamaño cambió {cambio * 100:.0f}% ({previo:,} → {tam:,} bytes) desde el {ant['fecha']}.")
    return ant


# ----------------------------------------------------------------------------
# Historial
# ----------------------------------------------------------------------------
def cargar_historial(ruta=RUTA_HISTORIAL):
    try:
        return json.loads(Path(ruta).read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return {"version": 1, "descripcion": "Registro diario de controles de calidad (ORAD). Una entrada por fecha de consulta.", "dias": []}


def guardar_historial(historial, hoy, r, ruta=RUTA_HISTORIAL):
    entrada = dict(hoy)
    entrada["controles"] = {"revisados": r.revisados, "alertas": r.alertas, "avisos": r.avisos}
    entrada["registrado"] = datetime.now().astimezone().replace(microsecond=0).isoformat()
    dias = [d for d in historial.get("dias", []) if d.get("fecha") != hoy["fecha"]]
    dias.append(entrada)
    dias.sort(key=lambda d: d.get("fecha") or "")
    historial["dias"] = dias[-MAX_DIAS_HISTORIAL:]
    Path(ruta).parent.mkdir(parents=True, exist_ok=True)
    Path(ruta).write_text(json.dumps(historial, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")


# ----------------------------------------------------------------------------
# Punto de entrada reutilizable
# ----------------------------------------------------------------------------
def ejecutar(guardar=False, carpeta_xls=CARPETA_XLS, ruta_historial=RUTA_HISTORIAL):
    archivos = leer_todos(carpeta_xls)
    r, ues = controles_del_dia(archivos)
    hoy = resumen_del_dia(archivos, ues)
    historial = cargar_historial(ruta_historial)
    controles_vs_anterior(r, hoy, historial)
    if guardar and hoy["fecha"]:
        guardar_historial(historial, hoy, r, ruta_historial)
    return {
        "revisados": r.revisados,
        "alertas": r.alertas,
        "avisos": r.avisos,
        "detalle": r.detalle[:20],            # el estado.json solo guarda lo más importante
        "notas": r.notas,
        "_resultado": r,
        "_resumen": hoy,
    }


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="Controles de calidad de los archivos del MEF.")
    ap.add_argument("--guardar", action="store_true", help="actualiza data/calidad_historial.json")
    args = ap.parse_args()
    res = ejecutar(guardar=args.guardar)
    r = res["_resultado"]
    print(f"CONTROLES DE CALIDAD — consulta MEF {res['_resumen']['fecha']}\n")
    for cid, c in sorted(r.controles.items()):
        marca = "OK " if c["fallos"] == 0 else ("ALERTA" if c["severidad"] == "alerta" else "AVISO ")
        print(f"  {cid} {marca:6s} {c['nombre']:48s} revisados={c['revisados']:4d} fallos={c['fallos']}")
    print(f"\nTotal: {res['revisados']} verificaciones · {res['alertas']} alerta(s) · {res['avisos']} aviso(s)")
    for d in r.detalle:
        print(f"  [{d['severidad'].upper()}] {d['id']}: {d['mensaje']}")
    for n in r.notas:
        print(f"  (nota) {n}")
    if args.guardar:
        print("\nHistorial actualizado:", RUTA_HISTORIAL)


if __name__ == "__main__":
    main()
