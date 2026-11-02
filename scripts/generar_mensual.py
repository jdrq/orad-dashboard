#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# =============================================================================
# generar_mensual.py
# ORAD - Gobierno Regional de Lambayeque
#
# Genera data/mensual-actual.js, la ÚNICA fuente de los valores mensuales que
# alimentan los Bloques 2D (Pliego), 6D (Sede Central) y 9 (proyección).
#
# POR QUÉ EXISTE (02/11/2026): antes esos valores estaban escritos a mano dentro
# de index.html (B9_REAL_BASE, B2M_DEVENGADO_MES, B6M_SC_DEVENGADO_MES) y cada
# cierre de mes obligaba a editar código. Ahora index.html solo los LEE de
# data/mensual-actual.js y este script los calcula a partir de los dos archivos
# mensuales del MEF:
#     xls/gore_devengado_mes.xls   -> Pliego 452            (Bloques 2D y 9)
#     xls/sede_devengado_mes.xls   -> UE 001-855 Sede Central (Bloque 6D)
#
# QUÉ VALIDA ANTES DE ESCRIBIR (cada punto corresponde a un error real que ya
# ocurrió en este proyecto; si algo falla NO se escribe nada):
#   1. El archivo es de GASTO (no de Ingresos: columna "Recaudado").
#   2. El encabezado dice "Sólo Proyectos".
#   3. El archivo de Sede Central trae la fila "Unidad Ejecutora 001-855" y el
#      del Pliego NO la trae (si no, los montos serían los de otro nivel).
#   4. La suma de los meses del archivo cuadra con el totalizador del MEF
#      (tolerancia de 5 soles por redondeo).
#   5. Están todos los meses desde enero hasta el mes que se cierra.
#   6. La consulta se hizo DESPUÉS de terminar el mes que se cierra.
#
# MES PROVISIONAL: el MEF reclasifica meses ya cerrados (pasó con junio, agosto
# y setiembre). Por eso el último mes cerrado queda marcado "provisional" con
# una fecha "revisar_desde" (3 días hábiles después). Si vuelves a correr el
# script pasada esa fecha y el mes no cambió, pasa a "confirmado". Si cambió,
# el cambio se APLICA y queda registrado en "cambios" para que se vea.
#
# USO (desde la raíz del repo):
#   python scripts\generar_mensual.py              cierra el mes anterior al de hoy
#   python scripts\generar_mensual.py --mes 10     fuerza el mes de cierre
#   python scripts\generar_mensual.py --dry-run    valida y muestra, sin escribir
# =============================================================================

import argparse
import json
import os
import re
import sys
from datetime import date, timedelta

from bs4 import BeautifulSoup

# ---------------- CONFIGURACIÓN ----------------
RUTA_PLIEGO = os.path.join("xls", "gore_devengado_mes.xls")
RUTA_SEDE = os.path.join("xls", "sede_devengado_mes.xls")
RUTA_SALIDA = os.path.join("data", "mensual-actual.js")
PREFIJO_JS = "window.MENSUAL_ACTUAL = "
CODIGO_UE = "001-855"
TOLERANCIA_CUADRE = 5          # soles de diferencia aceptados contra el totalizador MEF
DIAS_HABILES_REVISION = 3      # días hábiles para volver a revisar un mes provisional
MAX_CAMBIOS = 60               # tope del historial de cambios guardado en el archivo

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "setiembre", "octubre", "noviembre", "diciembre"]
MESES_LARGO = [m.capitalize() for m in MESES]
# -------------------------------------------------


class ErrorValidacion(Exception):
    """Algo del archivo o del estado no cuadra: no se escribe nada."""


def _num(s):
    s = str(s).replace(",", "").replace("S/", "").strip()
    if re.fullmatch(r"-?\d+(\.\d+)?", s):
        return int(round(float(s)))
    return None


def _mes_de_texto(txt):
    """'9: 'Setiembre' -> 9 ; devuelve None si no es una fila de mes."""
    m = re.match(r"^(\d{1,2}):\s*(.*)$", txt)
    if not m:
        return None
    nombre = re.sub(r"[^a-záéíóúñ]", "", m.group(2).lower()).replace("septiembre", "setiembre")
    if nombre in MESES and MESES.index(nombre) + 1 == int(m.group(1)):
        return int(m.group(1))
    return None


def leer_archivo(ruta, es_sede):
    """
    Lee un archivo mensual del MEF (HTML disfrazado de .xls) y devuelve:
    {"fecha": date, "anio": int, "meses": {n: devengado}, "total": int}
    Lanza ErrorValidacion si el archivo no es el que se espera.
    """
    nombre = os.path.basename(ruta)
    if not os.path.isfile(ruta):
        raise ErrorValidacion(f"No existe {ruta}.")
    with open(ruta, "rb") as f:
        crudo = f.read()
    texto_crudo = crudo.decode("cp1252", errors="replace")
    soup = BeautifulSoup(texto_crudo, "html.parser")
    texto = re.sub(r"\s+", " ", soup.get_text(" "))

    if "Recaudado" in texto_crudo:
        raise ErrorValidacion(f"{nombre} es de INGRESOS (columna 'Recaudado'). "
                              f"Re-exportar desde 'Consulta de Ejecución del Gasto'.")
    if not re.search(r"Ejecuci.n del Gasto", texto):
        raise ErrorValidacion(f"{nombre}: el encabezado no dice 'Ejecución del Gasto'.")
    if not re.search(r"S.lo Proyectos", texto):
        raise ErrorValidacion(f"{nombre}: el encabezado no dice 'Sólo Proyectos' "
                              f"(con 'Actividades y Proyectos' los montos cambian).")

    m = re.search(r"Fecha de la Consulta:\s*(\d{1,2})-(\w+)-(\d{4})", texto)
    if not m:
        raise ErrorValidacion(f"{nombre}: no se encontró la 'Fecha de la Consulta'.")
    mes_txt = m.group(2).lower().replace("septiembre", "setiembre")
    if mes_txt not in MESES:
        raise ErrorValidacion(f"{nombre}: mes de consulta desconocido '{m.group(2)}'.")
    fecha = date(int(m.group(3)), MESES.index(mes_txt) + 1, int(m.group(1)))

    m = re.search(r"A.o de Ejecuci.n:\s*(\d{4})", texto)
    if not m:
        raise ErrorValidacion(f"{nombre}: no se encontró el 'Año de Ejecución'.")
    anio = int(m.group(1))

    total_pliego = total_ue = None
    meses = {}
    for tr in soup.find_all("tr"):
        cols = [c.get_text(strip=True) for c in tr.find_all(["td", "th"])]
        if not cols:
            continue
        c0 = cols[0]
        if c0.startswith("Pliego 452") and len(cols) >= 7:
            total_pliego = _num(cols[6])
        elif c0.startswith(f"Unidad Ejecutora {CODIGO_UE}") and len(cols) >= 7:
            total_ue = _num(cols[6])
        else:
            n = _mes_de_texto(c0)
            if n is not None and len(cols) >= 7:
                dev = _num(cols[6])
                if dev is None:
                    continue
                if n in meses:
                    raise ErrorValidacion(f"{nombre}: el mes {n} aparece repetido.")
                meses[n] = dev

    tiene_ue = total_ue is not None
    if es_sede and not tiene_ue:
        raise ErrorValidacion(f"{nombre}: falta la fila 'Unidad Ejecutora {CODIGO_UE}'. "
                              f"El archivo está a otro nivel (montos distintos). "
                              f"Re-exportar entrando a la UE Sede Central.")
    if not es_sede and tiene_ue:
        raise ErrorValidacion(f"{nombre}: trae la fila 'Unidad Ejecutora {CODIGO_UE}'; "
                              f"el archivo del Pliego debe exportarse sin entrar a una UE.")
    total = total_ue if es_sede else total_pliego
    if total is None:
        raise ErrorValidacion(f"{nombre}: no se encontró la fila totalizadora "
                              f"({'UE ' + CODIGO_UE if es_sede else 'Pliego 452'}).")
    if not meses:
        raise ErrorValidacion(f"{nombre}: no se encontraron filas de meses.")

    return {"fecha": fecha, "anio": anio, "meses": meses, "total": total}


def sumar_habiles(desde, n):
    """Suma n días hábiles (lunes a viernes; no considera feriados)."""
    d = desde
    while n > 0:
        d += timedelta(days=1)
        if d.weekday() < 5:
            n -= 1
    return d


def cargar_previo(ruta=RUTA_SALIDA):
    if not os.path.isfile(ruta):
        return None
    with open(ruta, encoding="utf-8") as f:
        txt = f.read()
    i = txt.find(PREFIJO_JS)
    if i < 0:
        raise ErrorValidacion(f"{ruta} no tiene el formato esperado ('{PREFIJO_JS}...').")
    cuerpo = txt[i + len(PREFIJO_JS):].strip().rstrip(";")
    return json.loads(cuerpo)


def _serie(meses, cierre):
    return [meses.get(n) if n <= cierre else None for n in range(1, 13)]


def generar(mes=None, ruta_pliego=RUTA_PLIEGO, ruta_sede=RUTA_SEDE, ruta_salida=RUTA_SALIDA,
            dry_run=False, forzar=False, hoy=None):
    """Valida y escribe data/mensual-actual.js. Devuelve un dict-resumen."""
    hoy = hoy or date.today()
    avisos = []

    pl = leer_archivo(ruta_pliego, es_sede=False)
    se = leer_archivo(ruta_sede, es_sede=True)

    if pl["anio"] != se["anio"]:
        raise ErrorValidacion(f"Los archivos son de años distintos "
                              f"(Pliego {pl['anio']} / Sede {se['anio']}).")
    anio = pl["anio"]
    fecha = min(pl["fecha"], se["fecha"])
    if pl["fecha"] != se["fecha"]:
        avisos.append(f"Los dos archivos tienen distinta fecha de consulta "
                      f"({pl['fecha']:%d/%m/%Y} vs {se['fecha']:%d/%m/%Y}); "
                      f"se usa la más antigua. Conviene bajarlos el mismo día.")

    # --- mes que se cierra ---
    cierre = mes if mes is not None else fecha.month - 1
    if not 1 <= cierre <= 12:
        raise ErrorValidacion("No hay un mes cerrado para esa fecha de consulta "
                              "(en enero aún no hay mes anterior del mismo año). Usa --mes.")
    # El mes debe haber terminado antes de la consulta
    primer_dia_siguiente = date(anio + (cierre == 12), cierre % 12 + 1, 1)
    if fecha < primer_dia_siguiente:
        raise ErrorValidacion(f"La consulta es del {fecha:%d/%m/%Y}, pero {MESES_LARGO[cierre-1]} "
                              f"termina recién el {primer_dia_siguiente - timedelta(days=1):%d/%m/%Y}. "
                              f"Vuelve a descargar los archivos cuando el mes haya terminado.")

    # --- cuadre contra el totalizador y meses completos ---
    resultado = {}
    for clave, d, etiqueta in (("pliego", pl, "Pliego 452"), ("sede_central", se, f"UE {CODIGO_UE}")):
        falta = [MESES_LARGO[n - 1] for n in range(1, cierre + 1) if n not in d["meses"]]
        if falta:
            raise ErrorValidacion(f"{etiqueta}: faltan meses en el archivo: {', '.join(falta)}.")
        suma = sum(d["meses"].values())
        dif = suma - d["total"]
        if abs(dif) > TOLERANCIA_CUADRE:
            raise ErrorValidacion(f"{etiqueta}: la suma de los meses ({suma:,}) NO cuadra con el "
                                  f"totalizador del MEF ({d['total']:,}); diferencia {dif:,} soles "
                                  f"(tolerancia {TOLERANCIA_CUADRE}).")
        mensual = _serie(d["meses"], cierre)
        acumulado = sum(v for v in mensual if v is not None)
        resultado[clave] = {"mensual": mensual, "totalizador_mef": d["total"],
                            "diferencia_cuadre": dif, "acumulado": acumulado,
                            "mes_cierre": mensual[cierre - 1]}
        if any(v is not None and v < 0 for v in mensual):
            avisos.append(f"{etiqueta}: algún mes tiene devengado NEGATIVO (extornos); revisar.")

    # --- comparación con lo que ya estaba guardado ---
    previo = cargar_previo(ruta_salida)
    if previo is not None and previo.get("anio") != anio:
        avisos.append(f"El archivo previo es del año {previo.get('anio')}; se reinicia para {anio}.")
        previo = None
    cierre_previo = previo.get("ultimo_mes_cerrado") if previo else None
    if previo is not None and cierre < cierre_previo and not forzar:
        raise ErrorValidacion(f"Ya hay un cierre más reciente guardado ({MESES_LARGO[cierre_previo-1]}). "
                              f"Si de verdad quieres retroceder usa --forzar.")

    cambios_nuevos = []
    if previo is not None:
        for clave, ambito in (("pliego", "Pliego"), ("sede_central", "Sede Central")):
            old = previo.get(clave, {}).get("mensual", [None] * 12)
            for n in range(1, min(cierre_previo, cierre) + 1):
                a, b = old[n - 1], resultado[clave]["mensual"][n - 1]
                if a != b:
                    cambios_nuevos.append({"fecha": hoy.isoformat(), "ambito": ambito, "mes": n,
                                           "antes": a, "despues": b,
                                           "diferencia": (b - a) if (a is not None and b is not None) else None})

    # --- estado provisional / confirmado ---
    est_prev = (previo or {}).get("estado") or {}
    if previo is None or cierre != cierre_previo:
        estado = {"provisional": True, "revisar_desde": sumar_habiles(fecha, DIAS_HABILES_REVISION).isoformat(),
                  "confirmado_el": None}
    elif cambios_nuevos:
        estado = {"provisional": True, "revisar_desde": sumar_habiles(hoy, DIAS_HABILES_REVISION).isoformat(),
                  "confirmado_el": None}
    else:
        revisar = est_prev.get("revisar_desde")
        if est_prev.get("provisional") and revisar and hoy >= date.fromisoformat(revisar):
            estado = {"provisional": False, "revisar_desde": None, "confirmado_el": hoy.isoformat()}
        else:
            estado = {"provisional": bool(est_prev.get("provisional", True)),
                      "revisar_desde": est_prev.get("revisar_desde"),
                      "confirmado_el": est_prev.get("confirmado_el")}

    historial = ((previo or {}).get("cambios") or []) + cambios_nuevos
    salida = {
        "anio": anio,
        "generado": hoy.isoformat(),
        "fuente": {"pliego": ruta_pliego.replace("\\", "/"), "sede_central": ruta_sede.replace("\\", "/"),
                   "fecha_consulta": fecha.isoformat()},
        "ultimo_mes_cerrado": cierre,
        "estado": estado,
        "pliego": {k: resultado["pliego"][k] for k in ("mensual", "totalizador_mef", "diferencia_cuadre")},
        "sede_central": {k: resultado["sede_central"][k] for k in ("mensual", "totalizador_mef", "diferencia_cuadre")},
        "cambios": historial[-MAX_CAMBIOS:],
    }

    escrito = False
    if not dry_run:
        cabecera = ("/* ARCHIVO GENERADO por scripts/generar_mensual.py (o scripts/cierre_mensual.py).\n"
                    "   NO EDITAR A MANO: se pisa en cada cierre de mes.\n"
                    "   Lo leen los Bloques 2D, 6D y 9 de index.html. */\n")
        cuerpo = cabecera + PREFIJO_JS + json.dumps(salida, ensure_ascii=False, indent=2) + ";\n"
        tmp = ruta_salida + ".tmp"
        with open(tmp, "w", encoding="utf-8", newline="\n") as f:
            f.write(cuerpo)
        os.replace(tmp, ruta_salida)
        escrito = True

    return {"anio": anio, "cierre": cierre, "fecha_consulta": fecha, "resultado": resultado,
            "estado": estado, "cambios_nuevos": cambios_nuevos, "avisos": avisos,
            "escrito": escrito, "ruta_salida": ruta_salida}


def imprimir_resumen(r):
    m = MESES_LARGO[r["cierre"] - 1]
    print(f"\n  Cierre de {m} {r['anio']}  (consulta del {r['fecha_consulta']:%d/%m/%Y})")
    for clave, etiqueta in (("pliego", "Pliego 452    "), ("sede_central", "Sede Central   ")):
        x = r["resultado"][clave]
        print(f"   {etiqueta} mensual {x['mes_cierre']:>13,}   acumulado {x['acumulado']:>14,}   "
              f"MEF {x['totalizador_mef']:>14,}   dif. {x['diferencia_cuadre']:+d}")
    for c in r["cambios_nuevos"]:
        dif = f"{c['diferencia']:+,}" if c["diferencia"] is not None else "n/d"
        marca = "  <-- mes que se consideraba inamovible" if c["mes"] <= 7 else ""
        print(f"   [!!] RECLASIFICACION MEF en {c['ambito']} / {MESES_LARGO[c['mes']-1]}: "
              f"{c['antes']:,} -> {c['despues']:,} ({dif}){marca}")
    est = r["estado"]
    if est["provisional"]:
        print(f"   Estado: PROVISIONAL. Volver a bajar los 2 archivos mensuales y correr de nuevo "
              f"desde el {date.fromisoformat(est['revisar_desde']):%d/%m/%Y}.")
    else:
        print(f"   Estado: CONFIRMADO el {est['confirmado_el']}.")
    for a in r["avisos"]:
        print(f"   [aviso] {a}")
    print(f"   {'Escrito' if r['escrito'] else 'NO escrito (dry-run)'}: {r['ruta_salida']}")


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="Genera data/mensual-actual.js desde los archivos mensuales del MEF.")
    ap.add_argument("--mes", type=int, help="mes que se cierra (1-12). Por defecto: el anterior al de la consulta.")
    ap.add_argument("--dry-run", action="store_true", help="valida y muestra, sin escribir.")
    ap.add_argument("--forzar", action="store_true", help="permite retroceder a un mes anterior al ya guardado.")
    args = ap.parse_args()
    try:
        imprimir_resumen(generar(mes=args.mes, dry_run=args.dry_run, forzar=args.forzar))
    except ErrorValidacion as e:
        print(f"\n  [ERROR] {e}\n  No se escribió nada.")
        sys.exit(1)


if __name__ == "__main__":
    main()
