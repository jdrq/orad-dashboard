#!/usr/bin/env python3
# =============================================================================
# actualizar_rb_hist_sc.py  (v2 — motor progresivo, escalable mes a mes)
# ORPMI - Gobierno Regional de Lambayeque
# Actualiza: data/rb_hist_sc_enejul.json
#
# DIFERENCIA CON v1: v1 estaba hardcodeada a 3 períodos fijos (T1, T2,
# Julio). Esta versión usa una lista de configuración PERIODOS que vos
# extendés con una sola línea cada vez que cierra un mes nuevo (agosto,
# setiembre, octubre...) — el motor de acumulación progresiva y de
# validación es el mismo para todos, no hay que tocar lógica.
#
# CÓMO AGREGAR UN MES NUEVO (ej. cuando cierre agosto):
#   1) Re-exportar en Consulta Amigable: Mes 8 = agosto, columna Rubro
#      seleccionada, UE 001-855 Sede Central. Guardar como
#      AGOSTO_RUBRO_{año}.xls en xls/historico_rubro/ (para 2022-2025).
#   2) Descomentar (o agregar) la línea correspondiente a "ago" en la
#      lista PERIODOS más abajo.
#   3) Correr: python actualizar_rb_hist_sc.py
#
# CAMPOS QUE GENERA EN EL JSON (acumulado progresivo, Ene -> fin de mes):
#   dev_t1  = acumulado a marzo      (Ene-Mar)
#   dev_t2  = acumulado a junio      (Ene-Jun)
#   dev_sem = acumulado a julio      (Ene-Jul)   <- nombre heredado, ver nota
#   dev_ago = acumulado a agosto     (Ene-Ago)   <- se agrega solo si activás "ago"
#   dev_set = acumulado a setiembre  (Ene-Set)   <- etc.
#
# NOTA sobre el nombre "dev_sem": es heredado de una versión anterior del
# dashboard donde ese campo representaba el semestre (Ene-Jun). Hoy
# representa Ene-Jul. No se renombra para no romper index.html, que ya
# lo consume con ese nombre. Los meses nuevos (ago, set...) usan nombres
# nuevos y correctos desde el inicio: dev_ago, dev_set, etc.
#
# VALIDACIÓN:
#   - Julio (dev_sem) tiene benchmark preexistente en el JSON -> se valida
#     por diferencia porcentual contra ese valor (igual que v1).
#   - Agosto en adelante NO tiene benchmark preexistente (nadie cargó ese
#     dato antes) -> se valida por MONOTONICIDAD: el acumulado de cada
#     mes nuevo debe ser mayor o igual al del mes anterior. Si un mes
#     nuevo da un acumulado MENOR al anterior, es señal casi segura de
#     un error de exportación (grupo mal seleccionado, año equivocado,
#     archivo corrupto) y el script bloquea la escritura para ese año.
#
# USO:
#   python actualizar_rb_hist_sc.py
# =============================================================================

import os
import re
import json
from bs4 import BeautifulSoup

# ---------------- CONFIGURACIÓN ----------------
AÑOS = [2022, 2023, 2024, 2025]
CARPETA_XLS = "xls/historico_rubro"
CARPETA_DATA = "data"
ARCHIVO_JSON = os.path.join(CARPETA_DATA, "rb_hist_sc_enejul.json")

# Lista ordenada de períodos a acumular progresivamente.
# Cada entrada: (clave_campo_json, patrón_de_archivo, tiene_benchmark)
#   - clave_campo_json: nombre del campo en el JSON de salida
#   - patrón_de_archivo: con {año} como placeholder
#   - tiene_benchmark: True solo para "dev_sem" (Julio), que ya existía
#     en el JSON antes de este script y sirve de validación cruzada.
#     Los meses nuevos van con False (se validan por monotonicidad).
PERIODOS = [
    ("dev_t1", "1T_RUBRO_{año}.xls", False),
    ("dev_t2", "2T_RUBRO_{año}.xls", False),
    ("dev_sem", "JULIO_RUBRO_{año}.xls", True),

    # --- Descomentar la línea del mes correspondiente cuando cierre ---
    # ("dev_ago", "AGOSTO_RUBRO_{año}.xls", False),
    # ("dev_set", "SETIEMBRE_RUBRO_{año}.xls", False),
    # ("dev_oct", "OCTUBRE_RUBRO_{año}.xls", False),
    # ("dev_nov", "NOVIEMBRE_RUBRO_{año}.xls", False),
    # ("dev_dic", "DICIEMBRE_RUBRO_{año}.xls", False),
]
# -------------------------------------------------


def limpiar_numero(s):
    """Convierte string con comas y S/ a float."""
    s = str(s).replace(",", "").replace("S/", "").strip()
    try:
        return float(s) if s else 0.0
    except ValueError:
        return 0.0


def parsear_rubros(path):
    """
    Lee un archivo XLS del MEF (HTML disfrazado) filtrado por Rubro,
    a nivel UE 001-855 Sede Central. Retorna el devengado total del
    período (suma de todos los rubros) o None si el archivo no existe
    o no tiene el formato esperado.
    """
    if not os.path.exists(path):
        print(f"   ⚠️  ARCHIVO NO ENCONTRADO: {path}")
        return None

    with open(path, "rb") as f:
        content = f.read()

    soup = BeautifulSoup(content, "html.parser")
    tables = soup.find_all("table")

    if len(tables) < 4:
        print(f"   ⚠️  Formato inesperado: solo {len(tables)} tablas en {path}")
        return None

    tabla_rubros = tables[3]
    dev_total = 0.0
    filas_encontradas = 0

    for r in tabla_rubros.find_all("tr"):
        cols = [c.get_text(strip=True) for c in r.find_all(["td", "th"])]
        if len(cols) >= 8 and re.match(r"^\d{2}:", cols[0]):
            dev_total += limpiar_numero(cols[6])
            filas_encontradas += 1

    if filas_encontradas == 0:
        print(f"   ⚠️  Sin filas de detalle por Rubro en {path} "
              f"(¿se exportó sin la columna 'Rubro' seleccionada?)")
        return None

    return dev_total


def procesar_año(año, data_existente):
    print(f"\n--- Año {año} ---")
    año_str = str(año)
    entrada_actual = data_existente.get(año_str, {})
    entrada_propuesta = entrada_actual.copy()

    acumulado = 0.0
    bloqueado = False
    valor_anterior_acumulado = None  # para el chequeo de monotonicidad

    for clave_campo, patron_archivo, tiene_benchmark in PERIODOS:
        path = os.path.join(CARPETA_XLS, patron_archivo.format(año=año))
        print(f"   Leyendo {clave_campo} ({patron_archivo.format(año=año)})")
        periodo = parsear_rubros(path)

        if periodo is None:
            print(f"   ❌ No se pudo leer {clave_campo} para {año} — "
                  f"se conserva el valor anterior de TODOS los campos de este año.")
            return entrada_actual  # aborta todo el año, no solo el campo

        acumulado += periodo

        # --- Validación por benchmark preexistente (hoy: solo Julio/dev_sem) ---
        if tiene_benchmark:
            benchmark = entrada_actual.get(clave_campo)
            if benchmark is not None:
                diferencia = abs(acumulado - benchmark)
                tolerancia = max(1.0, benchmark * 0.005)  # 0.5%
                if diferencia > tolerancia:
                    print(f"   ⚠️  DIFERENCIA en {clave_campo}: JSON actual="
                          f"{benchmark:,.0f} vs recién calculado={acumulado:,.0f} "
                          f"(diferencia {diferencia:,.0f}). BLOQUEADO.")
                    bloqueado = True
                else:
                    print(f"   ✅ {clave_campo} validado contra benchmark existente "
                          f"(diferencia {diferencia:,.0f})")

        # --- Validación por monotonicidad (meses sin benchmark propio) ---
        if valor_anterior_acumulado is not None and acumulado < valor_anterior_acumulado - 1:
            print(f"   ⚠️  ALERTA: {clave_campo} (S/{acumulado:,.0f}) es MENOR al "
                  f"acumulado anterior (S/{valor_anterior_acumulado:,.0f}). El "
                  f"devengado acumulado no debería bajar. BLOQUEADO.")
            bloqueado = True

        entrada_propuesta[clave_campo] = round(acumulado)
        valor_anterior_acumulado = acumulado

    if bloqueado:
        print(f"   🚫 Año {año}: NO se escriben cambios (algún check falló).")
        return entrada_actual

    print(f"   ✅ Año {año} completo y coherente.")
    return entrada_propuesta


def main():
    print("=" * 70)
    print("Actualización progresiva — Rubro Sede Central")
    print("=" * 70)
    print(f"Períodos configurados: {[p[0] for p in PERIODOS]}")

    if not os.path.exists(ARCHIVO_JSON):
        print(f"❌ No se encontró {ARCHIVO_JSON}. Abortando.")
        return

    with open(ARCHIVO_JSON, "r", encoding="utf-8") as f:
        data = json.load(f)

    for año in AÑOS:
        resultado = procesar_año(año, data)
        if resultado is not None:
            data[str(año)] = resultado

    with open(ARCHIVO_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))

    print("\n" + "=" * 70)
    print(f"✅ Archivo actualizado: {ARCHIVO_JSON}")
    print("=" * 70)


if __name__ == "__main__":
    main()
