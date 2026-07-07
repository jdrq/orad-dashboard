#!/usr/bin/env python3
# =============================================================================
# actualizar_hist_gores_enejul.py
# ORPMI - Gobierno Regional de Lambayeque
# Actualiza: data/historico_enejul.json  (Bloque 8 — Ranking nacional GORES)
#
# CONTEXTO: este JSON alimentaba antes de forma MANUAL (armado a mano en
# sesión de trabajo). Este script lo formaliza como parte del pipeline
# reproducible, con la misma filosofía progresiva que
# actualizar_rb_hist_sc.py v2 (Bloque 6 — Rubro Sede Central):
#   - Lista PERIODOS configurable: agregar un mes nuevo = una línea, no
#     reescribir lógica.
#   - Validación cruzada contra el valor ya existente en el JSON (que fue
#     armado a mano y sirve de benchmark de confianza para Julio).
#   - Validación de monotonicidad para meses sin benchmark previo (Agosto
#     en adelante).
#   - Si algo falla, NO se sobrescribe — se conserva el año tal cual estaba.
#
# ARCHIVOS REQUERIDOS (en xls/historico/, ya existentes en el repo):
#   1T_{año}.xls, 2T_{año}.xls, julio_{año}.xls, anual_{año}_gores.xls
#   (para 2022-2025)
#
# CÓMO AGREGAR UN MES NUEVO (ej. cuando cierre agosto):
#   1) Re-exportar en Consulta Amigable: Sector 99 (todos los GOREs), sin
#      Pliego específico, filtro Mes 8 = agosto. Guardar como
#      agosto_{año}.xls en xls/historico/ (para 2022-2025).
#   2) Agregar la línea correspondiente a "agosto" en PERIODOS más abajo.
#   3) Correr: python actualizar_hist_gores_enejul.py
#
# USO:
#   python actualizar_hist_gores_enejul.py
# =============================================================================

import os
import re
import json
from datetime import datetime
from bs4 import BeautifulSoup

# ---------------- CONFIGURACIÓN ----------------
AÑOS = [2022, 2023, 2024, 2025]
CARPETA_XLS = "xls/historico"
CARPETA_DATA = "data"
ARCHIVO_JSON = os.path.join(CARPETA_DATA, "historico_enejul.json")
CODIGO_LAMBAYEQUE = "452"

# Períodos a acumular progresivamente para dev y cert.
# (etiqueta, patrón_de_archivo, tiene_benchmark_en_json_existente)
# Julio SÍ tiene benchmark porque el JSON ya fue armado a mano con datos
# Ene-Jul en la sesión anterior. Meses nuevos (agosto...) NO lo tienen,
# se validan solo por monotonicidad.
PERIODOS = [
    ("T1", "1T_{año}.xls", False),
    ("T2", "2T_{año}.xls", False),
    ("JULIO", "julio_{año}.xls", True),

    # --- Descomentar cuando cierre el mes correspondiente ---
    # ("AGOSTO", "agosto_{año}.xls", False),
    # ("SETIEMBRE", "setiembre_{año}.xls", False),
    # ("OCTUBRE", "octubre_{año}.xls", False),
    # ("NOVIEMBRE", "noviembre_{año}.xls", False),
    # ("DICIEMBRE", "diciembre_{año}.xls", False),
]

ARCHIVO_ANUAL = "anual_{año}_gores.xls"  # fuente del PIM (fijo, no se acumula)

# Etiqueta legible para el campo "label" del JSON, según el último período activo
NOMBRE_MES_FINAL = {"T1": "Marzo", "T2": "Junio", "JULIO": "Julio",
                     "AGOSTO": "Agosto", "SETIEMBRE": "Setiembre",
                     "OCTUBRE": "Octubre", "NOVIEMBRE": "Noviembre",
                     "DICIEMBRE": "Diciembre"}
# -------------------------------------------------


def limpiar_numero(s):
    """Convierte string con comas y S/ a float."""
    s = str(s).replace(",", "").replace("S/", "").strip()
    try:
        return float(s) if s else 0.0
    except ValueError:
        return 0.0


def parsear_gores(path):
    """
    Lee un archivo XLS del MEF (HTML disfrazado) con el ranking de 26
    GOREs (Sector 99, sin Pliego específico). Estructura confirmada
    (misma que usa convertir_semestral.py):
      Tabla 3 = detalle por Pliego/GORE. Columnas:
        [0] Nombre (código: nombre), [1] PIA, [2] PIM, [3] Certificación,
        [4] Compromiso Anual, [5] Atención Comp. Mensual,
        [6] Devengado, [7] Girado, [8] Avance %

    Retorna: dict {codigo: {nombre, pim, cert, dev}} o None si falla.
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

    tabla = tables[3]
    gores = {}

    for r in tabla.find_all("tr"):
        cols = [c.get_text(strip=True) for c in r.find_all(["td", "th"])]
        if len(cols) >= 7 and re.match(r"^\d{3}:", cols[0]):
            codigo = cols[0][:3]
            gores[codigo] = {
                "codigo": codigo,
                "nombre": cols[0],
                "pim": limpiar_numero(cols[2]),
                "cert": limpiar_numero(cols[3]),
                "dev": limpiar_numero(cols[6]),
            }

    if not gores:
        print(f"   ⚠️  Sin filas de GOREs reconocidas en {path}")
        return None

    return gores


def construir_ranking(gores_dict, metrica):
    """Ordena los GOREs por una métrica (dev|dev_pct|cert_pct) desc y
    asigna 'posicion' dentro de ese ranking específico."""
    lista = sorted(gores_dict.values(), key=lambda g: g[metrica], reverse=True)
    for i, g in enumerate(lista):
        g_copia = dict(g)
        g_copia["posicion"] = i + 1
        lista[i] = g_copia
    return lista


def procesar_año(año, data_existente):
    print(f"\n--- Año {año} ---")
    año_str = str(año)
    entrada_actual = data_existente.get("semestres", {}).get(año_str, {})

    # --- 1) PIM anual (fijo, no se acumula) ---
    path_anual = os.path.join(CARPETA_XLS, ARCHIVO_ANUAL.format(año=año))
    print(f"   Leyendo PIM anual: {path_anual}")
    gores_anual = parsear_gores(path_anual)
    if gores_anual is None:
        print(f"   ❌ No se pudo leer el PIM anual de {año}. Año sin cambios.")
        return entrada_actual

    # --- 2) Acumulación progresiva de dev y cert ---
    acumulado = {cod: {"dev": 0.0, "cert": 0.0, "nombre": g["nombre"]}
                 for cod, g in gores_anual.items()}
    ultimo_periodo_activo = None
    dev_lambayeque_anterior = None
    bloqueado = False

    for etiqueta, patron, tiene_benchmark in PERIODOS:
        path = os.path.join(CARPETA_XLS, patron.format(año=año))
        print(f"   Leyendo {etiqueta} ({patron.format(año=año)})")
        gores_periodo = parsear_gores(path)

        if gores_periodo is None:
            print(f"   ❌ No se pudo leer {etiqueta} para {año}. Año sin cambios.")
            return entrada_actual

        for cod, g in gores_periodo.items():
            if cod not in acumulado:
                acumulado[cod] = {"dev": 0.0, "cert": 0.0, "nombre": g["nombre"]}
            acumulado[cod]["dev"] += g["dev"]
            acumulado[cod]["cert"] += g["cert"]

        ultimo_periodo_activo = etiqueta
        dev_lambayeque_actual = acumulado.get(CODIGO_LAMBAYEQUE, {}).get("dev", 0.0)

        # --- Validación por benchmark preexistente (hoy: solo Julio) ---
        if tiene_benchmark:
            dev_benchmark = entrada_actual.get("lambayeque", {}).get("dev")
            if dev_benchmark is not None:
                diferencia = abs(dev_lambayeque_actual - dev_benchmark)
                tolerancia = max(1.0, dev_benchmark * 0.005)
                if diferencia > tolerancia:
                    print(f"   ⚠️  DIFERENCIA en {etiqueta}: JSON actual dev(Lambayeque)="
                          f"{dev_benchmark:,.0f} vs recién calculado={dev_lambayeque_actual:,.0f} "
                          f"(diferencia {diferencia:,.0f}). BLOQUEADO.")
                    bloqueado = True
                else:
                    print(f"   ✅ {etiqueta} validado contra benchmark existente "
                          f"(diferencia {diferencia:,.0f})")

            # También valida el promedio nacional ponderado (prom_dev_pct),
            # no solo Lambayeque — cierra el punto ciego detectado en
            # sesión de trabajo (07/07/2026): antes solo se validaba
            # Lambayeque y una diferencia de metodología en el promedio
            # nacional pasaba sin bloquear.
            prom_benchmark = entrada_actual.get("prom_dev_pct")
            if prom_benchmark is not None:
                suma_dev_check = sum(acc["dev"] for acc in acumulado.values())
                suma_pim_check = sum(gores_anual.get(cod, {}).get("pim", 0) for cod in acumulado)
                prom_actual = (suma_dev_check / suma_pim_check * 100) if suma_pim_check else 0.0
                diferencia_prom = abs(prom_actual - prom_benchmark)
                if diferencia_prom > 0.5:  # medio punto porcentual de tolerancia
                    print(f"   ⚠️  DIFERENCIA en prom_dev_pct: JSON actual={prom_benchmark}% "
                          f"vs recién calculado={prom_actual:.1f}% "
                          f"(diferencia {diferencia_prom:.1f}pp). BLOQUEADO.")
                    bloqueado = True
                else:
                    print(f"   ✅ prom_dev_pct validado contra benchmark existente "
                          f"(diferencia {diferencia_prom:.1f}pp)")

        # --- Validación por monotonicidad ---
        if dev_lambayeque_anterior is not None and dev_lambayeque_actual < dev_lambayeque_anterior - 1:
            print(f"   ⚠️  ALERTA: devengado de Lambayeque bajó de "
                  f"S/{dev_lambayeque_anterior:,.0f} a S/{dev_lambayeque_actual:,.0f} "
                  f"en {etiqueta}. BLOQUEADO.")
            bloqueado = True

        dev_lambayeque_anterior = dev_lambayeque_actual

    if bloqueado:
        print(f"   🚫 Año {año}: NO se escriben cambios (algún check falló).")
        return entrada_actual

    # --- 3) Calcular pct y armar la lista final por GORE ---
    lista_gores = []
    for cod, acc in acumulado.items():
        pim = gores_anual.get(cod, {}).get("pim", 0.0)
        dev = acc["dev"]
        cert = acc["cert"]
        dev_pct = round((dev / pim * 100), 2) if pim else 0.0
        cert_pct = round((cert / pim * 100), 2) if pim else 0.0
        lista_gores.append({
            "codigo": cod,
            "nombre": acc["nombre"],
            "pim": round(pim),
            "dev": round(dev),
            "cert": round(cert),
            "dev_pct": dev_pct,
            "cert_pct": cert_pct,
            "es_lambayeque": cod == CODIGO_LAMBAYEQUE,
        })

    total_gores = len(lista_gores)

    # --- 4) Construir los 3 rankings y asignar posiciones cruzadas ---
    ranking_dev_sol = construir_ranking({g["codigo"]: g for g in lista_gores}, "dev")
    ranking_dev_pct = construir_ranking({g["codigo"]: g for g in lista_gores}, "dev_pct")
    ranking_cert_pct = construir_ranking({g["codigo"]: g for g in lista_gores}, "cert_pct")

    pos_dev_sol = {g["codigo"]: g["posicion"] for g in ranking_dev_sol}
    pos_dev_pct = {g["codigo"]: g["posicion"] for g in ranking_dev_pct}
    pos_cert_pct = {g["codigo"]: g["posicion"] for g in ranking_cert_pct}

    def enriquecer(lista):
        enriquecida = []
        for g in lista:
            g2 = dict(g)
            g2["pos_dev_sol"] = pos_dev_sol[g["codigo"]]
            g2["pos_dev_pct"] = pos_dev_pct[g["codigo"]]
            g2["pos_cert_pct"] = pos_cert_pct[g["codigo"]]
            enriquecida.append(g2)
        return enriquecida

    ranking_dev_sol = enriquecer(ranking_dev_sol)
    ranking_dev_pct = enriquecer(ranking_dev_pct)
    ranking_cert_pct = enriquecer(ranking_cert_pct)

    lam = next((g for g in ranking_dev_sol if g["es_lambayeque"]), {})

    # Promedio NACIONAL ponderado por PIM (no promedio simple de los 26 %).
    # Responde "¿cuánto del presupuesto nacional ya se ejecutó?", que es
    # la métrica que usa el propio MEF y la que ya traía el dato manual
    # anterior (confirmado por auditoría: Σdev/Σpim reproduce exacto el
    # valor histórico, un promedio simple NO).
    suma_dev_nacional = sum(g["dev"] for g in lista_gores)
    suma_cert_nacional = sum(g["cert"] for g in lista_gores)
    suma_pim_nacional = sum(g["pim"] for g in lista_gores)
    prom_dev_pct = round(suma_dev_nacional / suma_pim_nacional * 100, 1) if suma_pim_nacional else 0.0
    prom_cert_pct = round(suma_cert_nacional / suma_pim_nacional * 100, 1) if suma_pim_nacional else 0.0

    mes_final = NOMBRE_MES_FINAL.get(ultimo_periodo_activo, ultimo_periodo_activo)
    label_mes_abrev = {"Marzo": "Mar", "Junio": "Jun", "Julio": "Jul", "Agosto": "Ago",
                        "Setiembre": "Set", "Octubre": "Oct", "Noviembre": "Nov",
                        "Diciembre": "Dic"}.get(mes_final, mes_final)

    print(f"   ✅ Año {año} completo: Lambayeque dev=S/{lam.get('dev', 0):,.0f} "
          f"({lam.get('dev_pct', 0)}%), posición {lam.get('pos_dev_sol', '?')}°/{total_gores}")

    return {
        "label": f"Ene–{label_mes_abrev} {año}",
        "nota": f"Acumulado T1+T2{'+ ' + mes_final if mes_final not in ('Marzo', 'Junio') else ''}",
        "total_gores": total_gores,
        "prom_dev_pct": prom_dev_pct,
        "prom_cert_pct": prom_cert_pct,
        "lambayeque": {
            "codigo": CODIGO_LAMBAYEQUE,
            "nombre": lam.get("nombre", ""),
            "pim": lam.get("pim", 0),
            "dev": lam.get("dev", 0),
            "cert": lam.get("cert", 0),
            "dev_pct": lam.get("dev_pct", 0),
            "cert_pct": lam.get("cert_pct", 0),
            "es_lambayeque": True,
            "pos_dev_sol": lam.get("pos_dev_sol", 0),
            "pos_dev_pct": lam.get("pos_dev_pct", 0),
            "pos_cert_pct": lam.get("pos_cert_pct", 0),
        },
        "ranking_dev_sol": ranking_dev_sol,
        "ranking_dev_pct": ranking_dev_pct,
        "ranking_cert_pct": ranking_cert_pct,
    }


def main():
    print("=" * 70)
    print("Actualización progresiva — Ranking nacional GORES (Ene-Jul y en adelante)")
    print("=" * 70)
    print(f"Períodos configurados: {[p[0] for p in PERIODOS]}")

    if not os.path.exists(ARCHIVO_JSON):
        print(f"❌ No se encontró {ARCHIVO_JSON}. Abortando.")
        return

    with open(ARCHIVO_JSON, "r", encoding="utf-8") as f:
        data = json.load(f)

    if "semestres" not in data:
        data["semestres"] = {}

    for año in AÑOS:
        resultado = procesar_año(año, data)
        if resultado is not None:
            data["semestres"][str(año)] = resultado

    data["generado"] = datetime.today().strftime("%Y-%m-%d")

    with open(ARCHIVO_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))

    print("\n" + "=" * 70)
    print(f"✅ Archivo actualizado: {ARCHIVO_JSON}")
    print("=" * 70)


if __name__ == "__main__":
    main()
