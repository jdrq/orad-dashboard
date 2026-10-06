#!/usr/bin/env python3
# =============================================================================
# actualizar_rb_hist_sc.py  (v3 — motor progresivo + desglose por rubro)
# ORAD - Gobierno Regional de Lambayeque
# Actualiza: data/rb_hist_sc_progresivo.json
#
# QUÉ CAMBIÓ EN v3 (06/10/2026)
#   Problema que resuelve: hasta v2 el script solo actualizaba los TOTALES
#   acumulados (dev_t1, dev_t2, dev_sem, dev_ago, dev_set, dev_oct). El
#   desglose por rubro (campo "rubros": cert / comp / dev de cada rubro)
#   quedó congelado en Ene-Jul. Resultado visible en el dashboard: las filas
#   de la tabla "Rubro — Sede Central" (años 2022-2025) sumaban Ene-Jul y la
#   fila TOTAL mostraba Ene-Oct. Además la columna "comp" era una copia de
#   "cert" (nunca fue el Compromiso real).
#
#   v3 hace tres cosas nuevas:
#     1) Acumula POR RUBRO (cert, comp, dev) a lo largo de todos los
#        períodos activos y reescribe "rubros" en el JSON. El PIM de cada
#        rubro NO viene en estos archivos mensuales (la columna sale vacía),
#        así que se CONSERVA el que ya está en el JSON.
#     2) Compromiso = columna "Compromiso Anual" del MEF (antes era una
#        copia de la certificación).
#     3) Valida cada archivo ANTES de sumarlo:
#          - que sea de "Ejecución del Gasto" (no de Ingresos / "Recaudado")
#          - que traiga la fila "Unidad Ejecutora 001-855" (Sede Central)
#          - que el "Año de Ejecución" del encabezado coincida con el año
#            del nombre del archivo
#        Estos tres errores ya ocurrieron (octubre_2022/23/24 eran de
#        Ingresos; AGOSTO_RUBRO venía a nivel Pliego sin la UE) y el script
#        v2 no los detectaba, porque solo bloqueaba si el acumulado bajaba.
#     También actualiza "label" y "av_pct" del año al último período activo
#     (antes quedaban en "Ene–Ago"). index.html no los usa para pintar —
#     los recalcula en vivo — pero así el JSON no queda con datos viejos.
#
# CÓMO AGREGAR UN MES NUEVO (ej. cuando cierre noviembre):
#   1) Re-exportar en Consulta Amigable: Mes 11, columna Rubro seleccionada,
#      UE 001-855 Sede Central, "Sólo Proyectos". Guardar como
#      NOVIEMBRE_RUBRO_{año}.xls en xls/historico_rubro/ (2022-2025).
#   2) Descomentar la línea "dev_nov" en PERIODOS (y verificar que TODOS los
#      meses anteriores sigan activos: si solo activás el último, el
#      acumulado queda incompleto).
#   3) Correr DESDE LA RAÍZ DEL REPO:  python scripts\actualizar_rb_hist_sc.py
#
# CAMPOS QUE GENERA EN EL JSON (acumulado progresivo, Ene -> fin de mes):
#   dev_t1  = acumulado a marzo      dev_t2  = acumulado a junio
#   dev_sem = acumulado a julio  (nombre heredado, ver nota)
#   dev_ago / dev_set / dev_oct ... = acumulado al mes respectivo
#   rubros  = desglose por rubro AL ÚLTIMO PERÍODO ACTIVO (cert, comp, dev)
#
# NOTA sobre "dev_sem": nombre heredado de cuando representaba el semestre
# (Ene-Jun). Hoy representa Ene-Jul. No se renombra para no romper
# index.html. Los meses nuevos usan nombres correctos (dev_ago, dev_set...).
#
# VALIDACIÓN DE COHERENCIA:
#   - Julio (dev_sem) se valida contra el benchmark que ya existía en el JSON.
#   - Del resto se valida MONOTONICIDAD: el acumulado no puede bajar.
#   - NUEVO: la suma de dev de las filas de "rubros" debe cuadrar con el
#     total acumulado del último período (tolerancia por redondeo).
#   Si algo falla, NO se escribe nada de ese año.
# =============================================================================

import os
import re
import json
from bs4 import BeautifulSoup

# ---------------- CONFIGURACIÓN ----------------
AÑOS = [2021, 2022, 2023, 2024, 2025]
CARPETA_XLS = "xls/historico_rubro"
CARPETA_DATA = "data"
ARCHIVO_JSON = os.path.join(CARPETA_DATA, "rb_hist_sc_progresivo.json")
CODIGO_UE = "001-855"          # Sede Central

# Lista ordenada de períodos a acumular progresivamente.
# Cada entrada: (clave_campo_json, patrón_de_archivo, tiene_benchmark)
PERIODOS = [
    ("dev_t1", "1T_RUBRO_{año}.xls", False),
    ("dev_t2", "2T_RUBRO_{año}.xls", False),
    ("dev_sem", "JULIO_RUBRO_{año}.xls", True),
    ("dev_ago", "AGOSTO_RUBRO_{año}.xls", False),
    ("dev_set", "SETIEMBRE_RUBRO_{año}.xls", False),
    ("dev_oct", "OCTUBRE_RUBRO_{año}.xls", False),
    # --- Descomentar la línea del mes correspondiente cuando cierre ---
    # ("dev_nov", "NOVIEMBRE_RUBRO_{año}.xls", False),
    # ("dev_dic", "DICIEMBRE_RUBRO_{año}.xls", False),
]

# Filtro de mes que DEBE traer cada archivo (número, nombre en el MEF).
# None = acumulado sin filtro de mes. Detecta, p. ej., un archivo de Setiembre
# exportado sin filtro de mes o agrupado por mes en vez de por Rubro.
MES_ESPERADO = {"dev_t1": None, "dev_t2": None,
                "dev_sem": (7, r"julio"), "dev_ago": (8, r"agosto"),
                "dev_set": (9, r"sep?tiembre"), "dev_oct": (10, r"octubre"),
                "dev_nov": (11, r"noviembre"), "dev_dic": (12, r"diciembre")}

# Subcarpeta donde vive cada tipo de archivo dentro de xls/historico_rubro/.
# (Si la subcarpeta no existe, se busca también en la carpeta principal.)
SUBCARPETA = {"dev_t1": "TRIMESTRE", "dev_t2": "TRIMESTRE", "ANUAL": "ANUAL"}   # el resto: "MES"

# Archivo con el PIM por rubro (sin filtro de mes). Solo se usa para años que
# aún no están en el JSON; se busca sin distinguir mayúsculas.
ARCHIVO_ANUAL = "ANUAL_RUBRO_{año}.xls"

# Nombres cortos de rubros conocidos (los que usa el dashboard).
NOMBRES_RUBRO = {"00": "Recursos Ordinarios", "09": "R. Dir. Recaudados",
                 "13": "Donaciones y Transf.", "15": "FONCOR",
                 "18": "Canon y Sobrecanon", "19": "Op. Oficiales de Crédito"}

# Abreviatura del mes de corte para el campo "label" ("Ene–Oct 2022")
MES_ABREV = {"dev_t1": "Mar", "dev_t2": "Jun", "dev_sem": "Jul",
             "dev_ago": "Ago", "dev_set": "Set", "dev_oct": "Oct",
             "dev_nov": "Nov", "dev_dic": "Dic"}
# -------------------------------------------------


def ubicar_archivo(nombre, clave):
    """
    Busca el archivo en xls/historico_rubro/<TRIMESTRE|MES|ANUAL>/ y, si no está
    ahí, en xls/historico_rubro/ (estructura antigua). Ignora mayúsculas.
    Devuelve la ruta encontrada o, si no existe, la ruta esperada (para el aviso).
    """
    sub = SUBCARPETA.get(clave, "MES")
    candidatas = [os.path.join(CARPETA_XLS, sub), CARPETA_XLS]
    for carpeta in candidatas:
        if os.path.isdir(carpeta):
            for f in os.listdir(carpeta):
                if f.casefold() == nombre.casefold() and os.path.isfile(os.path.join(carpeta, f)):
                    return os.path.join(carpeta, f)
    return os.path.join(candidatas[0], nombre)


def limpiar_numero(s):
    """Convierte string con comas y S/ a float."""
    s = str(s).replace(",", "").replace("S/", "").strip()
    try:
        return float(s) if s else 0.0
    except ValueError:
        return 0.0


def validar_archivo(path, contenido, soup, año, clave=None):
    """
    Chequeos previos a sumar un archivo. Devuelve True si es utilizable.
    Cada chequeo corresponde a un error que YA ocurrió en este proyecto.
    """
    nombre = os.path.basename(path)
    texto_crudo = contenido.decode("latin-1")
    texto = re.sub(r"\s+", " ", soup.get_text(" "))

    if "Recaudado" in texto_crudo:
        print(f"   ⚠️  {nombre} es de INGRESOS (columna 'Recaudado'), no de Gasto. "
              f"Re-exportar desde 'Consulta de Ejecución del Gasto'.")
        return False

    if not re.search(r"Ejecuci.n del Gasto", texto):
        print(f"   ⚠️  {nombre}: el encabezado no dice 'Ejecución del Gasto'. "
              f"¿Se exportó de otra consulta?")
        return False

    if f"Unidad Ejecutora {CODIGO_UE}" not in texto:
        print(f"   ⚠️  {nombre}: falta la fila 'Unidad Ejecutora {CODIGO_UE}'. "
              f"El archivo está a nivel Pliego completo (montos inflados). "
              f"Re-exportar filtrando la UE Sede Central.")
        return False

    m = re.search(r"A.o de Ejecuci.n:\s*(\d{4})", texto)
    if m and int(m.group(1)) != año:
        print(f"   ⚠️  {nombre}: el encabezado dice año {m.group(1)} pero se "
              f"esperaba {año}.")
        return False

    # Filtro de mes (v4): el archivo debe corresponder al período esperado.
    meses = re.findall(r"Mes (\d{1,2}): ?(\w+)", texto)
    esperado = MES_ESPERADO.get(clave)
    if clave is not None and esperado is None and meses:
        print(f"   ⚠️  {nombre}: trae filtro de mes ({meses[0][1]}) pero debe ser "
              f"un acumulado sin filtro de mes.")
        return False
    if esperado is not None:
        num, patron = esperado
        if not any(int(n) == num and re.fullmatch(patron, nom, re.I) for n, nom in meses):
            hallado = ", ".join(f"{n}:{nom}" for n, nom in meses) or "ninguno"
            print(f"   ⚠️  {nombre}: el filtro de mes no es el {num} (hallado: "
                  f"{hallado}). Re-exportar con el mes correcto y agrupado por Rubro.")
            return False

    return True


def parsear_rubros(path, año, clave=None):
    """
    Lee un archivo XLS del MEF (HTML disfrazado) filtrado por Rubro, a nivel
    UE 001-855 Sede Central. Devuelve un dict por rubro:
        { "18": {"cert": ..., "comp": ..., "dev": ...}, ... }
    del período del archivo, o None si no existe / no es válido.

    Columnas de la tabla de detalle (índice 3 del HTML):
      [0] "NN: nombre del rubro"  [1] PIA  [2] PIM (vacío en estos archivos)
      [3] Certificación  [4] Compromiso Anual  [5] Atención Comp. Mensual
      [6] Devengado  [7] Girado  [8] Avance %
    """
    if not os.path.exists(path):
        print(f"   ⚠️  ARCHIVO NO ENCONTRADO: {path}")
        return None

    with open(path, "rb") as f:
        contenido = f.read()

    soup = BeautifulSoup(contenido, "html.parser")
    tables = soup.find_all("table")

    if len(tables) < 4:
        print(f"   ⚠️  Formato inesperado: solo {len(tables)} tablas en {path}")
        return None

    if not validar_archivo(path, contenido, soup, año, clave):
        return None

    rubros = {}
    for r in tables[3].find_all("tr"):
        cols = [c.get_text(strip=True) for c in r.find_all(["td", "th"])]
        if len(cols) >= 8 and re.match(r"^\d{2}:", cols[0]):
            codigo = cols[0][:2]
            rubros[codigo] = {
                "cert": limpiar_numero(cols[3]),
                "comp": limpiar_numero(cols[4]),
                "dev": limpiar_numero(cols[6]),
            }

    if not rubros:
        print(f"   ⚠️  Sin filas de detalle por Rubro en {path} "
              f"(¿se exportó sin la columna 'Rubro' seleccionada?)")
        return None

    return rubros


def leer_anual(año):
    """
    Para un año NUEVO (sin entrada en el JSON): lee ANUAL_RUBRO_{año}.xls y
    devuelve (lista_de_rubros_con_PIM, pim_total) o None.
    Los rubros salen ordenados por PIM descendente (como en los años ya cargados).
    """
    nombre = ARCHIVO_ANUAL.format(año=año)
    path = ubicar_archivo(nombre, "ANUAL")
    if not os.path.isfile(path):
        print(f"   ❌ Año nuevo sin {nombre} en {os.path.dirname(path)} (se necesita para el PIM por rubro).")
        return None
    with open(path, "rb") as f:
        contenido = f.read()
    soup = BeautifulSoup(contenido, "html.parser")
    tables = soup.find_all("table")
    if len(tables) < 4 or not validar_archivo(path, contenido, soup, año, None):
        return None
    filas = []
    for r in tables[3].find_all("tr"):
        cols = [c.get_text(strip=True) for c in r.find_all(["td", "th"])]
        if len(cols) >= 8 and re.match(r"^\d{2}:", cols[0]):
            cod = cols[0][:2]
            nom = NOMBRES_RUBRO.get(cod) or cols[0][3:].strip().title()
            filas.append({"codigo": cod, "nombre": nom,
                          "pim": round(limpiar_numero(cols[2]))})
    if not filas:
        print(f"   ⚠️  {nombre}: sin filas por Rubro (¿exportado sin agrupar por Rubro?).")
        return None
    filas.sort(key=lambda x: -x["pim"])
    return filas, sum(x["pim"] for x in filas)


def reconstruir_rubros(rubros_json, acum):
    """
    Reescribe la lista "rubros" del JSON con los acumulados recién calculados.
    Conserva nombre, PIM y orden de los rubros que ya estaban en el JSON
    (el PIM por rubro no viene en los archivos mensuales).
    Si aparece un rubro nuevo en los archivos, lo agrega al final con PIM 0
    y avisa: hay que completarle nombre y PIM a mano.
    """
    resultado, vistos = [], set()
    for r in rubros_json:
        cod = r["codigo"]
        vistos.add(cod)
        a = acum.get(cod, {"cert": 0.0, "comp": 0.0, "dev": 0.0})
        nuevo = dict(r)
        nuevo["cert"] = round(a["cert"])
        nuevo["comp"] = round(a["comp"])
        nuevo["dev"] = round(a["dev"])
        resultado.append(nuevo)

    for cod, a in sorted(acum.items()):
        if cod in vistos:
            continue
        if abs(a["dev"]) < 1 and abs(a["cert"]) < 1 and abs(a["comp"]) < 1:
            continue
        print(f"   ⚠️  Rubro {cod} aparece en los archivos pero NO está en el "
              f"JSON. Se agrega con PIM 0: completar nombre y PIM a mano.")
        resultado.append({"codigo": cod,
                          "nombre": f"Rubro {cod} (completar nombre y PIM)",
                          "pim": 0,
                          "cert": round(a["cert"]),
                          "comp": round(a["comp"]),
                          "dev": round(a["dev"])})
    return resultado


def procesar_año(año, data_existente):
    print(f"\n--- Año {año} ---")
    año_str = str(año)
    entrada_actual = data_existente.get(año_str, {})
    entrada_propuesta = entrada_actual.copy()

    acumulado = 0.0
    acum_rubros = {}                 # codigo -> {"cert","comp","dev"} acumulados
    bloqueado = False
    valor_anterior_acumulado = None  # para el chequeo de monotonicidad
    ultima_clave = None

    for clave_campo, patron_archivo, tiene_benchmark in PERIODOS:
        path = ubicar_archivo(patron_archivo.format(año=año), clave_campo)
        print(f"   Leyendo {clave_campo} ({patron_archivo.format(año=año)})")
        rubros_periodo = parsear_rubros(path, año, clave_campo)

        if rubros_periodo is None:
            print(f"   ❌ No se pudo leer {clave_campo} para {año} — "
                  f"se conserva el valor anterior de TODOS los campos de este año.")
            return entrada_actual  # aborta todo el año, no solo el campo

        periodo = sum(r["dev"] for r in rubros_periodo.values())
        acumulado += periodo
        for cod, r in rubros_periodo.items():
            a = acum_rubros.setdefault(cod, {"cert": 0.0, "comp": 0.0, "dev": 0.0})
            for campo in a:
                a[campo] += r[campo]
        ultima_clave = clave_campo

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

        # --- Validación por monotonicidad ---
        if valor_anterior_acumulado is not None and acumulado < valor_anterior_acumulado - 1:
            print(f"   ⚠️  ALERTA: {clave_campo} (S/{acumulado:,.0f}) es MENOR al "
                  f"acumulado anterior (S/{valor_anterior_acumulado:,.0f}). El "
                  f"devengado acumulado no debería bajar. BLOQUEADO.")
            bloqueado = True

        entrada_propuesta[clave_campo] = round(acumulado)
        valor_anterior_acumulado = acumulado

    # --- Año nuevo: el PIM por rubro sale del archivo ANUAL_RUBRO ---
    base_rubros = entrada_actual.get("rubros", [])
    if not base_rubros:
        anual = leer_anual(año)
        if anual is None:
            print(f"   🚫 Año {año}: sin PIM por rubro, no se escribe.")
            return entrada_actual
        base_rubros, pim_total = anual
        entrada_propuesta["pim"] = pim_total
        print(f"   ℹ️  Año nuevo: PIM por rubro tomado de ANUAL_RUBRO (total S/{pim_total:,.0f}).")

    # --- Desglose por rubro al último período activo ---
    rubros_nuevos = reconstruir_rubros(base_rubros, acum_rubros)

    # Coherencia: las filas de rubros deben sumar el total acumulado
    suma_filas = sum(r["dev"] for r in rubros_nuevos)
    total = entrada_propuesta[ultima_clave]
    if abs(suma_filas - total) > len(rubros_nuevos) + 1:   # tolerancia por redondeo
        print(f"   ⚠️  Las filas de rubros suman S/{suma_filas:,.0f} pero el total "
              f"{ultima_clave} es S/{total:,.0f}. BLOQUEADO.")
        bloqueado = True
    else:
        print(f"   ✅ Filas de rubros cuadran con {ultima_clave} "
              f"(S/{suma_filas:,.0f} vs S/{total:,.0f})")

    if bloqueado:
        print(f"   🚫 Año {año}: NO se escriben cambios (algún check falló).")
        return entrada_actual

    entrada_propuesta["rubros"] = rubros_nuevos
    entrada_propuesta["label"] = f"Ene–{MES_ABREV.get(ultima_clave, ultima_clave)} {año}"
    pim = entrada_propuesta.get("pim")
    if pim:
        entrada_propuesta["av_pct"] = round(acumulado / pim * 100, 1)

    print(f"   ✅ Año {año} completo y coherente: {entrada_propuesta['label']} "
          f"= S/{acumulado:,.0f} ({entrada_propuesta.get('av_pct', '?')}% del PIM)")
    return entrada_propuesta


def main():
    print("=" * 70)
    print("Actualización progresiva — Rubro Sede Central")
    print("=" * 70)
    print(f"Períodos configurados: {[p[0] for p in PERIODOS]}")

    if not os.path.exists(ARCHIVO_JSON):
        print(f"❌ No se encontró {ARCHIVO_JSON}. ¿Corriste el script desde la "
              f"raíz del repo? Abortando.")
        return

    with open(ARCHIVO_JSON, "r", encoding="utf-8") as f:
        data = json.load(f)

    for año in AÑOS:
        resultado = procesar_año(año, data)
        if resultado is not None:
            data[str(año)] = resultado

    data = {k: data[k] for k in sorted(data)}   # años en orden ascendente
    with open(ARCHIVO_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))

    print("\n" + "=" * 70)
    print(f"✅ Archivo actualizado: {ARCHIVO_JSON}")
    print("=" * 70)


if __name__ == "__main__":
    main()
