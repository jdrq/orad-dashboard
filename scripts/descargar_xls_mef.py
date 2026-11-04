r"""
descargar_xls_mef.py
---------------------
Automatiza la descarga de los 17 archivos .xls de Consulta Amigable (MEF)
que alimentan el dashboard de ORAD - GORE Lambayeque.

MODO DE USO (Fase 2a - manual, supervisado):
    python scripts\descargar_xls_mef.py

Requiere:
    pip install playwright
    playwright install chromium

Arquitecto: este script corre en modo "headed" (con ventana visible) a
propósito durante esta fase - así Juan puede observar cada descarga y
detectar fallos en el momento, en vez de un cron job a ciegas.

GIT AUTO-PUSH (agregado en sesión 7 - 24/07/2026):
    Al terminar la descarga, el script verifica automáticamente que los
    17 archivos existan y tengan peso > 0. Si pasan todos:
        git add xls\ → git commit → git push
    Si falla cualquiera, NO se toca Git en absoluto. El usuario ve
    exactamente qué archivo falló y debe resolverlo antes de subir.

BLOQUE 2D - DEVENGADO MENSUAL (agregado en sesión 10 - 14/08/2026):
    Se agregó gore_devengado_mes.xls como 14º archivo, para alimentar
    data/devengado-mensual.js (Bloque 2D del dashboard). A diferencia de
    los demás, este NO pivota por Rubro ni Función, sino por "Mes" - un
    botón normal por rol/nombre visible (get_by_role("button", name="Mes")),
    SIN ID fijo como BtnRubro/BtnFuncion. Confirmado con playwright codegen
    real el 14/08/2026 (ver campo pivot_final_rol en ARCHIVOS).

BLOQUE 6D - DEVENGADO MENSUAL, SEDE CENTRAL (agregado en sesión 10 - 14/08/2026):
    Se agregó sede_devengado_mes.xls como 15º archivo, para alimentar
    data/devengado-mensual-sc.js (Bloque 6D del dashboard). Mismo patrón
    que el Bloque 2D (pivote por "Mes" vía pivot_final_rol), pero un nivel
    más abajo en la jerarquía: Ejecutora 001-855 (Sede Central). Confirmado
    con playwright codegen real el 14/08/2026.

BLOQUE EMR - CATEGORÍA PRESUPUESTAL 0068, EMERGENCIAS/DESASTRES (agregado en
sesión 12 - 18/08/2026):
    Se agregó proyecto_emergencia.xls como 16º archivo, para alimentar el
    Bloque EMR del dashboard (Reducción de Vulnerabilidad y Atención de
    Emergencias). A diferencia de todos los demás, usa el filtro "Actividades
    y Proyectos" (ap=ActProy), NO "Sólo Proyectos" (ap=Proyecto) — por eso
    URL_BASE dejó de ser una sola constante global y pasó a resolverse por
    archivo (campo "url" en ARCHIVOS, con fallback al valor de siempre si no
    se especifica). Además, el drill llega por una rama nueva: Pliego 452 ->
    "Categoría Presupuestal" (botón de cadena, mismo patrón que Nivel de
    Gobierno/Sector/Pliego) -> fila "0068:" -> pivote final "Genérica" por
    DOBLE clic (a diferencia de "Mes", que es clic simple). Confirmado con
    playwright codegen real el 18/08/2026.

REINTENTOS CON DOBLE CLIC (agregado en sesión 13 - 18/08/2026):
    Antes, solo el clic en la FILA tenía reintentos (3x, siempre clic
    simple) - el botón de cadena, el clic inicial en "TOTAL" y "Exportar"
    no tenían ninguno. En la práctica, Juan tenía que ayudar al script
    haciendo clic manualmente cuando se quedaba esperando una selección
    que nunca se marcaba. Causa: el UpdatePanel de Consulta Amigable a
    veces "traga" un clic simple si llega justo cuando termina el postback
    anterior. Ahora TODOS los puntos de clic (TOTAL, botón de cadena, fila,
    Exportar) reintentan, y desde el 2º intento escalan a DOBLE clic -
    automatizando exactamente la ayuda manual que antes hacía Juan.

FIXES DE ROBUSTEZ (sesión 14 - 19/08/2026, tras un fallo real en producción):
    1) clic_con_reintento ahora hace POLLING (hasta 12s, revisando cada 1s)
       en vez de un solo chequeo temprano. Causa del fallo real: en un día
       con el MEF más lento, el chequeo único llegaba demasiado pronto,
       se interpretaba como "el clic no sirvió" y disparaba un 2º clic
       MIENTRAS el primero seguía procesándose - ese 2º clic sí se
       quedaba colgado 30s esperando que el botón volviera a ser
       accionable. Además, cada locator.click() ahora tiene su propio
       timeout corto (10s) envuelto en try/except, para no depender del
       timeout por defecto de Playwright (30s) sin control.
    2) exito_total ahora SÍ se usa: si algún archivo falla, el script NO
       llega a Git - antes igual llamaba a git_push_si_completo(), y como
       la validación de archivos solo revisa existencia+peso (no fecha),
       un archivo viejo de ayer pasaba la validación como si nada, con
       riesgo real de que el dashboard quedara desactualizado en
       silencio.
    3) La detección de "nada que commitear" ahora también reconoce el
       texto real de Git "no changes added to commit" (antes solo
       buscaba "nothing added to commit", que Git no usa en este caso, y
       por eso reportaba error falso).

BLOQUE EMU - CATEGORÍA PRESUPUESTAL 0068, MUNICIPALIDADES (agregado en
sesión 13 - 18/08/2026):
    Se agregó munis_emergencia.xls como 17º archivo, para alimentar el
    bloque "Reducción de Vulnerabilidad y Atención de Emergencias —
    Municipalidades". Mismo filtro "Actividades y Proyectos" que el archivo
    anterior (url=URL_ACT_PROY), pero cadena de drill distinta: Nivel de
    Gobierno "M: GOBIERNOS LOCALES" -> "Gob.Loc./Mancom." -> "M:
    MUNICIPALIDADES" -> botón "Departamento" (sin nombre accesible estable,
    se ubica por ID fijo, ver soporte nuevo de selector_id en
    clic_con_reintento) -> "LAMBAYEQUE" -> "Categoría Presupuestal" ->
    "0068:" -> pivote final "Municipalidad" por clic simple. Confirmado con
    playwright codegen real el 18/08/2026.

NOTA DE DISEÑO IMPORTANTE (descubierto con playwright codegen el 06/07/2026):
Consulta Amigable usa UN SOLO botón para toda la jerarquía "Nivel de
Gobierno -> Sector -> Pliego -> Ejecutora -> Proyecto". Ese botón se
RE-ETIQUETA solo según la profundidad en la que estás parado (primero
dice "Nivel de Gobierno", luego "Sector", luego "Pliego"...), pero es
el mismo elemento con el mismo ID (#ctl00_CPH1_BtnTipoGobierno) todo
el tiempo. Por eso el "drill" real es: (clic en ese botón + clic en la
fila deseada), repetido tantas veces como niveles se quiera bajar.
Solo "Rubro", "Función", "Fuente", etc. son botones DISTINTOS que
pivotan el eje de la tabla en el nivel donde estés parado.
"""

from playwright.sync_api import sync_playwright
import hashlib
import os
import re
import sys
import time
import shutil
import smtplib
import subprocess
import urllib.request
from datetime import date, datetime
from email.message import EmailMessage
from pathlib import Path

# --------------------------------------------------------------------
# CONFIGURACIÓN
# --------------------------------------------------------------------
CARPETA_DESTINO = Path(__file__).resolve().parent.parent / "xls"
REPO_RAIZ = Path(__file__).resolve().parent.parent
CARPETA_LOGS = REPO_RAIZ / "logs"          # *.log ya está en .gitignore; ALERTA_*.txt también (ver README)
PAGES_URL = "https://jdrq.github.io/orad-dashboard"
# Archivos que se comparan contra la página publicada para saber si GitHub Pages ya desplegó.
ARCHIVOS_CENTINELA = ["rubro_pliego.xls", "ue_pliego.xls", "nacional_gores.xls"]
ESPERA_PAGES_SEG = 240      # cuánto esperar a que Pages publique antes de forzar el despliegue
INTERVALO_PAGES_SEG = 30
ANIO = "2026"
URL_BASE = f"https://apps5.mineco.gob.pe/transparencia/Navegador/default.aspx?y={ANIO}&ap=Proyecto"
# Filtro "Actividades y Proyectos" (Bloque EMR — Categoría 0068). Distinto del
# URL_BASE de siempre, que fija "Sólo Proyectos" (ap=Proyecto). Confirmado con
# playwright codegen real el 18/08/2026.
URL_ACT_PROY = f"https://apps5.mineco.gob.pe/transparencia/Navegador/default.aspx?y={ANIO}&ap=ActProy"
FRAME_SELECTOR = "#frame0"

BTN_RUBRO = "#ctl00_CPH1_BtnRubro"
BTN_FUNCION = "#ctl00_CPH1_BtnFuncion"

# Cada entrada describe:
#   pasos: lista ordenada de (etiqueta_boton, texto_fila) - en cada paso
#          se busca el botón de la cadena POR SU ETIQUETA VISIBLE en ese
#          momento (Nivel de Gobierno -> Sector -> Pliego -> Ejecutora),
#          NO por ID fijo, porque el ID cambia en cada nivel
#          (BtnTipoGobierno, BtnSector, BtnPliego, BtnEjecutora...).
#          Esto es exactamente lo que capturó el codegen real de Juan.
#   pivot_final: (selector, None) - botón de eje distinto a clickear al
#          final (Rubro, Función), identificado por ID fijo (#ctl00_...).
#   pivot_final_rol: (nombre_de_rol, None) - variante de pivot_final para
#          botones que NO tienen ID fijo confirmado y se ubican por rol/
#          nombre visible, igual que los pasos de la jerarquía (ej. "Mes",
#          confirmado con playwright codegen el 14/08/2026 — a diferencia
#          de Rubro/Función, no es un botón de ID fijo #ctl00_CPH1_Btn...).
#   pivot_dblclick: True/False - si el pivote de pivot_final_rol necesita
#          DOBLE clic en vez de clic simple (ej. "Genérica" en el archivo de
#          Categoría 0068, confirmado con playwright codegen el 18/08/2026).
#          False/ausente = clic simple, como "Mes".
#   url: URL base propia de este archivo (con su propio filtro ap=...). Si
#          no se especifica, se usa URL_BASE (ap=Proyecto, "Sólo Proyectos").
#   boton_final_sin_fila: etiqueta de un último clic al botón de cadena
#          SIN fila después (revela el desglose natural de ese nivel:
#          por Ejecutora, por Pliego, por Proyecto). None si no aplica.
ARCHIVOS = [
    {"nombre": "rubro_sede_central.xls",
     "pasos": [("Nivel de Gobierno", "R: GOBIERNOS REGIONALES"),
               ("Sector", "99: GOBIERNOS REGIONALES"),
               ("Pliego", "452:"),
               ("Ejecutora", "001-855:")],
     "pivot_final": BTN_RUBRO, "boton_final_sin_fila": None},
    {"nombre": "rubro_peot.xls",
     "pasos": [("Nivel de Gobierno", "R: GOBIERNOS REGIONALES"),
               ("Sector", "99: GOBIERNOS REGIONALES"),
               ("Pliego", "452:"),
               ("Ejecutora", "002-1133:")],
     "pivot_final": BTN_RUBRO, "boton_final_sin_fila": None},
    {"nombre": "rubro_agricultura.xls",
     "pasos": [("Nivel de Gobierno", "R: GOBIERNOS REGIONALES"),
               ("Sector", "99: GOBIERNOS REGIONALES"),
               ("Pliego", "452:"),
               ("Ejecutora", "100-856:")],
     "pivot_final": BTN_RUBRO, "boton_final_sin_fila": None},
    {"nombre": "rubro_transportes.xls",
     "pasos": [("Nivel de Gobierno", "R: GOBIERNOS REGIONALES"),
               ("Sector", "99: GOBIERNOS REGIONALES"),
               ("Pliego", "452:"),
               ("Ejecutora", "200-857:")],
     "pivot_final": BTN_RUBRO, "boton_final_sin_fila": None},
    {"nombre": "rubro_salud.xls",
     "pasos": [("Nivel de Gobierno", "R: GOBIERNOS REGIONALES"),
               ("Sector", "99: GOBIERNOS REGIONALES"),
               ("Pliego", "452:"),
               ("Ejecutora", "400-860:")],
     "pivot_final": BTN_RUBRO, "boton_final_sin_fila": None},
    {"nombre": "rubro_h_mercedes.xls",
     "pasos": [("Nivel de Gobierno", "R: GOBIERNOS REGIONALES"),
               ("Sector", "99: GOBIERNOS REGIONALES"),
               ("Pliego", "452:"),
               ("Ejecutora", "401-1001:")],
     "pivot_final": BTN_RUBRO, "boton_final_sin_fila": None},
    {"nombre": "rubro_h_belen.xls",
     "pasos": [("Nivel de Gobierno", "R: GOBIERNOS REGIONALES"),
               ("Sector", "99: GOBIERNOS REGIONALES"),
               ("Pliego", "452:"),
               ("Ejecutora", "402-1002:")],
     "pivot_final": BTN_RUBRO, "boton_final_sin_fila": None},
    {"nombre": "rubro_h_regional.xls",
     "pasos": [("Nivel de Gobierno", "R: GOBIERNOS REGIONALES"),
               ("Sector", "99: GOBIERNOS REGIONALES"),
               ("Pliego", "452:"),
               ("Ejecutora", "403-1422:")],
     "pivot_final": BTN_RUBRO, "boton_final_sin_fila": None},
    {"nombre": "rubro_pliego.xls",
     "pasos": [("Nivel de Gobierno", "R: GOBIERNOS REGIONALES"),
               ("Sector", "99: GOBIERNOS REGIONALES"),
               ("Pliego", "452:")],
     "pivot_final": BTN_RUBRO, "boton_final_sin_fila": None},
    {"nombre": "ue_pliego.xls",
     "pasos": [("Nivel de Gobierno", "R: GOBIERNOS REGIONALES"),
               ("Sector", "99: GOBIERNOS REGIONALES"),
               ("Pliego", "452:")],
     "pivot_final": None, "boton_final_sin_fila": "Ejecutora"},
    {"nombre": "nacional_gores.xls",
     "pasos": [("Nivel de Gobierno", "R: GOBIERNOS REGIONALES"),
               ("Sector", "99: GOBIERNOS REGIONALES")],
     "pivot_final": None, "boton_final_sin_fila": "Pliego"},
    {"nombre": "funciones_pliego.xls",
     "pasos": [("Nivel de Gobierno", "R: GOBIERNOS REGIONALES"),
               ("Sector", "99: GOBIERNOS REGIONALES"),
               ("Pliego", "452:")],
     "pivot_final": BTN_FUNCION, "boton_final_sin_fila": None},
    {"nombre": "proyectos_sede_central.xls",
     "pasos": [("Nivel de Gobierno", "R: GOBIERNOS REGIONALES"),
               ("Sector", "99: GOBIERNOS REGIONALES"),
               ("Pliego", "452:"),
               ("Ejecutora", "001-855:")],
     "pivot_final": None, "boton_final_sin_fila": "Producto/Proyecto"},
    # Bloque 2D — Devengado Mensual. Alimenta data/devengado-mensual.js.
    # Pivota por "Mes" en vez de Rubro/Función. Confirmado con playwright
    # codegen real el 14/08/2026: NO es un botón de ID fijo como BTN_RUBRO/
    # BTN_FUNCION, sino un botón normal por rol/nombre visible ("Mes"),
    # igual que los pasos de la jerarquía — por eso usa pivot_final_rol
    # en vez de pivot_final.
    {"nombre": "gore_devengado_mes.xls",
     "pasos": [("Nivel de Gobierno", "R: GOBIERNOS REGIONALES"),
               ("Sector", "99: GOBIERNOS REGIONALES"),
               ("Pliego", "452:")],
     "pivot_final": None, "pivot_final_rol": "Mes", "boton_final_sin_fila": None},
    # Bloque 6D — Devengado Mensual, Sede Central. Alimenta
    # data/devengado-mensual-sc.js. Mismo patrón que el anterior (pivote por
    # "Mes"), pero un nivel más abajo: Ejecutora 001-855 (Sede Central).
    # Confirmado con playwright codegen real el 14/08/2026.
    {"nombre": "sede_devengado_mes.xls",
     "pasos": [("Nivel de Gobierno", "R: GOBIERNOS REGIONALES"),
               ("Sector", "99: GOBIERNOS REGIONALES"),
               ("Pliego", "452:"),
               ("Ejecutora", "001-855:")],
     "pivot_final": None, "pivot_final_rol": "Mes", "boton_final_sin_fila": None},
    # Bloque EMR — Categoría Presupuestal 0068 (Emergencias/Desastres).
    # Alimenta el bloque "Reducción de Vulnerabilidad y Atención de
    # Emergencias" del dashboard. Filtro "Actividades y Proyectos"
    # (url=URL_ACT_PROY, NO ap=Proyecto como los demás 15). El drill llega
    # por una rama nueva: Pliego -> "Categoría Presupuestal" (botón de
    # cadena, mismo patrón que Nivel de Gobierno/Sector/Pliego) -> fila
    # "0068:" -> pivote final "Genérica" por DOBLE clic. Confirmado con
    # playwright codegen real el 18/08/2026.
    {"nombre": "proyecto_emergencia.xls",
     "url": URL_ACT_PROY,
     "pasos": [("Nivel de Gobierno", "R: GOBIERNOS REGIONALES"),
               ("Sector", "99: GOBIERNOS REGIONALES"),
               ("Pliego", "452:"),
               ("Categoría Presupuestal", "0068:")],
     "pivot_final": None, "pivot_final_rol": "Genérica", "pivot_dblclick": True,
     "boton_final_sin_fila": None},
    # Bloque EMU — Categoría 0068, Municipalidades (Gobiernos Locales).
    # Alimenta el bloque "Reducción de Vulnerabilidad y Atención de
    # Emergencias — Municipalidades" del dashboard. Igual que el archivo
    # anterior, filtro "Actividades y Proyectos" (url=URL_ACT_PROY). Cadena
    # distinta: Nivel de Gobierno "M: GOBIERNOS LOCALES" -> "Gob.Loc./
    # Mancom." -> "M: MUNICIPALIDADES" -> botón "Departamento" (SIN nombre
    # accesible estable, se ubica por ID fijo #ctl00_CPH1_BtnDepartamento)
    # -> "LAMBAYEQUE" -> "Categoría Presupuestal" -> "0068:" -> pivote
    # final "Municipalidad" por CLIC SIMPLE (a diferencia de "Genérica" en
    # el archivo anterior, que necesita doble clic). Confirmado con
    # playwright codegen real el 18/08/2026.
    {"nombre": "munis_emergencia.xls",
     "url": URL_ACT_PROY,
     "pasos": [("Nivel de Gobierno", "M: GOBIERNOS LOCALES"),
               ("Gob.Loc./Mancom.", "M: MUNICIPALIDADES"),
               (None, ": LAMBAYEQUE", "#ctl00_CPH1_BtnDepartamento"),
               ("Categoría Presupuestal", "0068:")],
     "pivot_final": None, "pivot_final_rol": "Municipalidad", "pivot_dblclick": False,
     "boton_final_sin_fila": None},
]

# Lista de los 17 nombres esperados — se usa en la validación final.
# Debe ser idéntica a los "nombre" de ARCHIVOS arriba.
NOMBRES_ESPERADOS = [a["nombre"] for a in ARCHIVOS]


# --------------------------------------------------------------------
# DESCARGA
# --------------------------------------------------------------------

def preparar_carpeta():
    CARPETA_DESTINO.mkdir(exist_ok=True)


def clic_con_reintento(fl, page, rol=None, nombre=None, selector_id=None,
                        intentos=3, verificar=None, exacto=True,
                        espera_verificacion=12):
    """
    Hace clic en un elemento con reintentos automáticos, ESCALANDO A DOBLE
    CLIC desde el 2º intento en adelante.

    Se puede ubicar el elemento por rol+nombre accesible (rol, nombre) o,
    si el botón no expone un nombre accesible estable (confirmado con
    playwright codegen el 18/08/2026 para "Departamento" en la cadena de
    Municipalidades), por selector_id (ej. "#ctl00_CPH1_BtnDepartamento").

    Agregado en sesión 13 (18/08/2026) para reemplazar la intervención
    manual que Juan tenía que hacer cuando el script se quedaba "congelado"
    esperando que una selección se marcara. Causa confirmada en la práctica:
    el UpdatePanel de Consulta Amigable a veces "traga" un clic simple si
    llega justo cuando termina el postback anterior; el doble clic casi
    inmediato (igual que hace Juan a mano) lo desatasca.

    verificar: función sin argumentos que retorna True si el clic surtió
    efecto (ej. una fila que debería quedar visible). Si no se pasa, solo
    se espera el postback y se asume éxito tras el primer clic.

    espera_verificacion: AJUSTADO en sesión 14 (19/08/2026) tras un fallo
    real en producción. Antes solo se revisaba UNA vez, ~5.8s después del
    clic. En un día con el servidor del MEF más lento de lo normal, esa
    única revisión llegó demasiado temprano, se interpretó como "el clic
    no sirvió", y disparó un 2º clic MIENTRAS el primero todavía se estaba
    procesando - ese 2º clic sí se quedó colgado 30s esperando a que el
    botón volviera a ser accionable (probablemente deshabilitado durante
    el postback). Ahora se hace polling: hasta `espera_verificacion`
    segundos revisando cada 1s, antes de considerar que el clic falló.
    """
    if selector_id:
        locator = fl.locator(selector_id)
    else:
        locator = fl.get_by_role(rol, name=nombre, exact=exacto)
    etiqueta_log = nombre or selector_id
    for intento in range(intentos):
        try:
            if intento == 0:
                locator.click(timeout=10000)
            else:
                # Escalar a doble clic a partir del 2º intento.
                locator.click(timeout=10000)
                time.sleep(0.3)
                locator.click(timeout=10000)
        except Exception as e:
            # El botón no estaba accionable (ej. deshabilitado durante un
            # postback en curso) - no abortar todo el archivo por esto,
            # solo pasar al siguiente intento tras una pausa.
            print(f"  [AVISO] '{etiqueta_log}' no fue accionable a tiempo "
                  f"({e.__class__.__name__}), reintentando...")
            time.sleep(2)
            continue

        try:
            page.wait_for_load_state("networkidle", timeout=5000)
        except Exception:
            pass

        if verificar is None:
            time.sleep(0.8)
            return

        # Polling: revisa cada 1s en vez de una sola vez, para no
        # confundir "el servidor está tardando" con "el clic no sirvió".
        for _ in range(espera_verificacion):
            try:
                if verificar():
                    return
            except Exception:
                pass  # lectura transitoria fallida durante un postback
            time.sleep(1)

        print(f"  [REINTENTO clic {intento + 1}/{intentos}] '{etiqueta_log}' "
              f"no surtió efecto tras {espera_verificacion}s de espera, "
              f"reintentando...")

    raise RuntimeError(
        f"El clic en '{etiqueta_log}' no surtió efecto tras {intentos} intentos "
        f"(ni con doble clic, ni con {espera_verificacion}s de espera cada uno). "
        f"Revisar manualmente."
    )


def validar_contenido(ruta, config):
    """
    Revisa que el archivo descargado SEA el que se pidió (v3, 04/11/2026).

    Antes solo se comprobaba que existiera y pesara >5 KB. Así, el 04/11/2026,
    rubro_pliego.xls y ue_pliego.xls llegaron con el contenido de nacional_gores
    (26 GOREs, sin Pliego 452): pesaban 49 KB, pasaron la validación y se
    subieron a GitHub. Ahora se lee el encabezado de la consulta y se compara
    con los pasos del drill-down. Devuelve (True, "") o (False, motivo).
    """
    import html as _html
    import re
    crudo = Path(ruta).read_bytes().decode("latin-1", errors="replace")
    texto = re.sub(r"\s+", " ", _html.unescape(re.sub(r"<[^>]+>", " ", crudo)))

    if "Recaudado" in crudo or not re.search(r"Ejecuci.n del Gasto", texto):
        return False, "no es una consulta de Ejecución del GASTO (¿Ingresos?)"

    url = config.get("url", URL_BASE)
    m = re.search(r"[?&]y=(\d{4})", url)
    if m:
        a = re.search(r"A.o de Ejecuci.n:\s*(\d{4})", texto)
        if not a or a.group(1) != m.group(1):
            return False, f"año del archivo = {a.group(1) if a else '?'}, se esperaba {m.group(1)}"

    if "ap=ActProy" in url:
        if not re.search(r"Incluye:\s*Actividades y Proyectos", texto):
            return False, 'falta el filtro "Actividades y Proyectos"'
    elif not re.search(r"Incluye:\s*S.lo Proyectos", texto):
        return False, 'falta el filtro "Sólo Proyectos"'

    hay_pliego = hay_ue = False
    for paso in config["pasos"]:
        fila = paso[1]
        if fila.startswith("R:"):
            req = r"Nivel de Gobierno R:"
        elif fila.startswith("M: GOBIERNOS"):
            req = r"Nivel de Gobierno M:"
        elif re.fullmatch(r"99:", fila) or fila.startswith("99:"):
            req = r"Sector 99:"
        elif re.fullmatch(r"\d{3}:", fila):
            req, hay_pliego = r"Pliego " + fila[:3] + ":", True
        elif re.fullmatch(r"\d{3}-\d+:", fila):
            req, hay_ue = r"Unidad Ejecutora " + fila.rstrip(":"), True
        elif re.fullmatch(r"\d{4}:", fila):
            req = r"Categor.a Presupuestal " + fila[:4]
        else:
            continue          # pasos sin rótulo fijo (período, municipalidad, etc.)
        if not re.search(req, texto):
            return False, (f"el encabezado no dice '{req}': el archivo está en un nivel "
                           f"distinto al pedido (no se llegó al nivel '{fila}')")

    # Un nivel de MÁS (consulta más amplia o más profunda de lo pedido)
    if not hay_pliego and re.search(r"Pliego \d{3}:", texto):
        return False, "trae un Pliego seleccionado pero este archivo es a nivel Sector"
    if hay_pliego and not hay_ue and re.search(r"Unidad Ejecutora \d", texto):
        return False, "trae una Unidad Ejecutora seleccionada pero este archivo es a nivel Pliego"

    if len(re.findall(r"<tr", crudo)) < 8:
        return False, "la tabla de detalle está vacía o incompleta"
    return True, ""


def procesar_archivo(page, config):
    """
    Ejecuta la secuencia completa de un archivo: navegar, filtrar,
    drillear la jerarquía Gobierno, pivotar (si aplica) y exportar.

    Usa frame_locator (no page.frame(name=...)) porque se auto-resuelve
    solo ante cada recarga del frame - esto es lo que finalmente
    resolvió todos los problemas de "frame detached" de las pruebas
    anteriores, confirmado con playwright codegen real.
    """
    page.goto(config.get("url", URL_BASE))
    fl = page.frame_locator(FRAME_SELECTOR)

    clic_con_reintento(fl, page, "cell", "TOTAL", intentos=2)
    # Nota: NO seleccionamos Actividades/Proyectos por dropdown porque
    # la URL ya lo fija (?ap=Proyecto o ?ap=ActProy, según el archivo) -
    # hacerlo de nuevo disparaba
    # una recarga completa de página que descarrilaba el resto del flujo.
    try:
        page.wait_for_load_state("networkidle", timeout=5000)
    except Exception:
        pass

    # v2 (02/11/2026): gancho opcional para filtrar por período (lo usa
    # scripts/descargar_historicos.py). Sin "hook_periodo" no hace nada, así
    # que la descarga diaria de los 17 archivos se comporta exactamente igual.
    if config.get("hook_periodo") and config.get("periodo_momento", "inicio") == "inicio":
        config["hook_periodo"](page, fl)

    # Bajar por la jerarquía (Gobierno -> Sector/Gob.Loc./Departamento -> ...).
    # El botón se busca POR SU ETIQUETA VISIBLE en cada paso (no por ID
    # fijo) porque el ID real cambia en cada nivel: BtnTipoGobierno,
    # BtnSector, BtnPliego, etc. - confirmado en vivo el 06/07/2026.
    # EXCEPCIÓN: si el paso trae un 3er elemento (selector_id), ese botón
    # no expone nombre accesible estable y se ubica por ID fijo en su lugar
    # (ej. "Departamento" en la cadena de Municipalidades, confirmado con
    # playwright codegen el 18/08/2026).
    for paso in config["pasos"]:
        etiqueta_boton, texto_fila = paso[0], paso[1]
        selector_id_boton = paso[2] if len(paso) > 2 else None

        clic_con_reintento(
            fl, page, rol="button", nombre=etiqueta_boton,
            selector_id=selector_id_boton, intentos=2,
            verificar=lambda tf=texto_fila: fl.get_by_role("cell", name=tf, exact=(tf in config.get("filas_exactas", ()))).first.is_visible()
        )
        time.sleep(1)  # margen extra: el servidor del MEF a veces es lento

        # Clic en la fila con hasta 3 intentos: si el checkpoint no
        # confirma la selección (fallo transitorio del servidor), se
        # vuelve a clickear en vez de rendirse al primer intento.
        # Desde el 2º intento se escala a DOBLE clic (ver clic_con_reintento
        # más arriba para el porqué) - esto es justo lo que antes requería
        # ayuda manual de Juan.
        confirmado = False
        for intento in range(3):
            _celda = fl.get_by_role("cell", name=texto_fila,
                                    exact=(texto_fila in config.get("filas_exactas", ())))
            if intento == 0:
                _celda.click()
            else:
                _celda.click()
                time.sleep(0.3)
                _celda.click()
            try:
                page.wait_for_load_state("networkidle", timeout=5000)
            except Exception:
                pass
            time.sleep(1)

            # CHECKPOINT: confirmar que el radio de la fila quedó
            # realmente marcado. El breadcrumb ("History") NO es buena
            # señal acá - ese resumen solo aparece una vez que se avanza
            # al SIGUIENTE nivel, no apenas se selecciona la fila actual
            # (confirmado con captura real el 06/07/2026). El radio, en
            # cambio, se marca de inmediato al clickear.
            for _ in range(15):
                try:
                    fila = fl.locator("tr", has_text=texto_fila).first   # sin comillas CSS: algunas filas llevan apóstrofo (": 'Julio")
                    if fila.locator("input:checked").count() > 0:
                        confirmado = True
                        break
                except Exception:
                    pass  # lectura transitoria fallida durante un postback
                time.sleep(1)

            if confirmado:
                break
            print(f"  [REINTENTO {intento + 1}/3] '{texto_fila}' no se "
                  f"marcó, volviendo a clickear (doble clic)...")

        if not confirmado:
            raise RuntimeError(
                f"El nivel '{texto_fila}' no quedó marcado (radio) tras "
                f"3 intentos. El drill pudo haberse saltado este nivel - "
                f"revisar manualmente."
            )

    # Paso final: pivotar a otro eje (Rubro/Función) o revelar el
    # desglose natural del último nivel (un clic más a la cadena)
    if config["pivot_final"] is not None:
        fl.locator(config["pivot_final"]).click()
    elif config.get("pivot_final_rol") is not None:
        boton = fl.get_by_role("button", name=config["pivot_final_rol"], exact=True)
        if config.get("pivot_dblclick"):
            boton.dblclick()
        else:
            boton.click()
    elif config["boton_final_sin_fila"] is not None:
        fl.get_by_role("button", name=config["boton_final_sin_fila"], exact=True).click()

    try:
        page.wait_for_load_state("networkidle", timeout=5000)
    except Exception:
        pass
    time.sleep(1.5)

    # CHECKPOINT POST-PIVOTE: confirmar que el último nivel drilleado
    # (ej. la UE 001-855) SIGUE presente en el breadcrumb después del
    # clic final. Algunos botones de pivote (ej. "Producto/Proyecto")
    # podrían resetear la selección de vuelta a un nivel más alto -
    # si eso pasa, el archivo exportado tendría datos de MÁS de lo
    # pedido (ej. todo el Pliego en vez de solo una UE), silenciosamente.
    # Con polling (no un solo chequeo) para no repetir el mismo error
    # de impaciencia del primer checkpoint.
    if config["pasos"]:
        ultimo_texto = config.get("texto_breadcrumb", config["pasos"][-1][1]).rstrip(":")
        confirmado_final = False
        for _ in range(15):
            try:
                breadcrumb_final = fl.locator(".History").inner_text(timeout=3000)
                if ultimo_texto in breadcrumb_final:
                    confirmado_final = True
                    break
            except Exception:
                pass
            time.sleep(1)
        if not confirmado_final:
            raise RuntimeError(
                f"Después del clic final, el nivel '{ultimo_texto}' ya "
                f"no aparece en el breadcrumb tras 15s - el pivote final "
                f"pudo haber reseteado la selección a un nivel más "
                f"amplio. Archivo NO exportado, revisar manualmente."
            )

    if config.get("hook_periodo") and config.get("periodo_momento") == "antes_exportar":
        config["hook_periodo"](page, fl)

    # Clic en "Exportar" con reintento: si el primer clic no dispara la
    # descarga dentro de 15s, se reintenta con doble clic antes de rendirse
    # (mismo patrón de escalada que el resto del script).
    descarga = None
    for intento in range(2):
        try:
            with page.expect_download(timeout=15000) as descarga_info:
                if intento == 0:
                    fl.get_by_role("link", name="Exportar").click()
                else:
                    print("  [REINTENTO Exportar] no se disparó la "
                          "descarga, reintentando con doble clic...")
                    fl.get_by_role("link", name="Exportar").click()
                    time.sleep(0.3)
                    fl.get_by_role("link", name="Exportar").click()
            descarga = descarga_info.value
            break
        except Exception:
            if intento == 1:
                raise RuntimeError(
                    "El clic en 'Exportar' no disparó la descarga tras "
                    "2 intentos. Revisar manualmente."
                )

    # v2: "destino" permite guardar en otra ruta (descargas históricas, que
    # se validan antes de moverse a su carpeta definitiva y no usan respaldo).
    destino = Path(config["destino"]) if config.get("destino") else CARPETA_DESTINO / config["nombre"]
    destino.parent.mkdir(parents=True, exist_ok=True)

    # v3: se descarga a un .tmp, se VALIDA el contenido y solo entonces reemplaza al
    # archivo bueno. Un archivo equivocado nunca pisa al anterior.
    tmp = destino.with_name(destino.name + ".tmp")
    descarga.save_as(tmp)
    ok, motivo = validar_contenido(tmp, config)
    if not ok:
        tmp.unlink(missing_ok=True)
        raise RuntimeError(f"ARCHIVO INCORRECTO descartado ({config['nombre']}): {motivo}. "
                           f"Se conserva la versión anterior.")

    if destino.exists() and not config.get("destino"):
        respaldo = CARPETA_DESTINO / "_respaldo_anterior" / config["nombre"]
        respaldo.parent.mkdir(exist_ok=True)
        shutil.copy(destino, respaldo)
    tmp.replace(destino)
    print(f"  [OK] {config['nombre']} guardado en {destino}")


# --------------------------------------------------------------------
# VALIDACIÓN Y GIT AUTO-PUSH
# --------------------------------------------------------------------

# --------------------------------------------------------------------
# REGISTRO (log) Y AVISOS — modo desatendido (v4, 04/11/2026)
# --------------------------------------------------------------------

class _Tee:
    """Duplica todo lo que se imprime a un archivo de registro (para revisar corridas
    que nadie estaba mirando)."""
    def __init__(self, *flujos):
        self.flujos = flujos
        self.encoding = getattr(flujos[0], "encoding", None) or "utf-8"

    def write(self, texto):
        for f in self.flujos:
            try:
                f.write(texto)
                f.flush()
            except Exception:
                pass
        return len(texto)

    def flush(self):
        for f in self.flujos:
            try:
                f.flush()
            except Exception:
                pass


def iniciar_log():
    """Guarda la salida en logs/descarga_AAAA-MM-DD.log y borra registros de más de 45 días."""
    CARPETA_LOGS.mkdir(exist_ok=True)
    limite = time.time() - 45 * 86400
    for viejo in CARPETA_LOGS.glob("*"):
        try:
            if viejo.stat().st_mtime < limite:
                viejo.unlink()
        except Exception:
            pass
    ruta = CARPETA_LOGS / f"descarga_{date.today():%Y-%m-%d}.log"
    f = open(ruta, "a", encoding="utf-8")
    f.write(f"\n===== Inicio {datetime.now():%Y-%m-%d %H:%M:%S} =====\n")
    sys.stdout = _Tee(sys.__stdout__, f)
    sys.stderr = _Tee(sys.__stderr__, f)
    return ruta


def avisar(asunto, detalle):
    """Deja constancia de un problema que el script no pudo resolver solo.
    1) archivo logs/ALERTA_*.txt   2) mensaje en pantalla de Windows (msg.exe, si existe)
    3) correo, SOLO si defines las variables de entorno ORAD_SMTP_HOST, ORAD_SMTP_USER,
       ORAD_SMTP_PASS y ORAD_ALERTA_A (opcionales: ORAD_SMTP_PORT, por defecto 587).
    Nunca escribas la contraseña dentro del script."""
    CARPETA_LOGS.mkdir(exist_ok=True)
    ruta = CARPETA_LOGS / f"ALERTA_{datetime.now():%Y-%m-%d_%H%M}.txt"
    ruta.write_text(f"{asunto}\n\n{detalle}\n", encoding="utf-8")
    print(f"\n🔔 ALERTA: {asunto}\n   Detalle guardado en: {ruta}")

    if os.name == "nt":
        try:
            subprocess.Popen(["msg", "*", "/TIME:900", f"ORAD dashboard: {asunto}"],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            pass

    host, usuario, clave, destino = (os.environ.get(k) for k in
                                     ("ORAD_SMTP_HOST", "ORAD_SMTP_USER", "ORAD_SMTP_PASS", "ORAD_ALERTA_A"))
    if host and usuario and clave and destino:
        try:
            msg = EmailMessage()
            msg["Subject"], msg["From"], msg["To"] = f"[ORAD] {asunto}", usuario, destino
            msg.set_content(detalle)
            with smtplib.SMTP(host, int(os.environ.get("ORAD_SMTP_PORT", "587")), timeout=30) as srv:
                srv.starttls()
                srv.login(usuario, clave)
                srv.send_message(msg)
            print("   Correo de aviso enviado.")
        except Exception as e:
            print(f"   (no se pudo enviar el correo de aviso: {e.__class__.__name__}: {e})")


# --------------------------------------------------------------------
# DESCARGA CON REINTENTOS
# --------------------------------------------------------------------

def descargar_con_reintentos(browser, config, intentos=3, pausa=5):
    """Reintenta un archivo con una pestaña NUEVA en cada intento (estado limpio del
    sitio del MEF). Retorna (True, None) o (False, motivo_del_último_error)."""
    ultimo = None
    for n in range(1, intentos + 1):
        page = browser.new_page()
        try:
            procesar_archivo(page, config)
            page.close()
            return True, None
        except Exception as e:
            ultimo = str(e)
            print(f"  [ERROR intento {n}/{intentos}] {config['nombre']}: {e}")
            if n == intentos:
                captura = REPO_RAIZ / f"error_{config['nombre']}.png"
                try:
                    page.screenshot(path=str(captura), full_page=True)
                    print(f"  [DIAGNÓSTICO] Captura guardada en: {captura}")
                except Exception:
                    print("  [DIAGNÓSTICO] No se pudo guardar la captura.")
            try:
                page.close()
            except Exception:
                pass
            if n < intentos:
                print(f"  Reintentando en {pausa * n} s...")
                time.sleep(pausa * n)
    return False, ultimo


# --------------------------------------------------------------------
# VALIDACIÓN FINAL Y GIT AUTO-PUSH
# --------------------------------------------------------------------

_MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7,
          "agosto": 8, "setiembre": 9, "septiembre": 9, "octubre": 10, "noviembre": 11,
          "diciembre": 12}


def hoy():
    return date.today()


def fecha_consulta(ruta):
    """Fecha de la consulta que trae el archivo ('04-noviembre-2026') o None."""
    texto = Path(ruta).read_bytes().decode("latin-1", errors="replace")
    m = re.search(r"Fecha de la Consulta:\s*(\d{1,2})-([A-Za-z]+)-(\d{4})", texto)
    if not m or m.group(2).lower() not in _MESES:
        return None
    return date(int(m.group(3)), _MESES[m.group(2).lower()], int(m.group(1)))


def validar_archivos():
    """
    Compuerta final antes de Git (v4). Para cada uno de los 17 archivos verifica:
      1) existe y pesa más de 5 KB;
      2) su CONTENIDO es el pedido (año, Gasto, filtro y nivel: validar_contenido);
      3) fue consultado HOY (un archivo de ayer no se sube como si fuera de hoy).
    Retorna (ok: bool, problemas: list).
    """
    TAMANIO_MINIMO = 5 * 1024  # 5 KB
    problemas = []

    for config in ARCHIVOS:
        nombre = config["nombre"]
        ruta = CARPETA_DESTINO / nombre
        if not ruta.exists():
            problemas.append(f"{nombre} — NO EXISTE")
            continue
        if ruta.stat().st_size < TAMANIO_MINIMO:
            problemas.append(f"{nombre} — DEMASIADO PEQUEÑO ({ruta.stat().st_size / 1024:.1f} KB, mínimo 5 KB)")
            continue
        ok, motivo = validar_contenido(ruta, config)
        if not ok:
            problemas.append(f"{nombre} — CONTENIDO INCORRECTO: {motivo}")
            continue
        f = fecha_consulta(ruta)
        if f != hoy():
            problemas.append(f"{nombre} — DESACTUALIZADO: consultado el {f or '(fecha ilegible)'}, hoy es {hoy()}")

    return len(problemas) == 0, problemas


def _git(*args):
    return subprocess.run(["git", *args], cwd=str(REPO_RAIZ), capture_output=True,
                          text=True, encoding="utf-8")


def git_push_si_completo():
    """
    git add xls/ → git commit → git push SOLO si los 17 archivos pasan la validación.
    Retorna: "subido" | "sin_cambios" | "bloqueado" | "error_git".
    """
    print("\n" + "=" * 60)
    print("VALIDACIÓN FINAL — verificando 17/17 archivos")
    print("=" * 60)

    ok, problemas = validar_archivos()
    if not ok:
        print(f"\n🚫 GIT PUSH BLOQUEADO — {len(problemas)} archivo(s) con problema:\n")
        for f in problemas:
            print(f"   ❌ {f}")
        print("\n⚠️  No se subió NADA a GitHub.")
        return "bloqueado", problemas

    mensaje_commit = f"data: actualización diaria XLS - {hoy():%d/%m/%Y}"
    print(f"\n✅ 17/17 archivos OK — procediendo con Git...\n")

    for cmd, descripcion in (
        (["add", "xls/"], "git add xls/"),
        (["commit", "-m", mensaje_commit], f'git commit -m "{mensaje_commit}"'),
        (["push", "origin", "main"], "git push origin main"),
    ):
        print(f"  ▶ {descripcion}")
        r = _git(*cmd)
        if r.stdout.strip():
            print(f"    {r.stdout.strip()}")
        if r.stderr.strip():
            print(f"    {r.stderr.strip()}")

        if r.returncode != 0:
            salida = r.stdout + r.stderr
            if any(t in salida for t in ("nothing to commit", "nothing added to commit",
                                         "no changes added to commit")):
                print("\nℹ️  Los archivos XLS no cambiaron respecto al último commit — nada nuevo que subir.")
                return "sin_cambios", []
            print(f"\n❌ Error en '{descripcion}' (código {r.returncode}). "
                  f"Revisa tu conexión, credenciales de Git o el estado del repo.")
            return "error_git", [f"'{descripcion}' falló: {salida.strip()[:300]}"]

    print(f"\n🚀 Push exitoso. Commit: \"{mensaje_commit}\"")
    return "subido", []


# --------------------------------------------------------------------
# VERIFICACIÓN DE LA PÁGINA PUBLICADA (y reparación automática)
# --------------------------------------------------------------------

def _normalizado(b):
    """Git en Windows convierte CRLF→LF al subir: se compara sin esa diferencia."""
    return hashlib.sha256(b.replace(b"\r\n", b"\n")).hexdigest()


def contenido_publicado(nombre):
    req = urllib.request.Request(f"{PAGES_URL}/xls/{nombre}?v={int(time.time())}",
                                 headers={"Cache-Control": "no-cache", "Pragma": "no-cache",
                                          "User-Agent": "orad-dashboard-verificacion"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read()


def pendientes_en_pagina():
    """Archivos centinela cuyo contenido publicado NO coincide con el local."""
    pend = []
    for nombre in ARCHIVOS_CENTINELA:
        try:
            publicado = contenido_publicado(nombre)
        except Exception as e:
            pend.append(f"{nombre} (no se pudo leer: {e.__class__.__name__})")
            continue
        if _normalizado(publicado) != _normalizado((CARPETA_DESTINO / nombre).read_bytes()):
            pend.append(nombre)
    return pend


def esperar_publicacion(max_seg, cada=None):
    cada = INTERVALO_PAGES_SEG if cada is None else cada
    t0 = time.time()
    while True:
        pend = pendientes_en_pagina()
        if not pend:
            return True, []
        if time.time() - t0 >= max_seg:
            return False, pend
        print(f"  ... la página aún no muestra los datos nuevos ({', '.join(pend)}); espero {cada} s")
        time.sleep(cada)


def verificar_y_reparar_publicacion():
    """Comprueba que GitHub Pages ya sirve los archivos de hoy. Si no aparecen tras
    ESPERA_PAGES_SEG, fuerza UN despliegue con un commit vacío (lo que se hizo a mano el
    04/11/2026, cuando GitHub no disparó el build del último push) y vuelve a esperar."""
    print("\n" + "=" * 60)
    print("VERIFICANDO LA PÁGINA PUBLICADA")
    print("=" * 60)
    ok, pend = esperar_publicacion(ESPERA_PAGES_SEG)
    if ok:
        print("✅ La página ya muestra los datos de hoy.")
        return True, []

    print(f"\n⚠️  Tras {ESPERA_PAGES_SEG} s la página sigue con datos viejos ({', '.join(pend)}).")
    print("   Fuerzo el despliegue con un commit vacío (una sola vez)...")
    for cmd in (["commit", "--allow-empty", "-m", "chore: forzar despliegue de Pages (automático)"],
                ["push", "origin", "main"]):
        r = _git(*cmd)
        print(f"  ▶ git {' '.join(cmd[:2])}: {(r.stdout + r.stderr).strip()[:200]}")
        if r.returncode != 0:
            return False, [f"no se pudo forzar el despliegue: {(r.stdout + r.stderr).strip()[:200]}"]

    ok, pend = esperar_publicacion(ESPERA_PAGES_SEG)
    if ok:
        print("✅ Tras forzar el despliegue, la página ya muestra los datos de hoy.")
        return True, []
    return False, [f"la página sigue sin mostrar: {', '.join(pend)}"]


# --------------------------------------------------------------------
# MAIN
# --------------------------------------------------------------------

def main():
    import argparse
    ap = argparse.ArgumentParser(description="Descarga diaria de los 17 XLS del MEF.")
    ap.add_argument("--solo", nargs="+", metavar="ARCHIVO",
                    help="descargar solo estos archivos (ej.: --solo rubro_pliego ue_pliego)")
    ap.add_argument("--sin-pausa", action="store_true",
                    help="no esperar ENTER al terminar (para ejecución programada)")
    ap.add_argument("--headless", action="store_true", help="sin ventana de navegador")
    ap.add_argument("--reintentos", type=int, default=3, help="intentos por archivo (defecto 3)")
    ap.add_argument("--sin-git", action="store_true", help="descargar y validar, sin tocar Git")
    ap.add_argument("--sin-verificar-web", action="store_true",
                    help="no comprobar la página publicada después del push")
    args = ap.parse_args()

    ruta_log = iniciar_log()
    print(f"Registro de esta corrida: {ruta_log}")

    lista = ARCHIVOS
    if args.solo:
        pedidos = {a.lower().removesuffix(".xls") for a in args.solo}
        lista = [a for a in ARCHIVOS if a["nombre"].removesuffix(".xls") in pedidos]
        desconocidos = pedidos - {a["nombre"].removesuffix(".xls") for a in lista}
        if desconocidos:
            raise SystemExit(f"Nombres no reconocidos: {sorted(desconocidos)}. Válidos: {NOMBRES_ESPERADOS}")
    preparar_carpeta()
    fallidos = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=args.headless, slow_mo=0 if args.headless else 200)
        for i, config in enumerate(lista, 1):
            print(f"\n[{i}/{len(lista)}] Procesando {config['nombre']} ...")
            ok, motivo = descargar_con_reintentos(browser, config, args.reintentos)
            if not ok:
                fallidos.append((config["nombre"], motivo))
        if not args.sin_pausa:
            input("\nDescarga terminada. Presiona ENTER para cerrar el navegador...")
        browser.close()

    # Si algún archivo no se pudo obtener tras los reintentos NO se sube nada: los demás
    # quedarían de hoy y ese de ayer, con la página mezclando fechas sin que nadie lo note.
    if fallidos:
        detalle = "\n".join(f"- {n}: {m}" for n, m in fallidos)
        print("\n⚠️  NO se ejecuta Git: archivos que no se pudieron descargar bien:\n" + detalle)
        avisar(f"{len(fallidos)} archivo(s) del MEF fallaron tras {args.reintentos} intentos",
               "Estos archivos conservan su versión anterior y NO se subió nada a GitHub:\n" + detalle +
               f"\n\nRepetir solo esos: python scripts\\descargar_xls_mef.py --solo "
               + " ".join(n.removesuffix('.xls') for n, _ in fallidos))
        sys.exit(1)

    if args.sin_git:
        return

    estado, problemas = git_push_si_completo()
    if estado in ("bloqueado", "error_git"):
        avisar("No se subió la actualización diaria a GitHub", "\n".join(f"- {x}" for x in problemas))
        sys.exit(1)

    if not args.sin_verificar_web:
        ok, problemas = verificar_y_reparar_publicacion()
        if not ok:
            avisar("La página publicada NO muestra los datos de hoy",
                   "\n".join(f"- {x}" for x in problemas) +
                   "\n\nRevisa https://github.com/jdrq/orad-dashboard/actions (¿falló el despliegue de Pages?).")
            sys.exit(2)

    print("\n🎉 Actualización diaria completa y verificada en la página publicada.")


if __name__ == "__main__":
    main()
