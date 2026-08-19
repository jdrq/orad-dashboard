"""
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
import time
import shutil
import subprocess
from datetime import date
from pathlib import Path

# --------------------------------------------------------------------
# CONFIGURACIÓN
# --------------------------------------------------------------------
CARPETA_DESTINO = Path(__file__).resolve().parent.parent / "xls"
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
            verificar=lambda tf=texto_fila: fl.get_by_role("cell", name=tf).first.is_visible()
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
            if intento == 0:
                fl.get_by_role("cell", name=texto_fila).click()
            else:
                fl.get_by_role("cell", name=texto_fila).click()
                time.sleep(0.3)
                fl.get_by_role("cell", name=texto_fila).click()
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
                    fila = fl.locator(f"tr:has-text('{texto_fila}')").first
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
        ultimo_texto = config["pasos"][-1][1].rstrip(":")
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

    destino = CARPETA_DESTINO / config["nombre"]

    if destino.exists():
        respaldo = CARPETA_DESTINO / "_respaldo_anterior" / config["nombre"]
        respaldo.parent.mkdir(exist_ok=True)
        shutil.copy(destino, respaldo)

    descarga.save_as(destino)
    print(f"  [OK] {config['nombre']} guardado en {destino}")


# --------------------------------------------------------------------
# VALIDACIÓN Y GIT AUTO-PUSH
# --------------------------------------------------------------------

def validar_archivos():
    """
    Verifica que los 17 archivos esperados existan en xls/ y tengan
    tamaño > 0 bytes. Retorna (ok: bool, faltantes: list).
    Un archivo descargado con error en el MEF a veces llega como HTML
    de 1-2 KB — el tamaño mínimo de 5 KB descarta esos casos.
    """
    TAMANIO_MINIMO = 5 * 1024  # 5 KB
    faltantes = []

    for nombre in NOMBRES_ESPERADOS:
        ruta = CARPETA_DESTINO / nombre
        if not ruta.exists():
            faltantes.append(f"{nombre} — NO EXISTE")
        elif ruta.stat().st_size < TAMANIO_MINIMO:
            faltantes.append(
                f"{nombre} — DEMASIADO PEQUEÑO "
                f"({ruta.stat().st_size / 1024:.1f} KB, mínimo 5 KB)"
            )

    return len(faltantes) == 0, faltantes


def git_push_si_completo():
    """
    Ejecuta git add xls/ → git commit → git push SOLO si los 17
    archivos pasan la validación. Si falla cualquiera, no toca Git.

    Usa subprocess con cwd apuntando a la raíz del repo (un nivel
    arriba del script, que vive en scripts/).
    """
    REPO_RAIZ = Path(__file__).resolve().parent.parent

    print("\n" + "=" * 60)
    print("VALIDACIÓN FINAL — verificando 17/17 archivos")
    print("=" * 60)

    ok, faltantes = validar_archivos()

    if not ok:
        print(f"\n🚫 GIT PUSH BLOQUEADO — {len(faltantes)} archivo(s) con problema:\n")
        for f in faltantes:
            print(f"   ❌ {f}")
        print(
            "\n⚠️  No se subió NADA a GitHub."
            "\nResuelve los archivos marcados y vuelve a ejecutar el script,"
            "\no haz git add / commit / push manualmente una vez corregido."
        )
        return

    # 17/17 OK — proceder con Git
    hoy = date.today().strftime("%d/%m/%Y")
    mensaje_commit = f"data: actualización diaria XLS - {hoy}"

    print(f"\n✅ 17/17 archivos OK — procediendo con Git...\n")

    comandos = [
        (["git", "add", "xls/"], "git add xls/"),
        (["git", "commit", "-m", mensaje_commit], f'git commit -m "{mensaje_commit}"'),
        (["git", "push", "origin", "main"], "git push origin main"),
    ]

    for cmd, descripcion in comandos:
        print(f"  ▶ {descripcion}")
        resultado = subprocess.run(
            cmd,
            cwd=str(REPO_RAIZ),
            capture_output=True,
            text=True,
            encoding="utf-8",
        )

        if resultado.stdout.strip():
            print(f"    {resultado.stdout.strip()}")
        if resultado.stderr.strip():
            print(f"    {resultado.stderr.strip()}")

        # git commit devuelve código 1 si no hay cambios (nada que commitear).
        # Eso NO es un error real — puede pasar si los datos del MEF no
        # cambiaron desde el último push. Se detecta por el mensaje.
        if resultado.returncode != 0:
            sin_cambios = (
                "nothing to commit" in resultado.stdout
                or "nothing to commit" in resultado.stderr
                or "nothing added to commit" in resultado.stdout
                or "nothing added to commit" in resultado.stderr
                or "no changes added to commit" in resultado.stdout
                or "no changes added to commit" in resultado.stderr
            )
            if sin_cambios:
                print(
                    "\nℹ️  Los archivos XLS no cambiaron respecto al último"
                    " commit — no hay nada nuevo que subir. Esto es normal."
                )
                return

            # Error real de Git
            print(
                f"\n❌ Error en '{descripcion}' "
                f"(código {resultado.returncode}). "
                f"Revisa tu conexión, credenciales de Git o el estado del repo."
            )
            return

    print(f"\n🚀 Push exitoso — dashboard actualizado en GitHub Pages.")
    print(f"   Commit: \"{mensaje_commit}\"")


# --------------------------------------------------------------------
# MAIN
# --------------------------------------------------------------------

def main():
    preparar_carpeta()
    exito_total = True

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=200)
        page = browser.new_page()

        for i, config in enumerate(ARCHIVOS, 1):
            print(f"\n[{i}/17] Procesando {config['nombre']} ...")
            try:
                procesar_archivo(page, config)
            except Exception as e:
                print(f"  [ERROR] Falló {config['nombre']}: {e}")
                captura = Path(__file__).resolve().parent.parent / f"error_{config['nombre']}.png"
                try:
                    page.screenshot(path=str(captura), full_page=True)
                    print(f"  [DIAGNÓSTICO] Captura guardada en: {captura}")
                except Exception:
                    print("  [DIAGNÓSTICO] No se pudo guardar la captura.")
                print("  Deteniendo el script para revisar manualmente.")
                exito_total = False
                break

        input("\nDescarga terminada. Presiona ENTER para cerrar el navegador...")
        browser.close()

    # CORREGIDO en sesión 14 (19/08/2026): antes esto se llamaba SIEMPRE,
    # sin importar si el loop de arriba terminó con error. Como la
    # validación de archivos solo revisa que EXISTAN y pesen > 0 (no que se
    # hayan actualizado HOY), un archivo viejo de un día anterior pasaba la
    # validación igual, y el script intentaba subir a Git como si todo
    # hubiera salido bien - riesgo real de que el dashboard quedara con un
    # bloque desactualizado sin que nadie se entere. Ahora, si algún
    # archivo falló, NO se valida ni se sube nada - se avisa explícitamente
    # cuál quedó con datos viejos.
    if not exito_total:
        print(
            "\n⚠️  El script se detuvo por un error antes de terminar los "
            "17 archivos (ver [ERROR] arriba). NO se ejecuta Git - los "
            "archivos que no llegaron a descargarse hoy conservan su "
            "versión de un día anterior en xls/, y el dashboard publicado "
            "seguirá mostrando esos bloques con datos desactualizados "
            "hasta que corras el script de nuevo con éxito."
        )
        return

    git_push_si_completo()


if __name__ == "__main__":
    main()
