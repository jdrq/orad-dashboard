"""
descargar_xls_mef.py
---------------------
Automatiza la descarga de los 13 archivos .xls de Consulta Amigable (MEF)
que alimentan el dashboard de ORAD - GORE Lambayeque.

MODO DE USO (Fase 2a - manual, supervisado):
    python descargar_xls_mef.py

Requiere:
    pip install playwright
    playwright install chromium

Arquitecto: este script corre en modo "headed" (con ventana visible) a
propósito durante esta fase - así Juan puede observar cada descarga y
detectar fallos en el momento, en vez de un cron job a ciegas.

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
from pathlib import Path

# --------------------------------------------------------------------
# CONFIGURACIÓN
# --------------------------------------------------------------------
CARPETA_DESTINO = Path(__file__).resolve().parent.parent / "xls"
ANIO = "2026"
URL_BASE = f"https://apps5.mineco.gob.pe/transparencia/Navegador/default.aspx?y={ANIO}&ap=Proyecto"
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
#          final (Rubro, Función), o None
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
]


def preparar_carpeta():
    CARPETA_DESTINO.mkdir(exist_ok=True)


def procesar_archivo(page, config):
    """
    Ejecuta la secuencia completa de un archivo: navegar, filtrar,
    drillear la jerarquía Gobierno, pivotar (si aplica) y exportar.

    Usa frame_locator (no page.frame(name=...)) porque se auto-resuelve
    solo ante cada recarga del frame - esto es lo que finalmente
    resolvió todos los problemas de "frame detached" de las pruebas
    anteriores, confirmado con playwright codegen real.
    """
    page.goto(URL_BASE)
    fl = page.frame_locator(FRAME_SELECTOR)

    fl.get_by_role("cell", name="TOTAL", exact=True).click()
    # Nota: NO seleccionamos Actividades/Proyectos por dropdown porque
    # la URL ya lo fija con ?ap=Proyecto - hacerlo de nuevo disparaba
    # una recarga completa de página que descarrilaba el resto del flujo.
    try:
        page.wait_for_load_state("networkidle", timeout=5000)
    except Exception:
        pass

    # Bajar por la jerarquía Gobierno -> Sector -> Pliego -> Ejecutora.
    # El botón se busca POR SU ETIQUETA VISIBLE en cada paso (no por ID
    # fijo) porque el ID real cambia en cada nivel: BtnTipoGobierno,
    # BtnSector, BtnPliego, etc. - confirmado en vivo el 06/07/2026.
    for etiqueta_boton, texto_fila in config["pasos"]:
        fl.get_by_role("button", name=etiqueta_boton, exact=True).click()
        try:
            page.wait_for_load_state("networkidle", timeout=5000)
        except Exception:
            pass
        time.sleep(1)  # margen extra: el servidor del MEF a veces es lento

        # Clic en la fila con hasta 3 intentos: si el checkpoint no
        # confirma la selección (fallo transitorio del servidor), se
        # vuelve a clickear en vez de rendirse al primer intento.
        confirmado = False
        for intento in range(3):
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
                  f"marcó, volviendo a clickear...")

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

    with page.expect_download(timeout=30000) as descarga_info:
        fl.get_by_role("link", name="Exportar").click()
    descarga = descarga_info.value

    destino = CARPETA_DESTINO / config["nombre"]

    if destino.exists():
        respaldo = CARPETA_DESTINO / "_respaldo_anterior" / config["nombre"]
        respaldo.parent.mkdir(exist_ok=True)
        shutil.copy(destino, respaldo)

    descarga.save_as(destino)
    print(f"  [OK] {config['nombre']} guardado en {destino}")


def main():
    preparar_carpeta()
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=200)
        page = browser.new_page()

        for i, config in enumerate(ARCHIVOS, 1):
            print(f"\n[{i}/13] Procesando {config['nombre']} ...")
            try:
                procesar_archivo(page, config)
            except Exception as e:
                print(f"  [ERROR] Falló {config['nombre']}: {e}")
                captura = Path(f"error_{config['nombre']}.png")
                try:
                    page.screenshot(path=str(captura), full_page=True)
                    print(f"  [DIAGNÓSTICO] Captura guardada en: {captura.resolve()}")
                except Exception:
                    print("  [DIAGNÓSTICO] No se pudo guardar la captura.")
                print("  Deteniendo el script para revisar manualmente.")
                break

        input("\nProceso terminado. Presiona ENTER para cerrar el navegador...")
        browser.close()


if __name__ == "__main__":
    main()
