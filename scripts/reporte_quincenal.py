#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# =============================================================================
# reporte_quincenal.py  —  informe quincenal en Word desde data/serie/*.csv
# ORAD - Gobierno Regional de Lambayeque
#
# Compara dos "fotos" de la serie diaria: el corte BASE (último dato anterior a la
# quincena) y el CIERRE (último dato dentro de la quincena). Genera un .docx
# EDITABLE en reportes/ (carpeta que no se sube a GitHub: es un borrador sin firmar).
#
# USO (desde la raíz del repo, o doble clic en generar_reporte.bat)
#   python scripts\reporte_quincenal.py                       última quincena completa
#   python scripts\reporte_quincenal.py --quincena 2026-11-1  1ra quincena de noviembre
#   python scripts\reporte_quincenal.py --desde 2026-11-01 --hasta 2026-11-26   rango libre
#
# DATOS DEL ENCABEZADO (opcional): crear reportes/datos_firma.json con
#   {"numero": "001-2026-XXX", "para": "Nombre", "cargo_para": "Cargo",
#    "de": "Nombre", "cargo_de": "Cargo"}
# Requiere: pip install python-docx matplotlib
# =============================================================================

import argparse
import calendar
import io
import json
import sys
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import alertas as al                       # noqa: E402

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.dates as mdates
    from docx import Document
    from docx.enum.section import WD_ORIENT  # noqa: F401
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
except ImportError as e:
    raise SystemExit(f"Falta una librería ({e.name}). Instalar una sola vez:  pip install python-docx matplotlib")

RAIZ = Path(__file__).resolve().parent.parent
SALIDA = RAIZ / "reportes"
PERU = timezone(timedelta(hours=-5))
VINO, GRIS = "7B1E1E", "F2F2F2"
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "setiembre", "octubre", "noviembre", "diciembre"]


# ---------------------------------------------------------------- utilidades
def ultima_quincena_completa(hoy):
    if hoy.day > 15:
        return hoy.replace(day=1), hoy.replace(day=15)
    fin = hoy.replace(day=1) - timedelta(days=1)
    return fin.replace(day=16), fin


def rango_quincena(texto):
    a, m, q = (int(x) for x in texto.split("-"))
    if q not in (1, 2):
        raise SystemExit("--quincena es AAAA-MM-1 o AAAA-MM-2")
    return (date(a, m, 1), date(a, m, 15)) if q == 1 else (date(a, m, 16), date(a, m, calendar.monthrange(a, m)[1]))


def cortes(fechas, desde, hasta):
    """(base, cierre): último dato antes del periodo y último dato dentro de él."""
    previas = [f for f in fechas if f < desde.isoformat()]
    dentro = [f for f in fechas if desde.isoformat() <= f <= hasta.isoformat()]
    if not previas or not dentro:
        raise SystemExit(f"No hay datos suficientes para {desde} – {hasta}. La serie va de {fechas[0]} a {fechas[-1]}.")
    return max(previas), max(dentro)


n = al.num
f_fecha = al.f_fecha


def mill(v, signo=False):
    return f"{v / 1e6:+,.2f}" if signo else f"{v / 1e6:,.2f}"


def pct(v):
    return f"{v:.1f}%"


def avance(f):
    return n(f, "devengado") / n(f, "pim") * 100 if n(f, "pim") else 0.0


def indice(filas, fecha, *claves):
    return {tuple(f[c] for c in claves): f for f in filas if f["fecha"] == fecha}


def corto(texto, largo=90):
    return texto if len(texto) <= largo else texto[:largo].rsplit(" ", 1)[0] + "…"


def cambios_alertas(antes, despues, proy0, proy1):
    """Clasifica las diferencias entre las alertas de dos cortes. Distingue 'resuelta' de
    'salió de la lista sin resolverse' (p. ej. le redujeron el PIM bajo el umbral)."""
    clave = lambda a: (a["regla"], a["ambito"])
    ka, kd = {clave(a) for a in antes}, {clave(a) for a in despues}
    ambitos_d = {a["ambito"] for a in despues}
    p0 = {f'{k[1]}': v for k, v in proy0.items()}
    p1 = {f'{k[1]}': v for k, v in proy1.items()}
    nuevas = [a for a in despues if clave(a) not in ka]
    salidas = []
    for a in antes:
        if clave(a) in kd:
            continue
        motivo = "Ya no cumple la condición de la alerta"
        if a["regla"] in ("A4", "A5"):
            cod = a["ambito"].split()[0]
            x, y = p0.get(cod), p1.get(cod)
            if a["ambito"] in ambitos_d:
                motivo = "Cambió de tipo de alerta; sigue observado"
            elif y is None:
                motivo = "Ya no figura en la consulta"
            elif x and n(y, "devengado") > n(x, "devengado"):
                motivo = f"Registró nuevo devengado (S/ {n(y, 'devengado') - n(x, 'devengado'):,.0f})"
            elif n(y, "pim") < al.PIM_PROYECTO_GRANDE:
                motivo = f"NO resuelta: el PIM se redujo a S/ {n(y, 'pim'):,.0f} (bajo el umbral)"
        salidas.append(dict(a, motivo=motivo))
    return nuevas, salidas


# ---------------------------------------------------------------- Word
def sombrear(celda, color):
    tcPr = celda._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    tcPr.append(shd)


def repetir_encabezado(fila):
    marca = OxmlElement("w:tblHeader")
    fila._tr.get_or_add_trPr().append(marca)


def tabla(doc, encabezados, filas, anchos, num_desde=1, total=False):
    t = doc.add_table(rows=1, cols=len(encabezados))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for i, fila in enumerate([encabezados] + filas):
        celdas = t.rows[0].cells if i == 0 else t.add_row().cells
        for j, valor in enumerate(fila):
            c = celdas[j]
            c.width = Cm(anchos[j])
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if j >= num_desde and i > 0:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r = p.add_run(str(valor))
            r.font.size = Pt(8.5)
            if i == 0:
                r.bold = True
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                sombrear(c, VINO)
            elif total and i == len(filas):
                r.bold = True
                sombrear(c, GRIS)
    for j, ancho in enumerate(anchos):
        t._tbl.tblGrid.findall(qn("w:gridCol"))[j].set(qn("w:w"), str(int(ancho * 567)))
    repetir_encabezado(t.rows[0])
    return t


def vinetas(doc, textos):
    for x in textos:
        doc.add_paragraph(x, style="List Bullet")


def grafico(serie, base, cierre):
    xs = [date.fromisoformat(f["fecha"]) for f in serie]
    ys = [n(f, "devengado") / 1e6 for f in serie]
    fig, ax = plt.subplots(figsize=(6.5, 2.6), dpi=180)
    ax.plot(xs, ys, color="#" + VINO, linewidth=1.8)
    ax.axvspan(date.fromisoformat(base), date.fromisoformat(cierre), color="#E8A33D", alpha=0.25, label="Periodo del informe")
    ax.set_ylabel("Millones de S/", fontsize=8)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d/%m"))
    ax.tick_params(labelsize=7)
    ax.grid(alpha=0.25)
    ax.spines[["top", "right"]].set_visible(False)
    ax.legend(fontsize=7, frameon=False, loc="upper left")
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png")
    plt.close(fig)
    buf.seek(0)
    return buf


def numero_pagina(parrafo):
    campo = OxmlElement("w:fldSimple")
    campo.set(qn("w:instr"), "PAGE")
    r = OxmlElement("w:r")
    texto = OxmlElement("w:t")
    texto.text = "1"
    r.append(texto)
    campo.append(r)
    parrafo._p.append(campo)


def preparar_documento():
    doc = Document()
    zoom = doc.settings.element.find(qn("w:zoom"))
    if zoom is not None:                  # la plantilla de python-docx omite este atributo obligatorio
        zoom.set(qn("w:percent"), "100")
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21), Cm(29.7)
    s.left_margin = s.right_margin = Cm(2.2)
    s.top_margin = s.bottom_margin = Cm(2)
    normal = doc.styles["Normal"]
    normal.font.name, normal.font.size = "Calibri", Pt(11)
    normal.paragraph_format.space_after = Pt(6)
    for nombre, tam in (("Heading 1", 13), ("Heading 2", 11.5)):
        e = doc.styles[nombre]
        e.font.name, e.font.size, e.font.bold = "Calibri", Pt(tam), True
        e.font.color.rgb = RGBColor.from_string(VINO)
        rf = e.element.rPr.rFonts
        for a in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
            rf.attrib.pop(qn(a), None)
        e.paragraph_format.space_before, e.paragraph_format.space_after = Pt(12), Pt(6)
    pie = s.footer.paragraphs[0]
    pie.alignment = WD_ALIGN_PARAGRAPH.CENTER
    numero_pagina(pie)
    return doc


def leer_firma():
    datos = {"numero": "______", "para": "[Nombre del destinatario]", "cargo_para": "[Cargo]",
             "de": "[Nombre del remitente]", "cargo_de": "[Cargo]"}
    ruta = SALIDA / "datos_firma.json"
    if ruta.exists():
        datos.update(json.loads(ruta.read_text(encoding="utf-8-sig")))
    return datos


# ---------------------------------------------------------------- informe
def generar(desde, hasta, carpeta=al.SERIE, salida=None):
    leer = lambda nombre: al.leer(nombre, carpeta)
    pliego = sorted(leer("fact_pliego_diario"), key=lambda f: f["fecha"])
    base, cierre = cortes([f["fecha"] for f in pliego], desde, hasta)
    p0 = next(f for f in pliego if f["fecha"] == base)
    p1 = next(f for f in pliego if f["fecha"] == cierre)
    ue, rubro = leer("fact_ue_diario"), leer("fact_rubro_diario")
    funcion, proy = leer("fact_funcion_diario"), leer("fact_proyecto_diario")
    a0, a1 = al.calcular(carpeta, hasta=base), al.calcular(carpeta, hasta=cierre)

    dev_q = n(p1, "devengado") - n(p0, "devengado")
    pim_q = n(p1, "pim") - n(p0, "pim")
    saldo = n(p1, "pim") - n(p1, "devengado")
    nombre_periodo = (f"del {desde.day} al {hasta.day} de {MESES[hasta.month - 1]} de {hasta.year}"
                      if (desde.year, desde.month) == (hasta.year, hasta.month)
                      else f"del {f_fecha(desde.isoformat())} al {f_fecha(hasta.isoformat())}")

    doc = preparar_documento()
    firma = leer_firma()
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run(f"INFORME Nº {firma['numero']}")
    r.bold, r.font.size = True, Pt(14)
    hoy = datetime.now(PERU).date()
    for etiqueta, valor in (("PARA", f"{firma['para']} — {firma['cargo_para']}"),
                            ("DE", f"{firma['de']} — {firma['cargo_de']}"),
                            ("ASUNTO", f"Ejecución de inversiones del Pliego 452, {nombre_periodo}"),
                            ("FECHA", f"Chiclayo, {hoy.day} de {MESES[hoy.month - 1]} de {hoy.year}")):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(2.6))
        p.paragraph_format.left_indent, p.paragraph_format.first_line_indent = Cm(2.6), Cm(-2.6)
        p.add_run(etiqueta).bold = True
        p.add_run(f"\t{valor}")

    doc.add_heading("1. Objetivo y alcance", 1)
    doc.add_paragraph(
        f"Informar la ejecución presupuestal de las inversiones del Gobierno Regional de Lambayeque (Pliego 452) "
        f"{nombre_periodo}. Las cifras corresponden a la consulta de Ejecución del Gasto, filtro «Sólo Proyectos», "
        f"de Consulta Amigable del MEF. Se comparan la consulta del {f_fecha(base)} (base) y la del "
        f"{f_fecha(cierre)} (cierre del periodo).")

    doc.add_heading("2. Resumen ejecutivo", 1)
    act, ant = a1["ranking_gores"]["actual"], al.calcular(carpeta, hasta=base)["ranking_gores"]["actual"]
    resumen = [
        f"El devengado acumulado llegó a S/ {mill(n(p1, 'devengado'))} millones, un avance de {pct(avance(p1))} "
        f"del PIM ({avance(p1) - avance(p0):+.1f} puntos en el periodo).",
        f"En el periodo se devengaron S/ {mill(dev_q)} millones.",
        f"El PIM {'aumentó' if pim_q > 0 else 'se redujo' if pim_q < 0 else 'no varió'}"
        + (f" en S/ {mill(abs(pim_q))} millones y quedó en" if pim_q else " y se mantiene en")
        + f" S/ {mill(n(p1, 'pim'))} millones.",
        f"Quedan S/ {mill(saldo)} millones por devengar al cierre del periodo.",
    ]
    if act and ant:
        resumen.append(f"Por avance, el Gobierno Regional ocupa el puesto {act['puesto']} de {act['total']} "
                       f"gobiernos regionales (en la base: {ant['puesto']}).")
    vinetas(doc, resumen)
    filas = []
    for etiqueta, campo in (("PIM", "pim"), ("Certificación", "certificacion"), ("Compromiso anual", "compromiso_anual"),
                            ("Devengado", "devengado"), ("Girado", "girado")):
        filas.append([etiqueta, mill(n(p0, campo)), mill(n(p1, campo)), mill(n(p1, campo) - n(p0, campo), True)])
    filas.append(["Avance (Devengado / PIM)", pct(avance(p0)), pct(avance(p1)), f"{avance(p1) - avance(p0):+.1f} pts"])
    tabla(doc, ["Millones de S/", f_fecha(base), f_fecha(cierre), "Variación"], filas, [6, 3.3, 3.3, 3.3])

    doc.add_heading("3. Evolución diaria del devengado", 1)
    serie = [f for f in pliego if f["fecha"] <= cierre]
    doc.add_picture(grafico(serie, base, cierre), width=Cm(16))

    doc.add_heading("4. Ejecución por unidad ejecutora", 1)
    u0 = indice(ue, base, "ue_codigo")
    u1 = sorted((f for f in ue if f["fecha"] == cierre), key=lambda f: -n(f, "pim"))
    filas = []
    for f in u1:
        x = u0.get((f["ue_codigo"],))
        filas.append([f"{f['ue_codigo']} {corto(f['ue_nombre'], 46)}", mill(n(f, "pim")), mill(n(f, "devengado")),
                      pct(avance(f)), mill(n(f, "devengado") - (n(x, "devengado") if x else 0), True),
                      mill(n(f, "pim") - (n(x, "pim") if x else 0), True)])
    filas.append(["TOTAL PLIEGO", mill(n(p1, "pim")), mill(n(p1, "devengado")), pct(avance(p1)),
                  mill(dev_q, True), mill(pim_q, True)])
    tabla(doc, ["Unidad ejecutora (millones de S/)", "PIM", "Devengado", "Avance", "Dev. periodo", "Var. PIM"],
          filas, [6.6, 1.9, 2, 1.6, 2.2, 1.9], total=True)

    def seccion_simple(titulo, datos, codigo, nombre, encabezado, extra=lambda f: True):
        doc.add_heading(titulo, 1)
        antes = indice([f for f in datos if extra(f)], base, codigo)
        ahora = sorted((f for f in datos if f["fecha"] == cierre and extra(f) and (n(f, "pim") or n(f, "devengado"))),
                       key=lambda f: -n(f, "pim"))
        filas = [[corto(f[nombre], 52), mill(n(f, "pim")), mill(n(f, "devengado")), pct(avance(f)),
                  mill(n(f, "devengado") - (n(antes[(f[codigo],)], "devengado") if (f[codigo],) in antes else 0), True)]
                 for f in ahora]
        tabla(doc, [encabezado, "PIM", "Devengado", "Avance", "Dev. periodo"], filas, [8, 2, 2.1, 1.7, 2.3])

    seccion_simple("5. Ejecución por fuente (rubro)", rubro, "rubro_codigo", "rubro_nombre",
                   "Rubro (millones de S/)", lambda f: f["ambito"] == "PLIEGO")
    seccion_simple("6. Ejecución por función", funcion, "funcion_codigo", "funcion_nombre", "Función (millones de S/)")

    doc.add_heading("7. Proyectos de la Sede Central", 1)
    j0, j1 = indice(proy, base, "ue_codigo", "proyecto_codigo"), indice(proy, cierre, "ue_codigo", "proyecto_codigo")
    delta = lambda k, c: n(j1[k], c) - (n(j0[k], c) if k in j0 else 0)
    top = [k for k in sorted(j1, key=lambda k: -delta(k, "devengado"))[:10] if delta(k, "devengado") > 0]
    doc.add_heading("7.1 Mayor devengado en el periodo", 2)
    tabla(doc, ["Proyecto (millones de S/)", "PIM", "Dev. periodo", "Avance"],
          [[f"{k[1]} {corto(j1[k]['proyecto_nombre'])}", mill(n(j1[k], "pim")), mill(delta(k, "devengado"), True),
            pct(avance(j1[k]))] for k in top], [10.4, 1.9, 2.2, 1.6])
    mod = [k for k in sorted(j1, key=lambda k: -abs(delta(k, "pim")))[:8] if abs(delta(k, "pim")) >= 500_000]
    doc.add_heading("7.2 Principales modificaciones del PIM", 2)
    if mod:
        tabla(doc, ["Proyecto (millones de S/)", "PIM base", "PIM cierre", "Variación"],
              [[f"{k[1]} {corto(j1[k]['proyecto_nombre'])}", mill(n(j0[k], "pim") if k in j0 else 0),
                mill(n(j1[k], "pim")), mill(delta(k, "pim"), True)] for k in mod], [10, 2, 2.1, 2])
    else:
        doc.add_paragraph("No hubo modificaciones de S/ 500 mil o más en proyectos de la Sede Central.")

    doc.add_heading("8. Alertas de ejecución", 1)
    doc.add_paragraph(f"Al {f_fecha(cierre)} hay {a1['total']} alertas, {a1['altas']} de nivel alto "
                      f"(en la base: {a0['total']}). Los umbrales son criterios de seguimiento interno, no normativos.")
    for cod, titulo in al.REGLAS.items():
        grupo = [a for a in a1["alertas"] if a["regla"] == cod]
        if grupo:
            doc.add_heading(f"{titulo} ({len(grupo)})", 2)
            tabla(doc, ["Ámbito", "Situación", "Nivel"], [[corto(a["ambito"]), a["mensaje"], a["nivel"]] for a in grupo],
                  [7.6, 7.1, 1.4], num_desde=9)
    nuevas, salidas = cambios_alertas(a0["alertas"], a1["alertas"], j0, j1)
    doc.add_heading("Cambios respecto de la base", 2)
    if nuevas or salidas:
        tabla(doc, ["Ámbito", "Cambio"],
              [[corto(a["ambito"]), f"Nueva: {al.REGLAS[a['regla']].lower()}"] for a in nuevas] +
              [[corto(a["ambito"]), f"Salió ({al.REGLAS[a['regla']].lower()}). {a['motivo']}"] for a in salidas],
              [8.2, 7.9], num_desde=9)
    else:
        doc.add_paragraph("Sin cambios en la lista de alertas.")

    doc.add_heading("9. Conclusiones", 1)
    conclusiones = [f"Al {f_fecha(cierre)} el avance de las inversiones del Pliego es {pct(avance(p1))}, "
                    f"con S/ {mill(saldo)} millones por devengar."]
    mayores = sorted(u1, key=lambda f: -(n(f, "pim") - n(f, "devengado")))[:2]
    conclusiones.append("El saldo por devengar se concentra en " + " y ".join(
        f"{f['ue_codigo']} (S/ {mill(n(f, 'pim') - n(f, 'devengado'))} millones)" for f in mayores) + ".")
    sin_dev = [a for a in a1["alertas"] if a["regla"] == "A4"]
    if sin_dev:
        conclusiones.append(f"{len(sin_dev)} proyecto(s) con PIM de al menos S/ 1 millón no registran devengado.")
    vinetas(doc, conclusiones)

    doc.add_heading("10. Recomendaciones", 1)
    rec = [f"Solicitar a las unidades ejecutoras con mayor saldo ({', '.join(f['ue_codigo'] for f in mayores)}) "
           f"su programación de devengado hasta el cierre del año."]
    if sin_dev:
        rec.append("Pedir a los responsables de los proyectos sin devengado el estado de su proceso y la fecha "
                   "estimada del primer devengado, o evaluar la reasignación de los recursos según la normativa vigente.")
    if any(a["regla"] in ("A2", "A5") for a in a1["alertas"]):
        rec.append("Revisar con las áreas competentes las causas de la falta de movimiento en las unidades y proyectos observados.")
    if any("NO resuelta" in a["motivo"] for a in salidas):
        rec.append("Mantener en seguimiento los proyectos que salieron de la lista solo por reducción de su PIM.")
    vinetas(doc, rec)

    doc.add_paragraph("Es todo cuanto informo para su conocimiento y fines.").paragraph_format.space_before = Pt(12)
    doc.add_paragraph("Atentamente,")
    p = doc.add_paragraph(f"_______________________________\n{firma['de']}\n{firma['cargo_de']}")
    p.paragraph_format.space_before = Pt(42)
    p.paragraph_format.keep_together = True

    destino = Path(salida) if salida else SALIDA / f"informe_quincenal_{desde.isoformat()}_a_{hasta.isoformat()}.docx"
    destino.parent.mkdir(parents=True, exist_ok=True)
    if destino.exists():                       # no pisar un informe que ya pudo haber sido editado
        destino = destino.with_name(f"{destino.stem}_{datetime.now():%H%M%S}.docx")
    doc.save(destino)
    return destino, base, cierre


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="Informe quincenal en Word desde la serie diaria.")
    ap.add_argument("--quincena", help="AAAA-MM-1 o AAAA-MM-2")
    ap.add_argument("--desde", type=date.fromisoformat)
    ap.add_argument("--hasta", type=date.fromisoformat)
    a = ap.parse_args()
    if a.quincena:
        desde, hasta = rango_quincena(a.quincena)
    elif a.desde and a.hasta:
        desde, hasta = a.desde, a.hasta
    elif a.desde or a.hasta:
        raise SystemExit("--desde y --hasta van juntos.")
    else:
        desde, hasta = ultima_quincena_completa(datetime.now(PERU).date())
    destino, base, cierre = generar(desde, hasta)
    print(f"Informe creado: {destino}")
    print(f"Periodo {desde:%d/%m/%Y} – {hasta:%d/%m/%Y} · base {f_fecha(base)} · cierre {f_fecha(cierre)}")
    if cierre < hasta.isoformat():
        print(f"[AVISO] No hay consulta del {hasta:%d/%m}: se usó la del {f_fecha(cierre)}.")


if __name__ == "__main__":
    main()
