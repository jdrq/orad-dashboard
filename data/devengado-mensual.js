/* ============================================================================
 *  BLOQUE 2D — DEVENGADO MENSUAL (GORE Lambayeque, "Sólo Proyectos")
 *  ORAD - Gobierno Regional de Lambayeque
 *
 *  Lee el archivo de Consulta Amigable "por Mes" (el que usa Juan para ver
 *  cuánto se devengó cada mes, ej. gore_devengado_mes.xls). Ese archivo NO
 *  tiene el mismo esquema que los demás .xls del MEF: las filas mensuales
 *  vienen sin PIA/PIM (solo aparecen en la fila TOTAL/Pliego), así que
 *  mef-reader.js las descarta por diseño (exige PIM numérico por fila).
 *  Por eso este módulo las parsea aparte, de forma independiente, sin tocar
 *  el lector universal.
 *
 *  Regla de negocio (definida por Juan, 14/08/2026):
 *   - Enero a Julio son valores INAMOVIBLES. Una vez cerrados, este módulo
 *     JAMÁS los sobreescribe, así el archivo del día los vuelva a traer.
 *   - Agosto en adelante SÍ se actualiza cada vez que se carga un archivo
 *     nuevo — automáticamente, sin tocar código.
 *   - El mes calendario en curso se marca siempre como "parcial" en el
 *     gráfico (barra naranja + nota), los demás como cerrados (vino).
 *
 *  Uso (ya cableado en index.html, ver hooks al final de este archivo):
 *     DevengadoMensual.detectar(wb, XLSX)      -> boolean
 *     DevengadoMensual.parsearMeses(wb, XLSX)  -> {1:{devengado,...}, ..., 12:{...}}
 *     DevengadoMensual.actualizar(archivos)    -> aplica al Bloque 2D y re-renderiza
 * ==========================================================================*/
(function (global) {
  "use strict";

  const MESES_CERRADOS_FIJOS = 7; // Ene(1)..Jul(7) — índices 0..6, jamás se tocan

  const NOMBRES_MES = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO",
                        "AGOSTO","SETIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"];

  function norm(s) {
    return String(s == null ? "" : s)
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .toUpperCase().replace(/\s+/g, " ").trim();
  }

  function toNum(v) {
    if (typeof v === "number") return isFinite(v) ? v : null;
    let s = String(v == null ? "" : v).trim();
    if (s === "" || s === "-" || s === "--") return null;
    const neg = /^\(.*\)$/.test(s);
    s = s.replace(/[^0-9.\-]/g, "");
    if (s === "" || s === "-" || s === ".") return null;
    const n = parseFloat(s);
    if (!isFinite(n)) return null;
    return neg ? -Math.abs(n) : n;
  }

  // ¿Esta fila luce como "N: 'NombreDeMes" con Certificación/Compromiso/
  // Atención/Devengado/Girado numéricos en las columnas 3-7? (columnas 1 y 2,
  // PIA/PIM, vienen vacías en este reporte — por eso mef-reader.js la ignora).
  function filaEsMes(row) {
    if (!row || row.length < 8) return false;
    const desc = norm(row[0]);
    return NOMBRES_MES.some(m => desc.includes(m)) && toNum(row[6]) !== null;
  }

  // Parsea el workbook ya leído por SheetJS y devuelve {1:{...},...,12:{...}},
  // indexado por el número de mes REAL detectado por nombre (no por posición
  // de fila), por si el archivo alguna vez trae menos de 12 filas.
  function parsearMeses(wb, XLSX) {
    const porMes = {};
    for (const name of wb.SheetNames) {
      const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, blankrows: false, raw: false });
      for (const row of aoa) {
        if (!filaEsMes(row)) continue;
        const desc = norm(row[0]);
        const idxMes = NOMBRES_MES.findIndex(m => desc.includes(m));
        if (idxMes < 0) continue;
        porMes[idxMes + 1] = {
          certificacion: toNum(row[3]),
          compromiso:    toNum(row[4]),
          atencion:      toNum(row[5]), // "Atención de Compromiso Mensual" — NO es el devengado
          devengado:     toNum(row[6]),
          girado:        toNum(row[7])
        };
      }
    }
    return porMes;
  }

  // ¿El archivo cargado es este tipo de reporte?
  function detectar(wb, XLSX) {
    for (const name of wb.SheetNames) {
      const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, blankrows: false, raw: false });
      if (aoa.some(filaEsMes)) return true;
    }
    return false;
  }

  // Aplica los datos parseados al Bloque 2D y dispara su re-render.
  // `archivos` es el array completo de archivos cargados (manual o autoload);
  // toma el más reciente marcado como esDevengadoMensual.
  function actualizar(archivos) {
    if (typeof global.B2M_DEVENGADO_MES === "undefined") return; // Bloque 2D no está en esta página
    const a = [...(archivos || [])].reverse().find(x => x.esDevengadoMensual);
    if (!a || !a.mesesDevengado) return;

    const mesEnCurso = new Date().getMonth() + 1; // 1-12, calendario real del navegador

    Object.keys(a.mesesDevengado).forEach(k => {
      const mes = parseInt(k, 10);
      if (mes <= MESES_CERRADOS_FIJOS) return; // Ene-Jul: inamovibles, nunca se tocan
      const dev = a.mesesDevengado[mes].devengado;
      if (dev == null) return;
      global.B2M_DEVENGADO_MES[mes - 1] = dev;
    });

    // El mes parcial es siempre el mes calendario en curso (si cae después de
    // Julio); si el mes en curso ya no está en el archivo o es Ene-Jul, no
    // marca ninguno como parcial.
    global.B2M_IDX_PARCIAL = (mesEnCurso > MESES_CERRADOS_FIJOS) ? (mesEnCurso - 1) : -1;

    if (typeof global.renderBloque2Mensual === "function") global.renderBloque2Mensual();
  }

  global.DevengadoMensual = { detectar, parsearMeses, actualizar, MESES_CERRADOS_FIJOS, NOMBRES_MES };

})(window);
