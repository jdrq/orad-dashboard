/* ============================================================================
 *  BLOQUE 6D — DEVENGADO MENSUAL, SEDE CENTRAL (GORE Lambayeque, UE 001-855,
 *  "Sólo Proyectos")
 *  ORAD - Gobierno Regional de Lambayeque
 *
 *  Hermano de data/devengado-mensual.js (Bloque 2D, alcance Pliego completo),
 *  pero acotado a la Unidad Ejecutora 001-855 (Sede Central). Mismo esquema
 *  de archivo — filas "N: 'NombreDeMes" sin PIA/PIM — por eso tampoco lo
 *  reconoce mef-reader.js, y por eso se parsea aparte, igual que el otro.
 *
 *  IMPORTANTE — mutua exclusión con el Bloque 2D:
 *   Ambos archivos (gore_devengado_mes.xls y sede_devengado_mes.xls) tienen
 *   la MISMA estructura de filas mensuales. Lo único que los distingue es
 *   que este trae una fila de contexto "Unidad Ejecutora 001-855: ...".
 *   detectar() exige esa fila explícitamente (código 001-855, no cualquier
 *   UE) para no confundirse si algún día se carga el archivo de otra UE.
 *
 *  Regla de negocio (misma que Bloque 2D, definida por Juan, 14/08/2026):
 *   - Enero a Julio son valores INAMOVIBLES — nunca se sobreescriben.
 *   - Agosto en adelante se actualiza solo con cada archivo nuevo cargado.
 *   - El mes calendario en curso se marca "parcial" (barra naranja + nota).
 *
 *  Uso (cableado en index.html igual que Bloque 2D):
 *     DevengadoMensualSC.detectar(wb, XLSX)      -> boolean
 *     DevengadoMensualSC.parsearMeses(wb, XLSX)  -> {1:{...},...,12:{...}}
 *     DevengadoMensualSC.actualizar(archivos)    -> aplica al Bloque 6D y re-renderiza
 * ==========================================================================*/
(function (global) {
  "use strict";

  const MESES_CERRADOS_FIJOS = 7; // Ene(1)..Jul(7) — índices 0..6, jamás se tocan
  const CODIGO_UE_SEDE_CENTRAL = "001-855";

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

  function filaEsMes(row) {
    if (!row || row.length < 8) return false;
    const desc = norm(row[0]);
    return NOMBRES_MES.some(m => desc.includes(m)) && toNum(row[6]) !== null;
  }

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

  // ¿Trae la fila de contexto "Unidad Ejecutora 001-855: ..."? (Sede Central,
  // código exacto — no basta con "cualquier Unidad Ejecutora").
  function esContextoSedeCentral(wb, XLSX) {
    for (const name of wb.SheetNames) {
      const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, blankrows: false, raw: false });
      for (const row of aoa) {
        const desc = norm(row && row[0]);
        if (desc.startsWith("UNIDAD EJECUTORA") && desc.includes(CODIGO_UE_SEDE_CENTRAL)) return true;
      }
    }
    return false;
  }

  function detectar(wb, XLSX) {
    if (!esContextoSedeCentral(wb, XLSX)) return false; // no es Sede Central específicamente
    for (const name of wb.SheetNames) {
      const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, blankrows: false, raw: false });
      if (aoa.some(filaEsMes)) return true;
    }
    return false;
  }

  function actualizar(archivos) {
    if (typeof global.B6M_SC_DEVENGADO_MES === "undefined") return; // Bloque 6D no está en esta página
    const a = [...(archivos || [])].reverse().find(x => x.esDevengadoMensualSC);
    if (!a || !a.mesesDevengadoSC) return;

    const mesEnCurso = new Date().getMonth() + 1;

    Object.keys(a.mesesDevengadoSC).forEach(k => {
      const mes = parseInt(k, 10);
      if (mes <= MESES_CERRADOS_FIJOS) return; // Ene-Jul: inamovibles, nunca se tocan
      const dev = a.mesesDevengadoSC[mes].devengado;
      if (dev == null) return;
      global.B6M_SC_DEVENGADO_MES[mes - 1] = dev;
    });

    global.B6M_SC_IDX_PARCIAL = (mesEnCurso > MESES_CERRADOS_FIJOS) ? (mesEnCurso - 1) : -1;

    if (typeof global.renderBloque6MensualSC === "function") global.renderBloque6MensualSC();
  }

  global.DevengadoMensualSC = { detectar, parsearMeses, actualizar, MESES_CERRADOS_FIJOS, NOMBRES_MES, CODIGO_UE_SEDE_CENTRAL };

})(window);
