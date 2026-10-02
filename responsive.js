/* ============================================================================
   PARTE 0 — Vista "página completa" en celulares y tablets (decisión de Jefatura)
   En pantallas físicas de 1024px o menos fija el ancho de la página en 1260px
   (el mismo del escritorio), así el navegador reduce TODO el diseño para que
   entre en la pantalla: se ve igual que en la PC, sin cortes, y se puede
   ampliar con los dedos (pellizcar). Con ancho 1260 los ajustes de tarjetas
   y nombres cortos de abajo no se activan.
   Para ver la versión de tarjetas: agregar ?movil=1 al final de la dirección.
   ========================================================================== */
(function () {
  try {
    var corto = Math.min(screen.width || 9999, screen.height || 9999);
    if (corto <= 1024 && !/[?&]movil=1/.test(location.search)) {
      var m = document.querySelector('meta[name="viewport"]');
      if (m) m.setAttribute("content", "width=1260, initial-scale=" + (Math.min(corto, window.innerWidth || corto) / 1260).toFixed(4) + ", minimum-scale=0.2, maximum-scale=5, user-scalable=yes");
    }
  } catch (e) {}
})();

/* ============================================================================
   responsive.js — Dashboard ORAD · GORE Lambayeque
   Ajusta los gráficos de Chart.js cuando la pantalla es de celular (≤700px).
   - NO modifica datos ni lógica del dashboard: solo márgenes y etiquetas.
   - Debe cargarse DESPUÉS de Chart.js y ANTES del script principal de la página.
   - Si algo falla, el try/catch lo ignora y el gráfico se dibuja como siempre.
   - Para revertir: borrar la línea <script src="responsive.js"> en index.html.
   ========================================================================== */
(function () {
  "use strict";
  if (!window.Chart || !window.matchMedia) return;
  var esCelular = window.matchMedia("(max-width: 700px)");

  Chart.register({
    id: "orad-mobile",
    beforeInit: function (chart) {
      try {
        if (!esCelular.matches) return;
        var o = chart.config.options || {};

        // 1) Menos margen lateral: los gráficos reservan 60px/110px a los lados
        //    para etiquetas de escritorio, y en un celular eso aplasta el área útil.
        var p = o.layout && o.layout.padding;
        if (p && typeof p === "object") {
          if (p.left  != null) p.left  = Math.min(p.left, 10);
          if (p.right != null) p.right = Math.min(p.right, 34);
        }

        var dl = o.plugins && o.plugins.datalabels;
        if (!dl) return;

        // 2) Etiquetas un poco más pequeñas (solo si la fuente es un objeto fijo)
        if (dl.font && typeof dl.font === "object" && dl.font.size) {
          dl.font.size = Math.min(dl.font.size, 11);
        }

        // 3) Gráficos mensuales: montos en millones ("S/ 68.6M") en vez de "S/ 68,587,989"
        var id = chart.canvas && chart.canvas.id;
        if (id === "b2mChart" || id === "b6mChartSC") {
          var clave = id === "b2mChart" ? "B2M_IDX_PARCIAL" : "B6M_SC_IDX_PARCIAL";
          dl.formatter = function (v, ctx) {
            if (v == null) return "";
            return "S/ " + (v / 1e6).toFixed(1) + "M" + (ctx.dataIndex === window[clave] ? "*" : "");
          };
        }
      } catch (e) {
        console.warn("[responsive.js] ajuste omitido:", e.message);
      }
    }
  });
})();

/* ============================================================================
   PARTE 2 — Nombres cortos en rankings y tablas como tarjetas (celular/tablet)
   - Nombres: "GOBIERNO REGIONAL DEL DEPARTAMENTO DE JUNIN" -> "GORE JUNIN",
     solo en pantallas <=1100px. En escritorio se restaura el texto original.
   - Tarjetas: en pantallas <=840px agrega clases/etiquetas a las tablas para que
     responsive.css las muestre como tarjetas con todas sus columnas.
   - En escritorio (>1100px) este código no modifica nada de la página.
   ========================================================================== */
(function () {
  "use strict";
  if (!window.matchMedia || !window.MutationObserver) return;
  var mqCard  = window.matchMedia("(max-width: 840px)");
  var mqCorto = window.matchMedia("(max-width: 1100px)");

  // ── Nombres cortos ──────────────────────────────────────────────────────
  function acortar(n) {
    var s = String(n || "").replace(/\s+/g, " ").trim();
    var m = /^MUNICIPALIDAD METROPOLITANA DE\s+(.+)$/i.exec(s);
    if (m) return "MUN. METROP. " + m[1];
    m = /^GOBIERNO REGIONAL\s+(?:DE LA PROVINCIA CONSTITUCIONAL DEL|DEL DEPARTAMENTO DE|DEL DEPARTAMENTO|DEL|DE)\s+(.+)$/i.exec(s);
    if (m) return "GORE " + m[1];
    return s;
  }
  var RANKINGS = ["goreRowsA", "goreRowsB", "goreRowsC"];
  function aplicarNombres() {
    var cortar = mqCorto.matches;
    RANKINGS.forEach(function (id) {
      var tb = document.getElementById(id);
      if (!tb) return;
      Array.prototype.forEach.call(tb.rows, function (tr) {
        var td = tr.cells[1];
        if (!td || tr.cells.length < 3) return;          // fila de mensaje: no tocar
        if (cortar) {
          if (td.getAttribute("data-full") == null) td.setAttribute("data-full", td.textContent);
          var full = td.getAttribute("data-full");
          var corto = acortar(full);
          if (td.textContent !== corto) td.textContent = corto;
          td.title = full;
        } else if (td.getAttribute("data-full") != null) {
          td.textContent = td.getAttribute("data-full");
          td.removeAttribute("data-full");
          td.removeAttribute("title");
        }
      });
    });
  }

  // ── Tarjetas ────────────────────────────────────────────────────────────
  var RX_RANK  = /^(N[°º]?|#)$/i;
  var RX_TITLE = /PLIEGO|UNIDAD EJECUTORA|PROYECTO|DESCRIPCI|FUNCI[ÓO]N|MUNICIPALIDAD|GEN[ÉE]RICA/i;
  var RX_AV    = /^AVANCE/i;
  var RX_BAR   = /GR[ÁA]FICO/i;

  function etiquetasDe(table) {
    var head = table.tHead && table.tHead.rows[0];
    if (!head) return null;
    return Array.prototype.filter.call(head.cells, function (th) { return th.style.display !== "none"; })
      .map(function (th) { return th.textContent.replace(/\s+/g, " ").trim(); });
  }

  function procesarTabla(table) {
    var labels = etiquetasDe(table);
    if (!labels) return;
    Array.prototype.forEach.call(table.querySelectorAll("tbody tr, tfoot tr"), function (tr) {
      if (tr.getAttribute("data-oc") === "1") return;       // ya procesada
      var col = 0, hasR = false, hasA = false, hasTitle = false;
      Array.prototype.forEach.call(tr.cells, function (td) {
        var span = td.colSpan || 1;
        var label = labels[col] || "";
        td.setAttribute("data-label", label);
        if (span > 1 && !hasTitle) { td.classList.add("oc-title"); hasTitle = true; }          // "TOTAL" o mensaje
        else if (RX_RANK.test(label)) { td.classList.add("oc-rank"); hasR = true; }
        else if (!hasTitle && RX_TITLE.test(label)) { td.classList.add("oc-title"); hasTitle = true; }
        else if (RX_AV.test(label)) { td.classList.add("oc-avance"); hasA = true; }
        else if (RX_BAR.test(label)) { td.classList.add("oc-bar"); }
        col += span;
      });
      if (hasR) tr.classList.add("oc-r");
      if (hasA) tr.classList.add("oc-a");
      tr.setAttribute("data-oc", "1");
    });
  }

  function procesarTodo() {
    try {
      aplicarNombres();
      var tarjetas = mqCard.matches;
      Array.prototype.forEach.call(document.querySelectorAll("table.pdf-table"), function (t) {
        if (tarjetas) { procesarTabla(t); t.classList.add("ocard"); }
        else t.classList.remove("ocard");
      });
    } catch (e) {
      console.warn("[responsive.js] tarjetas omitidas:", e.message);
    }
  }

  var pendiente = 0;
  function programar() {
    if (pendiente) return;
    pendiente = setTimeout(function () { pendiente = 0; procesarTodo(); }, 60);
  }
  function iniciar() {
    procesarTodo();
    new MutationObserver(programar).observe(document.body, { childList: true, subtree: true });
    [mqCard, mqCorto].forEach(function (mq) {
      if (mq.addEventListener) mq.addEventListener("change", procesarTodo);
      else if (mq.addListener) mq.addListener(procesarTodo);
    });
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", iniciar);
  else iniciar();
})();
