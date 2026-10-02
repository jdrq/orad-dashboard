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
