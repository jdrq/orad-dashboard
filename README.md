# Dashboard BI — ORAD · Gobierno Regional de Lambayeque

Dashboard web de ejecución presupuestal de inversiones del **Pliego 452 - GORE Lambayeque**, construido en un único `index.html` (HTML + JavaScript + Chart.js + SheetJS), publicado vía GitHub Pages. Funciona de forma idéntica abriéndolo por doble clic (`file://`) o desde la web.

**Elaborado por:** ORPMI — Oficina Regional de Programación Multianual de Inversiones
**Fuente de datos:** Consulta Amigable MEF · [apps5.mineco.gob.pe](https://apps5.mineco.gob.pe/transparencia/Navegador/default.aspx)
**Dashboard en vivo:** https://jdrq.github.io/orad-dashboard/

---

## Contenido del Dashboard

| Bloque | Descripción |
|--------|-------------|
| Bloque 1 | KPIs financieros del Pliego (PIM, Certificación, Compromiso, Devengado) |
| Bloque 2 | Ranking nacional de los 26 Gobiernos Regionales — Por Devengado %, Por Devengado Monto, Por Certificación % (año en curso en vivo + selector histórico 2022–2025) |
| Bloque 2D | Devengado Mensual — Gobierno Regional de Lambayeque (comparativa mes a mes, Enero–Diciembre) |
| Bloque 3 | Ejecución por Unidad Ejecutora (tabla, gráfico, Top Proyectos) |
| Bloque 3R | Proyectos con Ejecución en Riesgo — mayor PIM asignado sin devengar aún (Sede Central) |
| Bloque EMR | Reducción de Vulnerabilidad y Atención de Emergencias — Categoría Presupuestal 0068, Pliego 452 |
| Bloque EMU | Reducción de Vulnerabilidad y Atención de Emergencias — Municipalidades (38 municipalidades provinciales y distritales de Lambayeque, Categoría 0068) |
| Bloque 6 | Ejecución por Rubro / Fuente de Financiamiento (Sede Central, PEOT/Agricultura/Transporte/GERESA, Hospitales, consolidado Pliego), con comparación histórica 2022–2026 |
| Bloque 6D | Devengado Mensual — Sede Central (comparativa mes a mes, Enero–Diciembre) |
| Bloque 7 | Ejecución por Función a nivel Pliego |
| Bloque 8 | Comparativo Histórico Enero–Diciembre — solo Gobierno Regional de Lambayeque (Devengado y PIM, 2016–2025 + año en curso en vivo) |
| Bloque 9 | Proyección acumulada al 31/12 del año, con vista "Ver Primer Semestre" congelada al 30/06 |

> **Corregido 31/08/2026 — no confundir Bloque 2 con Bloque 8:** el ranking nacional de los 26 GORES (alimentado por `data/historico_progresivo.json` / `scripts/actualizar_hist_gores.py`) es el **Bloque 2** (`id="b2secDevP"` / `"b2secDevS"` / `"b2secCert"` en `index.html`), ubicado cerca del inicio de la página. El **Bloque 8** real (`id="b8secHistorico"`) es un bloque totalmente distinto: la tendencia histórica de Devengado/PIM **solo de Lambayeque**, sin comparación con otros GORES. Versiones anteriores de este README y de los scripts rotulaban por error el ranking nacional como "Bloque 8" — si encuentras ese rótulo en un comentario de código o en un commit viejo, es el error heredado, no una referencia válida.

---

## Arquitectura

El dashboard **no depende de ningún backend en producción**. El parseo de los `.xls` del día (que en realidad son HTML disfrazado, exportado por Consulta Amigable) ocurre **en vivo, en el navegador**, vía SheetJS embebido en `index.html`. Los datos históricos (años cerrados 2022–2025) se pre-procesan una sola vez con Python y se sirven como JSON estático. Dos módulos JS aparte (`data/devengado-mensual.js` y `data/devengado-mensual-sc.js`) alimentan los bloques de Devengado Mensual (2D y 6D) con lógica propia de "mes cerrado vs. mes parcial en curso", calculada en vivo contra la fecha real del navegador.

```
                    ┌─ Descarga diaria (Fase 2a — Playwright) ─┐
                    │  python scripts/descargar_xls_mef.py     │
                    │  → 17 archivos en xls/ (raíz)             │
                    └───────────────────────────────────────────┘
                                      ↓
                    index.html los lee y parsea EN EL NAVEGADOR (SheetJS)
                    data/devengado-mensual.js / -sc.js resuelven mes cerrado vs. parcial
                                      ↓
                    git add / commit / push (manual, vía VS Code / PowerShell)
                                      ↓
                    GitHub Pages publica automáticamente
                                      ↓
                    Tu jefe abre el enlace y ve los datos del día

                    ┌─ Devengado Mensual (Bloques 2D y 6D) ────────────────┐
                    │  gore_devengado_mes.xls / sede_devengado_mes.xls     │
                    │  → valores MENSUALES NO ACUMULADOS (hay que sumar    │
                    │    progresivamente para obtener el acumulado)        │
                    │            ↓                                        │
                    │  Meses cerrados: hardcodeados en index.html          │
                    │  (B2M_DEVENGADO_MES / B6M_SC_DEVENGADO_MES) — se     │
                    │  actualizan a mano al cierre de cada mes, con el     │
                    │  dato OFICIAL (no el parcial de mitad de mes: el MEF │
                    │  reclasifica meses "cerrados" días después, ya       │
                    │  pasó con junio y con agosto — ver Hoja de Ruta)     │
                    │  Mes en curso: se rellena en vivo desde el archivo   │
                    │  del día, sin tocar código                           │
                    └────────────────────────────────────────────────────┘
                                      ↓
                    ┌─ Datos históricos (una vez, al cerrar cada período) ─┐
                    │  xls/historico/*.xls        (Bloque 2 — ranking      │
                    │                               nacional 26 GORES)     │
                    │  xls/historico_rubro/*.xls  (Bloque 6 — Rubro        │
                    │                               Sede Central)          │
                    │            ↓                                        │
                    │  scripts/convertir_semestral.py         (Ene-Jun)   │
                    │  scripts/actualizar_hist_gores.py       (Ene-Jul+) │
                    │  scripts/actualizar_rb_hist_sc.py        (Ene-Jul+) │
                    │            ↓                                        │
                    │  data/*.json (historico_semestral, historico_progresivo,│
                    │               rb_hist_sc_progresivo, semestre1_2026)    │
                    └────────────────────────────────────────────────────┘
                                      ↓
                    index.html los carga con fetch() (Bloques 2 y 6)
```

> **Nota histórica:** versiones anteriores de este proyecto usaban un script Python distinto (`convertir_xls_a_json.py`) para pre-procesar todo el dashboard. Ese enfoque quedó **obsoleto** y fue retirado — hoy Python solo procesa los históricos cerrados (2022–2025), nunca los datos del día. No es necesario tener Python instalado para operar el dashboard día a día; sí es necesario para correr `scripts/descargar_xls_mef.py` (descarga automatizada) y los demás scripts de `scripts/`.

---

## Estructura del Proyecto

```
orad-dashboard/
├── index.html                       ← Dashboard (un solo archivo, autocontenido)
├── .gitignore
├── data/
│   ├── devengado-mensual.js         ← Bloque 2D: resuelve mes cerrado vs. parcial en vivo
│   ├── devengado-mensual-sc.js      ← Bloque 6D: ídem, para Sede Central
│   ├── semestre1_2026.json          ← Snapshot congelado al 30/06/2026 ("Ver Primer Semestre")
│   ├── historico_progresivo.json    ← Comparación histórica nacional Ene-Jul+ (2022-2025) — Bloque 2
│   ├── historico_semestral.json     ← Comparación histórica nacional Ene-Jun (2022-2025) — Bloque 2
│   └── rb_hist_sc_progresivo.json   ← Comparación histórica Rubro Sede Central Ene-Jul+ (2022-2025) — Bloque 6
├── scripts/
│   ├── descargar_xls_mef.py         ← Fase 2a: descarga automatizada (Playwright) de los 17 XLS diarios
│   ├── actualizar_rb_hist_sc.py     ← Actualiza dev_t1/dev_t2/dev_sem/dev_ago/dev_set en rb_hist_sc_progresivo.json (Bloque 6)
│   ├── actualizar_hist_gores.py     ← Actualiza historico_progresivo.json — ranking nacional GORES (Bloque 2)
│   └── convertir_semestral.py       ← Genera historico_semestral.json (ranking GORES) desde xls/historico/
└── xls/
    ├── (17 archivos del día — ver tabla abajo, se suben a Git)
    ├── historico/                   ← Manual. Fuente para scripts/convertir_semestral.py y
    │                                   scripts/actualizar_hist_gores.py (Bloque 2). SÍ se sube a Git.
    ├── historico_rubro/             ← Manual. Fuente para scripts/actualizar_rb_hist_sc.py (Bloque 6). SÍ se sube a Git.
    └── _respaldo_anterior/          ← Automática (la crea scripts/descargar_xls_mef.py en cada corrida).
                                        EXCLUIDA de Git vía .gitignore — nunca se sube.
```

> **Regla de oro para no confundir las carpetas de `xls/`:** si la carpeta la llenas **tú manualmente** desde Consulta Amigable, se sube a Git. Si la carpeta la **crea el script solo**, es un subproducto transitorio y va al `.gitignore`.

---

## Cómo Actualizar Diariamente

### Paso 1 — Descargar los 17 XLS (Fase 2a — automatizado)

```bash
python scripts\descargar_xls_mef.py
```

El script abre una ventana de navegador visible (Playwright), hace el drill-down completo para las 17 combinaciones de UE/agrupación, valida cada descarga y guarda un respaldo de la versión anterior en `xls/_respaldo_anterior/` antes de sobrescribir. Al terminar, la consola debe mostrar **`17/17 archivos OK`** — ese mensaje es tu compuerta de calidad antes de continuar. Si falta alguno, revisar el `error_<archivo>.png` generado y reintentar antes de seguir al paso 2.

> **Modo manual (respaldo, por si el script falla o el sitio del MEF cambia):** exportar los archivos a mano desde Consulta Amigable con el drill-down correspondiente a cada uno (ver tabla siguiente), y renombrar cada uno según corresponda. **Ojo con los archivos de Rubro:** al exportarlos hay que seleccionar explícitamente la columna de agrupación antes de bajar el archivo — si trae solo una fila "TOTAL", no se seleccionó la agrupación. **Ojo también con "Guardar como" desde Excel:** genera un archivo *frameset* (un cascarón que apunta a una subcarpeta `_archivos/` con la data real) en vez del HTML de una sola pieza que necesita el dashboard — usar siempre el botón de exportar nativo de la propia página de Consulta Amigable.

| # | Archivo a guardar como | Nivel de drill / Agrupación | Bloque que alimenta |
|---|------------------------|------------------------------|----------------------|
| 1 | `rubro_sede_central.xls` | UE 001-855 (Sede Central) → por Rubro | 6 |
| 2 | `rubro_peot.xls` | UE 002-1133 (Proy. Esp. Olmos Tinajones) → por Rubro | 6 |
| 3 | `rubro_agricultura.xls` | UE 100-856 (Agricultura) → por Rubro | 6 |
| 4 | `rubro_transportes.xls` | UE 200-857 (Transportes) → por Rubro | 6 |
| 5 | `rubro_salud.xls` | UE 400-860 (Salud) → por Rubro | 6 |
| 6 | `rubro_h_mercedes.xls` | UE 401-1001 (Hospital Las Mercedes) → por Rubro | 6 |
| 7 | `rubro_h_belen.xls` | UE 402-1002 (Hospital Belén) → por Rubro | 6 |
| 8 | `rubro_h_regional.xls` | UE 403-1422 (Hospital Regional) → por Rubro | 6 |
| 9 | `rubro_pliego.xls` | Pliego 452 consolidado (sin UE) → por Rubro | 6 |
| 10 | `ue_pliego.xls` | Pliego 452 consolidado → por Unidad Ejecutora | 1, 3 |
| 11 | `nacional_gores.xls` | Sector 99 (sin Pliego específico) → por Pliego (26 GOREs) | 2 |
| 12 | `funciones_pliego.xls` | Pliego 452 consolidado → por Función | 7 |
| 13 | `proyectos_sede_central.xls` | UE 001-855 (Sede Central) → por Proyecto | 3R |
| 14 | `gore_devengado_mes.xls` | Pliego 452, Sólo Proyectos → por Mes | 2D, 9 |
| 15 | `sede_devengado_mes.xls` | UE 001-855 (Sede Central), Sólo Proyectos → por Mes | 6D |
| 16 | `proyecto_emergencia.xls` | Pliego 452, Categoría 0068, "Actividades y Proyectos" → por Genérica de Gasto | EMR |
| 17 | `munis_emergencia.xls` | Sector 99 → Gobiernos Locales, Categoría 0068, Departamento Lambayeque, "Actividades y Proyectos" → por Municipalidad | EMU |

### Paso 2 — Subir a GitHub (vía VS Code o PowerShell)

```powershell
git status                 # revisar qué cambió antes de subir nada — el 17/17 OK ya fue tu compuerta de calidad
git add xls\*.xls          # solo los XLS de primer nivel, nunca "git add ." a ciegas
git commit -m "data: actualización diaria XLS - DD/MM/YYYY"
git push origin main
```

> `_respaldo_anterior/` queda excluida automáticamente por el `.gitignore` — no hace falta revisar manualmente que no se cuele.

### Paso 3 — Verificar

Esperar 1-2 minutos y abrir `https://jdrq.github.io/orad-dashboard/` (Ctrl+Shift+R para forzar recarga sin caché) y confirmar que los datos del día se reflejan.

---

## Archivos Históricos (años cerrados 2022–2025)

Estos **no** forman parte de la actualización diaria — se exportan una sola vez por período (o se corrigen puntualmente, como ocurrió al extender a Julio, Agosto y Setiembre) y viven en dos subcarpetas separadas por el bloque del dashboard que alimentan:

### `xls/historico/` → Bloque 2 (ranking nacional GORES)

Por año (2022–2025): `1T_{año}.xls`, `2T_{año}.xls`, `julio_{año}.xls`, `agosto_{año}.xls`, `setiembre_{año}.xls`, `anual_{año}_gores.xls`. Se procesan en dos etapas, con dos scripts distintos, según qué JSON alimentan:

- **`scripts/convertir_semestral.py`** → genera `data/historico_semestral.json` (Ene-Jun, T1+T2). Este es el corte "Primer Semestre" que se congela para la vista `Ver Primer Semestre`.
- **`scripts/actualizar_hist_gores.py`** → genera `data/historico_progresivo.json` (Ene-Jul+, T1+T2+Julio+Agosto+Setiembre...). Usa la misma filosofía de acumulación progresiva y validación cruzada que `scripts/actualizar_rb_hist_sc.py` (ver más abajo), con lista `PERIODOS` configurable para extender a los meses siguientes sin reescribir lógica. **Recalcula desde cero en cada corrida**, así que necesita TODOS los archivos de los períodos activos presentes a la vez en `xls/historico/`, no solo el del mes nuevo.

> **Metodología de promedios nacionales (corregido 07/07/2026):** `prom_dev_pct` y `prom_cert_pct` se calculan como **promedio ponderado por PIM** — `Σ(devengado de los 26 GOREs) / Σ(PIM de los 26 GOREs) × 100` — no como promedio simple de los 26 porcentajes individuales. Es la misma metodología que usa el propio MEF para el avance nacional agregado. Un promedio simple da un resultado distinto (confirmado con auditoría matemática: hasta ~3 puntos porcentuales de diferencia), porque le da el mismo peso a un GORE grande como Arequipa que a uno pequeño como Tumbes.

### `xls/historico_rubro/` → Bloque 6 (Rubro — Sede Central)

Por año (2022–2025): `1T_RUBRO_{año}.xls`, `2T_RUBRO_{año}.xls`, `JULIO_RUBRO_{año}.xls`, `AGOSTO_RUBRO_{año}.xls`, `SETIEMBRE_RUBRO_{año}.xls`. Se procesan con `scripts/actualizar_rb_hist_sc.py`, que actualiza `data/rb_hist_sc_progresivo.json`.

> **Requisito de exportación (crítico):** al pedir estos archivos en Consulta Amigable, seleccionar explícitamente la columna de agrupación ("Rubro" o el nivel correspondiente) **antes** de exportar. Si el archivo trae solo una fila "TOTAL" en la tabla de detalle, significa que no se seleccionó la agrupación — no hay forma de recuperar el desglose después, hay que reexportar desde cero.

> **Metodología de acumulación (importante, verificado empíricamente):** los filtros "Trimestre I", "Trimestre II" y "Mes N" de Consulta Amigable devuelven valores **por período (incrementales)**, no acumulados. Es decir, `2T_RUBRO_{año}.xls` (o `2T_{año}.xls`) trae solo Abr-Jun, no Ene-Jun; `AGOSTO_RUBRO_{año}.xls` trae solo Agosto, no Ene-Ago. Para obtener acumulados hay que sumar progresivamente:
> - `dev_t1` (acumulado a marzo) = valor de `1T`
> - `dev_t2` (acumulado a junio) = `1T` + `2T`
> - `dev_ago` (acumulado a agosto) = `1T` + `2T` + `Julio` + `Agosto`
> - `dev_set` (acumulado a setiembre) = lo anterior + `Setiembre`
>
> Tanto `scripts/actualizar_rb_hist_sc.py` como `scripts/actualizar_hist_gores.py` aplican esta suma automáticamente y validan el resultado contra el dato ya existente en el JSON (tolerancia 0.5%, o por monotonicidad para meses sin benchmark previo) **antes** de sobrescribir — si algún año no pasa la validación, ese año se conserva sin cambios (nunca se escribe un dato dudoso) y verás `⚠️ ... BLOQUEADO` en la consola. No lo ignores: revisa manualmente antes de reintentar.
>
> **Extender a meses futuros:** ambos scripts tienen una lista `PERIODOS` configurable al inicio del archivo. Agregar un mes nuevo es: (1) re-exportar el archivo del mes con la agrupación correcta seleccionada, (2) descomentar/agregar la línea correspondiente en `PERIODOS`, (3) correr el script. No requiere reescribir lógica.

### Devengado Mensual (Bloques 2D y 6D) — un tercer mecanismo de cierre de mes, manual y separado

A diferencia de los dos pipelines anteriores (JSON generado por script), los arrays `B9_REAL_BASE`, `B2M_DEVENGADO_MES` y `B6M_SC_DEVENGADO_MES` viven **hardcodeados directamente en `index.html`** y se cierran a mano cada fin de mes:

1. Descargar `gore_devengado_mes.xls` (Pliego) y/o `sede_devengado_mes.xls` (Sede Central) — trae el Devengado **mensual no acumulado** de cada mes del año, en una tabla tipo "1: Enero", "2: Febrero", etc.
2. Sumar progresivamente los meses no acumulados hasta el mes que se quiere cerrar, y cruzar el resultado contra el corte oficial que trae el mismo archivo en la fila "Pliego 452" / "Unidad Ejecutora 001-855" (deben coincidir, con 1 sol de tolerancia por redondeo).
3. Reemplazar el valor del mes recién cerrado en los tres arrays (`B9_REAL_BASE`, `B2M_DEVENGADO_MES`, y si aplica `B6M_SC_DEVENGADO_MES`), y avanzar el índice de "mes actual" (`B9_MES_ACTUAL_IDX`) al mes siguiente.

> **⚠️ El MEF reclasifica meses "cerrados" varios días después del cierre.** Ya pasó dos veces en 2026: junio se corrigió +S/2.12M una semana después de darlo por cerrado, y agosto se corrigió +S/3.9M entre el 31/08 y el 02/09. **No dar un mes por definitivamente fijo el mismo día en que termina** — esperar unos días hábiles a que el corte se estabilice en Consulta Amigable, y estar dispuesto a corregirlo una vez más si al re-exportar el archivo el monto cambió.

---

## Modo Manual (sin conexión)

El dashboard funciona igual abriendo `index.html` con doble clic y con los XLS en la carpeta `xls/` local — no requiere servidor ni conexión a internet. El fetch de los JSON de `data/` (históricos y snapshot semestral) es la única parte que depende de estar servido vía GitHub Pages o similar; si falta, el dashboard sigue funcionando con los datos en vivo de los 17 XLS.

---

## Hoja de Ruta

- **Fase 2a (actual):** descarga automatizada supervisada, con ventana de navegador visible y checkpoints de calidad (17/17 archivos, drill-down verificado, respaldo automático).
- **Fase 2b (futura, diferida deliberadamente):** automatización desatendida vía Windows Task Scheduler — solo después de varias semanas adicionales de Fase 2a estable, sin fallos de descarga ni necesidad de intervención manual.
- **Cerrado (07/07/2026):** `historico_progresivo.json` dejó de armarse manualmente — ahora lo genera `scripts/actualizar_hist_gores.py`, con la misma filosofía de validación cruzada y acumulación progresiva que Rubro. De paso se detectó y corrigió un error metodológico real: los promedios nacionales (`prom_dev_pct`/`prom_cert_pct`) usaban promedio simple en vez de ponderado por PIM.
- **Cerrado (18/08/2026):** se agregaron los Bloques EMR y EMU (Categoría Presupuestal 0068 — Reducción de Vulnerabilidad y Atención de Emergencias), llevando el total de archivos diarios de 15 a 17. Se corrigió además un bug de raíz de codificación (usar `TextDecoder` con windows-1252 en vez de `codepage:1252` de SheetJS) que afectaba Ñ/tildes en cualquier archivo.
- **Cerrado (31/08–02/09/2026):** cierre de Agosto y apertura de Setiembre en los tres pipelines (Bloque 2, Bloque 6, y Devengado Mensual 2D/6D/9). Agosto tuvo que corregirse dos veces por reclasificación tardía del MEF (ver nota en la sección de Devengado Mensual arriba). Se corrigió también un bug real en el render de Bloque 6 histórico: `index.html` leía el campo estático `av_pct`/`label` en vez de recalcular en vivo desde `dev_set`/`dev_ago`/PIM, lo que habría dejado setiembre invisible en el dashboard aunque el JSON ya tuviera el dato correcto — se agregó el helper `rbHistCorte()` para que esto no se repita cada mes.
- **Cerrado (31/08/2026):** corregido el rótulo "Bloque 8" mal aplicado al ranking nacional GORES en `scripts/actualizar_hist_gores.py` y en este README — el bloque correcto es **Bloque 2** (ver aviso al inicio de este documento).
- **Próximo hito (cuando cierre Setiembre 2026):** extender `PERIODOS` en `scripts/actualizar_rb_hist_sc.py` y `scripts/actualizar_hist_gores.py` agregando Octubre, y repetir el cierre manual de Bloques 2D/6D/9 con el mismo cuidado por la reclasificación tardía del MEF.
- **Pendiente no urgente:** migrar los datos hardcodeados en `index.html` (constante `RB_HIST_SC`, bloque "Ver Primer Semestre") a un JSON externo en `data/`, consolidando toda la data histórica en un solo formato — hoy coexisten dos fuentes (HTML hardcodeado y JSON externo) para el mismo bloque, lo cual es deuda técnica aceptable a corto plazo pero no ideal a largo plazo.
