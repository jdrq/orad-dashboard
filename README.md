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
| Bloque 2 | Ranking nacional de los 26 Gobiernos Regionales |
| Bloque 3 | Ejecución por Unidad Ejecutora |
| Bloque 6 | Ejecución por Rubro / Fuente de Financiamiento (por UE y consolidado Pliego), con comparación histórica 2022–2026 |
| Bloque 7 | Ejecución por Función a nivel Pliego |
| Bloque 8 | Comparativo histórico nacional GORES (2022–2026, extendido mes a mes) |
| Bloque 9 | Proyección acumulada al 31/12 del año, con vista "Ver Primer Semestre" congelada al 30/06 |

---

## Arquitectura

El dashboard **no depende de ningún backend en producción**. El parseo de los `.xls` del día (que en realidad son HTML disfrazado, exportado por Consulta Amigable) ocurre **en vivo, en el navegador**, vía SheetJS embebido en `index.html`. Los datos históricos (años cerrados 2022–2025) se pre-procesan una sola vez con Python y se sirven como JSON estático.

```
                    ┌─ Descarga diaria (Fase 2a — Playwright) ─┐
                    │  python descargar_xls_mef.py             │
                    │  → 13 archivos en xls/ (raíz)             │
                    └───────────────────────────────────────────┘
                                      ↓
                    index.html los lee y parsea EN EL NAVEGADOR (SheetJS)
                                      ↓
                    git add / commit / push (manual, vía VS Code)
                                      ↓
                    GitHub Pages publica automáticamente
                                      ↓
                    Tu jefe abre el enlace y ve los datos del día


                    ┌─ Datos históricos (una vez, al cerrar cada período) ─┐
                    │  xls/historico/*.xls        (ranking GORES)          │
                    │  xls/historico_rubro/*.xls  (Rubro Sede Central)     │
                    │            ↓                                        │
                    │  scripts/convertir_semestral.py                     │
                    │  actualizar_rb_hist_sc.py                           │
                    │            ↓                                        │
                    │  data/*.json (historico_semestral, historico_enejul,│
                    │               rb_hist_sc_enejul, semestre1_2026)    │
                    └────────────────────────────────────────────────────┘
                                      ↓
                    index.html los carga con fetch() (Bloques 6 y 8)
```

> **Nota histórica:** versiones anteriores de este proyecto usaban un script Python distinto (`convertir_xls_a_json.py`) para pre-procesar todo el dashboard. Ese enfoque quedó **obsoleto** y fue retirado — hoy Python solo procesa los históricos cerrados (2022–2025), nunca los datos del día. No es necesario tener Python instalado para operar el dashboard día a día; sí es necesario para correr `descargar_xls_mef.py` (descarga automatizada) y los scripts de `scripts/`.

---

## Estructura del Proyecto

```
orad-dashboard/
├── index.html                       ← Dashboard (un solo archivo, autocontenido)
├── descargar_xls_mef.py             ← Fase 2a: descarga automatizada (Playwright) de los 13 XLS diarios
├── actualizar_rb_hist_sc.py         ← Actualiza dev_t1/dev_t2 en rb_hist_sc_enejul.json
├── actualizar.bat                   ← Script de referencia (validación 13/13 + push)
├── .gitignore
├── data/
│   ├── semestre1_2026.json          ← Snapshot congelado al 30/06/2026 ("Ver Primer Semestre")
│   ├── historico_enejul.json        ← Comparación histórica nacional Ene-Jul (2022-2025)
│   ├── historico_semestral.json     ← Comparación histórica nacional Ene-Jun (2022-2025)
│   └── rb_hist_sc_enejul.json       ← Comparación histórica Rubro Sede Central Ene-Jul (2022-2025)
├── scripts/
│   └── convertir_semestral.py       ← Genera historico_semestral.json (ranking GORES) desde xls/historico/
└── xls/
    ├── (13 archivos del día — ver tabla abajo, se suben a Git)
    ├── historico/                   ← Manual. Fuente para convertir_semestral.py (ranking GORES). SÍ se sube a Git.
    ├── historico_rubro/             ← Manual. Fuente para actualizar_rb_hist_sc.py (Rubro Sede Central). SÍ se sube a Git.
    └── _respaldo_anterior/          ← Automática (la crea descargar_xls_mef.py en cada corrida).
                                        EXCLUIDA de Git vía .gitignore — nunca se sube.
```

> **Regla de oro para no confundir las carpetas de `xls/`:** si la carpeta la llenas **tú manualmente** desde Consulta Amigable, se sube a Git. Si la carpeta la **crea el script solo**, es un subproducto transitorio y va al `.gitignore`.

---

## Cómo Actualizar Diariamente

### Paso 1 — Descargar los 13 XLS (Fase 2a — automatizado)

```bash
python descargar_xls_mef.py
```

El script abre una ventana de navegador visible (Playwright), hace el drill-down completo para las 13 combinaciones de UE/agrupación, valida cada descarga y guarda un respaldo de la versión anterior en `xls/_respaldo_anterior/` antes de sobrescribir. Al terminar, la consola debe mostrar **`13/13 archivos OK`** — ese mensaje es tu compuerta de calidad antes de continuar. Si falta alguno, revisar el `error_<archivo>.png` generado y reintentar antes de seguir al paso 2.

> **Modo manual (respaldo, por si el script falla o el sitio del MEF cambia):** exportar los 13 archivos a mano desde Consulta Amigable, con Año = 2026, Actividades/Proyectos = "Sólo Proyectos", Nivel de Gobierno = R → Sector 99 → Pliego 452, y renombrar cada uno según la tabla siguiente.

| # | Archivo a guardar como | Nivel de drill / Agrupación |
|---|------------------------|------------------------------|
| 1 | `rubro_sede_central.xls` | UE 001-855 (Sede Central) → por Rubro |
| 2 | `rubro_peot.xls` | UE 002-1133 (Proy. Esp. Olmos Tinajones) → por Rubro |
| 3 | `rubro_agricultura.xls` | UE 100-856 (Agricultura) → por Rubro |
| 4 | `rubro_transportes.xls` | UE 200-857 (Transportes) → por Rubro |
| 5 | `rubro_salud.xls` | UE 400-860 (Salud) → por Rubro |
| 6 | `rubro_h_mercedes.xls` | UE 401-1001 (Hospital Las Mercedes) → por Rubro |
| 7 | `rubro_h_belen.xls` | UE 402-1002 (Hospital Belén) → por Rubro |
| 8 | `rubro_h_regional.xls` | UE 403-1422 (Hospital Regional) → por Rubro |
| 9 | `rubro_pliego.xls` | Pliego 452 consolidado (sin UE) → por Rubro |
| 10 | `ue_pliego.xls` | Pliego 452 consolidado → por Unidad Ejecutora |
| 11 | `nacional_gores.xls` | Sector 99 (sin Pliego específico) → por Pliego (26 GOREs) |
| 12 | `funciones_pliego.xls` | Pliego 452 consolidado → por Función |
| 13 | `proyectos_sede_central.xls` | UE 001-855 (Sede Central) → por Proyecto |

### Paso 2 — Subir a GitHub (vía VS Code)

```bash
git status                 # revisar qué cambió antes de subir nada — el 13/13 OK ya fue tu compuerta de calidad
git add xls\*.xls          # solo los XLS de primer nivel, nunca "git add ." a ciegas
git commit -m "data: actualización diaria XLS - DD/MM/YYYY"
git push origin main
```

> `_respaldo_anterior/` queda excluida automáticamente por el `.gitignore` — no hace falta revisar manualmente que no se cuele.

### Paso 3 — Verificar

Esperar 1-2 minutos y abrir `https://jdrq.github.io/orad-dashboard/` (Ctrl+Shift+R para forzar recarga sin caché) y confirmar que los datos del día se reflejan.

---

## Archivos Históricos (años cerrados 2022–2025)

Estos **no** forman parte de la actualización diaria — se exportan una sola vez (o se corrigen puntualmente, como ocurrió con la extensión a julio) y viven en dos subcarpetas separadas por el bloque del dashboard que alimentan:

### `xls/historico/` → Bloque 8 (ranking nacional GORES)

Por año (2022–2025): `1T_{año}.xls`, `2T_{año}.xls`, `anual_{año}_gores.xls`, `julio_{año}.xls`. Se procesan con `scripts/convertir_semestral.py`, que lee la **Tabla 3** de cada archivo (ranking de 26 GOREs) y genera `data/historico_semestral.json`.

### `xls/historico_rubro/` → Bloque 6 (Rubro — Sede Central)

Por año (2022–2025): `1T_RUBRO_{año}.xls`, `2T_RUBRO_{año}.xls`, `JULIO_RUBRO_{año}.xls`. Se procesan con `actualizar_rb_hist_sc.py`, que actualiza `data/rb_hist_sc_enejul.json`.

> **Requisito de exportación (crítico):** al pedir estos archivos en Consulta Amigable, seleccionar explícitamente la columna de agrupación ("Rubro" o el nivel correspondiente) **antes** de exportar. Si el archivo trae solo una fila "TOTAL" en la tabla de detalle, significa que no se seleccionó la agrupación — no hay forma de recuperar el desglose después, hay que reexportar desde cero.

> **Metodología de acumulación (importante, verificado empíricamente):** los filtros "Trimestre I", "Trimestre II" y "Mes N" de Consulta Amigable devuelven valores **por período (incrementales)**, no acumulados. Es decir, `2T_RUBRO_{año}.xls` trae solo Abr-Jun, no Ene-Jun. Para obtener acumulados hay que sumar progresivamente:
> - `dev_t1` (acumulado a marzo) = valor de `1T_RUBRO`
> - `dev_t2` (acumulado a junio) = `1T_RUBRO` + `2T_RUBRO`
> - `dev_sem`/Ene-Jul = `1T_RUBRO` + `2T_RUBRO` + `JULIO_RUBRO`
>
> `actualizar_rb_hist_sc.py` ya aplica esta suma automáticamente y valida el resultado contra el dato existente (tolerancia 0.5%) antes de sobrescribir — si ves un `⚠️ DIFERENCIA detectada` en la consola, no lo ignores, revisa manualmente antes de continuar.

---

## Modo Manual (sin conexión)

El dashboard funciona igual abriendo `index.html` con doble clic y con los XLS en la carpeta `xls/` local — no requiere servidor ni conexión a internet. El fetch de los JSON de `data/` (históricos y snapshot semestral) es la única parte que depende de estar servido vía GitHub Pages o similar; si falta, el dashboard sigue funcionando con los datos en vivo de los 13 XLS.

---

## Hoja de Ruta

- **Fase 2a (actual):** descarga automatizada supervisada, con ventana de navegador visible y checkpoints de calidad (13/13 archivos, drill-down verificado, respaldo automático).
- **Fase 2b (futura, diferida deliberadamente):** automatización desatendida vía Windows Task Scheduler — solo después de varias semanas adicionales de Fase 2a estable, sin fallos de descarga ni necesidad de intervención manual.
- **Pendiente no urgente:** migrar los datos hardcodeados en `index.html` (constante `RB_HIST_SC`, bloque "Ver Primer Semestre") a un JSON externo en `data/`, consolidando toda la data histórica en un solo formato — hoy coexisten dos fuentes (HTML hardcodeado y JSON externo) para el mismo bloque, lo cual es deuda técnica aceptable a corto plazo pero no ideal a largo plazo.
