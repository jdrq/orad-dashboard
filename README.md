# Dashboard BI — ORAD · Gobierno Regional de Lambayeque

Dashboard web de ejecución presupuestal de inversiones del **Pliego 452 - GORE Lambayeque**, construido con HTML + JavaScript + Chart.js, publicado vía GitHub Pages.

**Elaborado por:** ORPMI — Oficina Regional de Programación Multianual de Inversiones  
**Fuente de datos:** Consulta Amigable MEF · [apps5.mineco.gob.pe](https://apps5.mineco.gob.pe/transparencia/Navegador/default.aspx)

---

## Contenido del Dashboard

| Bloque | Descripción |
|--------|-------------|
| Bloque 1 | KPIs financieros del Pliego (PIM, Certificación, Compromiso, Devengado) |
| Bloque 2 | Ranking nacional de los 26 Gobiernos Regionales |
| Bloque 3 | Ejecución por Unidad Ejecutora |
| Bloque 6 | Ejecución por Rubro / Fuente de Financiamiento |
| Bloque 7 | Ejecución por Función a nivel Pliego |
| Bloque 8 | Comparativo histórico (2016–2026) |
| Bloque 9 | Proyección acumulada al cierre del año |

---

## Arquitectura

```
Consulta Amigable MEF
        ↓
  (descargar 5 XLS)
        ↓
  python scripts/convertir_xls_a_json.py
        ↓
  data/ejecucion.json
        ↓
  git push → GitHub Pages
        ↓
  Tu jefe abre el enlace y ve los datos
```

---

## Estructura del Proyecto

```
orpmi-dashboard/
├── index.html                  ← Dashboard (no modificar)
├── actualizar.bat              ← Doble clic para actualizar (Windows)
├── data/
│   ├── ejecucion.json          ← Datos publicados (se actualiza diariamente)
│   └── historico/              ← XLS históricos 2016-2025 (cargar una vez)
│       ├── historico_2016.xls
│       ├── historico_2017.xls
│       └── ...
├── scripts/
│   └── convertir_xls_a_json.py ← Script de conversión
├── xls/                        ← Carpeta para los XLS del día (NO se sube a Git)
│   ├── ue_pliego.xls
│   ├── rubro_pliego.xls
│   ├── rubro_sede_central.xls
│   ├── nacional_gores.xls
│   └── funciones_pliego.xls
└── .gitignore
```

---

## Cómo Actualizar Diariamente (5 minutos)

### Paso 1 — Descargar los 5 XLS de Consulta Amigable

| Archivo a guardar como | Consulta Amigable — configuración |
|------------------------|-----------------------------------|
| `ue_pliego.xls` | Pliego 452 → Por Unidad Ejecutora → Solo Proyectos |
| `rubro_pliego.xls` | Pliego 452 → Por Rubro → Solo Proyectos |
| `rubro_sede_central.xls` | UE 001-855 → Por Rubro → Solo Proyectos |
| `nacional_gores.xls` | Todos los GOREs → Por Pliego → Solo Proyectos |
| `funciones_pliego.xls` | Pliego 452 → Por Función → Solo Proyectos |

> Pegar todos en la carpeta `xls/`

### Paso 2 — Ejecutar el actualizador

Doble clic en `actualizar.bat`

O desde la terminal:
```bash
python scripts/convertir_xls_a_json.py
git add data/ejecucion.json
git commit -m "Actualización datos MEF - $(date +%Y-%m-%d)"
git push
```

### Paso 3 — Listo

Tu jefe puede ver los datos en: `https://[tu-usuario].github.io/orpmi-dashboard/`

---

## Archivos Históricos (cargar una vez)

Para el Bloque 8 (comparativo histórico), coloca los XLS de años anteriores en `data/historico/` con el nombre `historico_YYYY.xls` (mismo formato que el archivo `ue_pliego.xls` pero del año correspondiente).

Luego ejecuta el script normalmente. Los años históricos quedarán fijos en el JSON.

---

## Instalación de Dependencias (primera vez)

```bash
pip install beautifulsoup4 lxml
```

---

## Modo Manual (sin servidor)

El dashboard también funciona abriendo `index.html` directamente con doble clic y cargando los XLS manualmente — igual que antes. El módulo JSON solo se activa cuando hay un servidor HTTP (GitHub Pages).
