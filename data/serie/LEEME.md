# Serie diaria ORAD (capa de datos)

Tablas limpias, una fila por día y elemento. Se actualizan solas con la corrida diaria (`scripts/generar_serie.py`).

| Tabla | Grano |
|---|---|
| fact_pliego_diario | día × Pliego 452 |
| fact_ue_diario | día × unidad ejecutora |
| fact_rubro_diario | día × ámbito (PLIEGO o UE) × rubro |
| fact_funcion_diario | día × función |
| fact_gore_diario | día × gobierno regional (nacional) |
| fact_proyecto_diario | día × proyecto (Sede Central) |
| dim_* | códigos y nombres |

Medidas: pia, pim, certificacion, compromiso_anual, atencion_mensual, devengado, girado, avance_pct (montos en soles).
`fecha` = fecha de consulta del MEF.

## Conectar Power BI
Obtener datos > Web > URL, por ejemplo:
`https://raw.githubusercontent.com/jdrq/orad-dashboard/main/data/serie/fact_ue_diario.csv`
Elegir "Texto/CSV", codificación UTF-8. Relacionar fact_ue_diario[ue_codigo] con dim_ue[ue_codigo] (varios a uno).
