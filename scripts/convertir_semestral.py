#!/usr/bin/env python3
# =============================================================================
# convertir_semestral.py
# ORAD - Gobierno Regional de Lambayeque
# Genera: data/historico_semestral.json
#
# USO:
#   python convertir_semestral.py
#
# ARCHIVOS REQUERIDOS (colocar en xls/historico/):
#   1T_2022.xls, 2T_2022.xls
#   1T_2023.xls, 2T_2023.xls
#   1T_2024.xls, 2T_2024.xls
#   1T_2025.xls, 2T_2025.xls
#
# NOTA 2026: Se inyecta automáticamente desde data/ejecucion.json
# =============================================================================

import os, re, json
from datetime import datetime
from bs4 import BeautifulSoup

# ---------------- CONFIGURACIÓN ----------------
AÑOS_HISTORICOS = [2022, 2023, 2024, 2025]
CARPETA_XLS     = "xls/historico"
CARPETA_DATA    = "data"
OUTPUT_FILE     = os.path.join(CARPETA_DATA, "historico_semestral.json")
EJECUCION_JSON  = os.path.join(CARPETA_DATA, "ejecucion.json")
CODIGO_LAMBAYEQUE = "452"

# -----------------------------------------------

def limpiar_numero(s):
    """Convierte string con comas y S/ a float."""
    s = str(s).replace(',', '').replace('S/', '').strip()
    try:
        return float(s) if s else 0.0
    except ValueError:
        return 0.0

def parsear_gores_trimestre(path):
    """
    Lee un archivo XLS del MEF (HTML disfrazado) con el ranking de 26 GOREs.
    Retorna dict {codigo: {nombre, dev, cert}}
    Estructura del MEF: Tabla 3 tiene los GOREs. Columnas:
      [0] Nombre, [1] PIA, [2] PIM, [3] Cert, [4] CompAnual,
      [5] AtenCompMens, [6] Devengado, [7] Girado, [8] Avance%
    """
    print(f"   Leyendo: {path}")
    if not os.path.exists(path):
        print(f"   ⚠️  ARCHIVO NO ENCONTRADO: {path}")
        return {}

    with open(path, 'rb') as f:
        content = f.read()

    soup = BeautifulSoup(content, 'html.parser')
    tables = soup.find_all('table')

    if len(tables) < 4:
        print(f"   ⚠️  Formato inesperado: solo {len(tables)} tablas en {path}")
        return {}

    tabla_gores = tables[3]
    gores = {}

    for r in tabla_gores.find_all('tr'):
        cols = [c.get_text(strip=True) for c in r.find_all(['td', 'th'])]
        if len(cols) >= 7 and re.match(r'^\d{3}:', cols[0]):
            codigo = cols[0][:3]
            gores[codigo] = {
                'codigo':   codigo,
                'nombre':   cols[0],
                'dev':      limpiar_numero(cols[6]),
                'cert':     limpiar_numero(cols[3]),
            }

    print(f"   → {len(gores)} GOREs encontrados")
    return gores

def construir_semestre(año):
    """Suma T1 + T2 para cada GORE y construye el ranking semestral."""
    path_t1 = os.path.join(CARPETA_XLS, f"1T_{año}.xls")
    path_t2 = os.path.join(CARPETA_XLS, f"2T_{año}.xls")

    t1 = parsear_gores_trimestre(path_t1)
    t2 = parsear_gores_trimestre(path_t2)

    codigos = sorted(set(t1.keys()) | set(t2.keys()))
    gores_sem = []

    for cod in codigos:
        g1 = t1.get(cod, {'nombre': '', 'dev': 0, 'cert': 0})
        g2 = t2.get(cod, {'nombre': '', 'dev': 0, 'cert': 0})
        nombre = g1['nombre'] or g2['nombre']
        dev_sem  = g1['dev'] + g2['dev']
        cert_sem = g1['cert'] + g2['cert']
        gores_sem.append({
            'codigo':         cod,
            'nombre':         nombre,
            'dev':            dev_sem,
            'dev_t1':         g1['dev'],
            'dev_t2':         g2['dev'],
            'es_lambayeque':  cod == CODIGO_LAMBAYEQUE,
        })

    # Ordenar por devengado y asignar posición
    gores_sem.sort(key=lambda x: x['dev'], reverse=True)
    for i, g in enumerate(gores_sem):
        g['posicion'] = i + 1

    lam = next((g for g in gores_sem if g['codigo'] == CODIGO_LAMBAYEQUE), {})

    return {
        'label':        f'Ene–Jun {año}',
        'total_gores':  len(gores_sem),
        'lambayeque': {
            'dev':      lam.get('dev', 0),
            'dev_t1':   lam.get('dev_t1', 0),
            'dev_t2':   lam.get('dev_t2', 0),
            'posicion': lam.get('posicion', 0),
        },
        'ranking': gores_sem,
    }

def inyectar_2026():
    """
    Lee ejecucion.json y extrae el dato acumulado de Lambayeque
    y el ranking actual de los 26 GOREs para representar el semestre 2026.
    """
    if not os.path.exists(EJECUCION_JSON):
        print(f"   ⚠️  {EJECUCION_JSON} no encontrado — 2026 quedará vacío")
        return None

    with open(EJECUCION_JSON, 'r', encoding='utf-8') as f:
        ej = json.load(f)

    ranking_raw = ej.get('ranking_nacional', {}).get('ranking_dev_sol', [])
    if not ranking_raw:
        ranking_raw = ej.get('ranking_nacional', {}).get('ranking_dev_pct', [])

    gores_2026 = []
    for g in ranking_raw:
        cod = g.get('codigo', '000')
        gores_2026.append({
            'codigo':        cod,
            'nombre':        g.get('nombre', ''),
            'dev':           g.get('dev', 0),
            'dev_t1':        0,  # No aplica para 2026
            'dev_t2':        0,
            'es_lambayeque': g.get('es_lambayeque', cod == CODIGO_LAMBAYEQUE),
        })

    # Re-ordenar por devengado (ya viene ordenado, pero aseguramos)
    gores_2026.sort(key=lambda x: x['dev'], reverse=True)
    for i, g in enumerate(gores_2026):
        g['posicion'] = i + 1

    lam = next((g for g in gores_2026 if g['es_lambayeque']), {})
    fecha = ej.get('fecha', datetime.today().strftime('%d/%m/%Y'))

    return {
        'label':       f'Ene–Jun 2026 (al {fecha})',
        'total_gores': len(gores_2026),
        'nota':        'Dato diario acumulado desde ejecucion.json. Aún no cierra el 2T 2026.',
        'lambayeque': {
            'dev':      lam.get('dev', 0),
            'dev_t1':   lam.get('dev', 0),
            'dev_t2':   0,
            'posicion': lam.get('posicion', 0),
        },
        'ranking': gores_2026,
    }

# ======================== MAIN ========================

def main():
    print("=" * 60)
    print("  ORAD — Generador de Histórico Semestral")
    print(f"  Fecha: {datetime.today().strftime('%d/%m/%Y %H:%M')}")
    print("=" * 60)

    os.makedirs(CARPETA_DATA, exist_ok=True)

    semestres = {}

    for año in AÑOS_HISTORICOS:
        print(f"\n📅 Procesando {año}...")
        semestres[str(año)] = construir_semestre(año)
        lam = semestres[str(año)]['lambayeque']
        print(f"   ✅ Lambayeque: T1=S/{lam['dev_t1']:,.0f} + T2=S/{lam['dev_t2']:,.0f} = S/{lam['dev']:,.0f}  (Posición {lam['posicion']}°)")

    print(f"\n📅 Procesando 2026 (desde ejecucion.json)...")
    dato_2026 = inyectar_2026()
    if dato_2026:
        semestres['2026'] = dato_2026
        lam26 = dato_2026['lambayeque']
        print(f"   ✅ Lambayeque: S/{lam26['dev']:,.0f}  (Posición {lam26['posicion']}°)")
    else:
        print("   ⚠️  2026 no procesado — revisa ejecucion.json")

    output = {
        "meta": {
            "descripcion":    "Primer semestre de cada año — Ranking GOREs por Devengado en Monto",
            "elaborado_por":  "ORAD - Oficina Regional de Administración",
            "gore":           "Gobierno Regional del Departamento de Lambayeque — Pliego 452",
            "fuente":         "Consulta Amigable MEF — apps5.mineco.gob.pe",
            "nota_metodologia": (
                "Primer semestre = T1 (Ene-Mar) + T2 (Abr-Jun). "
                "Archivos MEF trimestrales representan el período, no acumulado. "
                "2026 usa el dato diario acumulado del archivo ejecucion.json."
            ),
            "generado":       datetime.today().strftime('%Y-%m-%d %H:%M'),
        },
        "semestres": semestres,
    }

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    size_kb = os.path.getsize(OUTPUT_FILE) / 1024
    print(f"\n✅ Archivo generado: {OUTPUT_FILE}  ({size_kb:.1f} KB)")

    # Resumen final
    print("\n" + "=" * 60)
    print("  RESUMEN — DEVENGADO SEMESTRAL LAMBAYEQUE")
    print("=" * 60)
    for año in AÑOS_HISTORICOS + [2026]:
        if str(año) in semestres:
            lam = semestres[str(año)]['lambayeque']
            pos = lam['posicion']
            dev = lam['dev']
            total = semestres[str(año)]['total_gores']
            print(f"  {año}: S/ {dev:>15,.0f}   Posición: {pos}° / {total}")
    print("=" * 60)

if __name__ == "__main__":
    main()
