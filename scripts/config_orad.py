"""Configuración única del proyecto ORAD: lee data/config.js (el mismo archivo que usa index.html).

Uso en cualquier script de scripts/:
    import config_orad as cfg
    cfg.anio_activo()          # 2026
    cfg.anios_historicos()     # [2015, ..., 2025]  (todos los años anteriores al activo)

Para cambiar de año se edita SOLO data/config.js (campo "anio_activo").
"""
import json
import re
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
RUTA_CONFIG = RAIZ / "data" / "config.js"


def _cargar():
    try:
        texto = RUTA_CONFIG.read_text(encoding="utf-8")
    except FileNotFoundError:
        raise SystemExit(f"[config] Falta {RUTA_CONFIG}. Debe existir en data/ (lo usa también index.html).")
    m = re.search(r"window\.ORAD_CONFIG\s*=\s*(\{.*?\})\s*;", texto, re.S)
    if not m:
        raise SystemExit("[config] No encuentro 'window.ORAD_CONFIG = {...};' en data/config.js")
    try:
        cfg = json.loads(m.group(1))
    except json.JSONDecodeError as e:
        raise SystemExit(f"[config] data/config.js: el bloque ORAD_CONFIG no es JSON válido: {e}")
    anio = cfg.get("anio_activo")
    primero = cfg.get("primer_anio_historico", 2015)
    if not isinstance(anio, int) or not (2016 <= anio <= 2100):
        raise SystemExit(f"[config] anio_activo inválido: {anio!r} (debe ser un entero, ej. 2026)")
    if not isinstance(primero, int) or primero >= anio:
        raise SystemExit(f"[config] primer_anio_historico inválido: {primero!r}")
    return cfg


def anio_activo() -> int:
    return _cargar()["anio_activo"]


def anios_historicos() -> list:
    c = _cargar()
    return list(range(c.get("primer_anio_historico", 2015), c["anio_activo"]))
