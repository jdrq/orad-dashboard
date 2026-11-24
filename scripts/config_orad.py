"""Configuración única del proyecto ORAD (lee config.json de la raíz del repo).

Uso en cualquier script de scripts/:
    import config_orad as cfg
    cfg.anio_activo()          # 2026
    cfg.anios_historicos()     # [2015, ..., 2025]  (todos los años anteriores al activo)
"""
import json
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
RUTA_CONFIG = RAIZ / "config.json"


def _cargar():
    try:
        cfg = json.loads(RUTA_CONFIG.read_text(encoding="utf-8"))
    except FileNotFoundError:
        raise SystemExit(f"[config] Falta {RUTA_CONFIG}. Debe existir en la raíz del repo.")
    except json.JSONDecodeError as e:
        raise SystemExit(f"[config] config.json inválido: {e}")
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
