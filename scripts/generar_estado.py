#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# =============================================================================
# generar_estado.py  —  deja constancia de la ÚLTIMA ACTUALIZACIÓN VERIFICADA
# ORAD - Gobierno Regional de Lambayeque
#
# Escribe data/estado.json, que el dashboard muestra como banner:
#   "Datos verificados · Consulta MEF: 24/11/2026 · 17/17 archivos válidos · ..."
#
# Se corre en el workflow diario DESPUÉS de la descarga y del cierre mensual, y solo
# si los pasos anteriores terminaron bien (comportamiento normal de GitHub Actions).
# Aun así vuelve a comprobar los 17 archivos (mismo control que la compuerta de Git):
# si algo no está bien NO escribe nada y termina con error. Así, estado.json solo
# puede decir "verificado" cuando de verdad lo está; si algún día falla la
# actualización, el banner envejece solo y se pone en ámbar/rojo.
#
# USO (desde la raíz del repo):
#   python scripts\generar_estado.py              escribe, hace commit y push
#   python scripts\generar_estado.py --sin-git    escribe pero no sube
# También corre los controles de calidad (scripts/controles_calidad.py): informan, no bloquean.
# CÓDIGOS DE SALIDA: 0 = ok · 1 = no verificado o falló Git.
# =============================================================================

import argparse
import json
import os
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
RAIZ = Path(__file__).resolve().parent.parent
RUTA_ESTADO = RAIZ / "data" / "estado.json"

import config_orad as cfg                 # noqa: E402
import descargar_xls_mef as d             # noqa: E402  (reutiliza la compuerta 17/17 y los helpers de Git)
import generar_mensual as gm              # noqa: E402
import controles_calidad as cc            # noqa: E402

PERU = timezone(timedelta(hours=-5))


def construir_estado():
    ok, problemas = d.validar_archivos()
    if not ok:
        print("[ERROR] No se puede declarar 'verificado': " + "; ".join(problemas))
        return None
    total = len(d.ARCHIVOS)
    consulta = d.fecha_consulta(d.CARPETA_DESTINO / d.ARCHIVOS[0]["nombre"])
    ahora = datetime.now(PERU).replace(microsecond=0)

    cierre = {}
    try:
        m = gm.cargar_previo() or {}
        est = m.get("estado") or {}
        cierre = {
            "anio": m.get("anio"),
            "ultimo_mes_cerrado": m.get("ultimo_mes_cerrado"),
            "provisional": bool(est.get("provisional")),
            "revisar_desde": est.get("revisar_desde"),
            "confirmado_el": est.get("confirmado_el"),
        }
    except Exception as e:                # el estado diario no debe caerse por esto
        print(f"[AVISO] No se pudo leer data/mensual-actual.js: {e}")

    # Controles de calidad (informan, NO bloquean). Si fallan por un error propio, no tumban el estado.
    try:
        res = cc.ejecutar(guardar=True)
        calidad = {"revisados": res["revisados"], "alertas": res["alertas"], "avisos": res["avisos"],
                   "detalle": res["detalle"], "notas": res["notas"]}
    except Exception as e:
        print(f"[AVISO] Controles de calidad no disponibles: {e}")
        calidad = {"error": str(e)[:200]}

    r = d._git("rev-parse", "--short", "HEAD")
    return {
        "version": 1,
        "actualizado": ahora.isoformat(),          # con zona horaria de Lima (-05:00)
        "fecha_consulta_mef": consulta.isoformat() if consulta else None,
        "archivos": {"esperados": total, "validos": total},
        "anio_activo": cfg.anio_activo(),
        "cierre_mensual": cierre,
        "calidad": calidad,
        "origen": "github-actions" if d.en_actions() else "manual",
        "corrida": os.environ.get("GITHUB_RUN_NUMBER"),
        "commit_datos": r.stdout.strip() if r.returncode == 0 else None,
    }


def publicar():
    d.preparar_git_ci()
    mensaje = "data: estado de la actualización diaria (automático)"
    for cmd in (["add", "data/estado.json", "data/calidad_historial.json"], ["commit", "-m", mensaje]):
        r = d._git(*cmd)
        if r.returncode != 0:
            salida = r.stdout + r.stderr
            if "nothing to commit" in salida or "no changes added" in salida:
                print("  estado.json sin cambios.")
                return True
            print(f"  [ERROR] git {cmd[0]}: {salida.strip()[:300]}")
            return False
    r = d._git("push", "origin", "main")
    if r.returncode != 0 and any(t in (r.stdout + r.stderr) for t in ("rejected", "fetch first", "non-fast-forward")):
        print("  El push fue rechazado (el repo avanzó). Integro con 'git pull --rebase' y reintento...")
        if d._git("pull", "--rebase", "origin", "main").returncode == 0:
            r = d._git("push", "origin", "main")
    if r.returncode != 0:
        print(f"  [ERROR] git push: {(r.stdout + r.stderr).strip()[:300]}")
        return False
    print(f"  Git: subido -> {mensaje}")
    return True


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="Escribe data/estado.json (última actualización verificada).")
    ap.add_argument("--sin-git", action="store_true", help="escribe el archivo pero no hace commit/push.")
    args = ap.parse_args()
    os.chdir(RAIZ)

    estado = construir_estado()
    if estado is None:
        if d.en_actions():
            print("::error title=Estado no verificado::los 17 archivos no pasan la validación; estado.json no se actualizó")
        sys.exit(1)
    RUTA_ESTADO.write_text(json.dumps(estado, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Estado verificado: {estado['archivos']['validos']}/{estado['archivos']['esperados']} archivos · "
          f"consulta MEF {estado['fecha_consulta_mef']} · {estado['actualizado']}")
    cal = estado["calidad"]
    if "error" not in cal:
        print(f"Calidad: {cal['revisados']} verificaciones · {cal['alertas']} alerta(s) · {cal['avisos']} aviso(s)")
        for dd in cal["detalle"]:
            print(f"  [{dd['severidad'].upper()}] {dd['id']}: {dd['mensaje']}")
            if d.en_actions() and dd["severidad"] == "alerta":
                print(f"::warning title=Calidad de datos {dd['id']}::{dd['mensaje']}")
    if args.sin_git:
        return
    if not publicar():
        if d.en_actions():
            print("::error title=Estado::no se pudo subir data/estado.json")
        sys.exit(1)


if __name__ == "__main__":
    main()
