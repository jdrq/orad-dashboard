#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# =============================================================================
# registrar_fallo.py  —  deja constancia de un intento de actualización FALLIDO
# ORAD - Gobierno Regional de Lambayeque
#
# Lo ejecuta el workflow SOLO cuando un paso falla (if: failure()). Escribe
# data/ultimo_intento.json; el banner lo muestra si es posterior a la última
# actualización verificada (data/estado.json NO se toca: sigue significando "verificado").
# Cuando una corrida posterior sale bien, el aviso desaparece solo.
#
# CAUSAS (se deducen del registro logs/descarga_*.log de la corrida):
#   mef        archivos que no se pudieron descargar tras los reintentos
#   tiempo     la descarga fue cortada por el límite de tiempo
#   validacion los archivos llegaron, pero no pasaron la compuerta 17/17
#   proceso    falló otro paso (Git, publicación, cierre, estado)
#
# USO:  python scripts/registrar_fallo.py [--sin-git]
# =============================================================================

import argparse
import json
import os
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
RUTA = RAIZ / "data" / "ultimo_intento.json"
PERU = timezone(timedelta(hours=-5))

MENSAJES = {
    "mef": "el portal del MEF no entregó todos los archivos",
    "tiempo": "el portal del MEF no respondió a tiempo",
    "validacion": "los archivos descargados no pasaron la validación",
    "proceso": "hubo un error interno del proceso",
}


def log_de_la_corrida():
    """Texto de la última corrida registrada en logs/descarga_*.log ('' si no hay)."""
    logs = sorted((RAIZ / "logs").glob("descarga_*.log"), key=lambda p: p.stat().st_mtime)
    if not logs:
        return ""
    return logs[-1].read_text(encoding="utf-8", errors="replace").split("===== Inicio ")[-1]


def deducir_causa(texto, resultado_descarga):
    """resultado_descarga: 'success' | 'failure' | '' (outcome del paso de descarga)."""
    if resultado_descarga == "success" or not texto:
        return "proceso"
    if "NO se ejecuta Git: archivos que no se pudieron descargar" in texto:
        return "mef"
    if "GIT PUSH BLOQUEADO" in texto:
        return "validacion"
    if "VALIDACIÓN FINAL" in texto:          # descargó todo; falló Git o la publicación
        return "proceso"
    return "tiempo"                           # se cortó en plena descarga


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="Registra un intento de actualización fallido.")
    ap.add_argument("--sin-git", action="store_true")
    args = ap.parse_args()
    os.chdir(RAIZ)

    causa = deducir_causa(log_de_la_corrida(), os.environ.get("RESULTADO_DESCARGA", ""))
    dato = {
        "fecha": datetime.now(PERU).replace(microsecond=0).isoformat(),
        "causa": causa,
        "mensaje": MENSAJES[causa],
        "corrida": os.environ.get("GITHUB_RUN_NUMBER"),
    }
    RUTA.write_text(json.dumps(dato, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Intento fallido registrado: {causa} — {MENSAJES[causa]}")
    if args.sin_git:
        return

    import descargar_xls_mef as d            # reutiliza los helpers de Git
    d.preparar_git_ci()
    for cmd in (["add", "data/ultimo_intento.json"],
                ["commit", "-m", "data: registro de intento fallido (automático)", "--", "data/ultimo_intento.json"]):
        r = d._git(*cmd)
        if r.returncode != 0:
            print(f"[ERROR] git {cmd[0]}: {(r.stdout + r.stderr).strip()[:300]}")
            sys.exit(1)
    r = d._git("push", "origin", "main")
    if r.returncode != 0:                     # --autostash: puede haber XLS a medio descargar sin commitear
        if d._git("pull", "--rebase", "--autostash", "origin", "main").returncode == 0:
            r = d._git("push", "origin", "main")
    if r.returncode != 0:
        print(f"[ERROR] git push: {(r.stdout + r.stderr).strip()[:300]}")
        sys.exit(1)
    print("  Git: aviso publicado.")


if __name__ == "__main__":
    main()
