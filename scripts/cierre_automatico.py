#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# =============================================================================
# cierre_automatico.py  —  decide SOLO si hoy toca una acción mensual y la ejecuta
# ORAD - Gobierno Regional de Lambayeque
#
# Se corre DESPUÉS de la descarga diaria (scripts/descargar_xls_mef.py), en la nube
# (GitHub Actions) o en tu PC. Es IDEMPOTENTE: si no toca nada, termina en segundos
# sin tocar el repositorio, así que puede correr todos los días y varias veces al día.
#
# QUÉ DECIDE (con la fecha de hoy en Perú y el estado de data/mensual-actual.js):
#   CERRAR     el último mes cerrado es anterior al mes pasado
#              -> cierre_mensual.py --mes N        (mensual + históricos)
#   CONFIRMAR  el último mes cerrado sigue "provisional" y ya llegó su fecha
#              de revisión (3 días hábiles) -> cierre_mensual.py --sin-historicos
#              (el MEF reclasifica meses ya cerrados; esto lo vuelve a validar:
#               si no cambió pasa a CONFIRMADO; si cambió, lo aplica y lo registra)
#   NADA       cualquier otro caso
#
# Los controles son los de siempre (cuadre contra el totalizador del MEF, mes
# completo, consulta posterior al fin de mes, validación de contenido de cada
# histórico). Si algo no cuadra NO se escribe nada y el script termina con error
# (en GitHub Actions eso deja la corrida en rojo y GitHub avisa por correo).
#
# USO (desde la raíz del repo):
#   python scripts\cierre_automatico.py              decide y ejecuta
#   python scripts\cierre_automatico.py --dry-run    solo muestra qué haría (no escribe)
#   python scripts\cierre_automatico.py --sin-git    ejecuta pero no hace commit/push
#
# CÓDIGOS DE SALIDA: 0 = nada que hacer o todo bien · 1 = falló la validación o Git ·
#                    2 = lo mensual quedó bien pero faltan/fallaron históricos.
# LIMITACIÓN CONOCIDA: en enero NO cierra diciembre del año anterior (requiere los
# archivos del año anterior); avisa y deja ese cierre como paso manual.
# =============================================================================

import argparse
import os
import subprocess
import sys
from datetime import date, datetime, timedelta, timezone

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)
sys.path.insert(0, AQUI)

import generar_mensual as gm  # noqa: E402

ARCHIVOS_DATOS = ["data/mensual-actual.js", "data/historico_progresivo.json",
                  "data/rb_hist_sc_progresivo.json"]


def hoy_peru():
    """Fecha de hoy en Perú (UTC-5 fijo, sin horario de verano). En GitHub Actions el
    reloj está en UTC; así la decisión es la misma en la nube y en tu PC."""
    return datetime.now(timezone(timedelta(hours=-5))).date()


def en_actions():
    return os.environ.get("GITHUB_ACTIONS") == "true"


def decidir(hoy, previo):
    """Devuelve (accion, mes, motivo). accion: 'cerrar' | 'confirmar' | 'nada' | 'aviso'."""
    objetivo = hoy.month - 1
    if objetivo == 0:
        return ("aviso", None, "Enero: el cierre de diciembre del año anterior no es automático "
                               "(usa los archivos del año anterior). Hacerlo manualmente.")
    mismo_anio = previo is not None and previo.get("anio") == hoy.year
    ultimo = previo.get("ultimo_mes_cerrado", 0) if mismo_anio else 0
    if ultimo < objetivo:
        return ("cerrar", objetivo,
                f"el último cierre guardado es {gm.MESES_LARGO[ultimo-1] if ultimo else 'ninguno'} "
                f"y ya terminó {gm.MESES_LARGO[objetivo-1]}")
    est = (previo or {}).get("estado") or {}
    revisar = est.get("revisar_desde")
    if est.get("provisional") and revisar and hoy >= date.fromisoformat(revisar):
        return ("confirmar", ultimo,
                f"{gm.MESES_LARGO[ultimo-1]} sigue provisional y su fecha de revisión "
                f"({date.fromisoformat(revisar):%d/%m/%Y}) ya llegó")
    if est.get("provisional"):
        return ("nada", None, f"{gm.MESES_LARGO[ultimo-1]} está provisional; se revisa desde "
                              f"el {date.fromisoformat(revisar):%d/%m/%Y}")
    return ("nada", None, f"al día: {gm.MESES_LARGO[ultimo-1]} cerrado y confirmado")


def _git(*args):
    return subprocess.run(["git", *args], cwd=RAIZ, capture_output=True, text=True, encoding="utf-8")


def publicar(mensaje):
    """git add de los datos mensuales/históricos -> commit -> push (con rebase si se rechaza)."""
    if en_actions():
        _git("config", "user.name", "orad-bot")
        _git("config", "user.email", "orad-bot@users.noreply.github.com")
    cambios = _git("status", "--porcelain", "--", *ARCHIVOS_DATOS).stdout.strip()
    if not cambios:
        print("  Git: los archivos de datos no cambiaron; no hay nada que subir.")
        return True
    for cmd in (["add", *ARCHIVOS_DATOS], ["commit", "-m", mensaje]):
        r = _git(*cmd)
        if r.returncode != 0:
            print(f"  [ERROR] git {cmd[0]}: {(r.stdout + r.stderr).strip()[:300]}")
            return False
    r = _git("push", "origin", "main")
    if r.returncode != 0 and any(t in (r.stdout + r.stderr) for t in ("rejected", "fetch first", "non-fast-forward")):
        print("  El push fue rechazado (el repo avanzó). Integro con 'git pull --rebase' y reintento...")
        r = _git("pull", "--rebase", "origin", "main")
        if r.returncode == 0:
            r = _git("push", "origin", "main")
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
    ap = argparse.ArgumentParser(description="Cierre mensual automático del dashboard ORAD.")
    ap.add_argument("--dry-run", action="store_true", help="muestra qué haría, sin escribir nada.")
    ap.add_argument("--sin-git", action="store_true", help="ejecuta pero no hace commit/push.")
    args = ap.parse_args()

    os.chdir(RAIZ)
    hoy = hoy_peru()
    try:
        previo = gm.cargar_previo()
    except Exception as e:
        print(f"[ERROR] No se pudo leer data/mensual-actual.js: {e}")
        sys.exit(1)

    accion, mes, motivo = decidir(hoy, previo)
    print(f"Cierre automático — hoy (Perú): {hoy:%d/%m/%Y} — decisión: {accion.upper()}"
          + (f" (mes {mes})" if mes else ""))
    print(f"  Motivo: {motivo}")

    if accion == "aviso":
        if en_actions():
            print(f"::warning title=Cierre automático::{motivo}")
        sys.exit(0)
    if accion == "nada":
        sys.exit(0)

    cmd = [sys.executable, os.path.join("scripts", "cierre_mensual.py"), "--mes", str(mes)]
    if accion == "confirmar":
        cmd.append("--sin-historicos")
    if args.dry_run:
        cmd.append("--dry-run")
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    p = subprocess.run(cmd, cwd=RAIZ, capture_output=True, text=True, encoding="utf-8",
                       errors="replace", env=env)
    print((p.stdout or "") + (p.stderr or ""))
    rc = p.returncode

    if args.dry_run:
        print("(dry-run: no se escribió ni se subió nada)")
        sys.exit(rc)

    anio = (gm.cargar_previo() or {}).get("anio", hoy.year)
    nombre_mes = gm.MESES_LARGO[mes - 1].lower()
    mensaje = (f"data: cierre {nombre_mes} {anio} (automático)" if accion == "cerrar"
               else f"data: confirmación de {nombre_mes} {anio} (automático)")
    if rc in (0, 2) and not args.sin_git:
        # Con código 2 lo mensual quedó bien y faltan históricos: se sube lo válido y se avisa.
        if not publicar(mensaje):
            rc = 1
    if rc != 0:
        detalle = ("falló la validación; no se escribió nada" if rc == 1
                   else "lo mensual quedó bien pero faltan o fallaron históricos")
        if en_actions():
            print(f"::error title=Cierre automático de {nombre_mes}::{detalle}")
        print(f"[!!] Cierre automático terminó con código {rc}: {detalle}.")
    sys.exit(rc)


if __name__ == "__main__":
    main()
