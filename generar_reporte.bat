@echo off
cd /d "%~dp0"
echo Actualizando datos desde GitHub...
git pull --ff-only origin main
if errorlevel 1 echo AVISO: no se pudo actualizar; se usaran los datos que hay en esta PC.
python scripts\reporte_quincenal.py %*
if errorlevel 1 (echo. & echo No se genero el informe. & pause & exit /b 1)
start "" reportes
pause
