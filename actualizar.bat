@echo off
title ORPMI - Actualizador de Dashboard
color 0A

echo.
echo ==============================================================
echo    ORPMI - GORE LAMBAYEQUE  ^|  Actualizador BI
echo ==============================================================
echo.
echo  PASO 1: Verifica que tienes los 12 archivos en la carpeta xls\
echo  --------------------------------------------------------------
echo    ue_pliego.xls           (Pliego 452 - UE)
echo    nacional_gores.xls      (Todos los GOREs - Pliego)
echo    funciones_pliego.xls    (Pliego 452 - Funcion)
echo    rubro_pliego.xls        (Pliego 452 - Rubro)
echo    rubro_sede_central.xls  (UE 001-855 - Rubro)
echo    rubro_peot.xls          (UE 002-1133 - Rubro)
echo    rubro_agricultura.xls   (UE 100-856 - Rubro)
echo    rubro_transportes.xls   (UE 200-857 - Rubro)
echo    rubro_salud.xls         (UE 400-860 - Rubro)
echo    rubro_h_mercedes.xls    (UE 401-1001 - Rubro)
echo    rubro_h_belen.xls       (UE 402-1002 - Rubro)
echo    rubro_h_regional.xls    (UE 403-1422 - Rubro)
echo.
set /p continuar="  Los tienes todos? Presiona ENTER para continuar..."
echo.

:: Verificar Python
python --version > nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python no instalado. Descargalo de python.org
    pause & exit /b 1
)

:: Instalar dependencias si faltan
python -c "import bs4" > nul 2>&1
if errorlevel 1 (
    echo  Instalando dependencias...
    pip install beautifulsoup4 lxml --quiet
)

:: Ejecutar conversion XLS a JSON
echo  Convirtiendo XLS a JSON...
echo.
python scripts\convertir_xls_a_json.py
if errorlevel 1 (
    echo.
    echo  [ERROR] Revisa los mensajes arriba.
    pause & exit /b 1
)

echo.
echo ==============================================================
set /p subir="  Subir a GitHub ahora? (S/N): "
if /i "%subir%"=="S" (
    git add data\ejecucion.json
    git commit -m "data: actualizacion diaria MEF - %date%"
    git push origin main
    echo.
    echo  [OK] Dashboard actualizado en GitHub Pages.
    echo  Tu jefe ya puede verlo en unos minutos.
) else (
    echo  JSON listo en data\ejecucion.json
    echo  Subelo manualmente con:
    echo    git add data\ejecucion.json
    echo    git commit -m "data: actualizacion diaria"
    echo    git push origin main
)
echo.
pause
