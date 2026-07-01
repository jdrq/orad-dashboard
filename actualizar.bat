@echo off
title ORPMI - Actualizador de Dashboard
color 0A

echo.
echo ==============================================================
echo    ORPMI - GORE LAMBAYEQUE  ^|  Actualizador BI
echo ==============================================================
echo.

:: Verificar Git
git --version > nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Git no esta instalado. Descargalo de git-scm.com
    pause & exit /b 1
)

echo  PASO 1: Verificando los 13 archivos en la carpeta xls\ ...
echo  --------------------------------------------------------------
set FALTAN=0

call :verificar ue_pliego.xls              "Pliego 452 - UE"
call :verificar nacional_gores.xls         "Todos los GOREs - Pliego"
call :verificar funciones_pliego.xls       "Pliego 452 - Funcion"
call :verificar rubro_pliego.xls           "Pliego 452 - Rubro"
call :verificar rubro_sede_central.xls     "UE 001-855 - Rubro"
call :verificar rubro_peot.xls             "UE 002-1133 - Rubro"
call :verificar rubro_agricultura.xls      "UE 100-856 - Rubro"
call :verificar rubro_transportes.xls      "UE 200-857 - Rubro"
call :verificar rubro_salud.xls            "UE 400-860 - Rubro"
call :verificar rubro_h_mercedes.xls       "UE 401-1001 - Rubro"
call :verificar rubro_h_belen.xls          "UE 402-1002 - Rubro"
call :verificar rubro_h_regional.xls       "UE 403-1422 - Rubro"
call :verificar proyectos_sede_central.xls "UE 001-855 - Por Proyecto"

echo.
if %FALTAN%==1 (
    echo  [ERROR] Faltan uno o mas archivos en xls\ ^(ver arriba^).
    echo  No se subio nada a GitHub. Copia los archivos que faltan
    echo  y vuelve a correr este script.
    echo.
    pause & exit /b 1
)

echo  [OK] Los 13 archivos estan completos.
echo.
echo  IMPORTANTE: deben llamarse EXACTAMENTE asi. El dashboard los
echo  busca por nombre dentro de xls\ cuando lo abre tu jefe.
echo.

echo  PASO 2: Subiendo los 13 archivos a GitHub...
echo  --------------------------------------------------------------
git add xls\*.xls
git commit -m "data: actualizacion diaria MEF - %date%"
git push origin main

if errorlevel 1 (
    echo.
    echo  [ERROR] El push fallo. Revisa los mensajes de Git arriba
    echo  ^(login/credenciales, conflicto, sin conexion, etc.^)
    pause & exit /b 1
)

echo.
echo ==============================================================
echo  [OK] Dashboard actualizado en GitHub Pages.
echo  Tu jefe abre el enlace de siempre y en 1-2 minutos ya ve
echo  los datos del dia, SIN tener que subir ningun archivo.
echo ==============================================================
echo.
pause
exit /b 0

:verificar
if not exist "xls\%~1" (
    echo    [FALTA] %~1                ^(%~2^)
    set FALTAN=1
) else (
    echo    [OK]    %~1
)
exit /b 0
