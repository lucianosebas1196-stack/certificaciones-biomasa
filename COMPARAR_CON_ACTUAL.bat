@echo off
chcp 65001 >nul
REM Compara ESTA version (la carpeta extraida del zip) con la instalacion que se usa todos los dias (Sistema_web),
REM con los ultimos archivos subidos alli. No cambia nada en ninguna de las dos. Tarda unos minutos.
REM Conviene correrlo ANTES de ACTUALIZAR_PROGRAMA.bat: tiene que decir IDENTICAS.
cd /d "%~dp0"
call "%~dp0_python.bat"
if not defined PY set PY=python
title Comparar version nueva con la actual
%PY% pruebas\comparar.py %1
echo.
pause
