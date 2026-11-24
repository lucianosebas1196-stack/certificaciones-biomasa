@echo off
chcp 65001 >nul
REM Prueba el programa en ESTA PC (calculos, seguridad, base de usuarios, config abierto en Excel, copias de
REM seguridad y un mes completo con los ultimos archivos subidos). Trabaja sobre una COPIA: no cambia nada real.
REM Conviene correrlo despues de instalar o actualizar. Tarda 1 a 3 minutos.
cd /d "%~dp0"
call "%~dp0_python.bat"
if not defined PY set PY=python
title Pruebas del programa de certificaciones
%PY% pruebas\probar.py
echo.
pause
