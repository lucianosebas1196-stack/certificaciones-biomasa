@echo off
chcp 65001 >nul
REM Copia de seguridad de la carpeta "trabajo" (certificaciones, numeracion, config, usuarios) en respaldos\
REM El programa ya hace una copia SOLA por dia mientras esta abierto. Este archivo sirve para hacer una al momento.
REM Para guardar tambien una copia FUERA de esta PC: escribir la carpeta en trabajo\respaldo_destino.txt
cd /d "%~dp0"
call "%~dp0_python.bat"
if not defined PY set PY=python
%PY% respaldo.py
echo.
pause
