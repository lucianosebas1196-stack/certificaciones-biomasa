@echo off
chcp 65001 >nul
REM Vuelve a la version que estaba antes de la ultima actualizacion (guardada en versiones_anteriores).
REM NO toca la carpeta trabajo. Cerrar el servidor antes de ejecutarlo.
cd /d "%~dp0"
call "%~dp0_python.bat"
if not defined PY set PY=python
title Volver a la version anterior
%PY% actualizar.py --volver
echo.
echo  Apreta una tecla para cerrar esta ventana.
pause >nul
