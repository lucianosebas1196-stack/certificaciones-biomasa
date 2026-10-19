@echo off
REM Crea un Administrador nuevo o le pone una clave nueva a uno existente (por si se olvido).
REM Solo funciona sentado en la PC servidor: quien tiene acceso a esta PC puede administrar el programa.
cd /d "%~dp0"
set PY=python
python --version >nul 2>nul || set PY=py
%PY% servidor.py --nuevo-admin
echo.
echo  Apreta una tecla para cerrar esta ventana.
pause >nul
