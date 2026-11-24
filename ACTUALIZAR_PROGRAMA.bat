@echo off
chcp 65001 >nul
REM Actualiza la instalacion que se usa todos los dias (Sistema_web) a esta version.
REM NO toca la carpeta trabajo. Antes hace copia de seguridad; despues corre las pruebas y,
REM si alguna falla, vuelve sola a la version anterior. Cerrar el servidor antes de ejecutarlo.
cd /d "%~dp0"
if not exist "actualizar.py" (
  echo  ERROR: abriste el .bat desde ADENTRO del zip. Extrae el zip primero y abri el .bat de la carpeta extraida.
  goto fin
)
call "%~dp0_python.bat"
if not defined PY set PY=python
title Actualizar programa de certificaciones
%PY% actualizar.py %1
:fin
echo.
echo  Apreta una tecla para cerrar esta ventana.
pause >nul
