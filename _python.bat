@echo off
REM Elige el Python a usar y lo deja en la variable PY (lo llaman los demas .bat).
REM 1) el Python portatil de la carpeta python\ (no hace falta instalar nada)
REM 2) si la carpeta python\ todavia no tiene python.exe, lo prepara una sola vez (baja 11 MB de python.org)
REM 3) si no hay portatil, el Python instalado en la PC
set PY=
if exist "%~dp0python\python.exe" goto portatil
if exist "%~dp0python\Lib\site-packages" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0preparar_python.ps1"
  if exist "%~dp0python\python.exe" goto portatil
)
python --version >nul 2>nul && set "PY=python"
if not defined PY py --version >nul 2>nul && set "PY=py"
goto :eof
:portatil
set PY="%~dp0python\python.exe"
goto :eof
