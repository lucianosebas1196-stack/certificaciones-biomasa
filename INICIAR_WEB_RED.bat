@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "servidor.py" (
  echo.
  echo  ERROR: no encuentro servidor.py en esta carpeta.
  echo  Seguramente abriste el .bat desde ADENTRO del zip.
  echo  Solucion: clic derecho en Sistema_web.zip - "Extraer todo", y abri el .bat de la carpeta extraida.
  goto fin
)
set PY=
python --version >nul 2>nul && set PY=python
if not defined PY py --version >nul 2>nul && set PY=py
if not defined PY (
  echo.
  echo  ERROR: Python no esta instalado en esta PC ^(o no esta en el PATH^).
  echo  Instalalo desde https://www.python.org/downloads/ y en la primera pantalla
  echo  marca la casilla "Add python.exe to PATH". Despues volve a abrir este archivo.
  goto fin
)
%PY% -c "import fastapi, uvicorn, multipart, pandas, pyxlsb, openpyxl, reportlab, pypdf, ipywidgets, cryptography" >nul 2>nul
if errorlevel 1 (
  echo Instalando lo necesario ^(solo la primera vez, 1-2 minutos^)...
  %PY% -m pip install -r requirements_web.txt
  if errorlevel 1 (
    echo.
    echo  ERROR: no se pudieron instalar las librerias. Revisa la conexion a internet y volve a intentar.
    goto fin
  )
)
title Certificaciones biomasa - SERVIDOR DE RED (no cerrar esta ventana)
echo.
echo  ==============================================================
echo   SERVIDOR DE CERTIFICACIONES - dejar esta ventana ABIERTA
echo   (se puede minimizar). No hagas clic adentro: si lo hiciste
echo   y aparece "Seleccionar" en el titulo, apreta Esc.
echo  ==============================================================
%PY% servidor.py 8000 --red
echo.
echo  El servidor se detuvo. Si hay un mensaje de error arriba, sacale una foto.
:fin
echo.
echo  Apreta una tecla para cerrar esta ventana.
pause >nul
