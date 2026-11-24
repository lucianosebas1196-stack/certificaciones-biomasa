@echo off
chcp 65001 >nul
REM Crea en el escritorio (y en el menu Inicio) el acceso directo "Certificaciones de biomasa" con el logo.
REM Va en la carpeta del programa (Sistema_web), junto a INICIAR_WEB_RED.bat. Sirve para la version 2.1 y la 3.0.
cd /d "%~dp0"
title Crear acceso directo - Certificaciones de biomasa
if not exist "servidor.py" (
  echo.
  echo  ERROR: este archivo tiene que estar en la carpeta del programa ^(la que tiene servidor.py^).
  echo  Copia CREAR_ACCESO_DIRECTO.bat, ABRIR_CERTIFICACIONES.bat y imagenes\icono.ico a Sistema_web.
  goto fin
)
if not exist "imagenes\icono.ico" (
  echo  ERROR: falta imagenes\icono.ico. Copialo a la carpeta imagenes del programa.
  goto fin
)
echo.
echo  Como se usa el programa en esta PC?
echo    1 = Es la PC SERVIDOR de la red ^(abre INICIAR_WEB_RED^)
echo    2 = Solo en esta PC, sin red ^(abre INICIAR_WEB^)
choice /c 12 /n /m "  Elegi 1 o 2: "
if errorlevel 2 (set MODO=LOCAL) else (set MODO=RED)
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$aqui='%~dp0'; $sh=New-Object -ComObject WScript.Shell;" ^
  "foreach($dir in @([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('Programs'))){" ^
  "  $s=$sh.CreateShortcut((Join-Path $dir 'Certificaciones de biomasa.lnk'));" ^
  "  $s.TargetPath=(Join-Path $aqui 'ABRIR_CERTIFICACIONES.bat'); $s.Arguments='%MODO%';" ^
  "  $s.WorkingDirectory=$aqui; $s.IconLocation=(Join-Path $aqui 'imagenes\icono.ico')+',0';" ^
  "  $s.Description='Programa de certificaciones de biomasa'; $s.Save() }"
if errorlevel 1 (
  echo.
  echo  No se pudo crear el acceso directo.
) else (
  echo.
  echo  Listo: en el escritorio y en el menu Inicio quedo "Certificaciones de biomasa" con el logo.
  echo  Para tenerlo en la barra de tareas: clic derecho en el acceso del menu Inicio - "Anclar a la barra de tareas".
)
:fin
echo.
echo  Apreta una tecla para cerrar.
pause >nul
