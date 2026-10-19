@echo off
REM Permite que las otras PC de la red de la empresa entren al puerto 8000 (solo red privada/de dominio).
REM Ejecutar UNA vez: clic derecho -> "Ejecutar como administrador".
netsh advfirewall firewall delete rule name="Certificaciones biomasa" >nul 2>&1
netsh advfirewall firewall add rule name="Certificaciones biomasa" dir=in action=allow protocol=TCP localport=8000 profile=private,domain
if errorlevel 1 (echo No se pudo: ejecutalo con clic derecho - Ejecutar como administrador.) else (echo Listo: puerto 8000 abierto solo para la red de la empresa.)
pause
