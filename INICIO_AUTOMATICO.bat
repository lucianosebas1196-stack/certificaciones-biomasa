@echo off
REM Hace que el servidor arranque solo cada vez que se inicia sesion en esta PC.
REM Ejecutar UNA vez (clic derecho -> Ejecutar como administrador). Para quitarlo: QUITAR_INICIO_AUTOMATICO.bat
schtasks /create /f /tn "Certificaciones biomasa" /sc onlogon /rl limited /tr "\"%~dp0INICIAR_WEB_RED.bat\""
if errorlevel 1 (echo No se pudo crear la tarea.) else (echo Listo: el servidor va a arrancar solo al iniciar sesion.)
REM Que la PC no se suspenda mientras esta enchufada (si se duerme, nadie puede certificar)
powercfg /change standby-timeout-ac 0
powercfg /change hibernate-timeout-ac 0
pause
