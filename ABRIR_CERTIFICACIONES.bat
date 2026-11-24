@echo off
chcp 65001 >nul
REM Lo abre el acceso directo "Certificaciones de biomasa".
REM Si el programa ya esta funcionando (por ejemplo, por el inicio automatico) solo abre el navegador;
REM si no, lo arranca como siempre. Asi nunca quedan dos servidores a la vez.
cd /d "%~dp0"
set MODO=%1
if /i "%MODO%"=="LOCAL" (set ESQUEMA=http) else (set ESQUEMA=https)
powershell -NoProfile -Command "try{$c=New-Object Net.Sockets.TcpClient;if($c.ConnectAsync('127.0.0.1',8000).Wait(1500)){exit 0}else{exit 1}}catch{exit 1}" >nul 2>nul
if not errorlevel 1 (
  start "" "%ESQUEMA%://localhost:8000/"
  exit /b
)
if /i "%MODO%"=="LOCAL" (call INICIAR_WEB.bat) else (call INICIAR_WEB_RED.bat)
