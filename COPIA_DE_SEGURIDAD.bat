@echo off
REM Copia la carpeta "trabajo" (estado, PDF, Excel, config) a "respaldos\fecha_hora". Se puede programar a diario.
cd /d "%~dp0"
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HHmm"') do set FECHA=%%i
robocopy "trabajo" "respaldos\%FECHA%" /E /NFL /NDL /NJH /NJS >nul
echo Copia hecha en respaldos\%FECHA%
