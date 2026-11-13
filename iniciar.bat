@echo off
cd /d "%~dp0"

set "PY="
py -3 -c "import openpyxl" >nul 2>nul
if not errorlevel 1 set "PY=py -3"
if defined PY goto run

python -c "import openpyxl" >nul 2>nul
if not errorlevel 1 set "PY=python"
if defined PY goto run

if exist "%USERPROFILE%\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe" set "PY=%USERPROFILE%\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe"
if defined PY goto run

echo.
echo  [ERROR] No se encontro un Python con openpyxl.
echo  Ejecuta en una consola:  py -m pip install openpyxl
echo  y vuelve a intentarlo.
echo.
pause
exit /b 1

:run
echo  Arrancando Panel de Gastos 2026...
echo  (cierra esta ventana para detener)
echo.
%PY% servidor.py
pause