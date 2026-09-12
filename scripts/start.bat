@echo off
setlocal
chcp 65001 >nul
set "PYTHONIOENCODING=utf-8"
cd /d "%~dp0.."
if not exist ".venv\Scripts\python.exe" goto missing
".venv\Scripts\python.exe" -c "import flask, openpyxl, basyx.aas, aas_core3" >nul 2>&1
if errorlevel 1 goto missing
".venv\Scripts\python.exe" -B "scripts\app.py" %*
if errorlevel 1 goto failed
exit /b 0
:missing
echo Python environment or dependencies are missing. Run setup.bat in the AAS directory first.
pause
exit /b 1
:failed
echo AAS could not start. See the error above.
pause
exit /b 1
