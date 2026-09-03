@echo off
setlocal

set "SCRIPT_PATH=C:\CAMINHO\DO\arquivo.py"

for %%F in ("%SCRIPT_PATH%") do set "PROJETO_DIR=%%~dpF"

chcp 65001 >nul

set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

cd /d "%PROJETO_DIR%"
if errorlevel 1 exit /b 1

call ".venv\Scripts\activate.bat"
if errorlevel 1 exit /b 1

python "%SCRIPT_PATH%"
set "PYTHON_EXIT=%errorlevel%"

call deactivate

exit /b %PYTHON_EXIT%