@echo off
setlocal
set "APP_DIR=%~dp0"
set "GUI_PYTHON=%APP_DIR%.venv\Scripts\python.exe"

if not exist "%GUI_PYTHON%" goto runtime_missing
"%GUI_PYTHON%" "%APP_DIR%catia_bom_gui.py"
pause
exit /b

:runtime_missing
echo.
echo CATIA BOM Manager runtime is not configured.
echo Run SETUP_GUI_RUNTIME.cmd once, then start this file again.
echo.
pause
