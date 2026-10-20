@echo off
setlocal EnableExtensions
set "APP_DIR=%~dp0"
set "BOOTSTRAP_EXE="
set "BOOTSTRAP_ARGS="

for %%P in (
  "%LocalAppData%\Programs\Python\Python313\python.exe"
  "%LocalAppData%\Programs\Python\Python312\python.exe"
  "%LocalAppData%\Programs\Python\Python311\python.exe"
  "%LocalAppData%\Programs\Python\Python310\python.exe"
  "%ProgramFiles%\Python313\python.exe"
  "%ProgramFiles%\Python312\python.exe"
  "%ProgramFiles%\Python311\python.exe"
  "%ProgramFiles%\Python310\python.exe"
) do (
  if not defined BOOTSTRAP_EXE if exist "%%~fP" set "BOOTSTRAP_EXE=%%~fP"
)
if defined BOOTSTRAP_EXE goto python_found

for %%V in (3.13 3.12 3.11 3.10) do (
  py.exe -%%V -c "import sys" >nul 2>nul
  if not errorlevel 1 (
    set "BOOTSTRAP_EXE=py.exe"
    set "BOOTSTRAP_ARGS=-%%V"
    goto python_found
  )
)

echo.
echo A standard CPython 3.10+ installation was not found.
echo Install 64-bit Python from python.org with Add Python to PATH enabled.
echo Do not use the C:\Espressif Python environment for this application.
echo.
pause
exit /b 1

:python_found
echo.
echo Creating isolated application runtime...
"%BOOTSTRAP_EXE%" %BOOTSTRAP_ARGS% -c "import struct,sys; sys.exit(0 if struct.calcsize('P') == 8 else 1)" >nul 2>nul
if errorlevel 1 goto python_not_64bit

"%BOOTSTRAP_EXE%" %BOOTSTRAP_ARGS% -m venv "%APP_DIR%.venv"
if errorlevel 1 goto setup_failed

"%APP_DIR%.venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto setup_failed

"%APP_DIR%.venv\Scripts\python.exe" -m pip install -r "%APP_DIR%requirements.txt"
if errorlevel 1 goto setup_failed

echo.
echo Setup complete. Double-click START_GUI.cmd to launch the application.
echo.
pause
exit /b 0

:setup_failed
echo.
echo Setup failed. Keep this window open and send a screenshot of the error.
echo.
pause
exit /b 1

:python_not_64bit
echo.
echo A 64-bit CPython 3.10+ installation is required for 64-bit CATIA.
echo Install 64-bit Python from python.org, then run this setup again.
echo.
pause
exit /b 1
