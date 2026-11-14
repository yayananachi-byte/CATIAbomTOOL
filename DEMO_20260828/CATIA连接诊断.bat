@echo off
if not defined ZYR_CATIA_CHECK_CONSOLE (
    set "ZYR_CATIA_CHECK_CONSOLE=1"
    start "ZYR CATIA Connection Check" /wait cmd.exe /k call "%~f0"
    exit /b
)
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"
call :find_python
if not defined ZYR_PYTHON_EXE (
    echo Python 未找到，无法执行 CATIA 连接诊断。
    echo 请安装 64 位 Python 3.9 或更高版本，并在安装时勾选 Add Python to PATH。
    goto :finish
)
"%ZYR_PYTHON_EXE%" %ZYR_PYTHON_ARGS% catia_connection_check.py
:finish
echo.
echo 请保留本窗口内容或截图后发给 ZYR。
pause >nul
exit /b

:find_python
set "ZYR_PYTHON_EXE="
set "ZYR_PYTHON_ARGS="
where python.exe >nul 2>nul && set "ZYR_PYTHON_EXE=python"
if defined ZYR_PYTHON_EXE exit /b
where py.exe >nul 2>nul && set "ZYR_PYTHON_EXE=py" && set "ZYR_PYTHON_ARGS=-3"
if defined ZYR_PYTHON_EXE exit /b
for %%P in ("C:\Python39\python.exe" "C:\Python310\python.exe" "%LocalAppData%\Programs\Python\Python39\python.exe" "%LocalAppData%\Programs\Python\Python310\python.exe" "%LocalAppData%\Programs\Python\Python311\python.exe" "%LocalAppData%\Programs\Python\Python312\python.exe") do (
    if exist "%%~fP" set "ZYR_PYTHON_EXE=%%~fP"
    if defined ZYR_PYTHON_EXE exit /b
)
exit /b
