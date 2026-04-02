@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "SCRIPT_PATH=%SCRIPT_DIR%md2word.py"

if defined PYTHON_EXE (
    "%PYTHON_EXE%" "%SCRIPT_PATH%" %*
    exit /b %ERRORLEVEL%
)

where py >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    py -3 "%SCRIPT_PATH%" %*
    exit /b %ERRORLEVEL%
)

where python >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    python "%SCRIPT_PATH%" %*
    exit /b %ERRORLEVEL%
)

where python3 >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    python3 "%SCRIPT_PATH%" %*
    exit /b %ERRORLEVEL%
)

echo [error] Python 3 not found. Please install Python and make py -3 or python available. 1>&2
exit /b 1
