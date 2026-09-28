@echo off
setlocal

REM Run the GUI from source on Windows.
REM Always work from the project root, regardless of the launch directory.
pushd "%~dp0"

echo ============================================================
echo  pdf2xlAI - Windows GUI
echo  Project: %CD%
echo ============================================================
echo.

REM Prefer the Python launcher only when Python 3 is actually installed.
py -3 -c "import sys" >nul 2>&1
if not errorlevel 1 (
    set "PY=py -3"
) else (
    python -c "import sys" >nul 2>&1
    if errorlevel 1 (
        echo ERROR: Python was not found.
        echo Install Python 3.10 or newer and enable Add Python to PATH.
        goto :FAIL
    )
    set "PY=python"
)

echo Python command: %PY%
echo.

if not exist ".venv\Scripts\python.exe" (
    echo [1/3] Creating virtual environment...
    %PY% -m venv .venv
    if errorlevel 1 goto :FAIL
) else (
    echo [1/3] Virtual environment already exists.
)
echo.

echo [2/3] Installing dependencies...
.venv\Scripts\python.exe -m pip install --upgrade pip
if errorlevel 1 goto :FAIL

if not exist "requirements.txt" (
    echo ERROR: requirements.txt was not found in %CD%.
    goto :FAIL
)
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto :FAIL
echo.

echo [3/3] Starting GUI...
if not exist "gui.py" (
    echo ERROR: gui.py was not found in %CD%.
    goto :FAIL
)
.venv\Scripts\python.exe gui.py
echo.
echo GUI closed.
popd
pause
exit /b 0

:FAIL
echo.
echo ------------------------------------------------------------
echo RUN FAILED
echo Copy the error output above if you need help.
echo ------------------------------------------------------------
echo.
popd
pause
exit /b 1
