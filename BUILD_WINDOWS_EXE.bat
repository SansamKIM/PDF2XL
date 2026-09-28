@echo off
setlocal

REM Build the Windows executable with PyInstaller.
REM Always work from the project root, regardless of the launch directory.
pushd "%~dp0"

echo ============================================================
echo  pdf2xlAI - Windows EXE build
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
    echo [1/4] Creating virtual environment...
    %PY% -m venv .venv
    if errorlevel 1 goto :FAIL
) else (
    echo [1/4] Virtual environment already exists.
)
echo.

echo [2/4] Upgrading pip...
.venv\Scripts\python.exe -m pip install --upgrade pip
if errorlevel 1 goto :FAIL
echo.

echo [3/4] Installing dependencies...
if not exist "requirements.txt" (
    echo ERROR: requirements.txt was not found in %CD%.
    goto :FAIL
)
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto :FAIL

if exist "requirements-dev.txt" (
    .venv\Scripts\python.exe -m pip install -r requirements-dev.txt
    if errorlevel 1 goto :FAIL
)
echo.

echo [4/4] Running PyInstaller...
if not exist "pdf2xlAI.spec" (
    echo ERROR: pdf2xlAI.spec was not found in %CD%.
    goto :FAIL
)
.venv\Scripts\python.exe -m PyInstaller pdf2xlAI.spec --noconfirm --clean
if errorlevel 1 goto :FAIL

echo.
echo ------------------------------------------------------------
echo BUILD SUCCEEDED
echo EXE: %CD%\dist\pdf2xlAI\pdf2xlAI.exe
echo ------------------------------------------------------------
echo.
popd
pause
exit /b 0

:FAIL
echo.
echo ------------------------------------------------------------
echo BUILD FAILED
echo Copy the error output above if you need help.
echo ------------------------------------------------------------
echo.
popd
pause
exit /b 1
