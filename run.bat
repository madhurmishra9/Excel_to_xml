@echo off
REM Set up dependencies in a local virtual environment and launch the menu.
setlocal
cd /d "%~dp0"

if not exist ".venv" (
    echo Creating virtual environment ...
    python -m venv .venv
    if errorlevel 1 (
        echo Failed to create virtual environment. Is Python installed and on PATH?
        exit /b 1
    )
)

call ".venv\Scripts\activate.bat"

echo Installing dependencies ...
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo Failed to install dependencies.
    exit /b 1
)

echo.
echo Launching menu ...
python menu.py

endlocal
