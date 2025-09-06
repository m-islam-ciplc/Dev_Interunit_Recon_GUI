@echo off
echo Starting Interunit Loan Matcher GUI...
echo.

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python is not installed or not in PATH
    echo Please install Python 3.8+ and try again
    pause
    exit /b 1
)

REM Check if required packages are installed
python -c "import PyQt6" >nul 2>&1
if errorlevel 1 (
    echo Installing required packages...
    pip install -r requirements_gui.txt
    if errorlevel 1 (
        echo Error: Failed to install required packages
        pause
        exit /b 1
    )
)

REM Run the GUI application
python main_gui.py

if errorlevel 1 (
    echo.
    echo An error occurred while running the GUI
    pause
)
