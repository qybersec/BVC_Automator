@echo off
echo ========================================
echo   BVC Automator - TMS Data Processor
echo ========================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found!
    echo Please install Python from https://python.org
    echo Make sure to check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
)

echo Python found. Installing dependencies...
echo.

:: Install required packages
pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo ERROR: Failed to install dependencies!
    echo Please check your internet connection and try again.
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo   Dependencies installed successfully!
echo   Starting BVC Automator...
echo ========================================
echo.

:: Run the application
python main.py

echo.
echo Application closed.
pause