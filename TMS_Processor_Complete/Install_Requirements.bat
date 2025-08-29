@echo off
echo ========================================
echo    TMS Processor - Requirements Installer
echo ========================================
echo.
echo This will install the required components...
echo.
pause

echo.
echo Checking Python installation...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo.
    echo Please install Python from: https://www.python.org/downloads/
    echo IMPORTANT: Check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
)

echo Python found! Installing required packages...
echo.
pip install -r requirements.txt

echo.
echo Installation complete!
echo.
echo If you see any errors above, please contact your IT department.
echo Otherwise, you can now run "Run TMS Processor.bat"
echo.
pause
