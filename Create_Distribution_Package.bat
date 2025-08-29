@echo off
echo ========================================
echo   Creating Distribution Package...
echo ========================================
echo.
echo Creating TMS_Processor_Complete.zip...
echo.

powershell -Command "Compress-Archive -Path 'TMS_Processor_Complete\*' -DestinationPath 'TMS_Processor_Complete.zip' -Force"

echo.
echo Distribution package created: TMS_Processor_Complete.zip
echo.
echo This file contains everything needed for distribution:
echo - All program files
echo - Installation scripts  
echo - Documentation
echo - START HERE.txt guide
echo.
echo Ready to share with your boss or colleagues!
echo.
pause