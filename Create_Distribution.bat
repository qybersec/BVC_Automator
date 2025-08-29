@echo off
echo ========================================
echo    Creating TMS Processor Distribution
echo ========================================
echo.

echo Creating distribution folder...
if exist "TMS_Processor_Complete" rmdir /s /q "TMS_Processor_Complete"
mkdir "TMS_Processor_Complete"

echo Copying core files...
copy "tms_processor.py" "TMS_Processor_Complete\"
copy "run_tms_processor.py" "TMS_Processor_Complete\"
copy "requirements.txt" "TMS_Processor_Complete\"

echo Copying launcher files...
copy "Run TMS Processor.bat" "TMS_Processor_Complete\"
copy "Install_Requirements.bat" "TMS_Processor_Complete\"

echo Copying documentation...
copy "README.md" "TMS_Processor_Complete\"
copy "INSTALLATION_GUIDE.md" "TMS_Processor_Complete\"

echo Copying test files (for verification)...
copy "test_processor.py" "TMS_Processor_Complete\"
copy "test_improvements.py" "TMS_Processor_Complete\"

echo.
echo Distribution package created successfully!
echo.
echo Files included:
echo - tms_processor.py (main processor)
echo - run_tms_processor.py (GUI launcher)
echo - Run TMS Processor.bat (easy launcher)
echo - Install_Requirements.bat (requirements installer)
echo - requirements.txt (dependencies)
echo - README.md (comprehensive guide)
echo - INSTALLATION_GUIDE.md (user-friendly guide)
echo - test files (for verification)
echo.
echo The "TMS_Processor_Complete" folder is ready for distribution!
echo.
pause
