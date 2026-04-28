@echo off
REM Master Payment Verification Automation - Windows Batch Script
REM This script extracts zip files and runs payment verification

echo ========================================
echo Payment Verification Master Automation
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.7 or higher from https://www.python.org/
    pause
    exit /b 1
)

echo Python found!
echo.

REM Check if required packages are installed
echo Checking dependencies...
python -c "import pandas" >nul 2>&1
if errorlevel 1 (
    echo.
    echo Installing required packages...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo ERROR: Failed to install required packages
        pause
        exit /b 1
    )
)

echo Dependencies OK!
echo.

REM Run the master automation script
echo Starting master automation process...
echo.
python master_automation.py

if errorlevel 1 (
    echo.
    echo ========================================
    echo Automation Failed!
    echo ========================================
    echo Please check the error messages above.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Automation Complete!
echo ========================================
echo.
echo All steps completed successfully!
echo Check the individual folder reports and the overall summary file.
echo.

pause
