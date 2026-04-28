@echo off
REM Payment Verification Automation - Windows Batch Script
REM This script runs the payment verification for all store folders

echo ========================================
echo Payment Verification Automation
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

REM Run the verification script
echo Starting verification process...
echo.
python payment_verification.py

echo.
echo ========================================
echo Verification Complete!
echo ========================================
echo.
echo Check the individual folder reports and the overall summary file.
echo.

pause
