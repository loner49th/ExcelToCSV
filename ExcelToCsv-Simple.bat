@echo off
echo Excel to CSV Converter
echo ======================

if "%~1"=="" (
    echo Please drag and drop an Excel file to this batch file
    pause
    exit /b 1
)

echo Processing file: %~1
echo.

powershell.exe -ExecutionPolicy Bypass -File "%~dp0Convert-ExcelToCsv.ps1" -ExcelFilePath "%~1"

echo.
echo Exit code: %ERRORLEVEL%
echo Press any key to continue...
pause > nul