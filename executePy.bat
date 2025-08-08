@echo off
:: Excel Table Merger Batch Script
:: This script runs the mergeTable.py Python script

echo ======================================================
echo           Excel Table Merger
echo ======================================================
echo Starting merge process...
echo.

:: Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python and try again
    pause
    exit /b 1
)

:: Check if mergeTable.py exists
if not exist "mergeTable.py" (
    echo ERROR: mergeTable.py not found in current directory
    echo Please make sure mergeTable.py is in the same folder as this bat file
    pause
    exit /b 1
)

:: Run the Python script
echo Running mergeTable.py...
echo.
python mergeTable.py

:: Check if the script ran successfully
if errorlevel 1 (
    echo.
    echo ERROR: Script execution failed
    echo Check the error messages above
) else (
    echo.
    echo ======================================================
    echo           Merge Process Completed!
    echo ======================================================
    echo Check the logs folder for detailed information
    echo Output file: 1.xlsx
)

echo.
pause