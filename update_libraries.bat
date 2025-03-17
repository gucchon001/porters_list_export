@echo off
chcp 65001 >nul
setlocal

echo ===== Python Library Update Tool =====
echo.

:: Check if Python is available
python --version > nul 2>&1
if errorlevel 1 (
    echo Error: Python not found.
    echo Please ensure Python is installed.
    goto END
)

:: Check for virtual environment
if not exist "venv" (
    echo Virtual environment not found. Creating a new one...
    python -m venv venv
    if errorlevel 1 (
        echo Failed to create virtual environment.
        goto END
    )
    echo Virtual environment created successfully.
)

:: Activate virtual environment
call venv\Scripts\activate.bat
if errorlevel 1 (
    echo Failed to activate virtual environment.
    goto END
)

echo Virtual environment activated successfully.

:: Upgrade pip
echo Upgrading pip...
python -m pip install --upgrade pip
if errorlevel 1 (
    echo Failed to upgrade pip.
    goto DEACTIVATE
)

:: Check for requirements.txt
if not exist "requirements.txt" (
    echo requirements.txt not found.
    goto DEACTIVATE
)

:: Update libraries
echo.
echo Updating libraries...
pip install -r requirements.txt --upgrade
if errorlevel 1 (
    echo Failed to update libraries.
    goto DEACTIVATE
)

echo.
echo Current installed package list:
pip list
echo.
echo Libraries updated successfully.

:DEACTIVATE
:: Deactivate virtual environment
call venv\Scripts\deactivate.bat

:END
echo.
echo Processing completed.
pause
endlocal 