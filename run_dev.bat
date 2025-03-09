@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

rem Initialize environment variables
set "VENV_PATH=.\venv"
set "PYTHON_CMD=python"
set "PIP_CMD=pip"
set "DEFAULT_SCRIPT=src.main"
set "APP_ENV=development"
set "SCRIPT_TO_RUN=%DEFAULT_SCRIPT%"
set "SCRIPT_ARGS="
set "INTERACTIVE=true"

rem Parse command line arguments
:parse_args
if "%~1"=="-m" (
    set "SCRIPT_TO_RUN=src.%~2"
    shift
    shift
    set "INTERACTIVE=false"
    goto parse_args
)
if "%~1"=="--block" (
    set "SCRIPT_ARGS=!SCRIPT_ARGS! --block %~2"
    shift
    shift
    set "INTERACTIVE=false"
    goto parse_args
)
if "%~1"=="--ids" (
    set "SCRIPT_ARGS=!SCRIPT_ARGS! --ids %~2"
    shift
    shift
    goto parse_args
)
if "%~1"=="--all" (
    set "SCRIPT_ARGS=!SCRIPT_ARGS! --all"
    shift
    set "INTERACTIVE=false"
    goto parse_args
)
if "%~1"=="--env" (
    set "APP_ENV=%~2"
    shift
    shift
    goto parse_args
)

rem Display help message if --help is provided
if "%~1"=="--help" (
    echo Usage:
    echo   run_dev.bat [options]
    echo.
    echo Options:
    echo   --env [dev^|prd]    : Specify environment
    echo                         ^(dev=development, prd=production^)
    echo   --block [1^|2^|3]    : Specify process block to run
    echo                         ^(1=Check consult flags, 2=Get survey data, 3=Import to PORTERS^)
    echo   --ids ID1,ID2,...   : Specify target IDs
    echo   --all               : Run all process blocks
    echo   --help              : Show this help
    echo.
    echo Environment modes:
    echo   dev  : Development environment with detailed logs
    echo   prd  : Production environment for stable operation
    echo.
    echo Examples:
    echo   run_dev.bat --env dev
    echo   run_dev.bat --env prd
    echo   run_dev.bat --block 1
    echo   run_dev.bat --block 2 --ids 123456789,987654321
    echo   run_dev.bat --all
    set "INTERACTIVE=false"
    goto END
)

if "%~1" neq "" (
    echo Unknown argument: %~1
    echo Use --help to show usage
    exit /b 1
)

rem If no arguments are provided or interactive mode is enabled, prompt the user
if "%INTERACTIVE%"=="true" (
    echo Select environment:
    echo   1. Development (dev)
    echo   2. Production (prd)
    set /p "CHOICE=Enter your choice (1/2): "
    if "!CHOICE!"=="1" set "APP_ENV=development"
    if "!CHOICE!"=="2" set "APP_ENV=production"
    if not defined APP_ENV (
        echo Error: Invalid choice. Please run again.
        exit /b 1
    )
    
    echo.
    echo Select process to run:
    echo   1. Check consult flags and extract new IDs
    echo   2. Get survey data and export to CSV
    echo   3. Import data to PORTERS
    echo   4. Run all processes
    echo.
    
    set /p "PROCESS=Enter your choice (1/2/3/4): "
    
    if "!PROCESS!"=="1" set "SCRIPT_ARGS=--block 1"
    if "!PROCESS!"=="2" (
        set /p "IDS=Enter target IDs (comma-separated, optional): "
        if not "!IDS!"=="" (
            set "SCRIPT_ARGS=--block 2 --ids !IDS!"
        ) else (
            set "SCRIPT_ARGS=--block 2"
        )
    )
    if "!PROCESS!"=="3" (
        set /p "IDS=Enter target IDs (comma-separated, optional): "
        if not "!IDS!"=="" (
            set "SCRIPT_ARGS=--block 3 --ids !IDS!"
        ) else (
            set "SCRIPT_ARGS=--block 3"
        )
    )
    if "!PROCESS!"=="4" set "SCRIPT_ARGS=--all"
    
    if not defined SCRIPT_ARGS (
        echo Error: Invalid choice. Please run again.
        exit /b 1
    )
)

rem Check if Python is installed
%PYTHON_CMD% --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python is not installed or not in PATH.
    pause
    exit /b 1
)

rem Create virtual environment if it doesn't exist
if not exist "%VENV_PATH%\Scripts\activate.bat" (
    echo [LOG] Virtual environment does not exist. Creating...
    %PYTHON_CMD% -m venv "%VENV_PATH%"
    if errorlevel 1 (
        echo Error: Failed to create virtual environment.
        pause
        exit /b 1
    )
    echo [LOG] Virtual environment created successfully.
)

rem Activate virtual environment
if exist "%VENV_PATH%\Scripts\activate" (
    call "%VENV_PATH%\Scripts\activate"
) else (
    echo Error: Failed to activate virtual environment. Activate script not found.
    pause
    exit /b 1
)

rem Check requirements.txt
if not exist requirements.txt (
    echo Error: requirements.txt not found.
    pause
    exit /b 1
)

rem Install requirements if needed
for /f "skip=1 delims=" %%a in ('certutil -hashfile requirements.txt SHA256') do if not defined CURRENT_HASH set "CURRENT_HASH=%%a"

if exist .req_hash (
    set /p STORED_HASH=<.req_hash
) else (
    set "STORED_HASH="
)

if not "%CURRENT_HASH%"=="%STORED_HASH%" (
    echo [LOG] Installing required packages...
    %PIP_CMD% install -r requirements.txt
    if errorlevel 1 (
        echo Error: Failed to install packages.
        pause
        exit /b 1
    )
    echo %CURRENT_HASH%>.req_hash
)

rem Run the script
echo [LOG] Environment: %APP_ENV%
echo [LOG] Running script: %SCRIPT_TO_RUN%
if defined SCRIPT_ARGS echo [LOG] Arguments: %SCRIPT_ARGS%

%PYTHON_CMD% -m %SCRIPT_TO_RUN% --env %APP_ENV% %SCRIPT_ARGS%
if errorlevel 1 (
    echo Error: Script execution failed.
    pause
    exit /b 1
)

:END
endlocal
