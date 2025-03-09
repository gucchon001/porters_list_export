@echo off
rem Ensure this file is saved with UTF-8 encoding
@echo off
setlocal enabledelayedexpansion

chcp 65001 > nul

:: スクリプト情報
set "TOOLS_DIR=spec_tools"
set "LOGS_DIR=%TOOLS_DIR%\logs"
set "DOCS_DIR=docs"

:: ヘルプメッセージ
if "%~1"=="" (
    echo 使用方法:
    echo   spec_tools_run.bat --merge         : merge_files.py を実行
    echo   spec_tools_run.bat --spec          : generate_spec.py を実行
    echo   spec_tools_run.bat --detailed-spec : generate_detailed_spec.py を実行
    echo   spec_tools_run.bat --all           : 全てを一括実行
    echo   spec_tools_run.bat --help          : このヘルプを表示
    goto END
)

:: 仮想環境のチェック
if not exist ".\venv\Scripts\activate.bat" (
    echo Error: 仮想環境が存在しません。run.bat で仮想環境を作成してください。
    goto END
)

call .\venv\Scripts\activate

:: コマンドの処理
if "%~1"=="--merge" (
    echo [LOG] merge_files.py を実行中...
    python %TOOLS_DIR%\merge_files.py > %LOGS_DIR%\merge_files.log 2>&1
    if errorlevel 1 (
        echo Error: merge_files.py の実行に失敗しました。ログを確認してください。
        goto END
    )
    echo [LOG] merge_files.py の実行が完了しました。
)

if "%~1"=="--spec" (
    echo [LOG] generate_spec.py を実行中...
    python %TOOLS_DIR%\generate_spec.py > %LOGS_DIR%\generate_spec.log 2>&1
    if errorlevel 1 (
        echo Error: generate_spec.py の実行に失敗しました。ログを確認してください。
        goto END
    )
    echo [LOG] generate_spec.py の実行が完了しました。
)

if "%~1"=="--detailed-spec" (
    echo [LOG] generate_detailed_spec.py を実行中...
    python %TOOLS_DIR%\generate_detailed_spec.py > %LOGS_DIR%\generate_detailed_spec.log 2>&1
    if errorlevel 1 (
        echo Error: generate_detailed_spec.py の実行に失敗しました。ログを確認してください。
        goto END
    )
    echo [LOG] generate_detailed_spec.py の実行が完了しました。
)

if "%~1"=="--all" (
    echo [LOG] 全てのスクリプトを一括実行中...
    call %0 --merge
    call %0 --spec
    call %0 --detailed-spec
    echo [LOG] 全てのスクリプトの実行が完了しました。
)

if "%~1"=="--help" (
    echo 使用方法:
    echo   spec_tools_run.bat --merge         : merge_files.py を実行
    echo   spec_tools_run.bat --spec          : generate_spec.py を実行
    echo   spec_tools_run.bat --detailed-spec : generate_detailed_spec.py を実行
    echo   spec_tools_run.bat --all           : 全てを一括実行
    echo   spec_tools_run.bat --help          : このヘルプを表示
)

:END
endlocal
