@echo off
chcp 65001 >nul

REM 仮想環境のパスを設定（必要に応じて変更してください）
set "VENV_PATH=.\venv"

REM 実行するPythonモジュールの名前を設定（必要に応じて変更してください）
set "MODULE_NAME=src.main"

REM 仮想環境が存在するか確認
if exist "%VENV_PATH%\Scripts\activate.bat" (
    echo [INFO] 仮想環境をアクティブ化しています...
    call "%VENV_PATH%\Scripts\activate.bat"
) else (
    echo [ERROR] 仮想環境が見つかりません: %VENV_PATH%
    echo 仮想環境を作成するか、正しいパスを設定してください。
    pause
    exit /b 1
)

REM モジュールを実行
echo [INFO] モジュールを実行しています: %MODULE_NAME%
python -m %MODULE_NAME%
if errorlevel 1 (
    echo [ERROR] モジュールの実行中にエラーが発生しました。
    pause
    exit /b 1
)

REM 仮想環境をディアクティブ化
echo [INFO] 仮想環境をディアクティブ化しています...
deactivate

echo [INFO] 実行が完了しました。
pause
