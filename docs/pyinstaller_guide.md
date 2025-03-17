# PyInstallerによる実行ファイル生成指示書

## 概要

この指示書は、Pythonプロジェクトを実行ファイル（.exe）に変換するための手順を説明します。PyInstallerを使用して、依存ライブラリや設定ファイルを含む単一の実行ファイルを生成します。

## 前提条件

- Python 3.8以上がインストールされていること
- pip（Pythonパッケージマネージャ）が利用可能であること
- 対象のPythonプロジェクトが正常に動作すること

## 手順

### 1. 必要なライブラリのインストール

```bash
# PyInstallerのインストール
pip install pyinstaller

# プロジェクトの依存ライブラリをインストール（例）
pip install -r requirements.txt
```

### 2. プロジェクト構成の確認

以下のような標準的なプロジェクト構成を想定しています：

```
project_root/
├── config/                  # 設定ファイル
│   ├── settings.ini
│   ├── secrets.env          # 環境変数ファイル
│   └── その他の設定ファイル
├── src/                     # ソースコード
│   ├── main.py              # エントリーポイント
│   ├── utils/
│   │   ├── environment.py   # 環境設定関連
│   │   └── logging_config.py
│   └── modules/             # 機能モジュール
│       ├── module1.py
│       ├── module2.py
│       └── ...
└── setup.py                 # セットアップファイル（オプション）
```

### 3. specファイルの作成

以下のコマンドでspecファイルの雛形を生成します：

```bash
pyinstaller --name your_app_name src/main.py
```

### 4. specファイルの編集

生成されたspecファイル（`your_app_name.spec`）を以下のテンプレートを参考に編集します：

```python
# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

block_cipher = None

# 必要なデータファイルを定義
added_files = [
    ('config/settings.ini', 'config'),
    ('config/secrets.env', 'config'),
    # 他の必要なファイルを追加
]

a = Analysis(
    ['src/main.py'],  # エントリーポイントのスクリプト
    pathex=[str(Path('.').absolute())],
    binaries=[],
    datas=added_files,
    hiddenimports=[
        # プロジェクトで使用している主要なライブラリ
        'pandas',
        'requests',
        'dotenv',
        'python-dotenv',
        # プロジェクト内のモジュール
        'src.utils.environment',
        'src.utils.logging_config',
        'src.modules.module1',
        'src.modules.module2',
        # 動的にインポートされるモジュール
        'configparser',
        # 他の必要なモジュールを追加
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(
    a.pure, 
    a.zipped_data,
    cipher=block_cipher
)

# --onefile モードで実行ファイルを作成
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='your_app_name',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None
)
```

### 5. 実行ファイルの生成

編集したspecファイルを使用して実行ファイルを生成します：

```bash
pyinstaller your_app_name.spec
```

### 6. 生成された実行ファイルの確認

生成された実行ファイルは `dist` ディレクトリに配置されます：

```
dist/
└── your_app_name.exe
```

### 7. 実行ファイルのテスト

生成された実行ファイルが正しく動作するか確認します：

```bash
cd dist
./your_app_name.exe
```

## トラブルシューティング

### 一般的な問題と解決策

1. **モジュールが見つからないエラー**
   - `hiddenimports` リストに不足しているモジュールを追加します
   - 動的にインポートされるモジュールも明示的に追加する必要があります

2. **ファイルが見つからないエラー**
   - `added_files` リストに不足しているファイルを追加します
   - パスの解決方法を確認します

3. **環境変数が読み込まれないエラー**
   - 環境変数ファイルが正しく含まれているか確認します
   - `added_files` リストに環境変数ファイルが含まれているか確認します

4. **実行時エラー**
   - コンソール出力を確認してエラーメッセージを特定します
   - デバッグモードで実行ファイルを生成して詳細なログを確認します：
     ```python
     exe = EXE(..., debug=True, ...)
     ```

## 注意事項

- 機密情報（APIキーなど）を含む環境変数ファイルを実行ファイルに含める場合は、セキュリティに注意してください
- 大きなライブラリを含む場合、実行ファイルのサイズが大きくなる可能性があります
- 一部のライブラリはPyInstallerとの互換性に問題がある場合があります
- 標準的な環境変数の読み込み方法（`dotenv`ライブラリなど）を使用している場合、追加の修正は不要なことが多いです

## 参考リンク

- [PyInstaller公式ドキュメント](https://pyinstaller.org/en/stable/)
- [PyInstallerのトラブルシューティングガイド](https://pyinstaller.org/en/stable/when-things-go-wrong.html)