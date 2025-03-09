
# Project Template

## 仕様
このプロジェクトテンプレートは、Python プロジェクトの開発、テスト、および仕様書生成を効率化するために作成されました。以下の主な機能を提供します。

- 仮想環境のセットアップ (`run.bat`)。
- デフォルトのプロジェクト構造と基本ファイル。
- 仕様書生成ツールのテンプレートと実行スクリプト (`spec_tools_run.bat`)。
- 環境ごとに設定可能な `settings.ini` と秘密情報の管理ファイル `secrets.env`。

---

## フォルダ構成
プロジェクトのフォルダ構成は以下の通りです。

```
%PROJECT_NAME%/
├── src/               # ソースコード
│   ├── main.py        # メインスクリプト
│   ├── utils/         # ユーティリティスクリプト
│   │   ├── config.py  # 設定ファイルの読み込み
│   │   └── logging_config.py # ログ設定
│   └── modules/       # モジュールディレクトリ
│       └── module1.py # サンプルモジュール
├── tests/             # テスト用コード
├── data/              # データ保存用ディレクトリ
├── logs/              # 実行時のログファイル保存先
├── config/            # 設定ファイル
│   ├── settings.ini   # 環境ごとの設定
│   └── secrets.env    # APIキーなどの秘密情報
├── docs/              # ドキュメント用ディレクトリ
├── spec_tools/        # 仕様書生成ツール
│   ├── merge_files.py         # 複数ファイルを統合するスクリプト
│   ├── generate_spec.py        # 仕様書生成スクリプト
│   ├── generate_detailed_spec.py # 詳細仕様書生成スクリプト
│   └── logs/          # 仕様書生成のログ
├── requirements.txt   # 必要なパッケージ
├── run.bat            # プロジェクト実行用スクリプト
└── spec_tools_run.bat # 仕様書生成用スクリプト
```

---

## 実行方法

### 1. 仮想環境の作成と有効化
初回実行時には仮想環境を作成します。以下のコマンドを使用してください。

```bash
run.bat
```

- 仮想環境が存在しない場合、自動的に作成されます。
- 必要なパッケージは `requirements.txt` から自動インストールされます。

仮想環境を手動で有効化する場合：
```bash
.\env\Scripts ctivate
```

---

### 2. メインスクリプトの実行
デフォルトでは `src\main.py` が実行されます。他のスクリプトを指定する場合は、引数にスクリプトパスを渡します。

```bash
run.bat src\your_script.py
```

環境を指定する場合、`--env` オプションを使用します（例: `development`, `production`, `test`）。

```bash
run.bat --env production
```

---

### 3. 仕様書生成ツールの使用
仕様書生成スクリプトは `spec_tools_run.bat` を使用して実行できます。

- **merge_files.py の実行**:
  ```bash
  spec_tools_run.bat --merge
  ```

- **仕様書生成**:
  ```bash
  spec_tools_run.bat --spec
  ```

- **詳細仕様書生成**:
  ```bash
  spec_tools_run.bat --detailed-spec
  ```

- **すべてを一括実行**:
  ```bash
  spec_tools_run.bat --all
  ```

---

### 4. テストモード
テストモードでスクリプトを実行するには、`--test` オプションを使用します。

```bash
run.bat --test
```

---

## 注意事項

1. **仮想環境の存在確認**:
   `run.bat` または `spec_tools_run.bat` を初回実行時に仮想環境が作成されます。既存の仮想環境を削除する場合、手動で `.\env` を削除してください。

2. **環境変数の設定**:
   APIキーなどの秘密情報は `config\secrets.env` に格納し、共有しないよう注意してください。

3. **パッケージのアップデート**:
   必要に応じて、`requirements.txt` を更新してください。更新後、`run.bat` を実行すると自動的にインストールされます。

---

## サポート情報

- **開発者**: [Your Name or Team Name]
- **連絡先**: example@domain.com
- **最終更新日**: [YYYY-MM-DD]
