# PORTERSシステムテスト

このディレクトリには、PORTERSシステムとの連携をテストするためのスクリプトが含まれています。

## テスト内容

1. **ログインテスト**
   - PORTERSシステムへのログイン処理をテスト
   - 二重ログインポップアップの処理も含む
   - ログイン後の画面検証

## 前提条件

テストを実行する前に、以下の準備が必要です。

1. **環境変数の設定**
   - `config/secrets.env` ファイルに以下の環境変数を設定してください。
     ```
     ADMIN_URL=https://example.com/login
     ADMIN_ID=会社ID
     LOGIN_ID=ユーザーID
     LOGIN_PASSWORD=パスワード
     ```

2. **セレクタ情報の設定**
   - `config/selectors.csv` ファイルにPORTERSシステムの要素セレクタを設定してください。
   - 基本的なセレクタは既に設定されていますが、PORTERSシステムの仕様変更があった場合は更新が必要です。

3. **必要なパッケージのインストール**
   - 以下のコマンドで必要なパッケージをインストールしてください。
     ```
     pip install -r requirements.txt
     ```

## テスト実行方法

### ログインテスト

```bash
python -m tests.test_main
```

### オプション

- `--headless`: ヘッドレスモードで実行（画面表示なし）
  ```bash
  python -m tests.test_main --headless
  ```

- `--skip-login`: ログイン処理をスキップ
  ```bash
  python -m tests.test_main --skip-login
  ```

## テスト結果

テスト実行時に以下のファイルが生成されます。

1. **ログファイル**
   - `logs/app_YYYYMMDD.log`: 日付ごとのログファイル

2. **スクリーンショット**
   - `logs/screenshots/YYYYMMDD_HHMMSS/`: テスト実行時のスクリーンショット
   - 主なスクリーンショット:
     - `login_before.png`: ログイン前
     - `login_input.png`: ログイン情報入力後
     - `login_after.png`: ログイン後
     - `double_login_popup.png`: 二重ログインポップアップ（表示された場合）
     - `login_success_verification.png`: ログイン成功確認
     - `before_logout.png`: ログアウト前
     - `after_logout.png`: ログアウト後

## トラブルシューティング

1. **セレクタが見つからない場合**
   - PORTERSシステムの仕様変更により、セレクタが変更された可能性があります。
   - `config/selectors.csv` ファイルを更新してください。

2. **ログインに失敗する場合**
   - 環境変数の設定を確認してください。
   - ログファイルとスクリーンショットを確認して、エラーの原因を特定してください。
   - `logs/screenshots/YYYYMMDD_HHMMSS/login_failed.html` ファイルにログイン失敗時のHTMLが保存されています。

3. **二重ログインポップアップの処理に失敗する場合**
   - ポップアップのセレクタが変更された可能性があります。
   - `tests/test_login.py` の `_handle_double_login_popup` メソッドを確認してください。 