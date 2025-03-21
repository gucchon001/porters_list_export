# PORTERSリストエクスポート

## プロジェクトの概要
このプロジェクトは、PORTERSシステムから「求職者情報」と「選考プロセス情報」という2種類のデータを取得し、所定のGoogleスプレッドシートへ転記するツールです。また、取得したデータを集計し、別シートへ集計結果を出力する機能も備えています。

### 主な機能
- PORTERSシステムへの自動ログイン（二重ログイン回避機能付き）
- 求職者一覧データの自動取得・CSVエクスポート
- 選考プロセス一覧データの自動取得・CSVエクスポート
- 取得したCSVデータのGoogleスプレッドシートへの自動転記
- 求職者データのフェーズ別集計処理
- 選考プロセスデータの日次集計処理

## バッチファイルの実行方法

### 開発環境用実行ファイル（run_dev.bat）
開発やテスト目的で使用します。実行したいパラメータを手動で指定できます。

```bash
# 基本実行（求職者一覧のみ取得）
run_dev.bat

# 求職者一覧と選考プロセスの両方を取得
run_dev.bat --process both

# 求職者一覧取得後に選考プロセスを取得（連続実行）
run_dev.bat --process sequential

# データ取得と集計を一度に行う
run_dev.bat --process sequential --aggregate both

# ヘッドレスモード（ブラウザを表示せず）で実行
run_dev.bat --headless

# 集計処理のみを実行（データ取得をスキップ）
run_dev.bat --skip-operations --aggregate users
```

### 本番環境用実行ファイル（run.bat）
本番環境での定期実行用に最適化されています。すべてのデータ取得と集計処理を一度に実行します。

```bash
# すべての処理を自動実行（求職者一覧・選考プロセス取得および集計）
run.bat
```

run.batは内部で以下のコマンドを実行しています：
```
python -m src.main --process sequential --aggregate both
```

## コマンドラインパラメータの説明

### 処理フロー指定 (--process)
- `candidates`: 求職者一覧のみエクスポート（デフォルト）
- `entryprocess`: 選考プロセス一覧のみエクスポート
- `both`: 両方を順に実行（別々のセッションで処理）
- `sequential`: 求職者一覧取得後に選考プロセスも取得（同一セッションで連続処理）

### 集計処理指定 (--aggregate)
- `none`: 集計処理を実行しない（デフォルト）
- `users`: 求職者フェーズ別集計を実行
- `entryprocess`: 選考プロセス集計を実行
- `both`: 両方の集計処理を実行

注意: デフォルトでは、集計処理の前にデータ取得処理も実行されます。集計処理のみを実行するには、`--skip-operations` フラグを追加する必要があります。例：
```
python -m src.main --skip-operations --aggregate users
```

### その他のオプション
- `--headless`: ブラウザを表示せずにヘッドレスモードで実行
- `--env [development|production]`: 実行環境の指定（設定ファイルの分岐用）
- `--skip-operations`: ログイン処理とデータ取得をスキップし、集計処理のみを実行
- `--log-level [DEBUG|INFO|WARNING|ERROR|CRITICAL]`: ログレベルの指定

## 集計処理について

### 求職者フェーズ別集計
求職者一覧データからフェーズごとの人数をカウントし、COUNT_USERSシートに記録します。
集計されるフェーズ：
- 相談前×推薦前(新規エントリー)
- 相談前×推薦前(open)
- 推薦済(仮エントリー)
- 面談設定済
- 終了

また、登録経路ごとの内訳も集計されます：
- LINE
- 自社サイト(応募後架電)
- 自社サイト(ダイレクトコミュニケーション)

### 選考プロセス集計
選考プロセスデータを日付ごとにLIST_ENTRYPROCESSシートに記録します。
記録される項目：
- 求職者ID
- 性名/名前
- 企業コード
- 企業名
- 選考プロセス
- 担当CA

## システム要件
- Python 3.8以上
- Google APIアクセス用のサービスアカウント（JSONファイル）
- インターネット接続環境
- ChromeブラウザがインストールされたWindows環境

## 設定ファイル
- `config/settings.ini`: シート名やURLなどの設定
- `config/secrets.env`: PORTERSログイン情報やAPI認証情報
- `config/selectors.csv`: ブラウザ操作に使用するHTML要素のセレクタ定義
