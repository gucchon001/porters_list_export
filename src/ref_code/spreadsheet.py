import pandas as pd
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from googleapiclient.discovery import build
from ..utils.environment import EnvironmentUtils as env
from ..utils.logging_config import get_logger
import time
from pathlib import Path

logger = get_logger(__name__)

class Spreadsheet:
    def __init__(self):
        """スプレッドシート操作を管理するクラス"""
        self.credentials = self._get_credentials()
        self.service = build('sheets', 'v4', credentials=self.credentials)
        self.spreadsheet_id = env.get_config_value('SPREADSHEET', 'SSID')
        
        # シート名の設定を読み込む
        self.friend_data_key = env.get_config_value('SHEET_NAMES', 'FRIEND_DATA')  # "友達リストDLデータ"
        self.anq_data_key = env.get_config_value('SHEET_NAMES', 'ANQ_DATA')  # "アンケートDLデータ"
        self.log_sheet_name = 'logsheet'  # ハードコーディング
        
        # スプレッドシートから実際のシート名を取得
        self.friend_sheet_name = None
        self.anq_sheet_name = None
        self._load_sheet_settings()
        
    def _get_credentials(self):
        """認証情報を取得"""
        try:
            SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
            service_account_file = env.get_env_var('SERVICE_ACCOUNT_FILE')
            credentials = service_account.Credentials.from_service_account_file(
                service_account_file, scopes=SCOPES)
            logger.info(f"✓ サービスアカウント認証情報を読み込みました: {service_account_file}")
            return credentials
        except Exception as e:
            logger.error(f"認証情報の取得に失敗: {str(e)}")
            raise

    def _load_sheet_settings(self):
        """settingsシートから各シート名を読み込む"""
        try:
            # settingsシートから設定を読み込む
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range='settings!A:B'  # A列とB列を読み込む
            ).execute()
            
            values = result.get('values', [])
            if not values:
                logger.error("❌ settingsシートが空です")
                raise ValueError("Settings sheet is empty")
            
            # 設定値を辞書に変換（ヘッダー行をスキップ）
            settings = {row[0]: row[1] for row in values[1:] if len(row) >= 2}
            
            # settings.iniの設定値をキーとして使用
            self.friend_sheet_name = settings.get(self.friend_data_key.strip('"'))  # クォートを除去
            self.anq_sheet_name = settings.get(self.anq_data_key.strip('"'))  # クォートを除去
            
            if not self.friend_sheet_name or not self.anq_sheet_name:
                logger.error("❌ 必要なシート名の設定が見つかりません")
                logger.error(f"友達リストDLデータのキー: {self.friend_data_key}")
                logger.error(f"アンケートDLデータのキー: {self.anq_data_key}")
                raise ValueError("Required sheet names not found in settings")
            
            logger.info(f"✓ シート名の設定を読み込みました（友達: {self.friend_sheet_name}, アンケート: {self.anq_sheet_name}）")
            
        except Exception as e:
            logger.error(f"❌ シート名の設定読み込みに失敗: {str(e)}")
            raise

    def update_sheet(self, csv_path, sheet_type='friend'):
        """CSVファイルの内容をスプレッドシートに転記"""
        try:
            logger.info(f"📊 CSVファイルの読み込みを開始: {csv_path}")
            
            # シートタイプに応じてシート名を設定
            self.sheet_name = self.anq_sheet_name if sheet_type == 'anq_data' else self.friend_sheet_name
            logger.info(f"📝 転記先シート: {self.sheet_name}")
            
            # CSVファイルを読み込み（Windows向けにcp932エンコーディングを指定）
            df = pd.read_csv(csv_path, encoding='cp932')
            logger.info(f"✓ CSVファイルを読み込みました（{len(df)}行）")
            
            # データを2次元配列に変換し、特殊文字を処理
            values = []
            # ヘッダー行を追加
            headers = [str(col).strip() for col in df.columns]
            values.append(headers)
            
            # データ行を追加
            for _, row in df.iterrows():
                # 各セルの値を文字列に変換し、空白を処理
                processed_row = []
                for value in row:
                    if pd.isna(value):  # None や NaN の処理
                        processed_row.append('')
                    else:
                        # 文字列に変換して空白を処理
                        processed_value = str(value).strip()
                        processed_row.append(processed_value)
                values.append(processed_row)
            
            # スプレッドシートをクリア
            range_name = f"{self.sheet_name}!A1:Z"
            self.service.spreadsheets().values().clear(
                spreadsheetId=self.spreadsheet_id,
                range=range_name
            ).execute()
            logger.info("✓ スプレッドシートをクリアしました")
            
            # データを書き込み
            body = {
                'values': values
            }
            result = self.service.spreadsheets().values().update(
                spreadsheetId=self.spreadsheet_id,
                range=f"{self.sheet_name}!A1",
                valueInputOption='USER_ENTERED',
                body=body
            ).execute()
            
            logger.info(f"✅ スプレッドシートの更新が完了しました（{result.get('updatedRows')}行）")
            return True
            
        except Exception as e:
            logger.error(f"❌ スプレッドシートの更新に失敗: {str(e)}")
            return False 

    def log_operation(self, operation_type, status, error_message=None):
        """操作ログをスプレッドシートに記録"""
        try:
            from datetime import datetime
            
            # ログデータの作成
            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            log_data = [[
                now,  # 日時
                operation_type,  # 操作種別
                status,  # ステータス
                error_message or ''  # エラーメッセージ
            ]]
            
            # ログシートに追記（ハードコーディングされたシート名を使用）
            body = {
                'values': log_data
            }
            
            # 最終行に追加
            result = self.service.spreadsheets().values().append(
                spreadsheetId=self.spreadsheet_id,
                range=f"{self.log_sheet_name}!A1",
                valueInputOption='RAW',
                insertDataOption='INSERT_ROWS',
                body=body
            ).execute()
            
            logger.info(f"✓ 操作ログを記録しました: {operation_type} - {status}")
            return True
            
        except Exception as e:
            logger.error(f"❌ 操作ログの記録に失敗: {str(e)}")
            return False 