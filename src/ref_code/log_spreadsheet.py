import pandas as pd
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime
from ..utils.environment import EnvironmentUtils as env
from ..utils.logging_config import get_logger

logger = get_logger(__name__)

class LogSpreadsheet:
    def __init__(self):
        """ログ用スプレッドシート操作を管理するクラス"""
        self.credentials = self._get_credentials()
        self.service = build('sheets', 'v4', credentials=self.credentials)
        self.spreadsheet_id = env.get_config_value('SPREADSHEET', 'SSID')
        self.log_sheet_name = 'logsheet'

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

    def log_operation(self, operation_type, status, error_message=None):
        """操作ログをスプレッドシートに記録"""
        try:
            # ログデータの作成
            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            log_data = [[
                now,  # 日時
                operation_type,  # 操作種別
                status,  # ステータス
                error_message or ''  # エラーメッセージ
            ]]
            
            # ログシートに追記
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