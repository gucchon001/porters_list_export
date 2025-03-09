import gspread
from oauth2client.service_account import ServiceAccountCredentials
from ...utils.environment import EnvironmentUtils as env
from ...utils.logging_config import get_logger

logger = get_logger(__name__)

def get_spreadsheet_connection():
    """
    Google スプレッドシートへの接続を取得する
    """
    try:
        # 設定値の取得
        spreadsheet_id = env.get_config_value("SPREADSHEET", "SSID")
        logger.info(f"Spreadsheet ID: {spreadsheet_id}")
        
        # サービスアカウントの認証情報を取得
        service_account_file = env.get_service_account_file()
        logger.info(f"Service Account File: {service_account_file}")
        
        # APIへのアクセス範囲を設定
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        
        # 認証情報を取得
        credentials = ServiceAccountCredentials.from_json_keyfile_name(service_account_file, scope)
        
        # gspread クライアントを作成
        client = gspread.authorize(credentials)
        
        # スプレッドシートを開く
        spreadsheet = client.open_by_key(spreadsheet_id)
        logger.info(f"✓ スプレッドシートへの接続に成功しました: {spreadsheet.title}")
        
        return spreadsheet
        
    except Exception as e:
        logger.error(f"❌ スプレッドシートへの接続中にエラーが発生しました: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return None

def get_column_index(headers, column_name):
    """
    ヘッダー行から特定の列名のインデックスを取得する
    """
    try:
        index = headers.index(column_name)
        return index
    except ValueError:
        logger.error(f"❌ 列 '{column_name}' がヘッダー行に見つかりません: {headers}")
        return None 