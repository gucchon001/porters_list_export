import gspread
from ...utils.logging_config import get_logger
from ..common.spreadsheet import get_spreadsheet_connection, get_column_index
from ..common.settings import load_sheet_settings

logger = get_logger(__name__)

def find_ids_with_matching_flags():
    """
    フラグに一致するIDを見つける
    """
    try:
        # スプレッドシートに接続
        spreadsheet = get_spreadsheet_connection()
        if not spreadsheet:
            return None
        
        # シート設定を読み込む
        friend_sheet_name, consult_sheet_name, _ = load_sheet_settings(spreadsheet)
        if not friend_sheet_name or not consult_sheet_name:
            return None
        
        # 1. 相談フラグマスタシートからデータを取得
        consult_sheet = spreadsheet.worksheet(consult_sheet_name)
        consult_data = consult_sheet.get_all_values()
        logger.info(f"✓ 相談フラグマスタシートからデータ取得: {len(consult_data)}行")
        
        # 相談フラグマスタのヘッダー行
        consult_headers = consult_data[0]
        logger.info(f"相談フラグマスタのヘッダー: {consult_headers}")
        
        # 「項目」と「フラグ」列のインデックスを取得
        item_index = get_column_index(consult_headers, "項目")
        flag_index = get_column_index(consult_headers, "フラグ")
        
        if item_index is None or flag_index is None:
            return None
        
        # フラグが1の項目を抽出 (B)
        flagged_items = []
        for row in consult_data[1:]:  # ヘッダー行をスキップ
            if len(row) > max(item_index, flag_index) and row[flag_index] == "1":
                flagged_items.append(row[item_index])
        
        logger.info(f"フラグが1の項目 (B): {flagged_items}")
        
        # 2. 友達リストDLデータシートからデータを取得
        friend_sheet = spreadsheet.worksheet(friend_sheet_name)
        friend_data = friend_sheet.get_all_values()
        logger.info(f"✓ 友達リストDLデータシートからデータ取得: {len(friend_data)}行")
        
        # 友達リストのヘッダー行（2行目を使用）
        if len(friend_data) < 2:
            logger.error("❌ 友達リストデータが不足しています")
            return None
            
        friend_headers = friend_data[1]  # 2行目をヘッダーとして使用
        logger.info(f"友達リストのヘッダー (2行目): {friend_headers}")
        
        # 「ID」と「対応マーク」列のインデックスを取得
        id_index = get_column_index(friend_headers, "ID")
        mark_index = get_column_index(friend_headers, "対応マーク")
        
        if id_index is None or mark_index is None:
            return None
        
        # 対応マークとフラグ項目を比較して一致するIDを抽出
        matching_ids = []
        for row in friend_data[2:]:  # ヘッダー行をスキップ（3行目から開始）
            if len(row) <= max(id_index, mark_index):
                continue
                
            mark_value = row[mark_index]
            
            # マークがフラグ付き項目のいずれかに一致するか確認
            if any(item in mark_value for item in flagged_items):
                matching_ids.append(row[id_index])
                logger.info(f"一致: ID={row[id_index]}, マーク={mark_value}")
        
        logger.info(f"一致するID数: {len(matching_ids)}")
        logger.info(f"一致するID: {matching_ids}")
        
        return matching_ids
        
    except gspread.exceptions.WorksheetNotFound as e:
        logger.error(f"❌ シートが見つかりません: {str(e)}")
        return None
    except Exception as e:
        logger.error(f"❌ データ検索中にエラーが発生しました: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return None 