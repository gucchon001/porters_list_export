import gspread
from ...utils.logging_config import get_logger
from ..common.spreadsheet import get_spreadsheet_connection, get_column_index
from ..common.settings import load_sheet_settings

logger = get_logger(__name__)

def find_ids_with_matching_flags():
    """
    フラグに一致するIDを検索する
    
    Returns:
        list: 一致するIDのリスト。エラーまたは該当なしの場合は空リスト
    """
    try:
        # スプレッドシートに接続
        spreadsheet = get_spreadsheet_connection()
        if not spreadsheet:
            return None
        
        # シート設定を読み込む
        friend_sheet_name, consult_sheet_name, transfer_sheet_name, anq_sheet_name = load_sheet_settings(spreadsheet)
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
        
        # 相談転記先リストから既存のIDを取得
        transfer_sheet = spreadsheet.worksheet(transfer_sheet_name)
        transfer_data = transfer_sheet.get_all_values()
        logger.info(f"✓ 相談転記先リストからデータ取得: {len(transfer_data)}行")
        
        # 相談転記先リストのヘッダー行
        transfer_headers = transfer_data[0]
        transfer_id_index = get_column_index(transfer_headers, "ID")
        
        if transfer_id_index is None:
            return None
            
        # 既存のIDをセットとして保持
        existing_ids = set()
        for row in transfer_data[1:]:  # ヘッダー行をスキップ
            if len(row) > transfer_id_index:
                existing_ids.add(row[transfer_id_index])
        
        logger.info(f"既存ID数: {len(existing_ids)}")
        
        # 対応マークとフラグ項目を比較して一致する新規IDを抽出
        new_matching_ids = []
        for row in friend_data[2:]:  # ヘッダー行をスキップ（3行目から開始）
            if len(row) <= max(id_index, mark_index):
                continue
                
            mark_value = row[mark_index]
            current_id = row[id_index]
            
            # マークがフラグ付き項目のいずれかに一致し、かつ既存IDに含まれていない場合
            if any(item in mark_value for item in flagged_items) and current_id not in existing_ids:
                new_matching_ids.append(current_id)
                logger.info(f"新規一致: ID={current_id}, マーク={mark_value}")
        
        logger.info(f"新規一致するID数: {len(new_matching_ids)}")
        logger.info(f"新規一致するID: {new_matching_ids}")
        
        if not new_matching_ids:
            logger.info("✓ 一致するIDは見つかりませんでした。新規データはありません。")
        
        return new_matching_ids
        
    except gspread.exceptions.WorksheetNotFound as e:
        logger.error(f"❌ シートが見つかりません: {str(e)}")
        return None
    except Exception as e:
        logger.error(f"❌ データ検索中にエラーが発生しました: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return [] 

def get_new_ids(self):
    """新規の相談フラグが立っているIDを取得する"""
    try:
        # 既存のIDリストを取得
        existing_ids = self._get_existing_ids()
        logger.info(f"既存ID数: {len(existing_ids)}")
        
        # 新規IDを取得
        new_ids = self._get_matching_ids(existing_ids)
        logger.info(f"新規一致するID数: {len(new_ids)}")
        logger.info(f"新規一致するID: {new_ids}")
        
        if not new_ids:
            logger.info("✓ 一致するIDは見つかりませんでした。新規データはありません。")
        
        return new_ids
        
    except Exception as e:
        logger.error(f"新規IDの取得中にエラーが発生しました: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return [] 