from ...utils.environment import EnvironmentUtils as env
from ...utils.logging_config import get_logger

logger = get_logger(__name__)

def load_sheet_settings(spreadsheet):
    """
    settingsシートから各シート名マッピングを読み込む
    """
    try:
        # settings.iniからキー名を取得
        friend_data_key = env.get_config_value('SHEET_NAMES', 'FRIEND_DATA').strip('"')  # "友達リストDLデータ"
        consult_flag_key = env.get_config_value('SHEET_NAMES', 'CONSULT_FLAG_MASTER').strip('"')  # "相談フラグマスタ"
        transfer_list_key = env.get_config_value('SHEET_NAMES', 'CONSULT_TRANSFER_LIST').strip('"')  # "相談転記先リスト"
        anq_data_key = env.get_config_value('SHEET_NAMES', 'ANQ_DATA').strip('"')  # "アンケートDLデータ"
        
        logger.info(f"設定から取得したキー - 友達リスト: {friend_data_key}, 相談フラグマスタ: {consult_flag_key}, 相談転記先: {transfer_list_key}, アンケートデータ: {anq_data_key}")
        
        # settingsシートを取得
        settings_sheet = spreadsheet.worksheet('settings')
        logger.info(f"✓ settingsシートを取得しました")
        
        # 設定値を取得
        data = settings_sheet.get_all_values()
        logger.info(f"settingsシートのデータ行数: {len(data)}")
        
        # 設定値を辞書に変換（ヘッダー行をスキップ）
        settings = {row[0]: row[1] for row in data[1:] if len(row) >= 2}
        logger.info(f"設定マッピング: {settings}")
        
        # キーが存在するか確認
        if friend_data_key not in settings:
            logger.error(f"❌ 設定キー '{friend_data_key}' が settingsシートに見つかりません")
            logger.error(f"利用可能な設定: {settings}")
            return None, None, None, None
            
        if consult_flag_key not in settings:
            logger.error(f"❌ 設定キー '{consult_flag_key}' が settingsシートに見つかりません")
            logger.error(f"利用可能な設定: {settings}")
            return None, None, None, None
            
        if transfer_list_key not in settings:
            logger.error(f"❌ 設定キー '{transfer_list_key}' が settingsシートに見つかりません")
            logger.error(f"利用可能な設定: {settings}")
            return None, None, None, None
            
        if anq_data_key not in settings:
            logger.error(f"❌ 設定キー '{anq_data_key}' が settingsシートに見つかりません")
            logger.error(f"利用可能な設定: {settings}")
            return None, None, None, None
            
        # 実際のシート名を取得
        friend_sheet_name = settings[friend_data_key]
        consult_sheet_name = settings[consult_flag_key]
        transfer_sheet_name = settings[transfer_list_key]
        anq_sheet_name = settings[anq_data_key]
        
        logger.info(f"✓ 友達データ実際のシート名: {friend_sheet_name}")
        logger.info(f"✓ 相談フラグマスタ実際のシート名: {consult_sheet_name}")
        logger.info(f"✓ 相談転記先リスト実際のシート名: {transfer_sheet_name}")
        logger.info(f"✓ アンケートデータ実際のシート名: {anq_sheet_name}")
        
        return friend_sheet_name, consult_sheet_name, transfer_sheet_name, anq_sheet_name
        
    except Exception as e:
        logger.error(f"❌ シート設定の読み込み中にエラーが発生しました: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return None, None, None, None 