import datetime
from ...utils.logging_config import get_logger
from ..common.spreadsheet import get_spreadsheet_connection, get_column_index
from ..common.settings import load_sheet_settings
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from ...utils.environment import EnvironmentUtils as env

logger = get_logger(__name__)

def update_consult_transfer_list(matching_ids):
    """
    相談転記先リストにマッチしたIDを追加する
    IDがすでに登録されている場合はスキップする
    """
    try:
        # スプレッドシートに接続
        spreadsheet = get_spreadsheet_connection()
        if not spreadsheet:
            return False
        
        # シート設定を読み込む
        _, _, transfer_sheet_name = load_sheet_settings(spreadsheet)
        if not transfer_sheet_name:
            return False
        
        # 相談転記先リストシートを取得
        transfer_sheet = spreadsheet.worksheet(transfer_sheet_name)
        transfer_data = transfer_sheet.get_all_values()
        logger.info(f"✓ 相談転記先リストシートからデータ取得: {len(transfer_data)}行")
        
        # 相談転記先リストのヘッダー行
        if not transfer_data:
            logger.error("❌ 相談転記先リストが空です")
            return False
            
        transfer_headers = transfer_data[0]
        logger.info(f"相談転記先リストのヘッダー: {transfer_headers}")
        
        # IDと相談日のカラムインデックスを特定
        try:
            id_col_idx = transfer_headers.index('ID')
            date_col_idx = transfer_headers.index('相談日')
        except ValueError as e:
            logger.error(f"❌ 相談転記先リストに必要なカラムがありません: {str(e)}")
            return False
        
        # 既存のIDを取得（int型として扱う）
        existing_ids = []
        for row in transfer_data[1:]:
            if row and len(row) > id_col_idx and row[id_col_idx]:
                try:
                    existing_ids.append(int(row[id_col_idx]))
                except ValueError:
                    logger.warning(f"数値に変換できないID値をスキップ: {row[id_col_idx]}")
        
        logger.info(f"既存のID数: {len(existing_ids)}")
        
        # 追加すべきIDを特定（int型として扱う）
        ids_to_add = []
        for id_value in matching_ids:
            try:
                int_id = int(id_value)
                if int_id not in existing_ids:
                    ids_to_add.append(int_id)
            except ValueError:
                logger.warning(f"数値に変換できないID値をスキップ: {id_value}")
        
        logger.info(f"追加すべきID数: {len(ids_to_add)}")
        
        if not ids_to_add:
            logger.info("✓ 追加すべきIDがありません。すべて既存です。")
            return True
        
        # 今日の日付と時刻を取得 (YYYY-MM-DD HH:MM:SS形式)
        today = datetime.datetime.now().strftime('%Y/%m/%d')
        logger.info(f"相談日として設定するタイムスタンプ: {today} (yyyy/mm/dd形式)")
        
        # 新しい行を作成
        new_rows = []
        for id_value in ids_to_add:
            # 転記先シートのカラム数に合わせて空の行を作成
            new_row = [''] * len(transfer_headers)
            new_row[id_col_idx] = id_value  # int型のまま設定
            new_row[date_col_idx] = today   # yyyy-mm-dd hh:mm:ss形式のタイムスタンプを設定
            logger.info(f"追加行を作成: ID={id_value}, 相談日={today}")
            new_rows.append(new_row)
        
        # スプレッドシートに追加
        start_row = len(transfer_data) + 1  # 追加開始行（1-indexed）
        end_row = start_row + len(new_rows) - 1
        end_col = len(transfer_headers)
        
        # A1形式のレンジを構築
        range_notation = f"A{start_row}:{chr(64 + end_col)}{end_row}"
        logger.info(f"更新範囲: {range_notation}, {len(new_rows)}行 x {len(transfer_headers)}列")
        
        # 更新処理
        logger.info(f"スプレッドシートを更新: シート名={transfer_sheet_name}, レンジ={range_notation}, 行数={len(new_rows)}")
        transfer_sheet.update(
            values=new_rows, 
            range_name=range_notation, 
            value_input_option="USER_ENTERED"  # 日付を日付型として認識させる
        )
        
        # ログシート（SSID_log）の「相談者一覧」シートにもIDを転記
        try:
            # ログ用スプレッドシートIDを取得
            log_spreadsheet_id = env.get_config_value("SPREADSHEET", "SSID_log")
            logger.info(f"ログスプレッドシートID: {log_spreadsheet_id}")
            
            # ログ用スプレッドシートに接続
            log_client = gspread.authorize(ServiceAccountCredentials.from_json_keyfile_name(
                env.get_service_account_file(), 
                ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
            ))
            log_spreadsheet = log_client.open_by_key(log_spreadsheet_id)
            
            # 相談者一覧シートを取得
            log_sheet = log_spreadsheet.worksheet("相談者一覧")
            logger.info(f"✓ 「相談者一覧」シートを取得しました")
            
            # A列の値だけを取得
            a_column_values = log_sheet.col_values(1)  # 1はA列を意味する（gspreadではカラムインデックスは1始まり）
            
            # A列の最終行+1を計算（空白行があっても最終データ行の次を取得）
            log_last_row = len(a_column_values) + 1
            logger.info(f"A列の最終行: {log_last_row - 1}, 次の行: {log_last_row}")
            
            # 転記するIDの行を作成
            log_rows = []
            for id_value in ids_to_add:
                log_rows.append([id_value])  # A列にのみIDを設定
            
            # 転記処理
            if log_rows:
                log_range = f"A{log_last_row}:A{log_last_row + len(log_rows) - 1}"
                logger.info(f"ログシートに追加: レンジ={log_range}, 行数={len(log_rows)}")
                log_sheet.update(
                    values=log_rows, 
                    range_name=log_range,
                    value_input_option="USER_ENTERED"
                )
                logger.info(f"✓ ログシート「相談者一覧_haraguchi作業」にID {len(log_rows)}件を転記しました")
        except Exception as e:
            logger.error(f"❌ ログシートへの転記中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            # メイン処理は継続させるためにエラーは捕捉するだけ
        
        logger.info(f"✓ 相談転記先リストを正常に更新しました")
        return True
        
    except Exception as e:
        logger.error(f"❌ 相談転記先リストの更新中にエラーが発生しました: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return False