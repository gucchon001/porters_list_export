import argparse
import sys
import os
from pathlib import Path
from src.utils.environment import EnvironmentUtils as env
from src.utils.logging_config import get_logger
from src.modules.consult.consult_flags import find_ids_with_matching_flags
from src.modules.consult.transfer_list import update_consult_transfer_list
from src.modules.anq_data.analyzer import analyze_anq_data
from src.modules.porters.importer import import_to_porters
from src.modules.common.spreadsheet import get_spreadsheet_connection

logger = get_logger(__name__)

def setup_configurations():
    """設定ファイルと機密情報をロードしてデータを取得します。"""
    try:
        # 環境変数のロード
        env.load_env()
        
        # SERVICE_ACCOUNT_FILE のパスを確認する（変換は不要）
        service_account_file = os.getenv('SERVICE_ACCOUNT_FILE')
        if service_account_file:
            logger.info(f"SERVICE_ACCOUNT_FILE: {service_account_file}")
        
        return True
    except Exception as e:
        logger.error(f"設定のロードに失敗: {str(e)}")
        return False

def process_consult_flags():
    """
    処理ブロック1: 相談フラグの確認と新規IDの抽出
    - 相談フラグマスタからフラグが1に設定されている項目を抽出
    - 友達リストDLデータシートの"対応マーク"列を確認し、抽出したフラグ項目に一致するIDを特定
    - 既存の相談転記先リストに存在しないIDを新規IDとして抽出
    - 抽出した新規IDと相談日（現在の日付）を相談転記先リスト（相談Raw）に追加
    
    Returns:
        list: 一致するIDのリスト。エラーまたは該当なしの場合はNone
    """
    logger.info("=" * 80)
    logger.info("処理ブロック1: 相談フラグの確認と新規IDの抽出を開始します")
    logger.info("=" * 80)
    
    # スプレッドシートへの接続を取得
    spreadsheet = get_spreadsheet_connection()
    if not spreadsheet:
        logger.error("スプレッドシートへの接続に失敗したため、処理を中止します")
        return None
    
    # フラグに一致するIDを検索
    matching_ids = find_ids_with_matching_flags()
    
    if matching_ids:
        logger.info(f"✅ フラグ条件に一致するIDを{len(matching_ids)}件見つけました:")
        
        # IDを書式なしテキストとして表示（コピー用）
        logger.info("以下のIDをコピーして使用できます（書式なしテキスト）:")
        id_text = "\n".join(matching_ids)
        logger.info(f"\n{id_text}")
        
        # 個別のIDもログに記録
        for id_value in matching_ids:
            logger.info(f"  - {id_value}")
            
        # 相談転記先リストを更新
        logger.info("相談転記先リストの更新を開始します...")
        if update_consult_transfer_list(matching_ids):
            logger.info("✅ 相談転記先リストの更新が完了しました")
            logger.info("=" * 80)
            logger.info("処理ブロック1: 相談フラグの確認と新規IDの抽出が完了しました")
            logger.info("=" * 80)
            return matching_ids
        else:
            logger.error("❌ 相談転記先リストの更新に失敗しました")
            logger.info("=" * 80)
            logger.info("処理ブロック1: 相談フラグの確認と新規IDの抽出が失敗しました")
            logger.info("=" * 80)
            return None
    else:
        logger.warning("❌ 一致するIDが見つからないか、処理中にエラーが発生しました")
        logger.info("=" * 80)
        logger.info("処理ブロック1: 相談フラグの確認と新規IDの抽出が完了しました（該当IDなし）")
        logger.info("=" * 80)
        return None

def process_anq_data(matching_ids=None):
    """
    処理ブロック2: アンケートデータの取得とCSV出力
    - 新しく追加されたIDのユーザーのアンケートDLデータからレコードを取得
    - 取得したレコードをCSVファイルとして保存（cp932エンコーディング）
    
    Args:
        matching_ids (list, optional): 処理対象のIDリスト。指定がない場合は処理ブロック1の結果を使用
        
    Returns:
        bool: 処理が成功した場合はTrue、失敗した場合はFalse
    """
    # IDが指定されていない場合は処理ブロック1を実行
    if matching_ids is None:
        logger.info("IDが指定されていないため、処理ブロック1を実行します")
        matching_ids = process_consult_flags()
    
    if not matching_ids:
        logger.warning("一致するIDがないため、アンケートデータの取得をスキップします")
        return False
    
    logger.info("=" * 80)
    logger.info("処理ブロック2: アンケートデータの取得とCSV出力を開始します")
    logger.info("=" * 80)
    
    # アンケートデータを取得してCSV出力
    logger.info(f"{len(matching_ids)}件のIDに対するアンケートデータを取得します")
    
    success, _ = analyze_anq_data(matching_ids)
    if success:
        logger.info("✅ アンケートデータの取得とCSV出力が完了しました")
        logger.info("=" * 80)
        logger.info("処理ブロック2: アンケートデータの取得とCSV出力が完了しました")
        logger.info("=" * 80)
        return True
    else:
        logger.error("❌ アンケートデータの取得とCSV出力に失敗しました")
        logger.info("=" * 80)
        logger.info("処理ブロック2: アンケートデータの取得とCSV出力が失敗しました")
        logger.info("=" * 80)
        return False

def process_porters_import(run_anq_data=False, matching_ids=None):
    """
    処理ブロック3: PORTERSへのデータインポート
    - Seleniumを使用してPORTERSシステムにログイン
    - 求職者インポート機能を使用
    - 保存したCSVファイルをアップロード
    - LINE初回アンケート取込形式でインポート
    
    Args:
        run_anq_data (bool): アンケートデータ処理を実行するかどうか
        matching_ids (list, optional): アンケートデータ処理に使用するIDリスト
        
    Returns:
        bool: 処理が成功した場合はTrue、失敗した場合はFalse
    """
    # アンケートデータ処理を実行する場合
    if run_anq_data:
        logger.info("アンケートデータ処理を実行します")
        anq_success = process_anq_data(matching_ids)
        if not anq_success:
            logger.warning("アンケートデータの取得に失敗したため、PORTERSへのインポートをスキップします")
            return False
    
    logger.info("=" * 80)
    logger.info("処理ブロック3: PORTERSへのデータインポートを開始します")
    logger.info("=" * 80)
    
    # PORTERSにインポート
    logger.info("CSVファイルをPORTERSにインポートします")
    
    if import_to_porters():
        logger.info("✅ PORTERSへのデータインポートが完了しました")
        logger.info("=" * 80)
        logger.info("処理ブロック3: PORTERSへのデータインポートが完了しました")
        logger.info("=" * 80)
        return True
    else:
        logger.error("❌ PORTERSへのデータインポートに失敗しました")
        logger.info("=" * 80)
        logger.info("処理ブロック3: PORTERSへのデータインポートが失敗しました")
        logger.info("=" * 80)
        return False

def run_all():
    """
    すべての処理ブロックを順番に実行する
    
    Returns:
        tuple: (matching_ids, anq_success, porters_success) 各処理ブロックの結果
    """
    logger.info("相談フラグ管理システムの処理を開始します")
    
    # 処理ブロック1: 相談フラグの確認と新規IDの抽出
    matching_ids = process_consult_flags()
    
    # 処理ブロック2: アンケートデータの取得とCSV出力
    if matching_ids:
        anq_success = process_anq_data(matching_ids)
    else:
        logger.warning("一致するIDがないため、後続の処理をスキップします")
        anq_success = False
    
    # 処理ブロック3: PORTERSへのデータインポート
    if anq_success:
        porters_success = process_porters_import()
    else:
        logger.warning("アンケートデータの取得に失敗したため、PORTERSへのインポートをスキップします")
        porters_success = False
    
    # 処理結果のサマリーを表示
    logger.info("=" * 80)
    logger.info("相談フラグ管理システムの処理結果サマリー")
    logger.info("=" * 80)
    logger.info(f"処理ブロック1: 相談フラグの確認と新規IDの抽出 - {'成功' if matching_ids else '失敗または該当なし'}")
    logger.info(f"処理ブロック2: アンケートデータの取得とCSV出力 - {'成功' if anq_success else '失敗またはスキップ'}")
    logger.info(f"処理ブロック3: PORTERSへのデータインポート - {'成功' if porters_success else '失敗またはスキップ'}")
    logger.info("=" * 80)
    
    if matching_ids and anq_success and porters_success:
        logger.info("✅ すべての処理が正常に完了しました")
    else:
        logger.warning("⚠️ 一部の処理が失敗またはスキップされました")
    
    logger.info("相談フラグ管理システムの処理を終了します")
    
    return matching_ids, anq_success, porters_success

def parse_arguments():
    """
    コマンドライン引数を解析する
    
    Returns:
        argparse.Namespace: 解析された引数
    """
    parser = argparse.ArgumentParser(description='相談フラグ管理システム')
    parser.add_argument('--block', type=int, choices=[1, 2, 3], help='実行する処理ブロック (1: 相談フラグ確認, 2: アンケートデータ取得, 3: PORTERSインポート)')
    parser.add_argument('--ids', nargs='+', help='処理対象のIDリスト（カンマ区切りまたは複数指定）')
    parser.add_argument('--all', action='store_true', help='すべての処理ブロックを実行')
    parser.add_argument('--env', default='development', help='実行環境 (development または production)')
    
    return parser.parse_args()

if __name__ == "__main__":
    try:
        # コマンドライン引数の解析
        args = parse_arguments()
        
        # 設定ファイルと機密情報をロード
        if not setup_configurations():
            logger.error("設定のロードに失敗したため、処理を中止します")
            sys.exit(1)
        
        # スプレッドシートへの接続をテスト
        spreadsheet = get_spreadsheet_connection()
        if not spreadsheet:
            logger.error("スプレッドシートへの接続に失敗したため、処理を中止します")
            sys.exit(1)
        logger.info("スプレッドシートへの接続テストに成功しました")
        
        # IDリストの処理
        ids_list = None
        if args.ids:
            # カンマ区切りの場合は分割
            if len(args.ids) == 1 and ',' in args.ids[0]:
                ids_list = args.ids[0].split(',')
            else:
                ids_list = args.ids
            
            logger.info(f"コマンドラインから{len(ids_list)}件のIDを受け取りました")
        
        # 処理ブロックの実行
        if args.all or (not args.block and not args.ids):
            # すべての処理ブロックを実行
            run_all()
        elif args.block == 1:
            # 処理ブロック1: 相談フラグの確認と新規IDの抽出
            process_consult_flags()
        elif args.block == 2:
            # 処理ブロック2: アンケートデータの取得とCSV出力
            process_anq_data(ids_list)
        elif args.block == 3:
            # 処理ブロック3: PORTERSへのデータインポート
            process_porters_import(run_anq_data=(ids_list is not None), matching_ids=ids_list)
        else:
            logger.error("実行する処理ブロックが指定されていません")
            sys.exit(1)
            
    except Exception as e:
        logger.error(f"予期せぬエラーが発生しました: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        sys.exit(1)