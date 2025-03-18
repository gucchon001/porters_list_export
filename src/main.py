#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
PORTERSシステムへのログイン処理と二重ログイン回避を実装するメインスクリプト

このスクリプトは、PORTERSシステムへのログイン処理と二重ログイン回避の機能を提供します。
Seleniumを使用してブラウザを制御し、ログイン画面へのアクセス、認証情報の入力、
ログインボタンのクリック、二重ログインポップアップの処理を行います。
また、ログイン後に「その他業務」ボタンをクリックしてメニュー項目5を押す処理も実行します。
"""

import os
import sys
import time
import argparse
from pathlib import Path

# プロジェクトのルートディレクトリをPYTHONPATHに追加
root_dir = Path(__file__).parent.parent
sys.path.append(str(root_dir))

from src.utils.environment import EnvironmentUtils as env
from src.utils.logging_config import get_logger
from src.modules.porters.login import PortersLogin
from src.modules.porters.operations import PortersOperations
from src.modules.spreadsheet_aggregator import SpreadsheetAggregator

logger = get_logger(__name__)

def setup_environment():
    """
    実行環境のセットアップを行う
    - 必要なディレクトリの作成
    - 設定ファイルの読み込み
    """
    # 必要なディレクトリの作成
    os.makedirs("logs", exist_ok=True)
    os.makedirs("data", exist_ok=True)
    
    # 設定ファイルの読み込み
    try:
        # 環境変数の読み込み
        env.load_env()
        logger.info("環境変数を読み込みました")
        return True
    except Exception as e:
        logger.error(f"設定ファイルの読み込みに失敗しました: {str(e)}")
        return False

def parse_arguments():
    """
    コマンドライン引数を解析する
    
    Returns:
        argparse.Namespace: 解析された引数
    """
    parser = argparse.ArgumentParser(description='PORTERSシステムへのログイン処理')
    parser.add_argument('--headless', action='store_true', help='ヘッドレスモードで実行')
    parser.add_argument('--env', default='development', help='実行環境 (development または production)')
    parser.add_argument('--skip-operations', action='store_true', help='業務操作をスキップする')
    parser.add_argument('--log-level', default='INFO', choices=['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'], 
                        help='ログレベル (DEBUG, INFO, WARNING, ERROR, CRITICAL)')
    parser.add_argument('--process', choices=['candidates', 'entryprocess', 'both', 'sequential'], default='candidates',
                        help='実行する処理フロー (candidates: 求職者一覧のエクスポート, entryprocess: 選考プロセス一覧の表示, both: 両方を順に実行, sequential: 求職者一覧処理後に選考プロセスも実行)')
    parser.add_argument('--aggregate', choices=['users', 'entryprocess', 'both', 'none'], default='none',
                        help='実行する集計処理の種類 (users: 求職者フェーズ別集計, entryprocess: 選考プロセス集計, both: 両方を実行, none: 集計処理を実行しない)')
    
    return parser.parse_args()

def main():
    """
    メイン処理
    
    以下の処理を実行します：
    1. 環境設定
    2. PORTERSシステムへのログイン
    3. 選択された処理フローの実行（求職者一覧のエクスポート、選考プロセス一覧の表示、または両方）
    4. スプレッドシートへのデータ集計（オプション）
    5. ログアウト処理
    6. ブラウザの終了
    
    Returns:
        int: 処理が成功した場合は0、失敗した場合は1
    """
    try:
        logger.info("================================================================================")
        logger.info("PORTERSシステムログインツールを開始します")
        logger.info("================================================================================")
        
        # コマンドライン引数の解析
        args = parse_arguments()
        
        # 環境変数の設定
        os.environ['APP_ENV'] = args.env
        os.environ['LOG_LEVEL'] = args.log_level
        logger.info(f"実行環境: {args.env}")
        logger.info(f"ログレベル: {args.log_level}")
        logger.info(f"処理フロー: {args.process}")
        logger.info(f"集計処理: {args.aggregate}")
        
        # 環境設定
        if not setup_environment():
            logger.error("環境設定に失敗したため、処理を中止します")
            return 1
        
        # 設定ファイルのパス
        selectors_path = os.path.join(root_dir, "config", "selectors.csv")
        
        # ブラウザとログインオブジェクトの初期化
        browser = None
        login = None
        
        # PORTERS操作の結果ステータス
        porters_success = False
        
        try:
            # PORTERSへのログイン
            success, browser, login = PortersLogin.login_to_porters(
                selectors_path=selectors_path, 
                headless=args.headless
            )
            
            if success:
                logger.info("ログイン処理が完了しました")
                
                # 業務操作の実行（スキップオプションが指定されていない場合）
                if not args.skip_operations and browser:
                    logger.info("業務操作を開始します")
                    operations = PortersOperations(browser)
                    
                    # 処理フローの選択
                    if args.process == 'candidates':
                        # 求職者一覧の処理フロー
                        logger.info("求職者一覧のエクスポート処理フローを実行します")
                        if not operations.execute_operations_flow():
                            logger.error("求職者一覧のエクスポート処理フローに失敗しました")
                            porters_success = False
                        else:
                            logger.info("求職者一覧のエクスポート処理フローが正常に完了しました")
                            porters_success = True
                    elif args.process == 'entryprocess':
                        # 選考プロセスの処理フロー
                        logger.info("選考プロセス一覧の表示処理フローを実行します")
                        if not operations.access_selection_processes():
                            logger.error("選考プロセス一覧の表示処理フローに失敗しました")
                            porters_success = False
                        else:
                            logger.info("選考プロセス一覧の表示処理フローが正常に完了しました")
                            porters_success = True
                    elif args.process == 'both':
                        # 両方の処理フローを順に実行
                        logger.info("=== 両方の処理フローを順に実行します ===")
                        
                        # 求職者一覧の処理フロー
                        logger.info("1. 求職者一覧のエクスポート処理フローを実行します")
                        candidates_success = operations.execute_operations_flow()
                        if not candidates_success:
                            logger.error("求職者一覧のエクスポート処理フローに失敗しました")
                        else:
                            logger.info("求職者一覧のエクスポート処理フローが正常に完了しました")
                        
                        # 処理間の待機時間
                        logger.info("次の処理に進む前に10秒間待機します")
                        time.sleep(10)
                        
                        # 選考プロセスの処理フロー
                        logger.info("2. 選考プロセス一覧の表示処理フローを実行します")
                        entryprocess_success = operations.access_selection_processes()
                        if not entryprocess_success:
                            logger.error("選考プロセス一覧の表示処理フローに失敗しました")
                        else:
                            logger.info("選考プロセス一覧の表示処理フローが正常に完了しました")
                        
                        # 両方の処理結果をログに出力
                        if candidates_success and entryprocess_success:
                            logger.info("✅ 両方の処理フローが正常に完了しました")
                            porters_success = True
                        else:
                            logger.warning("⚠️ 一部の処理フローが失敗しました")
                            porters_success = False
                    elif args.process == 'sequential':
                        # 求職者一覧処理後に選考プロセスも実行
                        logger.info("=== 求職者一覧処理後に選考プロセスも実行します ===")
                        
                        # PortersOperationsクラスの新しいメソッドを使用
                        candidates_success, entryprocess_success = operations.execute_both_processes()
                        
                        # 両方の処理結果をログに出力
                        if candidates_success and entryprocess_success:
                            logger.info("✅ 両方の処理フローが正常に完了しました")
                            porters_success = True
                        else:
                            logger.warning("⚠️ 一部の処理フローが失敗しました")
                            porters_success = False
                    
                    # 操作完了後の待機時間
                    logger.info("操作完了後、5秒間待機します")
                    time.sleep(5)
                
                    # ログアウト処理
                    if login:
                        logger.info("ログアウト処理を開始します")
                        try:
                            logout_result = login.logout()
                            if logout_result:
                                logger.info("ログアウト処理が正常に完了しました")
                            else:
                                logger.warning("ログアウト処理に失敗しました")
                        except Exception as e:
                            logger.error(f"ログアウト処理中に例外が発生しました: {str(e)}")
                
                # 集計処理の実行（オプションが指定されている場合）
                if args.aggregate != 'none':
                    logger.info("=== スプレッドシート集計処理を開始します ===")
                    aggregator = SpreadsheetAggregator()
                    
                    # 集計処理を実行
                    users_success, entryprocess_success = aggregator.run_aggregation(args.aggregate)
                    
                    # 結果の集計
                    success_count = 0
                    total_count = 0
                    
                    if args.aggregate in ['users', 'both']:
                        total_count += 1
                        if users_success:
                            success_count += 1
                            logger.info("✅ 求職者フェーズ別集計処理が正常に完了しました")
                        else:
                            logger.error("❌ 求職者フェーズ別集計処理に失敗しました")
                    
                    if args.aggregate in ['entryprocess', 'both']:
                        total_count += 1
                        if entryprocess_success:
                            success_count += 1
                            logger.info("✅ 選考プロセス集計処理が正常に完了しました")
                        else:
                            logger.error("❌ 選考プロセス集計処理に失敗しました")
                    
                    # 集計処理の結果ステータス
                    if success_count == total_count:
                        logger.info("✅ すべての集計処理が正常に完了しました")
                    else:
                        logger.warning(f"⚠️ 実行された {total_count} 件の集計処理のうち、{success_count} 件が成功し、{total_count - success_count} 件が失敗しました")
                
                logger.info("================================================================================")
                logger.info("PORTERSシステムログインツールを正常に終了します")
                logger.info("================================================================================")
                return 0
            else:
                logger.error("ログイン処理に失敗しました")
                logger.info("================================================================================")
                logger.info("PORTERSシステムログインツールを異常終了します")
                logger.info("================================================================================")
                return 1
                
        except Exception as e:
            logger.error(f"処理中に例外が発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            logger.info("================================================================================")
            logger.info("PORTERSシステムログインツールを異常終了します")
            logger.info("================================================================================")
            return 1
            
        finally:
            # 必ずブラウザを終了する
            if browser:
                logger.info("ブラウザを終了します")
                try:
                    browser.quit()
                    logger.info("ブラウザを正常に終了しました")
                except Exception as e:
                    logger.error(f"ブラウザの終了中に例外が発生しました: {str(e)}")
            
    except Exception as e:
        logger.error(f"予期しないエラーが発生しました: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        logger.info("================================================================================")
        logger.info("PORTERSシステムログインツールを異常終了します")
        logger.info("================================================================================")
        return 1

if __name__ == "__main__":
    sys.exit(main())
