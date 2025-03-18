#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
PORTERSシステムへのログイン処理と二重ログイン回避を実装するメインスクリプト

このスクリプトは、PORTERSシステムへのログイン処理と二重ログイン回避の機能を提供します。
処理フローの選択と引数の準備に特化し、実際の実行はPortersBrowserクラスに委譲します。
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
from src.modules.porters.browser import PortersBrowser
from src.modules.porters.operations import PortersOperations
from src.modules.spreadsheet_aggregator import SpreadsheetAggregator
from src.utils.slack_notifier import SlackNotifier

logger = get_logger(__name__)

# SlackNotifierのインスタンスを作成 (環境変数から設定を読み込む)
slack = SlackNotifier()

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

def candidates_workflow(browser, login, **kwargs):
    """
    求職者一覧の処理フローを実行する
    
    Args:
        browser (PortersBrowser): ブラウザオブジェクト
        login (PortersLogin): ログインオブジェクト
        **kwargs: その他のパラメータ
        
    Returns:
        bool: 処理が成功した場合はTrue、失敗した場合はFalse
    """
    logger.info("求職者一覧のエクスポート処理フローを実行します")
    operations = PortersOperations(browser)
    success = operations.execute_operations_flow()
    
    if success:
        logger.info("求職者一覧のエクスポート処理フローが正常に完了しました")
    else:
        logger.error("求職者一覧のエクスポート処理フローに失敗しました")
        
    return success

def entryprocess_workflow(browser, login, **kwargs):
    """
    選考プロセス一覧の処理フローを実行する
    
    Args:
        browser (PortersBrowser): ブラウザオブジェクト
        login (PortersLogin): ログインオブジェクト
        **kwargs: その他のパラメータ
        
    Returns:
        bool: 処理が成功した場合はTrue、失敗した場合はFalse
    """
    logger.info("選考プロセス一覧の表示処理フローを実行します")
    operations = PortersOperations(browser)
    success = operations.access_selection_processes()
    
    if success:
        logger.info("選考プロセス一覧の表示処理フローが正常に完了しました")
    else:
        logger.error("選考プロセス一覧の表示処理フローに失敗しました")
        
    return success

def both_workflow(browser, login, **kwargs):
    """
    求職者一覧と選考プロセス一覧の両方の処理フローを順に実行する
    
    Args:
        browser (PortersBrowser): ブラウザオブジェクト
        login (PortersLogin): ログインオブジェクト
        **kwargs: その他のパラメータ
        
    Returns:
        tuple: (success, (candidates_success, entryprocess_success)) 全体の成功/失敗と、各処理の成功/失敗
    """
    logger.info("=== 両方の処理フローを順に実行します ===")
    operations = PortersOperations(browser)
    
    # 求職者一覧の処理フロー
    logger.info("1. 求職者一覧のエクスポート処理フローを実行します")
    candidates_success = operations.execute_operations_flow()
    
    if candidates_success:
        logger.info("求職者一覧のエクスポート処理フローが正常に完了しました")
    else:
        logger.error("求職者一覧のエクスポート処理フローに失敗しました")
    
    # 処理間の待機時間
    logger.info("次の処理に進む前に10秒間待機します")
    time.sleep(10)
    
    # 選考プロセスの処理フロー
    logger.info("2. 選考プロセス一覧の表示処理フローを実行します")
    entryprocess_success = operations.access_selection_processes()
    
    if entryprocess_success:
        logger.info("選考プロセス一覧の表示処理フローが正常に完了しました")
    else:
        logger.error("選考プロセス一覧の表示処理フローに失敗しました")
    
    # 両方の処理結果をログに出力
    overall_success = candidates_success and entryprocess_success
    if overall_success:
        logger.info("✅ 両方の処理フローが正常に完了しました")
    else:
        logger.warning("⚠️ 一部の処理フローが失敗しました")
    
    return overall_success, (candidates_success, entryprocess_success)

def sequential_workflow(browser, login, **kwargs):
    """
    求職者一覧処理後に選考プロセスも実行する連続処理フロー
    
    Args:
        browser (PortersBrowser): ブラウザオブジェクト
        login (PortersLogin): ログインオブジェクト
        **kwargs: その他のパラメータ
        
    Returns:
        tuple: (success, (candidates_success, entryprocess_success)) 全体の成功/失敗と、各処理の成功/失敗
    """
    logger.info("=== 求職者一覧処理後に選考プロセスも実行します ===")
    operations = PortersOperations(browser)
    
    # PortersOperationsクラスの新しいメソッドを使用
    candidates_success, entryprocess_success = operations.execute_both_processes()
    
    # 両方の処理結果をログに出力
    overall_success = candidates_success and entryprocess_success
    if overall_success:
        logger.info("✅ 両方の処理フローが正常に完了しました")
    else:
        logger.warning("⚠️ 一部の処理フローが失敗しました")
    
    return overall_success, (candidates_success, entryprocess_success)

def run_aggregation(aggregate_option):
    """
    集計処理を実行する
    
    Args:
        aggregate_option (str): 実行する集計処理の種類
        
    Returns:
        bool: 集計処理が成功した場合はTrue、失敗した場合はFalse
    """
    logger.info("=== スプレッドシート集計処理を開始します ===")
    aggregator = SpreadsheetAggregator()
    
    # 集計処理を実行
    users_success, entryprocess_success = aggregator.run_aggregation(aggregate_option)
    
    # 結果の集計
    success_count = 0
    total_count = 0
    
    if aggregate_option in ['users', 'both']:
        total_count += 1
        if users_success:
            success_count += 1
            logger.info("✅ 求職者フェーズ別集計処理が正常に完了しました")
        else:
            logger.error("❌ 求職者フェーズ別集計処理に失敗しました")
    
    if aggregate_option in ['entryprocess', 'both']:
        total_count += 1
        if entryprocess_success:
            success_count += 1
            logger.info("✅ 選考プロセス集計処理が正常に完了しました")
        else:
            logger.error("❌ 選考プロセス集計処理に失敗しました")
    
    # 集計処理の結果ステータス
    overall_success = success_count == total_count
    if overall_success:
        logger.info("✅ すべての集計処理が正常に完了しました")
    else:
        logger.warning(f"⚠️ 実行された {total_count} 件の集計処理のうち、{success_count} 件が成功し、{total_count - success_count} 件が失敗しました")
    
    return overall_success, (users_success, entryprocess_success)

def main():
    """
    メイン処理
    
    処理フローの選択と引数の準備に特化し、実際の実行はPortersBrowserクラスに委譲します。
    
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
        
        # 処理成功フラグ
        success = True
        
        # 業務操作のスキップフラグがOFFの場合、PORTERSへログインして処理実行
        if not args.skip_operations:
            # 処理フローの選択
            workflow_func = None
            if args.process == 'candidates':
                workflow_func = candidates_workflow
            elif args.process == 'entryprocess':
                workflow_func = entryprocess_workflow
            elif args.process == 'both':
                workflow_func = both_workflow
            elif args.process == 'sequential':
                workflow_func = sequential_workflow
            
            # ワークフローパラメータの準備
            workflow_params = {
                'env': args.env,
                'process': args.process
            }
            
            # ワークフローセッションの実行
            success, workflow_results = PortersBrowser.execute_workflow_session(
                workflow_func=workflow_func,
                selectors_path=selectors_path,
                headless=args.headless,
                workflow_params=workflow_params
            )
        else:
            logger.info("業務操作をスキップします")
        
        # 集計処理の実行（オプションが指定されている場合）
        if args.aggregate != 'none':
            aggregate_success, _ = run_aggregation(args.aggregate)
            success = success and aggregate_success
        
        # 終了処理
        if success:
            logger.info("================================================================================")
            logger.info("PORTERSシステムログインツールを正常に終了します")
            logger.info("================================================================================")
            return 0
        else:
            logger.error("一部の処理に失敗しました")
            logger.info("================================================================================")
            logger.info("PORTERSシステムログインツールを異常終了します")
            logger.info("================================================================================")
            return 1
                
    except Exception as e:
        error_message = "予期しないエラーが発生しました"
        logger.error(f"{error_message}: {str(e)}")
        import traceback
        trace = traceback.format_exc()
        logger.error(trace)
        
        # Slack通知を送信
        slack.send_error(
            error_message=error_message,
            exception=e,
            title="PORTERSシステム予期しないエラー",
            context={
                "実行環境": args.env if 'args' in locals() else "不明",
                "処理フロー": args.process if 'args' in locals() else "不明",
                "集計処理": args.aggregate if 'args' in locals() else "不明",
                "traceback": trace[:1000] + ("..." if len(trace) > 1000 else "")
            }
        )
        
        logger.info("================================================================================")
        logger.info("PORTERSシステムログインツールを異常終了します")
        logger.info("================================================================================")
        return 1

if __name__ == "__main__":
    sys.exit(main())
