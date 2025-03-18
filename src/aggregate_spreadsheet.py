#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
スプレッドシート集計処理の実行スクリプト

このスクリプトは、PORTERSシステムから取得したデータの集計処理を実行します。
1. 求職者情報（ユーザー）のフェーズ別集計
2. 選考プロセス情報の集計
上記の処理を単独またはまとめて実行することができます。
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
    parser = argparse.ArgumentParser(description='PORTERSデータの集計処理')
    parser.add_argument('--env', default='development', help='実行環境 (development または production)')
    parser.add_argument('--log-level', default='INFO', choices=['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'], 
                        help='ログレベル (DEBUG, INFO, WARNING, ERROR, CRITICAL)')
    parser.add_argument('--aggregation-type', choices=['users', 'entryprocess', 'both'], default='both',
                        help='実行する集計処理の種類 (users: 求職者フェーズ別集計, entryprocess: 選考プロセス集計, both: 両方を実行)')
    
    return parser.parse_args()

def main():
    """
    メイン処理
    
    以下の処理を実行します：
    1. 環境設定
    2. 指定された集計処理の実行
    
    Returns:
        int: 処理が成功した場合は0、失敗した場合は1
    """
    try:
        logger.info("================================================================================")
        logger.info("PORTERSデータ集計ツールを開始します")
        logger.info("================================================================================")
        
        # コマンドライン引数の解析
        args = parse_arguments()
        
        # 環境変数の設定
        os.environ['APP_ENV'] = args.env
        os.environ['LOG_LEVEL'] = args.log_level
        logger.info(f"実行環境: {args.env}")
        logger.info(f"ログレベル: {args.log_level}")
        logger.info(f"集計タイプ: {args.aggregation_type}")
        
        # 環境設定
        if not setup_environment():
            logger.error("環境設定に失敗したため、処理を中止します")
            return 1
        
        # 集計処理の実行
        logger.info("=== 集計処理を開始します ===")
        aggregator = SpreadsheetAggregator()
        
        users_success, entryprocess_success = aggregator.run_aggregation(args.aggregation_type)
        
        # 結果の集計
        success_count = 0
        total_count = 0
        
        if args.aggregation_type in ['users', 'both']:
            total_count += 1
            if users_success:
                success_count += 1
                logger.info("✅ 求職者フェーズ別集計処理が正常に完了しました")
            else:
                logger.error("❌ 求職者フェーズ別集計処理に失敗しました")
        
        if args.aggregation_type in ['entryprocess', 'both']:
            total_count += 1
            if entryprocess_success:
                success_count += 1
                logger.info("✅ 選考プロセス集計処理が正常に完了しました")
            else:
                logger.error("❌ 選考プロセス集計処理に失敗しました")
        
        # 終了ステータスの決定
        if success_count == total_count:
            logger.info("================================================================================")
            logger.info("PORTERSデータ集計ツールを正常に終了します")
            logger.info("================================================================================")
            return 0
        else:
            logger.warning(f"実行された {total_count} 件の処理のうち、{success_count} 件が成功し、{total_count - success_count} 件が失敗しました")
            logger.info("================================================================================")
            logger.info("PORTERSデータ集計ツールを異常終了します")
            logger.info("================================================================================")
            return 1
            
    except Exception as e:
        logger.error(f"予期しないエラーが発生しました: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        logger.info("================================================================================")
        logger.info("PORTERSデータ集計ツールを異常終了します")
        logger.info("================================================================================")
        return 1

if __name__ == "__main__":
    sys.exit(main()) 