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
    
    return parser.parse_args()

def main():
    """
    メイン処理
    
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
        
        # 環境設定
        if not setup_environment():
            logger.error("環境設定に失敗したため、処理を中止します")
            return 1
        
        # 設定ファイルのパス
        selectors_path = os.path.join(root_dir, "config", "selectors.csv")
        
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
                
                # 「その他業務」ボタンをクリックしてメニュー項目5を押す
                if not operations.execute_operations_flow():
                    logger.error("業務操作に失敗しました")
                else:
                    logger.info("業務操作が正常に完了しました")
                
                # 操作完了後の待機時間
                logger.info("操作完了後、5秒間待機します")
                time.sleep(5)
            
            # ログアウト処理
            if login:
                login.logout()
            
            # ブラウザを終了
            if browser:
                browser.quit()
                
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
        logger.error(f"予期しないエラーが発生しました: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        logger.info("================================================================================")
        logger.info("PORTERSシステムログインツールを異常終了します")
        logger.info("================================================================================")
        return 1

if __name__ == "__main__":
    sys.exit(main())
