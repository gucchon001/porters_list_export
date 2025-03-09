import os
import sys
import time
from pathlib import Path
import traceback

from ...utils.environment import EnvironmentUtils as env
from ...utils.logging_config import get_logger
from .browser import Browser
from .login import Login
from .csv_import import CsvImport

logger = get_logger(__name__)

class Importer:
    def __init__(self):
        """PORTERSインポート処理を管理するクラス"""
        self.browser = None
    
    def setup(self):
        """
        ブラウザのセットアップを行う
        """
        try:
            logger.info("=== ブラウザのセットアップを開始します ===")
            
            # 環境変数のロード
            env.load_env()
            
            # プロジェクトのルートディレクトリを取得
            root_dir = env.get_project_root()
            
            # セレクタファイルのパスを取得
            selectors_path = os.path.join(root_dir, "config", "selectors.csv")
            
            # ブラウザインスタンスの作成
            self.browser = Browser(selectors_path=selectors_path)
            
            # ブラウザのセットアップ
            headless_mode = env.get_config_value("BROWSER", "HEADLESS", "False").lower() == "true"
            if not self.browser.setup(headless=headless_mode):
                logger.error("ブラウザのセットアップに失敗しました")
                return False
            
            logger.info("✅ ブラウザのセットアップが完了しました")
            return True
            
        except Exception as e:
            logger.error(f"ブラウザのセットアップ中にエラーが発生しました: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def execute(self):
        """
        PORTERSインポート処理を実行する
        """
        try:
            logger.info("=== PORTERSインポート処理を開始します ===")
            
            # ブラウザのセットアップ
            if not self.setup():
                logger.error("ブラウザのセットアップに失敗しました")
                return False
            
            try:
                # ログイン処理
                login = Login(self.browser)
                if not login.execute():
                    logger.error("ログイン処理に失敗しました")
                    return False
                
                # CSVインポート処理
                csv_import = CsvImport(self.browser)
                if not csv_import.execute():
                    logger.error("CSVインポート処理に失敗しました")
                    
                    # ログアウト処理
                    logger.info("ログアウト処理を実行します")
                    login.logout()
                    
                    return False
                
                # ログアウト処理
                logger.info("ログアウト処理を実行します")
                if not login.logout():
                    logger.warning("ログアウト処理に失敗しました")
                
                logger.info("✅ PORTERSインポート処理が完了しました")
                return True
                
            finally:
                # ブラウザを終了
                if self.browser:
                    self.browser.quit()
                
        except Exception as e:
            logger.error(f"PORTERSインポート処理中にエラーが発生しました: {str(e)}")
            logger.error(traceback.format_exc())
            return False

def import_to_porters():
    """
    PORTERSにCSVファイルをインポートする関数
    """
    importer = Importer()
    return importer.execute() 