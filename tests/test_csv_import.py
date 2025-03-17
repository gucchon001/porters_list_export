import time
import os
from pathlib import Path
import sys

# プロジェクトのルートディレクトリをPYTHONPATHに追加
root_dir = Path(__file__).parent.parent
sys.path.append(str(root_dir))

from src.utils.environment import EnvironmentUtils as env
from src.utils.logging_config import get_logger
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

logger = get_logger(__name__)

class TestCsvImport:
    def __init__(self, browser):
        """CSVインポート処理を管理するクラス"""
        self.browser = browser
        self.screenshot_dir = browser.screenshot_dir
        self.download_dir = os.path.join(os.getcwd(), "downloads")
        os.makedirs(self.download_dir, exist_ok=True)
    
    def execute(self):
        """CSVインポート処理を実行"""
        try:
            logger.info("=== CSVインポート処理を開始します ===")
            
            # 求職者一覧ページに移動
            self._navigate_to_candidate_list()
            
            # CSVダウンロード
            if not self._download_csv():
                logger.error("CSVダウンロードに失敗しました")
                return False
            
            # CSVファイルの確認
            csv_file = self._find_latest_csv()
            if not csv_file:
                logger.error("ダウンロードしたCSVファイルが見つかりません")
                return False
            
            logger.info(f"CSVファイルを確認しました: {csv_file}")
            
            # スクリーンショット
            self.browser.save_screenshot("csv_import_complete.png")
            
            logger.info("✅ CSVインポート処理が正常に完了しました")
            return True
            
        except Exception as e:
            logger.error(f"CSVインポート処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return False
    
    def _navigate_to_candidate_list(self):
        """求職者一覧ページに移動"""
        try:
            logger.info("求職者一覧ページに移動します")
            
            # 求職者一覧リンクをクリック
            candidate_list_link = self.browser.get_element('porters_menu', 'candidate_list')
            if not candidate_list_link:
                logger.warning("求職者一覧リンクが見つかりません。URLで直接アクセスを試みます。")
                admin_url = env.get_env_var('ADMIN_URL')
                base_url = admin_url.split('/index/login')[0]
                candidate_list_url = f"{base_url}/candidate/list"
                self.browser.navigate_to(candidate_list_url)
            else:
                candidate_list_link.click()
            
            # ページ読み込み待機
            time.sleep(5)
            
            # スクリーンショット
            self.browser.save_screenshot("candidate_list.png")
            
            logger.info("✓ 求職者一覧ページに移動しました")
            return True
            
        except Exception as e:
            logger.error(f"求職者一覧ページへの移動中にエラーが発生しました: {str(e)}")
            return False
    
    def _download_csv(self):
        """CSVをダウンロード"""
        try:
            logger.info("CSVダウンロードを開始します")
            
            # CSVダウンロードボタンをクリック
            csv_download_button = self.browser.get_element('porters_menu', 'csv_download')
            if not csv_download_button:
                logger.error("CSVダウンロードボタンが見つかりません")
                return False
            
            # スクリーンショット
            self.browser.save_screenshot("before_csv_download.png")
            
            # ボタンクリック
            csv_download_button.click()
            logger.info("✓ CSVダウンロードボタンをクリックしました")
            
            # ダウンロード待機
            time.sleep(5)
            
            # スクリーンショット
            self.browser.save_screenshot("after_csv_download.png")
            
            logger.info("✓ CSVダウンロードが完了しました")
            return True
            
        except Exception as e:
            logger.error(f"CSVダウンロード中にエラーが発生しました: {str(e)}")
            return False
    
    def _find_latest_csv(self):
        """最新のCSVファイルを検索"""
        try:
            csv_files = [f for f in os.listdir(self.download_dir) if f.lower().endswith('.csv')]
            if not csv_files:
                return None
            
            # 最新のファイルを取得
            latest_file = max(csv_files, key=lambda f: os.path.getmtime(os.path.join(self.download_dir, f)))
            return os.path.join(self.download_dir, latest_file)
            
        except Exception as e:
            logger.error(f"CSVファイルの検索中にエラーが発生しました: {str(e)}")
            return None 