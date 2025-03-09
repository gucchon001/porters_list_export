from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
from pathlib import Path
import pandas as pd
import logging

# Seleniumとurllib3の全ログを無効化
logging.getLogger('selenium').setLevel(logging.ERROR)
logging.getLogger('urllib3').setLevel(logging.ERROR)
logging.getLogger('WDM').setLevel(logging.ERROR)

from ..utils.environment import EnvironmentUtils as env
from ..utils.logging_config import get_logger

logger = get_logger(__name__)

class Browser:
    def __init__(self, settings_path='config/settings.ini', selectors_path='config/selectors.csv'):
        self.driver = None
        self.settings = env
        self.selectors = self._load_selectors(selectors_path)
        self.wait = None

    def _load_selectors(self, selectors_path):
        """セレクター設定の読み込み"""
        try:
            # ファイルの存在確認
            if not Path(selectors_path).exists():
                logger.error(f"セレクターファイルが見つかりません: {selectors_path}")
                return {}

            df = pd.read_csv(selectors_path)
            
            # 必要なカラムの確認
            required_columns = ['page', 'element', 'selector_type', 'selector_value']
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                logger.error(f"必要なカラムが不足しています: {missing_columns}")
                return {}

            selectors = {}
            for _, row in df.iterrows():
                if row['page'] not in selectors:
                    selectors[row['page']] = {}
                selectors[row['page']][row['element']] = {
                    'type': row['selector_type'],
                    'value': row['selector_value']
                }
            return selectors
            
        except Exception as e:
            logger.error(f"セレクター設定の読み込みに失敗: {str(e)}")
            logger.debug("詳細なエラー情報:", exc_info=True)
            return {}

    def _get_element(self, page, element, wait=30):
        """指定された要素を取得"""
        try:
            selector = self.selectors[page][element]
            by_type = getattr(By, selector['type'].upper())
            return WebDriverWait(self.driver, wait).until(
                EC.visibility_of_element_located((by_type, selector['value']))
            )
        except Exception as e:
            logger.error(f"要素の取得に失敗: {str(e)}")
            raise

    def setup(self):
        """ChromeDriverのセットアップ"""
        try:
            options = webdriver.ChromeOptions()
            
            # 基本設定
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('start-maximized')
            options.add_argument('--ignore-certificate-errors')
            options.add_argument('--ignore-ssl-errors')
            
            # GPUエラー対策
            options.add_argument('--disable-gpu')
            options.add_argument('--disable-software-rasterizer')
            
            # 直接bool型として扱う
            headless = self.settings.get_config_value('BROWSER', 'headless', default=False)
            
            if headless:
                options.add_argument('--headless=new')
                logger.info("ヘッドレスモードで実行します")
            else:
                logger.info("通常モードで実行します")
            
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)
            self.wait = WebDriverWait(self.driver, 20)
            self.driver.maximize_window()
            
        except Exception as e:
            logger.error(f"ブラウザのセットアップに失敗: {str(e)}")
            self.driver = None  # 確実にNoneに設定
            raise

    def get_url(self):
        """現在のURLを取得"""
        return self.driver.current_url

    def access_site(self, url):
        """指定したURLにアクセス"""
        self.driver.get(url)

    def open_csv_download_page(self):
        """CSVダウンロードページを開く"""
        try:
            csv_url = self.settings.get_config_value('URL', 'csv_download_url')
            self.driver.get(csv_url)
            time.sleep(2)  # ページの読み込みを待機
            return True
        except Exception as e:
            return False

    def quit(self):
        """ブラウザを終了"""
        try:
            if hasattr(self, 'driver') and self.driver is not None:
                self.driver.quit()
                logger.info("ブラウザを正常に終了しました")
        except Exception as e:
            logger.error(f"ブラウザの終了に失敗: {str(e)}") 