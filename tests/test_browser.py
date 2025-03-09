import sys
import os
import time
import csv
from pathlib import Path
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup

# プロジェクトのルートディレクトリをPYTHONPATHに追加
root_dir = Path(__file__).parent.parent
sys.path.append(str(root_dir))

from src.utils.environment import EnvironmentUtils as env
from src.utils.logging_config import get_logger

logger = get_logger(__name__)

class TestBrowser:
    def __init__(self, settings_path=None, selectors_path=None):
        """ブラウザ制御クラスの初期化"""
        self.driver = None
        self.wait = None
        self.selectors = {}
        self.screenshot_dir = os.path.join(root_dir, "tests", "screenshots")
        os.makedirs(self.screenshot_dir, exist_ok=True)
        
        # セレクタを読み込む
        if selectors_path:
            self.load_selectors(selectors_path)
    
    def load_selectors(self, csv_path):
        """セレクタCSVファイルを読み込む"""
        try:
            with open(csv_path, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    page = row['page']
                    element = row['element']
                    
                    if page not in self.selectors:
                        self.selectors[page] = {}
                    
                    self.selectors[page][element] = {
                        'description': row['description'],
                        'action_type': row['action_type'],
                        'selector_type': row['selector_type'],
                        'selector_value': row['selector_value'],
                        'element_type': row['element_type'],
                        'parent_selector': row['parent_selector']
                    }
            logger.info(f"セレクタを読み込みました: {len(self.selectors)} ページ")
            return True
        except Exception as e:
            logger.error(f"セレクタの読み込みに失敗しました: {str(e)}")
            return False
    
    def setup(self):
        """WebDriverのセットアップ"""
        try:
            # ブラウザオプションの設定
            options = Options()
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('start-maximized')
            options.add_argument('--disable-gpu')
            
            # WebDriverのセットアップ
            logger.info("WebDriverをセットアップしています...")
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)
            self.wait = WebDriverWait(self.driver, 20)
            return True
        except Exception as e:
            logger.error(f"WebDriverのセットアップに失敗しました: {str(e)}")
            return False
    
    def get_element(self, page, element, wait_time=20):
        """指定ページの指定要素を取得"""
        try:
            if page not in self.selectors or element not in self.selectors[page]:
                logger.error(f"セレクタが見つかりません: page={page}, element={element}")
                return None
            
            selector_value = self.selectors[page][element]['selector_value']
            selector_type = self.selectors[page][element]['selector_type'].upper()
            
            logger.info(f"要素を探索: {page}.{element} ({selector_type}: {selector_value})")
            
            by_type = getattr(By, selector_type)
            element = WebDriverWait(self.driver, wait_time).until(
                EC.visibility_of_element_located((by_type, selector_value))
            )
            return element
        except Exception as e:
            logger.error(f"要素の取得に失敗: {str(e)}")
            return None
    
    def save_screenshot(self, filename):
        """スクリーンショットを保存"""
        try:
            filepath = os.path.join(self.screenshot_dir, filename)
            self.driver.save_screenshot(filepath)
            logger.info(f"スクリーンショットを保存: {filepath}")
            return True
        except Exception as e:
            logger.error(f"スクリーンショットの保存に失敗: {str(e)}")
            return False
    
    def navigate_to(self, url):
        """指定URLに移動"""
        try:
            logger.info(f"URLにアクセス: {url}")
            self.driver.get(url)
            time.sleep(3)  # ページロード待機
            return True
        except Exception as e:
            logger.error(f"URLへのアクセスに失敗: {str(e)}")
            return False
    
    def analyze_page_content(self, html_content=None):
        """Beautiful Soupを使用してページ内容を解析する"""
        if html_content is None:
            html_content = self.driver.page_source
            
        soup = BeautifulSoup(html_content, 'html.parser')
        result = {
            'page_title': soup.title.text if soup.title else 'No title',
            'main_heading': '',
            'error_messages': [],
            'menu_items': [],
            'welcome_message': '',
            'dashboard_elements': [],
        }
        
        # ページのメインの見出しを取得
        main_heading = soup.find('h1')
        if main_heading:
            result['main_heading'] = main_heading.text.strip()
        
        # エラーメッセージを探す
        error_elements = soup.find_all(['div', 'p', 'span'], class_=lambda c: c and ('error' in c.lower() or 'alert' in c.lower()))
        result['error_messages'] = [elem.text.strip() for elem in error_elements if elem.text.strip()]
        
        # メニュー項目を探す
        menu_items = soup.find_all(['a', 'li'], id=lambda x: x and ('menu' in x.lower() or 'nav' in x.lower()))
        result['menu_items'] = [item.text.strip() for item in menu_items if item.text.strip()]
        
        # ウェルカムメッセージを探す
        welcome_elements = soup.find_all(['div', 'p', 'span'], string=lambda s: s and ('welcome' in s.lower() or 'ようこそ' in s))
        if welcome_elements:
            result['welcome_message'] = welcome_elements[0].text.strip()
        
        # ダッシュボード要素を探す
        dashboard_elements = soup.find_all(['div', 'section'], class_=lambda c: c and ('dashboard' in c.lower() or 'summary' in c.lower()))
        result['dashboard_elements'] = [elem.get('id', 'No ID') for elem in dashboard_elements]
        
        return result
    
    def quit(self):
        """ブラウザを終了"""
        if self.driver:
            try:
                self.driver.quit()
                logger.info("ブラウザを終了しました")
                return True
            except Exception as e:
                logger.error(f"ブラウザの終了に失敗: {str(e)}")
                return False 