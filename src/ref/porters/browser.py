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

from ...utils.environment import EnvironmentUtils as env
from ...utils.logging_config import get_logger

logger = get_logger(__name__)

class Browser:
    """ブラウザ制御クラス"""
    
    def __init__(self, settings_path=None, selectors_path=None):
        """ブラウザ制御クラスの初期化"""
        self.driver = None
        self.wait = None
        self.selectors = {}
        
        # プロジェクトのルートディレクトリを取得
        root_dir = env.get_project_root()
        
        # スクリーンショットディレクトリの設定
        self.screenshot_dir = os.path.join(root_dir, "logs", "screenshots")
        os.makedirs(self.screenshot_dir, exist_ok=True)
        
        # セレクタを読み込む
        if selectors_path:
            self.load_selectors(selectors_path)
        else:
            # デフォルトのセレクタパス
            default_selectors_path = os.path.join(root_dir, "config", "selectors.csv")
            if os.path.exists(default_selectors_path):
                self.load_selectors(default_selectors_path)
    
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
    
    def setup(self, headless=False):
        """WebDriverのセットアップ"""
        try:
            # Chromeオプションの設定
            chrome_options = Options()
            if headless:
                chrome_options.add_argument("--headless")
            chrome_options.add_argument("--window-size=1920,1080")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            
            # WebDriverのセットアップ
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            self.wait = WebDriverWait(self.driver, 10)  # 10秒のタイムアウト
            
            # セレクタが不足している場合、デフォルト値を設定
            if not self.selectors or 'porters' not in self.selectors:
                logger.warning("PORTERSのセレクタ情報が見つかりません。デフォルト値を使用します。")
                self.selectors['porters'] = {
                    'company_id': {
                        'selector_type': 'CSS_SELECTOR',
                        'selector_value': "#Model_LoginForm_company_login_id"
                    },
                    'username': {
                        'selector_type': 'CSS_SELECTOR',
                        'selector_value': "#Model_LoginForm_login_id"
                    },
                    'password': {
                        'selector_type': 'CSS_SELECTOR',
                        'selector_value': "#Model_LoginForm_password"
                    },
                    'login_button': {
                        'selector_type': 'CSS_SELECTOR',
                        'selector_value': "#yt0"
                    }
                }
                
            # porters_menuがなければ初期化
            if 'porters_menu' not in self.selectors:
                logger.warning("PORTERSメニューのセレクタ情報が見つかりません。デフォルト値を使用します。")
                self.selectors['porters_menu'] = {
                    'import_menu': {
                        'selector_type': 'CSS_SELECTOR',
                        'selector_value': "#main-menu-id-5 > a"
                    },
                    'menu_item_5': {
                        'selector_type': 'CSS_SELECTOR',
                        'selector_value': "#main-menu-id-5 > a"
                    },
                    'search_button': {
                        'selector_type': 'CSS_SELECTOR',
                        'selector_value': "#main > div > main > section.original-search > header > div.others > button"
                    }
                }
            
            logger.info("WebDriverのセットアップが完了しました")
            return True
        except Exception as e:
            logger.error(f"WebDriverのセットアップに失敗しました: {str(e)}")
            return False
    
    def navigate_to(self, url):
        """指定されたURLに移動"""
        try:
            self.driver.get(url)
            logger.info(f"URLに移動しました: {url}")
            return True
        except Exception as e:
            logger.error(f"URLへの移動に失敗しました: {url}, エラー: {str(e)}")
            return False
    
    def get_element(self, page, element_name):
        """指定されたページの要素を取得"""
        try:
            if page not in self.selectors or element_name not in self.selectors[page]:
                logger.error(f"セレクタが見つかりません: ページ={page}, 要素={element_name}")
                return None
            
            selector_info = self.selectors[page][element_name]
            selector_type = selector_info['selector_type'].upper()  # 大文字に変換
            selector_value = selector_info['selector_value']
            
            logger.info(f"要素を探索: {page}.{element_name} ({selector_type}: {selector_value})")
            
            # セレクタタイプに基づいて要素を検索 (By クラスの属性として使用)
            by_type = getattr(By, selector_type)
            element = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((by_type, selector_value))
            )
            
            logger.info(f"要素を取得しました: ページ={page}, 要素={element_name}")
            return element
        except Exception as e:
            logger.error(f"要素の取得に失敗しました: ページ={page}, 要素={element_name}, エラー: {str(e)}")
            return None
    
    def save_screenshot(self, filename):
        """スクリーンショットを保存"""
        try:
            screenshot_path = os.path.join(self.screenshot_dir, filename)
            self.driver.save_screenshot(screenshot_path)
            logger.info(f"スクリーンショットを保存しました: {screenshot_path}")
            return True
        except Exception as e:
            logger.error(f"スクリーンショットの保存に失敗しました: {str(e)}")
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