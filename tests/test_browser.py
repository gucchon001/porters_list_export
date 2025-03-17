import os
import csv
import time
from pathlib import Path
from datetime import datetime
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, NoSuchElementException

from src.utils.logging_config import get_logger

logger = get_logger(__name__)

class TestBrowser:
    def __init__(self, selectors_path=None, headless=False, timeout=10):
        """ブラウザ操作を管理するクラス"""
        self.driver = None
        self.wait = None
        self.timeout = timeout
        self.headless = headless
        self.selectors_path = selectors_path
        self.selectors = {}
        
        # スクリーンショット保存ディレクトリ
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.screenshot_dir = os.path.join("logs", "screenshots", timestamp)
        os.makedirs(self.screenshot_dir, exist_ok=True)
        
        # セレクタファイルが指定されている場合は読み込む
        if selectors_path and os.path.exists(selectors_path):
            self._load_selectors()
    
    def setup(self):
        """WebDriverのセットアップ"""
        try:
            logger.info("WebDriverのセットアップを開始します")
            
            # Chromeオプションの設定
            chrome_options = Options()
            if self.headless:
                chrome_options.add_argument("--headless")
            
            chrome_options.add_argument("--window-size=1920,1080")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            
            # ダウンロード設定
            download_dir = os.path.join(os.getcwd(), "downloads")
            os.makedirs(download_dir, exist_ok=True)
            
            prefs = {
                "download.default_directory": download_dir,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": False
            }
            chrome_options.add_experimental_option("prefs", prefs)
            
            # WebDriverの初期化
            self.driver = webdriver.Chrome(options=chrome_options)
            self.driver.maximize_window()
            self.wait = WebDriverWait(self.driver, self.timeout)
            
            logger.info("✅ WebDriverのセットアップが完了しました")
            return True
            
        except Exception as e:
            logger.error(f"WebDriverのセットアップ中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return False
    
    def _load_selectors(self):
        """CSVファイルからセレクタ情報を読み込む"""
        try:
            logger.info(f"セレクタファイルを読み込みます: {self.selectors_path}")
            
            with open(self.selectors_path, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    if 'group' in row and 'name' in row and 'selector_type' in row and 'selector_value' in row:
                        group = row['group']
                        name = row['name']
                        
                        if group not in self.selectors:
                            self.selectors[group] = {}
                        
                        self.selectors[group][name] = {
                            'selector_type': row['selector_type'],
                            'selector_value': row['selector_value']
                        }
            
            logger.info(f"セレクタ情報を読み込みました: {len(self.selectors)} グループ")
            for group, selectors in self.selectors.items():
                logger.info(f"  - {group}: {len(selectors)} セレクタ")
            
            return True
            
        except Exception as e:
            logger.error(f"セレクタファイルの読み込み中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return False
    
    def navigate_to(self, url):
        """指定されたURLに移動する"""
        try:
            logger.info(f"URLに移動します: {url}")
            self.driver.get(url)
            return True
        except Exception as e:
            logger.error(f"URL移動中にエラーが発生しました: {str(e)}")
            return False
    
    def get_element(self, group, name, wait_time=None):
        """指定されたセレクタに一致する要素を取得する"""
        if not self.driver:
            logger.error("WebDriverが初期化されていません")
            return None
        
        if group not in self.selectors or name not in self.selectors[group]:
            logger.error(f"セレクタが見つかりません: {group}.{name}")
            return None
        
        selector_info = self.selectors[group][name]
        selector_type = selector_info['selector_type']
        selector_value = selector_info['selector_value']
        
        try:
            wait = WebDriverWait(self.driver, wait_time or self.timeout)
            
            if selector_type.lower() == 'css':
                element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector_value)))
            elif selector_type.lower() == 'xpath':
                element = wait.until(EC.presence_of_element_located((By.XPATH, selector_value)))
            elif selector_type.lower() == 'id':
                element = wait.until(EC.presence_of_element_located((By.ID, selector_value)))
            elif selector_type.lower() == 'name':
                element = wait.until(EC.presence_of_element_located((By.NAME, selector_value)))
            elif selector_type.lower() == 'class':
                element = wait.until(EC.presence_of_element_located((By.CLASS_NAME, selector_value)))
            else:
                logger.error(f"未対応のセレクタタイプです: {selector_type}")
                return None
            
            return element
            
        except TimeoutException:
            logger.warning(f"要素が見つかりませんでした: {group}.{name} ({selector_type}: {selector_value})")
            return None
        except Exception as e:
            logger.error(f"要素の取得中にエラーが発生しました: {str(e)}")
            return None
    
    def save_screenshot(self, filename):
        """スクリーンショットを保存する"""
        if not self.driver:
            logger.error("WebDriverが初期化されていません")
            return False
        
        try:
            filepath = os.path.join(self.screenshot_dir, filename)
            self.driver.save_screenshot(filepath)
            logger.info(f"スクリーンショットを保存しました: {filepath}")
            return True
        except Exception as e:
            logger.error(f"スクリーンショットの保存中にエラーが発生しました: {str(e)}")
            return False
    
    def analyze_page_content(self, html_content):
        """ページのHTML内容を解析する"""
        result = {
            'page_title': '',
            'main_heading': '',
            'error_messages': [],
            'menu_items': []
        }
        
        try:
            soup = BeautifulSoup(html_content, 'html.parser')
            
            # タイトルを取得
            title_tag = soup.find('title')
            if title_tag:
                result['page_title'] = title_tag.text.strip()
            
            # 主な見出しを取得
            h1_tags = soup.find_all('h1')
            if h1_tags:
                result['main_heading'] = h1_tags[0].text.strip()
            
            # エラーメッセージを探す
            error_elements = soup.find_all(class_=lambda c: c and ('error' in c.lower() or 'alert' in c.lower()))
            for error in error_elements:
                error_text = error.text.strip()
                if error_text:
                    result['error_messages'].append(error_text)
            
            # メニュー項目を探す
            menu_elements = soup.find_all(['a', 'button'], class_=lambda c: c and ('menu' in c.lower() or 'nav' in c.lower()))
            for menu in menu_elements:
                menu_text = menu.text.strip()
                if menu_text:
                    result['menu_items'].append(menu_text)
            
            # 一般的なナビゲーション要素も探す
            nav_elements = soup.find_all('nav')
            for nav in nav_elements:
                links = nav.find_all('a')
                for link in links:
                    link_text = link.text.strip()
                    if link_text and link_text not in result['menu_items']:
                        result['menu_items'].append(link_text)
            
            return result
            
        except Exception as e:
            logger.error(f"ページ内容の解析中にエラーが発生しました: {str(e)}")
            return result
    
    def quit(self):
        """WebDriverを終了する"""
        if self.driver:
            try:
                self.driver.quit()
                logger.info("WebDriverを終了しました")
            except Exception as e:
                logger.error(f"WebDriverの終了中にエラーが発生しました: {str(e)}") 