import time
import os
from pathlib import Path
import sys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

from src.utils.environment import EnvironmentUtils as env
from src.utils.logging_config import get_logger

logger = get_logger(__name__)

class PortersLogin:
    """
    PORTERSシステムへのログイン処理を管理するクラス
    
    このクラスは、PORTERSシステムへのログイン、二重ログイン回避、ログアウトなどの
    認証関連の処理を担当します。ブラウザ操作の基本機能はPortersBrowserクラスに依存します。
    """
    
    def __init__(self, browser):
        """
        ログイン処理クラスの初期化
        
        Args:
            browser (PortersBrowser): ブラウザ操作を管理するインスタンス
        """
        self.browser = browser
        self.screenshot_dir = browser.screenshot_dir
    
    @classmethod
    def login_to_porters(cls, selectors_path=None, headless=False):
        """
        PORTERSシステムへのログイン処理を実行する
        
        Args:
            selectors_path (str): セレクタ情報を含むCSVファイルのパス
            headless (bool): ヘッドレスモードで実行するかどうか
        
        Returns:
            tuple: (success, browser, login) 処理成功の場合はTrue、失敗の場合はFalse、およびブラウザとログインオブジェクト
        """
        from src.modules.porters.browser import PortersBrowser
        
        try:
            logger.info("=== PORTERSシステムへのログイン処理を開始します ===")
            
            # ブラウザセットアップ
            browser = PortersBrowser(selectors_path=selectors_path, headless=headless)
            
            # WebDriverのセットアップ
            if not browser.setup():
                logger.error("ブラウザのセットアップに失敗しました")
                return False, None, None
            
            # ログイン処理
            login = cls(browser)
            if not login.execute():
                logger.error("ログイン処理に失敗しました")
                browser.quit()
                return False, None, None
            
            # ログイン成功後の検証
            logger.info("ログイン後の画面を検証します")
            browser.save_screenshot("login_success_verification.png")
            
            logger.info("✅ PORTERSシステムへのログイン処理が正常に完了しました")
            return True, browser, login
            
        except Exception as e:
            logger.error(f"PORTERSシステムへのログイン処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            if 'browser' in locals() and browser:
                browser.quit()
            return False, None, None
    
    def execute(self):
        """
        ログイン処理を実行する
        
        Returns:
            bool: ログインが成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== ログイン処理を開始します ===")
            
            # 環境変数からログイン情報を取得
            admin_url = env.get_env_var('ADMIN_URL')
            admin_id = env.get_env_var('ADMIN_ID')  # 会社ID
            login_id = env.get_env_var('LOGIN_ID')
            password = env.get_env_var('LOGIN_PASSWORD')
            
            # ログインページにアクセス
            logger.info(f"ログインページにアクセスします: {admin_url}")
            self.browser.navigate_to(admin_url)
            
            # ログイン前のスクリーンショット
            self.browser.save_screenshot("login_before.png")
            
            # 会社ID入力
            company_id_field = self.browser.get_element('porters', 'company_id')
            if not company_id_field:
                logger.error("会社IDフィールドが見つかりません")
                return False
            
            company_id_field.clear()
            company_id_field.send_keys(admin_id)
            logger.info(f"✓ 会社IDを入力しました: {admin_id}")
            
            # ユーザー名入力
            username_field = self.browser.get_element('porters', 'username')
            if not username_field:
                logger.error("ユーザー名フィールドが見つかりません")
                return False
            
            username_field.clear()
            username_field.send_keys(login_id)
            logger.info("✓ ユーザー名を入力しました")
            
            # パスワード入力
            password_field = self.browser.get_element('porters', 'password')
            if not password_field:
                logger.error("パスワードフィールドが見つかりません")
                return False
            
            password_field.clear()
            password_field.send_keys(password)
            logger.info("✓ パスワードを入力しました")
            
            # 入力後のスクリーンショット
            self.browser.save_screenshot("login_input.png")
            
            # ログインボタンクリック
            login_button = self.browser.get_element('porters', 'login_button')
            if not login_button:
                logger.error("ログインボタンが見つかりません")
                return False
            
            login_button.click()
            logger.info("✓ ログインボタンをクリックしました")
            
            # ログイン処理待機
            time.sleep(5)
            
            # 二重ログインポップアップ対応
            self._handle_double_login_popup()
            
            # ログイン後のスクリーンショット
            self.browser.save_screenshot("login_after.png")
            
            # ログイン後のHTMLを解析
            after_login_html = self.browser.driver.page_source
            after_login_analysis = self.browser.analyze_page_content(after_login_html)
            
            # ログイン結果の詳細を記録
            logger.info(f"ログイン後の状態:")
            logger.info(f"  - タイトル: {after_login_analysis['page_title']}")
            logger.info(f"  - 見出し: {after_login_analysis['main_heading']}")
            logger.info(f"  - エラーメッセージ: {after_login_analysis['error_messages']}")
            logger.info(f"  - メニュー項目数: {len(after_login_analysis['menu_items'])}")
            if after_login_analysis['menu_items']:
                logger.info(f"  - メニュー項目例: {after_login_analysis['menu_items'][:5]}")
            
            # URLの変化も確認
            current_url = self.browser.driver.current_url
            logger.info(f"ログイン後のURL: {current_url}")
            
            # ログイン成功を判定
            login_success = (admin_url != current_url and "login" not in current_url.lower()) or len(after_login_analysis['menu_items']) > 0
            
            if login_success:
                logger.info("✅ ログインに成功しました！")
                return True
            else:
                logger.error("❌ ログインに失敗しました")
                
                # HTMLファイルとして保存（詳細分析用）
                html_path = os.path.join(self.screenshot_dir, "login_failed.html")
                with open(html_path, 'w', encoding='utf-8') as f:
                    f.write(after_login_html)
                logger.info(f"ログイン失敗時のHTMLを保存しました: {html_path}")
                
                return False
                
        except Exception as e:
            logger.error(f"ログイン処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return False
    
    def _handle_double_login_popup(self):
        """
        二重ログインポップアップの処理
        
        Returns:
            bool: ポップアップ処理が成功した場合はTrue、失敗した場合はFalse
        """
        # 二重ログインポップアップが表示されているか確認
        double_login_ok_button = "#pageDeny > div.ui-dialog.ui-widget.ui-widget-content.ui-corner-all.ui-front.p-ui-messagebox.ui-dialog-buttons.ui-draggable > div.ui-dialog-buttonpane.ui-widget-content.ui-helper-clearfix > div > button > span"
        
        try:
            # 短いタイムアウトで確認（ポップアップが表示されていない場合にテストが長時間停止しないように）
            popup_wait = WebDriverWait(self.browser.driver, 3)
            ok_button = popup_wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, double_login_ok_button)))
            
            # ポップアップが見つかった場合
            logger.info("⚠️ 二重ログインポップアップが検出されました。OKボタンをクリックします。")
            self.browser.save_screenshot("double_login_popup.png")
            
            # OKボタンをクリック
            ok_button.click()
            logger.info("✓ 二重ログインポップアップのOKボタンをクリックしました")
            time.sleep(2)
            return True
            
        except TimeoutException:
            # ポップアップが表示されていない場合は何もしない
            logger.info("二重ログインポップアップは表示されていません。処理を継続します。")
            return True
        except Exception as e:
            # その他のエラー
            logger.warning(f"二重ログインポップアップ処理中にエラーが発生しましたが、処理を継続します: {e}")
            
            # JavaScriptでのクリックを試行
            try:
                self.browser.driver.execute_script(f"document.querySelector('{double_login_ok_button}').click();")
                logger.info("✓ JavaScriptで二重ログインポップアップのOKボタンをクリックしました")
                time.sleep(2)
                return True
            except:
                logger.warning("JavaScriptでのクリックも失敗しましたが、処理を継続します")
                return False
    
    def logout(self):
        """
        明示的なログアウト処理を実行する
        
        Returns:
            bool: ログアウトが成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== ログアウト処理を開始します ===")
            
            # 現在のURLをログに記録
            current_url = self.browser.driver.current_url
            logger.info(f"ログアウト前のURL: {current_url}")
            
            # スクリーンショット
            self.browser.save_screenshot("before_logout.png")
            
            # ログアウトボタンを探す
            # まず、セレクタ情報を確認
            logout_selector = None
            if 'porters_menu' in self.browser.selectors and 'logout_button' in self.browser.selectors['porters_menu']:
                logout_selector = self.browser.selectors['porters_menu']['logout_button']['selector_value']
            
            if not logout_selector:
                # セレクタがない場合、一般的なログアウトボタンのセレクタを試す
                logger.info("ログアウトボタンのセレクタが設定されていません。一般的なセレクタを試します。")
                possible_selectors = [
                    'a[href*="logout"]', 
                    'button[id*="logout"]', 
                    'a[id*="logout"]',
                    '.logout', 
                    '#logout',
                    'a:contains("ログアウト")',
                    'button:contains("ログアウト")'
                ]
                
                # 各セレクタを試す
                for selector in possible_selectors:
                    try:
                        logout_elements = self.browser.driver.find_elements(By.CSS_SELECTOR, selector)
                        if logout_elements:
                            logger.info(f"ログアウトボタンを発見しました: {selector}")
                            logout_selector = selector
                            break
                    except:
                        continue
            
            if logout_selector:
                # ログアウトボタンをクリック
                try:
                    logger.info(f"ログアウトボタンを探索: {logout_selector}")
                    logout_button = self.browser.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, logout_selector)))
                    logout_button.click()
                    logger.info("✓ ログアウトボタンをクリックしました")
                    
                    # 確認ダイアログが表示される場合の処理
                    try:
                        confirm_buttons = self.browser.driver.find_elements(By.CSS_SELECTOR, 'button.confirm, button[id*="confirm"], button:contains("OK"), button:contains("はい")')
                        if confirm_buttons:
                            confirm_buttons[0].click()
                            logger.info("✓ 確認ダイアログのボタンをクリックしました")
                    except:
                        logger.info("確認ダイアログはありませんでした")
                    
                    # ログアウト後の待機
                    time.sleep(3)
                    
                    # ログアウト後のスクリーンショット
                    self.browser.save_screenshot("after_logout.png")
                    logger.info("✅ ログアウトに成功しました")
                    return True
                    
                except Exception as e:
                    logger.warning(f"通常のクリックでログアウトに失敗しました: {str(e)}")
                    # JavaScriptでログアウトを試みる
                    try:
                        self.browser.driver.execute_script(f"document.querySelector('{logout_selector}').click();")
                        logger.info("✓ JavaScriptを使用してログアウトボタンをクリックしました")
                        time.sleep(3)
                        self.browser.save_screenshot("after_js_logout.png")
                        logger.info("✅ JavaScriptでのログアウトに成功しました")
                        return True
                    except Exception as js_e:
                        logger.error(f"JavaScriptを使用したログアウトにも失敗しました: {str(js_e)}")
            else:
                # セレクタが見つからない場合、JavaScriptでログアウトを試みる
                logger.warning("ログアウトボタンが見つかりませんでした。JavaScriptでのログアウトを試みます。")
                try:
                    # 一般的なログアウトURLにリダイレクト
                    admin_url = env.get_env_var('ADMIN_URL')
                    base_url = admin_url.split('/index/login')[0]
                    logout_url = f"{base_url}/index/logout"
                    
                    self.browser.driver.get(logout_url)
                    logger.info(f"✓ ログアウトURLに直接アクセスしました: {logout_url}")
                    time.sleep(3)
                    self.browser.save_screenshot("after_direct_logout.png")
                    logger.info("✅ 直接URLアクセスでのログアウトに成功しました")
                    return True
                except Exception as url_e:
                    logger.error(f"直接URLアクセスでのログアウトにも失敗しました: {str(url_e)}")
            
            logger.error("❌ ログアウト処理に失敗しました")
            return False
            
        except Exception as e:
            logger.error(f"ログアウト処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return False 