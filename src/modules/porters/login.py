import time
import os
from pathlib import Path
import sys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

from ...utils.environment import EnvironmentUtils as env
from ...utils.logging_config import get_logger

logger = get_logger(__name__)

class Login:
    def __init__(self, browser):
        """ログイン処理を管理するクラス"""
        self.browser = browser
        self.screenshot_dir = browser.screenshot_dir
    
    def execute(self):
        """
        ログイン処理を実行する
        """
        try:
            logger.info("=== ログイン処理を開始します ===")
            
            # 環境変数から認証情報を取得
            admin_url = env.get_config_value("PORTERS", "ADMIN_URL")
            company_id = env.get_env_var("PORTERS_COMPANY_ID")
            username = env.get_env_var("PORTERS_USERNAME")
            password = env.get_env_var("PORTERS_PASSWORD")
            
            if not all([admin_url, company_id, username, password]):
                logger.error("認証情報が不足しています")
                return False
            
            # 管理画面URLに移動
            if not self.browser.navigate_to(admin_url):
                logger.error(f"管理画面URL({admin_url})への移動に失敗しました")
                return False
            
            logger.info(f"管理画面URL({admin_url})に移動しました")
            self.browser.save_screenshot("login_page.png")
            
            # ログインフォームの要素を取得
            company_id_input = self.browser.get_element("porters", "company_id")
            username_input = self.browser.get_element("porters", "username")
            password_input = self.browser.get_element("porters", "password")
            
            if not all([company_id_input, username_input, password_input]):
                logger.error("ログインフォームの要素の取得に失敗しました")
                return False
            
            # 認証情報を入力
            company_id_input.clear()
            company_id_input.send_keys(company_id)
            logger.info("会社IDを入力しました")
            
            username_input.clear()
            username_input.send_keys(username)
            logger.info("ユーザー名を入力しました")
            
            password_input.clear()
            password_input.send_keys(password)
            logger.info("パスワードを入力しました")
            
            self.browser.save_screenshot("login_form_filled.png")
            
            # ログインボタンをクリック
            login_button = self.browser.get_element("porters", "login_button")
            if not login_button:
                # ログインボタンが見つからない場合、フォームをサブミット
                logger.warning("ログインボタンが見つかりません。フォームをサブミットします。")
                password_input.submit()
            else:
                login_button.click()
            
            logger.info("ログインボタンをクリックしました")
            
            # ログイン後のページ読み込みを待機
            try:
                WebDriverWait(self.browser.driver, 10).until(
                    lambda driver: "login" not in driver.current_url.lower()
                )
                logger.info("ログイン後のページ読み込みが完了しました")
            except TimeoutException:
                logger.error("ログイン後のページ読み込みがタイムアウトしました")
                self.browser.save_screenshot("login_timeout.png")
                return False
            
            # ログイン成功の確認
            self.browser.save_screenshot("after_login.png")
            
            # ページ内容を解析してログイン成功を確認
            page_content = self.browser.analyze_page_content()
            
            if page_content.get('error_messages'):
                logger.error(f"ログインエラーが発生しました: {page_content['error_messages']}")
                return False
            
            if page_content.get('welcome_message') or page_content.get('dashboard_elements'):
                logger.info("✅ ログインに成功しました")
                return True
            
            # URLでログイン成功を確認
            current_url = self.browser.driver.current_url
            if "login" not in current_url.lower():
                logger.info(f"現在のURL: {current_url}")
                logger.info("✅ URLからログイン成功を確認しました")
                return True
            
            logger.error("ログインに失敗しました")
            return False
            
        except Exception as e:
            logger.error(f"ログイン処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("login_error.png")
            return False
    
    def logout(self):
        """
        ログアウト処理を実行する
        """
        try:
            logger.info("=== ログアウト処理を開始します ===")
            
            # ログアウトボタンを探す
            logout_button = None
            
            # 一般的なログアウトボタンのセレクタを試す
            selectors = [
                (By.CSS_SELECTOR, "a.logout, button.logout, a[href*='logout'], a[href*='sign-out']"),
                (By.XPATH, "//a[contains(text(), 'ログアウト') or contains(text(), 'Logout')]"),
                (By.XPATH, "//button[contains(text(), 'ログアウト') or contains(text(), 'Logout')]"),
                (By.XPATH, "//a[contains(@href, 'logout') or contains(@href, 'sign-out')]")
            ]
            
            for selector_type, selector in selectors:
                try:
                    elements = self.browser.driver.find_elements(selector_type, selector)
                    if elements:
                        logout_button = elements[0]
                        logger.info(f"ログアウトボタンを見つけました: {selector}")
                        break
                except Exception:
                    continue
            
            if logout_button:
                try:
                    # ログアウトボタンをクリック
                    logout_button.click()
                    logger.info("ログアウトボタンをクリックしました")
                    
                    # ログアウト後のページ読み込みを待機
                    time.sleep(3)
                    self.browser.save_screenshot("after_logout.png")
                    
                    # ログアウト成功の確認
                    current_url = self.browser.driver.current_url
                    if "login" in current_url.lower():
                        logger.info("✅ ログアウトに成功しました")
                        return True
                    
                    logger.warning("ログアウト後のURLにloginが含まれていません")
                    
                    # JavaScriptでログアウトを試みる
                    try:
                        self.browser.driver.execute_script("window.location.href = '/logout' || '/auth/logout' || '/porters/logout';")
                        time.sleep(3)
                        logger.info("✅ JavaScriptでのログアウトに成功しました")
                        return True
                    except Exception as js_e:
                        logger.error(f"JavaScriptを使用したログアウトにも失敗しました: {str(js_e)}")
                except Exception as click_e:
                    logger.error(f"ログアウトボタンのクリックに失敗しました: {str(click_e)}")
                    
                    # JavaScriptでログアウトを試みる
                    try:
                        self.browser.driver.execute_script("window.location.href = '/logout' || '/auth/logout' || '/porters/logout';")
                        time.sleep(3)
                        logger.info("✅ JavaScriptでのログアウトに成功しました")
                        return True
                    except Exception as js_e:
                        logger.error(f"JavaScriptを使用したログアウトにも失敗しました: {str(js_e)}")
            else:
                # セレクタが見つからない場合、JavaScriptでログアウトを試みる
                logger.warning("ログアウトボタンが見つかりませんでした。JavaScriptでのログアウトを試みます。")
                try:
                    # POSTERSの一般的なログアウトURLパターンを試す
                    self.browser.driver.execute_script("window.location.href = '/logout' || '/auth/logout' || '/porters/logout';")
                    time.sleep(3)
                    logger.info("✓ JavaScriptでログアウトURLにリダイレクトしました")
                    self.browser.save_screenshot("after_redirect_logout.png")
                    return True
                except Exception as js_e:
                    logger.error(f"JavaScriptを使用したログアウトリダイレクトにも失敗しました: {str(js_e)}")
            
            logger.warning("ログアウト処理ができませんでした。ブラウザを閉じて強制終了します。")
            return False
            
        except Exception as e:
            logger.error(f"ログアウト処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return False 