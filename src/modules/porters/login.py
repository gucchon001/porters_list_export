import time
import os
from pathlib import Path
import sys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

# プロジェクトのルートディレクトリをPYTHONPATHに追加
root_dir = Path(__file__).resolve().parent.parent.parent.parent
sys.path.append(str(root_dir))

from src.utils.environment import EnvironmentUtils as env
from src.utils.logging_config import get_logger

logger = get_logger(__name__)

class Login:
    def __init__(self, browser):
        """ログイン処理を管理するクラス"""
        self.browser = browser
        self.screenshot_dir = browser.screenshot_dir
    
    def execute(self):
        """ログイン処理を実行"""
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
            
            # ログイン結果の確認
            current_url = self.browser.driver.current_url
            logger.info(f"ログイン後のURL: {current_url}")
            
            # ログイン成功を判定
            login_success = (admin_url != current_url and "login" not in current_url.lower())
            
            if login_success:
                logger.info("✅ ログインに成功しました！")
                return True
            else:
                logger.error("❌ ログインに失敗しました")
                return False
                
        except Exception as e:
            logger.error(f"ログイン処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return False
    
    def _handle_double_login_popup(self):
        """二重ログインポップアップの処理"""
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
            
        except TimeoutException:
            # ポップアップが表示されていない場合は何もしない
            logger.info("二重ログインポップアップは表示されていません。処理を継続します。")
        except Exception as e:
            # その他のエラー
            logger.warning(f"二重ログインポップアップ処理中にエラーが発生しましたが、処理を継続します: {e}")
    
    def logout(self):
        """ログアウト処理を実行する"""
        try:
            logger.info("=== ログアウト処理を開始します ===")
            
            # 現在のURLを記録
            current_url = self.browser.driver.current_url
            logger.info(f"ログアウト前のURL: {current_url}")
            
            # スクリーンショットを撮る
            self.browser.save_screenshot("before_logout.png")
            
            # ユーザーメニューをクリック
            try:
                user_menu_selector = "#nav2-inner > div > ul > li.original-class-user > a > span"
                user_menu = self.browser.driver.find_element(By.CSS_SELECTOR, user_menu_selector)
                user_menu.click()
                logger.info("✓ ユーザーメニューをクリックしました")
                
                # メニューが表示されるまで少し待機
                time.sleep(2)
                
                # ログアウトオプションをクリック
                logout_option_selector = "#porters-contextmenu-column_1 > ul > li:nth-child(6)"
                logout_option = self.browser.driver.find_element(By.CSS_SELECTOR, logout_option_selector)
                logout_option.click()
                logger.info("✓ ログアウトオプションをクリックしました")
                
                # ログアウト完了を待機
                time.sleep(3)
                
                # ログアウト後のスクリーンショット
                self.browser.save_screenshot("after_logout.png")
                
                # ログアウト成功を確認（ログインページに戻っているか）
                if "login" in self.browser.driver.current_url.lower():
                    logger.info("✅ ログアウトに成功しました")
                    return True
                else:
                    logger.warning("ログアウト後もログインページに戻っていません")
                    return False
                
            except Exception as e:
                logger.warning(f"指定セレクタでのログアウト処理に失敗: {e}")
                
                # 代替方法: 一般的なセレクタを試す
                logger.info("ログアウトボタンのセレクタが設定されていません。一般的なセレクタを試します。")
                
                # 様々なセレクタを試す
                selectors = [
                    ".logout", 
                    "a.logout", 
                    "a[href*='logout']",
                    "button.logout",
                    "#logout",
                    "a:contains('ログアウト')",
                    "a:contains('Logout')"
                ]
                
                for selector in selectors:
                    try:
                        elements = self.browser.driver.find_elements(By.CSS_SELECTOR, selector)
                        if elements:
                            elements[0].click()
                            logger.info(f"✓ セレクタ '{selector}' でログアウトボタンをクリックしました")
                            time.sleep(3)
                            return True
                    except:
                        continue
                
                logger.warning("ログアウトボタンが見つかりませんでした")
                return False
                
        except Exception as e:
            logger.error(f"ログアウト処理中にエラーが発生しました: {str(e)}")
            return False 