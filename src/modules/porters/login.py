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
            
            # ユーザーメニューを開く処理
            user_menu_opened = False
            
            # 1. まず指定されたセレクタでユーザーメニューを探す
            try:
                logger.info("セレクタでユーザーメニューボタンを探索します")
                user_menu_selector = "#nav2-inner > div > ul > li.original-class-user > a > span"
                user_menu_element = WebDriverWait(self.browser.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, user_menu_selector))
                )
                user_menu_element.click()
                logger.info("✓ セレクタでユーザーメニューボタンをクリックしました")
                user_menu_opened = True
            except Exception as e:
                logger.warning(f"セレクタでのユーザーメニュークリックに失敗しました: {str(e)}")
            
            # 2. 「川島」テキストを含む要素を探す
            if not user_menu_opened:
                try:
                    logger.info("「川島」テキストを含む要素を探索します")
                    # XPathを使用して「川島」テキストを含む要素を検索
                    kawashima_xpath = "//*[contains(text(), '川島')]"
                    kawashima_elements = self.browser.driver.find_elements(By.XPATH, kawashima_xpath)
                    
                    if kawashima_elements:
                        logger.info(f"「川島」テキストを含む要素を {len(kawashima_elements)} 個発見しました")
                        for element in kawashima_elements:
                            try:
                                logger.info(f"「川島」テキスト要素をクリックします: {element.text}")
                                element.click()
                                logger.info("✓ 「川島」テキスト要素をクリックしました")
                                user_menu_opened = True
                                break
                            except Exception as click_e:
                                logger.warning(f"「川島」テキスト要素のクリックに失敗しました: {str(click_e)}")
                                # 親要素をクリックしてみる
                                try:
                                    parent = self.browser.driver.execute_script("return arguments[0].parentNode;", element)
                                    parent.click()
                                    logger.info("✓ 「川島」テキスト要素の親要素をクリックしました")
                                    user_menu_opened = True
                                    break
                                except Exception as parent_e:
                                    logger.warning(f"親要素のクリックにも失敗しました: {str(parent_e)}")
                except Exception as e:
                    logger.warning(f"「川島」テキスト要素の探索に失敗しました: {str(e)}")
            
            # 3. 一般的なユーザーメニューを探す
            if not user_menu_opened:
                try:
                    logger.info("一般的なユーザーメニューを探索します")
                    user_texts = ["ユーザー", "User", "user"]
                    for text in user_texts:
                        try:
                            user_xpath = f"//*[contains(text(), '{text}')]"
                            user_elements = self.browser.driver.find_elements(By.XPATH, user_xpath)
                            if user_elements:
                                logger.info(f"「{text}」テキストを含む要素を発見しました")
                                user_elements[0].click()
                                logger.info(f"✓ 「{text}」テキスト要素をクリックしました")
                                user_menu_opened = True
                                break
                        except Exception as text_e:
                            logger.warning(f"「{text}」テキスト要素のクリックに失敗しました: {str(text_e)}")
                except Exception as e:
                    logger.warning(f"一般的なユーザーメニューの探索に失敗しました: {str(e)}")
            
            # ユーザーメニューが開けなかった場合は直接ログアウトリンクを探す
            if not user_menu_opened:
                logger.warning("ユーザーメニューを開けませんでした。直接ログアウトリンクを探します。")
                try:
                    logger.info("直接ログアウトリンクを探索します")
                    logout_link_selector = "a[href*='logout']"
                    logout_link = WebDriverWait(self.browser.driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, logout_link_selector))
                    )
                    logout_link.click()
                    logger.info("✓ 直接ログアウトリンクをクリックしました")
                    time.sleep(3)
                    self.browser.save_screenshot("after_direct_logout_link.png")
                    
                    # ログアウト確認
                    if self._verify_logout():
                        return True
                except Exception as e:
                    logger.warning(f"直接ログアウトリンクのクリックに失敗しました: {str(e)}")
            
            # ユーザーメニューが開けた場合、ログアウトボタンをクリック
            if user_menu_opened:
                time.sleep(2)  # メニューが表示されるまで待機
                self.browser.save_screenshot("after_user_menu_open.png")
                
                # ログアウトボタンをクリック
                logout_clicked = False
                
                # まずセレクタでログアウトボタンを探す
                try:
                    logger.info("セレクタでログアウトボタンを探索します")
                    logout_selector = "#porters-contextmenu-column_1 > ul > li:nth-child(6)"
                    logout_element = WebDriverWait(self.browser.driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, logout_selector))
                    )
                    logout_element.click()
                    logger.info("✓ セレクタでログアウトボタンをクリックしました")
                    logout_clicked = True
                except Exception as e:
                    logger.warning(f"セレクタでのログアウトボタンクリックに失敗しました: {str(e)}")
                
                # テキストでログアウトボタンを探す
                if not logout_clicked:
                    try:
                        logger.info("テキストでログアウトボタンを探索します")
                        logout_texts = ["ログアウト", "Logout", "logout"]
                        for text in logout_texts:
                            try:
                                logout_xpath = f"//*[contains(text(), '{text}')]"
                                logout_elements = self.browser.driver.find_elements(By.XPATH, logout_xpath)
                                if logout_elements:
                                    logger.info(f"「{text}」テキストを含む要素を発見しました")
                                    logout_elements[0].click()
                                    logger.info(f"✓ 「{text}」テキスト要素をクリックしました")
                                    logout_clicked = True
                                    break
                            except Exception as text_e:
                                logger.warning(f"「{text}」テキスト要素のクリックに失敗しました: {str(text_e)}")
                    except Exception as e:
                        logger.warning(f"テキストでのログアウトボタン探索に失敗しました: {str(e)}")
                
                # href属性でログアウトリンクを探す
                if not logout_clicked:
                    try:
                        logger.info("href属性でログアウトリンクを探索します")
                        links = self.browser.driver.find_elements(By.TAG_NAME, "a")
                        for link in links:
                            try:
                                href = link.get_attribute("href")
                                if href and "logout" in href:
                                    logger.info(f"ログアウトを含むhrefを発見しました: {href}")
                                    link.click()
                                    logger.info("✓ ログアウトリンクをクリックしました")
                                    logout_clicked = True
                                    break
                            except:
                                continue
                    except Exception as e:
                        logger.warning(f"href属性でのログアウトリンク探索に失敗しました: {str(e)}")
                
                if logout_clicked:
                    # ログアウト後の待機
                    time.sleep(3)
                    self.browser.save_screenshot("after_logout.png")
                    
                    # ログアウト確認
                    if self._verify_logout():
                        return True
            
            # ここまでの方法でログアウトできなかった場合、直接ログアウトURLにアクセス
            logger.warning("通常の方法でログアウトできませんでした。直接ログアウトURLにアクセスします。")
            try:
                # 一般的なログアウトURLにリダイレクト
                admin_url = env.get_env_var('ADMIN_URL')
                base_url = admin_url.split('/index/login')[0]
                logout_url = f"{base_url}/index/logout"
                
                self.browser.driver.get(logout_url)
                logger.info(f"✓ ログアウトURLに直接アクセスしました: {logout_url}")
                time.sleep(3)
                self.browser.save_screenshot("after_direct_logout_url.png")
                
                # ログアウト確認
                if self._verify_logout():
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
    
    def _verify_logout(self):
        """
        ログアウトが正常に完了したかを確認する
        
        Returns:
            bool: ログアウトが成功した場合はTrue、失敗した場合はFalse
        """
        try:
            # ログイン画面に戻ったことを確認
            login_elements = self.browser.driver.find_elements(By.CSS_SELECTOR, "input[type='password']")
            if login_elements:
                logger.info("ログイン画面に戻ったことを確認しました")
                logger.info("✅ ログアウト処理が完了しました")
                return True
            else:
                # URLでログアウトを確認
                current_url = self.browser.driver.current_url
                logger.info(f"現在のURL: {current_url}")
                if "login" in current_url or "auth" in current_url:
                    logger.info("ログインページのURLを確認しました")
                    logger.info("✅ ログアウト処理が完了しました")
                    return True
                else:
                    logger.warning("ログイン画面への遷移が確認できませんでした")
                    return False
        except Exception as e:
            logger.warning(f"ログイン画面確認中にエラーが発生しました: {str(e)}")
            return False 