import time
import os
from pathlib import Path
import sys
import traceback
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException, NoSuchElementException

from ...utils.environment import EnvironmentUtils as env
from ...utils.logging_config import get_logger
from .browser import Browser
from .login import Login

logger = get_logger(__name__)

class CsvImport:
    def __init__(self, browser):
        """CSVインポート処理を管理するクラス"""
        self.browser = browser
        self.screenshot_dir = browser.screenshot_dir
    
    def execute(self):
        """
        CSVインポート処理を実行する
        """
        try:
            logger.info("=== CSVインポート処理を開始します ===")
            
            # 環境変数からCSVファイルパスを取得
            csv_path = env.get_config_value("PORTERS", "IMPORT_CSV_PATH")
            if not csv_path:
                logger.error("CSVファイルパスが設定されていません")
                return False
            
            # プロジェクトのルートディレクトリを取得
            root_dir = env.get_project_root()
            
            # 相対パスを絶対パスに変換
            if not os.path.isabs(csv_path):
                csv_path = os.path.join(root_dir, csv_path)
            
            if not os.path.exists(csv_path):
                logger.error(f"CSVファイルが存在しません: {csv_path}")
                return False
            
            logger.info(f"インポートするCSVファイル: {csv_path}")
            
            # インポートメニューに移動
            if not self.navigate_to_import_menu():
                logger.error("インポートメニューへの移動に失敗しました")
                return False
            
            # CSVファイルをアップロード
            if not self.upload_csv_file(csv_path):
                logger.error("CSVファイルのアップロードに失敗しました")
                return False
            
            # インポート設定を行う
            if not self.configure_import_settings():
                logger.error("インポート設定に失敗しました")
                return False
            
            # インポートを実行
            if not self.execute_import():
                logger.error("インポート実行に失敗しました")
                return False
            
            logger.info("✅ CSVインポート処理が完了しました")
            return True
            
        except Exception as e:
            logger.error(f"CSVインポート処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("csv_import_error.png")
            return False
    
    def navigate_to_import_menu(self):
        """
        インポートメニューに移動する
        """
        try:
            logger.info("インポートメニューに移動します")
            
            # メニュー項目をクリック
            menu_item = self.browser.get_element("porters_menu", "import_menu")
            if not menu_item:
                logger.warning("インポートメニュー項目が見つかりません。直接URLに移動を試みます。")
                
                # 直接インポートURLに移動
                import_url = env.get_config_value("PORTERS", "IMPORT_URL", "")
                if import_url:
                    return self.browser.navigate_to(import_url)
                
                # メニュー項目5をクリック（一般的なインポートメニュー）
                menu_item = self.browser.get_element("porters_menu", "menu_item_5")
                if not menu_item:
                    logger.error("メニュー項目が見つかりません")
                    self.browser.save_screenshot("menu_item_not_found.png")
                    return False
            
            # メニュー項目をクリック
            menu_item.click()
            logger.info("メニュー項目をクリックしました")
            time.sleep(2)
            self.browser.save_screenshot("after_menu_click.png")
            
            # インポートサブメニューをクリック（必要な場合）
            try:
                import_submenu = WebDriverWait(self.browser.driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'インポート') or contains(text(), 'Import')]"))
                )
                import_submenu.click()
                logger.info("インポートサブメニューをクリックしました")
                time.sleep(2)
                self.browser.save_screenshot("after_submenu_click.png")
            except TimeoutException:
                logger.info("インポートサブメニューは表示されませんでした")
            
            # 求職者インポートメニューをクリック（必要な場合）
            try:
                job_seeker_import = WebDriverWait(self.browser.driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), '求職者') or contains(text(), 'Job Seeker')]"))
                )
                job_seeker_import.click()
                logger.info("求職者インポートメニューをクリックしました")
                time.sleep(2)
                self.browser.save_screenshot("after_job_seeker_import_click.png")
            except TimeoutException:
                logger.info("求職者インポートメニューは表示されませんでした")
            
            # インポートページに移動したことを確認
            try:
                WebDriverWait(self.browser.driver, 10).until(
                    lambda driver: "インポート" in driver.page_source or "import" in driver.page_source.lower()
                )
                logger.info("インポートページに移動しました")
                self.browser.save_screenshot("import_page.png")
                return True
            except TimeoutException:
                logger.error("インポートページへの移動がタイムアウトしました")
                self.browser.save_screenshot("import_page_timeout.png")
                return False
            
        except Exception as e:
            logger.error(f"インポートメニューへの移動中にエラーが発生しました: {str(e)}")
            self.browser.save_screenshot("navigate_to_import_menu_error.png")
            return False
    
    def upload_csv_file(self, csv_path):
        """
        CSVファイルをアップロードする
        """
        try:
            logger.info(f"CSVファイルをアップロードします: {csv_path}")
            
            # ファイル選択ボタンを探す
            file_input = self.browser.get_element("porters_import", "file_input")
            if not file_input:
                logger.warning("ファイル選択ボタンが見つかりません。一般的な入力フィールドを探します。")
                
                # 一般的なファイル入力フィールドを探す
                try:
                    file_input = WebDriverWait(self.browser.driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file']"))
                    )
                    logger.info("ファイル入力フィールドを見つけました")
                except TimeoutException:
                    logger.error("ファイル入力フィールドが見つかりません")
                    self.browser.save_screenshot("file_input_not_found.png")
                    return False
            
            # CSVファイルをアップロード
            file_input.send_keys(csv_path)
            logger.info("CSVファイルをアップロードしました")
            time.sleep(2)
            self.browser.save_screenshot("after_file_upload.png")
            
            # アップロード完了を待機
            try:
                WebDriverWait(self.browser.driver, 30).until(
                    lambda driver: os.path.basename(csv_path) in driver.page_source
                )
                logger.info("ファイルアップロードの完了を確認しました")
                return True
            except TimeoutException:
                logger.warning("ファイル名がページに表示されませんでした。次のステップに進みます。")
                return True
            
        except Exception as e:
            logger.error(f"CSVファイルのアップロード中にエラーが発生しました: {str(e)}")
            self.browser.save_screenshot("upload_csv_file_error.png")
            return False
    
    def configure_import_settings(self):
        """
        インポート設定を行う
        """
        try:
            logger.info("インポート設定を行います")
            
            # 次へボタンをクリック（必要な場合）
            try:
                next_button = WebDriverWait(self.browser.driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '次へ') or contains(text(), 'Next')]"))
                )
                next_button.click()
                logger.info("次へボタンをクリックしました")
                time.sleep(2)
                self.browser.save_screenshot("after_next_button.png")
            except TimeoutException:
                logger.info("次へボタンは表示されませんでした")
            
            # 設定画面が表示されるまで待機
            try:
                WebDriverWait(self.browser.driver, 10).until(
                    lambda driver: "設定" in driver.page_source or "setting" in driver.page_source.lower() or "config" in driver.page_source.lower()
                )
                logger.info("設定画面が表示されました")
                self.browser.save_screenshot("settings_page.png")
            except TimeoutException:
                logger.warning("設定画面の表示がタイムアウトしました。現在のページで続行します。")
            
            # 設定オプションを選択（必要な場合）
            # ここでは一般的な設定を行います。実際のシステムに合わせて調整してください。
            
            # 次へボタンをクリック（必要な場合）
            try:
                next_button = WebDriverWait(self.browser.driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '次へ') or contains(text(), 'Next')]"))
                )
                next_button.click()
                logger.info("次へボタンをクリックしました")
                time.sleep(2)
                self.browser.save_screenshot("after_next_button_2.png")
            except TimeoutException:
                logger.info("次へボタンは表示されませんでした")
            
            # 確認画面が表示されるまで待機
            try:
                WebDriverWait(self.browser.driver, 10).until(
                    lambda driver: "確認" in driver.page_source or "confirm" in driver.page_source.lower()
                )
                logger.info("確認画面が表示されました")
                self.browser.save_screenshot("confirm_page.png")
            except TimeoutException:
                logger.warning("確認画面の表示がタイムアウトしました。現在のページで続行します。")
            
            return True
            
        except Exception as e:
            logger.error(f"インポート設定中にエラーが発生しました: {str(e)}")
            self.browser.save_screenshot("configure_import_settings_error.png")
            return False
    
    def execute_import(self):
        """
        インポートを実行する
        """
        try:
            logger.info("インポートを実行します")
            
            # インポート実行ボタンをクリック
            import_button = self.browser.get_element("porters_import", "import_button")
            if not import_button:
                logger.warning("インポート実行ボタンが見つかりません。一般的なボタンを探します。")
                
                # 一般的なインポートボタンを探す
                buttons = self.browser.driver.find_elements(By.TAG_NAME, "button")
                import_button = None
                
                for button in buttons:
                    button_text = button.text.strip().lower()
                    if "インポート" in button_text or "import" in button_text or "実行" in button_text or "execute" in button_text:
                        import_button = button
                        logger.info(f"インポート実行ボタンを見つけました: {button_text}")
                        break
                
                if not import_button:
                    logger.error("インポート実行ボタンが見つかりません")
                    self.browser.save_screenshot("import_button_not_found.png")
                    return False
            
            # インポート実行ボタンをクリック
            import_button.click()
            logger.info("インポート実行ボタンをクリックしました")
            time.sleep(2)
            self.browser.save_screenshot("after_import_button.png")
            
            # 確認ダイアログが表示される場合は確認ボタンをクリック
            try:
                confirm_button = WebDriverWait(self.browser.driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "button.confirm, button.ok, button[type='submit']"))
                )
                confirm_button.click()
                logger.info("確認ボタンをクリックしました")
                time.sleep(2)
                self.browser.save_screenshot("after_confirm_button.png")
            except TimeoutException:
                logger.info("確認ダイアログは表示されませんでした")
            
            # インポート完了を待機
            try:
                WebDriverWait(self.browser.driver, 60).until(
                    lambda driver: "完了" in driver.page_source or "success" in driver.page_source.lower() or "complete" in driver.page_source.lower()
                )
                logger.info("インポート完了を確認しました")
                self.browser.save_screenshot("import_complete.png")
                return True
            except TimeoutException:
                logger.warning("インポート完了の確認がタイムアウトしました。現在のページを確認します。")
                
                # 現在のページを確認
                page_content = self.browser.analyze_page_content()
                
                if page_content.get('error_messages'):
                    logger.error(f"インポートエラーが発生しました: {page_content['error_messages']}")
                    return False
                
                # エラーメッセージがなければ成功とみなす
                logger.info("エラーメッセージが見つからないため、インポートは成功したとみなします")
                return True
            
        except Exception as e:
            logger.error(f"インポート実行中にエラーが発生しました: {str(e)}")
            self.browser.save_screenshot("execute_import_error.png")
            
            # デバッグ情報を収集
            try:
                logger.info("=== デバッグ情報 ===")
                logger.info(f"現在のURL: {self.browser.driver.current_url}")
                
                # 画面上のボタンを列挙
                all_buttons = self.browser.driver.find_elements(By.TAG_NAME, "button")
                logger.info(f"画面上のボタン数: {len(all_buttons)}")
                for i, button in enumerate(all_buttons):
                    logger.info(f"ボタン{i+1}: テキスト='{button.text}', クラス='{button.get_attribute('class')}'")
                
                # HTMLを保存して後で分析できるようにする
                with open(os.path.join(self.screenshot_dir, "button_debug_html.html"), "w", encoding="utf-8") as f:
                    f.write(self.browser.driver.page_source)
            except Exception as debug_error:
                logger.error(f"デバッグ情報収集中にエラー: {debug_error}")
            
            return False

# import_to_porters() 関数は削除（importer.py に移動済み） 