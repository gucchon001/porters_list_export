import time
import os
from pathlib import Path
import sys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

# プロジェクトのルートディレクトリをPYTHONPATHに追加
root_dir = Path(__file__).resolve().parent.parent.parent.parent
sys.path.append(str(root_dir))

from src.utils.environment import EnvironmentUtils as env
from src.utils.logging_config import get_logger

logger = get_logger(__name__)

class CsvImport:
    def __init__(self, browser):
        """CSVインポート処理を管理するクラス"""
        self.browser = browser
        self.screenshot_dir = browser.screenshot_dir
    
    def execute(self, csv_file_path=None):
        """
        CSVインポート処理を実行する
        """
        try:
            logger.info("=== CSVインポート処理を開始します ===")
            
            # CSVファイルパスが指定されていない場合は、インスタンス変数を使用
            if csv_file_path is None:
                csv_file_path = self.csv_file_path
            
            logger.info(f"インポートするCSVファイル: {csv_file_path}")
            
            # インポートメニューを開く
            if not self._open_import_menu():
                logger.error("インポートメニューを開けませんでした")
                return False
            
            # ファイルをアップロード
            if not self._upload_csv_file(csv_file_path):
                logger.error("CSVファイルのアップロードに失敗しました")
                return False
            
            # インポート方法を選択
            if not self._select_import_method():
                logger.error("インポート方法の選択に失敗しました")
                return False
            
            # 「次へ」ボタンをクリック
            if not self._click_next_button():
                logger.error("「次へ」ボタンのクリックに失敗しました")
                return False
            
            # 3秒待機して画面遷移を確認
            time.sleep(3)
            
            # 現在のURLを確認
            current_url = self.browser.driver.current_url
            logger.info(f"「次へ」ボタンクリック後のURL: {current_url}")
            
            # スクリーンショットを取得
            self.browser.save_screenshot("after_next_button.png")
            
            # ダイアログが閉じてしまった場合（カレンダー画面に戻った場合）
            if "calendar" in current_url and not self._is_import_dialog_visible():
                logger.error("インポートダイアログが閉じてしまいました。処理が中断されました。")
                return False
            
            # インポート実行ボタンをクリック
            if not self._click_import_button():
                logger.error("インポート実行ボタンのクリックに失敗しました")
                return False
            
            # インポート結果を確認
            if not self._check_import_result():
                logger.error("インポート結果の確認に失敗しました")
                return False
            
            logger.info("✅ CSVインポート処理が正常に完了しました")
            return True
            
        except Exception as e:
            logger.error(f"CSVインポート処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("csv_import_error.png")
            return False

    def _is_import_dialog_visible(self):
        """
        インポートダイアログが表示されているかを確認する
        """
        try:
            # ダイアログの存在を確認
            dialogs = self.browser.driver.find_elements(By.CSS_SELECTOR, ".ui-dialog")
            if not dialogs:
                logger.warning("画面上にダイアログが見つかりません")
                return False
            
            # ダイアログのタイトルを確認
            for dialog in dialogs:
                try:
                    title_elem = dialog.find_element(By.CSS_SELECTOR, ".ui-dialog-title")
                    title = title_elem.text
                    if "求職者 - インポート" in title:
                        logger.info(f"インポートダイアログを確認: {title}")
                        return True
                except:
                    continue
            
            # HTMLを保存して後で分析できるようにする
            html_path = os.path.join(self.browser.screenshot_dir, "dialog_check.html")
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(self.browser.driver.page_source)
            
            logger.warning("インポートダイアログが見つかりません")
            return False
            
        except Exception as e:
            logger.error(f"ダイアログ確認中にエラー: {str(e)}")
            return False

    def _open_import_menu(self):
        """インポートメニューを開く"""
        try:
            logger.info("インポートメニューを探索します")
            
            # 現在のウィンドウハンドルを保存
            current_handles = self.browser.driver.window_handles
            logger.info(f"操作前のウィンドウハンドル一覧: {current_handles}")
            
            # 「その他業務」ボタンをクリック
            if not self._click_other_operations_button(current_handles):
                logger.error("「その他業務」ボタンのクリックに失敗しました")
                return False
            
            # メニュー項目5をクリック
            if not self._click_menu_item_5():
                logger.error("メニュー項目5のクリックに失敗しました")
                return False
            
            # 「求職者のインポート」リンクをクリック
            if not self._click_import_link():
                logger.error("「求職者のインポート」リンクのクリックに失敗しました")
                return False
            
            logger.info("✅ インポートメニューの操作が完了しました")
            return True
            
        except Exception as e:
            logger.error(f"インポートメニューの探索中にエラーが発生しました: {str(e)}")
            self.browser.save_screenshot("import_menu_error.png")
            return False
    
    def _upload_csv_file(self, csv_path):
        """CSVファイルをアップロードする"""
        try:
            # 方法1: 直接input[type="file"]要素を探す
            try:
                file_input = self.browser.driver.find_element(By.CSS_SELECTOR, 'input[type="file"]')
                file_input.send_keys(csv_path)
                logger.info(f"✓ ファイル入力要素にCSVパスを設定しました: {csv_path}")
                time.sleep(2)
                return True
            except NoSuchElementException:
                logger.warning("直接的なファイル入力要素が見つかりませんでした")
            
            # 方法2: 「添付」ボタンをクリックしてから隠れたinput要素を操作
            try:
                # 「添付」ボタンを探す
                attach_buttons = self.browser.driver.find_elements(By.XPATH, "//button[contains(text(), '添付') or contains(text(), 'ファイル選択')]")
                if attach_buttons:
                    # ボタンをクリック
                    attach_buttons[0].click()
                    logger.info("✓ 「添付」ボタンをクリックしました")
                    time.sleep(1)
                    
                    # 隠れたinput要素を探す
                    hidden_inputs = self.browser.driver.find_elements(By.CSS_SELECTOR, 'input[type="file"]')
                    if hidden_inputs:
                        # JavaScriptで表示状態を変更
                        self.browser.driver.execute_script("arguments[0].style.display = 'block';", hidden_inputs[0])
                        hidden_inputs[0].send_keys(csv_path)
                        logger.info(f"✓ 隠れたファイル入力要素にCSVパスを設定しました: {csv_path}")
                        time.sleep(2)
                        return True
            except Exception as e:
                logger.warning(f"「添付」ボタン方式でのアップロードに失敗しました: {str(e)}")
            
            # 方法3: ドラッグ&ドロップエリアを探す
            try:
                drop_area = self.browser.driver.find_element(By.CSS_SELECTOR, '.dropzone, [id*="drop"], [class*="drop"]')
                if drop_area:
                    # JavaScriptでファイルをドロップする代わりにinput要素を作成
                    self.browser.driver.execute_script("""
                        var input = document.createElement('input');
                        input.type = 'file';
                        input.style.display = 'block';
                        input.style.position = 'absolute';
                        input.style.top = '0';
                        input.style.left = '0';
                        document.body.appendChild(input);
                        return input;
                    """)
                    
                    # 作成したinput要素を取得
                    file_input = self.browser.driver.find_element(By.CSS_SELECTOR, 'input[type="file"][style*="position: absolute"]')
                    file_input.send_keys(csv_path)
                    logger.info(f"✓ 作成したファイル入力要素にCSVパスを設定しました: {csv_path}")
                    time.sleep(2)
                    return True
            except Exception as e:
                logger.warning(f"ドラッグ&ドロップエリア方式でのアップロードに失敗しました: {str(e)}")
            
            logger.error("すべての方法でCSVファイルのアップロードに失敗しました")
            self.browser.save_screenshot("file_upload_failed.png")
            return False
            
        except Exception as e:
            logger.error(f"CSVファイルのアップロード中にエラーが発生しました: {str(e)}")
            self.browser.save_screenshot("file_upload_error.png")
            return False
    
    def _select_import_method(self):
        """インポート方法を選択する（LINE初回アンケート取込）"""
        try:
            logger.info("=== インポート方法の選択処理を開始 ===")
            
            # スクリーンショットで状態を確認
            self.browser.save_screenshot("before_import_method.png")
            
            # standalone_test.pyと同じセレクタを使用
            import_method_selector = "#porters-pdialog_1 > div > div.subWrap.resize > div > div > div > ul > li:nth-child(9) > label > input[type=radio]"
            
            try:
                import_method = self.browser.driver.find_element(By.CSS_SELECTOR, import_method_selector)
                self.browser.driver.execute_script("arguments[0].click();", import_method)
                logger.info("✓ 「LINE初回アンケート取込」を選択しました")
                time.sleep(2)
                self.browser.save_screenshot("import_method_selected.png")
                return True
            except Exception as e:
                logger.warning(f"指定セレクタでのインポート方法選択に失敗: {e}")
                
                # 代替方法: すべてのラジオボタンを調査
                try:
                    radio_buttons = self.browser.driver.find_elements(By.CSS_SELECTOR, "input[type='radio']")
                    logger.info(f"画面上のラジオボタン数: {len(radio_buttons)}")
                    
                    # 9番目のラジオボタン（LINE初回アンケート取込）を選択
                    if len(radio_buttons) >= 9:
                        self.browser.driver.execute_script("arguments[0].click();", radio_buttons[8])  # 0-indexedで9番目
                        logger.info("✓ 9番目のラジオボタンを選択しました")
                        time.sleep(2)
                        return True
                    elif radio_buttons:
                        # 最後のラジオボタンを選択
                        self.browser.driver.execute_script("arguments[0].click();", radio_buttons[-1])
                        logger.info(f"✓ 最後のラジオボタン（{len(radio_buttons)}番目）を選択しました")
                        time.sleep(2)
                        return True
                except Exception as e2:
                    logger.warning(f"ラジオボタンの調査にも失敗: {e2}")
            
            logger.warning("インポート方法の選択に失敗しましたが、処理を続行します")
            return False
            
        except Exception as e:
            logger.warning(f"インポート方法の選択中にエラーが発生しました: {str(e)}")
            self.browser.save_screenshot("import_method_error.png")
            return False
    
    def _click_next_button(self):
        """「次へ」ボタンをクリックする"""
        try:
            logger.info("=== 「次へ」ボタンのクリック処理を開始 ===")
            
            # スクリーンショットを取得
            self.browser.save_screenshot("before_next_button.png")
            
            # 方法1: 直接ボタンを探す
            try:
                # ボタンパネルを探す
                button_pane = self.browser.driver.find_element(By.CSS_SELECTOR, ".ui-dialog-buttonpane")
                buttons = button_pane.find_elements(By.TAG_NAME, "button")
                
                # ボタンの情報をログに出力
                for i, btn in enumerate(buttons):
                    logger.info(f"ボタン {i+1}: テキスト='{btn.text}', クラス={btn.get_attribute('class')}")
                
                # 「次へ」ボタンを探す
                next_button = None
                for btn in buttons:
                    if btn.text == "次へ":
                        next_button = btn
                        break
                
                if next_button:
                    # スクロールして表示
                    self.browser.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                    time.sleep(1)
                    
                    # クリック
                    next_button.click()
                    logger.info("✓ 「次へ」ボタンをクリックしました")
                    
                    # 画面遷移を待機（長めに設定）
                    time.sleep(10)
                    
                    # 画面3への遷移を確認（より柔軟な条件）
                    page_source = self.browser.driver.page_source
                    if "求職者 - インポート (3/4)" in page_source or "インポート (3" in page_source:
                        logger.info("✓ 画面3への遷移を確認しました")
                        self.browser.save_screenshot("screen3_displayed.png")
                        return True
                    else:
                        logger.warning("画面3への遷移を確認できませんでした")
                        # HTMLを保存して後で分析
                        with open("screen_after_next.html", "w", encoding="utf-8") as f:
                            f.write(page_source)
                else:
                    logger.warning("「次へ」ボタンが見つかりませんでした")
            except Exception as e:
                logger.warning(f"ボタンパネルからのボタン検索に失敗: {e}")
            
            # 方法2: JavaScriptで直接実行
            try:
                logger.info("JavaScriptで直接2番目のボタンをクリックします")
                script = """
                var buttonPane = document.querySelector('.ui-dialog-buttonpane');
                if (buttonPane) {
                    var buttons = buttonPane.querySelectorAll('button');
                    if (buttons.length >= 2) {
                        buttons[1].click();
                        return true;
                    }
                }
                return false;
                """
                result = self.browser.driver.execute_script(script)
                if result:
                    logger.info("✓ JavaScriptで2番目のボタンのクリックに成功しました")
                    time.sleep(10)  # 長めに待機
                    
                    # 画面3への遷移を確認（より柔軟な条件）
                    page_source = self.browser.driver.page_source
                    if "求職者 - インポート (3/4)" in page_source or "インポート (3" in page_source:
                        logger.info("✓ 画面3への遷移を確認しました")
                        self.browser.save_screenshot("screen3_displayed.png")
                        return True
                    else:
                        logger.warning("JavaScriptでのクリック後も画面3への遷移を確認できませんでした")
                        # HTMLを保存して後で分析
                        with open("screen_after_js_next.html", "w", encoding="utf-8") as f:
                            f.write(page_source)
                else:
                    logger.warning("JavaScriptでのクリックも失敗しました")
            except Exception as js_error:
                logger.error(f"JavaScriptでのクリックにも失敗: {js_error}")
            
            # 現在のURLを確認
            current_url = self.browser.driver.current_url
            logger.info(f"現在のURL: {current_url}")
            
            # カレンダー画面に戻ってしまった場合
            if "calendar" in current_url:
                logger.error("ボタンクリック後にカレンダー画面に戻ってしまいました")
                self.browser.save_screenshot("calendar_redirect.png")
                return False
            
            # 画面状態を確認
            self.browser.save_screenshot("after_next_button.png")
            
            # すべての方法が失敗した場合
            logger.error("すべての方法で画面2の「次へ」ボタンのクリックに失敗しました")
            return False
        
        except Exception as e:
            logger.error(f"「次へ」ボタンのクリック中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("next_button_error.png")
            return False
    
    def _click_import_button(self):
        """インポート実行ボタンをクリックする"""
        try:
            logger.info("=== インポート実行ボタンのクリック処理を開始 ===")
            
            # 現在の画面状態を確認
            current_url = self.browser.driver.current_url
            logger.info(f"現在のURL: {current_url}")
            
            # スクリーンショットで状態を確認
            self.browser.save_screenshot("before_import_button.png")
            
            # HTMLを保存して詳細分析
            html_path = os.path.join(self.browser.screenshot_dir, "import_dialog.html")
            with open(html_path, 'w', encoding='utf-8') as f:
                f.write(self.browser.driver.page_source)
            logger.info(f"現在の画面HTMLを保存しました: {html_path}")
            
            # ダイアログの状態を確認
            dialogs = self.browser.driver.find_elements(By.CSS_SELECTOR, ".ui-dialog")
            logger.info(f"画面上のダイアログ数: {len(dialogs)}")
            
            if dialogs:
                # 最後のダイアログ（最前面）を使用
                dialog = dialogs[-1]
                dialog_id = dialog.get_attribute("id")
                dialog_class = dialog.get_attribute("class")
                logger.info(f"操作対象ダイアログ: ID={dialog_id}, クラス={dialog_class}")
                
                # ダイアログのタイトルを取得
                try:
                    title_element = dialog.find_element(By.CSS_SELECTOR, ".ui-dialog-title")
                    dialog_title = title_element.text
                    logger.info(f"ダイアログタイトル: {dialog_title}")
                except:
                    logger.info("ダイアログタイトルは取得できませんでした")
                
                # ボタンパネルを探す
                try:
                    button_pane = dialog.find_element(By.CSS_SELECTOR, ".ui-dialog-buttonpane")
                    logger.info("ダイアログのボタンパネルを見つけました")
                    
                    # ボタンパネル内のすべてのボタン要素を取得
                    buttons = button_pane.find_elements(By.TAG_NAME, "button")
                    logger.info(f"ボタンパネル内のボタン数: {len(buttons)}")
                    
                    # すべてのボタンの情報を記録
                    for i, btn in enumerate(buttons):
                        try:
                            btn_text = btn.text
                            btn_class = btn.get_attribute("class")
                            logger.info(f"ボタン {i+1}: テキスト='{btn_text}', クラス={btn_class}")
                        except:
                            logger.info(f"ボタン {i+1}: [情報取得不可]")
                    
                    # 「実行」ボタンを探す
                    import_button = None
                    for btn in buttons:
                        if btn.text == "実行":
                            import_button = btn
                            break
                    
                    if import_button:
                        # ボタンが無効化されていないか確認
                        if "ui-button-disabled" in import_button.get_attribute("class"):
                            logger.warning("「実行」ボタンが無効化されています。「次へ」ボタンをクリックして画面4に進みます。")
                            
                            # 「次へ」ボタンを探す
                            next_button = None
                            for btn in buttons:
                                if btn.text == "次へ":
                                    next_button = btn
                                    break
                            
                            if next_button:
                                # スクロールして表示
                                self.browser.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                                time.sleep(1)
                                
                                # クリック
                                next_button.click()
                                logger.info("✓ 画面3の「次へ」ボタンをクリックしました")
                                
                                # 画面遷移を待機
                                time.sleep(10)
                                
                                # 画面4への遷移を確認
                                if "求職者 - インポート (4/4)" in self.browser.driver.page_source or "インポート (4" in self.browser.driver.page_source:
                                    logger.info("✓ 画面4への遷移を確認しました")
                                    self.browser.save_screenshot("screen4_displayed.png")
                                    
                                    # 画面4で「実行」ボタンを探す
                                    return self._click_execute_button_on_screen4()
                                else:
                                    logger.warning("画面4への遷移を確認できませんでした")
                            else:
                                logger.warning("「次へ」ボタンが見つかりませんでした")
                        else:
                            # スクロールして表示
                            self.browser.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", import_button)
                            time.sleep(1)
                            
                            # クリック
                            import_button.click()
                            logger.info("✓ 「実行」ボタンをクリックしました")
                            
                            # 処理完了を待機
                            time.sleep(5)
                            
                            # 結果を確認
                            return self._check_import_result()
                    else:
                        logger.warning("「実行」ボタンが見つかりませんでした")
                except Exception as e:
                    logger.warning(f"ボタンパネルからのボタン検索に失敗: {e}")
            else:
                logger.warning("ダイアログが見つかりません")
        except Exception as e:
            logger.warning(f"直接ボタン検索に失敗: {e}")
        
        # 方法2: JavaScriptで直接実行
        try:
            logger.info("JavaScriptで直接ダイアログのボタンを探します")
            script = """
            var dialogs = document.querySelectorAll('.ui-dialog');
            if (dialogs.length > 0) {
                var dialog = dialogs[dialogs.length - 1];
                var buttonPane = dialog.querySelector('.ui-dialog-buttonpane');
                if (buttonPane) {
                    var buttons = buttonPane.querySelectorAll('button');
                    // 「実行」ボタンを探す
                    for (var i = 0; i < buttons.length; i++) {
                        if (buttons[i].textContent.trim() === '実行') {
                            if (!buttons[i].classList.contains('ui-button-disabled')) {
                                buttons[i].click();
                                return 'execute';
                            }
                        }
                    }
                    // 「次へ」ボタンを探す
                    for (var i = 0; i < buttons.length; i++) {
                        if (buttons[i].textContent.trim() === '次へ') {
                            buttons[i].click();
                            return 'next';
                        }
                    }
                }
            }
            return false;
            """
            result = self.browser.driver.execute_script(script)
            if result:
                logger.info(f"✓ JavaScriptでボタンのクリックに成功しました: {result}")
                
                # 処理完了を待機
                time.sleep(10)
                
                if result == 'next':
                    # 画面4への遷移を確認
                    if "求職者 - インポート (4/4)" in self.browser.driver.page_source or "インポート (4" in self.browser.driver.page_source:
                        logger.info("✓ 画面4への遷移を確認しました")
                        self.browser.save_screenshot("screen4_displayed.png")
                        
                        # 画面4で「実行」ボタンを探す
                        return self._click_execute_button_on_screen4()
                    else:
                        logger.warning("画面4への遷移を確認できませんでした")
                else:
                    # 結果を確認
                    return self._check_import_result()
            else:
                logger.warning("JavaScriptでのボタンクリックも失敗しました")
        except Exception as js_error:
            logger.error(f"JavaScriptでのクリックにも失敗: {js_error}")
        
        # スクリーンショットを取得
        self.browser.save_screenshot("after_js_import_button.png")
        
        # すべての方法が失敗した場合
        logger.error("すべての方法でインポート実行ボタンのクリックに失敗しました")
        return False
    
    def _click_execute_button_on_screen4(self):
        """画面4の「実行」ボタンをクリックする"""
        try:
            logger.info("=== 画面4の「実行」ボタンのクリック処理を開始 ===")
            
            # スクリーンショットを取得
            self.browser.save_screenshot("before_execute_button.png")
            
            # 方法1: 直接ボタンを探す
            try:
                # ダイアログの状態を確認
                dialogs = self.browser.driver.find_elements(By.CSS_SELECTOR, ".ui-dialog")
                
                if dialogs:
                    # 最後のダイアログ（最前面）を使用
                    dialog = dialogs[-1]
                    
                    # ボタンパネルを探す
                    button_pane = dialog.find_element(By.CSS_SELECTOR, ".ui-dialog-buttonpane")
                    buttons = button_pane.find_elements(By.TAG_NAME, "button")
                    
                    # ボタンの情報をログに出力
                    for i, btn in enumerate(buttons):
                        logger.info(f"ボタン {i+1}: テキスト='{btn.text}', クラス={btn.get_attribute('class')}")
                    
                    # 「実行」ボタンを探す
                    execute_button = None
                    for btn in buttons:
                        if btn.text == "実行":
                            execute_button = btn
                            break
                    
                    if execute_button:
                        # スクロールして表示
                        self.browser.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", execute_button)
                        time.sleep(1)
                        
                        # クリック
                        execute_button.click()
                        logger.info("✓ 画面4の「実行」ボタンをクリックしました")
                        
                        # 処理完了を待機
                        time.sleep(5)
                        
                        # 結果を確認
                        return self._check_import_result()
                    else:
                        logger.warning("「実行」ボタンが見つかりませんでした")
            except Exception as e:
                logger.warning(f"ボタンパネルからのボタン検索に失敗: {e}")
            
            # 方法2: JavaScriptで直接実行
            try:
                logger.info("JavaScriptで直接「実行」ボタンをクリックします")
                script = """
                var dialogs = document.querySelectorAll('.ui-dialog');
                if (dialogs.length > 0) {
                    var dialog = dialogs[dialogs.length - 1];
                    var buttonPane = dialog.querySelector('.ui-dialog-buttonpane');
                    if (buttonPane) {
                        var buttons = buttonPane.querySelectorAll('button');
                        for (var i = 0; i < buttons.length; i++) {
                            if (buttons[i].textContent.trim() === '実行') {
                                buttons[i].click();
                                return true;
                            }
                        }
                    }
                }
                return false;
                """
                result = self.browser.driver.execute_script(script)
                if result:
                    logger.info("✓ JavaScriptで「実行」ボタンのクリックに成功しました")
                    
                    # 処理完了を待機
                    time.sleep(5)
                    
                    # 結果を確認
                    return self._check_import_result()
                else:
                    logger.warning("JavaScriptでの「実行」ボタンクリックも失敗しました")
            except Exception as js_error:
                logger.error(f"JavaScriptでのクリックにも失敗: {js_error}")
            
            # スクリーンショットを取得
            self.browser.save_screenshot("after_execute_button.png")
            
            # すべての方法が失敗した場合
            logger.error("すべての方法で画面4の「実行」ボタンのクリックに失敗しました")
            return False
        
        except Exception as e:
            logger.error(f"画面4の「実行」ボタンのクリック中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("execute_button_error.png")
            return False

    def _check_import_result(self):
        """インポート結果を確認する"""
        try:
            # インポート後のスクリーンショット
            self.browser.save_screenshot("after_import.png")
            
            # ページソースを取得
            page_source = self.browser.driver.page_source
            
            # 成功メッセージを探す
            success_patterns = ["成功", "完了", "インポートが完了", "正常に取り込まれました", "success", "completed"]
            success_found = False
            for pattern in success_patterns:
                if pattern in page_source:
                    logger.info(f"✅ インポート成功メッセージを確認: '{pattern}'")
                    success_found = True
                    break
            
            if success_found:
                # 成功メッセージが見つかった場合、OKボタンをクリック
                try:
                    logger.info("インポート完了後の「OK」ボタンを探します")
                    
                    # OKボタンを探す (複数の方法で試行)
                    ok_button = None
                    ok_button_clicked = False
                    
                    # 方法1: 直接ボタンを探す
                    try:
                        # 方法1-1: 特定のセレクタで探す
                        try:
                            ok_button_selector = "#pageCalendar > div.ui-dialog.ui-widget.ui-widget-content.ui-corner-all.ui-front.p-ui-messagebox.ui-dialog-buttons.ui-draggable > div.ui-dialog-buttonpane.ui-widget-content.ui-helper-clearfix > div > button"
                            ok_button = self.browser.driver.find_element(By.CSS_SELECTOR, ok_button_selector)
                            logger.info("✓ 特定セレクタでOKボタンを発見しました")
                        except:
                            logger.info("特定セレクタでOKボタンが見つかりませんでした")
                        
                        # 方法1-2: テキストで探す
                        if not ok_button:
                            buttons = self.browser.driver.find_elements(By.TAG_NAME, "button")
                            for btn in buttons:
                                if btn.text.strip() in ["OK", "Ok", "ok"]:
                                    ok_button = btn
                                    logger.info("✓ テキスト検索でOKボタンを発見しました")
                                    break
                        
                        # 方法1-3: メッセージボックス内のボタンを探す
                        if not ok_button:
                            try:
                                message_box = self.browser.driver.find_element(By.CSS_SELECTOR, ".p-ui-messagebox")
                                if message_box:
                                    message_buttons = message_box.find_elements(By.TAG_NAME, "button")
                                    if message_buttons:
                                        ok_button = message_buttons[0]  # 最初のボタンを使用
                                        logger.info("✓ メッセージボックス内でOKボタンを発見しました")
                            except:
                                logger.info("メッセージボックス内でOKボタンが見つかりませんでした")
                        
                        # 方法1-4: クラスで探す
                        if not ok_button:
                            ok_buttons = self.browser.driver.find_elements(By.CSS_SELECTOR, ".ui-button.ui-widget.ui-state-default.ui-corner-all.ui-button-text-only")
                            if ok_buttons:
                                for btn in ok_buttons:
                                    if btn.is_displayed() and btn.is_enabled():
                                        ok_button = btn
                                        logger.info("✓ クラス検索でOKボタンを発見しました")
                                        break
                        
                        if ok_button:
                            # スクロールして表示
                            self.browser.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", ok_button)
                            time.sleep(1)
                            
                            # クリック
                            ok_button.click()
                            logger.info("✓ インポート完了後の「OK」ボタンをクリックしました")
                            time.sleep(3)  # クリック後の処理を待機
                            ok_button_clicked = True
                        else:
                            logger.warning("インポート完了後の「OK」ボタンが見つかりませんでした")
                    except Exception as e:
                        logger.warning(f"「OK」ボタンの検索に失敗: {e}")
                    
                    # 方法2: JavaScriptで直接実行
                    if not ok_button_clicked:
                        try:
                            logger.info("JavaScriptで直接「OK」ボタンをクリックします")
                            script = """
                            // 特定のセレクタでOKボタンを探す
                            var specificButton = document.querySelector("#pageCalendar > div.ui-dialog.ui-widget.ui-widget-content.ui-corner-all.ui-front.p-ui-messagebox.ui-dialog-buttons.ui-draggable > div.ui-dialog-buttonpane.ui-widget-content.ui-helper-clearfix > div > button");
                            if (specificButton) {
                                specificButton.click();
                                return true;
                            }
                            
                            // メッセージボックス内のボタンを探す
                            var messageBox = document.querySelector(".p-ui-messagebox");
                            if (messageBox) {
                                var messageButtons = messageBox.querySelectorAll("button");
                                if (messageButtons.length > 0) {
                                    messageButtons[0].click();
                                    return true;
                                }
                            }
                            
                            // 表示されているボタンを探す
                            var buttons = document.querySelectorAll('button');
                            for (var i = 0; i < buttons.length; i++) {
                                var btn = buttons[i];
                                if (btn.textContent.trim().toLowerCase() === 'ok' && 
                                    btn.offsetParent !== null && 
                                    !btn.disabled) {
                                    btn.click();
                                    return true;
                                }
                            }
                            
                            // ダイアログ内のボタンを探す
                            var dialogs = document.querySelectorAll('.ui-dialog');
                            if (dialogs.length > 0) {
                                for (var j = 0; j < dialogs.length; j++) {
                                    var dialog = dialogs[j];
                                    var dialogButtons = dialog.querySelectorAll('button');
                                    for (var k = 0; k < dialogButtons.length; k++) {
                                        var dialogBtn = dialogButtons[k];
                                        if (dialogBtn.offsetParent !== null && !dialogBtn.disabled) {
                                            dialogBtn.click();
                                            return true;
                                        }
                                    }
                                }
                            }
                            return false;
                            """
                            result = self.browser.driver.execute_script(script)
                            if result:
                                logger.info("✓ JavaScriptでインポート完了後のボタンクリックに成功しました")
                                time.sleep(3)  # クリック後の処理を待機
                                ok_button_clicked = True
                            else:
                                logger.warning("JavaScriptでのボタンクリックも失敗しました")
                        except Exception as js_error:
                            logger.error(f"JavaScriptでのクリックにも失敗: {js_error}")
                    
                    # 最終確認のスクリーンショット
                    self.browser.save_screenshot("after_ok_button.png")
                    
                    # OKボタンがクリックできなかった場合はエラーとする
                    if not ok_button_clicked:
                        logger.error("❌ インポート完了後の「OK」ボタンをクリックできませんでした")
                        return False
                    
                except Exception as ok_error:
                    logger.error(f"「OK」ボタンクリック処理中にエラー: {ok_error}")
                    return False
                
                return True
            
            # エラーメッセージを探す
            error_patterns = ["エラー", "失敗", "error", "failed"]
            for pattern in error_patterns:
                if pattern in page_source:
                    logger.error(f"❌ インポートエラーメッセージを検出: '{pattern}'")
                    return False
            
            # 明確なメッセージが見つからない場合
            logger.warning("⚠️ インポート結果を明確に判断できませんでした。処理は続行します")
            return True
            
        except Exception as e:
            logger.error(f"インポート結果の確認中にエラーが発生しました: {str(e)}")
            return False

    def _click_other_operations_button(self, current_handles):
        """「その他業務」ボタンをクリックして新しいウィンドウに切り替える"""
        try:
            # 「その他業務」ボタンのセレクタを取得
            others_button_selector = "#main > div > main > section.original-search > header > div.others > button"
            if hasattr(self.browser, 'selectors') and 'porters_menu' in self.browser.selectors and 'search_button' in self.browser.selectors['porters_menu']:
                others_button_selector = self.browser.selectors['porters_menu']['search_button']['selector_value']
            
            logger.info(f"「その他業務」ボタンを探索: {others_button_selector}")
            
            # 要素を見つける
            others_button = WebDriverWait(self.browser.driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, others_button_selector))
            )
            
            # スクロールして確実に表示（standalone_testと同様）
            self.browser.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", others_button)
            time.sleep(1)
            
            # ボタンの詳細をログに記録
            logger.info(f"「その他業務」ボタン情報: テキスト={others_button.text}, クラス={others_button.get_attribute('class')}")
            
            # クリック実行 (JavaScriptでクリック)
            self.browser.driver.execute_script("arguments[0].click();", others_button)
            logger.info("✓ 「その他業務」ボタンをクリックしました")
            
            # 新しいウィンドウが開くのを待機 - 延長（standalone_testと同様）
            time.sleep(8)  # 十分な待機時間
            
            # 新しいウィンドウに切り替え
            new_handles = self.browser.driver.window_handles
            logger.info(f"操作後のウィンドウハンドル一覧: {new_handles}")
            
            if len(new_handles) > len(current_handles):
                # 新しいウィンドウが開かれた場合
                new_window = [handle for handle in new_handles if handle not in current_handles][0]
                logger.info(f"新しいウィンドウに切り替えます: {new_window}")
                self.browser.driver.switch_to.window(new_window)
                logger.info("✓ 新しいウィンドウにフォーカスを切り替えました")
                
                # 新しいウィンドウで読み込みを待機（standalone_testと同様）
                time.sleep(5)
                self.browser.save_screenshot("new_window.png")
                
                # 新しいウィンドウでのページ状態を確認
                new_window_html = self.browser.driver.page_source
                html_path = os.path.join(self.screenshot_dir, "new_window.html")
                with open(html_path, "w", encoding="utf-8") as f:
                    f.write(new_window_html)
                logger.info("新しいウィンドウのHTMLを保存しました")
                return True
            else:
                logger.warning("新しいウィンドウが開かれませんでした")
                return False
                
        except Exception as e:
            logger.error(f"「その他業務」ボタンのクリック中にエラーが発生しました: {str(e)}")
            self.browser.save_screenshot("other_operations_error.png")
            return False

    def _click_menu_item_5(self):
        """メニュー項目5をクリック"""
        try:
            # メニュー項目5のセレクタを取得
            menu_item_selector = "#main-menu-id-5 > a"
            if hasattr(self.browser, 'selectors') and 'porters_menu' in self.browser.selectors and 'menu_item_5' in self.browser.selectors['porters_menu']:
                menu_item_selector = self.browser.selectors['porters_menu']['menu_item_5']['selector_value']
            
            logger.info(f"メニュー項目5を探索: {menu_item_selector}")
            
            # メニュー項目をクリック
            menu_item = WebDriverWait(self.browser.driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, menu_item_selector))
            )
            menu_item.click()
            logger.info("✓ メニュー項目5をクリックしました")
            return True
            
        except Exception as e:
            logger.error(f"メニュー項目5のクリック中にエラーが発生しました: {str(e)}")
            self.browser.save_screenshot("menu_item_error.png")
            return False

    def _click_import_link(self):
        """「求職者のインポート」リンクをクリック"""
        try:
            logger.info("=== 「求職者のインポート」リンクのクリック処理を開始 ===")
            
            # メニューコンテナをスクロール（standalone_testと同様）
            try:
                menu_container = self.browser.driver.find_element(By.CSS_SELECTOR, ".main-menu-scrollable")
                logger.info(f"メニューコンテナを発見: ID={menu_container.get_attribute('id')}")
                
                # メニューを最下部までスクロール
                self.browser.driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", menu_container)
                logger.info("✓ メニューコンテナを最下部までスクロールしました")
                time.sleep(2)
                self.browser.save_screenshot("menu_scrolled_bottom.png")
            except Exception as e:
                logger.warning(f"メニューコンテナのスクロールに失敗しました: {e}")
            
            # 「求職者のインポート」リンクを探す（standalone_testと同様の複数の方法）
            try:
                # 1. title属性による検索（最も確実）
                logger.info("title属性を使って「求職者のインポート」リンクを検索")
                import_link = self.browser.driver.find_element(By.CSS_SELECTOR, "a[title='求職者のインポート']")
                logger.info(f"「求職者のインポート」リンクを見つけました: ID={import_link.get_attribute('id')}")
                
                # JavaScriptでクリック
                self.browser.driver.execute_script("arguments[0].click();", import_link)
                logger.info("✓ 「求職者のインポート」リンクをクリックしました")
                time.sleep(5)
                
                # ポップアップが表示されたか確認
                popup_elements = self.browser.driver.find_elements(By.CSS_SELECTOR, ".ui-dialog, .popup, .modal")
                if popup_elements:
                    logger.info("✓ ポップアップが表示されました")
                    self.browser.save_screenshot("popup_displayed.png")
                    return True
                else:
                    logger.warning("! ポップアップが表示されていません。再試行します")
                    
            except Exception as e:
                logger.warning(f"title属性での「求職者のインポート」リンク検索に失敗しました: {e}")
                
                try:
                    # 2. "インポート"ヘッダーの次の項目を探す方法
                    logger.info("「インポート」ヘッダーの下の項目を検索")
                    
                    # ヘッダーを見つける
                    import_headers = self.browser.driver.find_elements(By.XPATH, "//li[contains(@class, 'header')]/a[@title='インポート']")
                    if import_headers:
                        header = import_headers[0]
                        logger.info(f"「インポート」ヘッダーを見つけました: {header.text}")
                        
                        # ヘッダーの親要素（li）を取得
                        header_li = header.find_element(By.XPATH, "..")
                        
                        # 次の兄弟要素を取得
                        next_li = self.browser.driver.execute_script("""
                            var current = arguments[0];
                            var next = current.nextElementSibling;
                            return next;
                        """, header_li)
                        
                        if next_li:
                            # 次の要素内のリンクをクリック
                            import_link = next_li.find_element(By.TAG_NAME, "a")
                            logger.info(f"「インポート」ヘッダーの次の項目を見つけました: {import_link.text}")
                            
                            self.browser.driver.execute_script("arguments[0].click();", import_link)
                            logger.info("✓ 「インポート」ヘッダーの次の項目をクリックしました")
                            time.sleep(5)
                            return True
                except Exception as e2:
                    logger.warning(f"「インポート」ヘッダー方式での検索にも失敗しました: {e2}")
            
            # 最終手段：JavaScriptでのダイレクトアクセス（standalone_testと同様）
            try:
                # デバッグ情報の収集
                self.browser.save_screenshot("menu_error.png")
                html_path = os.path.join(self.screenshot_dir, "menu_html.html")
                with open(html_path, "w", encoding="utf-8") as f:
                    f.write(self.browser.driver.page_source)
                logger.info("デバッグ情報を保存しました")
                
                # 「求職者のインポート」リンクのIDを探す
                links = self.browser.driver.find_elements(By.TAG_NAME, "a")
                for link in links:
                    try:
                        if link.get_attribute("title") == "求職者のインポート":
                            link_id = link.get_attribute("id")
                            logger.info(f"「求職者のインポート」リンクを検出: ID={link_id}")
                            # IDを使ってJavaScriptでクリック
                            self.browser.driver.execute_script(f'document.getElementById("{link_id}").click();')
                            logger.info("✓ JavaScriptでIDを使ってクリックしました")
                            time.sleep(5)
                            return True
                    except:
                        continue
            except Exception as e:
                logger.error(f"最終手段も失敗しました: {e}")
            
            logger.error("「求職者のインポート」リンクが見つかりませんでした")
            self.browser.save_screenshot("import_link_not_found.png")
            return False
            
        except Exception as e:
            logger.error(f"「求職者のインポート」リンクのクリック中にエラーが発生しました: {str(e)}")
            self.browser.save_screenshot("import_link_error.png")
            return False

    def select_file(self, file_path):
        """ファイルを選択する"""
        try:
            logger.info("=== ファイル選択処理を開始 ===")
            logger.info(f"選択するファイル: {file_path}")
            
            # ページが完全に読み込まれるまで十分待機（standalone_testと同様）
            time.sleep(8)
            self.browser.save_screenshot("before_file_select.png")
            
            # 「添付」ボタンをクリック（standalone_testと同様）
            try:
                attachment_button = self.browser.driver.find_element(By.CSS_SELECTOR, "#_ibb_lbl")
                logger.info("「添付」ボタンを見つけました")
                self.browser.driver.execute_script("arguments[0].click();", attachment_button)
                logger.info("✓ 「添付」ボタンをクリックしました")
                time.sleep(3)
            except Exception as e:
                logger.warning(f"「添付」ボタンのクリックに失敗: {e}")
                self.browser.save_screenshot("attachment_button_error.png")
                
                # HTMLを保存
                html_path = os.path.join(self.screenshot_dir, "attachment_html.html")
                with open(html_path, "w", encoding="utf-8") as f:
                    f.write(self.browser.driver.page_source)
                    
                # 現在のURLも記録
                logger.info(f"現在のURL: {self.browser.driver.current_url}")
            
            # ファイル選択（standalone_testと同様）
            try:
                # 通常、添付ボタンの近くに隠れたinput要素がある
                file_input = self.browser.driver.find_element(By.CSS_SELECTOR, "input[type='file']")
                
                # JavaScript経由で表示状態を変更し、ファイルパスを送信
                self.browser.driver.execute_script("arguments[0].style.display = 'block';", file_input)
                file_input.send_keys(file_path)
                logger.info(f"✓ ファイルを選択しました: {file_path}")
                time.sleep(3)
                self.browser.save_screenshot("after_file_select.png")
                return True
            except Exception as e:
                logger.error(f"ファイル選択に失敗: {e}")
                self.browser.save_screenshot("file_select_error.png")
                return False
            
        except Exception as e:
            logger.error(f"ファイル選択処理中にエラーが発生しました: {str(e)}")
            self.browser.save_screenshot("file_select_error.png")
            return False

    # 「次へ」ボタンをクリックして次の画面に進む関数
    def click_next_button_and_wait(self, current_screen, next_screen):
        """「次へ」ボタンをクリックして次の画面に遷移するのを待つ"""
        logger.info(f"=== 画面{current_screen}から画面{next_screen}への遷移処理を開始 ===")
        
        # スクリーンショットを撮る
        self.browser.take_screenshot(f"before_next_button_screen{current_screen}")
        
        try:
            # ボタンパネルを探す
            button_pane = self.browser.find_element(By.CSS_SELECTOR, ".ui-dialog-buttonpane")
            # 「次へ」ボタンを探す
            next_button = button_pane.find_element(By.XPATH, ".//button[text()='次へ']")
            next_button.click()
            logger.info(f"✓ 画面{current_screen}の「次へ」ボタンをクリックしました")
        except Exception as e:
            logger.warning(f"ボタンパネルからのボタン検索に失敗: {e}")
            logger.info("JavaScriptで直接2番目のボタンをクリックします")
            try:
                self.browser.execute_script("""
                    var buttons = document.querySelectorAll('.ui-dialog-buttonpane button');
                    if (buttons.length >= 2) {
                        buttons[1].click();
                    }
                """)
                logger.info("✓ JavaScriptで2番目のボタンのクリックに成功しました")
            except Exception as js_error:
                logger.error(f"JavaScriptでのボタンクリックに失敗: {js_error}")
                return False
        
        # 次の画面への遷移を待つ
        expected_title = f"求職者 - インポート ({next_screen}/4)"
        try:
            WebDriverWait(self.browser.driver, 10).until(
                lambda driver: expected_title in driver.page_source
            )
            logger.info(f"✓ 画面{next_screen}への遷移を確認しました")
            self.browser.take_screenshot(f"screen{next_screen}_displayed")
            return True
        except Exception as e:
            logger.warning(f"画面{next_screen}への遷移を確認できませんでした: {e}")
            return False

    # 実行ボタンをクリックする処理
    def click_execute_button(self):
        """実行ボタンをクリックする"""
        logger.info("=== インポート実行ボタンのクリック処理を開始 ===")
        
        # 現在のURLをログに記録
        logger.info(f"現在のURL: {self.browser.driver.current_url}")
        
        # スクリーンショットを撮る
        self.browser.take_screenshot("before_import_button")
        
        # 現在の画面のHTMLを保存
        html_path = os.path.join("logs", "screenshots", "import_dialog.html")
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(self.browser.driver.page_source)
        logger.info(f"現在の画面HTMLを保存しました: {os.path.abspath(html_path)}")
        
        # 画面上のダイアログを探す
        dialogs = self.browser.find_elements(By.CSS_SELECTOR, ".ui-dialog")
        logger.info(f"画面上のダイアログ数: {len(dialogs)}")
        
        # 最後のダイアログを操作対象とする
        dialog = dialogs[-1]
        logger.info(f"操作対象ダイアログ: ID={dialog.get_attribute('id')}, クラス={dialog.get_attribute('class')}")
        
        # ダイアログのタイトルを取得
        title_element = dialog.find_element(By.CSS_SELECTOR, ".ui-dialog-title")
        logger.info(f"ダイアログタイトル: {title_element.text}")
        
        # 現在の画面を確認
        current_screen = 0
        for i in range(1, 5):
            if f"求職者 - インポート ({i}/4)" in title_element.text:
                current_screen = i
                break
        
        logger.info(f"現在の画面: {current_screen}/4")
        
        # 画面4（最終確認画面）でない場合は、画面4まで進める
        if current_screen < 4:
            logger.info(f"画面{current_screen}を検出しました。画面4まで進みます")
            
            # 画面2から画面3へ
            if current_screen == 2:
                if not self.click_next_button_for_screen2():
                    logger.error("画面3への遷移に失敗しました")
                    return False
                current_screen = 3
            
            # 画面3から画面4へ
            if current_screen == 3:
                if not self.click_next_button_and_wait(3, 4):
                    logger.error("画面4への遷移に失敗しました")
                    return False
            
            # 画面4に遷移したので、再度ダイアログを取得
            dialogs = self.browser.find_elements(By.CSS_SELECTOR, ".ui-dialog")
            dialog = dialogs[-1]
        
        # ボタンパネルを探す
        try:
            button_pane = dialog.find_element(By.CSS_SELECTOR, ".ui-dialog-buttonpane")
            logger.info("ダイアログのボタンパネルを見つけました")
            
            # ボタンパネル内のボタンを全て取得
            buttons = button_pane.find_elements(By.TAG_NAME, "button")
            logger.info(f"ボタンパネル内のボタン数: {len(buttons)}")
            
            # 各ボタンの情報をログに出力
            for i, button in enumerate(buttons, 1):
                logger.info(f"ボタン {i}: テキスト='{button.text}', クラス={button.get_attribute('class')}")
            
            # 「実行」ボタンを探す
            execute_button = None
            for button in buttons:
                if button.text == "実行" and "ui-state-disabled" not in button.get_attribute("class"):
                    execute_button = button
                    break
            
            if execute_button:
                execute_button.click()
                logger.info("✓ 「実行」ボタンをクリックしました")
            else:
                logger.warning("有効な「実行」ボタンが見つかりませんでした")
                return False
            
        except Exception as e:
            logger.warning(f"ボタンパネルからのボタン検索に失敗: {e}")
            logger.info("JavaScriptで直接ダイアログのボタンを探します")
            
            try:
                # JavaScriptで「実行」ボタンをクリック
                self.browser.execute_script("""
                    var dialogs = document.querySelectorAll('.ui-dialog');
                    var dialog = dialogs[dialogs.length - 1];
                    var buttons = dialog.querySelectorAll('.ui-dialog-buttonpane button');
                    for (var i = 0; i < buttons.length; i++) {
                        if (buttons[i].textContent.trim() === '実行' && 
                            !buttons[i].classList.contains('ui-state-disabled')) {
                            buttons[i].click();
                            return true;
                        }
                    }
                    return false;
                """)
                logger.info("✓ JavaScriptでボタンのクリックに成功しました")
            except Exception as js_error:
                logger.error(f"JavaScriptでのボタンクリックに失敗: {js_error}")
                return False
        
        # インポート完了を待つ
        time.sleep(5)
        self.browser.take_screenshot("after_import_button")
        
        return True

    def click_next_button_for_screen2(self):
        """画面2の「次へ」ボタンをクリック - スクロール対応版"""
        logger.info("=== 画面2の「次へ」ボタンのクリック処理を開始 ===")
        
        # スクリーンショットを撮る
        self.browser.take_screenshot("before_next_button_screen2")
        
        # 画面2が完全に読み込まれるのを待つ
        try:
            WebDriverWait(self.browser.driver, 10).until(
                lambda driver: "求職者 - インポート (2/4)" in driver.page_source
            )
            logger.info("✓ 画面2の読み込みを確認しました")
        except Exception as e:
            logger.warning(f"画面2の読み込み確認に失敗: {e}")
        
        # 少し待機して画面が安定するのを待つ
        time.sleep(2)
        
        # 方法1: ダイアログ全体を表示範囲内にスクロール
        try:
            logger.info("ダイアログ全体を表示範囲内にスクロールします")
            self.browser.driver.execute_script("""
                var dialog = document.querySelector('.ui-dialog');
                if (dialog) {
                    dialog.scrollIntoView({block: 'center', behavior: 'smooth'});
                }
            """)
            time.sleep(1)
        except Exception as e:
            logger.warning(f"ダイアログへのスクロールに失敗: {e}")
        
        # 方法2: ボタンパネルを表示範囲内にスクロール
        try:
            logger.info("ボタンパネルを表示範囲内にスクロールします")
            self.browser.driver.execute_script("""
                var buttonPane = document.querySelector('.ui-dialog-buttonpane');
                if (buttonPane) {
                    buttonPane.scrollIntoView({block: 'end', behavior: 'smooth'});
                }
            """)
            time.sleep(1)
        except Exception as e:
            logger.warning(f"ボタンパネルへのスクロールに失敗: {e}")
        
        # 方法3: ウィンドウサイズを調整
        try:
            logger.info("ウィンドウサイズを調整します")
            original_size = self.browser.driver.get_window_size()
            self.browser.driver.set_window_size(1200, 800)  # より大きなサイズに設定
            time.sleep(1)
        except Exception as e:
            logger.warning(f"ウィンドウサイズの調整に失敗: {e}")
        
        # 方法4: ダイアログのサイズを調整
        try:
            logger.info("ダイアログのサイズを調整します")
            self.browser.driver.execute_script("""
                var dialog = document.querySelector('.ui-dialog');
                if (dialog) {
                    dialog.style.height = 'auto';
                    dialog.style.maxHeight = '90vh';
                }
            """)
            time.sleep(1)
        except Exception as e:
            logger.warning(f"ダイアログサイズの調整に失敗: {e}")
        
        # スクリーンショットを撮って状態を確認
        self.browser.take_screenshot("after_scroll_adjustments")
        
        # 「次へ」ボタンをJavaScriptで直接クリック
        try:
            logger.info("JavaScriptで「次へ」ボタンを直接クリックします")
            result = self.browser.driver.execute_script("""
                // まず、ダイアログのタイトルで画面2を特定
                var dialogs = document.querySelectorAll('.ui-dialog');
                var targetDialog = null;
                
                for (var i = 0; i < dialogs.length; i++) {
                    var title = dialogs[i].querySelector('.ui-dialog-title');
                    if (title && title.textContent.includes('求職者 - インポート (2/4)')) {
                        targetDialog = dialogs[i];
                        break;
                    }
                }
                
                if (!targetDialog) {
                    // タイトルが見つからない場合は最後のダイアログを使用
                    targetDialog = dialogs[dialogs.length - 1];
                }
                
                // ボタンパネルを探す
                var buttonPane = targetDialog.querySelector('.ui-dialog-buttonpane');
                if (!buttonPane) return false;
                
                // ボタンパネル内のすべてのボタンを取得
                var buttons = buttonPane.querySelectorAll('button');
                
                // 「次へ」ボタンを探す
                var nextButton = null;
                for (var j = 0; j < buttons.length; j++) {
                    if (buttons[j].textContent.trim() === '次へ') {
                        nextButton = buttons[j];
                        break;
                    }
                }
                
                // 「次へ」ボタンが見つからない場合は2番目のボタンを使用
                if (!nextButton && buttons.length >= 2) {
                    nextButton = buttons[1];
                }
                
                if (nextButton) {
                    // ボタンが見つかった場合、スクロールして表示してからクリック
                    nextButton.scrollIntoView({block: 'center', behavior: 'smooth'});
                    setTimeout(function() {
                        nextButton.click();
                    }, 500);
                    return true;
                }
                
                return false;
            """)
            
            if result:
                logger.info("✓ JavaScriptでの「次へ」ボタンのクリックに成功しました")
                # クリック後に少し待機
                time.sleep(3)
                
                # 画面3への遷移を確認
                if "求職者 - インポート (3/4)" in self.browser.driver.page_source:
                    logger.info("✓ 画面3への遷移を確認しました")
                    self.browser.take_screenshot("screen3_displayed")
                    return True
                else:
                    logger.warning("画面3への遷移が確認できませんでした")
            else:
                logger.warning("JavaScriptでの「次へ」ボタンのクリックに失敗しました")
        except Exception as e:
            logger.error(f"JavaScriptでの「次へ」ボタンのクリック中にエラー: {e}")
        
        # 最後の手段: タブキーでフォーカスを移動させてEnterキーを押す
        try:
            logger.info("タブキーとEnterキーを使用して「次へ」ボタンをクリックします")
            
            # アクティブな要素にフォーカスを当てる
            active_element = self.browser.driver.switch_to.active_element
            
            # タブキーを複数回押して「次へ」ボタンにフォーカスを移動
            for _ in range(10):  # 最大10回タブキーを押す
                active_element.send_keys(Keys.TAB)
                time.sleep(0.5)
                
                # 現在フォーカスされている要素のテキストを確認
                focused_text = self.browser.driver.execute_script("return document.activeElement.textContent.trim();")
                logger.info(f"現在フォーカスされている要素のテキスト: {focused_text}")
                
                if focused_text == "次へ":
                    # 「次へ」ボタンにフォーカスが当たったらEnterキーを押す
                    self.browser.driver.switch_to.active_element.send_keys(Keys.ENTER)
                    logger.info("✓ 「次へ」ボタンにフォーカスを当ててEnterキーを押しました")
                    time.sleep(3)
                    
                    # 画面3への遷移を確認
                    if "求職者 - インポート (3/4)" in self.browser.driver.page_source:
                        logger.info("✓ 画面3への遷移を確認しました")
                        self.browser.take_screenshot("screen3_displayed")
                        return True
                    break
        except Exception as e:
            logger.error(f"タブキーとEnterキーを使用した方法でエラー: {e}")
        
        # ウィンドウサイズを元に戻す
        try:
            self.browser.driver.set_window_size(original_size['width'], original_size['height'])
        except:
            pass
        
        # すべての方法が失敗した場合
        logger.error("すべての方法で画面2の「次へ」ボタンのクリックに失敗しました")
        self.browser.take_screenshot("screen2_click_failed")
        return False

# import_to_porters() 関数は削除（importer.py に移動済み） 