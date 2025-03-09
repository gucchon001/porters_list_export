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
    
    def execute(self, csv_path=None):
        """CSVインポート処理を実行"""
        try:
            logger.info("=== CSVインポート処理を開始します ===")
            
            if not csv_path:
                # 環境変数からCSVパスを取得
                csv_path = env.get_env_var('CSV_PATH')
                if not csv_path:
                    logger.error("CSVファイルのパスが指定されていません")
                    return False
            
            # CSVファイルの存在確認
            if not os.path.exists(csv_path):
                logger.error(f"指定されたCSVファイルが存在しません: {csv_path}")
                return False
            
            logger.info(f"インポートするCSVファイル: {csv_path}")
            
            # インポートメニューを探す - 複数の方法を試す
            if not self.find_and_click_import_menu():
                logger.error("インポートメニューが見つかりませんでした")
                return False
            
            # インポート前のスクリーンショット
            self.browser.save_screenshot("before_import.png")
            
            # ファイルアップロード処理
            if not self.upload_csv_file(csv_path):
                logger.error("CSVファイルのアップロードに失敗しました")
                return False
            
            # インポート方法の選択（必要に応じて）
            if not self.select_import_method():
                logger.warning("インポート方法の選択に失敗しましたが、処理を続行します")
            
            # 「次へ」ボタンのクリック
            if not self.click_next_button():
                logger.error("「次へ」ボタンのクリックに失敗しました")
                return False
            
            # インポート実行ボタンのクリック
            if not self.click_import_button():
                logger.error("インポート実行ボタンのクリックに失敗しました")
                return False
            
            # インポート結果の確認
            if not self.check_import_result():
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
    
    def find_and_click_import_menu(self):
        """インポートメニューを探して開く"""
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
    
    def upload_csv_file(self, csv_path):
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
    
    def select_import_method(self):
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
    
    def click_next_button(self):
        """「次へ」ボタンをクリック"""
        try:
            logger.info("=== 「次へ」ボタンのクリック処理を開始 ===")
            
            # スクリーンショットで状態を確認
            self.browser.save_screenshot("before_next_button.png")
            
            try:
                # ダイアログのボタンパネルを特定
                dialog = self.browser.driver.find_element(By.CSS_SELECTOR, ".ui-dialog")
                button_pane = dialog.find_element(By.CSS_SELECTOR, ".ui-dialog-buttonpane")
                logger.info("ダイアログのボタンパネルを見つけました")
                
                # ボタンパネル内のすべてのボタン要素を取得
                buttons = button_pane.find_elements(By.TAG_NAME, "button")
                logger.info(f"ボタンパネル内のボタン数: {len(buttons)}")
                
                # 各ボタンの情報をログに記録
                for i, btn in enumerate(buttons):
                    try:
                        logger.info(f"ボタン {i+1}: テキスト='{btn.text}', クラス={btn.get_attribute('class')}")
                    except:
                        logger.info(f"ボタン {i+1}: [情報取得不可]")
                
                if len(buttons) >= 2:
                    # 2番目（インデックス1）のボタンが「次へ」
                    next_button = buttons[1]  # インデックスは0始まり
                    
                    # スクロールして表示（standalone_testと同様）
                    self.browser.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                    time.sleep(1)
                    
                    # JavaScriptでクリック
                    self.browser.driver.execute_script("arguments[0].click();", next_button)
                    logger.info("✓ 2番目のボタン（「次へ」）をJavaScriptでクリックしました")
                    time.sleep(3)
                    
                    # クリック後のスクリーンショット
                    self.browser.save_screenshot("after_next_button.png")
                    return True
                else:
                    # ボタンが少ない場合の対処
                    logger.warning(f"期待するボタン数が見つかりません。見つかったボタン数: {len(buttons)}")
                    
                    # 最後の手段：一番右側のボタン（通常は「次へ」や「確定」）をクリック
                    if buttons:
                        rightmost_button = buttons[-1]
                        self.browser.driver.execute_script("arguments[0].click();", rightmost_button)
                        logger.info("✓ 一番右側のボタンをクリックしました")
                        time.sleep(3)
                        return True
                    else:
                        raise Exception("クリック可能なボタンが見つかりません")
            
            except Exception as e:
                logger.warning(f"ボタンパネルからのボタン検索に失敗: {e}")
                
                # 最終手段：JavaScriptで直接実行（test_csv_importと同様）
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
                    time.sleep(3)
                    return True
                else:
                    logger.warning("JavaScriptでのクリックも失敗しました")
            
            # 「次へ」ボタンが見つからない場合は警告を出して続行
            logger.warning("「次へ」ボタンが見つかりませんでした。この画面では不要かもしれません。")
            return False
            
        except Exception as e:
            logger.warning(f"「次へ」ボタンのクリック中にエラーが発生しました: {str(e)}")
            self.browser.save_screenshot("next_button_error.png")
            return False
    
    def click_import_button(self):
        """インポート実行ボタンをクリックする"""
        try:
            logger.info("=== インポート実行ボタンのクリック処理を開始 ===")
            
            # 現在の画面状態を確認
            current_url = self.browser.driver.current_url
            logger.info(f"現在のURL: {current_url}")
            
            # スクリーンショットで状態を確認
            self.browser.save_screenshot("before_import_button.png")
            
            # HTMLを保存して詳細分析
            html_path = os.path.join(self.screenshot_dir, "import_dialog.html")
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
                    
                    if len(buttons) >= 1:
                        # 最後のボタンが「インポート」または「取込」
                        import_button = buttons[-1]  # 最後のボタン
                        
                        # スクロールして表示
                        self.browser.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", import_button)
                        time.sleep(1)
                        
                        # JavaScriptでクリック
                        self.browser.driver.execute_script("arguments[0].click();", import_button)
                        logger.info(f"✓ 最後のボタン（'{import_button.text}'）をJavaScriptでクリックしました")
                        time.sleep(5)  # インポート処理の完了を待つ
                        
                        # クリック後のスクリーンショット
                        self.browser.save_screenshot("after_import_button.png")
                        return True
                except Exception as e:
                    logger.warning(f"ボタンパネルからのボタン検索に失敗: {e}")
            
            # 方法2: JavaScriptで直接ダイアログのボタンを探してクリック
            logger.info("JavaScriptで直接ダイアログのボタンを探します")
            script = """
            // すべてのダイアログを取得
            var dialogs = document.querySelectorAll('.ui-dialog');
            if (dialogs.length > 0) {
                // 最後のダイアログ（最前面）を使用
                var dialog = dialogs[dialogs.length - 1];
                console.log('ダイアログID: ' + dialog.id);
                
                // ボタンパネルを探す
                var buttonPane = dialog.querySelector('.ui-dialog-buttonpane');
                if (buttonPane) {
                    // すべてのボタンを取得
                    var buttons = buttonPane.querySelectorAll('button');
                    console.log('ボタン数: ' + buttons.length);
                    
                    if (buttons.length > 0) {
                        // 最後のボタン（通常は「インポート」）をクリック
                        var lastButton = buttons[buttons.length - 1];
                        console.log('クリックするボタン: ' + lastButton.textContent);
                        lastButton.click();
                        return true;
                    }
                } else {
                    // ボタンパネルが見つからない場合、ダイアログ内のすべてのボタンを探す
                    var allButtons = dialog.querySelectorAll('button');
                    if (allButtons.length > 0) {
                        // 最後のボタンをクリック
                        allButtons[allButtons.length - 1].click();
                        return true;
                    }
                }
            }
            
            // ダイアログが見つからない場合、ページ全体からボタンを探す
            var importButtons = Array.from(document.querySelectorAll('button')).filter(function(btn) {
                var text = btn.textContent.toLowerCase();
                return text.includes('インポート') || text.includes('取込') || text.includes('import');
            });
            
            if (importButtons.length > 0) {
                importButtons[0].click();
                return true;
            }
            
            return false;
            """
            
            result = self.browser.driver.execute_script(script)
            if result:
                logger.info("✓ JavaScriptでボタンのクリックに成功しました")
                time.sleep(5)  # インポート処理の完了を待つ
                
                # クリック後のスクリーンショット
                self.browser.save_screenshot("after_js_import_button.png")
                return True
            
            # 方法3: 最終手段 - 画面上のすべてのボタンを調査
            logger.info("最終手段: 画面上のすべてのボタンを調査します")
            all_buttons = self.browser.driver.find_elements(By.TAG_NAME, 'button')
            logger.info(f"画面上のボタン総数: {len(all_buttons)}")
            
            # インポート関連のボタンを探す
            import_buttons = []
            for btn in all_buttons:
                try:
                    btn_text = btn.text.lower()
                    if 'インポート' in btn_text or '取込' in btn_text or 'import' in btn_text:
                        import_buttons.append(btn)
                        logger.info(f"インポート関連ボタンを検出: '{btn.text}'")
                except:
                    continue
            
            if import_buttons:
                # 最初のインポートボタンをクリック
                self.browser.driver.execute_script("arguments[0].click();", import_buttons[0])
                logger.info(f"✓ インポート関連ボタン '{import_buttons[0].text}' をクリックしました")
                time.sleep(5)
                return True
            elif all_buttons:
                # 最後のボタンをクリック（最後のボタンが通常「確定」や「インポート」）
                self.browser.driver.execute_script("arguments[0].click();", all_buttons[-1])
                logger.info("✓ 画面上の最後のボタンをクリックしました")
                time.sleep(5)
                return True
            
            logger.error("インポート実行ボタンが見つかりませんでした")
            self.browser.save_screenshot("import_button_not_found.png")
            return False
            
        except Exception as e:
            logger.error(f"インポート実行ボタンのクリック中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("import_button_error.png")
            return False
    
    def check_import_result(self):
        """インポート結果を確認する"""
        try:
            # インポート後のスクリーンショット
            self.browser.save_screenshot("after_import.png")
            
            # ページソースを取得
            page_source = self.browser.driver.page_source
            
            # 成功メッセージを探す
            success_patterns = ["成功", "完了", "インポートが完了", "正常に取り込まれました", "success", "completed"]
            for pattern in success_patterns:
                if pattern in page_source:
                    logger.info(f"✅ インポート成功メッセージを確認: '{pattern}'")
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

# import_to_porters() 関数は削除（importer.py に移動済み） 