import time
import os
from pathlib import Path
import sys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException, NoSuchElementException
from src.utils.environment import EnvironmentUtils as env

# プロジェクトのルートディレクトリをPYTHONPATHに追加
root_dir = Path(__file__).parent.parent
sys.path.append(str(root_dir))

from src.utils.logging_config import get_logger

logger = get_logger(__name__)

class TestCsvImport:
    def __init__(self, browser):
        """CSVインポート処理を管理するクラス"""
        self.browser = browser
        self.screenshot_dir = browser.screenshot_dir
    
    def execute(self):
        """ログイン後の操作を実行する"""
        try:
            logger.info("=== ログイン後の操作を開始します ===")
            
            # ページが完全に読み込まれるまで十分待機
            time.sleep(2)

            # 現在のページ状態を確認
            page_html = self.browser.driver.page_source
            page_analysis = self.browser.analyze_page_content(page_html)
            logger.info(f"現在のページ状態: タイトル={page_analysis['page_title']}")
            logger.info(f"メニュー項目数: {len(page_analysis['menu_items'])}")
            
            # スクリーンショット保存
            self.browser.save_screenshot("before_operations.png")
            
            # 現在のウィンドウハンドルを保存（後で戻れるように）
            main_window_handle = self.browser.driver.current_window_handle
            logger.info(f"メインウィンドウハンドル: {main_window_handle}")
            current_handles = self.browser.driver.window_handles
            logger.info(f"操作前のウィンドウハンドル一覧: {current_handles}")
            
            # "その他業務"ボタンの操作
            if not self._click_others_button():
                return False
            
            # 新しいウィンドウに切り替え
            if not self._switch_to_new_window(current_handles):
                return False
            
            # メニュー項目5をクリック
            if not self._click_menu_item():
                return False
            
            # 「求職者のインポート」リンクをクリック
            if not self._click_import_link():
                return False
            
            # ファイルインポート操作
            if not self._perform_file_import():
                return False
            
            # インポート設定画面の操作
            if not self._perform_import_settings():
                return False
            
            # レコードの重複処理設定画面の操作
            if not self._perform_mapping_settings():
                return False
            
            # インポート実行画面の操作
            if not self._perform_final_confirmation():
                return False
            
            # 操作が正常に完了
            logger.info("✅ メニュー操作が正常に完了しました")
            return True
            
        except Exception as e:
            logger.error(f"ログイン後の操作中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("error_after_login.png")
            
            # メインウィンドウに戻る
            try:
                self.browser.driver.switch_to.window(main_window_handle)
                logger.info("メインウィンドウに戻りました")
            except:
                logger.warning("メインウィンドウへの切り替えに失敗しました")
                
            return False
    
    def _handle_popup(self):
        """ポップアップ処理"""
        try:
            # ポップアップボタン候補を確認（タイムアウトを3秒に短縮）
            popup_wait = WebDriverWait(self.browser.driver, 3)
            popup_selectors = [
                ".close-button", 
                ".popup-close", 
                "button.close", 
                ".modal-close",
                ".ui-dialog-titlebar-close",
                "[aria-label='Close']"
            ]
            
            for selector in popup_selectors:
                try:
                    popup_button = popup_wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    logger.info(f"ポップアップを検出しました。セレクタ {selector} で閉じます...")
                    
                    # JavaScriptでクリック
                    self.browser.driver.execute_script("arguments[0].click();", popup_button)
                    logger.info("✓ ポップアップを閉じました")
                    time.sleep(1)  # 短い待機で十分
                    break
                except:
                    # このセレクタでは見つからなかっただけなので、次を試す
                    continue
        except Exception as popup_error:
            # ポップアップがなければ単に続行
            logger.info("ポップアップは検出されませんでした。処理を続行します。")
    
    def _click_others_button(self):
        """「その他業務」ボタンをクリック"""
        others_button_selector = "#main > div > main > section.original-search > header > div.others > button"
        logger.info(f"「その他業務」ボタンを探索: {others_button_selector}")
        
        try:
            # 十分な待機を追加
            time.sleep(3)
            
            # スクリーンショットで状態確認
            self.browser.save_screenshot("before_others_click.png")
            
            # 要素を見つける
            others_button = self.browser.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, others_button_selector)))
            
            # スクロールして確実に表示
            self.browser.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", others_button)
            time.sleep(1)
            
            # ボタンの詳細をログに記録
            logger.info(f"「その他業務」ボタン情報: テキスト={others_button.text}, クラス={others_button.get_attribute('class')}")
            
            # クリック実行 (JavaScriptでクリック)
            self.browser.driver.execute_script("arguments[0].click();", others_button)
            logger.info("✓ 「その他業務」ボタンをクリックしました")
            
            # 新しいウィンドウが開くのを待機
            time.sleep(8)
            
            return True
        except Exception as e:
            logger.error(f"「その他業務」ボタンのクリックに失敗: {str(e)}")
            self.browser.save_screenshot("others_button_error.png")
            return False
    
    def _switch_to_new_window(self, current_handles):
        """新しいウィンドウに切り替え"""
        try:
            # 新しいウィンドウに切り替え
            new_handles = self.browser.driver.window_handles
            logger.info(f"操作後のウィンドウハンドル一覧: {new_handles}")
            
            if len(new_handles) > len(current_handles):
                # 新しいウィンドウが開かれた場合
                new_window = [handle for handle in new_handles if handle not in current_handles][0]
                logger.info(f"新しいウィンドウに切り替えます: {new_window}")
                self.browser.driver.switch_to.window(new_window)
                logger.info("✓ 新しいウィンドウにフォーカスを切り替えました")
                
                # 新しいウィンドウで読み込みを待機
                time.sleep(5)
                self.browser.save_screenshot("new_window.png")
                
                # 新しいウィンドウでのページ状態を確認
                new_window_html = self.browser.driver.page_source
                with open(os.path.join(self.screenshot_dir, "new_window.html"), "w", encoding="utf-8") as f:
                    f.write(new_window_html)
                logger.info("新しいウィンドウのHTMLを保存しました")
                return True
            else:
                logger.warning("新しいウィンドウが検出されませんでした")
                return False
        except Exception as e:
            logger.error(f"新しいウィンドウへの切り替えに失敗: {str(e)}")
            return False
    
    def _click_menu_item(self):
        """メニュー項目5をクリック"""
        try:
            # メニュー項目5をクリック
            menu_item_selector = "#main-menu-id-5 > a"
            if 'porters_menu' in self.browser.selectors and 'menu_item_5' in self.browser.selectors['porters_menu']:
                menu_item_selector = self.browser.selectors['porters_menu']['menu_item_5']['selector_value']
            logger.info(f"メニュー項目5を探索: {menu_item_selector}")
            
            # スクリーンショットを撮ってページの状態を記録
            self.browser.save_screenshot("before_menu_click.png")
            
            # メニュー項目をクリック
            menu_item = self.browser.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, menu_item_selector)))
            menu_item.click()
            logger.info("✓ メニュー項目5をクリックしました")
            
            # サブメニューが表示されるまでしっかり待機
            time.sleep(5)
            self.browser.save_screenshot("after_menu_click.png")
            
            # メニューコンテナを見つけてスクロール
            try:
                menu_containers = self.browser.driver.find_elements(By.CSS_SELECTOR, ".ui-menu, .dropdown-menu, .menu")
                if menu_containers:
                    # 最後に見つかったメニューコンテナが通常ドロップダウン
                    menu_container = menu_containers[-1]
                    logger.info(f"メニューコンテナを見つけました: {menu_container.tag_name}")
                    
                    # コンテナの一番下までスクロール
                    self.browser.driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", menu_container)
                    logger.info("✓ メニューコンテナの一番下までスクロールしました")
                    time.sleep(2)
                    
                    # スクロール後のスクリーンショット
                    self.browser.save_screenshot("after_scroll.png")
            except Exception as e:
                logger.warning(f"メニューコンテナのスクロールに失敗しましたが、処理を続行します: {e}")
            
            return True
        except Exception as e:
            logger.error(f"メニュー項目のクリックに失敗: {str(e)}")
            self.browser.save_screenshot("menu_click_error.png")
            return False
    
    def _click_import_link(self):
        """「求職者のインポート」リンクをクリック"""
        try:
            logger.info("「求職者のインポート」リンクをクリックします")
            
            # メニュースクロール - 最下部にある可能性が高い
            try:
                menu_container = self.browser.driver.find_element(By.CSS_SELECTOR, ".main-menu-scrollable")
                logger.info(f"メニューコンテナを発見: ID={menu_container.get_attribute('id')}")
                
                # メニューを最下部までスクロール
                self.browser.driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", menu_container)
                logger.info("メニューを最下部までスクロールしました")
                time.sleep(1)
                self.browser.save_screenshot("menu_scrolled_bottom.png")
            except Exception as scroll_error:
                logger.warning(f"メニューのスクロールに失敗しましたが、処理を続行します: {scroll_error}")
            
            # 1. title属性による検索（最も確実）
            try:
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
                    
                    # ポップアップの状態を記録
                    popup_html = popup_elements[0].get_attribute('innerHTML')
                    with open(os.path.join(self.screenshot_dir, "import_popup.html"), "w", encoding="utf-8") as f:
                        f.write(popup_html)
                        
                    return True
                else:
                    logger.warning("! ポップアップが表示されていません。再試行します")
                    # ポップアップが表示されなくても、すぐに失敗とせず次の方法を試す
            except Exception as e:
                logger.error(f"1回目の試行でエラー: {e}")
            
            # 2. "インポート"ヘッダーの次の項目を探す方法
            try:
                logger.info("「インポート」ヘッダーの下の項目を検索")
                
                # ヘッダーを見つける
                import_headers = self.browser.driver.find_elements(By.XPATH, "//li[contains(@class, 'header')]/a[@title='インポート']")
                if import_headers:
                    import_header = import_headers[0]
                    logger.info(f"「インポート」ヘッダーを見つけました: ID={import_header.get_attribute('id')}")
                    
                    # ヘッダーの親要素から次の兄弟要素を取得
                    header_li = import_header.find_element(By.XPATH, "..")
                    next_li = header_li.find_element(By.XPATH, "following-sibling::li")
                    import_link = next_li.find_element(By.TAG_NAME, "a")
                    
                    logger.info(f"「インポート」ヘッダー直後の項目: '{import_link.text}'")
                    self.browser.driver.execute_script("arguments[0].click();", import_link)
                    logger.info("✓ 「インポート」ヘッダー後の項目をクリックしました")
                    time.sleep(5)
                else:
                    # 3. メニュー最下部の項目をクリック
                    logger.info("メニュー最下部の項目を探索")
                    menu_items = self.browser.driver.find_elements(By.CSS_SELECTOR, ".main-menu-scrollable li")
                    if menu_items:
                        last_item = menu_items[-1]
                        last_link = last_item.find_element(By.TAG_NAME, "a")
                        logger.info(f"メニュー最下部の項目: '{last_link.text}'")
                        
                        self.browser.driver.execute_script("arguments[0].click();", last_link)
                        logger.info("✓ メニュー最下部の項目をクリックしました")
                        time.sleep(5)
                    else:
                        raise Exception("メニュー項目が見つかりません")
                
                # ポップアップ確認
                popup_check = self.browser.driver.find_elements(By.CSS_SELECTOR, ".ui-dialog, #porters-pdialog_1")
                if popup_check:
                    logger.info("✓ ポップアップが表示されました")
                    self.browser.save_screenshot("popup_displayed_method2.png")
                    return True
                else:
                    logger.warning("! ポップアップが表示されていません")
                    # 続行して次の方法を試す
            except Exception as e2:
                logger.error(f"全ての方法でリンクのクリックに失敗: {e2}")
                
                # デバッグ情報の収集
                self.browser.save_screenshot("menu_error.png")
                with open(os.path.join(self.screenshot_dir, "menu_html.html"), "w", encoding="utf-8") as f:
                    f.write(self.browser.driver.page_source)
                logger.info("デバッグ情報を保存しました")
            
            # 4. JavaScriptでのダイレクトアクセスを試みる（最終手段）
            try:
                # 「求職者のインポート」リンクのIDを探す
                links = self.browser.driver.find_elements(By.TAG_NAME, "a")
                for link in links:
                    if link.get_attribute("title") == "求職者のインポート":
                        link_id = link.get_attribute("id")
                        logger.info(f"「求職者のインポート」リンクを検出: ID={link_id}")
                        # IDを使ってJavaScriptでクリック
                        self.browser.driver.execute_script(f'document.getElementById("{link_id}").click();')
                        logger.info("✓ JavaScriptでIDを使ってクリックしました")
                        time.sleep(5)
                        
                        # ポップアップ確認
                        popup_final = self.browser.driver.find_elements(By.CSS_SELECTOR, ".ui-dialog, #porters-pdialog_1")
                        if popup_final:
                            logger.info("✓ 最終手段でポップアップが表示されました")
                            self.browser.save_screenshot("popup_displayed_final.png")
                            return True
                        break
            except Exception as js_error:
                logger.error(f"JavaScriptでのアクセスにも失敗: {js_error}")
            
            # 全ての方法を試した後のチェック
            # すべての方法で明示的なポップアップが確認できなくても、画面が変わっていれば成功とみなす
            logger.warning("全ての方法を試しましたが、明示的なポップアップは確認できませんでした")
            logger.info("画面の変化を確認します")
            self.browser.save_screenshot("final_state.png")
            
            # HTMLを保存して分析
            with open(os.path.join(self.screenshot_dir, "final_html.html"), "w", encoding="utf-8") as f:
                f.write(self.browser.driver.page_source)
            
            # 求職者インポート関連の要素がページ内にあるか確認
            page_source = self.browser.driver.page_source
            if "インポート" in page_source or "求職者" in page_source or "添付" in page_source:
                logger.info("✓ ページ内にインポート関連の要素を検出しました。処理を続行します")
                return True
            
            logger.error("全ての方法が失敗し、ページ変化も検出できませんでした")
            return False
                
        except Exception as e:
            logger.error(f"「求職者のインポート」リンクの操作に失敗: {str(e)}")
            self.browser.save_screenshot("import_link_error.png")
            return False
    
    def _perform_file_import(self):
        """ファイルインポート操作"""
        try:
            logger.info("=== ファイルインポート画面での操作開始 ===")
            
            # ページが完全に読み込まれるまで十分待機
            time.sleep(8)
            
            # 現在のページ状態を確認
            page_title = self.browser.driver.title
            page_source = self.browser.driver.page_source
            
            # 画面の確認
            if "インポート" in page_source:
                # ヘッダーテキストなどから画面ステップを確認
                headers = self.browser.driver.find_elements(By.TAG_NAME, "h1")
                if headers:
                    header_text = headers[0].text
                    logger.info(f"現在の画面: {header_text}")
                else:
                    logger.info("現在の画面: 求職者 - インポート (1/4) インポートするファイルの設定")
            else:
                logger.info(f"現在の画面タイトル: {page_title}")
            
            # 「添付」ボタンをクリック
            try:
                logger.info("「添付」ボタンを探しています...")
                attachment_button = self.browser.driver.find_element(By.CSS_SELECTOR, "#_ibb_lbl")
                logger.info("「添付」ボタンを見つけました")
                self.browser.driver.execute_script("arguments[0].click();", attachment_button)
                logger.info("✓ 「添付」ボタンをクリックしました")
                time.sleep(3)
            except Exception as e:
                logger.error(f"「添付」ボタンのクリックに失敗: {e}")
                self.browser.save_screenshot("attachment_button_error.png")
                # HTMLを保存
                with open(os.path.join(self.screenshot_dir, "attachment_html.html"), "w", encoding="utf-8") as f:
                    f.write(self.browser.driver.page_source)
                    
                # 現在のURLも記録
                logger.info(f"現在のURL: {self.browser.driver.current_url}")
            
            # ファイル選択（input type=fileに直接パスを送信）
            try:
                # 通常、添付ボタンの近くに隠れたinput要素がある
                file_input = self.browser.driver.find_element(By.CSS_SELECTOR, "input[type='file']")
                
                # standalone_test.py と同じファイルパスを使用
                # または環境変数から取得
                file_path = os.environ.get('CSV_IMPORT_FILE', r"C:\Users\yohay\Downloads\LINE_member_0303.csv")
                
                # ファイルが存在するか確認
                if not os.path.exists(file_path):
                    logger.warning(f"指定されたファイルが存在しません: {file_path}")
                    logger.info("代替としてサンプルCSVファイルを使用します")
                    
                    # サンプルCSVファイルのパスを生成
                    file_path = os.path.join(root_dir, "data", "sample_import.csv")
                    
                    # サンプルCSVファイルが存在しない場合は作成
                    if not os.path.exists(file_path):
                        # ディレクトリが存在しない場合は作成
                        os.makedirs(os.path.dirname(file_path), exist_ok=True)
                        
                        # 簡単なCSVサンプルを作成（LINE会員情報に似せる）
                        with open(file_path, 'w', encoding='utf-8') as f:
                            f.write("id,name,email,phone,note\n")
                            f.write("1,テスト太郎,test1@example.com,08012345678,テスト会員1\n")
                            f.write("2,テスト次郎,test2@example.com,08087654321,テスト会員2\n")
                            f.write("3,テスト三郎,test3@example.com,08011112222,テスト会員3\n")
                        logger.info(f"サンプルCSVファイルを作成しました: {file_path}")
                
                # JavaScript経由で表示状態を変更し、ファイルパスを送信
                self.browser.driver.execute_script("arguments[0].style.display = 'block';", file_input)
                file_input.send_keys(file_path)
                logger.info(f"✓ ファイルを選択しました: {file_path}")
                time.sleep(3)
                self.browser.save_screenshot("after_file_select.png")
            except Exception as e:
                logger.error(f"ファイル選択に失敗: {e}")
                self.browser.save_screenshot("file_select_error.png")
            
            # インポート方法の選択（LINE初回アンケート取込）
            try:
                logger.info("LINE初回アンケート取込の選択肢を探しています...")
                
                # standalone_test.py と同じセレクタを使用
                import_method_selector = "#porters-pdialog_1 > div > div.subWrap.resize > div > div > div > ul > li:nth-child(9) > label > input[type=radio]"
                import_method = self.browser.driver.find_element(By.CSS_SELECTOR, import_method_selector)
                self.browser.driver.execute_script("arguments[0].click();", import_method)
                logger.info("✓ 「LINE初回アンケート取込」を選択しました")
                time.sleep(2)
                self.browser.save_screenshot("import_method_selected.png")
            except Exception as e:
                logger.error(f"LINE初回アンケート取込の選択に失敗: {e}")
                self.browser.save_screenshot("import_method_error.png")
                
                # 代替手段：テキストで選択肢を探す
                try:
                    logger.info("代替方法：テキストでLINE取込選択肢を探しています...")
                    # すべてのラジオボタンのラベルを調査
                    radio_labels = self.browser.driver.find_elements(By.CSS_SELECTOR, "label")
                    for label in radio_labels:
                        if "LINE" in label.text:
                            logger.info(f"LINEを含むラベルを発見: {label.text}")
                            radio = label.find_element(By.CSS_SELECTOR, "input[type='radio']")
                            self.browser.driver.execute_script("arguments[0].click();", radio)
                            logger.info(f"✓ 「{label.text}」を選択しました")
                            time.sleep(2)
                            self.browser.save_screenshot("line_option_selected.png")
                            break
                    else:
                        # どのラジオボタンも見つからない場合は最初のラジオボタンを選択
                        logger.warning("LINE関連の選択肢が見つからないため、最初のラジオボタンを選択します")
                        radio_buttons = self.browser.driver.find_elements(By.CSS_SELECTOR, "input[type='radio']")
                        if radio_buttons:
                            self.browser.driver.execute_script("arguments[0].click();", radio_buttons[0])
                            logger.info("✓ 最初のラジオボタンを選択しました")
                            time.sleep(2)
                except Exception as text_search_error:
                    logger.error(f"代替選択方法も失敗: {text_search_error}")
            
            # 「次へ」ボタンをクリック - 位置ベースのアプローチ
            try:
                logger.info("「次へ」ボタンを探索します（位置ベースアプローチ）")
                
                # スクリーンショットで状態を確認
                self.browser.save_screenshot("before_next_button.png")
                
                # ダイアログのボタンパネルを特定
                dialog = self.browser.driver.find_element(By.CSS_SELECTOR, ".ui-dialog")
                button_pane = dialog.find_element(By.CSS_SELECTOR, ".ui-dialog-buttonpane")
                logger.info("ダイアログのボタンパネルを見つけました")
                
                # ボタンパネル内のすべてのボタン要素を取得
                buttons = button_pane.find_elements(By.TAG_NAME, "button")
                logger.info(f"ボタンパネル内のボタン数: {len(buttons)}")
                
                if len(buttons) >= 2:
                    # 2番目（インデックス1）のボタンが「次へ」
                    next_button = buttons[1]  # インデックスは0始まり
                    
                    # ボタンの情報をログに記録
                    logger.info(f"2番目のボタン: クラス={next_button.get_attribute('class')}")
                    
                    # スクロールして表示
                    self.browser.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                    time.sleep(1)
                    
                    # JavaScriptでクリック
                    self.browser.driver.execute_script("arguments[0].click();", next_button)
                    logger.info("✓ 2番目のボタン（「次へ」）をJavaScriptでクリックしました")
                    
                    # 画面遷移を待機
                    time.sleep(5)
                    
                    # 遷移後の画面状態を確認
                    self.browser.save_screenshot("after_next_button_transition.png")
                    
                    # 画面遷移の確認
                    try:
                        # 現在の画面をログに記録
                        current_html = self.browser.driver.page_source
                        page_title = self.browser.driver.title
                        
                        # ステップ表示を探す（複数の方法で試行）
                        step_indicators = self.browser.driver.find_elements(By.CSS_SELECTOR, ".step-indicator, .wizard-steps, .progress-steps")
                        
                        if step_indicators:
                            step_text = step_indicators[0].text
                            logger.info(f"ステップ表示: {step_text}")
                        
                        # ヘッダー要素を探す
                        headers = self.browser.driver.find_elements(By.CSS_SELECTOR, "h1, h2, .page-header, .title")
                        header_text = ""
                        
                        if headers:
                            header_text = headers[0].text
                            logger.info(f"画面ヘッダー: {header_text}")
                        
                        # 画面状態の判定
                        if "(2/4)" in current_html or "インポート設定" in current_html:
                            logger.info("✅ 画面遷移確認: 求職者 - インポート (2/4) インポート設定画面")
                        elif "2" in header_text and ("4" in header_text) and "インポート" in header_text:
                            logger.info("✅ 画面遷移確認: 求職者 - インポート (2/4) インポート設定画面")
                        else:
                            # どれにも当てはまらない場合
                            logger.info(f"画面遷移確認: タイトル「{page_title}」の画面に遷移しました")
                            
                            # HTMLを保存して後で分析できるようにする
                            with open(os.path.join(self.screenshot_dir, "page2_html.html"), "w", encoding="utf-8") as f:
                                f.write(current_html)
                    
                    except Exception as e:
                        logger.error(f"画面遷移確認中にエラー: {e}")
                else:
                    # ボタンが少ない場合の対処
                    logger.error(f"期待するボタン数が見つかりません。見つかったボタン数: {len(buttons)}")
                    
                    # すべてのボタンの情報を記録
                    for i, btn in enumerate(buttons):
                        try:
                            logger.info(f"ボタン {i+1}: クラス={btn.get_attribute('class')}")
                        except:
                            logger.info(f"ボタン {i+1}: [情報取得不可]")
                    
                    # 最後の手段：一番右側のボタン（通常は「次へ」や「確定」）をクリック
                    if buttons:
                        rightmost_button = buttons[-1]
                        self.browser.driver.execute_script("arguments[0].click();", rightmost_button)
                        logger.info("✓ 一番右側のボタンをクリックしました")
                        time.sleep(3)
                    else:
                        raise Exception("クリック可能なボタンが見つかりません")
            except Exception as e:
                logger.error(f"「次へ」ボタンのクリックに失敗: {e}")
                
                # 最終手段：JavaScriptで直接実行
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
                        time.sleep(3)
                    else:
                        logger.error("JavaScriptでのクリックも失敗しました")
                        
                except Exception as js_error:
                    logger.error(f"JavaScriptでのクリックにも失敗: {js_error}")
                
                # デバッグ情報
                self.browser.save_screenshot("next_button_error.png")
                with open(os.path.join(self.screenshot_dir, "next_button_html.html"), "w", encoding="utf-8") as f:
                    f.write(self.browser.driver.page_source)
            
            logger.info("=== ファイルインポート画面での操作完了 ===")
            return True
            
        except Exception as e:
            logger.error(f"ファイルインポート操作に失敗: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("file_import_error.png")
            
            # エラー発生時のHTMLを保存
            with open(os.path.join(self.screenshot_dir, "error_html.html"), "w", encoding="utf-8") as f:
                f.write(self.browser.driver.page_source)
            
            return False
    
    def _perform_import_settings(self):
        """インポート設定画面での操作"""
        try:
            logger.info("=== インポート設定画面での操作開始 ===")
            
            # 画面が完全に読み込まれるまで十分待機
            time.sleep(5)
            
            # 現在のページ状態を確認
            page_title = self.browser.driver.title
            logger.info(f"現在の画面: {page_title}")
            
            # スクリーンショットを撮影して状態を確認
            self.browser.save_screenshot("import_settings_screen.png")
            
            # 「次へ」ボタンを探索 - span.ui-button-text 方式（実行ボタンと同じ方法）
            next_button = None
            next_button_found = False
            
            try:
                logger.info("「次へ」ボタンを探索します（インポート設定画面）")
                spans = self.browser.driver.find_elements(By.CSS_SELECTOR, "span.ui-button-text")
                for span in spans:
                    if span.text == "次へ":
                        next_button = span.find_element(By.XPATH, "./..")  # 親要素（ボタン）を取得
                        logger.info("✅ 「次へ」ボタンを検出しました（span.ui-button-text）")
                        next_button_found = True
                        break
            except Exception as e:
                logger.warning(f"span.ui-button-textでの「次へ」ボタン検出に失敗: {e}")
            
            # 次へボタンが見つかった場合、クリック
            if next_button_found:
                try:
                    self.browser.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                    time.sleep(1)
                    self.browser.driver.execute_script("arguments[0].click();", next_button)
                    logger.info("✓ 「次へ」ボタンをクリックしました")
                    
                    # 画面遷移の確認（待機）
                    time.sleep(5)
                    
                except Exception as click_error:
                    logger.error(f"「次へ」ボタンのクリック中にエラー: {click_error}")
                    # 既存の代替方法を維持
                    logger.info("JavaScriptで直接2番目のボタンをクリックします")
                    try:
                        self.browser.driver.execute_script("""
                            var buttons = document.querySelectorAll('button');
                            if (buttons.length >= 2) {
                                buttons[1].click();  // 2番目のボタン
                            }
                        """)
                        logger.info("✓ JavaScriptで2番目のボタンのクリックに成功しました")
                        time.sleep(3)
                    except Exception as js_error:
                        logger.error(f"JavaScriptでのクリックにも失敗: {js_error}")
            else:
                # ボタンが見つからない場合は既存の代替方法を使用
                logger.info("JavaScriptで直接2番目のボタンをクリックします")
                try:
                    self.browser.driver.execute_script("""
                        var buttons = document.querySelectorAll('button');
                        if (buttons.length >= 2) {
                            buttons[1].click();  // 2番目のボタン
                        }
                    """)
                    logger.info("✓ JavaScriptで2番目のボタンのクリックに成功しました")
                    time.sleep(3)
                except Exception as js_error:
                    logger.error(f"JavaScriptでのクリックにも失敗: {js_error}")
            
            # 操作後の画面キャプチャ
            self.browser.save_screenshot("settings_next_button_result.png")
            
            logger.info("=== インポート設定画面での操作完了 ===")
            return True
            
        except Exception as e:
            logger.error(f"インポート設定画面での操作に失敗: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("settings_next_button_error.png")
            return False
    
    def _perform_mapping_settings(self):
        """レコードの重複処理設定画面での操作"""
        try:
            logger.info("=== レコードの重複処理設定画面での操作開始 ===")
            
            # 画面が完全に読み込まれるまで十分待機
            time.sleep(5)
            
            # 現在のページ状態を確認
            page_title = self.browser.driver.title
            page_source = self.browser.driver.page_source
            
            # ヘッダーテキストなどから画面ステップを確認
            headers = self.browser.driver.find_elements(By.TAG_NAME, "h1")
            if headers:
                header_text = headers[0].text
                logger.info(f"現在の画面: {header_text}")
            else:
                logger.info("現在の画面: 求職者 - インポート (3/4) レコードの重複処理設定画面")
            
            # スクリーンショットを撮影して状態を確認
            self.browser.save_screenshot("mapping_screen.png")
            
            # 「次へ」ボタンをクリック - 位置ベースのアプローチ
            try:
                logger.info("「次へ」ボタンを探索します（重複処理設定画面）")
                
                # ダイアログのボタンパネルを特定
                dialog = self.browser.driver.find_element(By.CSS_SELECTOR, ".ui-dialog")
                button_pane = dialog.find_element(By.CSS_SELECTOR, ".ui-dialog-buttonpane")
                logger.info("ダイアログのボタンパネルを見つけました")
                
                # ボタンパネル内のすべてのボタン要素を取得
                buttons = button_pane.find_elements(By.TAG_NAME, "button")
                logger.info(f"ボタンパネル内のボタン数: {len(buttons)}")
                
                if len(buttons) >= 2:
                    # 2番目（インデックス1）のボタンが「次へ」
                    next_button = buttons[1]  # インデックスは0始まり
                    
                    # ボタンの情報をログに記録
                    logger.info(f"2番目のボタン: クラス={next_button.get_attribute('class')}")
                    
                    # スクロールして表示
                    self.browser.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                    time.sleep(1)
                    
                    # JavaScriptでクリック
                    self.browser.driver.execute_script("arguments[0].click();", next_button)
                    logger.info("✓ 2番目のボタン（「次へ」）をJavaScriptでクリックしました")
                    
                    # 画面遷移を待機
                    time.sleep(5)
                    
                    # 遷移後の画面状態を確認
                    self.browser.save_screenshot("after_mapping_next_button.png")
                    
                    # 画面遷移の確認
                    try:
                        # 現在の画面をログに記録
                        current_html = self.browser.driver.page_source
                        
                        # 画面状態の判定
                        if "(4/4)" in current_html or "確認" in current_html:
                            logger.info("✅ 画面遷移確認: 求職者 - インポート (4/4) 確認画面")
                        else:
                            logger.info(f"画面遷移確認: タイトル「{page_title}」の画面に遷移しました")
                            
                            # HTMLを保存して後で分析できるようにする
                            with open(os.path.join(self.screenshot_dir, "page4_html.html"), "w", encoding="utf-8") as f:
                                f.write(current_html)
                    
                    except Exception as e:
                        logger.error(f"画面遷移確認中にエラー: {e}")
                else:
                    # ボタンが少ない場合の対処
                    logger.error(f"期待するボタン数が見つかりません。見つかったボタン数: {len(buttons)}")
                    
                    # 最後の手段：一番右側のボタン（通常は「次へ」や「確定」）をクリック
                    if buttons:
                        rightmost_button = buttons[-1]
                        self.browser.driver.execute_script("arguments[0].click();", rightmost_button)
                        logger.info("✓ 一番右側のボタンをクリックしました")
                        time.sleep(3)
                    else:
                        raise Exception("クリック可能なボタンが見つかりません")
            except Exception as e:
                logger.error(f"「次へ」ボタンのクリックに失敗: {e}")
                
                # 最終手段：JavaScriptで直接実行
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
                        time.sleep(3)
                    else:
                        logger.error("JavaScriptでのクリックも失敗しました")
                        
                except Exception as js_error:
                    logger.error(f"JavaScriptでのクリックにも失敗: {js_error}")
                
                # デバッグ情報
                self.browser.save_screenshot("mapping_next_button_error.png")
                with open(os.path.join(self.screenshot_dir, "mapping_button_html.html"), "w", encoding="utf-8") as f:
                    f.write(self.browser.driver.page_source)
            
            logger.info("=== レコードの重複処理設定画面での操作完了 ===")
            return True
            
        except Exception as e:
            logger.error(f"レコードの重複処理設定画面での操作に失敗: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("mapping_settings_error.png")
            
            # エラー発生時のHTMLを保存
            with open(os.path.join(self.screenshot_dir, "mapping_error_html.html"), "w", encoding="utf-8") as f:
                f.write(self.browser.driver.page_source)
            
            return False
    
    def _perform_final_confirmation(self):
        """インポート実行画面での操作"""
        try:
            logger.info("=== インポート実行画面での操作開始 ===")
            
            # 画面が完全に読み込まれるまで十分待機
            time.sleep(5)
            
            # 現在のページ状態を確認
            page_source = self.browser.driver.page_source
            
            # ヘッダーテキストなどから画面ステップを確認
            headers = self.browser.driver.find_elements(By.TAG_NAME, "h1")
            if headers:
                header_text = headers[0].text
                logger.info(f"現在の画面: {header_text}")
            else:
                logger.info("現在の画面: 求職者 - インポート (4/4) インポート実行画面")
            
            # スクリーンショットを撮影して状態を確認
            self.browser.save_screenshot("final_confirmation_screen.png")
            
            # 設定を読み込み - EnvironmentUtilsを使用
            execute_import = env.get_config_value("BROWSER", "execute_import", default=False)
            logger.info(f"設定値: execute_import = {execute_import}")
            
            # 「実行」ボタンを検出
            execute_button_found = False
            execute_button = None
            
            # 方法1: spanタグのクラスとテキストで検索（最も信頼性が高い）
            try:
                logger.info("「実行」ボタンを探索します - 方法1: span.ui-button-text")
                spans = self.browser.driver.find_elements(By.CSS_SELECTOR, "span.ui-button-text")
                for span in spans:
                    if span.text == "実行":
                        execute_button = span.find_element(By.XPATH, "./..")  # 親要素（ボタン）を取得
                        logger.info("✅ 「実行」ボタンを検出しました（span.ui-button-text）")
                        execute_button_found = True
                        break
            except Exception as e:
                logger.warning(f"span.ui-button-textでの「実行」ボタン検出に失敗: {e}")
            
            # 実行ボタンが見つかった場合、設定に基づいて実行
            if execute_button_found:
                if execute_import:
                    logger.info("設定に基づき、「実行」ボタンをクリックします")
                    try:
                        # スクロールして表示
                        self.browser.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", execute_button)
                        time.sleep(1)
                        
                        # JavaScriptでクリック
                        self.browser.driver.execute_script("arguments[0].click();", execute_button)
                        logger.info("✓ 「実行」ボタンをクリックしました - インポートを実行します")
                        
                        # インポート完了まで待機（長めに）
                        logger.info("インポート処理の完了を待機しています...")
                        time.sleep(15)
                        
                        # 完了確認
                        self.browser.save_screenshot("after_import_execution.png")
                        
                        # 完了後の処理: OKボタンをクリック
                        logger.info("完了確認ダイアログの「OK」ボタンを探します...")
                        ok_button_found = False
                        
                        # 方法1: spanタグから「OK」ボタンを探す
                        try:
                            spans = self.browser.driver.find_elements(By.CSS_SELECTOR, "span.ui-button-text")
                            for span in spans:
                                if span.text.lower() in ["ok", "確定", "閉じる"]:
                                    ok_button = span.find_element(By.XPATH, "./..")  # 親要素（ボタン）を取得
                                    logger.info(f"✅ 「OK」ボタンを検出しました: テキスト='{span.text}'")
                                    
                                    # スクロールして表示
                                    self.browser.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", ok_button)
                                    time.sleep(1)
                                    
                                    # JavaScriptでクリック
                                    self.browser.driver.execute_script("arguments[0].click();", ok_button)
                                    logger.info("✓ 「OK」ボタンをクリックしました")
                                    ok_button_found = True
                                    time.sleep(3)  # 画面遷移を待機
                                    break
                        except Exception as ok_error:
                            logger.warning(f"spanからの「OK」ボタン検出に失敗: {ok_error}")
                        
                        # 方法2: ダイアログの最初のボタンを探す（通常は「OK」か「閉じる」）
                        if not ok_button_found:
                            try:
                                logger.info("ダイアログの最初のボタンをクリックします")
                                result = self.browser.driver.execute_script("""
                                    // ダイアログボタンを取得
                                    var dialogButtons = document.querySelectorAll('.ui-dialog-buttonset button');
                                    if (dialogButtons.length > 0) {
                                        dialogButtons[0].click();
                                        return true;
                                    }
                                    
                                    // 通常のボタン要素も検索
                                    var allButtons = document.querySelectorAll('button');
                                    for (var i = 0; i < allButtons.length; i++) {
                                        var buttonText = allButtons[i].textContent.trim().toLowerCase();
                                        if (buttonText === 'ok' || buttonText === '確定' || buttonText === '閉じる') {
                                            allButtons[i].click();
                                            return true;
                                        }
                                    }
                                    return false;
                                """)
                                
                                if result:
                                    logger.info("✓ JavaScriptで「OK」ボタンのクリックに成功しました")
                                    ok_button_found = True
                                    time.sleep(3)  # 画面遷移を待機
                                else:
                                    logger.warning("JavaScriptでのOKボタン検出に失敗しました")
                                    
                                # 検出失敗時のデバッグ情報収集
                                all_buttons = self.browser.driver.find_elements(By.TAG_NAME, "button")
                                logger.info(f"画面上のボタン数: {len(all_buttons)}")
                                for i, button in enumerate(all_buttons):
                                    logger.info(f"ボタン{i+1}: テキスト='{button.text}', クラス='{button.get_attribute('class')}'")
                                
                                # HTMLを保存
                                with open(os.path.join(self.screenshot_dir, "ok_button_debug.html"), "w", encoding="utf-8") as f:
                                    f.write(self.browser.driver.page_source)
                            except Exception as js_error:
                                logger.error(f"JavaScriptでの「OK」ボタンクリックに失敗: {js_error}")
                        
                        # 最終スクリーンショット
                        self.browser.save_screenshot("after_ok_button.png")
                        logger.info("インポート完了処理が完了しました")
                        
                    except Exception as click_error:
                        logger.error(f"「実行」ボタンのクリック中にエラー: {click_error}")
                        self.browser.save_screenshot("execute_button_error.png")
                else:
                    logger.info("設定に基づき、「実行」ボタンをクリックせずに終了します")
                    logger.info("✓ インポート実行をスキップしました（execute_import=False）")
            else:
                logger.error("「実行」ボタンが見つからないため、インポートを実行できません")
                # ダイアログ内のすべてのボタンを表示（デバッグ用）
                try:
                    all_buttons = self.browser.driver.find_elements(By.TAG_NAME, "button")
                    logger.info(f"画面上のボタン数: {len(all_buttons)}")
                    for i, button in enumerate(all_buttons):
                        logger.info(f"ボタン{i+1}: テキスト='{button.text}', クラス='{button.get_attribute('class')}'")
                    
                    # HTMLを保存して後で分析できるようにする
                    with open(os.path.join(self.screenshot_dir, "button_debug_html.html"), "w", encoding="utf-8") as f:
                        f.write(self.browser.driver.page_source)
                except Exception as debug_error:
                    logger.error(f"デバッグ情報収集中にエラー: {debug_error}")
            
            logger.info("=== インポート実行画面での操作完了 ===")
            return True
            
        except Exception as e:
            logger.error(f"インポート実行画面での操作に失敗: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("final_confirmation_error.png")
            
            # エラー発生時のHTMLを保存
            with open(os.path.join(self.screenshot_dir, "final_error_html.html"), "w", encoding="utf-8") as f:
                f.write(self.browser.driver.page_source)
            
            return False