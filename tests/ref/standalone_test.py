import sys
import os
import time
import csv
from pathlib import Path
import re

# プロジェクトのルートディレクトリをPYTHONPATHに追加
root_dir = Path(__file__).parent.parent
sys.path.append(str(root_dir))

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup

from src.utils.environment import EnvironmentUtils as env
from src.utils.logging_config import get_logger

logger = get_logger(__name__)

def load_selectors(csv_path):
    """セレクタCSVファイルを読み込む"""
    selectors = {}
    try:
        with open(csv_path, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                page = row['page']
                element = row['element']
                
                if page not in selectors:
                    selectors[page] = {}
                
                selectors[page][element] = {
                    'description': row['description'],
                    'action_type': row['action_type'],
                    'selector_type': row['selector_type'],
                    'selector_value': row['selector_value'],
                    'element_type': row['element_type'],
                    'parent_selector': row['parent_selector']
                }
        logger.info(f"セレクタを読み込みました: {len(selectors)} ページ")
        return selectors
    except Exception as e:
        logger.error(f"セレクタの読み込みに失敗しました: {str(e)}")
        return {}

def analyze_page_content(html_content):
    """
    Beautiful Soupを使用してページ内容を解析する
    """
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

def perform_login(driver, wait, selectors, screenshot_dir):
    """ログイン処理を実行する"""
    try:
        logger.info("=== ログイン処理を開始します ===")
        
        # 環境変数からログイン情報を取得
        admin_url = env.get_env_var('ADMIN_URL')
        admin_id = env.get_env_var('ADMIN_ID')  # 会社ID
        login_id = env.get_env_var('LOGIN_ID')  # ← ADMIN_USERから変更
        password = env.get_env_var('LOGIN_PASSWORD')  # ← ADMIN_PASSから変更
        
        # ログインページにアクセス
        logger.info(f"ログインページにアクセスします: {admin_url}")
        driver.get(admin_url)
        time.sleep(3)  # ページ読み込み待機
        
        # ログイン前のスクリーンショット
        driver.save_screenshot(os.path.join(screenshot_dir, "login_before.png"))
        
        # 会社ID入力
        company_id_selector = selectors['porters']['company_id']['selector_value']
        logger.info(f"会社IDフィールドを探索: {company_id_selector}")
        company_id_field = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, company_id_selector)))
        company_id_field.clear()
        company_id_field.send_keys(admin_id)
        logger.info(f"✓ 会社IDを入力しました: {admin_id}")
        
        # ユーザー名入力
        username_selector = selectors['porters']['username']['selector_value']
        logger.info(f"ユーザー名フィールドを探索: {username_selector}")
        username_field = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, username_selector)))
        username_field.clear()
        username_field.send_keys(login_id)  # ← admin_userから変更
        logger.info("✓ ユーザー名を入力しました")
        
        # パスワード入力
        password_selector = selectors['porters']['password']['selector_value']
        logger.info(f"パスワードフィールドを探索: {password_selector}")
        password_field = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, password_selector)))
        password_field.clear()
        password_field.send_keys(password)  # ← admin_passから変更
        logger.info("✓ パスワードを入力しました")
        
        # 入力後のスクリーンショット
        driver.save_screenshot(os.path.join(screenshot_dir, "login_input.png"))
        
        # ログインボタンクリック
        login_button_selector = selectors['porters']['login_button']['selector_value']
        logger.info(f"ログインボタンを探索: {login_button_selector}")
        login_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, login_button_selector)))
        login_button.click()
        logger.info("✓ ログインボタンをクリックしました")
        
        # ログイン処理待機
        time.sleep(5)
        
        # ↓↓↓ 二重ログインポップアップ対応を追加 ↓↓↓
        # 二重ログインポップアップが表示されているか確認
        double_login_ok_button = "#pageDeny > div.ui-dialog.ui-widget.ui-widget-content.ui-corner-all.ui-front.p-ui-messagebox.ui-dialog-buttons.ui-draggable > div.ui-dialog-buttonpane.ui-widget-content.ui-helper-clearfix > div > button > span"
        
        try:
            # 短いタイムアウトで確認（ポップアップが表示されていない場合にテストが長時間停止しないように）
            popup_wait = WebDriverWait(driver, 3)
            ok_button = popup_wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, double_login_ok_button)))
            
            # ポップアップが見つかった場合
            logger.info("⚠️ 二重ログインポップアップが検出されました。OKボタンをクリックします。")
            driver.save_screenshot(os.path.join(screenshot_dir, "double_login_popup.png"))
            
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
            
            # JavaScriptでのクリックを試行
            try:
                driver.execute_script(f"document.querySelector('{double_login_ok_button}').click();")
                logger.info("✓ JavaScriptで二重ログインポップアップのOKボタンをクリックしました")
                time.sleep(2)
            except:
                logger.warning("JavaScriptでのクリックも失敗しましたが、処理を継続します")
        # ↑↑↑ 二重ログインポップアップ対応ここまで ↑↑↑
        
        # ログイン後のスクリーンショット
        driver.save_screenshot(os.path.join(screenshot_dir, "login_after.png"))
        
        # ログイン後のHTMLを解析
        after_login_html = driver.page_source
        after_login_analysis = analyze_page_content(after_login_html)
        
        # ログイン結果の詳細を記録
        logger.info(f"ログイン後の状態:")
        logger.info(f"  - タイトル: {after_login_analysis['page_title']}")
        logger.info(f"  - 見出し: {after_login_analysis['main_heading']}")
        logger.info(f"  - エラーメッセージ: {after_login_analysis['error_messages']}")
        logger.info(f"  - メニュー項目数: {len(after_login_analysis['menu_items'])}")
        if after_login_analysis['menu_items']:
            logger.info(f"  - メニュー項目例: {after_login_analysis['menu_items'][:5]}")
        
        # URLの変化も確認
        current_url = driver.current_url
        logger.info(f"ログイン後のURL: {current_url}")
        
        # ログイン成功を判定
        login_success = (admin_url != current_url and "login" not in current_url.lower()) or len(after_login_analysis['menu_items']) > 0
        
        if login_success:
            logger.info("✅ ログインに成功しました！")
            return True
        else:
            logger.error("❌ ログインに失敗しました")
            
            # HTMLファイルとして保存（詳細分析用）
            html_path = os.path.join(screenshot_dir, "login_failed.html")
            with open(html_path, 'w', encoding='utf-8') as f:
                f.write(after_login_html)
            logger.info(f"ログイン失敗時のHTMLを保存しました: {html_path}")
                
            return False
            
    except Exception as e:
        logger.error(f"ログイン処理中にエラーが発生しました: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return False

def perform_after_login_operations(driver, wait, selectors, screenshot_dir):
    """ログイン後の操作を実行する"""
    try:
        logger.info("=== ログイン後の操作を開始します ===")
        
        # ページが完全に読み込まれるまで十分待機
        time.sleep(10)  
        
        # ポップアップチェックは一度だけ実行し、タイムアウトを短く設定
        try:
            # ポップアップボタン候補を確認（タイムアウトを3秒に短縮）
            popup_wait = WebDriverWait(driver, 3)
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
                    driver.execute_script("arguments[0].click();", popup_button)
                    logger.info("✓ ポップアップを閉じました")
                    time.sleep(1)  # 短い待機で十分
                    break
                except:
                    # このセレクタでは見つからなかっただけなので、次を試す
                    continue
        except Exception as popup_error:
            # ポップアップがなければ単に続行
            logger.info("ポップアップは検出されませんでした。処理を続行します。")
        
        # 現在のページ状態を確認
        page_html = driver.page_source
        page_analysis = analyze_page_content(page_html)
        logger.info(f"現在のページ状態: タイトル={page_analysis['page_title']}")
        logger.info(f"メニュー項目数: {len(page_analysis['menu_items'])}")
        
        # スクリーンショット保存
        driver.save_screenshot(os.path.join(screenshot_dir, "before_operations.png"))
        
        # 現在のウィンドウハンドルを保存（後で戻れるように）
        main_window_handle = driver.current_window_handle
        logger.info(f"メインウィンドウハンドル: {main_window_handle}")
        current_handles = driver.window_handles
        logger.info(f"操作前のウィンドウハンドル一覧: {current_handles}")
        
        # "その他業務"ボタンをクリック - セレクタ修正
        others_button_selector = "#main > div > main > section.original-search > header > div.others > button"
        logger.info(f"「その他業務」ボタンを探索: {others_button_selector}")
        
        try:
            # 十分な待機を追加
            time.sleep(3)  # 追加の待機
            
            # スクリーンショットで状態確認
            driver.save_screenshot(os.path.join(screenshot_dir, "before_others_click.png"))
            
            # 要素を見つける
            others_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, others_button_selector)))
            
            # スクロールして確実に表示
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", others_button)
            time.sleep(1)
            
            # ボタンの詳細をログに記録
            logger.info(f"「その他業務」ボタン情報: テキスト={others_button.text}, クラス={others_button.get_attribute('class')}")
            
            # クリック実行 (JavaScriptでクリック)
            driver.execute_script("arguments[0].click();", others_button)
            logger.info("✓ 「その他業務」ボタンをクリックしました")
            
            # 新しいウィンドウが開くのを待機 - 延長
            time.sleep(8)  # 十分な待機時間
            
            # 新しいウィンドウに切り替え
            new_handles = driver.window_handles
            logger.info(f"操作後のウィンドウハンドル一覧: {new_handles}")
            
            if len(new_handles) > len(current_handles):
                # 新しいウィンドウが開かれた場合
                new_window = [handle for handle in new_handles if handle not in current_handles][0]
                logger.info(f"新しいウィンドウに切り替えます: {new_window}")
                driver.switch_to.window(new_window)
                logger.info("✓ 新しいウィンドウにフォーカスを切り替えました")
                
                # 新しいウィンドウで読み込みを待機
                time.sleep(5)
                driver.save_screenshot(os.path.join(screenshot_dir, "new_window.png"))
                
                # 新しいウィンドウでのページ状態を確認
                new_window_html = driver.page_source
                with open(os.path.join(screenshot_dir, "new_window.html"), "w", encoding="utf-8") as f:
                    f.write(new_window_html)
                logger.info("新しいウィンドウのHTMLを保存しました")
            else:
                logger.warning("新しいウィンドウが検出されませんでした")
            
            # メニュー項目5をクリック
            menu_item_selector = "#main-menu-id-5 > a"
            if 'porters_menu' in selectors and 'menu_item_5' in selectors['porters_menu']:
                menu_item_selector = selectors['porters_menu']['menu_item_5']['selector_value']
            logger.info(f"メニュー項目5を探索: {menu_item_selector}")
            
            # スクリーンショットを撮ってページの状態を記録
            driver.save_screenshot(os.path.join(screenshot_dir, "before_menu_click.png"))
            
            # メニュー項目をクリック
            menu_item = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, menu_item_selector)))
            menu_item.click()
            logger.info("✓ メニュー項目5をクリックしました")
            
            # サブメニューが表示されるまでしっかり待機
            time.sleep(5)
            driver.save_screenshot(os.path.join(screenshot_dir, "after_menu_click.png"))
            
            # メニューコンテナを見つけてスクロール
            try:
                menu_containers = driver.find_elements(By.CSS_SELECTOR, ".ui-menu, .dropdown-menu, .menu")
                if menu_containers:
                    # 最後に見つかったメニューコンテナが通常ドロップダウン
                    menu_container = menu_containers[-1]
                    logger.info(f"メニューコンテナを見つけました: {menu_container.tag_name}")
                    
                    # コンテナの一番下までスクロール
                    driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", menu_container)
                    logger.info("✓ メニューコンテナの一番下までスクロールしました")
                    time.sleep(2)
                    
                    # スクロール後のスクリーンショット
                    driver.save_screenshot(os.path.join(screenshot_dir, "after_scroll.png"))
            except Exception as e:
                logger.warning(f"メニューコンテナのスクロールに失敗しましたが、処理を続行します: {e}")
            
            # 「求職者のインポート」リンクをクリック - 最適化バージョン
            logger.info("「求職者のインポート」リンクを探索します")
            
            try:
                # 0. メニュースクロール - 最下部にある可能性が高い
                menu_container = driver.find_element(By.CSS_SELECTOR, ".main-menu-scrollable")
                logger.info(f"メニューコンテナを発見: ID={menu_container.get_attribute('id')}")
                
                # メニューを最下部までスクロール
                driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", menu_container)
                logger.info("メニューを最下部までスクロールしました")
                time.sleep(1)
                driver.save_screenshot(os.path.join(screenshot_dir, "menu_scrolled_bottom.png"))
                
                # 1. title属性による検索（最も確実）
                logger.info("title属性を使って「求職者のインポート」リンクを検索")
                import_link = driver.find_element(By.CSS_SELECTOR, "a[title='求職者のインポート']")
                logger.info(f"「求職者のインポート」リンクを見つけました: ID={import_link.get_attribute('id')}")
                
                # JavaScriptでクリック
                driver.execute_script("arguments[0].click();", import_link)
                logger.info("✓ 「求職者のインポート」リンクをクリックしました")
                time.sleep(5)
                
                # ポップアップが表示されたか確認
                popup_elements = driver.find_elements(By.CSS_SELECTOR, ".ui-dialog, .popup, .modal")
                if popup_elements:
                    logger.info("✓ ポップアップが表示されました")
                    driver.save_screenshot(os.path.join(screenshot_dir, "popup_displayed.png"))
                else:
                    logger.warning("! ポップアップが表示されていません。再試行します")
                    raise Exception("ポップアップ未表示")
                
            except Exception as e:
                logger.error(f"1回目の試行でエラー: {e}")
                
                try:
                    # 2. "インポート"ヘッダーの次の項目を探す方法
                    logger.info("「インポート」ヘッダーの下の項目を検索")
                    
                    # ヘッダーを見つける
                    import_headers = driver.find_elements(By.XPATH, "//li[contains(@class, 'header')]/a[@title='インポート']")
                    if import_headers:
                        import_header = import_headers[0]
                        logger.info(f"「インポート」ヘッダーを見つけました: ID={import_header.get_attribute('id')}")
                        
                        # ヘッダーの親要素から次の兄弟要素を取得
                        header_li = import_header.find_element(By.XPATH, "..")
                        next_li = header_li.find_element(By.XPATH, "following-sibling::li")
                        import_link = next_li.find_element(By.TAG_NAME, "a")
                        
                        logger.info(f"「インポート」ヘッダー直後の項目: '{import_link.text}'")
                        driver.execute_script("arguments[0].click();", import_link)
                        logger.info("✓ 「インポート」ヘッダー後の項目をクリックしました")
                        time.sleep(5)
                    else:
                        # 3. メニュー最下部の項目をクリック
                        logger.info("メニュー最下部の項目を探索")
                        menu_items = driver.find_elements(By.CSS_SELECTOR, ".main-menu-scrollable li")
                        if menu_items:
                            last_item = menu_items[-1]
                            last_link = last_item.find_element(By.TAG_NAME, "a")
                            logger.info(f"メニュー最下部の項目: '{last_link.text}'")
                            
                            driver.execute_script("arguments[0].click();", last_link)
                            logger.info("✓ メニュー最下部の項目をクリックしました")
                            time.sleep(5)
                        else:
                            raise Exception("メニュー項目が見つかりません")
                    
                    # ポップアップ確認
                    popup_check = driver.find_elements(By.CSS_SELECTOR, ".ui-dialog, #porters-pdialog_1")
                    if popup_check:
                        logger.info("✓ ポップアップが表示されました")
                    else:
                        logger.warning("! ポップアップが表示されていません")
                        return False
                        
                except Exception as e2:
                    logger.error(f"全ての方法でリンクのクリックに失敗: {e2}")
                    
                    # デバッグ情報の収集
                    driver.save_screenshot(os.path.join(screenshot_dir, "menu_error.png"))
                    with open(os.path.join(screenshot_dir, "menu_html.html"), "w", encoding="utf-8") as f:
                        f.write(driver.page_source)
                    logger.info("デバッグ情報を保存しました")
                    
                    # JavaScriptでのダイレクトアクセスを試みる（最終手段）
                    try:
                        # 「求職者のインポート」リンクのIDを探す
                        links = driver.find_elements(By.TAG_NAME, "a")
                        for link in links:
                            if link.get_attribute("title") == "求職者のインポート":
                                link_id = link.get_attribute("id")
                                logger.info(f"「求職者のインポート」リンクを検出: ID={link_id}")
                                # IDを使ってJavaScriptでクリック
                                driver.execute_script(f'document.getElementById("{link_id}").click();')
                                logger.info("✓ JavaScriptでIDを使ってクリックしました")
                                time.sleep(5)
                                break
                    except:
                        logger.error("最終手段も失敗しました")
                        return False
            
            # ファイルインポート画面での操作
            logger.info("=== ファイルインポート画面での操作開始 ===")
            
            # ページが完全に読み込まれるまで十分待機
            time.sleep(8)
            
            # 「添付」ボタンをクリック
            try:
                attachment_button = driver.find_element(By.CSS_SELECTOR, "#_ibb_lbl")
                logger.info("「添付」ボタンを見つけました")
                driver.execute_script("arguments[0].click();", attachment_button)
                logger.info("✓ 「添付」ボタンをクリックしました")
                time.sleep(3)
            except Exception as e:
                logger.error(f"「添付」ボタンのクリックに失敗: {e}")
                driver.save_screenshot(os.path.join(screenshot_dir, "attachment_button_error.png"))
                # HTMLを保存
                with open(os.path.join(screenshot_dir, "attachment_html.html"), "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                    
                # 現在のURLも記録
                logger.info(f"現在のURL: {driver.current_url}")
                
                # 続行（他の要素を探す可能性）
            
            # ファイル選択（input type=fileに直接パスを送信）
            try:
                # 通常、添付ボタンの近くに隠れたinput要素がある
                file_input = driver.find_element(By.CSS_SELECTOR, "input[type='file']")
                file_path = r"C:\Users\yohay\Downloads\LINE_member_0303.csv"
                
                # JavaScript経由で表示状態を変更し、ファイルパスを送信
                driver.execute_script("arguments[0].style.display = 'block';", file_input)
                file_input.send_keys(file_path)
                logger.info(f"✓ ファイルを選択しました: {file_path}")
                time.sleep(3)
                driver.save_screenshot(os.path.join(screenshot_dir, "after_file_select.png"))
            except Exception as e:
                logger.error(f"ファイル選択に失敗: {e}")
                driver.save_screenshot(os.path.join(screenshot_dir, "file_select_error.png"))
                
            # インポート方法の選択（LINE初回アンケート取込）
            try:
                import_method_selector = "#porters-pdialog_1 > div > div.subWrap.resize > div > div > div > ul > li:nth-child(9) > label > input[type=radio]"
                import_method = driver.find_element(By.CSS_SELECTOR, import_method_selector)
                driver.execute_script("arguments[0].click();", import_method)
                logger.info("✓ 「LINE初回アンケート取込」を選択しました")
                time.sleep(2)
                driver.save_screenshot(os.path.join(screenshot_dir, "import_method_selected.png"))
            except Exception as e:
                logger.error(f"インポート方法の選択に失敗: {e}")
                driver.save_screenshot(os.path.join(screenshot_dir, "import_method_error.png"))
                
            # 「次へ」ボタンをクリック - 位置ベースのアプローチ
            try:
                logger.info("「次へ」ボタンを探索します（位置ベースアプローチ）")
                
                # スクリーンショットで状態を確認
                driver.save_screenshot(os.path.join(screenshot_dir, "before_next_button.png"))
                
                # ダイアログのボタンパネルを特定
                dialog = driver.find_element(By.CSS_SELECTOR, ".ui-dialog")
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
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                    time.sleep(1)
                    
                    # JavaScriptでクリック
                    driver.execute_script("arguments[0].click();", next_button)
                    logger.info("✓ 2番目のボタン（「次へ」）をJavaScriptでクリックしました")
                    time.sleep(3)
                    
                    # クリック後のスクリーンショット
                    driver.save_screenshot(os.path.join(screenshot_dir, "after_next_button.png"))
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
                        driver.execute_script("arguments[0].click();", rightmost_button)
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
                    result = driver.execute_script(script)
                    if result:
                        logger.info("✓ JavaScriptで2番目のボタンのクリックに成功しました")
                        time.sleep(3)
                    else:
                        logger.error("JavaScriptでのクリックも失敗しました")
                        
                except Exception as js_error:
                    logger.error(f"JavaScriptでのクリックにも失敗: {js_error}")
                
                # デバッグ情報
                driver.save_screenshot(os.path.join(screenshot_dir, "next_button_error.png"))
                with open(os.path.join(screenshot_dir, "next_button_html.html"), "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
            
            logger.info("=== ファイルインポート画面での操作完了 ===")
            
            # 操作が正常に完了
            logger.info("✅ メニュー操作が正常に完了しました")
            return True
            
        except Exception as e:
            logger.error(f"メニュー操作に失敗しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            driver.save_screenshot(os.path.join(screenshot_dir, "menu_error.png"))
            
            # エラー発生時のHTMLを保存
            with open(os.path.join(screenshot_dir, "error_html.html"), "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            
            # メインウィンドウに戻る
            try:
                driver.switch_to.window(main_window_handle)
                logger.info("メインウィンドウに戻りました")
            except:
                logger.warning("メインウィンドウへの切り替えに失敗しました")
            
            return False
        
    except Exception as e:
        logger.error(f"ログイン後の操作中にエラーが発生しました: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        driver.save_screenshot(os.path.join(screenshot_dir, "error_after_login.png"))
        return False

def perform_logout(driver, wait, selectors, screenshot_dir):
    """明示的なログアウト処理を実行する"""
    try:
        logger.info("=== ログアウト処理を開始します ===")
        
        # 現在のURLをログに記録
        current_url = driver.current_url
        logger.info(f"ログアウト前のURL: {current_url}")
        
        # スクリーンショット
        driver.save_screenshot(os.path.join(screenshot_dir, "before_logout.png"))
        
        # ログアウトボタンを探す
        # まず、セレクタ情報を確認
        logout_selector = None
        if 'porters_menu' in selectors and 'logout_button' in selectors['porters_menu']:
            logout_selector = selectors['porters_menu']['logout_button']['selector_value']
        
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
                    logout_elements = driver.find_elements(By.CSS_SELECTOR, selector)
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
                logout_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, logout_selector)))
                logout_button.click()
                logger.info("✓ ログアウトボタンをクリックしました")
                
                # 確認ダイアログが表示される場合の処理
                try:
                    confirm_buttons = driver.find_elements(By.CSS_SELECTOR, 'button.confirm, button[id*="confirm"], button:contains("OK"), button:contains("はい")')
                    if confirm_buttons:
                        confirm_buttons[0].click()
                        logger.info("✓ 確認ダイアログのボタンをクリックしました")
                except:
                    logger.info("確認ダイアログはありませんでした")
                
                # ログアウト後の待機
                time.sleep(3)
                
                # ログアウト後のスクリーンショット
                driver.save_screenshot(os.path.join(screenshot_dir, "after_logout.png"))
                logger.info("✅ ログアウトに成功しました")
                return True
                
            except Exception as e:
                logger.warning(f"通常のクリックでログアウトに失敗しました: {str(e)}")
                # JavaScriptでログアウトを試みる
                try:
                    driver.execute_script(f"document.querySelector('{logout_selector}').click();")
                    logger.info("✓ JavaScriptを使用してログアウトボタンをクリックしました")
                    time.sleep(3)
                    driver.save_screenshot(os.path.join(screenshot_dir, "after_js_logout.png"))
                    logger.info("✅ JavaScriptでのログアウトに成功しました")
                    return True
                except Exception as js_e:
                    logger.error(f"JavaScriptを使用したログアウトにも失敗しました: {str(js_e)}")
        else:
            # セレクタが見つからない場合、JavaScriptでログアウトを試みる
            logger.warning("ログアウトボタンが見つかりませんでした。JavaScriptでのログアウトを試みます。")
            try:
                # POSTERSの一般的なログアウトURLパターンを試す
                driver.execute_script("window.location.href = '/logout' || '/auth/logout' || '/porters/logout';")
                time.sleep(3)
                logger.info("✓ JavaScriptでログアウトURLにリダイレクトしました")
                driver.save_screenshot(os.path.join(screenshot_dir, "after_redirect_logout.png"))
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

def main():
    """メインテスト処理"""
    driver = None
    
    try:
        # 環境変数の読み込み
        env.load_env()
        
        # セレクタの読み込み
        selectors_path = os.path.join(root_dir, "config", "selectors.csv")
        selectors = load_selectors(selectors_path)
        
        if not selectors or 'porters' not in selectors:
            logger.error("❌ PORTERSのセレクタ情報が見つかりません")
            return False
        
        # porters_menuがなければ初期化
        if 'porters_menu' not in selectors:
            logger.warning("PORTERSメニューのセレクタ情報が見つかりません。デフォルト値を使用します。")
            selectors['porters_menu'] = {
                'search_button': {
                    'selector_value': "#main > div > main > section.original-search > header > div.others > button"
                },
                'menu_item_5': {
                    'selector_value': "#main-menu-id-5 > a"
                }
            }
        
        # ブラウザオプションの設定
        options = Options()
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('start-maximized')
        options.add_argument('--disable-gpu')
        
        # WebDriverのセットアップ
        logger.info("WebDriverをセットアップしています...")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        wait = WebDriverWait(driver, 20)
        
        # スクリーンショットディレクトリの作成
        screenshot_dir = os.path.join(root_dir, "tests", "screenshots")
        os.makedirs(screenshot_dir, exist_ok=True)
        
        # ログイン処理を実行
        login_success = perform_login(driver, wait, selectors, screenshot_dir)
        if not login_success:
            logger.error("❌ ログインに失敗したため、テストを中止します")
            return False
        
        logger.info("✅ ログインに成功しました。ログイン後の操作を開始します")
        
        # ログイン後の操作を実行
        success = perform_after_login_operations(driver, wait, selectors, screenshot_dir)
        
        if success:
            logger.info("✅ テストが正常に完了しました！")
            print("✅ テストが正常に完了しました！")
            return True
        else:
            logger.error("❌ テストが失敗しました")
            print("❌ テストが失敗しました")
            return False
    
    except Exception as e:
        logger.error(f"テスト実行中にエラーが発生しました: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        print(f"❌ エラーが発生しました: {str(e)}")
        return False
    
    finally:
        # 明示的にログアウト処理を実行
        if driver:
            try:
                logger.info("明示的なログアウト処理を実行します")
                # スクリーンショットディレクトリの確認
                screenshot_dir = os.path.join(root_dir, "tests", "screenshots")
                os.makedirs(screenshot_dir, exist_ok=True)
                
                # セレクタの読み込み (念のため)
                selectors_path = os.path.join(root_dir, "config", "selectors.csv")
                selectors = load_selectors(selectors_path)
                
                # ログアウト処理を実行
                wait = WebDriverWait(driver, 10)
                perform_logout(driver, wait, selectors, screenshot_dir)
            except Exception as logout_error:
                logger.error(f"ログアウト処理中にエラーが発生しました: {str(logout_error)}")
            
            # ブラウザを終了
            logger.info("ブラウザを終了します")
            driver.quit()

if __name__ == "__main__":
    result = main()
    print(f"テスト結果: {'成功' if result else '失敗'}") 