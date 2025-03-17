#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
PORTERSシステムの業務操作を管理するモジュール

このモジュールは、PORTERSシステムでの「その他業務」ボタンのクリックや
メニュー項目の選択など、業務操作に関する機能を提供します。
"""

import os
import time
import glob
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
from pathlib import Path
from typing import Optional

from src.utils.logging_config import get_logger
from src.utils.helpers import wait_for_new_csv_in_downloads, move_file_to_data_dir, find_latest_csv_in_downloads
from src.utils.spreadsheet import SpreadsheetManager

logger = get_logger(__name__)

class PortersOperations:
    """
    PORTERSシステムの業務操作を管理するクラス
    
    このクラスは、PORTERSシステムでの「その他業務」ボタンのクリックや
    メニュー項目の選択など、業務操作に関する機能を提供します。
    """
    
    def __init__(self, browser):
        """
        業務操作クラスの初期化
        
        Args:
            browser (PortersBrowser): ブラウザ操作を管理するインスタンス
        """
        self.browser = browser
        self.screenshot_dir = browser.screenshot_dir
    
    def click_other_operations_button(self):
        """
        「その他業務」ボタンをクリックして新しいウィンドウに切り替える
        
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== 「その他業務」ボタンのクリック処理を開始します ===")
            
            # 現在のウィンドウハンドルを保存
            current_handles = self.browser.driver.window_handles
            logger.info(f"現在のウィンドウハンドル: {current_handles}")
            
            # 「その他業務」ボタンをクリック
            if not self.browser.click_element('porters_menu', 'search_button', use_javascript=True):
                logger.error("「その他業務」ボタンのクリックに失敗しました")
                return False
            
            # 新しいウィンドウに切り替え
            if not self.browser.switch_to_new_window(current_handles):
                logger.error("新しいウィンドウへの切り替えに失敗しました")
                return False
            
            # 新しいウィンドウでのページ状態を確認
            new_window_html = self.browser.driver.page_source
            html_path = os.path.join(self.screenshot_dir, "new_window.html")
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(new_window_html)
            logger.info("新しいウィンドウのHTMLを保存しました")
            
            logger.info("✅ 「その他業務」ボタンのクリックと新しいウィンドウへの切り替えが完了しました")
            return True
            
        except Exception as e:
            logger.error(f"「その他業務」ボタンのクリック処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("other_operations_error.png")
            return False
    
    def click_menu_item_5(self):
        """
        メニュー項目5をクリックする
        
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== メニュー項目5のクリック処理を開始します ===")
            
            # メニュー項目5をクリック
            if not self.browser.click_element('porters_menu', 'menu_item_5'):
                logger.error("メニュー項目5のクリックに失敗しました")
                return False
            
            # 処理待機
            time.sleep(3)
            self.browser.save_screenshot("after_menu_item_5_click.png")
            
            logger.info("✅ メニュー項目5のクリック処理が完了しました")
            return True
            
        except Exception as e:
            logger.error(f"メニュー項目5のクリック処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("menu_item_5_error.png")
            return False
    
    def click_all_candidates(self):
        """
        「すべての求職者」リンクをクリックする
        
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== 「すべての求職者」リンクのクリック処理を開始します ===")
            
            # 「すべての求職者」リンクをクリック
            if not self.browser.click_element('porters_menu', 'all_candidates'):
                logger.error("「すべての求職者」リンクのクリックに失敗しました")
                
                # セレクタが動的に変わる可能性があるため、直接CSSセレクタを使用して再試行
                try:
                    logger.info("直接CSSセレクタを使用して「すべての求職者」リンクを探索します")
                    all_candidates_selector = "#ui-id-189 > li:nth-child(7) > a"
                    all_candidates_element = WebDriverWait(self.browser.driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, all_candidates_selector))
                    )
                    all_candidates_element.click()
                    logger.info("✓ 直接CSSセレクタを使用して「すべての求職者」リンクをクリックしました")
                except Exception as css_e:
                    logger.error(f"直接CSSセレクタを使用したクリックにも失敗しました: {str(css_e)}")
                    
                    # テキストで要素を探す最終手段
                    try:
                        logger.info("テキスト内容で「すべての求職者」リンクを探索します")
                        links = self.browser.driver.find_elements(By.TAG_NAME, "a")
                        for link in links:
                            if "すべての求職者" in link.text:
                                logger.info(f"「すべての求職者」テキストを含むリンクを発見しました: {link.text}")
                                link.click()
                                logger.info("✓ テキスト内容で「すべての求職者」リンクをクリックしました")
                                break
                        else:
                            logger.error("「すべての求職者」テキストを含むリンクが見つかりませんでした")
                            return False
                    except Exception as text_e:
                        logger.error(f"テキスト内容での探索にも失敗しました: {str(text_e)}")
                        return False
            
            # 処理待機
            time.sleep(3)
            self.browser.save_screenshot("after_all_candidates_click.png")
            
            # ページ内容を確認
            page_html = self.browser.driver.page_source
            page_analysis = self.browser.analyze_page_content(page_html)
            logger.info(f"ページタイトル: {page_analysis['page_title']}")
            
            logger.info("✅ 「すべての求職者」リンクのクリック処理が完了しました")
            return True
            
        except Exception as e:
            logger.error(f"「すべての求職者」リンクのクリック処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("all_candidates_error.png")
            return False
    
    def select_all_candidates(self):
        """
        求職者一覧画面で「全てチェック」チェックボックスをクリックする
        
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== 「全てチェック」チェックボックスのクリック処理を開始します ===")
            
            # 画面の読み込みを待機
            time.sleep(3)
            self.browser.save_screenshot("before_select_all.png")
            
            # 「全てチェック」チェックボックスをクリック
            if not self.browser.click_element('candidates_list', 'select_all_checkbox'):
                logger.error("「全てチェック」チェックボックスのクリックに失敗しました")
                
                # 直接CSSセレクタを使用して再試行
                try:
                    logger.info("直接CSSセレクタを使用して「全てチェック」チェックボックスを探索します")
                    checkbox_selector = "#recordListView input[type='checkbox']"
                    checkbox_element = WebDriverWait(self.browser.driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, checkbox_selector))
                    )
                    checkbox_element.click()
                    logger.info("✓ 直接CSSセレクタを使用して「全てチェック」チェックボックスをクリックしました")
                except Exception as css_e:
                    logger.error(f"直接CSSセレクタを使用したクリックにも失敗しました: {str(css_e)}")
                    
                    # より具体的なセレクタを試す
                    try:
                        logger.info("より具体的なセレクタを使用して「全てチェック」チェックボックスを探索します")
                        specific_selector = "#recordListView > div.jss37 > div:nth-child(2) > div > div.jss45 > span > span > input"
                        specific_element = WebDriverWait(self.browser.driver, 10).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, specific_selector))
                        )
                        specific_element.click()
                        logger.info("✓ より具体的なセレクタを使用して「全てチェック」チェックボックスをクリックしました")
                    except Exception as specific_e:
                        logger.error(f"より具体的なセレクタを使用したクリックにも失敗しました: {str(specific_e)}")
                        return False
            
            # 処理待機
            time.sleep(2)
            self.browser.save_screenshot("after_select_all.png")
            
            logger.info("✅ 「全てチェック」チェックボックスのクリック処理が完了しました")
            return True
            
        except Exception as e:
            logger.error(f"「全てチェック」チェックボックスのクリック処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("select_all_error.png")
            return False
    
    def click_show_more_repeatedly(self, max_attempts=20, interval=5):
        """
        「もっと見る」ボタンを繰り返しクリックして、すべての求職者を表示する
        
        Args:
            max_attempts (int): 最大試行回数
            interval (int): クリック間隔（秒）
            
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== 「もっと見る」ボタンの繰り返しクリック処理を開始します ===")
            
            # 「もっと見る」ボタンが見つからなくなるまで繰り返しクリック
            attempt = 0
            while attempt < max_attempts:
                attempt += 1
                logger.info(f"「もっと見る」ボタンのクリック試行: {attempt}/{max_attempts}")
                
                # スクリーンショットを取得
                self.browser.save_screenshot(f"show_more_attempt_{attempt}.png")
                
                # 「もっと見る」ボタンを探す（クラス名で検索）
                try:
                    # まずセレクタ情報を使用
                    show_more_button = None
                    try:
                        show_more_button = self.browser.get_element('candidates_list', 'show_more_button')
                    except:
                        pass
                    
                    # セレクタ情報で見つからない場合、クラス名で検索
                    if not show_more_button:
                        logger.info("クラス名で「もっと見る」ボタンを探索します")
                        show_more_button = WebDriverWait(self.browser.driver, 5).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.list-view-show-more-button"))
                        )
                    
                    # ボタンが見つからない場合、テキスト内容で検索
                    if not show_more_button:
                        logger.info("テキスト内容で「もっと見る」ボタンを探索します")
                        buttons = self.browser.driver.find_elements(By.TAG_NAME, "button")
                        for button in buttons:
                            if "もっと見る" in button.text:
                                logger.info(f"「もっと見る」テキストを含むボタンを発見しました: {button.text}")
                                show_more_button = button
                                break
                    
                    # ボタンが見つかった場合、クリック
                    if show_more_button:
                        # 要素が画面内に表示されるようにスクロール
                        self.browser.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", show_more_button)
                        time.sleep(1)  # スクロール完了を待機
                        
                        # クリック実行
                        show_more_button.click()
                        logger.info(f"✓ 「もっと見る」ボタンをクリックしました（{attempt}回目）")
                        
                        # 次のデータ読み込みを待機
                        time.sleep(interval)
                    else:
                        logger.info("「もっと見る」ボタンが見つかりませんでした。すべてのデータが表示されたと思われます。")
                        break
                        
                except TimeoutException:
                    logger.info("「もっと見る」ボタンが見つかりませんでした。すべてのデータが表示されたと思われます。")
                    break
                except Exception as e:
                    logger.warning(f"{attempt}回目の「もっと見る」ボタンクリック中にエラーが発生しましたが、処理を継続します: {str(e)}")
                    # エラーが発生しても処理を継続
                    time.sleep(interval)
            
            # 「もっと見る」ボタンがなくなった後、データグリッドコンテナを一番下までスクロール
            logger.info("データグリッドコンテナを一番下までスクロールします")
            
            # データグリッドコンテナを探す（複数のセレクタを試す）
            data_grid_container = None
            grid_selectors = [
                "#dataGridContainer",
                ".data-grid-container",
                "#recordListView div[role='grid']",
                "#recordListView .jss37",
                "#recordListView .MuiTable-root",
                "#recordListView table",
                "div[role='grid']",
                ".grid-container",
                ".table-container"
            ]
            
            for selector in grid_selectors:
                try:
                    elements = self.browser.driver.find_elements(By.CSS_SELECTOR, selector)
                    if elements:
                        data_grid_container = elements[0]
                        logger.info(f"データグリッドコンテナを発見しました: {selector}")
                        break
                except Exception:
                    continue
            
            if data_grid_container:
                try:
                    # コンテナの高さ情報を取得
                    container_height = self.browser.driver.execute_script("return arguments[0].scrollHeight", data_grid_container)
                    logger.info(f"データグリッドコンテナの高さ: {container_height}px")
                    
                    # 段階的にスクロールして、すべてのデータを確実に読み込む
                    scroll_steps = 5  # スクロールのステップ数
                    for step in range(1, scroll_steps + 1):
                        scroll_position = int(container_height * step / scroll_steps)
                        self.browser.driver.execute_script(f"arguments[0].scrollTop = {scroll_position}", data_grid_container)
                        logger.info(f"データグリッドコンテナをスクロール: {scroll_position}px ({step}/{scroll_steps})")
                        time.sleep(1)  # 各ステップ後に少し待機
                    
                    # 最後に一番下までスクロール
                    self.browser.driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", data_grid_container)
                    logger.info("データグリッドコンテナを一番下までスクロールしました")
                    
                    # スクロール完了を待機
                    time.sleep(3)
                    self.browser.save_screenshot("after_grid_scroll_bottom.png")
                    
                except Exception as e:
                    logger.warning(f"データグリッドコンテナのスクロール中にエラーが発生しました: {str(e)}")
                    self._scroll_page_fallback()
            else:
                logger.warning("データグリッドコンテナが見つかりませんでした")
                self._scroll_page_fallback()
            
            # 最終的な画面のスクリーンショットを取得
            self.browser.save_screenshot("after_show_more_all.png")
            
            logger.info(f"✅ 「もっと見る」ボタンの繰り返しクリック処理が完了しました（{attempt}回実行）")
            return True
            
        except Exception as e:
            logger.error(f"「もっと見る」ボタンの繰り返しクリック処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("show_more_error.png")
            return False
    
    def _scroll_page_fallback(self):
        """
        ページ全体をスクロールするフォールバック処理
        
        データグリッドコンテナが見つからない場合や、
        スクロールに失敗した場合に使用する代替手段
        """
        try:
            logger.info("代替手段: ページ全体を一番下までスクロールします")
            
            # ページの高さを取得
            page_height = self.browser.driver.execute_script("return document.body.scrollHeight")
            logger.info(f"ページの高さ: {page_height}px")
            
            # 段階的にスクロール
            scroll_steps = 5
            for step in range(1, scroll_steps + 1):
                scroll_position = int(page_height * step / scroll_steps)
                self.browser.driver.execute_script(f"window.scrollTo(0, {scroll_position});")
                logger.info(f"ページを段階的にスクロール: {scroll_position}px ({step}/{scroll_steps})")
                time.sleep(1)
            
            # 最後に一番下までスクロール
            self.browser.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            self.browser.save_screenshot("after_page_scroll_bottom.png")
            
            # 一番上に戻ってから再度一番下までスクロール
            self.browser.driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(1)
            self.browser.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            
        except Exception as scroll_e:
            logger.warning(f"ページ全体のスクロール中にエラーが発生しました: {str(scroll_e)}")
    
    def export_candidates_data(self) -> bool:
        """
        候補者データをエクスポートする
        
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== 候補者データのエクスポート処理を開始します ===")
            
            # アクションリストボタンをクリック
            if not self.browser.click_element('candidates_list', 'action_button'):
                logger.error("アクションリストボタンのクリックに失敗しました")
                
                # 直接CSSセレクタを使用して再試行
                try:
                    logger.info("直接CSSセレクタを使用してアクションボタンを探索します")
                    action_button_selector = "#recordListView > div.jss37 > div:nth-child(2) > div > button > div"
                    action_button_element = WebDriverWait(self.browser.driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, action_button_selector))
                    )
                    action_button_element.click()
                    logger.info("✓ 直接CSSセレクタを使用してアクションボタンをクリックしました")
                except Exception as css_e:
                    logger.error(f"直接CSSセレクタを使用したクリックにも失敗しました: {str(css_e)}")
                    return False
            else:
                logger.info("✓ アクションリストボタンをクリックしました")
            
            # エクスポートボタンをクリック
            if not self.browser.click_element('candidates_list', 'export_button'):
                logger.error("エクスポートボタンのクリックに失敗しました")
                
                # 直接CSSセレクタを使用して再試行
                try:
                    logger.info("直接CSSセレクタを使用してエクスポートボタンを探索します")
                    export_button_selector = "#pageResume > div:nth-child(25) > div > ul > li.jss175.linkExport"
                    export_button_element = WebDriverWait(self.browser.driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, export_button_selector))
                    )
                    export_button_element.click()
                    logger.info("✓ 直接CSSセレクタを使用してエクスポートボタンをクリックしました")
                except Exception as css_e:
                    logger.warning(f"直接CSSセレクタを使用したクリックに失敗しました: {str(css_e)}")
                    
                    # XPathを使用して再試行
                    try:
                        logger.info("XPathを使用してエクスポートボタンを探索します")
                        export_button_xpath = "//*[@id='pageResume']/div[13]/div/ul/li[2]"
                        export_button_element = WebDriverWait(self.browser.driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, export_button_xpath))
                        )
                        export_button_element.click()
                        logger.info("✓ XPathを使用してエクスポートボタンをクリックしました")
                    except Exception as xpath_e:
                        logger.error(f"XPathを使用したクリックにも失敗しました: {str(xpath_e)}")
                        
                        # テキスト内容で検索する最終手段
                        try:
                            logger.info("テキスト内容でエクスポートボタンを探索します")
                            elements = self.browser.driver.find_elements(By.TAG_NAME, "li")
                            for element in elements:
                                if "エクスポート" in element.text and "linkExport" in element.get_attribute("class"):
                                    logger.info(f"「エクスポート」テキストを含む要素を発見しました: {element.text}")
                                    element.click()
                                    logger.info("✓ テキスト内容でエクスポートボタンをクリックしました")
                                    break
                            else:
                                logger.error("「エクスポート」テキストを含む要素が見つかりませんでした")
                                return False
                        except Exception as text_e:
                            logger.error(f"テキスト内容での探索にも失敗しました: {str(text_e)}")
                            return False
            else:
                logger.info("✓ エクスポートボタンをクリックしました")
            
            # エクスポートモーダルで「候補者リスト（全rawデータ）」を選択
            if not self.browser.click_element('export_dialog', 'all_raw_data'):
                logger.error("「候補者リスト（全rawデータ）」オプションの選択に失敗しました")
                return False
            logger.info("✓ 「候補者リスト（全rawデータ）」オプションを選択しました")
            
            # 「次へ」ボタンを3回クリック
            for i in range(3):
                button_name = f'next_button_{i+1}'
                # セレクタによる検索はスキップし、直接テキストで探す
                logger.info(f"テキストで{i+1}回目の「次へ」ボタンを探索します")
                buttons = self.browser.driver.find_elements(By.TAG_NAME, "span")
                next_button_found = False
                
                for button in buttons:
                    if "次へ" in button.text:
                        logger.info(f"「次へ」テキストを含むボタンを発見しました: {button.text}")
                        button.click()
                        logger.info(f"✓ テキストで{i+1}回目の「次へ」ボタンをクリックしました")
                        next_button_found = True
                        break
                
                if not next_button_found:
                    # 最後のボタンは「実行」の可能性がある
                    if i == 2:
                        logger.info("最後のボタンは「実行」の可能性があるため、「実行」ボタンを探索します")
                        for button in buttons:
                            if "実行" in button.text:
                                logger.info(f"「実行」テキストを含むボタンを発見しました: {button.text}")
                                button.click()
                                logger.info("✓ テキストで「実行」ボタンをクリックしました")
                                next_button_found = True
                                break
                    
                    if not next_button_found:
                        logger.error(f"{i+1}回目の「次へ」または「実行」ボタンが見つかりませんでした")
                        return False
                
                time.sleep(1)
            
            # 「実行」ボタンをクリック（3回目の「次へ」ボタンの後に実行ボタンをクリック）
            logger.info("テキストで「実行」ボタンを探索します")
            buttons = self.browser.driver.find_elements(By.TAG_NAME, "span")
            execute_button_found = False
            
            for button in buttons:
                if "実行" in button.text:
                    logger.info(f"「実行」テキストを含むボタンを発見しました: {button.text}")
                    button.click()
                    logger.info("✓ テキストで「実行」ボタンをクリックしました")
                    execute_button_found = True
                    break
            
            if not execute_button_found:
                # 「設定を保存」ボタンが表示されている可能性がある
                for button in buttons:
                    if "設定を保存" in button.text:
                        logger.info(f"「設定を保存」テキストを含むボタンを発見しました: {button.text}")
                        logger.warning("「設定を保存」ボタンではなく「実行」ボタンを探しています")
                        continue
                    
                    # ボタンのテキストをログに出力して確認
                    if button.text.strip():
                        logger.info(f"ボタンテキスト: '{button.text}'")
                
                # 親要素を探索してボタンを見つける
                try:
                    logger.info("ダイアログのボタンペインを探索します")
                    button_panes = self.browser.driver.find_elements(By.CSS_SELECTOR, ".ui-dialog-buttonpane")
                    if button_panes:
                        logger.info(f"{len(button_panes)}個のボタンペインを発見しました")
                        for pane in button_panes:
                            buttons = pane.find_elements(By.TAG_NAME, "button")
                            logger.info(f"ペイン内に{len(buttons)}個のボタンを発見しました")
                            
                            # 最後のボタンが「実行」ボタンの可能性が高い
                            if buttons:
                                last_button = buttons[-1]
                                logger.info(f"最後のボタンをクリックします: '{last_button.text}'")
                                last_button.click()
                                logger.info("✓ 最後のボタンをクリックしました")
                                execute_button_found = True
                                break
                except Exception as e:
                    logger.warning(f"ダイアログのボタンペイン探索中にエラーが発生しました: {str(e)}")
            
            if not execute_button_found:
                logger.error("「実行」ボタンが見つかりませんでした")
                return False
            
            # OKボタンをクリック
            time.sleep(30)  # エクスポート処理の完了を待つ
            
            # テキストで「OK」ボタンを探す
            logger.info("テキストで「OK」ボタンを探索します")
            ok_button_found = False
            
            # まずspanタグ内のテキストで探す
            buttons = self.browser.driver.find_elements(By.TAG_NAME, "span")
            for button in buttons:
                if button.text.strip().upper() == "OK":
                    logger.info(f"「OK」テキストを含むspanを発見しました: {button.text}")
                    button.click()
                    logger.info("✓ テキストで「OK」ボタンをクリックしました")
                    ok_button_found = True
                    break
            
            # spanで見つからない場合はbuttonタグで探す
            if not ok_button_found:
                logger.info("buttonタグで「OK」ボタンを探索します")
                buttons = self.browser.driver.find_elements(By.TAG_NAME, "button")
                for button in buttons:
                    if button.text.strip().upper() == "OK":
                        logger.info(f"「OK」テキストを含むbuttonを発見しました: {button.text}")
                        button.click()
                        logger.info("✓ buttonタグで「OK」ボタンをクリックしました")
                        ok_button_found = True
                        break
            
            # ダイアログのボタンを探す
            if not ok_button_found:
                try:
                    logger.info("OKダイアログのボタンを探索します")
                    # ダイアログのボタンペインを探す
                    dialog_buttons = self.browser.driver.find_elements(By.CSS_SELECTOR, ".ui-dialog-buttonpane button")
                    if dialog_buttons:
                        logger.info(f"{len(dialog_buttons)}個のダイアログボタンを発見しました")
                        # ボタンのテキストをログに出力
                        for i, btn in enumerate(dialog_buttons):
                            logger.info(f"ダイアログボタン {i+1}: テキスト='{btn.text}'")
                        
                        # 通常OKボタンは1つだけ
                        dialog_buttons[0].click()
                        logger.info("✓ ダイアログボタンをクリックしました")
                        ok_button_found = True
                    else:
                        # 一般的なダイアログボタンを探す
                        logger.info("一般的なダイアログボタンを探索します")
                        dialog_selectors = [
                            ".ui-dialog button",
                            ".modal-dialog button",
                            ".dialog button",
                            "div[role='dialog'] button"
                        ]
                        
                        for selector in dialog_selectors:
                            buttons = self.browser.driver.find_elements(By.CSS_SELECTOR, selector)
                            if buttons:
                                logger.info(f"{selector} で {len(buttons)}個のボタンを発見しました")
                                # 最初のボタンをクリック
                                buttons[0].click()
                                logger.info(f"✓ {selector} の最初のボタンをクリックしました")
                                ok_button_found = True
                                break
                except Exception as e:
                    logger.warning(f"OKダイアログのボタン探索中にエラーが発生しました: {str(e)}")
            
            if not ok_button_found:
                logger.error("「OK」ボタンが見つかりませんでした")
                return False
            
            # CSVファイルをダウンロード
            csv_file_path = self._download_exported_csv()
            if not csv_file_path:
                logger.error("CSVファイルのダウンロードに失敗しました")
                return False
            logger.info(f"✓ CSVファイルをダウンロードしました: {csv_file_path}")
            
            # スプレッドシートに転記
            if self.import_csv_to_spreadsheet(csv_file_path):
                logger.info("✓ CSVデータをスプレッドシートに転記しました")
            else:
                logger.error("CSVデータのスプレッドシートへの転記に失敗しました")
                return False
            
            logger.info("✅ 候補者データのエクスポート処理が完了しました")
            return True
            
        except Exception as e:
            logger.error(f"候補者データのエクスポート処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("export_candidates_error.png")
            return False
    
    def _download_exported_csv(self, max_retries: int = 3, retry_interval: int = 30) -> Optional[str]:
        """
        エクスポート結果リストを開き、CSVファイルをダウンロードする
        
        Args:
            max_retries (int): 最大リトライ回数
            retry_interval (int): リトライ間隔（秒）
            
        Returns:
            Optional[str]: ダウンロードしたCSVファイルのパス。失敗した場合はNone。
        """
        # エクスポート結果リストを開く
        for attempt in range(max_retries):
            try:
                logger.info(f"エクスポート結果リストを開く（試行 {attempt + 1}/{max_retries}）")
                
                # テキスト要素で「エクスポートの結果一覧を開く」を探す
                logger.info("テキストで「エクスポートの結果一覧を開く」ボタンを探索します")
                export_result_button_found = False
                
                # まずliタグ内のテキストで探す
                elements = self.browser.driver.find_elements(By.TAG_NAME, "li")
                for element in elements:
                    if "エクスポートの結果一覧を開く" in element.text:
                        logger.info(f"「エクスポートの結果一覧を開く」テキストを含む要素を発見しました: {element.text}")
                        element.click()
                        logger.info("✓ テキストで「エクスポートの結果一覧を開く」ボタンをクリックしました")
                        export_result_button_found = True
                        break
                
                # テキストで見つからない場合はタイトル属性で探す
                if not export_result_button_found:
                    logger.info("タイトル属性で「エクスポートの結果一覧を開く」ボタンを探索します")
                    for element in elements:
                        title = element.get_attribute("title")
                        if title and "エクスポートの結果一覧を開く" in title:
                            logger.info(f"「エクスポートの結果一覧を開く」タイトルを持つ要素を発見しました")
                            element.click()
                            logger.info("✓ タイトル属性で「エクスポートの結果一覧を開く」ボタンをクリックしました")
                            export_result_button_found = True
                            break
                
                # タイトル属性でも見つからない場合はクラス名で探す
                if not export_result_button_found:
                    logger.info("クラス名で「エクスポートの結果一覧を開く」ボタンを探索します")
                    elements = self.browser.driver.find_elements(By.CLASS_NAME, "p-notificationbar-item-export")
                    if elements:
                        logger.info("クラス名で「エクスポートの結果一覧を開く」ボタンを発見しました")
                        elements[0].click()
                        logger.info("✓ クラス名で「エクスポートの結果一覧を開く」ボタンをクリックしました")
                        export_result_button_found = True
                
                # 最後にセレクタを使用して探す
                if not export_result_button_found:
                    logger.info("セレクタで「エクスポートの結果一覧を開く」ボタンを探索します")
                    if self.browser.click_element('export_result', 'result_list_button'):
                        logger.info("✓ セレクタで「エクスポートの結果一覧を開く」ボタンをクリックしました")
                        export_result_button_found = True
                
                if not export_result_button_found:
                    raise Exception("「エクスポートの結果一覧を開く」ボタンが見つかりませんでした")
                
                time.sleep(3)  # リストが表示されるまで待機
                break
            except Exception as e:
                logger.warning(f"エクスポート結果リストを開く際にエラーが発生しました: {e}")
                if attempt < max_retries - 1:
                    logger.info(f"{retry_interval}秒後にリトライします")
                    time.sleep(retry_interval)
                else:
                    logger.error("エクスポート結果リストを開くことができませんでした")
                    return None
        
        # CSVダウンロードリンクをクリック
        for attempt in range(max_retries):
            try:
                logger.info(f"CSVダウンロードリンクをクリック（試行 {attempt + 1}/{max_retries}）")
                
                # セレクタを使用してCSVダウンロードリンクをクリック
                if not self.browser.click_element('export_result', 'csv_download_link'):
                    logger.warning("セレクタでCSVダウンロードリンクを見つけられませんでした")
                    
                    # テキストでリンクを探す
                    try:
                        logger.info("テキストでCSVダウンロードリンクを探索します")
                        links = self.browser.driver.find_elements(By.TAG_NAME, "a")
                        for link in links:
                            link_text = link.text.strip()
                            if "エクスポートしたデーターを取得する" in link_text:
                                logger.info(f"「エクスポートしたデーターを取得する」テキストを含むリンクを発見しました: {link_text}")
                                link.click()
                                logger.info("✓ テキストでCSVダウンロードリンクをクリックしました")
                                break
                            elif "CSV" in link_text:
                                logger.info(f"「CSV」テキストを含むリンクを発見しました: {link_text}")
                                link.click()
                                logger.info("✓ テキストでCSVダウンロードリンクをクリックしました")
                                break
                        else:
                            # href属性で探す
                            links = self.browser.driver.find_elements(By.CSS_SELECTOR, "a[href*='download']")
                            if links:
                                logger.info("href属性でCSVダウンロードリンクを発見しました")
                                links[0].click()
                                logger.info("✓ href属性でCSVダウンロードリンクをクリックしました")
                            else:
                                raise Exception("CSVダウンロードリンクが見つかりませんでした")
                    except Exception as e:
                        raise e
                
                # ダウンロードが完了するまで待機
                logger.info("CSVファイルのダウンロードを待機中...")
                time.sleep(10)  # ダウンロード開始のための初期待機
                
                # ダウンロードディレクトリで新しいCSVファイルを待機
                csv_path = find_latest_csv_in_downloads()
                if csv_path:
                    logger.info(f"CSVファイルのダウンロードが完了しました: {csv_path}")
                    return csv_path
                else:
                    logger.warning("CSVファイルのダウンロードを検出できませんでした")
                    if attempt < max_retries - 1:
                        logger.info(f"{retry_interval}秒後にリトライします")
                        time.sleep(retry_interval)
                    else:
                        # 最終試行でも失敗した場合は、最新のCSVファイルを探す
                        latest_csv = find_latest_csv_in_downloads()
                        if latest_csv:
                            logger.info(f"最新のCSVファイルを使用します: {latest_csv}")
                            return latest_csv
                        logger.error("CSVファイルをダウンロードできませんでした")
                        return None
            except Exception as e:
                logger.warning(f"CSVダウンロード中にエラーが発生しました: {e}")
                if attempt < max_retries - 1:
                    logger.info(f"{retry_interval}秒後にリトライします")
                    time.sleep(retry_interval)
                else:
                    logger.error("CSVファイルをダウンロードできませんでした")
                    return None
    
    def import_csv_to_spreadsheet(self, csv_file_path: str) -> bool:
        """
        CSVファイルのデータをスプレッドシートに転記する
        
        Args:
            csv_file_path (str): CSVファイルのパス
            
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== CSVデータのスプレッドシートへの転記処理を開始します ===")
            
            # CSVファイルの存在確認
            if not os.path.exists(csv_file_path):
                logger.error(f"CSVファイルが見つかりません: {csv_file_path}")
                return False
            
            # CSVファイルのサイズ確認
            file_size = os.path.getsize(csv_file_path)
            if file_size == 0:
                logger.error(f"CSVファイルが空です: {csv_file_path}")
                return False
            
            logger.info(f"CSVファイルのサイズ: {file_size} バイト")
            
            # CSVファイルの内容を確認（デバッグ用）
            try:
                # 複数の文字コードを試す
                encodings = ['utf-8-sig', 'utf-8', 'shift-jis', 'cp932']
                csv_content = None
                
                for encoding in encodings:
                    try:
                        with open(csv_file_path, 'r', encoding=encoding) as f:
                            csv_content = f.read(1024)  # 最初の1KBだけ読み込む
                            logger.info(f"CSVファイルの文字コード: {encoding}")
                            logger.info(f"CSVファイルの先頭部分: {csv_content[:100]}...")
                            break
                    except UnicodeDecodeError:
                        continue
                
                if csv_content is None:
                    logger.warning("CSVファイルの文字コードを特定できませんでした")
            except Exception as e:
                logger.warning(f"CSVファイルの内容確認中にエラーが発生: {str(e)}")
            
            # スプレッドシートマネージャーの初期化
            try:
                spreadsheet_manager = SpreadsheetManager()
                logger.info("SpreadsheetManagerを初期化しました")
            except Exception as e:
                logger.error(f"SpreadsheetManagerの初期化に失敗しました: {str(e)}")
                import traceback
                logger.error(traceback.format_exc())
                return False
            
            # スプレッドシートを開く
            try:
                spreadsheet_manager.open_spreadsheet()
                logger.info("スプレッドシートを開きました")
            except Exception as e:
                logger.error(f"スプレッドシートを開くことができませんでした: {str(e)}")
                import traceback
                logger.error(traceback.format_exc())
                return False
            
            # USERSALL シートのデータをクリア
            try:
                logger.info("USERSALLシートのデータをクリアします")
                spreadsheet_manager.clear_worksheet('users_all')
                logger.info("USERSALLシートのデータをクリアしました")
            except Exception as e:
                logger.error(f"USERSALLシートのデータクリアに失敗しました: {str(e)}")
                import traceback
                logger.error(traceback.format_exc())
                # クリアに失敗しても続行
            
            # CSVデータをスプレッドシートに転記
            try:
                logger.info(f"CSVデータをUSERSALLシートに転記します: {csv_file_path}")
                spreadsheet_manager.import_csv_to_sheet(csv_file_path, 'users_all')
                logger.info("CSVデータをUSERSALLシートに転記しました")
            except Exception as e:
                logger.error(f"CSVデータの転記に失敗しました: {str(e)}")
                import traceback
                logger.error(traceback.format_exc())
                return False
            
            # CSVファイルを削除
            try:
                logger.info(f"スプレッドシートへの転記が完了したため、CSVファイルを削除します: {csv_file_path}")
                os.remove(csv_file_path)
                logger.info(f"✓ CSVファイルを削除しました: {csv_file_path}")
            except Exception as e:
                logger.warning(f"CSVファイルの削除中にエラーが発生しました: {str(e)}")
                # ファイル削除に失敗しても処理は続行
            
            logger.info("✅ CSVデータのスプレッドシートへの転記処理が完了しました")
            return True
            
        except Exception as e:
            logger.error(f"CSVデータのスプレッドシートへの転記処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return False
    
    def execute_operations_flow(self):
        """
        PORTERSシステムでの一連の業務操作を実行する
        
        以下の処理を順番に実行します：
        1. 「その他業務」ボタンをクリックして新しいウィンドウを開く
        2. メニュー項目5をクリック
        3. 「すべての求職者」リンクをクリック
        4. 求職者一覧画面で「全てチェック」チェックボックスをクリック
        5. 「もっと見る」ボタンを繰り返しクリックして、すべての求職者を表示
        6. アクション一覧ボタンをクリックしてエクスポート処理を実行
        
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== 業務操作フローを開始します ===")
            
            # 「その他業務」ボタンをクリック
            if not self.click_other_operations_button():
                logger.error("「その他業務」ボタンのクリック処理に失敗しました")
                return False
            
            # メニュー項目5をクリック
            if not self.click_menu_item_5():
                logger.error("メニュー項目5のクリック処理に失敗しました")
                return False
            
            # 「すべての求職者」リンクをクリック
            if not self.click_all_candidates():
                logger.error("「すべての求職者」リンクのクリック処理に失敗しました")
                return False
            
            # 求職者一覧画面で「全てチェック」チェックボックスをクリック
            if not self.select_all_candidates():
                logger.error("「全てチェック」チェックボックスのクリック処理に失敗しました")
                return False
            
            # 「もっと見る」ボタンを繰り返しクリックして、すべての求職者を表示
            if not self.click_show_more_repeatedly():
                logger.error("「もっと見る」ボタンの繰り返しクリック処理に失敗しました")
                return False
            
            # 求職者データのエクスポート処理を実行
            if not self.export_candidates_data():
                logger.error("求職者データのエクスポート処理に失敗しました")
                return False
            
            logger.info("✅ 業務操作フローが正常に完了しました")
            return True
            
        except Exception as e:
            logger.error(f"業務操作フロー中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return False
    
    def logout(self):
        """
        PORTERSシステムからログアウトする
        
        ユーザーメニューを開いてログアウトボタンをクリックします。
        
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== ログアウト処理を開始します ===")
            logger.info(f"ログアウト前のURL: {self.browser.driver.current_url}")
            
            # ユーザーメニューボタンをクリック
            logger.info("ユーザーメニューボタンを探索します")
            user_menu_clicked = False
            
            # まずセレクタで試す
            try:
                user_menu_selector = "#nav2-inner > div > ul > li.original-class-user > a > span"
                user_menu_element = WebDriverWait(self.browser.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, user_menu_selector))
                )
                user_menu_element.click()
                logger.info("✓ セレクタでユーザーメニューボタンをクリックしました")
                user_menu_clicked = True
            except Exception as e:
                logger.warning(f"セレクタでのユーザーメニュークリックに失敗しました: {str(e)}")
            
            # セレクタで失敗した場合はテキストで探す
            if not user_menu_clicked:
                try:
                    logger.info("テキストでユーザーメニューボタンを探索します")
                    # 「川島」を含む要素を探す
                    elements = self.browser.driver.find_elements(By.XPATH, "//*[contains(text(), '川島')]")
                    for element in elements:
                        logger.info(f"「川島」テキストを含む要素を発見しました: {element.text}")
                        try:
                            element.click()
                            logger.info("✓ 「川島」テキストでユーザーメニューボタンをクリックしました")
                            user_menu_clicked = True
                            break
                        except Exception as click_e:
                            logger.warning(f"要素のクリックに失敗しました: {str(click_e)}")
                            # 親要素をクリックしてみる
                            try:
                                parent = self.browser.driver.execute_script("return arguments[0].parentNode;", element)
                                parent.click()
                                logger.info("✓ 「川島」テキストの親要素をクリックしました")
                                user_menu_clicked = True
                                break
                            except Exception as parent_e:
                                logger.warning(f"親要素のクリックにも失敗しました: {str(parent_e)}")
                                continue
                except Exception as e:
                    logger.warning(f"「川島」テキストでのユーザーメニュークリックに失敗しました: {str(e)}")
            
            # 一般的なユーザーテキストで探す
            if not user_menu_clicked:
                try:
                    logger.info("一般的なユーザーテキストでユーザーメニューボタンを探索します")
                    spans = self.browser.driver.find_elements(By.TAG_NAME, "span")
                    for span in spans:
                        if "ユーザー" in span.text or "User" in span.text or "user" in span.text:
                            logger.info(f"ユーザーメニューと思われる要素を発見しました: {span.text}")
                            span.click()
                            logger.info("✓ テキストでユーザーメニューボタンをクリックしました")
                            user_menu_clicked = True
                            break
                except Exception as e:
                    logger.warning(f"テキストでのユーザーメニュークリックに失敗しました: {str(e)}")
            
            # ユーザーメニューが見つからない場合は親要素を探す
            if not user_menu_clicked:
                try:
                    logger.info("ユーザーメニューの親要素を探索します")
                    user_menu_parent_selector = "#nav2-inner > div > ul > li.original-class-user > a"
                    user_menu_parent = WebDriverWait(self.browser.driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, user_menu_parent_selector))
                    )
                    user_menu_parent.click()
                    logger.info("✓ 親要素からユーザーメニューボタンをクリックしました")
                    user_menu_clicked = True
                except Exception as e:
                    logger.warning(f"親要素からのユーザーメニュークリックに失敗しました: {str(e)}")
            
            # 直接ログアウトリンクを探す
            if not user_menu_clicked:
                try:
                    logger.info("ログアウトボタンを直接探索: a[href*='logout']")
                    logout_link = WebDriverWait(self.browser.driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href*='logout']"))
                    )
                    logout_link.click()
                    logger.info("✓ ログアウトリンクを直接クリックしました")
                    # ログアウト完了を待機
                    time.sleep(5)
                    self.browser.save_screenshot("after_direct_logout.png")
                    return self._verify_logout()
                except Exception as e:
                    logger.warning(f"通常のクリックでログアウトに失敗しました: {str(e)}")
                    
                    # JavaScriptでログアウトを試みる
                    try:
                        logger.info("JavaScriptでログアウトを試みます")
                        self.browser.driver.execute_script("document.querySelector('a[href*=\"logout\"]').click();")
                        logger.info("✓ JavaScriptでログアウトリンクをクリックしました")
                        # ログアウト完了を待機
                        time.sleep(5)
                        self.browser.save_screenshot("after_js_logout.png")
                        return self._verify_logout()
                    except Exception as js_e:
                        logger.error(f"JavaScriptを使用したログアウトにも失敗しました: {str(js_e)}")
            
            if not user_menu_clicked:
                logger.error("ユーザーメニューボタンが見つかりませんでした")
                return False
            
            # メニューが表示されるまで待機
            time.sleep(2)
            self.browser.save_screenshot("after_user_menu_click.png")
            
            # ログアウトボタンをクリック
            logger.info("ログアウトボタンを探索します")
            logout_clicked = False
            
            # まずセレクタで試す
            try:
                logout_selector = "#porters-contextmenu-column_1 > ul > li:nth-child(6)"
                logout_element = WebDriverWait(self.browser.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, logout_selector))
                )
                logout_element.click()
                logger.info("✓ セレクタでログアウトボタンをクリックしました")
                logout_clicked = True
            except Exception as e:
                logger.warning(f"セレクタでのログアウトボタンクリックに失敗しました: {str(e)}")
            
            # セレクタで失敗した場合はテキストで探す
            if not logout_clicked:
                try:
                    logger.info("テキストでログアウトボタンを探索します")
                    elements = self.browser.driver.find_elements(By.TAG_NAME, "li")
                    for element in elements:
                        if "ログアウト" in element.text:
                            logger.info(f"「ログアウト」テキストを含む要素を発見しました: {element.text}")
                            element.click()
                            logger.info("✓ テキストでログアウトボタンをクリックしました")
                            logout_clicked = True
                            break
                except Exception as e:
                    logger.warning(f"テキストでのログアウトボタンクリックに失敗しました: {str(e)}")
            
            # href属性で探す
            if not logout_clicked:
                try:
                    logger.info("href属性でログアウトリンクを探索します")
                    logout_links = self.browser.driver.find_elements(By.CSS_SELECTOR, "a[href*='logout']")
                    if logout_links:
                        logout_links[0].click()
                        logger.info("✓ href属性でログアウトリンクをクリックしました")
                        logout_clicked = True
                except Exception as e:
                    logger.warning(f"href属性でのログアウトリンククリックに失敗しました: {str(e)}")
            
            if not logout_clicked:
                logger.error("ログアウトボタンが見つかりませんでした")
                return False
            
            # ログアウト完了を待機
            logger.info("ログアウト完了を待機中...")
            time.sleep(5)
            self.browser.save_screenshot("after_logout.png")
            
            return self._verify_logout()
            
        except Exception as e:
            logger.error(f"ログアウト処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("logout_error.png")
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
    
    def click_selection_process_menu(self):
        """
        メニューバーの「選考プロセス」をクリックする
        
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== 「選考プロセス」メニューのクリック処理を開始します ===")
            
            # 「選考プロセス」メニューをクリック
            selection_process_clicked = False
            
            # まず指定されたセレクタで試す
            try:
                logger.info("セレクタで「選考プロセス」メニューを探索します")
                selection_process_selector = "#main-menu-id-6 > a"
                selection_process_element = WebDriverWait(self.browser.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, selection_process_selector))
                )
                selection_process_element.click()
                logger.info("✓ セレクタで「選考プロセス」メニューをクリックしました")
                selection_process_clicked = True
            except Exception as e:
                logger.warning(f"セレクタでの「選考プロセス」メニュークリックに失敗しました: {str(e)}")
            
            # セレクタで失敗した場合はテキストで探す
            if not selection_process_clicked:
                try:
                    logger.info("テキストで「選考プロセス」メニューを探索します")
                    links = self.browser.driver.find_elements(By.TAG_NAME, "a")
                    for link in links:
                        if "選考プロセス" in link.text:
                            logger.info(f"「選考プロセス」テキストを含むリンクを発見しました: {link.text}")
                            link.click()
                            logger.info("✓ テキストで「選考プロセス」メニューをクリックしました")
                            selection_process_clicked = True
                            break
                except Exception as e:
                    logger.warning(f"テキストでの「選考プロセス」メニュークリックに失敗しました: {str(e)}")
            
            if not selection_process_clicked:
                logger.error("「選考プロセス」メニューが見つかりませんでした")
                return False
            
            # 処理待機
            time.sleep(3)
            self.browser.save_screenshot("after_selection_process_menu_click.png")
            
            logger.info("✅ 「選考プロセス」メニューのクリック処理が完了しました")
            return True
            
        except Exception as e:
            logger.error(f"「選考プロセス」メニューのクリック処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("selection_process_menu_error.png")
            return False
    
    def click_all_selection_processes(self):
        """
        「すべての選考プロセス」リンクをクリックする
        
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== 「すべての選考プロセス」リンクのクリック処理を開始します ===")
            
            # 「すべての選考プロセス」リンクをクリック
            all_processes_clicked = False
            
            # まず指定されたセレクタで試す
            try:
                logger.info("セレクタで「すべての選考プロセス」リンクを探索します")
                all_processes_selector = "#ui-id-196 > li:nth-child(12) > a"
                all_processes_element = WebDriverWait(self.browser.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, all_processes_selector))
                )
                all_processes_element.click()
                logger.info("✓ セレクタで「すべての選考プロセス」リンクをクリックしました")
                all_processes_clicked = True
            except Exception as e:
                logger.warning(f"セレクタでの「すべての選考プロセス」リンククリックに失敗しました: {str(e)}")
            
            # セレクタで失敗した場合はテキストで探す
            if not all_processes_clicked:
                try:
                    logger.info("テキストで「すべての選考プロセス」リンクを探索します")
                    links = self.browser.driver.find_elements(By.TAG_NAME, "a")
                    for link in links:
                        if "すべての選考プロセス" in link.text:
                            logger.info(f"「すべての選考プロセス」テキストを含むリンクを発見しました: {link.text}")
                            link.click()
                            logger.info("✓ テキストで「すべての選考プロセス」リンクをクリックしました")
                            all_processes_clicked = True
                            break
                except Exception as e:
                    logger.warning(f"テキストでの「すべての選考プロセス」リンククリックに失敗しました: {str(e)}")
            
            if not all_processes_clicked:
                logger.error("「すべての選考プロセス」リンクが見つかりませんでした")
                return False
            
            # 処理待機
            time.sleep(3)
            self.browser.save_screenshot("after_all_selection_processes_click.png")
            
            # ページ内容を確認
            page_html = self.browser.driver.page_source
            page_analysis = self.browser.analyze_page_content(page_html)
            logger.info(f"ページタイトル: {page_analysis['page_title']}")
            
            logger.info("✅ 「すべての選考プロセス」リンクのクリック処理が完了しました")
            return True
            
        except Exception as e:
            logger.error(f"「すべての選考プロセス」リンクのクリック処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("all_selection_processes_error.png")
            return False
    
    def access_selection_processes(self):
        """
        選考プロセス画面にアクセスする
        
        以下の処理を順番に実行します：
        1. メニューバーの「選考プロセス」をクリック
        2. 「すべての選考プロセス」リンクをクリック
        
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== 選考プロセス画面へのアクセス処理を開始します ===")
            
            # 「選考プロセス」メニューをクリック
            if not self.click_selection_process_menu():
                logger.error("「選考プロセス」メニューのクリック処理に失敗しました")
                return False
            
            # 「すべての選考プロセス」リンクをクリック
            if not self.click_all_selection_processes():
                logger.error("「すべての選考プロセス」リンクのクリック処理に失敗しました")
                return False
            
            logger.info("✅ 選考プロセス画面へのアクセス処理が完了しました")
            return True
            
        except Exception as e:
            logger.error(f"選考プロセス画面へのアクセス処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("access_selection_processes_error.png")
            return False 