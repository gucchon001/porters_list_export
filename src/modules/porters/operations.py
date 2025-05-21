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
import re
import importlib.util

from src.utils.logging_config import get_logger
from src.utils.helpers import wait_for_new_csv_in_downloads, move_file_to_data_dir, find_latest_csv_in_downloads
from src.utils.spreadsheet import SpreadsheetManager
from src.utils.environment import EnvironmentUtils as env
from src.modules.spreadsheet_aggregator import SpreadsheetAggregator

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
            current_handles = self.browser.get_window_handles()
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
            new_window_html = self.browser.get_page_source()
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
    
    def click_candidates_menu(self):
        """
        求職者メニューをクリックする
        
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== 求職者メニューのクリック処理を開始します ===")
            
            # メニュー項目5（求職者メニュー）をクリック
            if not self.browser.click_element('porters_menu', 'menu_item_5'):
                logger.error("求職者メニューのクリックに失敗しました")
                return False
            
            # 処理待機
            time.sleep(3)
            self.browser.save_screenshot("after_candidates_menu_click.png")
            
            logger.info("✅ 求職者メニューのクリック処理が完了しました")
            return True
            
        except Exception as e:
            logger.error(f"求職者メニューのクリック処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("candidates_menu_error.png")
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
                    all_candidates_element = self.browser.wait_for_element(
                        By.CSS_SELECTOR, 
                        all_candidates_selector,
                        condition=EC.element_to_be_clickable
                    )
                    
                    if all_candidates_element:
                        self.browser.click_element_direct(all_candidates_element)
                        logger.info("✓ 直接CSSセレクタを使用して「すべての求職者」リンクをクリックしました")
                    else:
                        logger.error("直接CSSセレクタを使用しても「すべての求職者」リンクが見つかりませんでした")
                        
                        # テキストで要素を探す最終手段
                        try:
                            logger.info("テキスト内容で「すべての求職者」リンクを探索します")
                            links = self.browser.find_elements_by_tag("a", "すべての求職者")
                            if links:
                                self.browser.click_element_direct(links[0])
                                logger.info("✓ テキスト内容で「すべての求職者」リンクをクリックしました")
                            else:
                                logger.error("「すべての求職者」テキストを含むリンクが見つかりませんでした")
                                return False
                        except Exception as text_e:
                            logger.error(f"テキスト内容での探索にも失敗しました: {str(text_e)}")
                            return False
                except Exception as css_e:
                    logger.error(f"直接CSSセレクタを使用したクリックにも失敗しました: {str(css_e)}")
                    return False
            
            # 処理待機
            time.sleep(3)
            self.browser.save_screenshot("after_all_candidates_click.png")
            
            # ページ内容を確認
            page_html = self.browser.get_page_source()
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
                    checkbox_element = self.browser.wait_for_element(
                        By.CSS_SELECTOR, 
                        checkbox_selector,
                        condition=EC.element_to_be_clickable,
                        timeout=10
                    )
                    self.browser.click_element_direct(checkbox_element)
                    logger.info("✓ 直接CSSセレクタを使用して「全てチェック」チェックボックスをクリックしました")
                except Exception as css_e:
                    logger.error(f"直接CSSセレクタを使用したクリックにも失敗しました: {str(css_e)}")
                    
                    # より具体的なセレクタを試す
                    try:
                        logger.info("より具体的なセレクタを使用して「全てチェック」チェックボックスを探索します")
                        specific_selector = "#recordListView > div.jss37 > div:nth-child(2) > div > div.jss45 > span > span > input"
                        specific_element = self.browser.wait_for_element(
                            By.CSS_SELECTOR, 
                            specific_selector,
                            condition=EC.element_to_be_clickable,
                            timeout=10
                        )
                        self.browser.click_element_direct(specific_element)
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
                
                # 直接CSSセレクタを使用してアクションボタンを探索します
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
                
                # 「次へ」というテキストを含むspanタグを探す
                buttons = self.browser.find_elements(By.TAG_NAME, "span")
                next_button_found = False
                
                for button in buttons:
                    try:
                        if "次へ" in button.text:
                            logger.info(f"「次へ」テキストを含むボタンを発見しました: {button.text}")
                            button.click()
                            logger.info(f"✓ テキストで{i+1}回目の「次へ」ボタンをクリックしました")
                            next_button_found = True
                            break
                    except:
                        continue
                
                if not next_button_found:
                    # 最後のボタンは「実行」の可能性がある
                    if i == 2:
                        logger.info("最後のボタンは「実行」の可能性があるため、「実行」ボタンを探索します")
                        for button in buttons:
                            try:
                                if "実行" in button.text:
                                    logger.info(f"「実行」テキストを含むボタンを発見しました: {button.text}")
                                    button.click()
                                    logger.info("✓ テキストで「実行」ボタンをクリックしました")
                                    next_button_found = True
                                    break
                            except:
                                continue
                    
                    if not next_button_found:
                        logger.error(f"{i+1}回目の「次へ」または「実行」ボタンが見つかりませんでした")
                        return False
                
                time.sleep(1)
            
            # 「実行」ボタンをクリック（3回目の「次へ」ボタンの後に実行ボタンをクリック）
            logger.info("テキストで「実行」ボタンを探索します")
            buttons = self.browser.driver.find_elements(By.TAG_NAME, "span")
            execute_button_found = False
            
            for button in buttons:
                try:
                    if "実行" in button.text:
                        logger.info(f"「実行」テキストを含むボタンを発見しました: {button.text}")
                        button.click()
                        logger.info("✓ テキストで「実行」ボタンをクリックしました")
                        execute_button_found = True
                        break
                except:
                    continue
            
            if not execute_button_found:
                # 「設定を保存」ボタンが表示されている可能性がある
                for button in buttons:
                    try:
                        if "設定を保存" in button.text:
                            logger.info(f"「設定を保存」テキストを含むボタンを発見しました: {button.text}")
                            logger.warning("「設定を保存」ボタンではなく「実行」ボタンを探しています")
                            continue
                        
                        # ボタンのテキストをログに出力して確認
                        if button.text.strip():
                            logger.info(f"ボタンテキスト: '{button.text}'")
                    except:
                        continue
                
                # 親要素を探索してボタンを見つける
                try:
                    logger.info("ダイアログのボタンペインを探索します")
                    button_panes = self.browser.driver.find_elements(By.CSS_SELECTOR, ".ui-dialog-buttonpane")
                    if button_panes:
                        logger.info(f"{len(button_panes)}個のボタンペインを発見しました")
                        for pane in button_panes:
                            buttons = []
                            try:
                                buttons = pane.find_elements(By.TAG_NAME, "button")
                                logger.info(f"ペイン内に{len(buttons)}個のボタンを発見しました")
                                
                                # 最後のボタンが「実行」ボタンの可能性が高い
                                if buttons:
                                    last_button = buttons[-1]
                                    button_text = last_button.text if hasattr(last_button, 'text') else ""
                                    logger.info(f"最後のボタンをクリックします: '{button_text}'")
                                    last_button.click()
                                    logger.info("✓ 最後のボタンをクリックしました")
                                    execute_button_found = True
                                    break
                            except Exception as btn_e:
                                logger.warning(f"ボタン探索中にエラーが発生しました: {str(btn_e)}")
                                continue
                except Exception as e:
                    logger.warning(f"ダイアログのボタンペイン探索中にエラーが発生しました: {str(e)}")
            
            if not execute_button_found:
                logger.error("「実行」ボタンが見つかりませんでした")
                return False
            
            # OKボタンをクリック
            time.sleep(45)  # エクスポート処理の完了を待つ（30秒から45秒に増加）
            
            # テキストで「OK」ボタンを探す
            logger.info("テキストで「OK」ボタンを探索します")
            ok_button_found = False
            
            # まずspanタグ内のテキストで探す
            buttons = self.browser.driver.find_elements(By.TAG_NAME, "span")
            for button in buttons:
                try:
                    if button.text.strip().upper() == "OK":
                        logger.info(f"「OK」テキストを含むspanを発見しました: {button.text}")
                        button.click()
                        logger.info("✓ テキストで「OK」ボタンをクリックしました")
                        ok_button_found = True
                        break
                except:
                    continue
            
            # spanで見つからない場合はbuttonタグで探す
            if not ok_button_found:
                logger.info("buttonタグで「OK」ボタンを探索します")
                buttons = self.browser.driver.find_elements(By.TAG_NAME, "button")
                for button in buttons:
                    try:
                        if button.text.strip().upper() == "OK":
                            logger.info(f"「OK」テキストを含むbuttonを発見しました: {button.text}")
                            button.click()
                            logger.info("✓ buttonタグで「OK」ボタンをクリックしました")
                            ok_button_found = True
                            break
                    except:
                        continue
            
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
                            try:
                                button_text = btn.text.strip() if hasattr(btn, 'text') else ""
                                logger.info(f"ダイアログボタン {i+1}: テキスト='{button_text}'")
                            except:
                                continue
                        
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
            
            # 設定ファイルからシート名を取得
            candidates_sheet = env.get_config_value('SHEET_NAMES', 'CANDIDATES', '"users_all"').strip('"')
            logger.info(f"求職者一覧データの転記先シート名: {candidates_sheet}")
            
            # スプレッドシートに転記
            if self.import_csv_to_spreadsheet(csv_file_path, candidates_sheet):
                logger.info(f"✓ CSVデータを{candidates_sheet}シートに転記しました")
            else:
                logger.error(f"CSVデータの{candidates_sheet}シートへの転記に失敗しました")
                return False
            
            logger.info("✅ 候補者データのエクスポート処理が完了しました")
            return True
            
        except Exception as e:
            logger.error(f"候補者データのエクスポート処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("export_candidates_error.png")
            return False
    
    def _download_exported_csv(self, max_retries: int = 3, retry_interval: int = 45) -> Optional[str]:
        """
        エクスポートされたCSVファイルをダウンロードする
        
        Args:
            max_retries (int): 最大リトライ回数（デフォルト: 3）
            retry_interval (int): リトライ間隔（秒）（デフォルト: 45）
            
        Returns:
            Optional[str]: ダウンロードされたCSVファイルのパス、失敗した場合はNone
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
                    elements = self.browser.driver.find_elements(By.TAG_NAME, "li")
                    for element in elements:
                        try:
                            title = element.get_attribute("title")
                            if title and "エクスポートの結果一覧を開く" in title:
                                logger.info(f"「エクスポートの結果一覧を開く」タイトルを持つ要素を発見しました")
                                element.click()
                                logger.info("✓ タイトル属性で「エクスポートの結果一覧を開く」ボタンをクリックしました")
                                export_result_button_found = True
                                break
                        except:
                            continue
                
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
                    logger.error("エクスポート結果リストを開くのを諦めます")
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
    
    def import_csv_to_spreadsheet(self, csv_file_path: str, sheet_name: str = 'users_all') -> bool:
        """
        CSVファイルのデータをスプレッドシートに転記する
        
        Args:
            csv_file_path (str): CSVファイルのパス
            sheet_name (str): 転記先のシート名（デフォルト: 'users_all'）
            
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info(f"=== CSVデータの{sheet_name}シートへの転記処理を開始します ===")
            
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
            
            # シートのデータをクリア
            try:
                logger.info(f"{sheet_name}シートのデータをクリアします")
                spreadsheet_manager.clear_worksheet(sheet_name)
                logger.info(f"{sheet_name}シートのデータをクリアしました")
            except Exception as e:
                logger.error(f"{sheet_name}シートのデータクリアに失敗しました: {str(e)}")
                import traceback
                logger.error(traceback.format_exc())
                # クリアに失敗しても続行
            
            # CSVデータをスプレッドシートに転記
            try:
                logger.info(f"CSVデータを{sheet_name}シートに転記します: {csv_file_path}")
                spreadsheet_manager.import_csv_to_sheet(csv_file_path, sheet_name)
                logger.info(f"CSVデータを{sheet_name}シートに転記しました")
            except Exception as e:
                logger.error(f"CSVデータの転記に失敗しました: {str(e)}")
                import traceback
                logger.error(traceback.format_exc())
                return False
            
            logger.info(f"✅ CSVデータの{sheet_name}シートへの転記処理が完了しました")
            return True
            
        except Exception as e:
            logger.error(f"CSVデータの{sheet_name}シートへの転記処理中にエラーが発生しました: {str(e)}")
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
    
    def execute_common_candidates_flow(self):
        """
        求職者関連の共通処理フローを実行する
        
        以下の処理を順番に実行します：
        1. 「その他業務」ボタンをクリックして新しいウィンドウを開く
        2. 求職者メニューをクリック
        
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== 求職者関連の共通処理フローを開始します ===")
            
            # 「その他業務」ボタンをクリック
            if not self.click_other_operations_button():
                logger.error("「その他業務」ボタンのクリック処理に失敗しました")
                return False
            
            # 求職者メニューをクリック
            if not self.click_candidates_menu():
                logger.error("求職者メニューのクリック処理に失敗しました")
                return False
            
            logger.info("✅ 求職者関連の共通処理フローが正常に完了しました")
            return True
            
        except Exception as e:
            logger.error(f"求職者関連の共通処理フロー中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("common_candidates_flow_error.png")
            return False
    
    def execute_operations_flow(self):
        """
        業務操作フローを実行する
        
        以下の処理を順番に実行します：
        1. 求職者関連の共通処理フロー（「その他業務」ボタンクリック、求職者メニュークリック）
        2. 「すべての求職者」リンクをクリック
        3. 「全てチェック」チェックボックスをクリック
        4. 「もっと見る」ボタンを繰り返しクリックして、すべての求職者を表示
        5. 求職者データのエクスポート処理を実行
        
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== 業務操作フローを開始します ===")
            
            # 求職者関連の共通処理フローを実行
            if not self.execute_common_candidates_flow():
                logger.error("求職者関連の共通処理フローに失敗しました")
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
            
            # スプレッドシートの集計処理を実行
            try:
                logger.info("スプレッドシートの日別フェーズカウント集計処理を実行します")

                aggregator = SpreadsheetAggregator()
                if not aggregator.initialize():
                    logger.error("SpreadsheetAggregatorの初期化に失敗しました。集計処理をスキップします。")
                    result = False
                else:
                    result = aggregator.record_daily_phase_counts()

                if result:
                    logger.info("スプレッドシートの日別フェーズカウント集計処理が正常に完了しました")
                else:
                    logger.warning("スプレッドシートの日別フェーズカウント集計処理に失敗しましたが、ブラウザ操作は継続します")
            except Exception as e:
                logger.error(f"スプレッドシートの日別フェーズカウント集計処理中に予期せぬエラーが発生しました: {str(e)}")
                import traceback
                logger.error(traceback.format_exc())
                # 集計処理の失敗は全体の処理に影響を与えないようにする
            
            logger.info("✅ 業務操作フローが正常に完了しました")
            return True
            
        except Exception as e:
            logger.error(f"業務操作フロー中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return False
    
    def execute_both_processes(self):
        """
        求職者一覧と選考プロセスの両方の処理を順に実行する
        
        以下の処理を順番に実行します：
        1. 求職者一覧の処理フロー
        2. 選考プロセスの処理フロー
        
        Returns:
            tuple: (求職者一覧の処理結果, 選考プロセスの処理結果)
        """
        try:
            logger.info("=== 求職者一覧と選考プロセスの両方の処理を順に実行します ===")
            
            # 求職者一覧の処理フロー
            logger.info("1. 求職者一覧のエクスポート処理フローを実行します")
            candidates_success = self.execute_operations_flow()
            if not candidates_success:
                logger.error("求職者一覧のエクスポート処理フローに失敗しました")
            else:
                logger.info("求職者一覧のエクスポート処理フローが正常に完了しました")
            
            # 処理間の待機時間
            logger.info("次の処理に進む前に10秒間待機します")
            time.sleep(10)
            
            # 現在の画面の状態を確認
            current_url = self.browser.driver.current_url
            current_title = self.browser.driver.title
            logger.info(f"現在のURL: {current_url}")
            logger.info(f"現在の画面のタイトル: {current_title}")
            
            # 選考プロセスの処理フロー
            logger.info("2. 選考プロセス一覧の表示処理フローを実行します")
            entryprocess_success = self.access_selection_processes()
            if not entryprocess_success:
                logger.error("選考プロセス一覧の表示処理フローに失敗しました")
            else:
                logger.info("選考プロセス一覧の表示処理フローが正常に完了しました")
            
            # 両方の処理結果をログに出力
            if candidates_success and entryprocess_success:
                logger.info("✅ 両方の処理フローが正常に完了しました")
            else:
                logger.warning("⚠️ 一部の処理フローが失敗しました")
            
            return (candidates_success, entryprocess_success)
            
        except Exception as e:
            logger.error(f"両方の処理フロー中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return (False, False)
    
    def execute_common_selection_flow(self):
        """
        選考プロセス関連の共通処理フローを実行する
        
        以下の処理を順番に実行します：
        1. 既に開いている「その他業務」画面で処理を続行
        
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== 選考プロセス関連の共通処理フローを開始します ===")
            
            # 現在のウィンドウハンドルを確認
            current_handles = self.browser.driver.window_handles
            logger.info(f"現在のウィンドウハンドル: {current_handles}")
            
            # 既に「その他業務」画面が開いていると想定して処理を続行
            logger.info("既に開いている「その他業務」画面で処理を続行します")
            
            # 現在のURLを確認
            current_url = self.browser.driver.current_url
            logger.info(f"現在のURL: {current_url}")
            
            # 現在の画面のタイトルを確認
            current_title = self.browser.driver.title
            logger.info(f"現在の画面のタイトル: {current_title}")
            
            # スクリーンショットを取得
            self.browser.save_screenshot("selection_process_common_flow.png")
            
            logger.info("✅ 選考プロセス関連の共通処理フローが正常に完了しました")
            return True
            
        except Exception as e:
            logger.error(f"選考プロセス関連の共通処理フロー中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("common_selection_flow_error.png")
            return False
    
    def access_selection_processes(self):
        """
        選考プロセス画面にアクセスする
        
        以下の処理を順番に実行します：
        1. 選考プロセス関連の共通処理フロー（既に開いている「その他業務」画面で処理を続行）
        2. メニューバーの「選考プロセス」をクリック
        3. 「すべての選考プロセス」リンクをクリック
        4. 「全てチェック」チェックボックスをクリック
        5. 「もっと見る」ボタンを繰り返しクリックして、すべての選考プロセスを表示
        6. 選考プロセスデータのエクスポート処理を実行
        
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== 選考プロセス画面へのアクセス処理を開始します ===")
            
            # 選考プロセス関連の共通処理フローを実行
            if not self.execute_common_selection_flow():
                logger.error("選考プロセス関連の共通処理フローに失敗しました")
                return False
            
            # 「選考プロセス」メニューをクリック
            if not self.click_selection_process_menu():
                logger.error("「選考プロセス」メニューのクリック処理に失敗しました")
                return False
            
            # 「すべての選考プロセス」リンクをクリック
            if not self.click_all_selection_processes():
                logger.error("「すべての選考プロセス」リンクのクリック処理に失敗しました")
                return False
            
            # 「全てチェック」チェックボックスをクリック
            if not self.select_all_selection_processes():
                logger.warning("選考プロセス一覧の「全てチェック」チェックボックスのクリック処理に失敗しましたが、処理を継続します")
                # 失敗しても処理を継続
            
            # 「もっと見る」ボタンを繰り返しクリックして、すべての選考プロセスを表示
            if not self.click_show_more_selection_processes():
                logger.warning("選考プロセス一覧の「もっと見る」ボタンの繰り返しクリック処理に失敗しましたが、処理を継続します")
                # 失敗しても処理を継続（データ件数が少ない場合は「もっと見る」ボタンが表示されないため）
            
            # 選考プロセスデータのエクスポート処理を実行
            if not self.export_selection_processes_data():
                logger.error("選考プロセスデータのエクスポート処理に失敗しました")
                return False
            
            logger.info("✅ 選考プロセス画面へのアクセス処理が完了しました")
            return True
            
        except Exception as e:
            logger.error(f"選考プロセス画面へのアクセス処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("access_selection_processes_error.png")
            return False
    
    def click_selection_process_menu(self):
        """
        メニューバーの「選考プロセス」をクリックする
        
        指定されたセレクタ:
        #main-menu-id-6 > a
        <a href="javascript:void(0)">選考プロセス</a>
        
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== 「選考プロセス」メニューのクリック処理を開始します ===")
            
            # 「選考プロセス」メニューをクリック
            selection_process_clicked = False
            
            # 指定されたセレクタで試す
            try:
                logger.info("指定されたセレクタで「選考プロセス」メニューを探索します: #main-menu-id-6 > a")
                selection_process_selector = "#main-menu-id-6 > a"
                selection_process_element = WebDriverWait(self.browser.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, selection_process_selector))
                )
                # 要素の情報をログに出力
                logger.info(f"要素のテキスト: '{selection_process_element.text}'")
                logger.info(f"要素のhref属性: '{selection_process_element.get_attribute('href')}'")
                
                # クリック前にスクリーンショットを取得
                self.browser.save_screenshot("before_selection_process_menu_click.png")
                
                # 要素をクリック
                selection_process_element.click()
                logger.info("✓ 指定されたセレクタで「選考プロセス」メニューをクリックしました")
                selection_process_clicked = True
            except Exception as e:
                logger.warning(f"指定されたセレクタでの「選考プロセス」メニュークリックに失敗しました: {str(e)}")
                
                # JavaScriptでクリックを試みる
                try:
                    logger.info("JavaScriptで「選考プロセス」メニューをクリックします")
                    script = """
                    var element = document.querySelector('#main-menu-id-6 > a');
                    if (element) {
                        element.click();
                        return true;
                    }
                    return false;
                    """
                    result = self.browser.driver.execute_script(script)
                    if result:
                        logger.info("✓ JavaScriptで「選考プロセス」メニューをクリックしました")
                        selection_process_clicked = True
                except Exception as js_e:
                    logger.warning(f"JavaScriptでの「選考プロセス」メニュークリックに失敗しました: {str(js_e)}")
            
            # セレクタで失敗した場合はテキストで探す
            if not selection_process_clicked:
                try:
                    logger.info("テキストで「選考プロセス」メニューを探索します")
                    links = self.browser.driver.find_elements(By.TAG_NAME, "a")
                    for link in links:
                        if link.text.strip() == "選考プロセス":
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
        
        指定されたセレクタ:
        #ui-id-196 > li:nth-child(12) > a
        <a title="すべての選考プロセス" href="javascript:void(0)">すべての選考プロセス</a>
        
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== 「すべての選考プロセス」リンクのクリック処理を開始します ===")
            
            # 「すべての選考プロセス」リンクをクリック
            all_processes_clicked = False
            
            # 指定されたセレクタで試す
            try:
                logger.info("指定されたセレクタで「すべての選考プロセス」リンクを探索します: #ui-id-196 > li:nth-child(12) > a")
                all_processes_selector = "#ui-id-196 > li:nth-child(12) > a"
                all_processes_element = WebDriverWait(self.browser.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, all_processes_selector))
                )
                # 要素の情報をログに出力
                logger.info(f"要素のテキスト: '{all_processes_element.text}'")
                logger.info(f"要素のtitle属性: '{all_processes_element.get_attribute('title')}'")
                logger.info(f"要素のhref属性: '{all_processes_element.get_attribute('href')}'")
                
                # クリック前にスクリーンショットを取得
                self.browser.save_screenshot("before_all_selection_processes_click.png")
                
                # 要素をクリック
                all_processes_element.click()
                logger.info("✓ 指定されたセレクタで「すべての選考プロセス」リンクをクリックしました")
                all_processes_clicked = True
            except Exception as e:
                logger.warning(f"指定されたセレクタでの「すべての選考プロセス」リンククリックに失敗しました: {str(e)}")
                
                # JavaScriptでクリックを試みる
                try:
                    logger.info("JavaScriptで「すべての選考プロセス」リンクをクリックします")
                    script = """
                    var element = document.querySelector('#ui-id-196 > li:nth-child(12) > a');
                    if (element) {
                        element.click();
                        return true;
                    }
                    return false;
                    """
                    result = self.browser.driver.execute_script(script)
                    if result:
                        logger.info("✓ JavaScriptで「すべての選考プロセス」リンクをクリックしました")
                        all_processes_clicked = True
                except Exception as js_e:
                    logger.warning(f"JavaScriptでの「すべての選考プロセス」リンククリックに失敗しました: {str(js_e)}")
            
            # title属性で探す
            if not all_processes_clicked:
                try:
                    logger.info("title属性で「すべての選考プロセス」リンクを探索します")
                    links = self.browser.driver.find_elements(By.CSS_SELECTOR, "a[title='すべての選考プロセス']")
                    if links:
                        logger.info("title属性で「すべての選考プロセス」リンクを発見しました")
                        links[0].click()
                        logger.info("✓ title属性で「すべての選考プロセス」リンクをクリックしました")
                        all_processes_clicked = True
                except Exception as e:
                    logger.warning(f"title属性での「すべての選考プロセス」リンククリックに失敗しました: {str(e)}")
            
            # セレクタで失敗した場合はテキストで探す（完全一致）
            if not all_processes_clicked:
                try:
                    logger.info("テキストで「すべての選考プロセス」リンクを探索します（完全一致）")
                    links = self.browser.driver.find_elements(By.TAG_NAME, "a")
                    for link in links:
                        if link.text.strip() == "すべての選考プロセス":
                            logger.info(f"「すべての選考プロセス」テキストを含むリンクを発見しました: {link.text}")
                            link.click()
                            logger.info("✓ テキストで「すべての選考プロセス」リンクをクリックしました")
                            all_processes_clicked = True
                            break
                except Exception as e:
                    logger.warning(f"テキストでの「すべての選考プロセス」リンククリックに失敗しました: {str(e)}")
            
            # 完全一致で失敗した場合は部分一致で探す
            if not all_processes_clicked:
                try:
                    logger.info("テキストで「選考プロセス」を含むリンクを探索します（部分一致）")
                    links = self.browser.driver.find_elements(By.TAG_NAME, "a")
                    for link in links:
                        if "選考プロセス" in link.text:
                            logger.info(f"「選考プロセス」テキストを含むリンクを発見しました: {link.text}")
                            link.click()
                            logger.info("✓ 「選考プロセス」テキストを含むリンクをクリックしました")
                            all_processes_clicked = True
                            break
                except Exception as e:
                    logger.warning(f"部分一致テキストでのリンククリックに失敗しました: {str(e)}")
            
            if not all_processes_clicked:
                # 現在のページのHTMLを保存して分析
                page_html = self.browser.driver.page_source
                html_path = os.path.join(self.screenshot_dir, "selection_process_page.html")
                with open(html_path, 'w', encoding='utf-8') as f:
                    f.write(page_html)
                logger.info(f"現在のページのHTMLを保存しました: {html_path}")
                
                # 利用可能なリンクを表示
                logger.info("利用可能なリンク一覧:")
                links = self.browser.driver.find_elements(By.TAG_NAME, "a")
                for i, link in enumerate(links[:20]):  # 最初の20個だけ表示
                    try:
                        logger.info(f"  {i+1}. テキスト: '{link.text}', title: '{link.get_attribute('title')}', href: '{link.get_attribute('href')}'")
                    except:
                        logger.info(f"  {i+1}. [取得不可]")
                
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
    
    def click_show_more_selection_processes(self, max_attempts=20, interval=5):
        """
        選考プロセス一覧画面で「もっと見る」ボタンを繰り返しクリックして、すべての選考プロセスを表示する
        
        Args:
            max_attempts (int): 最大試行回数
            interval (int): クリック間隔（秒）
            
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== 選考プロセス一覧の「もっと見る」ボタンの繰り返しクリック処理を開始します ===")
            
            # まず現在の表示件数を確認
            try:
                # 提供されたクラス情報を使用して表示件数を取得
                footer_selectors = [
                    ".jss154.data-grid-footer",
                    ".data-grid-footer",
                    ".jss152 .jss153 .jss154",
                    ".jss153 .jss154"
                ]
                
                footer_element = None
                for selector in footer_selectors:
                    elements = self.browser.driver.find_elements(By.CSS_SELECTOR, selector)
                    if elements:
                        footer_element = elements[0]
                        break
                
                if footer_element:
                    footer_text = footer_element.text
                    logger.info(f"データグリッドフッター情報: {footer_text}")
                    
                    # 「X件中Y件表示」の形式から件数を抽出
                    import re
                    match = re.search(r'(\d+)件中(\d+)件表示', footer_text)
                    if match:
                        total_count = int(match.group(1))
                        displayed_count = int(match.group(2))
                        logger.info(f"選考プロセス一覧: 全{total_count}件中{displayed_count}件表示中")
                        
                        # 全件表示されている場合は処理をスキップ
                        if total_count == displayed_count:
                            logger.info("すべての選考プロセスが既に表示されています。「もっと見る」ボタンのクリック処理をスキップします。")
                            return True
            except Exception as e:
                logger.warning(f"表示件数の確認中にエラーが発生しましたが、処理を継続します: {str(e)}")
            
            # スクリーンショットを取得
            self.browser.save_screenshot("selection_processes_before_show_more.png")
            
            # 「もっと見る」ボタンが見つからなくなるまで繰り返しクリック
            attempt = 0
            show_more_button_found = False
            
            while attempt < max_attempts:
                attempt += 1
                logger.info(f"「もっと見る」ボタンのクリック試行: {attempt}/{max_attempts}")
                
                # スクリーンショットを取得
                self.browser.save_screenshot(f"selection_processes_show_more_attempt_{attempt}.png")
                
                # 「もっと見る」ボタンを探す（複数の方法で試行）
                try:
                    # クラス名で検索
                    show_more_button = None
                    
                    # 求職者一覧と同様のクラス名で検索
                    try:
                        logger.info("クラス名で「もっと見る」ボタンを探索します")
                        show_more_button = WebDriverWait(self.browser.driver, 5).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.list-view-show-more-button"))
                        )
                    except:
                        pass
                    
                    # 一般的なボタンクラスで検索
                    if not show_more_button:
                        button_selectors = [
                            "button.list-view-show-more-button",
                            "button.show-more-button",
                            "button.more-button",
                            "button.load-more",
                            ".jss152 button",
                            ".jss153 button"
                        ]
                        
                        for selector in button_selectors:
                            try:
                                elements = self.browser.driver.find_elements(By.CSS_SELECTOR, selector)
                                if elements:
                                    for element in elements:
                                        if "もっと見る" in element.text:
                                            show_more_button = element
                                            logger.info(f"セレクタ '{selector}' で「もっと見る」ボタンを発見しました")
                                            break
                                    if show_more_button:
                                        break
                            except:
                                continue
                    
                    # テキスト内容で検索
                    if not show_more_button:
                        logger.info("テキスト内容で「もっと見る」ボタンを探索します")
                        buttons = self.browser.driver.find_elements(By.TAG_NAME, "button")
                        for button in buttons:
                            try:
                                if "もっと見る" in button.text:
                                    logger.info(f"「もっと見る」テキストを含むボタンを発見しました: {button.text}")
                                    show_more_button = button
                                    break
                            except:
                                continue
                    
                    # ボタンが見つかった場合、クリック
                    if show_more_button:
                        show_more_button_found = True
                        
                        # 要素が画面内に表示されるようにスクロール
                        self.browser.scroll_to_element(show_more_button)
                        
                        # クリック実行
                        show_more_button.click()
                        logger.info(f"✓ 「もっと見る」ボタンをクリックしました（{attempt}回目）")
                        
                        # 次のデータ読み込みを待機
                        time.sleep(interval)
                    else:
                        if show_more_button_found:
                            logger.info("「もっと見る」ボタンが見つかりませんでした。すべてのデータが表示されたと思われます。")
                        else:
                            if attempt == 1:
                                logger.info("「もっと見る」ボタンが見つかりませんでした。データ件数が少ないため表示されていない可能性があります。")
                            else:
                                logger.info("「もっと見る」ボタンが見つかりませんでした。すべてのデータが表示されたと思われます。")
                        break
                        
                except TimeoutException:
                    if show_more_button_found:
                        logger.info("「もっと見る」ボタンが見つかりませんでした。すべてのデータが表示されたと思われます。")
                    else:
                        if attempt == 1:
                            logger.info("「もっと見る」ボタンが見つかりませんでした。データ件数が少ないため表示されていない可能性があります。")
                        else:
                            logger.info("「もっと見る」ボタンが見つかりませんでした。すべてのデータが表示されたと思われます。")
                    break
                except Exception as e:
                    logger.warning(f"{attempt}回目の「もっと見る」ボタンクリック中にエラーが発生しましたが、処理を継続します: {str(e)}")
                    # エラーが発生しても処理を継続
                    time.sleep(interval)
            
            # データグリッドコンテナを一番下までスクロール
            logger.info("データグリッドコンテナを一番下までスクロールします")
            
            # データグリッドコンテナを探す（提供されたクラス情報を使用）
            data_grid_container = None
            grid_selectors = [
                ".jss152",
                ".jss153",
                "div[role='grid']",
                ".data-grid-container",
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
                    self.browser.save_screenshot("selection_processes_after_grid_scroll_bottom.png")
                    
                except Exception as e:
                    logger.warning(f"データグリッドコンテナのスクロール中にエラーが発生しました: {str(e)}")
                    self._scroll_page_fallback()
            else:
                logger.warning("データグリッドコンテナが見つかりませんでした")
                self._scroll_page_fallback()
            
            # 最終的な画面のスクリーンショットを取得
            self.browser.save_screenshot("selection_processes_after_show_more_all.png")
            
            # 最終的な表示件数を確認
            try:
                footer_element = None
                for selector in footer_selectors:
                    elements = self.browser.driver.find_elements(By.CSS_SELECTOR, selector)
                    if elements:
                        footer_element = elements[0]
                        break
                
                if footer_element:
                    footer_text = footer_element.text
                    logger.info(f"最終的なデータグリッドフッター情報: {footer_text}")
            except Exception as e:
                logger.warning(f"最終的な表示件数の確認中にエラーが発生しました: {str(e)}")
            
            logger.info(f"✅ 選考プロセス一覧の「もっと見る」ボタンの繰り返しクリック処理が完了しました（{attempt}回試行）")
            return True
            
        except Exception as e:
            logger.error(f"選考プロセス一覧の「もっと見る」ボタンの繰り返しクリック処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("selection_processes_show_more_error.png")
            return False
    
    def select_all_selection_processes(self):
        """
        選考プロセス一覧画面で「全てチェック」チェックボックスをクリックする
        
        指定されたセレクタ:
        #recordListView > div.jss37 > div:nth-child(2) > div > div.jss45 > span > span > input
        
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== 選考プロセス一覧の「全てチェック」チェックボックスのクリック処理を開始します ===")
            
            # 画面の読み込みを待機
            time.sleep(3)
            self.browser.save_screenshot("selection_processes_before_select_all.png")
            
            # 「全てチェック」チェックボックスをクリック
            checkbox_clicked = False
            
            # 指定されたセレクタで試す
            try:
                logger.info("指定されたセレクタで「全てチェック」チェックボックスを探索します: #recordListView > div.jss37 > div:nth-child(2) > div > div.jss45 > span > span > input")
                checkbox_selector = "#recordListView > div.jss37 > div:nth-child(2) > div > div.jss45 > span > span > input"
                checkbox_element = WebDriverWait(self.browser.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, checkbox_selector))
                )
                # 要素の情報をログに出力
                logger.info(f"要素のtype属性: '{checkbox_element.get_attribute('type')}'")
                logger.info(f"要素のclass属性: '{checkbox_element.get_attribute('class')}'")
                
                # クリック前にスクリーンショットを取得
                self.browser.save_screenshot("selection_processes_before_checkbox_click.png")
                
                # 要素をクリック
                checkbox_element.click()
                logger.info("✓ 指定されたセレクタで「全てチェック」チェックボックスをクリックしました")
                checkbox_clicked = True
            except Exception as e:
                logger.warning(f"指定されたセレクタでの「全てチェック」チェックボックスクリックに失敗しました: {str(e)}")
                
                # 一般的なセレクタを試す
                try:
                    logger.info("一般的なセレクタで「全てチェック」チェックボックスを探索します")
                    checkbox_selectors = [
                        "#recordListView input[type='checkbox']",
                        ".jss37 input[type='checkbox']",
                        ".jss45 input[type='checkbox']",
                        "input[type='checkbox']"
                    ]
                    
                    for selector in checkbox_selectors:
                        try:
                            elements = self.browser.driver.find_elements(By.CSS_SELECTOR, selector)
                            if elements:
                                # 最初のチェックボックスが「全てチェック」の可能性が高い
                                elements[0].click()
                                logger.info(f"✓ セレクタ '{selector}' で「全てチェック」チェックボックスをクリックしました")
                                checkbox_clicked = True
                                break
                        except:
                            continue
                except Exception as general_e:
                    logger.warning(f"一般的なセレクタでの「全てチェック」チェックボックスクリックにも失敗しました: {str(general_e)}")
            
            # JavaScriptでクリックを試みる
            if not checkbox_clicked:
                try:
                    logger.info("JavaScriptで「全てチェック」チェックボックスをクリックします")
                    script = """
                    var checkbox = document.querySelector("#recordListView > div.jss37 > div:nth-child(2) > div > div.jss45 > span > span > input");
                    if (checkbox) {
                        checkbox.click();
                        return true;
                    }
                    
                    // 一般的なセレクタで探す
                    var checkboxes = document.querySelectorAll("input[type='checkbox']");
                    if (checkboxes.length > 0) {
                        checkboxes[0].click();
                        return true;
                    }
                    
                    return false;
                    """
                    result = self.browser.driver.execute_script(script)
                    if result:
                        logger.info("✓ JavaScriptで「全てチェック」チェックボックスをクリックしました")
                        checkbox_clicked = True
                except Exception as js_e:
                    logger.warning(f"JavaScriptでの「全てチェック」チェックボックスクリックに失敗しました: {str(js_e)}")
            
            if not checkbox_clicked:
                logger.error("「全てチェック」チェックボックスが見つかりませんでした")
                
                # 現在のページのHTMLを保存して分析
                page_html = self.browser.driver.page_source
                html_path = os.path.join(self.screenshot_dir, "selection_processes_checkbox_page.html")
                with open(html_path, 'w', encoding='utf-8') as f:
                    f.write(page_html)
                logger.info(f"現在のページのHTMLを保存しました: {html_path}")
                
                # 利用可能なチェックボックスを表示
                logger.info("利用可能なチェックボックス一覧:")
                checkboxes = self.browser.driver.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
                for i, checkbox in enumerate(checkboxes):
                    try:
                        logger.info(f"  {i+1}. class: '{checkbox.get_attribute('class')}', name: '{checkbox.get_attribute('name')}', id: '{checkbox.get_attribute('id')}'")
                    except:
                        logger.info(f"  {i+1}. [取得不可]")
                
                return False
            
            # 処理待機
            time.sleep(2)
            self.browser.save_screenshot("selection_processes_after_select_all.png")
            
            logger.info("✅ 選考プロセス一覧の「全てチェック」チェックボックスのクリック処理が完了しました")
            return True
            
        except Exception as e:
            logger.error(f"選考プロセス一覧の「全てチェック」チェックボックスのクリック処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("selection_processes_select_all_error.png")
            return False
    
    def export_selection_processes_data(self) -> bool:
        """
        選考プロセスデータをエクスポートする
        
        指定されたセレクタ:
        アクションボタン: #recordListView > div.jss37 > div:nth-child(2) > div > button > div
        エクスポートリンク: #pageProcess > div:nth-child(23) > div > ul > li.jss157.linkExport
        <li class="jss157 linkExport" title="エクスポート">エクスポート</li>
        
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== 選考プロセスデータのエクスポート処理を開始します ===")
            
            # アクションリストボタンをクリック
            action_button_clicked = False
            
            # 指定されたセレクタで試す
            try:
                logger.info("指定されたセレクタでアクションボタンを探索します: #recordListView > div.jss37 > div:nth-child(2) > div > button > div")
                action_button_selector = "#recordListView > div.jss37 > div:nth-child(2) > div > button > div"
                action_button_element = WebDriverWait(self.browser.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, action_button_selector))
                )
                # クリック前にスクリーンショットを取得
                self.browser.save_screenshot("selection_processes_before_action_button_click.png")
                
                # 要素をクリック
                action_button_element.click()
                logger.info("✓ 指定されたセレクタでアクションボタンをクリックしました")
                action_button_clicked = True
            except Exception as e:
                logger.warning(f"指定されたセレクタでのアクションボタンクリックに失敗しました: {str(e)}")
                
                # 一般的なセレクタを試す
                try:
                    logger.info("一般的なセレクタでアクションボタンを探索します")
                    action_button_selectors = [
                        "#recordListView button",
                        ".jss37 button",
                        "button.action-button",
                        "button.menu-button"
                    ]
                    
                    for selector in action_button_selectors:
                        try:
                            elements = self.browser.driver.find_elements(By.CSS_SELECTOR, selector)
                            if elements:
                                # 最初のボタンをクリック
                                elements[0].click()
                                logger.info(f"✓ セレクタ '{selector}' でアクションボタンをクリックしました")
                                action_button_clicked = True
                                break
                        except:
                            continue
                except Exception as general_e:
                    logger.warning(f"一般的なセレクタでのアクションボタンクリックにも失敗しました: {str(general_e)}")
            
            # JavaScriptでクリックを試みる
            if not action_button_clicked:
                try:
                    logger.info("JavaScriptでアクションボタンをクリックします")
                    script = """
                    var button = document.querySelector("#recordListView > div.jss37 > div:nth-child(2) > div > button > div");
                    if (button) {
                        button.click();
                        return true;
                    }
                    
                    // 親要素をクリック
                    var parentButton = document.querySelector("#recordListView > div.jss37 > div:nth-child(2) > div > button");
                    if (parentButton) {
                        parentButton.click();
                        return true;
                    }
                    
                    return false;
                    """
                    result = self.browser.driver.execute_script(script)
                    if result:
                        logger.info("✓ JavaScriptでアクションボタンをクリックしました")
                        action_button_clicked = True
                except Exception as js_e:
                    logger.warning(f"JavaScriptでのアクションボタンクリックに失敗しました: {str(js_e)}")
            
            if not action_button_clicked:
                logger.error("アクションボタンが見つかりませんでした")
                return False
            
            # 少し待機してメニューが表示されるのを待つ
            time.sleep(2)
            self.browser.save_screenshot("selection_processes_after_action_button_click.png")
            
            # エクスポートボタンをクリック
            export_button_clicked = False
            
            # 指定されたセレクタで試す
            try:
                logger.info("指定されたセレクタでエクスポートボタンを探索します: #pageProcess > div:nth-child(23) > div > ul > li.jss157.linkExport")
                export_button_selector = "#pageProcess > div:nth-child(23) > div > ul > li.jss157.linkExport"
                export_button_element = WebDriverWait(self.browser.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, export_button_selector))
                )
                # クリック前にスクリーンショットを取得
                self.browser.save_screenshot("selection_processes_before_export_button_click.png")
                
                # 要素をクリック
                export_button_element.click()
                logger.info("✓ 指定されたセレクタでエクスポートボタンをクリックしました")
                export_button_clicked = True
            except Exception as e:
                logger.warning(f"指定されたセレクタでのエクスポートボタンクリックに失敗しました: {str(e)}")
                
                # クラス名で試す
                try:
                    logger.info("クラス名でエクスポートボタンを探索します: li.jss157.linkExport")
                    export_elements = self.browser.driver.find_elements(By.CSS_SELECTOR, "li.jss157.linkExport")
                    if export_elements:
                        export_elements[0].click()
                        logger.info("✓ クラス名でエクスポートボタンをクリックしました")
                        export_button_clicked = True
                except Exception as class_e:
                    logger.warning(f"クラス名でのエクスポートボタンクリックに失敗しました: {str(class_e)}")
                
                # タイトル属性で試す
                if not export_button_clicked:
                    try:
                        logger.info("タイトル属性でエクスポートボタンを探索します")
                        export_elements = self.browser.driver.find_elements(By.CSS_SELECTOR, "li[title='エクスポート']")
                        if export_elements:
                            export_elements[0].click()
                            logger.info("✓ タイトル属性でエクスポートボタンをクリックしました")
                            export_button_clicked = True
                    except Exception as title_e:
                        logger.warning(f"タイトル属性でのエクスポートボタンクリックに失敗しました: {str(title_e)}")
                
                # テキスト内容で試す
                if not export_button_clicked:
                    try:
                        logger.info("テキスト内容でエクスポートボタンを探索します")
                        elements = self.browser.driver.find_elements(By.TAG_NAME, "li")
                        for element in elements:
                            if "エクスポート" in element.text:
                                logger.info(f"「エクスポート」テキストを含む要素を発見しました: {element.text}")
                                element.click()
                                logger.info("✓ テキスト内容でエクスポートボタンをクリックしました")
                                export_button_clicked = True
                                break
                    except Exception as text_e:
                        logger.warning(f"テキスト内容でのエクスポートボタンクリックに失敗しました: {str(text_e)}")
            
            if not export_button_clicked:
                logger.error("エクスポートボタンが見つかりませんでした")
                return False
            
            # エクスポートモーダルで「求人打診~内定まで」を選択
            try:
                logger.info("「求人打診~内定まで」オプションを探索します")
                time.sleep(2)  # モーダルが表示されるまで待機
                self.browser.save_screenshot("selection_processes_export_modal.png")
                
                # 「求人打診~内定まで」のラジオボタンを選択
                option_selected = False
                
                # 指定されたセレクタで試す
                try:
                    logger.info("指定されたセレクタで「求人打診~内定まで」オプションを探索します: #porters-pdialog_2 > div.mapping > div > div > div > ul > li:nth-child(2) > label > span")
                    option_selector = "#porters-pdialog_2 > div.mapping > div > div > div > ul > li:nth-child(2) > label > span"
                    option_element = WebDriverWait(self.browser.driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, option_selector))
                    )
                    # 要素の情報をログに出力
                    logger.info(f"要素のテキスト: '{option_element.text}'")
                    
                    # クリック前にスクリーンショットを取得
                    self.browser.save_screenshot("before_option_click.png")
                    
                    # 要素をクリック
                    option_element.click()
                    logger.info("✓ 指定されたセレクタで「求人打診~内定まで」オプションをクリックしました")
                    option_selected = True
                except Exception as e:
                    logger.warning(f"指定されたセレクタでの「求人打診~内定まで」オプションクリックに失敗しました: {str(e)}")
                    
                    # 親要素のラベルを試す
                    try:
                        logger.info("親要素のラベルで「求人打診~内定まで」オプションを探索します")
                        label_selector = "#porters-pdialog_2 > div.mapping > div > div > div > ul > li:nth-child(2) > label"
                        label_element = WebDriverWait(self.browser.driver, 5).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, label_selector))
                        )
                        label_element.click()
                        logger.info("✓ 親要素のラベルで「求人打診~内定まで」オプションをクリックしました")
                        option_selected = True
                    except Exception as label_e:
                        logger.warning(f"親要素のラベルでの「求人打診~内定まで」オプションクリックに失敗しました: {str(label_e)}")
                
                # テキスト内容で探す
                if not option_selected:
                    try:
                        logger.info("テキスト内容で「求人打診~内定まで」オプションを探索します")
                        labels = self.browser.driver.find_elements(By.TAG_NAME, "label")
                        for label in labels:
                            if "求人打診" in label.text and "内定" in label.text:
                                logger.info(f"「求人打診~内定まで」テキストを含むラベルを発見しました: {label.text}")
                                label.click()
                                logger.info("✓ テキスト内容で「求人打診~内定まで」オプションをクリックしました")
                                option_selected = True
                                break
                    except Exception as text_e:
                        logger.warning(f"テキスト内容での「求人打診~内定まで」オプションクリックに失敗しました: {str(text_e)}")
                
                # ラジオボタンを直接探す
                if not option_selected:
                    try:
                        logger.info("ラジオボタンを直接探索します")
                        radio_buttons = self.browser.driver.find_elements(By.CSS_SELECTOR, "input[type='radio']")
                        if len(radio_buttons) >= 2:  # 少なくとも2つのラジオボタンがある場合
                            # 2番目のラジオボタンをクリック
                            radio_buttons[1].click()
                            logger.info("✓ 2番目のラジオボタンをクリックしました")
                            option_selected = True
                    except Exception as radio_e:
                        logger.warning(f"ラジオボタン直接クリックに失敗しました: {str(radio_e)}")
                
                if not option_selected:
                    logger.warning("「求人打診~内定まで」オプションの選択に失敗しましたが、処理を継続します")
                
                # 「次へ」ボタンをクリック
                logger.info("テキストで「次へ」ボタンを探索します")
                buttons = self.browser.driver.find_elements(By.TAG_NAME, "span")
                next_button_found = False
                
                for button in buttons:
                    if "次へ" in button.text:
                        logger.info(f"「次へ」テキストを含むボタンを発見しました: {button.text}")
                        button.click()
                        logger.info("✓ テキストで「次へ」ボタンをクリックしました")
                        next_button_found = True
                        break
                
                if not next_button_found:
                    logger.error("「次へ」ボタンが見つかりませんでした")
                    return False
                
                # 「次へ」ボタンを2回クリック
                for i in range(2):
                    time.sleep(1)  # 画面の遷移を待つ
                    logger.info(f"テキストで{i+2}回目の「次へ」ボタンを探索します")
                    buttons = self.browser.driver.find_elements(By.TAG_NAME, "span")
                    next_button_found = False
                    
                    for button in buttons:
                        if "次へ" in button.text:
                            logger.info(f"「次へ」テキストを含むボタンを発見しました: {button.text}")
                            button.click()
                            logger.info(f"✓ テキストで{i+2}回目の「次へ」ボタンをクリックしました")
                            next_button_found = True
                            break
                    
                    if not next_button_found:
                        # 最後のボタンは「実行」の可能性がある
                        if i == 1:
                            logger.info("最後のボタンは「実行」の可能性があるため、「実行」ボタンを探索します")
                            for button in buttons:
                                if "実行" in button.text:
                                    logger.info(f"「実行」テキストを含むボタンを発見しました: {button.text}")
                                    button.click()
                                    logger.info("✓ テキストで「実行」ボタンをクリックしました")
                                    next_button_found = True
                                    break
                        
                        if not next_button_found:
                            logger.warning(f"{i+2}回目の「次へ」または「実行」ボタンが見つかりませんでしたが、処理を継続します")
                
                # 「実行」ボタンをクリック
                time.sleep(1)  # 画面の遷移を待つ
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
                                buttons = []
                                try:
                                    buttons = pane.find_elements(By.TAG_NAME, "button")
                                    logger.info(f"ペイン内に{len(buttons)}個のボタンを発見しました")
                                    
                                    # 最後のボタンが「実行」ボタンの可能性が高い
                                    if buttons:
                                        last_button = buttons[-1]
                                        button_text = last_button.text if hasattr(last_button, 'text') else ""
                                        logger.info(f"最後のボタンをクリックします: '{button_text}'")
                                        last_button.click()
                                        logger.info("✓ 最後のボタンをクリックしました")
                                        execute_button_found = True
                                        break
                                except Exception as btn_e:
                                    logger.warning(f"ボタン探索中にエラーが発生しました: {str(btn_e)}")
                                    continue
                    except Exception as e:
                        logger.warning(f"ダイアログのボタンペイン探索中にエラーが発生しました: {str(e)}")
                
                if not execute_button_found:
                    logger.warning("「実行」ボタンが見つかりませんでしたが、処理を継続します")
                
                # OKボタンをクリック
                time.sleep(45)  # エクスポート処理の完了を待つ（30秒から45秒に増加）
                
                # テキストで「OK」ボタンを探す
                logger.info("テキストで「OK」ボタンを探索します")
                ok_button_found = False
                
                # まずspanタグ内のテキストで探す
                buttons = self.browser.driver.find_elements(By.TAG_NAME, "span")
                for button in buttons:
                    try:
                        if button.text.strip().upper() == "OK":
                            logger.info(f"「OK」テキストを含むspanを発見しました: {button.text}")
                            button.click()
                            logger.info("✓ テキストで「OK」ボタンをクリックしました")
                            ok_button_found = True
                            break
                    except:
                        continue
                
                # spanで見つからない場合はbuttonタグで探す
                if not ok_button_found:
                    logger.info("buttonタグで「OK」ボタンを探索します")
                    buttons = self.browser.driver.find_elements(By.TAG_NAME, "button")
                    for button in buttons:
                        try:
                            if button.text.strip().upper() == "OK":
                                logger.info(f"「OK」テキストを含むbuttonを発見しました: {button.text}")
                                button.click()
                                logger.info("✓ buttonタグで「OK」ボタンをクリックしました")
                                ok_button_found = True
                                break
                        except:
                            continue
                
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
                                try:
                                    button_text = btn.text.strip() if hasattr(btn, 'text') else ""
                                    logger.info(f"ダイアログボタン {i+1}: テキスト='{button_text}'")
                                except:
                                    continue
                            
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
                    logger.warning("「OK」ボタンが見つかりませんでした")
                    # 処理を継続
                
            except Exception as modal_e:
                logger.warning(f"エクスポートモーダル操作中にエラーが発生しました: {str(modal_e)}")
                # 処理を継続
            
            # CSVファイルをダウンロード
            csv_file_path = self._download_exported_csv()
            if not csv_file_path:
                logger.error("CSVファイルのダウンロードに失敗しました")
                return False
            logger.info(f"✓ CSVファイルをダウンロードしました: {csv_file_path}")
            
            # 設定ファイルからシート名を取得
            entryprocess_sheet = env.get_config_value('SHEET_NAMES', 'ENTRYPROCESS', '"entryprocess_all"').strip('"')
            logger.info(f"選考プロセス一覧データの転記先シート名: {entryprocess_sheet}")
            
            # スプレッドシートに転記
            if self.import_csv_to_spreadsheet(csv_file_path, entryprocess_sheet):
                logger.info(f"✓ CSVデータを{entryprocess_sheet}シートに転記しました")
            else:
                logger.error(f"CSVデータの{entryprocess_sheet}シートへの転記に失敗しました")
                return False
            
            logger.info("✅ 選考プロセスデータのエクスポート処理が完了しました")
            return True
            
        except Exception as e:
            logger.error(f"選考プロセスデータのエクスポート処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            self.browser.save_screenshot("export_selection_processes_error.png")
            return False