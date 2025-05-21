#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
スプレッドシートの集計処理を行うモジュール

このモジュールは、Google Spreadsheetのデータ集計機能を提供します。
ユーザーフェーズ別集計とエントリープロセス集計の両方の機能が含まれています。
"""

import os
import sys
import datetime
from pathlib import Path
from typing import Dict, List, Any, Tuple, Optional
import copy
import unicodedata # Unicode正規化のために追加

# プロジェクトのルートディレクトリをPYTHONPATHに追加
root_dir = Path(__file__).resolve().parent.parent.parent
sys.path.append(str(root_dir))

from src.utils.spreadsheet import SpreadsheetManager
from src.utils.environment import EnvironmentUtils as env
from src.utils.logging_config import get_logger

logger = get_logger(__name__)

def _custom_col_to_a1(col: int) -> str:
    """
    1から始まる列番号をA1形式の列名に変換します。
    例: 1 -> 'A', 27 -> 'AA'
    """
    if not isinstance(col, int) or col < 1:
        raise ValueError("列番号は正の整数である必要があります")
    
    string = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        string = chr(65 + remainder) + string
    return string

class SpreadsheetAggregator:
    """
    スプレッドシートの集計処理を行うクラス
    """
    
    def __init__(self):
        """
        SpreadsheetAggregatorの初期化
        """
        self.spreadsheet_manager = None
        self.error_handler = None
        
        # ErrorHandler初期化
        try:
            from src.utils.error_handler import ErrorHandler
            self.error_handler = ErrorHandler('spreadsheet_aggregator')
            logger.debug("ErrorHandler初期化完了")
        except Exception as e:
            logger.warning(f"ErrorHandler初期化に失敗: {e}")
        
        self.phase_counts = {
            "相談前×推薦前(新規エントリー)": 0,
            "相談済×推薦前(open)": 0,
            "推薦済(仮エントリー)": 0,
            "面談設定済": 0,
            "終了": 0,
            # "エージェント前相談": 0, # この行をコメントアウトまたは削除
            # "その他": 0
        }
        logger.debug(f"SpreadsheetAggregator initialized with phases: {list(self.phase_counts.keys())}")
    
    def _notify_error(self, error_message: str, exception: Exception, context: Dict[str, Any]):
        """
        エラーを通知する
        
        Args:
            error_message (str): エラーメッセージ
            exception (Exception): 発生した例外
            context (Dict[str, Any]): エラーのコンテキスト情報
        """
        logger.error(f"{error_message}: {exception}")
        import traceback
        logger.error(traceback.format_exc())
        
        # スプレッドシート情報も追加
        if self.spreadsheet_manager is not None:
            try:
                sheet_info = self.spreadsheet_manager.get_spreadsheet_info()
                if sheet_info:
                    context["シート情報"] = sheet_info
            except Exception as e:
                logger.warning(f"スプレッドシート情報の取得に失敗: {e}")
        
        # ErrorHandlerがあれば使用
        if self.error_handler:
            try:
                self.error_handler.handle(error_message, exception, context)
            except Exception as e:
                logger.error(f"ErrorHandlerでのエラー通知に失敗: {e}")
                
                # スラック通知（フォールバック）
                try:
                    from src.utils.slack_notifier import SlackNotifier
                    notifier = SlackNotifier.get_instance()
                    notifier.send_error(error_message, exception, context=context)
                except Exception as notifier_error:
                    logger.error(f"Slack通知に失敗: {notifier_error}")
        else:
            # ErrorHandlerがない場合は直接Slack通知
            try:
                from src.utils.slack_notifier import SlackNotifier
                notifier = SlackNotifier.get_instance()
                notifier.send_error(error_message, exception, context=context)
            except Exception as notifier_error:
                logger.error(f"Slack通知に失敗: {notifier_error}")
    
    def initialize(self) -> bool:
        """
        スプレッドシートマネージャーの初期化

        Returns:
            bool: 初期化が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            self.spreadsheet_manager = SpreadsheetManager()
            self.spreadsheet_manager.open_spreadsheet()
            return True
        except Exception as e:
            logger.error(f"スプレッドシートマネージャーの初期化に失敗しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return False
    
    def aggregate_users_by_phase(self, aggregation_date=None):
        """各フェーズごとのユーザー数を集計する
        
        Args:
            aggregation_date (datetime.date, optional): 集計日. デフォルトは今日

        Returns:
            bool: 処理が成功したか
        """
        try:
            logger.info("=== フェーズ別ユーザー数の集計処理を開始します ===")
            
            # settings.iniからシート名を取得
            try:
                # 設定ファイルから各シート名を取得
                users_all_sheet_name = env.get_config_value('SHEET_NAMES', 'USERSALL').strip('"\'')
                count_users_sheet_name = env.get_config_value('SHEET_NAMES', 'COUNT_USERS').strip('"\'')

                logger.info(f"設定ファイルのシート名:")
                logger.info(f"  USERSALL: '{users_all_sheet_name}'")
                logger.info(f"  COUNT_USERS: '{count_users_sheet_name}'")
            except Exception as e:
                logger.error(f"設定ファイルからのシート名取得に失敗: {str(e)}")
                return False
            
            # 現在の日付を取得 (yyyy/mm/dd形式)
            today = aggregation_date.strftime("%Y/%m/%d") if aggregation_date else datetime.datetime.now().strftime("%Y/%m/%d")
            logger.info(f"集計日: {today}")
            
            # USERS_ALLシートからデータを取得
            users_worksheet = self.spreadsheet_manager.get_worksheet(users_all_sheet_name)
            users_data = users_worksheet.get_all_values()
            
            if not users_data or len(users_data) < 2: # データがないかヘッダーのみの場合はエラー
                logger.error(f"{users_all_sheet_name}シートにデータがありません、またはヘッダーのみです。")
                return False
            
            headers_users_all = users_data[0]
            phase_col_idx_users_all = None
            for i, h in enumerate(headers_users_all):
                if h == "フェーズ":
                    phase_col_idx_users_all = i
                    break
            
            if phase_col_idx_users_all is None:
                logger.error(f"{users_all_sheet_name}シートに「フェーズ」列が見つかりません。")
                return False

            unique_phases_in_sheet = set()
            for r_idx, row_data in enumerate(users_data[1:], start=1): # start=1 でヘッダーを除いた行番号と一致
                if phase_col_idx_users_all < len(row_data):
                    phase_val = row_data[phase_col_idx_users_all].strip()
                    if phase_val: # 空文字は除外
                        unique_phases_in_sheet.add(phase_val)
                else:
                    logger.warning(f"{users_all_sheet_name}シートの {r_idx+1}行目はフェーズ列のデータが不足しています。") # r_idx+1 で実際のシート行番号

            logger.info(f"--- users_allシートのフェーズ名デバッグ ---")
            logger.info(f"定義済みフェーズ (self.phase_counts): {list(self.phase_counts.keys())}")
            logger.info(f"users_allシート上のユニークなフェーズ名: {list(unique_phases_in_sheet)}")
            
            defined_set = set(self.phase_counts.keys())
            # sheet_set は unique_phases_in_sheet をそのまま使う
            
            logger.info(f"定義にあってシートにないフェーズ: {list(defined_set - unique_phases_in_sheet)}")
            logger.info(f"シートにあって定義にないフェーズ: {list(unique_phases_in_sheet - defined_set)}")
            logger.info(f"--- デバッグ終了 ---")
            
            # ヘッダー行を取得 (これは users_data から再取得しても良いが、既に headers_users_all がある)
            headers = users_data[0] # ここでの headers は users_all のヘッダー
            
            # フェーズ列のインデックスを取得
            phase_index = -1
            for i, header_name in enumerate(headers):
                if header_name == "フェーズ":
                    phase_index = i
                    break
            
            if phase_index == -1:
                logger.error("users_allシートに「フェーズ」列が見つかりません。")
                return None # エラーとしてNoneを返す

            # 登録経路列のインデックスを取得
            registration_route_index = -1
            for i, header_name in enumerate(headers):
                if header_name == "登録経路":
                    registration_route_index = i
                    break
            # registration_route_index が -1 の場合でも処理は継続（「不明」として扱われる）

            
            # フェーズごとのカウントを初期化 (self.phase_countsをディープコピーして使用)
            phase_counts = {
                "全体": copy.deepcopy(self.phase_counts)
            }
            
            # 登録経路ごとのカウントも初期化
            defined_routes = [s for s in self.sections if s != "全体"]
            for route in defined_routes:
                phase_counts[route] = copy.deepcopy(self.phase_counts)

            logger.info(f"集計対象の登録経路 (phase_countsのキー): {list(phase_counts.keys())}")
            logger.info(f"集計対象のフェーズ定義（全体）: {list(phase_counts.get('全体', {}).keys())}")


            # データ行を処理してフェーズごとのカウントを集計
            for row_idx, row in enumerate(users_data[1:], start=2):  # ヘッダー行をスキップし、行番号は2から開始 (ログ表示用)
                if not any(row): # 空行はスキップ
                    logger.debug(f"Skipping empty row at {row_idx}")
                    continue

                raw_phase_from_sheet = row[phase_index] if phase_index < len(row) and row[phase_index] else "未分類"
                # Unicode正規化 (NFC) と strip を適用
                normalized_phase_from_sheet = unicodedata.normalize('NFC', raw_phase_from_sheet).strip()
                
                # strip()適用の前後の値をログに出力
                logger.debug(f"Row {row_idx}: Raw Phase='{raw_phase_from_sheet}', Normalized+Stripped Phase='{normalized_phase_from_sheet}' (len={len(normalized_phase_from_sheet)}, type={type(normalized_phase_from_sheet)})")
                if normalized_phase_from_sheet:
                    logger.debug(f"  Normalized+Stripped Phase char codes: {[ord(c) for c in normalized_phase_from_sheet]}")

                phase_to_compare = normalized_phase_from_sheet # 比較に使用するフェーズ名

                registration_route_raw = row[registration_route_index].strip() if registration_route_index != -1 and registration_route_index < len(row) and row[registration_route_index] else "不明"
                registration_route = unicodedata.normalize('NFC', registration_route_raw).strip()
                
                # 「全体」の集計
                counted_for_overall = False
                for defined_phase_key in phase_counts["全体"].keys():
                    normalized_defined_key = unicodedata.normalize('NFC', defined_phase_key)
                    # 比較前に両方の文字コードをログに出力
                    if phase_to_compare == "相談済×推薦前(open)" and normalized_defined_key == "相談済×推薦前(open)": # 特定のケースで詳細ログ
                        logger.debug(f"  Comparing for '全体': Sheet Phase='{phase_to_compare}' (codes: {[ord(c) for c in phase_to_compare]}), Defined Key='{normalized_defined_key}' (codes: {[ord(c) for c in normalized_defined_key]}) Trig: {phase_to_compare == normalized_defined_key}")
                    
                    if phase_to_compare == normalized_defined_key:
                        phase_counts["全体"][defined_phase_key] += 1 # 元のキーでカウント
                        counted_for_overall = True
                        logger.debug(f"    [全体] Match found: '{phase_to_compare}' == '{normalized_defined_key}'. Counted for '{defined_phase_key}'.")
                        break
                
                if not counted_for_overall and phase_to_compare and phase_to_compare != "未分類":
                    unknown_phase_codes = [ord(c) for c in phase_to_compare]
                    logger.warning(f"全体セクションで未知のフェーズ: '{phase_to_compare}' (char codes: {unknown_phase_codes}) (users_allシート {row_idx}行目) - 集計から除外されます。 Defined keys: {list(phase_counts['全体'].keys())}")

                # 登録経路別の集計
                if registration_route in phase_counts: 
                    counted_for_route = False
                    for defined_phase_key in phase_counts[registration_route].keys():
                        normalized_defined_key = unicodedata.normalize('NFC', defined_phase_key)
                        if phase_to_compare == "相談済×推薦前(open)" and normalized_defined_key == "相談済×推薦前(open)": # 特定のケースで詳細ログ
                             logger.debug(f"  Comparing for route '{registration_route}': Sheet Phase='{phase_to_compare}' (codes: {[ord(c) for c in phase_to_compare]}), Defined Key='{normalized_defined_key}' (codes: {[ord(c) for c in normalized_defined_key]}) Trig: {phase_to_compare == normalized_defined_key}")

                        if phase_to_compare == normalized_defined_key:
                            phase_counts[registration_route][defined_phase_key] += 1 # 元のキーでカウント
                            counted_for_route = True
                            logger.debug(f"    [{registration_route}] Match found: '{phase_to_compare}' == '{normalized_defined_key}'. Counted for '{defined_phase_key}'.")
                            break
                    
                    if not counted_for_route and phase_to_compare and phase_to_compare != "未分類":
                        unknown_phase_codes = [ord(c) for c in phase_to_compare]
                        logger.warning(f"登録経路 '{registration_route}' で未知のフェーズ: '{phase_to_compare}' (char codes: {unknown_phase_codes}) (users_allシート {row_idx}行目) - 集計から除外されます。 Defined keys: {list(phase_counts[registration_route].keys())}")
                elif registration_route and registration_route != "不明": 
                    logger.warning(f"未知の登録経路: '{registration_route}' (users_allシート {row_idx}行目) - この経路のデータは登録経路別集計には含まれません。")

            logger.info(f"フェーズごとのカウント（全体）最終結果: {phase_counts.get('全体', {})}")
            
            # COUNT_USERSシートからデータを取得
            count_worksheet = self.spreadsheet_manager.get_worksheet(count_users_sheet_name)
            count_data = count_worksheet.get_all_values()
            
            if not count_data:
                logger.error(f"{count_users_sheet_name}シートにデータがありません")
                return False
            
            # ヘッダー行を取得
            count_headers = count_data[0]
            
            # 日付列のインデックスを取得
            date_index = 0  # 通常は最初の列
            
            # 日付に対応する行を探す
            target_row_index = None
            for i, row in enumerate(count_data):
                if row and row[date_index] == today:
                    target_row_index = i
                    break
            
            if target_row_index is None:
                logger.warning(f"{count_users_sheet_name}シートに日付 {today} の行が見つかりません。新しい行を追加します。")
                # 新しい行を追加する
                try:
                    # 空の行を作成
                    new_row = [""] * len(count_data[0])
                    # 日付を設定
                    new_row[date_index] = today
                    
                    # 最初の空行を探す
                    empty_row_index = None
                    for i, row in enumerate(count_data):
                        if not row or all(cell == "" for cell in row):
                            empty_row_index = i
                            break
                    
                    if empty_row_index is not None:
                        # 空行を更新
                        target_row_index = empty_row_index
                        logger.info(f"空行を見つけました: {empty_row_index + 1}行目")
                        # 行全体を更新
                        count_worksheet.update(f"A{empty_row_index + 1}:{_custom_col_to_a1(len(new_row))}{empty_row_index + 1}", [new_row])
                    else:
                        # 新しい行を追加
                        count_worksheet.append_row(new_row)
                        target_row_index = len(count_data)
                        logger.info(f"新しい行を追加しました: {target_row_index + 1}行目")
                    
                    # 更新後のデータを再取得
                    count_data = count_worksheet.get_all_values()
                except Exception as e:
                    logger.error(f"新しい行の追加に失敗しました: {str(e)}")
                    return False
            
            logger.info(f"対象行のインデックス: {target_row_index}")
            
            # COUNT_USERSシートのヘッダー行を取得
            count_headers = count_data[0]
            
            # ヘッダーから各カラムの位置を特定
            sections = ["全体", "LINE", "自社サイト(応募後架電)", "自社サイト(ダイレクトコミュニケーション)", "LINE（エージェント）"]
            section_columns = {}
            phases = list(phase_counts["全体"].keys())  # フェーズのリスト
            
            # 各セクションのカラムマッピングを初期化
            for section in sections:
                section_columns[section] = {}
                for phase in phases:
                    section_columns[section][phase] = None
            
            # セクションの開始位置を特定
            section_starts = {}
            for section in sections:
                for j, header in enumerate(count_headers):
                    if header == section:
                        section_starts[section] = j
                        logger.info(f"セクション '{section}' の開始位置: 列{j+1} ({_custom_col_to_a1(j+1)})")
                        break
            
            # 各セクションごとにフェーズカラムを割り当て
            # セクションの開始位置から順に各フェーズに割り当てる
            for section in sections:
                if section in section_starts:
                    start_col = section_starts[section] + 1  # セクション名の次の列から開始
                    
                    # Date列を除外（通常は最初の列）
                    if section == "全体" and count_headers[0] == "Date":
                        start_col = 1  # 全体セクションはDate列の次から
                    
                    # 実際のヘッダー内容を確認して列のオフセットを決定
                    # この部分をヘッダー行の実際の内容に基づいて解析するよう修正
                    logger.info(f"セクション '{section}' のヘッダー解析:")
                    
                    # 各フェーズに対応する列を特定
                    phase_columns = {}
                    
                    # フェーズヘッダー行からフェーズカラムを特定する
                    # 実際のヘッダー行に基づいて割り当てを行う（1行目が空白でもできるようにする）
                    if len(count_data) > 1:
                        sub_headers = count_data[1] if len(count_data) > 1 else []
                        logger.info(f"サブヘッダー行の内容: {sub_headers}")
                        
                        # セクション開始位置からスキャンして各フェーズのカラムを特定
                        scan_start = section_starts[section]
                        scan_end = section_starts.get(sections[sections.index(section) + 1], len(count_headers)) if sections.index(section) < len(sections) - 1 else len(count_headers)
                        
                        logger.info(f"セクション '{section}' のスキャン範囲: 列{scan_start+1}～列{scan_end}")
                        
                        # サブヘッダー行（2行目）からフェーズ名を探す
                        for i in range(scan_start, scan_end):
                            if i < len(sub_headers) and sub_headers[i] and any(phase in sub_headers[i] for phase in phases):
                                for phase in phases:
                                    if phase in sub_headers[i]:
                                        section_columns[section][phase] = i
                                        logger.info(f"サブヘッダーからフェーズ '{phase}' を列{i+1} ({_custom_col_to_a1(i+1)})に割り当て")
                    
                    # サブヘッダーで見つからなかった場合、従来の方法で順番に割り当て
                    if all(section_columns[section][phase] is None for phase in phases):
                        logger.info(f"サブヘッダーからフェーズが見つかりませんでした。順番に割り当てます。")
                        for i, phase in enumerate(phases):
                            col_index = start_col + i
                            # 前日差分カラムと合計カラムを避ける
                            if col_index < len(count_headers) and \
                               "前日差分" not in count_headers[col_index] and \
                               "合計" not in count_headers[col_index]: # 「合計」をスキップする条件を追加
                                section_columns[section][phase] = col_index
                                logger.info(f"デフォルト割り当て: セクション '{section}' のフェーズ '{phase}' を列{col_index+1} ({_custom_col_to_a1(col_index+1)})に割り当て")
                    
                    # 特別なケース：LINEセクションの相談前×推薦前(新規エントリー)が特定の場所にある場合
                    if section == "LINE":
                        # 明示的にデバッグ
                        target_phase = "相談前×推薦前(新規エントリー)"
                        logger.info(f"LINE セクションの {target_phase} の特別チェック")
                        # 実際の列位置を特定するためのデバッグ情報
                        if len(count_data) > 1:
                            for i in range(max(0, section_starts["LINE"] - 1), min(len(count_data[1]), section_starts["LINE"] + 10)):
                                if i < len(count_data[1]):
                                    logger.info(f"LINE周辺のカラム内容 列{i+1}({_custom_col_to_a1(i+1)}): '{count_data[1][i] if i < len(count_data[1]) else '範囲外'}'")
            
            # デバッグ情報：ヘッダー情報とカラム位置のマッピングをログに記録
            logger.info(f"ヘッダー行の内容: {count_headers}")
            for section in sections:
                logger.info(f"セクション '{section}' のカラムマッピング:")
                for phase, col_index in section_columns[section].items():
                    logger.info(f"  - {phase}: {'None' if col_index is None else f'列{col_index+1} ({_custom_col_to_a1(col_index + 1)})'}")
            
            # 更新するセルの準備
            update_cells = []
            
            # 「全体」セクションの更新
            for phase, count in phase_counts["全体"].items():
                if phase in section_columns["全体"] and section_columns["全体"][phase] is not None:
                    col_index = section_columns["全体"][phase]
                    update_cells.append({
                        'row': target_row_index + 1,
                        'col': col_index + 1,
                        'value': count
                    })
                    # デバッグ: 相談前×推薦前(新規エントリー)の更新について詳細出力
                    if "新規エントリー" in phase:
                        logger.info(f"「{phase}」を列{col_index+1} ({_custom_col_to_a1(col_index + 1)})に値={count}で更新します")
            
            # 各登録経路セクションも更新
            for route in defined_routes:
                for phase, count in phase_counts[route].items():
                    if route in section_columns and phase in section_columns[route] and section_columns[route][phase] is not None:
                        col_index = section_columns[route][phase]
                        update_cells.append({
                            'row': target_row_index + 1,
                            'col': col_index + 1,
                            'value': count
                        })
                        # デバッグ: 相談前×推薦前(新規エントリー)の更新について詳細出力
                        if "新規エントリー" in phase:
                            logger.info(f"[{route}] 「{phase}」を列{col_index+1} ({_custom_col_to_a1(col_index + 1)})に値={count}で更新します")
            
            # デバッグ情報：更新するセルの情報をログに記録
            logger.info(f"更新予定のセル数: {len(update_cells)}")
            for i, cell in enumerate(update_cells[:10]):  # 最初の10個だけ表示
                logger.info(f"  セル{i+1}: 行{cell['row']}、列{cell['col']} ({_custom_col_to_a1(cell['col'])}{cell['row']}) = {cell['value']}")
            if len(update_cells) > 10:
                logger.info(f"  ...他{len(update_cells)-10}個のセル")
            
            # ユニコード文字列のデバッグ
            logger.info("=== フェーズ名のデバッグ ===")
            for phase in phase_counts["全体"]:
                logger.info(f"フェーズ名: '{phase}', バイト表現: {phase.encode('utf-8')}")
                for section in sections:
                    if section in section_columns and phase in section_columns[section]:
                        col_index = section_columns[section][phase]
                        logger.info(f"  セクション '{section}' のマッピング: {'None' if col_index is None else f'列{col_index+1} ({_custom_col_to_a1(col_index + 1)})'}")
            
            # セルを一括更新
            if update_cells:
                # バッチ更新用のデータを準備
                batch_data = []
                for cell_data in update_cells:
                    batch_data.append({
                        'range': f"{_custom_col_to_a1(cell_data['col'])}{cell_data['row']}",
                        'values': [[cell_data['value']]]
                    })
                
                # バッチ更新を実行
                count_worksheet.batch_update(batch_data, value_input_option='RAW')  # RAWオプションで数値を保持
                logger.info(f"{len(update_cells)}個のセルを数値として一括更新しました")
            else:
                logger.warning("更新するセルがありません")
            
            logger.info("✅ フェーズ別ユーザー数の集計処理が完了しました")
            return True
            
        except Exception as e:
            # エラー通知
            context = {
                "処理": "フェーズ別ユーザー集計",
                "集計日": str(aggregation_date) if aggregation_date else "指定なし",
            }
            self._notify_error(f"フェーズ別ユーザー集計に失敗しました", e, context)
            return False
    
    def aggregate_entry_process(self) -> bool:
        """
        ENTRYPROCESSシートのデータを集計して、LIST_ENTRYPROCESSシート（data_ep）に
        日付ごとの選考プロセスデータを記録する
        
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== 選考プロセスデータの集計処理を開始します ===")
            
            # settings.iniからシート名を取得
            try:
                # 設定ファイルから各シート名を取得
                entryprocess_sheet_name = env.get_config_value('SHEET_NAMES', 'ENTRYPROCESS').strip('"\'')
                list_entryprocess_sheet_name = env.get_config_value('SHEET_NAMES', 'LIST_ENTRYPROCESS').strip('"\'')

                logger.info(f"設定ファイルのシート名:")
                logger.info(f"  ENTRYPROCESS: '{entryprocess_sheet_name}'")
                logger.info(f"  LIST_ENTRYPROCESS: '{list_entryprocess_sheet_name}'")
            except Exception as e:
                logger.error(f"設定ファイルからのシート名取得に失敗: {str(e)}")
                return False
            
            # 現在の日付を取得 (yyyy/mm/dd形式)
            today = datetime.datetime.now().strftime("%Y/%m/%d")
            logger.info(f"集計日: {today}")
            
            # ENTRYPROCESSシートからデータを取得
            entryprocess_worksheet = self.spreadsheet_manager.get_worksheet(entryprocess_sheet_name)
            entryprocess_data = entryprocess_worksheet.get_all_values()
            
            if not entryprocess_data:
                logger.error(f"{entryprocess_sheet_name}シートにデータがありません")
                return False
            
            # ヘッダー行を取得
            headers = entryprocess_data[0]
            logger.info(f"ENTRYPROCESSシートのヘッダー行: {headers}")
            
            # 必要なカラムのインデックスを取得
            required_columns = {
                '求職者ID': None,
                '求職者名': None,   # 新しいフォーマットでは「求職者名」が追加
                '企業コード': None,
                '企業名': None,
                '選考プロセス': None,
                '担当CA': None,
                '登録経路': None,    # 新しいフォーマットでは「登録経路」が追加
                '企業 ID': None,     # 新しいフォーマットでは「企業 ID」が追加
                '選考プロセス日付': None, # 新しいフォーマットでは「選考プロセス日付」が追加
                '選考プロセスメモ': None, # 新しいフォーマットでは「選考プロセスメモ」が追加
                '終了フラグ': None,    # 新しいフォーマットでは「終了フラグ」が追加
                '終了理由': None,     # 新しいフォーマットでは「終了理由」が追加
            }
            
            # 名前関連のカラムのインデックスを取得（新しいフォーマットでも後方互換性のために保持）
            name_columns = {
                '性名': None,
                '名前': None,
                '性名（カナ）': None,  # 新しいフォーマットでは「性名（カナ）」が追加
                '名前（カナ）': None,  # 新しいフォーマットでは「名前（カナ）」が追加
                '年齢': None,         # 新しいフォーマットでは「年齢」が追加
                '生年月日(年齢)': None # 新しいフォーマットでは「生年月日(年齢)」が追加
            }
            
            for i, header in enumerate(headers):
                if header in required_columns:
                    required_columns[header] = i
                if header in name_columns:
                    name_columns[header] = i
            
            # 必要なカラムが存在するか確認
            # 企業コードは必須、他のカラムは少なくとも求職者名か(性名+名前)のどちらかが必要
            essential_columns = ['企業コード', '企業名', '選考プロセス', '担当CA']
            missing_columns = [col for col in essential_columns if required_columns[col] is None]
            
            if missing_columns:
                logger.error(f"ENTRYPROCESSシートに必要なカラムが見つかりません: {', '.join(missing_columns)}")
                return False
            
            # 名前関連のカラムチェック - 求職者名か(性名+名前)のどちらかが必要
            has_name_info = (required_columns['求職者名'] is not None) or (name_columns['性名'] is not None and name_columns['名前'] is not None)
            if not has_name_info:
                logger.error("ENTRYPROCESSシートに名前情報が見つかりません。「求職者名」または「性名」と「名前」が必要です。")
                return False
            
            logger.info(f"必要なカラムのインデックス: {required_columns}")
            logger.info(f"名前関連カラムのインデックス: {name_columns}")
            
            # データを処理して集計データを作成
            aggregated_data = []
            skipped_count = 0
            for row in entryprocess_data[1:]:  # ヘッダー行をスキップ
                if len(row) > max(filter(None, [required_columns[col] for col in essential_columns])):
                    # 企業コードの有無をチェック
                    has_company_code = False
                    if (required_columns['企業コード'] is not None and 
                        required_columns['企業コード'] < len(row) and 
                        row[required_columns['企業コード']].strip()):
                        has_company_code = True
                    
                    # 企業コードがない場合はスキップ
                    if not has_company_code:
                        skipped_count += 1
                        continue
                    
                    # 新しい行を作成
                    new_row = [today]  # Date列に今日の日付を設定
                    
                    # entryprocess_allシートの全カラムをそのままの順序で追加
                    # （Dateは既に追加済みなので、他のカラムを順に追加）
                    for i, value in enumerate(row):
                        new_row.append(value)
                    
                    # 集計データに追加
                    aggregated_data.append(new_row)
            
            if skipped_count > 0:
                logger.info(f"企業コードがないため {skipped_count}行をスキップしました")
            
            if not aggregated_data:
                logger.warning("選考プロセスのデータが見つかりませんでした")
                return True  # データがなくても成功と見なす
            
            # 重複データの処理
            # 重複キー: 求職者ID、選考プロセス、選考プロセス日付、企業コード、企業名
            unique_data = {}
            duplicate_count = 0
            
            # 重複チェックに使用するインデックス
            # インデックスを調整（Date列が先頭に追加されているため、元のインデックス+1）
            key_indices = [required_columns['求職者ID'] + 1]  # 求職者ID
            
            # 企業コードと企業名のインデックスを追加
            if required_columns['選考プロセス'] is not None:
                key_indices.append(required_columns['選考プロセス'] + 1)
            if required_columns['選考プロセス日付'] is not None:
                key_indices.append(required_columns['選考プロセス日付'] + 1)
            if required_columns['企業コード'] is not None:
                key_indices.append(required_columns['企業コード'] + 1)
            if required_columns['企業名'] is not None:
                key_indices.append(required_columns['企業名'] + 1)
            
            logger.info(f"重複チェックに使用するインデックス: {key_indices}")
            
            for row in aggregated_data:
                # キーとなる値を組み合わせてユニークキーを作成
                unique_key_values = [str(row[i]) if i < len(row) else "" for i in key_indices]
                unique_key = tuple(unique_key_values)
                
                if unique_key in unique_data:
                    duplicate_count += 1
                    logger.debug(f"重複データを検出しました: {row}")
                else:
                    unique_data[unique_key] = row
            
            # 重複除去後のデータに置き換え
            aggregated_data = list(unique_data.values())
            
            if duplicate_count > 0:
                logger.info(f"重複データを {duplicate_count}件 検出し、統合しました")
            
            logger.info(f"集計対象データ: {len(aggregated_data)}行")
            
            # 設定ファイルのシート名を使用してデータを記録するシートを取得
            list_ep_worksheet = self.spreadsheet_manager.get_worksheet(list_entryprocess_sheet_name)
            logger.info(f"シート '{list_entryprocess_sheet_name}' を使用してデータを集計します")
            
            # 取得したワークシートを使用
            list_ep_data = list_ep_worksheet.get_all_values()
            
            if not list_ep_data:
                logger.error(f"{list_entryprocess_sheet_name}シートにデータがありません")
                return False
            
            # ヘッダー行を確認
            # 実際のヘッダー行の構造を予測: Dateの後に元のヘッダーが続く
            expected_headers = ['Date']
            expected_headers.extend(headers)  # entryprocess_allのヘッダーをそのまま使用
            
            actual_headers = list_ep_data[0] if list_ep_data else []
            
            if actual_headers != expected_headers:
                logger.warning(f"{list_entryprocess_sheet_name}シートのヘッダー行が期待と異なります。期待: {expected_headers}, 実際: {actual_headers}")
                # ヘッダー行の検証は行うが、処理は続行する
            
            # 今日のデータを検索
            today_data_exists = False
            for i, row in enumerate(list_ep_data[1:], 1):  # ヘッダー行をスキップしてインデックスを1から始める
                if row and row[0] == today:
                    today_data_exists = True
                    logger.info(f"{list_entryprocess_sheet_name}シートに既に今日の日付 ({today}) のデータが存在します。データを上書きします。")
                    # 既存データを削除
                    column_count = len(expected_headers)
                    last_column_letter = _custom_col_to_a1(column_count)
                    delete_range = f"A{i+1}:{last_column_letter}{i+len(aggregated_data)}"
                    try:
                        list_ep_worksheet.batch_clear([delete_range])
                        logger.info(f"既存データを削除しました: {delete_range}")
                    except Exception as e:
                        logger.error(f"既存データの削除に失敗しました: {str(e)}")
                        return False
                    break
            
            # データを追加する位置を決定
            start_row = 1  # デフォルト値
            
            if not today_data_exists:
                # 最初の空行を探す
                empty_row_index = None
                for i, row in enumerate(list_ep_data[1:], 1):  # ヘッダー行をスキップしてインデックスを1から始める
                    if not row or all(cell == "" for cell in row):
                        empty_row_index = i
                        break
                
                if empty_row_index is not None:
                    # 空行が見つかった場合、その位置に追加
                    start_row = empty_row_index + 1  # 1-indexed
                    logger.info(f"空行が見つかりました: {start_row}行目から追加します")
                else:
                    # ワークシートの最後に追加
                    start_row = len(list_ep_data) + 1  # 1-indexed
                    logger.info(f"ワークシートの最後: {start_row}行目から追加します")
            else:
                # 削除した行と同じ位置に追加
                for i, row in enumerate(list_ep_data[1:], 1):  # ヘッダー行をスキップしてインデックスを1から始める
                    if row and row[0] == today:
                        start_row = i + 1  # 1-indexed
                        break
                else:
                    # 見つからない場合は最後に追加
                    start_row = len(list_ep_data) + 1  # 1-indexed
                logger.info(f"今日のデータを上書き: {start_row}行目から追加します")
            
            # データを一括更新
            column_count = len(expected_headers)
            last_column_letter = _custom_col_to_a1(column_count)
            update_range = f"A{start_row}:{last_column_letter}{start_row + len(aggregated_data) - 1}"
            
            try:
                # シートのサイズを確認
                current_rows = list_ep_worksheet.row_count
                current_cols = list_ep_worksheet.col_count
                
                # 必要な行数・列数を計算
                needed_rows = start_row + len(aggregated_data) - 1
                needed_cols = column_count
                
                # 必要に応じてシートのサイズを拡張
                if needed_rows > current_rows:
                    list_ep_worksheet.add_rows(needed_rows - current_rows)
                    logger.info(f"シートの行数を拡張しました: {current_rows} → {needed_rows}")
                
                if needed_cols > current_cols:
                    list_ep_worksheet.add_cols(needed_cols - current_cols)
                    logger.info(f"シートの列数を拡張しました: {current_cols} → {needed_cols}")
                
                # データを更新
                list_ep_worksheet.update(values=aggregated_data, range_name=update_range)
                logger.info(f"データを更新しました: {update_range}, {len(aggregated_data)}行")
            except Exception as e:
                logger.error(f"データの更新に失敗しました: {str(e)}")
                return False
            
            logger.info("✅ 選考プロセスデータの集計処理が完了しました")
            return True
            
        except Exception as e:
            logger.error(f"選考プロセスデータの集計処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return False
    
    def run_aggregation(self, aggregate_type: str) -> Tuple[bool, bool]:
        """
        指定された集計処理を実行する
        
        Args:
            aggregate_type (str): 実行する集計処理のタイプ ('users', 'entryprocess', 'both')
            
        Returns:
            Tuple[bool, bool]: (ユーザー集計の成功/失敗, エントリープロセス集計の成功/失敗)
        """
        if not self.initialize():
            logger.error("スプレッドシートマネージャーの初期化に失敗したため、処理を中止します")
            return False, False
        
        users_success = True
        entryprocess_success = True
        
        if aggregate_type in ['users', 'both']:
            logger.info("ユーザーフェーズ別集計処理を実行します")
            users_success = self.aggregate_users_by_phase()
            if users_success:
                logger.info("ユーザーフェーズ別集計処理が正常に完了しました")
            else:
                logger.error("ユーザーフェーズ別集計処理に失敗しました")
        
        if aggregate_type in ['entryprocess', 'both']:
            logger.info("選考プロセス集計処理を実行します")
            entryprocess_success = self.aggregate_entry_process()
            if entryprocess_success:
                logger.info("選考プロセス集計処理が正常に完了しました")
            else:
                logger.error("選考プロセス集計処理に失敗しました")
        
        return users_success, entryprocess_success

    def record_daily_phase_counts(self, aggregation_date: Optional[datetime.date] = None) -> bool:
        """
        USERS_ALLシートのデータを集計して、COUNT_USERSシートに
        日付ごとのフェーズ別カウントを記録する

        Args:
            aggregation_date (datetime.date, optional): 集計日. デフォルトは今日

        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== 日別フェーズ別ユーザー数の記録処理を開始します ===")
            
            if not self.spreadsheet_manager:
                logger.error("SpreadsheetManagerが初期化されていません。")
                return False

            # settings.iniからシート名を取得
            try:
                users_all_sheet_name = env.get_config_value('SHEET_NAMES', 'USERSALL').strip('"\'')
                count_users_sheet_name = env.get_config_value('SHEET_NAMES', 'COUNT_USERS').strip('"\'')
                logger.info(f"設定ファイルのシート名: USERSALL='{users_all_sheet_name}', COUNT_USERS='{count_users_sheet_name}'")
            except Exception as e:
                logger.error(f"設定ファイルからのシート名取得に失敗: {str(e)}")
                return False
            
            today_str = aggregation_date.strftime("%Y/%m/%d") if aggregation_date else datetime.datetime.now().strftime("%Y/%m/%d")
            logger.info(f"集計日: {today_str}")
            
            # ユーザーデータの取得
            users_worksheet = self.spreadsheet_manager.get_worksheet(users_all_sheet_name)
            users_data = users_worksheet.get_all_values()
            
            if not users_data or len(users_data) < 2: # ヘッダー行すらないか、ヘッダー行のみ
                logger.error(f"'{users_all_sheet_name}'シートにデータがありません（ヘッダー行を除く）。")
                return False
            
            # フェーズ列とオプションで登録経路列のインデックスを取得
            headers = users_data[0]
            try:
                phase_index = headers.index("フェーズ")
            except ValueError:
                logger.error(f"'{users_all_sheet_name}'シートに「フェーズ」列が見つかりません。ヘッダー: {headers}")
                return False
            
            try:
                route_index = headers.index("登録経路")
                logger.info(f"「登録経路」列のインデックス: {route_index}")
            except ValueError:
                route_index = -1
                logger.warning(f"'{users_all_sheet_name}'シートに「登録経路」列が見つかりません。登録経路別の集計は行いません。")
            
            logger.info(f"「フェーズ」列のインデックス: {phase_index}")
            
            # COUNT_USERSシートのデータを取得
            count_worksheet = self.spreadsheet_manager.get_worksheet(count_users_sheet_name)
            count_data = count_worksheet.get_all_values()
            if not count_data or len(count_data) < 2: # ヘッダー行とサブヘッダー行が必要
                logger.error(f"'{count_users_sheet_name}'シートにデータがありません（少なくともヘッダー行とサブヘッダー行が必要）。")
                return False

            # セクション行（1行目）とフェーズ行（2行目）を取得
            section_headers = count_data[0]
            if len(count_data) >= 2:
                phase_headers = count_data[1]
            else:
                logger.error(f"'{count_users_sheet_name}'シートにサブヘッダー行（フェーズ定義行）がありません。")
                return False
            
            logger.info(f"セクション行: {section_headers}")
            logger.info(f"フェーズ行: {phase_headers}")
            
            # セクションとフェーズの対応関係を作成
            sections = {}
            current_section = None
            
            for i, header in enumerate(section_headers):
                if header and i > 0:  # 最初の列は日付なのでスキップ
                    sections[i] = header
                    current_section = header
                elif current_section and i > 0:
                    # 空のヘッダーは前のセクションの続き
                    sections[i] = current_section
            
            logger.info(f"検出されたセクション: {sections}")
            
            # 実際のフェーズ名とその列インデックスのマッピングを作成
            phase_column_map = {}
            for i, phase_name in enumerate(phase_headers):
                if phase_name and i > 0:  # 最初の列は日付なのでスキップ
                    if "前日差分" not in phase_name and "合計" not in phase_name:  # 前日差分や合計列はスキップ
                        phase_column_map[phase_name] = i
                        section = sections.get(i, "不明")
                        logger.info(f"フェーズ '{phase_name}' をセクション '{section}' の列 {i+1} ({_custom_col_to_a1(i+1)}) に割り当て")
            
            if not phase_column_map:
                logger.error(f"'{count_users_sheet_name}'シートから有効なフェーズが見つかりませんでした。")
                return False
            
            logger.info(f"検出されたフェーズとカラム: {phase_column_map}")
            
            # フェーズごとのカウント初期化
            phase_counts = {phase: 0 for phase in phase_column_map.keys()}
            section_counts = {section: {} for section in set(sections.values())}
            
            for section_name in section_counts:
                for i, phase_name in enumerate(phase_headers):
                    if i > 0 and phase_name and sections.get(i) == section_name:
                        if "前日差分" not in phase_name and "合計" not in phase_name:
                            section_counts[section_name][phase_name] = 0
            
            logger.info(f"セクション別フェーズカウント初期値: {section_counts}")
            
            # users_allシート内の各行をフェーズ別にカウント
            unknown_phases = set()
            
            for row_num, row in enumerate(users_data[1:], start=2):  # ヘッダー行をスキップ
                if len(row) > phase_index:
                    # フェーズと登録経路を取得して正規化
                    phase_raw = row[phase_index].strip() if phase_index < len(row) else ""
                    phase = unicodedata.normalize('NFC', phase_raw).strip()
                    
                    route_raw = row[route_index].strip() if route_index >= 0 and route_index < len(row) else ""
                    route = unicodedata.normalize('NFC', route_raw).strip()
                    
                    # デバッグログ
                    if not phase in phase_counts and phase:
                        if phase not in unknown_phases:  # 同じフェーズは一度だけログに出力
                            logger.debug(f"未知のフェーズ: '{phase}' (users_allシート {row_num}行目)")
                            unknown_phases.add(phase)
                    
                    # 全体用のカウントを更新
                    if phase in phase_counts:
                        phase_counts[phase] += 1
                    
                    # セクション別のカウントを更新
                    if route and route in section_counts:
                        section_phase_counts = section_counts[route]
                        if phase in section_phase_counts:
                            section_counts[route][phase] += 1
            
            if unknown_phases:
                logger.warning(f"{len(unknown_phases)}種類の未知のフェーズがありました: {sorted(list(unknown_phases))}")
            
            logger.info(f"集計されたフェーズごとのカウント（全体）: {phase_counts}")
            logger.info(f"集計されたセクション別フェーズカウント: {section_counts}")
            
            # data_usersシート内の適切な行を探す（日付から）
            date_col_index = 0  # 日付列は通常A列（インデックス0）
            target_row_index = -1
            
            for i, row in enumerate(count_data):
                if row and len(row) > date_col_index and row[date_col_index] == today_str:
                    target_row_index = i
                    logger.info(f"日付 '{today_str}' の行が見つかりました (行 {i+1})")
                    break
            
            # 更新するセルの準備
            cells_to_update = []
            
            if target_row_index == -1:
                # 日付行が見つからない場合は新しい行を追加
                logger.info(f"'{count_users_sheet_name}'シートに日付 '{today_str}' の行が見つかりません。新しい行を追加します。")
                
                # 新しい行のデータを準備（初期値はすべて0）
                new_row = [today_str] + [0] * (max(sections.keys()) if sections else 0)
                
                # フェーズカウントを設定
                for phase, count in phase_counts.items():
                    col_index = phase_column_map.get(phase)
                    if col_index is not None and col_index < len(new_row):
                        new_row[col_index] = count
                
                try:
                    # 行を追加
                    count_worksheet.append_row(new_row, value_input_option='USER_ENTERED')
                    logger.info(f"新しい行をシートに追加しました: {new_row}")
                    return True
                except Exception as e:
                    logger.error(f"新しい行の追加に失敗しました: {e}")
                    return False
            else:
                # 既存の行を更新
                logger.info(f"既存の行 {target_row_index + 1} を更新します")
                
                # 全体のフェーズカウントを更新
                for phase, count in phase_counts.items():
                    col_index = phase_column_map.get(phase)
                    if col_index is not None:
                        cell_ref = f"{_custom_col_to_a1(col_index + 1)}{target_row_index + 1}"
                        cells_to_update.append({
                            'range': cell_ref,
                            'values': [[count]]
                        })
                        logger.debug(f"セル {cell_ref} を値 {count} で更新します（フェーズ: {phase}）")
                
                # セクション別のフェーズカウントを更新
                total_by_phase = {}  # フェーズごとの合計値を追跡
                
                # 各セクションのフェーズごとに集計
                for section_name, section_phases in section_counts.items():
                    if section_name != "全体":  # 全体セクションは後で計算するので除外
                        logger.info(f"セクション '{section_name}' のフェーズカウントを更新します")
                        for phase_name, count in section_phases.items():
                            # 合計値を集計
                            if phase_name not in total_by_phase:
                                total_by_phase[phase_name] = 0
                            total_by_phase[phase_name] += count
                            
                            # セクションとフェーズの組み合わせに対応する列を特定
                            for i, header in enumerate(phase_headers):
                                if i > 0 and header == phase_name and sections.get(i) == section_name:
                                    cell_ref = f"{_custom_col_to_a1(i + 1)}{target_row_index + 1}"
                                    cells_to_update.append({
                                        'range': cell_ref,
                                        'values': [[count]]
                                    })
                                    logger.debug(f"セル {cell_ref} を値 {count} で更新します（セクション: {section_name}, フェーズ: {phase_name}）")
                                    break
                
                # 全体セクションの更新 - すべての登録経路の合計
                for phase_name, total_count in total_by_phase.items():
                    # 全体セクションでのフェーズ列を見つける
                    for i, header in enumerate(phase_headers):
                        if i > 0 and header == phase_name and sections.get(i) == "全体":
                            cell_ref = f"{_custom_col_to_a1(i + 1)}{target_row_index + 1}"
                            cells_to_update.append({
                                'range': cell_ref,
                                'values': [[total_count]]
                            })
                            logger.info(f"セル {cell_ref} を合計値 {total_count} で更新します（全体セクション, フェーズ: {phase_name}）")
                            break
                
                # 合計列の更新
                for section_name in set(sections.values()):
                    # そのセクションの合計列を見つける
                    for i, header in enumerate(phase_headers):
                        if i > 0 and header == "合計" and sections.get(i) == section_name:
                            # そのセクションのすべてのフェーズの合計値を計算
                            section_total = 0
                            if section_name == "全体":
                                # 全体セクションの合計
                                section_total = sum(total_by_phase.values())
                            else:
                                # 各セクションの合計
                                section_total = sum(section_counts.get(section_name, {}).values())
                            
                            cell_ref = f"{_custom_col_to_a1(i + 1)}{target_row_index + 1}"
                            cells_to_update.append({
                                'range': cell_ref,
                                'values': [[section_total]]
                            })
                            logger.info(f"セル {cell_ref} を合計値 {section_total} で更新します（セクション: {section_name}, 合計列）")
                            break
                
                if cells_to_update:
                    try:
                        # セルを一括更新
                        count_worksheet.batch_update(cells_to_update, value_input_option='USER_ENTERED')
                        logger.info(f"{len(cells_to_update)}個のセルを更新しました")
                    except Exception as e:
                        logger.error(f"セルの更新に失敗しました: {e}")
                        return False
                else:
                    logger.warning("更新するセルがありませんでした")
            
            logger.info("✅ 日別フェーズ別ユーザー数の記録処理が完了しました")
            return True
            
        except Exception as e:
            logger.error(f"日別フェーズ別ユーザー数の記録処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return False