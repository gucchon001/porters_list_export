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
from typing import Dict, List, Any, Tuple

# プロジェクトのルートディレクトリをPYTHONPATHに追加
root_dir = Path(__file__).resolve().parent.parent.parent
sys.path.append(str(root_dir))

from src.utils.spreadsheet import SpreadsheetManager
from src.utils.environment import EnvironmentUtils as env
from src.utils.logging_config import get_logger

logger = get_logger(__name__)

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
            
            if not users_data:
                logger.error(f"{users_all_sheet_name}シートにデータがありません")
                return False
            
            # ヘッダー行を取得
            headers = users_data[0]
            
            # フェーズ列のインデックスを取得
            phase_index = None
            registration_route_index = None
            
            for i, header in enumerate(headers):
                if header == "フェーズ":
                    phase_index = i
                elif header == "登録経路":
                    registration_route_index = i
            
            if phase_index is None:
                logger.error(f"{users_all_sheet_name}シートに「フェーズ」列が見つかりません")
                return False
            
            logger.info(f"フェーズ列のインデックス: {phase_index}")
            
            if registration_route_index is None:
                logger.warning(f"{users_all_sheet_name}シートに「登録経路」列が見つかりません。全体のみを集計します。")
            else:
                logger.info(f"登録経路列のインデックス: {registration_route_index}")
            
            # フェーズごとのカウントを初期化
            phase_counts = {
                "全体": {
                    "相談前×推薦前(新規エントリー)": 0,  # 空白なしのキーを使用
                    "相談前×推薦前(open)": 0,
                    "推薦済(仮エントリー)": 0,
                    "面談設定済": 0,
                    "終了": 0
                }
            }
            
            # 登録経路ごとのカウントも初期化
            registration_routes = ["LINE", "自社サイト(応募後架電)", "自社サイト(ダイレクトコミュニケーション)"]
            for route in registration_routes:
                phase_counts[route] = {
                    "相談前×推薦前(新規エントリー)": 0,  # 空白なしのキーを使用
                    "相談前×推薦前(open)": 0,
                    "推薦済(仮エントリー)": 0,
                    "面談設定済": 0,
                    "終了": 0
                }
            
            # データ行を処理してフェーズごとのカウントを集計
            for row in users_data[1:]:  # ヘッダー行をスキップ
                if len(row) > phase_index:
                    phase = row[phase_index].strip()  # 前後の空白を除去
                    
                    # デバッグ: 特定のフェーズに関する詳細情報
                    if "新規エントリー" in phase:
                        logger.info(f"新規エントリー行検出: '{phase}', len={len(phase)}, code points={[ord(c) for c in phase]}")
                    
                    # 全体のカウント
                    for key in phase_counts["全体"].keys():
                        if phase == key or phase == " " + key:  # 空白ありとなしの両方に対応
                            phase_counts["全体"][key] += 1
                            # デバッグ: マッチした場合の詳細
                            if "新規エントリー" in key:
                                logger.info(f"マッチしたキー: '{key}' と フェーズ: '{phase}'")
                            break
                    
                    # 登録経路ごとのカウント
                    if registration_route_index is not None and len(row) > registration_route_index:
                        route = row[registration_route_index].strip()  # 前後の空白を除去
                        if route in registration_routes:
                            for key in phase_counts[route].keys():
                                if phase == key or phase == " " + key:  # 空白ありとなしの両方に対応
                                    phase_counts[route][key] += 1
                                    break
            
            logger.info(f"全体のフェーズごとのカウント: {phase_counts['全体']}")
            for route in registration_routes:
                logger.info(f"{route}のフェーズごとのカウント: {phase_counts[route]}")
            
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
                        count_worksheet.update(f"A{empty_row_index + 1}:{chr(65 + len(new_row) - 1)}{empty_row_index + 1}", [new_row])
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
            sections = ["全体", "LINE", "自社サイト(応募後架電)", "自社サイト(ダイレクトコミュニケーション)"]
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
                        logger.info(f"セクション '{section}' の開始位置: 列{j+1} ({chr(65+j)})")
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
                                        logger.info(f"サブヘッダーからフェーズ '{phase}' を列{i+1} ({chr(65+i)})に割り当て")
                    
                    # サブヘッダーで見つからなかった場合、従来の方法で順番に割り当て
                    if all(section_columns[section][phase] is None for phase in phases):
                        logger.info(f"サブヘッダーからフェーズが見つかりませんでした。順番に割り当てます。")
                        for i, phase in enumerate(phases):
                            col_index = start_col + i
                            # 前日差分カラムを避ける
                            if col_index < len(count_headers) and "前日差分" not in count_headers[col_index]:
                                section_columns[section][phase] = col_index
                                logger.info(f"デフォルト割り当て: セクション '{section}' のフェーズ '{phase}' を列{col_index+1} ({chr(65+col_index)})に割り当て")
                    
                    # 特別なケース：LINEセクションの相談前×推薦前(新規エントリー)が特定の場所にある場合
                    if section == "LINE":
                        # 明示的にデバッグ
                        target_phase = "相談前×推薦前(新規エントリー)"
                        logger.info(f"LINE セクションの {target_phase} の特別チェック")
                        # 実際の列位置を特定するためのデバッグ情報
                        if len(count_data) > 1:
                            for i in range(max(0, section_starts["LINE"] - 1), min(len(count_data[1]), section_starts["LINE"] + 10)):
                                if i < len(count_data[1]):
                                    logger.info(f"LINE周辺のカラム内容 列{i+1}({chr(65+i)}): '{count_data[1][i] if i < len(count_data[1]) else '範囲外'}'")
            
            # デバッグ情報：ヘッダー情報とカラム位置のマッピングをログに記録
            logger.info(f"ヘッダー行の内容: {count_headers}")
            for section in sections:
                logger.info(f"セクション '{section}' のカラムマッピング:")
                for phase, col_index in section_columns[section].items():
                    logger.info(f"  - {phase}: {'None' if col_index is None else f'列{col_index+1} ({chr(65+col_index)})'}")
            
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
                        logger.info(f"「{phase}」を列{col_index+1} ({chr(65+col_index)})に値={count}で更新します")
            
            # 各登録経路セクションも更新
            for route in registration_routes:
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
                            logger.info(f"[{route}] 「{phase}」を列{col_index+1} ({chr(65+col_index)})に値={count}で更新します")
            
            # デバッグ情報：更新するセルの情報をログに記録
            logger.info(f"更新予定のセル数: {len(update_cells)}")
            for i, cell in enumerate(update_cells[:10]):  # 最初の10個だけ表示
                logger.info(f"  セル{i+1}: 行{cell['row']}、列{cell['col']} ({chr(64+cell['col'])}{cell['row']}) = {cell['value']}")
            if len(update_cells) > 10:
                logger.info(f"  ...他{len(update_cells)-10}個のセル")
            
            # ユニコード文字列のデバッグ
            logger.info("=== フェーズ名のデバッグ ===")
            for phase in phase_counts["全体"]:
                logger.info(f"フェーズ名: '{phase}', バイト表現: {phase.encode('utf-8')}")
                for section in sections:
                    if section in section_columns and phase in section_columns[section]:
                        col_index = section_columns[section][phase]
                        logger.info(f"  セクション '{section}' のマッピング: {'None' if col_index is None else f'列{col_index+1} ({chr(65+col_index)})'}")
            
            # セルを一括更新
            if update_cells:
                # バッチ更新用のデータを準備
                batch_data = []
                for cell_data in update_cells:
                    batch_data.append({
                        'range': f"{chr(64 + cell_data['col'])}{cell_data['row']}",
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
                    last_column_letter = chr(64 + column_count) if column_count <= 26 else chr(64 + column_count // 26) + chr(64 + column_count % 26)
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
            last_column_letter = chr(64 + column_count) if column_count <= 26 else chr(64 + column_count // 26) + chr(64 + column_count % 26)
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