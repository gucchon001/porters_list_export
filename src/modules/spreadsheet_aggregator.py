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
    
    def aggregate_users_by_phase(self) -> bool:
        """
        USERS_ALLシートのデータを集計して、COUNT_USERSシートに
        日付ごとのフェーズ別カウントを記録する
        
        Returns:
            bool: 処理が成功した場合はTrue、失敗した場合はFalse
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
            today = datetime.datetime.now().strftime("%Y/%m/%d")
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
            for i, header in enumerate(headers):
                if header == "フェーズ":
                    phase_index = i
                    break
            
            if phase_index is None:
                logger.error(f"{users_all_sheet_name}シートに「フェーズ」列が見つかりません")
                return False
            
            logger.info(f"フェーズ列のインデックス: {phase_index}")
            
            # フェーズごとのカウントを初期化
            phase_counts = {
                "相談前×推薦前(新規エントリー)": 0,
                "相談前×推薦前(open)": 0,
                "推薦済(仮エントリー)": 0,
                "面談設定済": 0,
                "終了": 0
            }
            
            # データ行を処理してフェーズごとのカウントを集計
            for row in users_data[1:]:  # ヘッダー行をスキップ
                if len(row) > phase_index:
                    phase = row[phase_index]
                    if phase in phase_counts:
                        phase_counts[phase] += 1
                    elif phase:
                        logger.warning(f"未知のフェーズ: {phase}")
            
            logger.info(f"フェーズごとのカウント: {phase_counts}")
            
            # COUNT_USERSシートからデータを取得
            count_worksheet = self.spreadsheet_manager.get_worksheet(count_users_sheet_name)
            count_data = count_worksheet.get_all_values()
            
            if not count_data:
                logger.error(f"{count_users_sheet_name}シートにデータがありません")
                return False
            
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
            
            # フェーズ名とカラムのマッピングを作成
            phase_column_map = {}
            for phase_name in phase_counts.keys():
                for j, header in enumerate(count_headers):
                    if header == phase_name:
                        phase_column_map[phase_name] = j
                        break
            
            logger.info(f"フェーズと列のマッピング: {phase_column_map}")
            
            # 更新するセルの値を準備
            update_cells = []
            for phase_name, count in phase_counts.items():
                if phase_name in phase_column_map:
                    col_index = phase_column_map[phase_name]
                    update_cells.append({
                        'row': target_row_index + 1,
                        'col': col_index + 1,
                        'value': count  # 文字列から数値に変更
                    })
            
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
            logger.error(f"フェーズ別ユーザー数の集計処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
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
                '企業コード': None,
                '企業名': None,
                '選考プロセス': None,
                '担当CA': None
            }
            
            # 名前関連のカラムのインデックスを取得
            name_columns = {
                '性名': None,
                '名前': None
            }
            
            for i, header in enumerate(headers):
                if header in required_columns:
                    required_columns[header] = i
                if header in name_columns:
                    name_columns[header] = i
            
            # 必要なカラムが存在するか確認
            missing_columns = [col for col, idx in required_columns.items() if idx is None]
            if missing_columns:
                logger.error(f"ENTRYPROCESSシートに必要なカラムが見つかりません: {', '.join(missing_columns)}")
                return False
            
            logger.info(f"必要なカラムのインデックス: {required_columns}")
            logger.info(f"名前関連カラムのインデックス: {name_columns}")
            
            # データ行を処理して集計データを作成
            aggregated_data = []
            skipped_count = 0
            for row in entryprocess_data[1:]:  # ヘッダー行をスキップ
                if len(row) > max(required_columns.values() or [0]):
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
                    
                    # 求職者ID
                    if required_columns['求職者ID'] is not None and required_columns['求職者ID'] < len(row):
                        new_row.append(row[required_columns['求職者ID']])
                    else:
                        new_row.append("")
                    
                    # 性名と名前の列を追加
                    # 性名の列
                    if '性名' in name_columns and name_columns['性名'] is not None and name_columns['性名'] < len(row):
                        new_row.append(row[name_columns['性名']])
                    else:
                        new_row.append("")
                    
                    # 名前の列
                    if '名前' in name_columns and name_columns['名前'] is not None and name_columns['名前'] < len(row):
                        new_row.append(row[name_columns['名前']])
                    else:
                        new_row.append("")
                    
                    # 残りの列を追加
                    for col_name in ['企業コード', '企業名', '選考プロセス', '担当CA']:
                        if col_name in required_columns and required_columns[col_name] is not None and required_columns[col_name] < len(row):
                            new_row.append(row[required_columns[col_name]])
                        else:
                            new_row.append("")
                    
                    # 集計データに追加
                    aggregated_data.append(new_row)
            
            if skipped_count > 0:
                logger.info(f"企業コードがないため {skipped_count}行をスキップしました")
            
            if not aggregated_data:
                logger.warning("選考プロセスのデータが見つかりませんでした")
                return True  # データがなくても成功と見なす
            
            # 重複データの処理
            # 重複キー: 求職者ID（インデックス1）、性名（インデックス2）、名前（インデックス3）、企業コード（インデックス4）、企業名（インデックス5）、選考プロセス（インデックス6）、担当CA（インデックス7）
            unique_data = {}
            duplicate_count = 0
            
            for row in aggregated_data:
                # キーとなる値を組み合わせてユニークキーを作成
                # Dateを除いた全ての項目をキーとする
                unique_key = tuple(row[1:])
                
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
            expected_headers = ['Date', '求職者ID', '性名', '名前', '企業コード', '企業名', '選考プロセス', '担当CA']
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
                    delete_range = f"A{i+1}:H{i+len(aggregated_data)}"  # A〜H (8列)に修正
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
            update_range = f"A{start_row}:H{start_row + len(aggregated_data) - 1}"
            try:
                # シートのサイズを確認
                current_rows = list_ep_worksheet.row_count
                current_cols = list_ep_worksheet.col_count
                
                # 必要な行数・列数を計算
                needed_rows = start_row + len(aggregated_data) - 1
                needed_cols = 8  # A〜H (8列)
                
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