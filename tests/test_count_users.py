#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
スプレッドシートの集計処理をテストするモジュール

このモジュールは、USERS_ALLシートのデータを集計して、COUNT_USERSシートに
日付ごとのフェーズ別カウントを記録する処理をテストします。
"""

import os
import sys
import pytest
import datetime
from pathlib import Path
from typing import Dict, List, Any

# プロジェクトのルートディレクトリをPYTHONPATHに追加
root_dir = Path(__file__).parent.parent
sys.path.append(str(root_dir))

from src.utils.spreadsheet import SpreadsheetManager
from src.utils.environment import EnvironmentUtils as env
from src.utils.logging_config import get_logger

logger = get_logger(__name__)

def list_sheets_info(spreadsheet_manager):
    """
    スプレッドシート内の全シートの情報を取得して表示する
    
    Args:
        spreadsheet_manager (SpreadsheetManager): スプレッドシート操作用のインスタンス
    """
    try:
        # スプレッドシートのプロパティを取得
        spreadsheet = spreadsheet_manager.spreadsheet
        
        # 各シートの情報を取得
        for sheet in spreadsheet.worksheets():
            sheet_title = sheet.title
            logger.info(f"シート: '{sheet_title}'")
        
    except Exception as e:
        logger.error(f"シート情報の取得に失敗しました: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())

class TestCountUsers:
    """
    スプレッドシートの集計処理をテストするクラス
    """
    
    @pytest.fixture
    def spreadsheet_manager(self):
        """
        SpreadsheetManagerのインスタンスを提供するフィクスチャ
        """
        manager = SpreadsheetManager()
        manager.open_spreadsheet()
        return manager
    
    def test_count_users_by_phase(self, spreadsheet_manager):
        """
        USERS_ALLシートのデータを集計して、COUNT_USERSシートに
        日付ごとのフェーズ別カウントを記録するテスト
        """
        # スプレッドシート内の全シートの情報を取得して表示
        list_sheets_info(spreadsheet_manager)
        
        # 集計処理を実行
        result = self.count_users_by_phase(spreadsheet_manager)
        
        # 結果を検証
        assert result is True, "集計処理が失敗しました"
    
    def count_users_by_phase(self, spreadsheet_manager: SpreadsheetManager) -> bool:
        """
        USERS_ALLシートのデータを集計して、COUNT_USERSシートに
        日付ごとのフェーズ別カウントを記録する
        
        Args:
            spreadsheet_manager (SpreadsheetManager): SpreadsheetManagerのインスタンス
            
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
            users_worksheet = spreadsheet_manager.get_worksheet(users_all_sheet_name)
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
            count_worksheet = spreadsheet_manager.get_worksheet(count_users_sheet_name)
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

if __name__ == "__main__":
    # 直接実行された場合は集計処理を実行
    manager = SpreadsheetManager()
    manager.open_spreadsheet()
    
    test = TestCountUsers()
    result = test.count_users_by_phase(manager)
    
    if result:
        print("集計処理が正常に完了しました")
    else:
        print("集計処理に失敗しました") 