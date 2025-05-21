#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
SpreadsheetAggregator の日別集計機能をテストするモジュール
"""

import os
import sys
import pytest
import datetime # 必要に応じて
from pathlib import Path
# from typing import Dict, List, Any # 不要になる可能性

# プロジェクトのルートディレクトリをPYTHONPATHに追加
root_dir = Path(__file__).parent.parent
sys.path.append(str(root_dir))

from src.utils.spreadsheet import SpreadsheetManager # SpreadsheetManagerはAggregator内で使われる
from src.modules.spreadsheet_aggregator import SpreadsheetAggregator # <<< SpreadsheetAggregator をインポート
from src.utils.environment import EnvironmentUtils as env
from src.utils.logging_config import get_logger

logger = get_logger(__name__)

# list_sheets_info 関数は残しても良いが、テストからは直接呼び出さないようにする
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

class TestCountUsers: # クラス名は TestSpreadsheetAggregatorFeatures など、より実態に合わせても良い
    """
    SpreadsheetAggregator の集計処理をテストするクラス
    """

    @pytest.fixture
    def spreadsheet_aggregator_instance(self): # fixture名を変更
        """
        SpreadsheetAggregator のインスタンス（初期化済み）を提供するフィクスチャ
        """
        aggregator = SpreadsheetAggregator()
        initialized = aggregator.initialize() # SpreadsheetManagerの初期化もここで行う
        if not initialized:
            pytest.fail("SpreadsheetAggregator の初期化に失敗しました。")
        return aggregator

    def test_record_daily_phase_counts_successfully(self, spreadsheet_aggregator_instance: SpreadsheetAggregator): # メソッド名を変更
        """
        SpreadsheetAggregator.record_daily_phase_counts が正常に実行されることをテストする。
        注意: このテストは実際のシートを更新する可能性があるため、
              テスト専用のスプレッドシートIDや、実行前にバックアップを取得するなどの対策を推奨します。
              または、SpreadsheetManagerをモック化して、実際の書き込みを行わないようにします。
        """
        logger.info("=== SpreadsheetAggregator.record_daily_phase_counts のテストを開始します ===")

        # 特定の日付でテストしたい場合は、ここで日付オブジェクトを作成
        # test_date = datetime.date(2023, 1, 15)
        # result = spreadsheet_aggregator_instance.record_daily_phase_counts(aggregation_date=test_date)

        # メソッドのデフォルト（今日の日付）で実行
        result = spreadsheet_aggregator_instance.record_daily_phase_counts()

        assert result is True, "SpreadsheetAggregator.record_daily_phase_counts が失敗しました"
        logger.info("SpreadsheetAggregator.record_daily_phase_counts のテストが正常に完了しました。")

    # count_users_by_phase メソッドは削除 (SpreadsheetAggregatorに移植済み)

if __name__ == "__main__":
    # この部分は、コマンドラインから直接 pytest を実行することを推奨するため、
    # 必ずしも必要ではありません。残す場合は、呼び出し方を修正する必要があります。
    logger.info("テストスクリプトを直接実行します (pytest経由での実行を推奨)。")
    
    # テストを簡易的に実行する例 (pytestのフィクスチャなどは利用できない)
    try:
        aggregator = SpreadsheetAggregator()
        if aggregator.initialize():
            logger.info("SpreadsheetAggregator を初期化しました。")
            # テストしたい日付を指定
            # test_run_date = datetime.date(2024, 5, 20) # 例
            # success = aggregator.record_daily_phase_counts(aggregation_date=test_run_date)
            success = aggregator.record_daily_phase_counts() # 今日で実行

            if success:
                print("record_daily_phase_counts の直接実行が成功しました。")
            else:
                print("record_daily_phase_counts の直接実行が失敗しました。")
        else:
            print("SpreadsheetAggregator の初期化に失敗しました。")
    except Exception as e:
        print(f"テストの直接実行中にエラーが発生しました: {e}")
        import traceback
        traceback.print_exc() 