#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
スプレッドシート集計処理テスト実行スクリプト
"""

import os
import sys
from pathlib import Path

# プロジェクトのルートディレクトリをPYTHONPATHに追加
root_dir = Path(__file__).resolve().parent
sys.path.append(str(root_dir))

from src.utils.logging_config import get_logger
from src.modules.spreadsheet_aggregator import SpreadsheetAggregator

logger = get_logger(__name__)

def main():
    """
    データユーザー集計処理を実行する
    """
    logger.info("=== データユーザー集計処理テスト開始 ===")
    
    # SpreadsheetAggregatorインスタンスを作成
    aggregator = SpreadsheetAggregator()
    
    # ユーザー集計のみを実行
    users_success, _ = aggregator.run_aggregation('users')
    
    if users_success:
        logger.info("✅ ユーザー集計処理が正常に完了しました")
    else:
        logger.error("❌ ユーザー集計処理が失敗しました")
    
    logger.info("=== データユーザー集計処理テスト完了 ===")

if __name__ == "__main__":
    main() 