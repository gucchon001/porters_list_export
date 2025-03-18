#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
設定ファイル関連のユーティリティ関数
"""

import os
import sys
from pathlib import Path
from typing import Dict, Optional, Any

# プロジェクトのルートディレクトリをPYTHONPATHに追加
root_dir = Path(__file__).resolve().parent.parent.parent
sys.path.append(str(root_dir))

from src.utils.environment import EnvironmentUtils as env
from src.utils.logging_config import get_logger

logger = get_logger(__name__)

def get_porters_account() -> Dict[str, str]:
    """
    PORTERSアカウント情報を取得する
    
    Returns:
        Dict[str, str]: ユーザー名とパスワードを含む辞書
    """
    try:
        # 環境変数が読み込まれていない場合は読み込む
        env.load_env()
        
        # 環境変数からPORTERSのアカウント情報を取得
        username = os.environ.get('PORTERS_USERNAME', '')
        password = os.environ.get('PORTERS_PASSWORD', '')
        
        if not username or not password:
            logger.error("環境変数からPORTERSアカウント情報を取得できませんでした")
            return {'username': '', 'password': ''}
        
        return {
            'username': username,
            'password': password
        }
    except Exception as e:
        logger.error(f"PORTERSアカウント情報の取得に失敗しました: {str(e)}")
        return {'username': '', 'password': ''}

def get_spreadsheet_id() -> str:
    """
    Google SpreadsheetのスプレッドシートのスプレッドシートのスプレッドシートのスプレッドシートIDを取得する
    
    Returns:
        str: スプレッドシートID
    """
    try:
        # 環境変数が読み込まれていない場合は読み込む
        env.load_env()
        
        # 環境変数からスプレッドシートIDを取得
        spreadsheet_id = os.environ.get('SHEET_ID', '')
        
        if not spreadsheet_id:
            # 環境変数から取得できない場合は設定ファイルから取得
            spreadsheet_id = env.get_config_value('SHEET', 'ID').strip('"\'')
        
        if not spreadsheet_id:
            logger.error("スプレッドシートIDを取得できませんでした")
            return ''
        
        return spreadsheet_id
    except Exception as e:
        logger.error(f"スプレッドシートIDの取得に失敗しました: {str(e)}")
        return '' 