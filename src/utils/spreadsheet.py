#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Google Spreadsheetを操作するためのユーティリティモジュール

このモジュールは、Google Spreadsheetへの接続、データの読み書き、
CSVファイルからのデータインポートなどの機能を提供します。
"""

import os
import csv
import time
import configparser
from pathlib import Path
from typing import List, Dict, Any, Optional, Union

import gspread
from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError, SpreadsheetNotFound

from src.utils.logging_config import get_logger

logger = get_logger(__name__)

class SpreadsheetManager:
    """
    Google Spreadsheetを操作するためのクラス
    
    このクラスは、Google Spreadsheetへの接続、データの読み書き、
    CSVファイルからのデータインポートなどの機能を提供します。
    """
    
    def __init__(self, credential_path: str = None):
        """
        SpreadsheetManagerの初期化
        
        Args:
            credential_path (str, optional): サービスアカウントのJSONファイルパス。
                                            指定しない場合はデフォルトのパスを使用。
        """
        self.config = self._load_config()
        self.spreadsheet_id = self.config.get('SPREADSHEET', 'SSID').strip('"\'')
        self.sheet_names = {
            'users_all': self.config.get('SHEET_NAMES', 'USERSALL').strip('"\''),
            'entryprocess_all': self.config.get('SHEET_NAMES', 'ENTRYPROCESS').strip('"\''),
            'logging': self.config.get('SHEET_NAMES', 'LOGSHEET').strip('"\'')
        }
        
        # 認証情報の設定
        if credential_path is None:
            # デフォルトのパスを使用
            base_dir = Path(__file__).resolve().parent.parent.parent
            credential_path = os.path.join(base_dir, 'config', 'boxwood-dynamo-384411-6dec80faabfc.json')
        
        self.credential_path = credential_path
        self.client = self._authenticate()
        self.spreadsheet = None
        
        logger.info(f"SpreadsheetManager initialized with spreadsheet ID: {self.spreadsheet_id}")
    
    def _load_config(self) -> configparser.ConfigParser:
        """
        設定ファイルを読み込む
        
        Returns:
            configparser.ConfigParser: 設定情報
        """
        config = configparser.ConfigParser()
        base_dir = Path(__file__).resolve().parent.parent.parent
        config_path = os.path.join(base_dir, 'config', 'settings.ini')
        
        if not os.path.exists(config_path):
            logger.error(f"Config file not found: {config_path}")
            raise FileNotFoundError(f"Config file not found: {config_path}")
        
        config.read(config_path)
        return config
    
    def _authenticate(self) -> gspread.Client:
        """
        Google APIに認証する
        
        Returns:
            gspread.Client: 認証済みのgspreadクライアント
        """
        try:
            scopes = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
            
            credentials = Credentials.from_service_account_file(
                self.credential_path, 
                scopes=scopes
            )
            
            client = gspread.authorize(credentials)
            logger.info("Successfully authenticated with Google API")
            return client
            
        except Exception as e:
            logger.error(f"Authentication failed: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            raise
    
    def open_spreadsheet(self) -> gspread.Spreadsheet:
        """
        スプレッドシートを開く
        
        Returns:
            gspread.Spreadsheet: 開いたスプレッドシート
        """
        try:
            self.spreadsheet = self.client.open_by_key(self.spreadsheet_id)
            logger.info(f"Successfully opened spreadsheet: {self.spreadsheet.title}")
            return self.spreadsheet
            
        except SpreadsheetNotFound:
            logger.error(f"Spreadsheet not found with ID: {self.spreadsheet_id}")
            raise
        except Exception as e:
            logger.error(f"Failed to open spreadsheet: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            raise
    
    def get_worksheet(self, sheet_key: str) -> gspread.Worksheet:
        """
        ワークシートを取得する
        
        Args:
            sheet_key (str): シートのキー ('users_all', 'entryprocess_all', 'logging')
            
        Returns:
            gspread.Worksheet: 取得したワークシート
        """
        if self.spreadsheet is None:
            self.open_spreadsheet()
        
        sheet_name = self.sheet_names.get(sheet_key)
        if not sheet_name:
            logger.error(f"Invalid sheet key: {sheet_key}")
            raise ValueError(f"Invalid sheet key: {sheet_key}")
        
        try:
            worksheet = self.spreadsheet.worksheet(sheet_name)
            logger.info(f"Successfully got worksheet: {sheet_name}")
            return worksheet
            
        except Exception as e:
            logger.error(f"Failed to get worksheet '{sheet_name}': {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            raise
    
    def clear_worksheet(self, sheet_key: str) -> None:
        """
        ワークシートのデータをクリアする
        
        Args:
            sheet_key (str): シートのキー ('users_all', 'entryprocess_all', 'logging')
        """
        worksheet = self.get_worksheet(sheet_key)
        
        try:
            # ヘッダー行を保持するため、2行目以降をクリア
            if worksheet.row_count > 1:
                worksheet.batch_clear(["A2:ZZ"])
                logger.info(f"Successfully cleared worksheet: {self.sheet_names.get(sheet_key)}")
            else:
                logger.warning(f"Worksheet {self.sheet_names.get(sheet_key)} has no data to clear")
                
        except Exception as e:
            logger.error(f"Failed to clear worksheet: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            raise
    
    def import_csv_to_sheet(self, csv_path: str, sheet_key: str, has_header: bool = True) -> None:
        """
        CSVファイルからデータをシートにインポートする
        
        Args:
            csv_path (str): CSVファイルのパス
            sheet_key (str): シートのキー ('users_all', 'entryprocess_all', 'logging')
            has_header (bool, optional): CSVファイルにヘッダーがあるかどうか。デフォルトはTrue。
        """
        if not os.path.exists(csv_path):
            logger.error(f"CSV file not found: {csv_path}")
            raise FileNotFoundError(f"CSV file not found: {csv_path}")
        
        worksheet = self.get_worksheet(sheet_key)
        
        try:
            # CSVファイルを読み込む（複数の文字コードを試す）
            encodings = ['utf-8-sig', 'utf-8', 'shift-jis', 'cp932']
            data = None
            used_encoding = None
            
            for encoding in encodings:
                try:
                    with open(csv_path, 'r', encoding=encoding) as f:
                        csv_reader = csv.reader(f)
                        data = list(csv_reader)
                    used_encoding = encoding
                    logger.info(f"Successfully read CSV file with encoding: {encoding}")
                    break
                except UnicodeDecodeError:
                    logger.warning(f"Failed to read CSV file with encoding: {encoding}")
                    continue
            
            if data is None:
                logger.error("Failed to read CSV file with any encoding")
                raise UnicodeError("Failed to read CSV file with any encoding")
            
            if not data:
                logger.warning(f"CSV file is empty: {csv_path}")
                return
            
            # データの行数と列数を確認
            row_count = len(data)
            col_count = max(len(row) for row in data) if data else 0
            logger.info(f"CSV data: {row_count} rows, {col_count} columns")
            
            # ヘッダー行の処理
            start_row = 1
            if has_header:
                # ヘッダー行を更新
                header = data[0]
                logger.info(f"Header row: {header}")
                worksheet.update('A1', [header])
                # データは2行目から
                data = data[1:]
                start_row = 2
            
            # データが空の場合は終了
            if not data:
                logger.warning(f"CSV file has no data rows: {csv_path}")
                return
            
            # データを一括でアップロード
            if data:
                # APIの制限に対応するため、一度に更新する行数を制限
                batch_size = 1000  # Google Sheets APIの制限に基づく適切な値
                total_batches = (len(data) + batch_size - 1) // batch_size  # 切り上げ除算
                
                for i in range(0, len(data), batch_size):
                    batch_num = i // batch_size + 1
                    logger.info(f"Uploading batch {batch_num}/{total_batches} ({len(data[i:i+batch_size])} rows)")
                    
                    batch_data = data[i:i+batch_size]
                    range_str = f'A{start_row + i}:ZZ{start_row + i + len(batch_data) - 1}'
                    
                    try:
                        worksheet.update(range_str, batch_data)
                        logger.info(f"Successfully updated range: {range_str}")
                    except Exception as e:
                        logger.error(f"Failed to update range {range_str}: {str(e)}")
                        # エラーが発生しても処理を継続
                        import traceback
                        logger.error(traceback.format_exc())
                    
                    # API制限に引っかからないよう少し待機
                    if i + batch_size < len(data):
                        time.sleep(1)
            
            logger.info(f"Successfully imported {len(data)} rows from CSV to worksheet: {self.sheet_names.get(sheet_key)}")
            
        except Exception as e:
            logger.error(f"Failed to import CSV to worksheet: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            raise
    
    def append_log(self, log_data: List[str]) -> None:
        """
        ログシートにデータを追加する
        
        Args:
            log_data (List[str]): 追加するログデータ（1行分）
        """
        worksheet = self.get_worksheet('logging')
        
        try:
            worksheet.append_row(log_data)
            logger.info(f"Successfully appended log data: {log_data}")
            
        except Exception as e:
            logger.error(f"Failed to append log data: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            # ログの追加に失敗してもプログラムは継続 