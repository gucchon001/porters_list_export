import os
import sys
from pathlib import Path
import logging
import time
import json
from datetime import datetime
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import traceback

from ...utils.environment import EnvironmentUtils as env
from ...utils.logging_config import get_logger
from ..common.settings import load_sheet_settings

logger = get_logger(__name__)

class AnqDataAnalysis:
    """
    アンケートDLデータの分析を行うクラス
    """
    
    def __init__(self, target_ids):
        """
        アンケートデータ分析クラス
        
        Args:
            target_ids (list): 分析対象のIDリスト
        """
        logger.info("アンケートデータ分析クラスを初期化します")
        
        # 環境設定の読み込み
        env.load_env()
        logger.info("環境変数を読み込みました")
        
        # プロジェクトのルートディレクトリを取得
        root_dir = env.get_project_root()
        
        # 出力ディレクトリの設定
        self.output_dir = os.path.join(root_dir, "output")
        os.makedirs(self.output_dir, exist_ok=True)
        logger.info(f"出力ディレクトリを作成しました: {self.output_dir}")
        
        # サービスアカウントファイルのパスを取得
        self.service_account_file = env.get_config_value("GOOGLE_SHEETS", "SERVICE_ACCOUNT_FILE")
        logger.info(f"サービスアカウントファイル: {self.service_account_file}")
        
        # スプレッドシートIDの取得
        self.spreadsheet_id = env.get_config_value("GOOGLE_SHEETS", "SPREADSHEET_ID")
        logger.info(f"スプレッドシートID: {self.spreadsheet_id}")
        
        # シート名の取得
        self.anq_data_key = env.get_config_value("SHEET_NAMES", "ANQ_DATA")
        logger.info(f"アンケートデータキー: {self.anq_data_key}")
        
        # 対象IDの設定
        self.target_ids = target_ids
        logger.info(f"対象ID: {self.target_ids}")
        
        # Google Sheets APIへの接続設定
        self.gc = None
        self.spreadsheet = None
        self.anq_data_sheet = None
        
        # データフレーム
        self.data_df = None
    
    def connect_to_spreadsheet(self):
        """
        Google Spreadsheetに接続する
        
        Returns:
            bool: 接続に成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("Google Spreadsheetへの接続を開始します")
            
            # サービスアカウントの認証情報を取得
            scope = ['https://spreadsheets.google.com/feeds',
                     'https://www.googleapis.com/auth/drive']
            credentials = Credentials.from_service_account_file(
                self.service_account_file, scopes=scope)
            
            # gspreadクライアントを作成
            self.gc = gspread.authorize(credentials)
            logger.info("✓ Google Sheets APIに接続しました")
            
            # スプレッドシートを開く
            self.spreadsheet = self.gc.open_by_key(self.spreadsheet_id)
            logger.info(f"✓ スプレッドシート '{self.spreadsheet.title}' を開きました")
            
            # アンケートDLデータシートを取得
            self.anq_data_sheet = self.spreadsheet.worksheet(self.anq_data_key)
            logger.info(f"✓ '{self.anq_data_key}' シートを取得しました")
            
            return True
            
        except Exception as e:
            logger.error(f"スプレッドシートへの接続に失敗しました: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def get_anq_data(self):
        """
        アンケートDLデータから指定されたIDのレコードを取得する
        
        Returns:
            bool: データ取得に成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info(f"{len(self.target_ids)}件のIDに対するアンケートデータを取得します")
            
            if not self.anq_data_sheet:
                logger.error("アンケートDLデータシートが取得されていません")
                return False
            
            # すべてのデータを取得
            all_values = self.anq_data_sheet.get_all_values()
            if not all_values:
                logger.error("アンケートDLデータが空です")
                return False
            
            # ヘッダー行を取得
            headers = all_values[0]
            logger.info(f"ヘッダー行: {headers}")
            
            # ID列のインデックスを取得
            id_col_index = None
            for i, header in enumerate(headers):
                if "ID" in header:
                    id_col_index = i
                    logger.info(f"ID列を見つけました: {header} (インデックス: {i})")
                    break
            
            if id_col_index is None:
                logger.error("ID列が見つかりませんでした")
                return False
            
            # 対象IDのレコードを抽出
            records = []
            for row in all_values[1:]:  # ヘッダー行をスキップ
                if row[id_col_index] in self.target_ids:
                    record = {}
                    for i, value in enumerate(row):
                        record[headers[i]] = value
                    records.append(record)
            
            logger.info(f"取得したレコード数: {len(records)}")
            
            # 取得したデータをDataFrameに変換
            self.data_df = pd.DataFrame(records)
            logger.info(f"DataFrame作成完了: {self.data_df.shape[0]}行 x {self.data_df.shape[1]}列")
            
            # DataFrameの先頭5行を表示
            logger.info("=== DataFrame先頭5行 ===")
            logger.info(self.data_df.head().to_string())
            
            return True
            
        except Exception as e:
            logger.error(f"アンケートデータの取得に失敗しました: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def save_to_csv(self, df, filename=None):
        """
        DataFrameをCSVファイルとして保存する
        
        Args:
            df (pandas.DataFrame): 保存するDataFrame
            filename (str, optional): 保存するファイル名。指定しない場合は日時から自動生成
            
        Returns:
            str or None: 保存したCSVファイルのパス。失敗した場合はNone
        """
        try:
            logger.info("アンケートデータをCSVファイルとして保存します")
            
            if filename is None:
                # 現在の日時をファイル名に使用
                now = datetime.now()
                date_str = now.strftime("%Y%m%d_%H%M%S")
                filename = f"anq_data_{date_str}.csv"
            
            csv_path = os.path.join(self.output_dir, filename)
            
            # 最新のCSVファイルへのリンク用のパス
            latest_csv_path = os.path.join(self.output_dir, "anq_data_latest.csv")
            
            try:
                # cp932エンコーディングで保存を試みる
                df.to_csv(csv_path, encoding='cp932', index=False)
                logger.info(f"✓ CSVファイルをcp932エンコーディングで保存しました: {csv_path}")
            except UnicodeEncodeError:
                # cp932で保存できない文字が含まれている場合はUTF-8で保存
                logger.warning("cp932エンコーディングでの保存に失敗しました。UTF-8で保存を試みます。")
                df.to_csv(csv_path, encoding='utf-8', index=False)
                logger.info(f"✓ CSVファイルをUTF-8エンコーディングで保存しました: {csv_path}")
            
            # 最新のCSVファイルへのシンボリックリンクを作成（Windowsの場合はコピー）
            if os.path.exists(latest_csv_path):
                os.remove(latest_csv_path)
            
            if os.name == "nt":  # Windows
                import shutil
                shutil.copy2(csv_path, latest_csv_path)
                logger.info(f"最新のCSVファイルをコピーしました: {latest_csv_path}")
            else:  # Linux/Mac
                os.symlink(csv_path, latest_csv_path)
                logger.info(f"最新のCSVファイルへのシンボリックリンクを作成しました: {latest_csv_path}")
            
            # 環境変数にCSVファイルパスを設定
            env.set_config_value("PORTERS", "IMPORT_CSV_PATH", "output/anq_data_latest.csv")
            
            return csv_path
            
        except Exception as e:
            logger.error(f"CSVファイルの保存に失敗しました: {str(e)}")
            logger.error(traceback.format_exc())
            return None
    
    def run(self):
        """
        アンケートデータ分析を実行する
        
        Returns:
            bool: 処理に成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info("=== アンケートデータ分析を開始します ===")
            
            # スプレッドシートに接続
            if not self.connect_to_spreadsheet():
                logger.error("スプレッドシートへの接続に失敗しました")
                return False
            
            # アンケートデータを取得
            if not self.get_anq_data():
                logger.error("アンケートデータの取得に失敗しました")
                return False
            
            # CSVファイルとして保存
            csv_path = self.save_to_csv(self.data_df)
            if not csv_path:
                logger.error("CSVファイルの保存に失敗しました")
                return False
            
            logger.info(f"✓ CSVファイルを保存しました: {csv_path}")
            logger.info("=== アンケートデータ分析が完了しました ===")
            return True
            
        except Exception as e:
            logger.error(f"アンケートデータ分析中にエラーが発生しました: {str(e)}")
            logger.error(traceback.format_exc())
            return False

def analyze_anq_data(target_ids):
    """
    アンケートデータ分析を実行する関数
    
    Args:
        target_ids (list): 分析対象のIDリスト
        
    Returns:
        tuple: (bool, pandas.DataFrame) 処理結果とデータフレーム。失敗した場合は (False, None)
    """
    try:
        logger.info("=" * 80)
        logger.info("アンケートデータ分析を開始します")
        logger.info("=" * 80)
        
        if not target_ids:
            logger.warning("分析対象のIDが指定されていません")
            return False, None
        
        logger.info(f"分析対象のID数: {len(target_ids)}")
        for i, id_value in enumerate(target_ids):
            logger.info(f"  ID {i+1}: {id_value}")
        
        analyzer = AnqDataAnalysis(target_ids)
        result = analyzer.run()
        
        if result:
            logger.info("✅ アンケートデータ分析が正常に完了しました")
            logger.info("=" * 80)
            return True, analyzer.data_df
        else:
            logger.error("❌ アンケートデータ分析に失敗しました")
            logger.info("=" * 80)
            return False, None
        
    except Exception as e:
        logger.error(f"予期せぬエラーが発生しました: {str(e)}")
        logger.error(traceback.format_exc())
        logger.error(f"❌ エラーが発生しました: {str(e)}")
        logger.info("=" * 80)
        return False, None 