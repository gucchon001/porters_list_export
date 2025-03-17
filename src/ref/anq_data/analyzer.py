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
from ..common.spreadsheet import get_spreadsheet_connection

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
        self.service_account_file = env.get_service_account_file()
        logger.info(f"サービスアカウントファイル: {self.service_account_file}")
        
        # スプレッドシートIDの取得
        self.spreadsheet_id = env.get_config_value("SPREADSHEET", "SSID")
        logger.info(f"スプレッドシートID: {self.spreadsheet_id}")
        
        # シート名の取得（キーのみ保存）
        self.anq_data_key = env.get_config_value("SHEET_NAMES", "ANQ_DATA").strip('"')
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
            
            # 共通のスプレッドシート接続を使用
            self.spreadsheet = get_spreadsheet_connection()
            if not self.spreadsheet:
                logger.error("スプレッドシートへの接続に失敗しました")
                return False
            
            logger.info(f"✓ スプレッドシート '{self.spreadsheet.title}' に接続しました")
            
            # settings.pyを使用してシート名を取得
            # アンケートデータシートは他のシートと一緒に取得されないため、別途処理
            try:
                # settingsシートを取得
                settings_sheet = self.spreadsheet.worksheet('settings')
                logger.info(f"✓ settingsシートを取得しました")
                
                # 設定値を取得
                data = settings_sheet.get_all_values()
                logger.info(f"settingsシートのデータ行数: {len(data)}")
                
                # 設定値を辞書に変換（ヘッダー行をスキップ）
                settings = {row[0]: row[1] for row in data[1:] if len(row) >= 2}
                logger.info(f"設定マッピング: {settings}")
                
                # アンケートデータシートの実際の名前を取得
                if self.anq_data_key in settings:
                    anq_sheet_name = settings[self.anq_data_key]
                    logger.info(f"✓ アンケートデータ実際のシート名: {anq_sheet_name}")
                else:
                    logger.warning(f"❌ 設定キー '{self.anq_data_key}' が settingsシートに見つかりません")
                    logger.warning(f"利用可能な設定: {settings}")
                    logger.warning(f"キー '{self.anq_data_key}' をそのままシート名として使用します")
                    anq_sheet_name = self.anq_data_key
            
            except Exception as e:
                logger.warning(f"設定の読み込み中にエラーが発生しました: {str(e)}")
                logger.warning(f"キー '{self.anq_data_key}' をそのままシート名として使用します")
                anq_sheet_name = self.anq_data_key
            
            # アンケートDLデータシートを取得
            try:
                self.anq_data_sheet = self.spreadsheet.worksheet(anq_sheet_name)
                logger.info(f"✓ '{anq_sheet_name}' シートを取得しました")
            except gspread.exceptions.WorksheetNotFound:
                logger.error(f"❌ シート '{anq_sheet_name}' が見つかりません")
                # 利用可能なシート一覧を表示
                all_worksheets = self.spreadsheet.worksheets()
                logger.info("利用可能なシート一覧:")
                for i, worksheet in enumerate(all_worksheets):
                    logger.info(f"  {i+1}. {worksheet.title}")
                return False
            
            return True
            
        except Exception as e:
            logger.error(f"スプレッドシートへの接続に失敗しました: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def get_anq_data(self):
        """
        アンケートデータを取得する
        
        Returns:
            bool: 取得に成功した場合はTrue、失敗した場合はFalse
        """
        try:
            logger.info(f"{len(self.target_ids)}件のIDに対するアンケートデータを取得します")
            logger.info(f"検索対象ID: {self.target_ids}")
            
            # スプレッドシートからデータを取得
            all_data = self.anq_data_sheet.get_all_values()
            logger.info(f"スプレッドシートから取得した全データ行数: {len(all_data)}")
            
            if not all_data or len(all_data) <= 1:
                logger.error("スプレッドシートからデータを取得できませんでした")
                return False
            
            # ヘッダー行を取得
            headers = all_data[0]
            logger.info(f"ヘッダー行: {headers}")
            
            # "回答者ID"列のインデックスを特定
            id_column_name = "回答者ID"
            
            if id_column_name not in headers:
                logger.error(f"'{id_column_name}'列が見つかりません。ヘッダー: {headers}")
                return False
            
            id_column_index = headers.index(id_column_name)
            logger.info(f"'{id_column_name}'列を見つけました (インデックス: {id_column_index})")
            
            # 対象IDに一致する行を抽出
            matching_rows = []
            for row in all_data[1:]:  # ヘッダー行をスキップ
                if id_column_index < len(row):
                    row_id = row[id_column_index]
                    # IDが文字列として保存されている可能性があるため、文字列に変換して比較
                    if str(row_id) in [str(target_id) for target_id in self.target_ids]:
                        matching_rows.append(row)
                        logger.info(f"一致するIDを持つ行を見つけました: {row_id}")
            
            logger.info(f"取得したレコード数: {len(matching_rows)}")
            
            if not matching_rows:
                # 最初の10行のIDを表示して確認
                sample_ids = [row[id_column_index] if id_column_index < len(row) else "N/A" for row in all_data[1:11]]
                logger.warning(f"指定されたID {self.target_ids} に一致するデータが見つかりませんでした")
                logger.warning(f"スプレッドシートの最初の10行のID: {sample_ids}")
                # 空のDataFrameを作成
                self.data_df = pd.DataFrame(columns=headers)
            else:
                # DataFrameを作成
                self.data_df = pd.DataFrame(matching_rows, columns=headers)
            
            logger.info(f"DataFrame作成完了: {self.data_df.shape[0]}行 x {self.data_df.shape[1]}列")
            logger.info("=== DataFrame先頭5行 ===")
            logger.info(self.data_df.head())
            
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
            
            # 環境変数にCSVファイルパスを設定する代わりに、相対パスを記録
            logger.info("CSVファイルパス 'output/anq_data_latest.csv' を後続の処理で使用します")
            
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
    アンケートデータを分析してCSVファイルに出力する
    
    Args:
        target_ids (list): 分析対象のIDリスト
        
    Returns:
        tuple: (成功したかどうか, CSVファイルのパス)
    """
    try:
        logger.info(f"{len(target_ids)}件のIDに対するアンケートデータ分析を開始します")
        
        # アンケートデータ分析クラスのインスタンスを作成
        analyzer = AnqDataAnalysis(target_ids)
        
        # スプレッドシートに接続
        if not analyzer.connect_to_spreadsheet():
            logger.error("スプレッドシートへの接続に失敗しました")
            return False, None
        
        # アンケートデータを取得
        if not analyzer.get_anq_data():
            logger.error("アンケートデータの取得に失敗しました")
            return False, None
        
        # CSVファイルに保存
        csv_path = analyzer.save_to_csv(analyzer.data_df)
        if not csv_path:
            logger.error("CSVファイルの保存に失敗しました")
            return False, None
        
        logger.info(f"✅ アンケートデータの分析とCSV出力が完了しました: {csv_path}")
        return True, csv_path
        
    except Exception as e:
        logger.error(f"アンケートデータの分析中にエラーが発生しました: {str(e)}")
        logger.error(traceback.format_exc())
        return False, None 