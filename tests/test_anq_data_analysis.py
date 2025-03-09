import os
import sys
from pathlib import Path
import logging
import time
import json
from datetime import datetime

# プロジェクトのルートディレクトリをPYTHONPATHに追加
root_dir = Path(__file__).parent.parent
sys.path.append(str(root_dir))

import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

# EnvironmentUtils を env という名前でインポート
from src.utils.environment import EnvironmentUtils as env
from src.utils.logging_config import get_logger
# settings モジュールから load_sheet_settings 関数をインポート
from src.modules.common.settings import load_sheet_settings

logger = get_logger(__name__)

class AnqDataAnalysis:
    """
    アンケートDLデータの分析を行うクラス
    """
    
    def __init__(self):
        """初期化処理"""
        logger.info("アンケートデータ分析クラスを初期化します")
        
        # 環境設定の読み込み - standalone_test.py と同様にシンプルに
        env.load_env()
        logger.info("環境変数を読み込みました")
        
        # スクリーンショットディレクトリの設定
        self.screenshot_dir = os.path.join(root_dir, "tests", "screenshots")
        os.makedirs(self.screenshot_dir, exist_ok=True)
        
        # 出力ディレクトリの設定
        self.output_dir = os.path.join(root_dir, "output")
        os.makedirs(self.output_dir, exist_ok=True)
        logger.info(f"出力ディレクトリを作成しました: {self.output_dir}")
        
        # サービスアカウントファイルのパスを取得 - 環境変数から直接取得
        service_account_path = env.get_env_var('SERVICE_ACCOUNT_FILE', 
                                              'config/credentials/service_account.json')
        self.service_account_file = os.path.join(env.get_project_root(), service_account_path)
        logger.info(f"サービスアカウントファイル: {self.service_account_file}")
        
        # スプレッドシートIDの取得 - settings.iniから取得
        self.spreadsheet_id = env.get_config_value(
            "SPREADSHEET", "SSID", 
            default="1ttg8heypq0fSbE2CYdSdhEOK7pLAvQG9655wgiFtMVk"
        )
        logger.info(f"スプレッドシートID: {self.spreadsheet_id}")
        
        # シート名の取得 - settings.iniから取得
        self.anq_data_key = env.get_config_value(
            "SHEET_NAMES", "ANQ_DATA", 
            default="アンケートDLデータ"
        ).strip('"')  # ダブルクォートを除去
        logger.info(f"アンケートデータキー: {self.anq_data_key}")
        
        # Google Sheets APIへの接続設定
        self.gc = None
        self.spreadsheet = None
        self.anq_data_sheet = None
        
        # 結果格納用
        self.headers = []
        self.data_df = None
        
        # 検索対象のID（複数）
        self.target_ids = [
            "147989543",  # 元のID
            "147879809",  # 新しく追加したID
            "147599677",  # 新しく追加したID
            "147523237"   # 新しく追加したID
        ]
        logger.info(f"検索対象ID一覧: {self.target_ids}")
    
    def connect_to_spreadsheet(self):
        """
        Google Sheetsに接続し、スプレッドシートとワークシートを取得
        """
        try:
            logger.info("Google Sheetsへの接続を開始します")
            
            # サービスアカウントファイルの存在確認
            if not os.path.exists(self.service_account_file):
                logger.error(f"サービスアカウントファイルが見つかりません: {self.service_account_file}")
                return False
            
            # スコープ設定
            scope = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
            
            # 認証情報の設定
            creds = Credentials.from_service_account_file(
                self.service_account_file, scopes=scope
            )
            
            # Google Sheetsクライアントの作成
            self.gc = gspread.authorize(creds)
            logger.info("✓ Google Sheets API認証に成功しました")
            
            # スプレッドシートを開く
            self.spreadsheet = self.gc.open_by_key(self.spreadsheet_id)
            logger.info(f"✓ スプレッドシート '{self.spreadsheet.title}' を開きました")
            
            # settings シートからシート名マッピングを取得
            try:
                # まず settings シートが存在するか確認
                settings_sheet = self.spreadsheet.worksheet('settings')
                logger.info("✓ settings シートを取得しました")
                
                # settings シートからマッピングを取得
                settings_data = settings_sheet.get_all_values()
                logger.info(f"settings シートのデータ行数: {len(settings_data)}")
                
                # 設定値を辞書に変換（ヘッダー行をスキップ）
                settings = {row[0]: row[1] for row in settings_data[1:] if len(row) >= 2}
                logger.info(f"設定マッピング: {settings}")
                
                # アンケートデータシートの実際の名前を取得
                if self.anq_data_key in settings:
                    self.anq_data_sheet_name = settings[self.anq_data_key]
                    logger.info(f"✓ アンケートデータ実際のシート名: {self.anq_data_sheet_name}")
                else:
                    logger.warning(f"❌ 設定キー '{self.anq_data_key}' が settings シートに見つかりません")
                    logger.warning(f"利用可能な設定: {settings}")
                    logger.warning(f"キー '{self.anq_data_key}' をそのままシート名として使用します")
                    self.anq_data_sheet_name = self.anq_data_key
            
            except gspread.exceptions.WorksheetNotFound:
                logger.warning("settings シートが見つかりません。キーをそのままシート名として使用します")
                self.anq_data_sheet_name = self.anq_data_key
            
            except Exception as e:
                logger.warning(f"設定の読み込み中にエラーが発生しました: {str(e)}")
                logger.warning(f"キー '{self.anq_data_key}' をそのままシート名として使用します")
                self.anq_data_sheet_name = self.anq_data_key
            
            # アンケートDLデータシートを取得
            try:
                self.anq_data_sheet = self.spreadsheet.worksheet(self.anq_data_sheet_name)
                logger.info(f"✓ シート '{self.anq_data_sheet_name}' を取得しました")
            except gspread.exceptions.WorksheetNotFound:
                logger.error(f"❌ シート '{self.anq_data_sheet_name}' が見つかりません")
                
                # 利用可能なシート一覧を表示
                all_worksheets = self.spreadsheet.worksheets()
                logger.info("利用可能なシート一覧:")
                for i, worksheet in enumerate(all_worksheets):
                    logger.info(f"  {i+1}. {worksheet.title}")
                
                return False
            
            return True
            
        except Exception as e:
            logger.error(f"Google Sheetsへの接続に失敗しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return False
    
    def get_headers(self):
        """
        ヘッダー行を取得
        """
        try:
            logger.info(f"シート '{self.anq_data_sheet_name}' のヘッダー行を取得します")
            
            # シートの1行目を取得
            self.headers = self.anq_data_sheet.row_values(1)
            
            # ヘッダー情報をログに出力
            logger.info(f"ヘッダー数: {len(self.headers)}")
            for i, header in enumerate(self.headers):
                logger.info(f"  {i+1}. {header}")
            
            return self.headers
            
        except Exception as e:
            logger.error(f"ヘッダー行の取得に失敗しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return []
    
    def get_record_by_id(self, target_id):
        """
        指定されたIDのレコードを取得
        
        Args:
            target_id (str): 検索対象のID
            
        Returns:
            dict: ヘッダーとデータの辞書、見つからない場合はNone
        """
        try:
            logger.info(f"ID '{target_id}' のレコードを検索します")
            
            # A列のすべての値を取得
            a_column = self.anq_data_sheet.col_values(1)
            logger.info(f"A列のデータ数: {len(a_column)}")
            
            # ヘッダー行をスキップしてIDを検索
            row_index = None
            for i, value in enumerate(a_column[1:], start=2):  # 2行目から開始
                if value == target_id:
                    row_index = i
                    logger.info(f"✓ ID '{target_id}' が {row_index} 行目に見つかりました")
                    break
            
            if row_index is None:
                logger.warning(f"❌ ID '{target_id}' が見つかりませんでした")
                return None
            
            # 該当行のデータを取得
            row_data = self.anq_data_sheet.row_values(row_index)
            logger.info(f"取得したデータ項目数: {len(row_data)}")
            
            # ヘッダーとデータを辞書に変換
            result = {}
            for i, header in enumerate(self.headers):
                if i < len(row_data):
                    result[header] = row_data[i]
                else:
                    result[header] = ""  # データがない場合は空文字
            
            # 結果をログに出力
            logger.info("=== 取得したレコードの内容 ===")
            for header, value in result.items():
                logger.info(f"  {header}: {value}")
            
            return result
            
        except Exception as e:
            logger.error(f"レコードの取得に失敗しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return None
    
    def get_records_by_ids(self, target_ids):
        """
        複数の指定されたIDのレコードを取得
        
        Args:
            target_ids (list): 検索対象のID一覧
            
        Returns:
            list: レコード辞書のリスト
        """
        try:
            logger.info(f"複数ID {target_ids} のレコードを検索します")
            
            # 結果を格納するリスト
            results = []
            
            # A列のすべての値を一度だけ取得（効率化のため）
            a_column = self.anq_data_sheet.col_values(1)
            logger.info(f"A列のデータ数: {len(a_column)}")
            
            # 各IDについて検索
            for target_id in target_ids:
                logger.info(f"ID '{target_id}' の検索を開始します")
                
                # ヘッダー行をスキップしてIDを検索
                row_index = None
                for i, value in enumerate(a_column[1:], start=2):  # 2行目から開始
                    if value == target_id:
                        row_index = i
                        logger.info(f"✓ ID '{target_id}' が {row_index} 行目に見つかりました")
                        break
                
                if row_index is None:
                    logger.warning(f"❌ ID '{target_id}' が見つかりませんでした")
                    continue  # 次のIDへ
                
                # 該当行のデータを取得
                row_data = self.anq_data_sheet.row_values(row_index)
                logger.info(f"取得したデータ項目数: {len(row_data)}")
                
                # ヘッダーとデータを辞書に変換
                result = {}
                for i, header in enumerate(self.headers):
                    if i < len(row_data):
                        result[header] = row_data[i]
                    else:
                        result[header] = ""  # データがない場合は空文字
                
                # 結果をログに出力
                logger.info(f"=== ID '{target_id}' のレコード内容 ===")
                for header, value in result.items():
                    logger.info(f"  {header}: {value}")
                
                # 結果リストに追加
                results.append(result)
            
            logger.info(f"合計 {len(results)} 件のレコードを取得しました")
            return results
            
        except Exception as e:
            logger.error(f"複数レコードの取得に失敗しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return []
    
    def save_to_csv(self, df, filename=None):
        """
        DataFrameをCSVファイルとして保存
        
        Args:
            df (DataFrame): 保存するDataFrame
            filename (str, optional): ファイル名。指定しない場合は日時から自動生成
            
        Returns:
            str: 保存したファイルのパス
        """
        try:
            if df is None or df.empty:
                logger.error("保存するデータがありません")
                return None
            
            # ファイル名が指定されていない場合は日時から自動生成
            if filename is None:
                now = datetime.now()
                timestamp = now.strftime("%Y%m%d_%H%M%S")
                filename = f"anq_data_{timestamp}.csv"
            
            # 出力ファイルパスの作成
            filepath = os.path.join(self.output_dir, filename)
            
            # CSVファイルとして保存（cp932エンコーディング）
            df.to_csv(filepath, encoding='cp932', index=False)
            logger.info(f"✓ CSVファイルを保存しました: {filepath}")
            
            return filepath
            
        except UnicodeEncodeError as e:
            logger.error(f"文字コードエラーが発生しました: {str(e)}")
            logger.info("cp932で保存できない文字が含まれています。utf-8で保存を試みます。")
            
            try:
                # UTF-8で保存を試みる
                filepath = os.path.join(self.output_dir, f"utf8_{filename}")
                df.to_csv(filepath, encoding='utf-8', index=False)
                logger.info(f"✓ CSVファイルをUTF-8で保存しました: {filepath}")
                return filepath
            except Exception as utf8_error:
                logger.error(f"UTF-8での保存にも失敗しました: {str(utf8_error)}")
                return None
                
        except Exception as e:
            logger.error(f"CSVファイルの保存に失敗しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return None
    
    def run(self):
        """
        メイン実行処理
        """
        try:
            logger.info("=== アンケートデータ分析を開始します ===")
            
            # Google Sheetsに接続
            if not self.connect_to_spreadsheet():
                logger.error("スプレッドシートへの接続に失敗したため、処理を中止します")
                return False
            
            # ヘッダー行を取得
            headers = self.get_headers()
            if not headers:
                logger.error("ヘッダー行の取得に失敗したため、処理を中止します")
                return False
            
            # 複数IDのレコードを取得
            records = self.get_records_by_ids(self.target_ids)
            if not records:
                logger.error(f"指定されたIDのレコード取得に失敗したため、処理を中止します")
                return False
            
            logger.info(f"取得したレコード数: {len(records)}")
            
            # 取得したデータをDataFrameに変換
            self.data_df = pd.DataFrame(records)
            logger.info(f"DataFrame作成完了: {self.data_df.shape[0]}行 x {self.data_df.shape[1]}列")
            
            # DataFrameの先頭5行を表示
            logger.info("=== DataFrame先頭5行 ===")
            logger.info(self.data_df.head().to_string())
            
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
            import traceback
            logger.error(traceback.format_exc())
            return False

def main():
    """
    メイン処理
    """
    try:
        analyzer = AnqDataAnalysis()
        result = analyzer.run()
        
        if result:
            print("✅ アンケートデータ分析が正常に完了しました")
        else:
            print("❌ アンケートデータ分析に失敗しました")
        
        return result
        
    except Exception as e:
        logger.error(f"予期せぬエラーが発生しました: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        print(f"❌ エラーが発生しました: {str(e)}")
        return False

if __name__ == "__main__":
    result = main()
    print(f"テスト結果: {'成功' if result else '失敗'}") 