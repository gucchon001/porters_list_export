import os
import sys
import time
from pathlib import Path
import argparse
import traceback

# プロジェクトのルートディレクトリをPYTHONPATHに追加
root_dir = Path(__file__).parent.parent
sys.path.append(str(root_dir))

from src.utils.logging_config import get_logger
from src.utils.environment import EnvironmentUtils as env
from .test_browser import TestBrowser
from .test_login import TestLogin
from .test_csv_import import TestCsvImport

logger = get_logger(__name__)

def main():
    """テスト実行のメイン関数"""
    parser = argparse.ArgumentParser(description='ログインとCSVインポートのテスト')
    parser.add_argument('--headless', action='store_true', help='ヘッドレスモードで実行')
    parser.add_argument('--skip-login', action='store_true', help='ログインをスキップ')
    parser.add_argument('--skip-import', action='store_true', help='CSVインポートをスキップ')
    args = parser.parse_args()
    
    logger.info("=== テスト実行を開始します ===")
    
    # 環境変数のロード - 重要！
    env.load_env()
    
    # 設定ファイルのパス
    selectors_path = os.path.join(root_dir, "config", "selectors.csv")
    
    # ブラウザセットアップ
    browser = TestBrowser(selectors_path=selectors_path)
    try:
        # セレクタの検証とフォールバック
        if not browser.selectors or 'porters' not in browser.selectors:
            logger.warning("PORTERSのセレクタ情報が見つかりません。デフォルト値を使用します。")
            browser.selectors['porters'] = {
                'company_id': {
                    'selector_type': 'css',
                    'selector_value': "#Model_LoginForm_company_login_id"
                },
                'username': {
                    'selector_type': 'css',
                    'selector_value': "#Model_LoginForm_login_id"
                },
                'password': {
                    'selector_type': 'css',
                    'selector_value': "#Model_LoginForm_password"
                }
            }
            
        # porters_menuがなければ初期化
        if 'porters_menu' not in browser.selectors:
            logger.warning("PORTERSメニューのセレクタ情報が見つかりません。デフォルト値を使用します。")
            browser.selectors['porters_menu'] = {
                'search_button': {
                    'selector_type': 'css',
                    'selector_value': "#main > div > main > section.original-search > header > div.others > button"
                },
                'menu_item_5': {
                    'selector_type': 'css',
                    'selector_value': "#main-menu-id-5 > a"
                }
            }
        
        # WebDriverのセットアップ
        if not browser.setup():
            logger.error("ブラウザのセットアップに失敗しました")
            return False
        
        # ログイン処理
        if not args.skip_login:
            login = TestLogin(browser)
            if not login.execute():
                logger.error("ログイン処理に失敗しました")
                return False
        else:
            logger.info("ログイン処理をスキップします")
        
        # CSVインポート処理
        if not args.skip_import:
            csv_import = TestCsvImport(browser)
            if not csv_import.execute():
                logger.error("CSVインポート処理に失敗しました")
                return False
        else:
            logger.info("CSVインポート処理をスキップします")
        
        # ログアウト処理
        if not args.skip_login:
            login.logout()
        
        logger.info("✅ テスト実行が正常に完了しました")
        return True
        
    except Exception as e:
        logger.error(f"テスト実行中にエラーが発生しました: {str(e)}")
        logger.error(traceback.format_exc())
        return False
    finally:
        # ブラウザを終了
        browser.quit()

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1) 