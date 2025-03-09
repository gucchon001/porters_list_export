import time
from ..utils.environment import EnvironmentUtils as env
from ..utils.logging_config import get_logger
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from datetime import datetime

logger = get_logger(__name__)

class Login:
    def __init__(self, browser):
        """
        ログイン機能を管理するクラス
        
        Args:
            browser: Browserクラスのインスタンス
        """
        self.browser = browser
        logger.info("=== ログインプロセスを開始 ===")

    def execute(self):
        """ログイン処理を実行"""
        try:
            logger.info("👤 新規ログインを開始します")
            success, _ = self._perform_login()
            return success
            
        except Exception as e:
            logger.error(f"❌ ログイン処理でエラー: {str(e)}")
            return False

    def _perform_login(self):
        """実際のログイン処理を実行"""
        try:
            login_url = "https://manager.linestep.net/account/login"
            self.browser.driver.get(login_url)
            logger.info(f"📍 ログインページにアクセス: {login_url}")
            time.sleep(2)

            login_credentials = {
                'id': env.get_env_var('LOGIN_ID'),
                'password': env.get_env_var('LOGIN_PASSWORD')
            }
            logger.debug(f"ログイン情報を取得: ID={login_credentials['id'][:3]}***")

            # ログインフォームの入力
            logger.info("ユーザー名フィールドを探索中...")
            username_field = self.browser._get_element('login', 'username')
            username_field.clear()
            username_field.send_keys(login_credentials['id'])
            logger.info(f"✓ ユーザー名を入力: {login_credentials['id'][:3]}***")
            time.sleep(1)

            logger.info("パスワードフィールドを探索中...")
            password_field = self.browser._get_element('login', 'password')
            password_field.clear()
            password_field.send_keys(login_credentials['password'])
            logger.info("✓ パスワードを入力: ********")
            time.sleep(1)

            logger.info("⚠️ reCAPTCHA認証を待機中...")
            print("\n⚠️ reCAPTCHA認証を手動で完了し、ログインボタンをクリックしてください")
            
            # URLの変更を監視
            current_url = self.browser.driver.current_url
            for _ in range(300):  # 5分間待機
                time.sleep(1)
                try:
                    new_url = self.browser.driver.current_url
                    if new_url != current_url and 'login' not in new_url:
                        logger.info("✅ ログイン成功")
                        return True, new_url
                except Exception as e:
                    logger.error(f"URL確認でエラー: {str(e)}")
                    continue
            
            logger.error("ログインタイムアウト")
            return False, None

        except Exception as e:
            logger.error(f"❌ ログイン処理でエラー: {str(e)}")
            return False, None 