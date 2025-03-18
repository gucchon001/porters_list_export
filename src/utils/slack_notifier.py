#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Slack通知機能

エラーやイベント発生時にSlackへ通知を送信するためのユーティリティクラスを提供します。
環境変数からWebhook URLを取得し、指定されたメッセージをSlackに送信します。
"""

import os
import json
import requests
from typing import Dict, Any, Optional
import traceback
import platform
import socket
from datetime import datetime

from src.utils.logging_config import get_logger

logger = get_logger(__name__)

class SlackNotifier:
    """
    Slack通知を送信するユーティリティクラス
    """
    
    def __init__(self, webhook_url: Optional[str] = None):
        """
        SlackNotifierの初期化
        
        Args:
            webhook_url (Optional[str]): Slack Webhook URL。指定しない場合は環境変数から取得
        """
        self.webhook_url = webhook_url or os.environ.get('SLACK_WEBHOOK')
        
        if not self.webhook_url:
            logger.warning("Slack Webhook URLが設定されていません。Slack通知は無効です。")
    
    def send_message(self, message: str, title: Optional[str] = None, 
                     color: str = "#36a64f", fields: Optional[Dict[str, str]] = None) -> bool:
        """
        Slackにメッセージを送信
        
        Args:
            message (str): 送信するメッセージ本文
            title (Optional[str]): メッセージのタイトル (デフォルト: None)
            color (str): メッセージの色 (デフォルト: 緑)
            fields (Optional[Dict[str, str]]): 追加のフィールド情報
            
        Returns:
            bool: 送信が成功した場合はTrue、失敗した場合はFalse
        """
        if not self.webhook_url:
            logger.warning("Webhook URLが設定されていないため、Slackへの通知はスキップされました")
            return False
            
        try:
            # 現在のホスト名とIPアドレスを取得
            hostname = platform.node()
            try:
                ip_address = socket.gethostbyname(socket.gethostname())
            except:
                ip_address = "不明"
                
            # 現在の日時を取得
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # 基本的なフィールド情報を設定
            default_fields = [
                {
                    "title": "環境",
                    "value": os.environ.get('APP_ENV', 'development'),
                    "short": True
                },
                {
                    "title": "ホスト",
                    "value": f"{hostname} ({ip_address})",
                    "short": True
                },
                {
                    "title": "発生時刻",
                    "value": current_time,
                    "short": True
                }
            ]
            
            # 追加のフィールド情報があれば追加
            if fields:
                for key, value in fields.items():
                    default_fields.append({
                        "title": key,
                        "value": value,
                        "short": True
                    })
            
            # Slackメッセージのペイロードを作成
            payload = {
                "attachments": [
                    {
                        "fallback": title or "PORTERS List Export 通知",
                        "color": color,
                        "title": title or "PORTERS List Export 通知",
                        "text": message,
                        "fields": default_fields,
                        "footer": "PORTERS List Export",
                        "footer_icon": "https://platform.slack-edge.com/img/default_application_icon.png",
                        "ts": int(datetime.now().timestamp())
                    }
                ]
            }
            
            # POSTリクエストを送信
            response = requests.post(
                self.webhook_url,
                data=json.dumps(payload),
                headers={"Content-Type": "application/json"}
            )
            
            # レスポンスをチェック
            if response.status_code == 200 and response.text == 'ok':
                logger.info(f"Slack通知が正常に送信されました: {title}")
                return True
            else:
                logger.error(f"Slack通知の送信に失敗しました: ステータスコード={response.status_code}, レスポンス={response.text}")
                return False
                
        except Exception as e:
            logger.error(f"Slack通知の送信中にエラーが発生しました: {str(e)}")
            return False
    
    def send_error(self, error_message: str, exception: Optional[Exception] = None, 
                  title: str = "エラー発生", context: Optional[Dict[str, str]] = None) -> bool:
        """
        エラー情報をSlackに送信
        
        Args:
            error_message (str): エラーの説明メッセージ
            exception (Optional[Exception]): 発生した例外オブジェクト
            title (str): メッセージのタイトル (デフォルト: 'エラー発生')
            context (Optional[Dict[str, str]]): エラー発生時のコンテキスト情報
            
        Returns:
            bool: 送信が成功した場合はTrue、失敗した場合はFalse
        """
        # エラーメッセージを構築
        message = f"*{error_message}*\n"
        
        # 例外情報があれば追加
        if exception:
            message += f"\n```\n{str(exception)}\n```"
            
            # スタックトレースも追加
            stack_trace = traceback.format_exc()
            if stack_trace and stack_trace != "NoneType: None\n":
                message += f"\n*スタックトレース:*\n```\n{stack_trace[:1000]}```"
                if len(stack_trace) > 1000:
                    message += "\n(スタックトレースが長すぎるため省略されました)"
        
        # 追加のコンテキスト情報
        fields = {}
        if context:
            fields.update(context)
            
        # エラーメッセージを送信
        return self.send_message(message, title, color="#ff0000", fields=fields)
    
    @staticmethod
    def get_instance() -> 'SlackNotifier':
        """
        SlackNotifierのシングルトンインスタンスを取得
        
        Returns:
            SlackNotifier: SlackNotifierのインスタンス
        """
        return SlackNotifier() 