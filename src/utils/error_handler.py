#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
共通エラーハンドリングユーティリティ

エラーの捕捉、ログ記録、Slack通知を統一的に管理するためのユーティリティクラスを提供します。
モジュール間でエラー処理と通知を標準化するために使用します。
"""

import os
import sys
import traceback
import inspect
from datetime import datetime
from typing import Dict, Any, Optional, Callable, TypeVar, cast

from src.utils.logging_config import get_logger
from src.utils.slack_notifier import SlackNotifier

logger = get_logger(__name__)

# 戻り値の型を定義
T = TypeVar('T')

class ErrorHandler:
    """
    統一的なエラー処理と通知を提供するクラス
    """
    
    def __init__(self, module_name: str = None, slack_notifier: SlackNotifier = None):
        """
        ErrorHandlerの初期化
        
        Args:
            module_name (str, optional): エラーが発生したモジュール名。指定しない場合は自動検出
            slack_notifier (SlackNotifier, optional): SlackNotifierのインスタンス。指定しない場合は新規作成
        """
        self.module_name = module_name or self._detect_module_name()
        self.slack_notifier = slack_notifier or SlackNotifier.get_instance()
    
    def _detect_module_name(self) -> str:
        """
        呼び出し元のモジュール名を自動的に検出
        
        Returns:
            str: 検出されたモジュール名
        """
        frame = inspect.currentframe()
        if frame:
            try:
                # 呼び出し元のフレームを取得
                caller_frame = frame.f_back
                if caller_frame:
                    # モジュール名を取得
                    module = inspect.getmodule(caller_frame)
                    if module:
                        return module.__name__
            finally:
                # 循環参照を防ぐためにフレームを明示的に解放
                del frame
        
        # 検出できない場合はデフォルト値を返す
        return "unknown_module"
    
    def handle_exception(self, error_message: str, exception: Exception = None, context: dict = None, slack_title: str = None) -> None:
        """
        例外をハンドリングしてログ記録とSlack通知を行う
        
        Args:
            error_message: エラーメッセージ
            exception: 例外オブジェクト
            context: エラーの追加コンテキスト情報
            slack_title: Slack通知のタイトル（指定がない場合はデフォルトタイトルを使用）
        """
        # コンテキスト情報の準備
        ctx = context or {}
        
        # 例外情報を追加
        if exception:
            import traceback
            tb_str = "".join(traceback.format_exception(type(exception), exception, exception.__traceback__))
            
            # ログにエラー詳細を記録
            self.logger.error(f"{error_message}: {str(exception)}")
            self.logger.error(tb_str)
            
            # コンテキストに例外情報を追加
            ctx["例外内容"] = str(exception)
            ctx["例外タイプ"] = exception.__class__.__name__
        else:
            # 例外がない場合もエラーメッセージをログに記録
            self.logger.error(error_message)
        
        # Slack通知
        if self.slack_notifier:
            try:
                # コンテキスト情報を整形
                formatted_context = self.format_error_context(ctx)
                
                # タイトルの設定
                title = slack_title or f"{self.module_name}でエラーが発生しました"
                
                # Slack通知本文の作成
                message = f"*{error_message}*\n\n{formatted_context}"
                if exception:
                    # スタックトレースの最初の3行を追加
                    stack_lines = tb_str.splitlines()[:3]
                    stack_preview = "\n".join(stack_lines)
                    message += f"\n\n*スタックトレース（抜粋）*:\n```{stack_preview}```"
                
                # Slack通知を送信
                self.slack_notifier.send_notification(title, message)
                self.logger.info("Slackにエラー通知を送信しました")
            except Exception as e:
                self.logger.error(f"Slack通知の送信に失敗しました: {str(e)}")
    
    def with_error_handling(
        self, 
        func: Callable[..., T], 
        error_message: str, 
        context: Dict[str, Any] = None,
        screenshot_func: Callable[[Exception], Optional[str]] = None,
        slack_title: str = None,
        default_return: Any = None
    ) -> T:
        """
        関数をエラーハンドリングのコンテキストで実行
        
        Args:
            func (Callable): 実行する関数
            error_message (str): エラー時のメッセージ
            context (Dict[str, Any], optional): エラー発生時のコンテキスト
            screenshot_func (Callable, optional): エラー時にスクリーンショットを取得する関数
            slack_title (str, optional): Slack通知のタイトル
            default_return (Any, optional): エラー時の戻り値
            
        Returns:
            T: 関数の実行結果、エラー時はdefault_return
        """
        try:
            return func()
        except Exception as e:
            # スクリーンショットを取得
            if screenshot_func:
                try:
                    screenshot_path = screenshot_func(e)
                    if screenshot_path and context is not None:
                        context["スクリーンショット"] = screenshot_path
                except Exception as screenshot_e:
                    self.logger.error(f"スクリーンショット取得中にエラーが発生しました: {str(screenshot_e)}")
            
            # エラーを処理
            self.handle_exception(
                error_message=error_message,
                exception=e,
                context=context,
                slack_title=slack_title
            )
            
            # デフォルト値を返す
            return cast(T, default_return)
    
    @staticmethod
    def for_module(module_name: str = None) -> 'ErrorHandler':
        """
        指定されたモジュール用のErrorHandlerインスタンスを取得
        
        Args:
            module_name (str, optional): モジュール名
            
        Returns:
            ErrorHandler: ErrorHandlerのインスタンス
        """
        return ErrorHandler(module_name)

    def format_error_context(self, context: dict) -> str:
        """
        エラーコンテキスト情報を整形された文字列に変換する
        
        Args:
            context (dict): エラーコンテキスト情報の辞書
            
        Returns:
            str: 整形されたコンテキスト情報
        """
        if not context:
            return "コンテキスト情報なし"
            
        formatted_text = "【エラーコンテキスト】\n"
        
        # 重要な情報を先に表示
        priority_keys = ["操作", "処理", "詳細", "シート情報", "URL"]
        for key in priority_keys:
            if key in context:
                value = context[key]
                # 複数行の場合はインデントを追加
                if isinstance(value, str) and "\n" in value:
                    value_lines = value.split("\n")
                    formatted_value = value_lines[0]
                    for line in value_lines[1:]:
                        formatted_value += f"\n    {line}"
                    formatted_text += f"・{key}: {formatted_value}\n"
                else:
                    formatted_text += f"・{key}: {value}\n"
                    
        # 残りの情報を表示
        for key, value in context.items():
            if key not in priority_keys:
                if isinstance(value, str) and "\n" in value:
                    value_lines = value.split("\n")
                    formatted_value = value_lines[0]
                    for line in value_lines[1:]:
                        formatted_value += f"\n    {line}"
                    formatted_text += f"・{key}: {formatted_value}\n"
                else:
                    formatted_text += f"・{key}: {value}\n"
                    
        return formatted_text 