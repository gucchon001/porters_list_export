#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
汎用ユーティリティ関数を提供するモジュール

このモジュールは、ファイル操作やデータ処理など、
プロジェクト全体で使用される汎用的な関数を提供します。
"""

import os
import glob
import time
from pathlib import Path
from typing import List, Optional
from datetime import datetime
import logging

from src.utils.logging_config import get_logger

logger = get_logger(__name__)

def find_latest_file(directory: str, pattern: str) -> Optional[str]:
    """
    指定されたディレクトリ内で、指定されたパターンに一致する最新のファイルを探す
    
    Args:
        directory (str): 検索対象のディレクトリパス
        pattern (str): ファイル名のパターン（例: "*.csv"）
        
    Returns:
        Optional[str]: 最新のファイルのパス。ファイルが見つからない場合はNone。
    """
    if not os.path.exists(directory):
        return None
    
    files = glob.glob(os.path.join(directory, pattern))
    if not files:
        return None
    
    # 最終更新日時でソート
    latest_file = max(files, key=os.path.getmtime)
    return latest_file

def find_latest_csv_in_downloads() -> Optional[str]:
    """
    ダウンロードディレクトリ内で最新のCSVファイルを探す
    
    Returns:
        Optional[str]: 最新のCSVファイルのパス。ファイルが見つからない場合はNone。
    """
    # プロジェクト内のdownloadsディレクトリを最初に確認
    project_download_dir = os.path.join(os.getcwd(), "downloads")
    if os.path.exists(project_download_dir):
        latest_file = find_latest_file(project_download_dir, "*.csv")
        if latest_file:
            logger.info(f"プロジェクト内のdownloadsディレクトリで最新のCSVファイルを発見: {latest_file}")
            return latest_file
    
    # プロジェクト内で見つからない場合は、ユーザーのダウンロードディレクトリを確認
    home_dir = os.path.expanduser("~")
    download_dir = os.path.join(home_dir, "Downloads")
    
    # Windowsの場合は別のパスを試す
    if not os.path.exists(download_dir):
        download_dir = os.path.join(home_dir, "ダウンロード")
    
    # それでも見つからない場合は環境変数を確認
    if not os.path.exists(download_dir) and "USERPROFILE" in os.environ:
        download_dir = os.path.join(os.environ["USERPROFILE"], "Downloads")
    
    return find_latest_file(download_dir, "*.csv")

def wait_for_new_csv_in_downloads(timeout: int = 60, check_interval: float = 1.0) -> Optional[str]:
    """
    ダウンロードディレクトリ内に新しいCSVファイルが現れるのを待つ
    
    Args:
        timeout (int): タイムアウト時間（秒）
        check_interval (float): チェック間隔（秒）
        
    Returns:
        Optional[str]: 新しいCSVファイルのパス。タイムアウトした場合はNone。
    """
    # プロジェクト内のdownloadsディレクトリを最初に確認
    project_download_dir = os.path.join(os.getcwd(), "downloads")
    download_dirs = []
    
    if os.path.exists(project_download_dir):
        download_dirs.append(project_download_dir)
        logger.info(f"プロジェクト内のダウンロードディレクトリを監視します: {project_download_dir}")
    
    # ユーザーのダウンロードディレクトリも念のため確認
    home_dir = os.path.expanduser("~")
    user_download_dirs = [
        os.path.join(home_dir, "Downloads"),
        os.path.join(home_dir, "ダウンロード"),
        os.path.join(home_dir, "Desktop"),
        os.path.join(home_dir, "デスクトップ")
    ]
    
    # 環境変数からダウンロードディレクトリを追加
    for env_var in ["USERPROFILE", "HOME", "HOMEPATH"]:
        if env_var in os.environ:
            user_download_dirs.append(os.path.join(os.environ[env_var], "Downloads"))
            user_download_dirs.append(os.path.join(os.environ[env_var], "ダウンロード"))
    
    # Windowsの場合、OneDriveのダウンロードフォルダも確認
    if "USERPROFILE" in os.environ:
        onedrive_dir = os.path.join(os.environ["USERPROFILE"], "OneDrive")
        if os.path.exists(onedrive_dir):
            user_download_dirs.append(os.path.join(onedrive_dir, "Downloads"))
            user_download_dirs.append(os.path.join(onedrive_dir, "ダウンロード"))
    
    # 存在するユーザーのダウンロードディレクトリを追加
    for dir_path in user_download_dirs:
        if os.path.exists(dir_path):
            download_dirs.append(dir_path)
            logger.info(f"ユーザーのダウンロードディレクトリも監視します: {dir_path}")
    
    if not download_dirs:
        logger.error("有効なダウンロードディレクトリが見つかりませんでした")
        return None
    
    # 現在のCSVファイルとその更新時刻を記録
    current_files = {}
    for download_dir in download_dirs:
        for file_path in glob.glob(os.path.join(download_dir, "*.csv")):
            current_files[file_path] = os.path.getmtime(file_path)
            logger.debug(f"既存のCSVファイルを検出: {file_path}")
    
    logger.info(f"ダウンロード監視を開始します。既存のCSVファイル数: {len(current_files)}")
    
    start_time = time.time()
    while time.time() - start_time < timeout:
        # 少し待機
        time.sleep(check_interval)
        
        # 新しいCSVファイルを探す
        for download_dir in download_dirs:
            for file_path in glob.glob(os.path.join(download_dir, "*.csv")):
                # 新しいファイルか、更新されたファイルを検出
                if file_path not in current_files or os.path.getmtime(file_path) > current_files.get(file_path, 0):
                    # ファイルサイズが0でないことを確認（ダウンロード中でない）
                    if os.path.getsize(file_path) > 0:
                        # ファイルが完全にダウンロードされるまで少し待機
                        time.sleep(2)
                        
                        # ファイルサイズが変わらなくなったことを確認（ダウンロード完了）
                        initial_size = os.path.getsize(file_path)
                        time.sleep(1)
                        if os.path.getsize(file_path) == initial_size:
                            logger.info(f"新しいCSVファイルを検出: {file_path}")
                            return file_path
    
    # タイムアウト
    logger.warning(f"タイムアウト（{timeout}秒）: 新しいCSVファイルは検出されませんでした")
    
    # タイムアウト時に最新のCSVファイルを返す（代替手段）
    latest_csv = None
    latest_time = 0
    
    # プロジェクト内のdownloadsディレクトリを優先
    if os.path.exists(project_download_dir):
        for file_path in glob.glob(os.path.join(project_download_dir, "*.csv")):
            file_time = os.path.getmtime(file_path)
            if file_time > latest_time:
                latest_time = file_time
                latest_csv = file_path
    
    # プロジェクト内で見つからない場合は他のディレクトリも確認
    if not latest_csv:
        for download_dir in download_dirs:
            if download_dir != project_download_dir:  # プロジェクトディレクトリは既に確認済み
                for file_path in glob.glob(os.path.join(download_dir, "*.csv")):
                    file_time = os.path.getmtime(file_path)
                    if file_time > latest_time:
                        latest_time = file_time
                        latest_csv = file_path
    
    if latest_csv and latest_time > start_time - 300:  # 5分以内に更新されたファイルなら使用
        logger.info(f"タイムアウトしましたが、最近更新されたCSVファイルを使用します: {latest_csv}")
        return latest_csv
    
    return None

def move_file_to_data_dir(file_path: str, new_filename: Optional[str] = None, keep_original: bool = False) -> Optional[str]:
    """
    ファイルをdataディレクトリに移動する
    
    Args:
        file_path (str): 移動するファイルのパス
        new_filename (Optional[str]): 新しいファイル名（指定しない場合は元のファイル名を使用）
        keep_original (bool): 元のファイルを保持するかどうか（デフォルトはFalse）
        
    Returns:
        Optional[str]: 移動後のファイルパス。失敗した場合はNone。
    """
    if not os.path.exists(file_path):
        logger.error(f"移動元ファイルが存在しません: {file_path}")
        return None
    
    # dataディレクトリのパスを取得
    base_dir = Path(__file__).resolve().parent.parent.parent
    data_dir = os.path.join(base_dir, "data")
    
    # dataディレクトリが存在しない場合は作成
    if not os.path.exists(data_dir):
        os.makedirs(data_dir)
        logger.info(f"dataディレクトリを作成しました: {data_dir}")
    
    # 新しいファイル名を設定
    if new_filename is None:
        new_filename = os.path.basename(file_path)
    
    # タイムスタンプを付与
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename, ext = os.path.splitext(new_filename)
    new_filename = f"{filename}_{timestamp}{ext}"
    
    # 移動先のパス
    dest_path = os.path.join(data_dir, new_filename)
    
    try:
        import shutil
        if keep_original:
            # ファイルをコピー
            shutil.copy2(file_path, dest_path)
            logger.info(f"ファイルをコピーしました: {file_path} -> {dest_path}")
        else:
            # ファイルを移動
            shutil.move(file_path, dest_path)
            logger.info(f"ファイルを移動しました: {file_path} -> {dest_path}")
        
        return dest_path
    except Exception as e:
        logger.error(f"ファイルの移動/コピー中にエラーが発生しました: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return None 
