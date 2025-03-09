# utils.py
import os
import logging
import fnmatch
import configparser
from typing import List, Tuple, Optional
from pathlib import Path
from datetime import datetime

from anytree import Node, RenderTree
from openai import OpenAI

logger = logging.getLogger(__name__)

def normalize_path(path: str) -> str:
    """パスを正規化"""
    return os.path.normpath(path).replace('\\', '/')

def read_settings(settings_path: str = 'config/settings.ini') -> dict:
    """設定ファイルを読み込む"""
    config = configparser.ConfigParser()
    default_settings = {
        'source_directory': '.',  # デフォルトのソースディレクトリ
        'output_file': 'merge.txt',  # デフォルトの出力ファイル名
        'exclusions': 'env,myenv,*__pycache__*,downloads,sample_file,*.log',  # 除外パターン
        'openai_api_key': '',  # デフォルトのOpenAI APIキー
        'openai_model': 'gpt-4o'  # デフォルトのOpenAIモデル
    }
    
    try:
        if os.path.exists(settings_path):
            config.read(settings_path, encoding='utf-8')
            settings = {
                'source_directory': config['DEFAULT'].get('SourceDirectory', default_settings['source_directory']),
                'output_file': config['DEFAULT'].get('OutputFile', default_settings['output_file']),
                'exclusions': config['DEFAULT'].get('Exclusions', default_settings['exclusions']).replace(' ', '')
            }
            
            # APIセクションの設定を追加
            if 'API' in config:
                settings['openai_api_key'] = config['API'].get('openai_api_key', default_settings['openai_api_key'])
                settings['openai_model'] = config['API'].get('openai_model', default_settings['openai_model'])
            else:
                logger.warning("API section not found in settings.ini")
                settings['openai_api_key'] = default_settings['openai_api_key']
                settings['openai_model'] = default_settings['openai_model']
        else:
            logger.warning(f"Settings file not found at {settings_path}. Using default settings.")
            settings = default_settings

        # APIキーが設定されていない場合の警告
        if not settings['openai_api_key']:
            logger.warning("OpenAI API key is not set. Some functionality may be limited.")
        
        return settings
    except Exception as e:
        logger.error(f"Error reading settings file {settings_path}: {e}")
        return default_settings

def read_file_safely(filepath: str) -> Optional[str]:
    """ファイルを安全に読み込む"""
    for encoding in ['utf-8', 'cp932']:
        try:
            with open(filepath, 'r', encoding=encoding) as f:
                return f.read()
        except UnicodeDecodeError:
            continue  # 次のエンコーディングで試す
        except Exception as e:
            logger.error(f"Error reading file {filepath}: {e}")
            return None
    logger.warning(f"Failed to read file {filepath} with supported encodings.")
    return None

def write_file_content(filepath: str, content: str) -> bool:
    """ファイルに内容を書き込む"""
    try:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)
        return True
    except Exception as e:
        logger.error(f"Error writing to file {filepath}: {e}")
        return False

def get_python_files(directory: str, exclude_patterns: List[str]) -> List[Tuple[str, str]]:
    """指定ディレクトリ配下のPythonファイルを取得"""
    python_files = []
    exclude_patterns = [pattern.strip() for pattern in exclude_patterns if pattern.strip()]  # パターンの正規化
    
    try:
        for root, dirs, files in os.walk(directory):
            # 除外パターンに一致するディレクトリをスキップ
            dirs[:] = [d for d in dirs if not any(fnmatch.fnmatch(d, pattern) for pattern in exclude_patterns)]
            
            for file in files:
                if file.endswith('.py') and not any(fnmatch.fnmatch(file, pattern) for pattern in exclude_patterns):
                    filepath = os.path.join(root, file)
                    rel_path = os.path.relpath(filepath, directory)
                    python_files.append((rel_path, filepath))
        
        return sorted(python_files)
    except Exception as e:
        logger.error(f"Error getting Python files from {directory}: {e}")
        return []

# 以下は共通化された関数

# utils.py の setup_logger 関数を更新
def setup_logger(name: str, log_dir: Optional[str] = None, log_file: str = "merge_files.log") -> logging.Logger:
    """
    ロガーをセットアップします。

    Args:
        name (str): ロガーの名前。
        log_dir (Optional[str]): ログファイルの保存ディレクトリ。
        log_file (str): ログファイル名。

    Returns:
        logging.Logger: セットアップ済みのロガー。
    """
    if log_dir is None:
        log_dir = os.path.join(os.getcwd(), "logs")  # デフォルトでカレントディレクトリ内の 'logs' フォルダを使用

    try:
        # ディレクトリが存在しない場合は作成
        os.makedirs(log_dir, exist_ok=True)
    except PermissionError as e:
        print(f"ログディレクトリの作成に失敗しました: {log_dir}. エラー: {e}")
        raise

    # ログファイルパスを構築
    log_path = os.path.join(log_dir, log_file)
    print(f"Logging to: {log_path}")  # デバッグ用の出力

    # ロガーを構築
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    # 既存のハンドラーをクリア（重複ログを防ぐため）
    if logger.hasHandlers():
        logger.handlers.clear()

    try:
        # ログハンドラーを設定
        file_handler = logging.FileHandler(log_path, encoding='utf-8')
    except PermissionError as e:
        print(f"ログファイルの作成に失敗しました: {log_path}. エラー: {e}")
        raise
    except Exception as e:
        print(f"ログファイルハンドラーの設定中にエラーが発生しました: {e}")
        raise

    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # コンソールハンドラー（オプション）
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    return logger

class PromptLogger:
    """プロンプトとレスポンスをログに記録するクラス"""
    def __init__(self, logger: logging.Logger):
        self.logger = logger

    def log_prompt(self, prompt: str):
        """プロンプトをログに記録"""
        self.logger.debug("\n=== Prompt ===\n" + prompt + "\n")

    def log_response(self, response: str):
        """レスポンスをログに記録"""
        self.logger.debug("\n=== Response ===\n" + response + "\n")

def ensure_directories_exist(dirs: List[str]) -> None:
    """指定されたディレクトリが存在しない場合は作成する"""
    for dir_path in dirs:
        os.makedirs(dir_path, exist_ok=True)

def initialize_openai_client(api_key: Optional[str] = None) -> OpenAI:
    """OpenAIクライアントを初期化する"""
    if not api_key:
        api_key = os.getenv('OPENAI_API_KEY')
        if not api_key:
            logger.error("OpenAI APIキーが設定されていません。環境変数または設定ファイルを確認してください。")
            raise ValueError("OpenAI APIキーが設定されていません。")
    client = OpenAI(api_key=api_key)
    return client

def generate_tree_structure(root_dir: str, exclude_dirs: List[str], exclude_files: List[str]) -> str:
    """ディレクトリ構造をツリー形式で生成"""
    def add_nodes(parent_node: Node, current_path: str):
        try:
            for item in sorted(os.listdir(current_path)):
                item_path = os.path.join(current_path, item)
                if os.path.isdir(item_path):
                    if item not in exclude_dirs:
                        dir_node = Node(item, parent=parent_node)
                        add_nodes(dir_node, item_path)
                elif not any(fnmatch.fnmatch(item, pattern) for pattern in exclude_files):
                    Node(item, parent=parent_node)
        except PermissionError:
            logger.warning(f"アクセス権限がないためスキップしました: {current_path}")

    root_node = Node(os.path.basename(root_dir))
    add_nodes(root_node, root_dir)
    return "\n".join([f"{pre}{node.name}" for pre, _, node in RenderTree(root_node)])

def update_readme(template_path: str, readme_path: str, spec_path: str, merge_path: str) -> bool:
    """README.md をテンプレートから更新する"""
    try:
        template_content = read_file_safely(template_path)
        if not template_content:
            logger.error(f"{template_path} の読み込みに失敗しました。README.md は更新されませんでした。")
            return False

        # [spec] プレースホルダーを仕様書で置換
        spec_content = read_file_safely(spec_path)
        if not spec_content:
            logger.error(f"{spec_path} の読み込みに失敗しました。README.md の [spec] プレースホルダーは置換されませんでした。")
            spec_content = "[仕様書の内容が取得できませんでした。]"
        updated_content = template_content.replace("[spec]", spec_content)

        # [tree] プレースホルダーをディレクトリ構造で置換
        merge_content = read_file_safely(merge_path)
        if not merge_content:
            logger.error(f"{merge_path} の読み込みに失敗しました。README.md の [tree] プレースホルダーは置換されませんでした。")
            tree_content = "[フォルダ構成の取得に失敗しました。]"
        else:
            # " # Merged Python Files" までの内容を抽出
            split_marker = "# Merged Python Files"
            if split_marker in merge_content:
                tree_section = merge_content.split(split_marker)[0]
            else:
                tree_section = merge_content  # マーカーがなければ全体を使用
            tree_content = tree_section.strip()
        updated_content = updated_content.replace("[tree]", f"```\n{tree_content}\n```")

        # 現在の日付を挿入（オプション）
        current_date = datetime.now().strftime("%Y-%m-%d")
        updated_content = updated_content.replace("[YYYY-MM-DD]", current_date)

        # README.md に書き込む
        success = write_file_content(readme_path, updated_content)
        if success:
            logger.info(f"README.md が正常に更新されました: {readme_path}")
            return True
        else:
            logger.error(f"README.md の更新に失敗しました: {readme_path}")
            return False

    except Exception as e:
        logger.error(f"README.md の更新中にエラーが発生しました: {e}")
        return False

def initialize_openai_client(api_key: Optional[str] = None) -> OpenAI:
    """OpenAIクライアントを初期化する"""
    if not api_key:
        api_key = os.getenv('OPENAI_API_KEY')
        if not api_key:
            logger.error("OpenAI APIキーが設定されていません。環境変数または設定ファイルを確認してください。")
            raise ValueError("OpenAI APIキーが設定されていません。")
    client = OpenAI(api_key=api_key)
    return client

def get_ai_response(client: OpenAI, prompt: str, model: str = "o1-mini", temperature: float = 0.7, 
                   system_content: str = "あなたは仕様書を作成するAIです。") -> str:
    """OpenAI APIを使用してAI応答を生成"""
    try:
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": system_content},
                {"role": "user", "content": prompt}
            ],
            temperature=temperature
        )
        logger.info("AI応答の取得に成功しました。")
        return response.choices[0].message.content
    except Exception as e:
        logger.error(f"AI応答取得中にエラーが発生しました: {e}")
        return ""

class OpenAIConfig:
    """OpenAI設定を管理するクラス"""
    def __init__(self, api_key: Optional[str] = None, model: str = "gpt-4o", temperature: float = 0.7):
        self.api_key = api_key or os.getenv('OPENAI_API_KEY')
        if not self.api_key:
            raise ValueError("OpenAI APIキーが設定されていません。環境変数または設定ファイルを確認してください。")
        self.model = model
        self.temperature = temperature
        self.client = OpenAI(api_key=self.api_key)

    def get_response(self, prompt: str, system_content: str = "あなたは仕様書を作成するAIです。") -> str:
        """AI応答を取得"""
        return get_ai_response(
            self.client,
            prompt,
            self.model,
            self.temperature,
            system_content
        )