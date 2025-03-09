# merge_files.py

import os
import argparse
from typing import Optional, List, Tuple
import logging

from utils import (
    setup_logger,
    read_settings,
    get_python_files,
    read_file_safely,
    write_file_content,
    ensure_directories_exist,
    generate_tree_structure
)

class PythonFileMerger:
    """
    Pythonファイルをマージするクラス。
    指定されたディレクトリからPythonファイルを収集し、
    フォルダ構造とファイル内容を統合した出力ファイルを生成する。
    """

    def __init__(self, settings_path: str = 'config/settings.ini', logger: Optional[logging.Logger] = None):
        """
        INI設定を読み込んでマージャーを初期化。

        Args:
            settings_path (str): 設定ファイルのパス。
            logger (Optional[logging.Logger]): 使用するロガー。指定がなければ新規作成。
        """
        self.logger = logger or setup_logger("PythonFileMerger", log_file="merge_files.log")
        self.settings = read_settings(settings_path)

        self.project_dir = os.path.abspath(self.settings['source_directory'])
        self.output_dir = os.path.join(self.project_dir, 'docs')
        self.output_filename = self.settings['output_file']

        self.logger.debug(f"Project directory: {self.project_dir}")
        self.logger.debug(f"Output directory: {self.output_dir}")
        self.logger.debug(f"Output filename: {self.output_filename}")

        ensure_directories_exist([self.output_dir])

        # 除外ディレクトリとファイルパターン
        self.exclude_dirs: List[str] = self.settings.get('exclusions', '').split(',')
        self.exclude_files: List[str] = ['*.log']

        self.logger.debug(f"Excluded directories: {self.exclude_dirs}")
        self.logger.debug(f"Excluded file patterns: {self.exclude_files}")

        self.logger.info("PythonFileMergerの初期化に成功しました。")

    def _generate_tree_structure(self) -> str:
        """
        anytreeを使用してディレクトリ構造を生成する。
        """
        return generate_tree_structure(self.project_dir, self.exclude_dirs, self.exclude_files)

    def _collect_python_files(self) -> List[Tuple[str, str]]:
        """
        プロジェクトディレクトリからPythonファイルを収集する。
        """
        self.logger.info(f"{self.project_dir} からPythonファイルを収集中...")
        python_files = get_python_files(self.project_dir, self.exclude_dirs)
        self.logger.debug(f"収集したPythonファイル: {python_files}")
        return python_files

    def _merge_files_content(self, python_files: List[Tuple[str, str]]) -> str:
        """
        Pythonファイルの内容をマージする。
        """
        merged_content = ""

        if not python_files:
            self.logger.warning("マージするPythonファイルが見つかりません。")
            return merged_content

        tree_structure = self._generate_tree_structure()
        merged_content += f"{tree_structure}\n\n# Merged Python Files\n\n"

        for rel_path, filepath in sorted(python_files):
            content = read_file_safely(filepath)
            if content is not None:
                merged_content += f"\n{'='*80}\nFile: {rel_path}\n{'='*80}\n\n{content}\n"
                self.logger.debug(f"ファイルをマージしました: {rel_path}")
            else:
                self.logger.warning(f"読み込みエラーのためスキップしました: {filepath}")

        return merged_content

    def _write_output(self, content: str) -> Optional[str]:
        """
        マージされた内容を出力ファイルに書き込む。
        """
        if not content:
            self.logger.warning("出力する内容がありません。")
            return None

        output_path = os.path.join(self.output_dir, self.output_filename)
        success = write_file_content(output_path, content)

        if success:
            self.logger.info(f"マージされた内容を正常に書き込みました: {output_path}")
            return output_path
        else:
            self.logger.error(f"マージされた内容の書き込みに失敗しました: {output_path}")
            return None

    def process(self) -> Optional[str]:
        """
        ファイルマージ処理を実行する。
        """
        try:
            python_files = self._collect_python_files()

            if not python_files:
                self.logger.warning("マージ処理を中止します。Pythonファイルが見つかりません。")
                return None

            merged_content = self._merge_files_content(python_files)
            return self._write_output(merged_content)
        except Exception as e:
            self.logger.error(f"ファイルマージ処理中にエラーが発生しました: {e}")
            return None


def merge_py_files(settings_path: str = 'config/settings.ini', logger: Optional[logging.Logger] = None) -> Optional[str]:
    """
    マージ処理のエントリーポイント。
    """
    logger = logger or setup_logger("merge_py_files", log_file="merge_py_files.log")
    logger.info("Pythonファイルのマージ処理を開始します。")
    try:
        merger = PythonFileMerger(settings_path=settings_path, logger=logger)
        return merger.process()
    except Exception as e:
        logger.error(f"マージ処理中に予期せぬエラーが発生しました: {e}")
        return None


def parse_arguments() -> argparse.Namespace:
    """
    コマンドライン引数を解析する。
    """
    parser = argparse.ArgumentParser(description="Merge Python files into a single output.")
    parser.add_argument(
        "--settings",
        type=str,
        default="config/settings.ini",
        help="Path to the settings.ini file"
    )
    return parser.parse_args()


def main():
    """
    Main function to execute the Python file merging process.
    """
    args = parse_arguments()
    log_dir = os.path.join(os.getcwd(), "\spec_tools\logs")  # "spec_tools" を削除
    logger = setup_logger("merge_files", log_dir=log_dir, log_file="merge_files.log")

    try:
        output_file = merge_py_files(settings_path=args.settings, logger=logger)
        if output_file:
            logger.info(f"Python files successfully merged. Output saved to: {output_file}")
        else:
            logger.error("File merging failed. Check logs for more details.")
    except Exception as e:
        logger.error(f"Unexpected error occurred: {e}")


if __name__ == "__main__":
    main()
