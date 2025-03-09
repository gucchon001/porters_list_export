#generate_detailed_spec.py
import os
from utils import setup_logger  # setup_logger をインポート
from dotenv import load_dotenv
from icecream import ic
from utils import read_file_safely, write_file_content, OpenAIConfig

# 環境変数をロード
load_dotenv(dotenv_path="config/secrets.env")

# ロガーを設定
logger = setup_logger("generate_detailed_spec", log_file="\spec_tools\generate_detailed_spec.log")

class SpecificationGenerator:
    """仕様書生成を管理するクラス"""

    def __init__(self):
        """設定を読み込んで初期化"""
        try:
            # OpenAIConfigを使用してAI関連の設定を初期化
            self.ai_config = OpenAIConfig(
                model="gpt-4o",
                temperature=0.7
            )

            # ディレクトリ設定
            self.source_dir = os.path.abspath(".")
            self.document_dir = os.path.join(self.source_dir, 'docs')
            self.prompt_file = os.path.join(self.source_dir, 'spec_tools', 'prompt', 'prompt_generate_detailed_spec.txt')
            ic(self.source_dir, self.document_dir, self.prompt_file)  # デバッグ: パスを確認

            logger.debug("SpecificationGenerator initialized")
        except Exception as e:
            logger.error(f"SpecificationGeneratorの初期化に失敗しました: {e}")
            raise

    def generate(self) -> str:
        """仕様書を生成してファイルに保存"""
        try:
            # merge.txtの内容を読み込む
            code_content = self._read_merge_file()
            ic(code_content)  # デバッグ: merge.txtの内容を確認
            if not code_content:
                logger.error("コード内容が空です。")
                return ""

            # プロンプトファイルを読み込む
            prompt = self._read_prompt_file()
            ic(prompt)  # デバッグ: プロンプト内容を確認
            if not prompt:
                logger.error("プロンプトファイルの読み込みに失敗しました。")
                return ""

            # プロンプトの最終形を作成
            full_prompt = f"{prompt}\n\nコード:\n{code_content}"
            ic(full_prompt)  # デバッグ: 完成したプロンプトを確認

            # OpenAIConfigを使用してAI応答を取得
            specification = self.ai_config.get_response(full_prompt)
            ic(specification)  # デバッグ: 生成された仕様書を確認
            if not specification:
                return ""

            # 出力ファイルのパスを設定
            output_path = os.path.join(self.document_dir, 'detail_spec.txt')
            ic(output_path)  # デバッグ: 出力パスを確認
            if write_file_content(output_path, specification):
                logger.info(f"仕様書が正常に出力されました: {output_path}")
                return output_path
            return ""
        except Exception as e:
            logger.error(f"仕様書生成中にエラーが発生しました: {e}")
            return ""

    def _read_merge_file(self) -> str:
        """merge.txt ファイルの内容を読み込む"""
        merge_path = os.path.join(self.document_dir, 'merge.txt')
        ic(merge_path)  # デバッグ: merge.txtのパスを確認
        content = read_file_safely(merge_path)
        if content:
            logger.info("merge.txt の読み込みに成功しました。")
        else:
            logger.error("merge.txt の読み込みに失敗しました。")
        return content or ""

    def _read_prompt_file(self) -> str:
        """プロンプトファイルを読み込む"""
        ic(self.prompt_file)  # デバッグ: プロンプトファイルのパスを確認
        content = read_file_safely(self.prompt_file)
        if content:
            logger.info("prompt_generate_detailed_spec.txt の読み込みに成功しました。")
        else:
            logger.error(f"プロンプトファイルの読み込みに失敗しました: {self.prompt_file}")
        return content or ""

def generate_detailed_specification():
    """詳細仕様書を生成"""
    try:
        generator = SpecificationGenerator()
        output_file = generator.generate()
        if output_file:
            logger.info(f"Detailed specification generated successfully. Output saved to: {output_file}")
        else:
            logger.error("Detailed specification generation failed. Check logs for details.")
    except Exception as e:
        logger.error(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    generate_detailed_specification()
