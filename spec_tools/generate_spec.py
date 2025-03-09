#generate_spec.py
import os
import logging
from typing import Optional
from utils import read_file_safely, write_file_content, OpenAIConfig, setup_logger
from icecream import ic
from dotenv import load_dotenv
from pathlib import Path
from datetime import datetime

# .envファイルを読み込む
load_dotenv(dotenv_path="config/secrets.env")


class SpecificationGenerator:
    """仕様書生成を管理するクラス"""

    def __init__(self, logger: logging.Logger):
        """設定を読み込んで初期化"""
        self.logger = logger
        try:
            # OpenAI設定の初期化
            # OpenAIConfigクラスで既にAPIキーのチェックが行われる
            self.ai_config = OpenAIConfig(
                model="gpt-4o",
                temperature=0.7
            )

            # ソースディレクトリの設定
            self.source_dir = os.path.abspath(".")
            self.document_dir = os.path.join(self.source_dir, 'docs')
            self.spec_tools_dir = os.path.join(self.source_dir, 'spec_tools/prompt')
            self.prompt_file = os.path.join(self.spec_tools_dir, 'prompt_requirements_spec.txt')
            ic(self.source_dir, self.document_dir, self.prompt_file)

            self.logger.debug("SpecificationGenerator initialized")
        except ValueError as ve:
            self.logger.error(str(ve))
            self._write_api_key_error()
            raise
        except Exception as e:
            self.logger.error(f"SpecificationGeneratorの初期化に失敗しました: {e}")
            raise

    def generate(self) -> str:
        """仕様書を生成してファイルに保存"""
        try:
            code_content = self._read_merge_file()
            ic(code_content)
            if not code_content:
                self.logger.error("コード内容が空です。")
                return ""

            prompt = self._read_prompt_file()
            ic(prompt)
            if not prompt:
                self.logger.error("プロンプトファイルの読み込みに失敗しました。")
                return ""

            full_prompt = f"{prompt}\n\nコード:\n{code_content}"
            ic(full_prompt)

            # OpenAIConfigを使用してAI応答を取得
            specification = self.ai_config.get_response(full_prompt)
            ic(specification)
            if not specification:
                return ""

            output_path = os.path.join(self.document_dir, 'requirements_spec.txt')
            ic(output_path)
            if write_file_content(output_path, specification):
                self.logger.info(f"仕様書が正常に出力されました: {output_path}")
                self.update_readme()
                return output_path
            return ""
        except Exception as e:
            self.logger.error(f"仕様書生成中にエラーが発生しました: {e}")
            return ""

    def _read_merge_file(self) -> str:
        """merge.txt ファイルの内容を読み込む"""
        merge_path = os.path.join(self.document_dir, 'merge.txt')
        ic(merge_path)  # デバッグ: merge.txtのパスを確認
        content = read_file_safely(merge_path)
        if content:
            self.logger.info("merge.txt の読み込みに成功しました。")
        else:
            self.logger.error("merge.txt の読み込みに失敗しました。")
        return content or ""

    def _read_prompt_file(self) -> str:
        """プロンプトファイル (prompt_requirements_spec.txt) を読み込む"""
        ic(self.prompt_file)  # デバッグ: プロンプトファイルのパスを確認
        try:
            with open(self.prompt_file, 'r', encoding='utf-8') as f:
                content = f.read()
                self.logger.info("prompt_requirements_spec.txt の読み込みに成功しました。")
                return content
        except FileNotFoundError:
            self.logger.error(f"プロンプトファイルが見つかりません: {self.prompt_file}")
        except UnicodeDecodeError as e:
            self.logger.error(f"エンコードエラー: {e}")
        except Exception as e:
            self.logger.error(f"プロンプトファイルの読み込み中にエラーが発生しました: {e}")
        return ""

    def _write_api_key_error(self) -> None:
        """APIキーが設定されていない場合のエラーメッセージをファイルに出力"""
        error_message = """
# OpenAI APIキー設定エラー

## エラー内容
OpenAI APIキーが設定されていないため、仕様書を生成できません。

## 解決方法
以下のいずれかの方法でAPIキーを設定してください：

1. 環境変数での設定:
   - OPENAI_API_KEY環境変数にAPIキーを設定

2. secrets.envファイルでの設定:
   - config/secrets.env ファイルを作成
   - 以下の形式でOpenAI APIキーを設定：
     ```
     OPENAI_API_KEY=sk-あなたのAPIキー
     ```

3. settings.iniファイルでの設定:
   - config/settings.ini ファイルのAPI セクションに設定：
     ```
     [API]
     openai_api_key=sk-あなたのAPIキー
     ```

## 補足
- APIキーはOpenAIのウェブサイトから取得できます
- APIキーは機密情報のため、GitHubなどに公開しないように注意してください
- .envファイルと設定ファイルは.gitignoreに追加することを推奨します
"""
        output_path = os.path.join(self.document_dir, 'requirements_spec.txt')
        write_file_content(output_path, error_message)
        self.logger.info(f"APIキー設定エラーメッセージを出力しました: {output_path}")

        # README.mdも更新
        self.update_readme(error_content=error_message)

    def update_readme(self, error_content: Optional[str] = None) -> None:
        """README.md を README_tmp.md のテンプレートを使用して更新する"""
        try:
            readme_tmp_path = os.path.join(self.spec_tools_dir, 'README_tmp.md')
            readme_path = os.path.join(self.source_dir, 'README.md')

            self.logger.info(f"README.md を更新するためにテンプレートを読み込みます: {readme_tmp_path}")

            # README_tmp.md の内容を読み込む
            template_content = read_file_safely(readme_tmp_path)
            if not template_content:
                self.logger.error(f"{readme_tmp_path} の読み込みに失敗しました。README.md は更新されませんでした。")
                return

            if error_content:
                # エラーの場合は、エラーメッセージで置換
                updated_content = template_content.replace("[spec]", error_content)
                updated_content = updated_content.replace("[tree]", "[APIキーが設定されていないため、フォルダ構成を取得できません。]")
            else:
                # 通常の更新処理
                # [spec] プレースホルダーを requirements_spec.txt の内容で置換
                spec_path = os.path.join(self.document_dir, 'requirements_spec.txt')
                spec_content = read_file_safely(spec_path)
                if not spec_content:
                    self.logger.error(f"{spec_path} の読み込みに失敗しました。README.md の [spec] プレースホルダーは置換されませんでした。")
                    spec_content = "[仕様書の内容が取得できませんでした。]"
                updated_content = template_content.replace("[spec]", spec_content)

                # [tree] プレースホルダーを merge.txt の内容で置換
                merge_path = os.path.join(self.document_dir, 'merge.txt')
                merge_content = read_file_safely(merge_path)
                if not merge_content:
                    self.logger.error(f"{merge_path} の読み込みに失敗しました。README.md の [tree] プレースホルダーは置換されませんでした。")
                    tree_content = "[フォルダ構成の取得に失敗しました。]"
                else:
                    tree_content = merge_content.strip()
                updated_content = updated_content.replace("[tree]", f"```\n{tree_content}\n```")

            # 現在の日付を挿入
            current_date = datetime.now().strftime("%Y-%m-%d")
            updated_content = updated_content.replace("[YYYY-MM-DD]", current_date)

            # README.md に書き込む
            if write_file_content(readme_path, updated_content):
                self.logger.info(f"README.md が正常に更新されました: {readme_path}")
            else:
                self.logger.error(f"README.md の更新に失敗しました: {readme_path}")

        except Exception as e:
            self.logger.error(f"README.md の更新中にエラーが発生しました: {e}")
            raise


def generate_specification():
    """仕様書を生成"""
    logger = setup_logger("generate_spec", log_dir="\spec_tools\logs")  # ロガーを設定
    try:
        generator = SpecificationGenerator(logger=logger)
        output_file = generator.generate()
        if output_file:
            logger.info(f"Specification generated successfully. Output saved to: {output_file}")
        else:
            logger.error("Specification generation failed. Check logs for details.")
    except ValueError as ve:
        logger.error(f"APIキーエラー: {ve}")
    except Exception as e:
        logger.error(f"An unexpected error occurred: {e}")


if __name__ == "__main__":
    generate_specification()