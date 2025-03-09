# tests/conftest.py

import sys
from pathlib import Path

# プロジェクトルートのパスを取得
project_root = Path(__file__).resolve().parent.parent

# src ディレクトリをPYTHONPATHに追加
src_path = project_root / 'src'
if src_path not in sys.path:
    sys.path.insert(0, str(src_path))
