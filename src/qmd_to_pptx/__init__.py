"""
qmd_to_pptx パッケージのエントリーポイント。

外部に公開する render() 関数を定義する。

ロギング設定:
- 非MCPモード（直接呼び出し）: 標準出力（stdout）にログを出力する
- MCPモード: mcp_server.py の main() がロガーを標準エラー（stderr）に切り替える
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path

from .preprocessor import Preprocessor
from .slide_renderer import SlideRenderer

# パッケージロガーを取得し、デフォルトで標準出力（stdout）ハンドラーを設定する
# MCPモードでは mcp_server.py の main() がハンドラーをstderrに切り替える
_pkg_logger = logging.getLogger("qmd_to_pptx")
if not _pkg_logger.handlers:
    _stdout_handler = logging.StreamHandler(sys.stdout)
    _stdout_handler.setFormatter(logging.Formatter("%(levelname)s %(name)s: %(message)s"))
    _pkg_logger.addHandler(_stdout_handler)
    _pkg_logger.setLevel(logging.WARNING)
    # 親（ルート）ロガーへの伝播を無効にして二重出力を防止する
    _pkg_logger.propagate = False


def render(
    input: str,
    output: str,
    reference_doc: str | None = None,
) -> None:
    """
    QMDまたはMarkdownのファイルまたはテキストをPPTXに変換して保存する。

    inputがファイルシステム上の既存ファイルパスの場合はファイルを読み込む。
    そうでない場合は入力文字列をテキストとしてそのまま使用する。

    reference_docが指定された場合はそのテンプレートのデザインを継承する。
    YAMLフロントマターの format.pptx.reference-doc フィールドよりも
    reference_doc 引数の値を優先する。

    Parameters
    ----------
    input : str
        QMDまたはMarkdownのファイルパス、あるいはテキスト文字列。
    output : str
        出力先PPTXファイルのパス。
    reference_doc : str | None
        ベースとなるPPTXテンプレートファイルのパス（省略可）。
    """
    # inputがファイルパスとして有効かどうかを確認する
    input_path = Path(input)
    try:
        is_file = input_path.exists() and input_path.is_file()
    except OSError:
        # ファイル名が長すぎる場合などはテキストとして扱う
        is_file = False

    if is_file:
        text = input_path.read_text(encoding="utf-8")
    else:
        # ファイルでない場合はテキストとして直接使用する
        text = input

    # 前処理器でQMD互換形式に正規化する
    preprocessor = Preprocessor()
    normalized_text = preprocessor.normalize(text)

    # スライドレンダラーで全スライドを生成して保存する
    renderer = SlideRenderer()
    renderer.render_all(normalized_text, output, reference_doc)
