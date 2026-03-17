"""
MCPサーバーのエントリーポイントモジュール。

FastMCPを使用して markdown_to_pptx および list_templates ツールを公開する。
stdioトランスポートでは標準出力はMCPプロトコル専用のため、
ログはすべて標準エラー出力（stderr）に出力する。
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
import tempfile
import urllib.request
import urllib.error
from contextlib import contextmanager
from typing import Iterator

from mcp.server.fastmcp import FastMCP

# MCPサーバーは非同期フレームワーク（FastMCP）上で動作するため、
# mermaid-parser-py の MermaidParser.parse() が内部で呼ぶ asyncio.run() が
# 既存イベントループに衝突して RuntimeWarning を引き起こす。
# nest_asyncio を適用することで、実行中のループ内でも asyncio.run() を
# ネスト呼び出しできるようにする（適用はサーバー起動モジュール内に限定する）。
import nest_asyncio
nest_asyncio.apply()

from . import render
from .template_registry import TemplateRegistry

# FastMCPサーバーインスタンスを生成する
mcp = FastMCP("qmd_to_pptx")


@contextmanager
def _resolve_template(template_id: str | None) -> Iterator[str | None]:
    """
    template_id をテンプレートファイルパスに解決するコンテキストマネージャー。

    template_id が URL（http:// または https:// で始まる）の場合は
    一時ファイルにダウンロードし、コンテキスト終了後に削除する。
    template_id が登録済みIDの場合はレジストリから解決する。
    template_id が None の場合は None を返す。

    Parameters
    ----------
    template_id : str | None
        テンプレートID、URL、または None。

    Yields
    ------
    str | None
        テンプレートPPTXファイルの絶対パス、または None。

    Raises
    ------
    ValueError
        template_id が未登録のIDの場合。
    urllib.error.URLError
        URL からのダウンロードに失敗した場合。
    """
    if template_id is None:
        yield None
        return

    # URL として認識する（http:// または https:// で始まる場合）
    if template_id.startswith("http://") or template_id.startswith("https://"):
        # 一時ファイルを .pptx 拡張子で作成する
        tmp_fd, tmp_path = tempfile.mkstemp(suffix=".pptx")
        os.close(tmp_fd)
        try:
            logging.info("テンプレートをダウンロード中: %s", template_id)
            urllib.request.urlretrieve(template_id, tmp_path)  # noqa: S310
            logging.info("ダウンロード完了: %s -> %s", template_id, tmp_path)
            yield tmp_path
        finally:
            # render 完了後（または例外発生時）に一時ファイルを削除する
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
                logging.info("一時テンプレートファイルを削除しました: %s", tmp_path)
        return

    # 登録済みテンプレートID として解決する
    registry = TemplateRegistry()
    reference_doc = registry.resolve(template_id)
    yield reference_doc


@mcp.tool()
def markdown_to_pptx(
    content: str,
    output: str,
    template_id: str | None = None,
) -> str:
    """
    QMDまたはMarkdownのテキストからPPTXファイルを生成する。

    Parameters
    ----------
    content : str
        QMDまたはMarkdownのテキスト文字列（ファイルパスではなくテキスト内容そのもの）。
    output : str
        出力先PPTXファイルパス（サーバー側ファイルシステム上のパス）。
    template_id : str | None
        テンプレートの指定。以下の形式を受け付ける。
        - config/templates.yaml に登録済みのテンプレートID
        - http:// または https:// で始まるPPTXテンプレートのURL
          （URLの場合は一時ファイルにダウンロードし、処理後に削除する）
        - 省略時はデフォルトレイアウトを使用する

    Returns
    -------
    str
        処理結果を示す文字列メッセージ。
    """
    try:
        with _resolve_template(template_id) as reference_doc:
            render(content, output, reference_doc)
        return f"PPTXファイルを生成しました: {output}"
    except ValueError as e:
        # 未登録のIDを指定した場合はエラーメッセージを返して変換しない
        return f"エラー: {e}"
    except urllib.error.URLError as e:
        return f"テンプレートのダウンロードに失敗しました: {e}"
    except Exception as e:
        logging.error("PPTX生成中にエラーが発生しました: %s", e, exc_info=True)
        return f"エラーが発生しました: {e}"


@mcp.tool()
def list_templates() -> str:
    """
    config/templates.yaml に登録済みのテンプレート一覧を返す。

    Returns
    -------
    str
        テンプレートIDと説明の一覧を示す文字列。
        テンプレートが登録されていない場合はその旨を示す文字列を返す。
    """
    registry = TemplateRegistry()
    templates = registry.list_templates()

    if not templates:
        return (
            "テンプレートが登録されていません。\n"
            "config/templates.yaml にテンプレートを追加するか、"
            "環境変数 QMD_TO_PPTX_TEMPLATES に設定ファイルのパスを指定してください。"
        )

    # {id: description} を読みやすい形式に整形する
    lines = ["登録済みテンプレート一覧:"]
    for tid, description in templates.items():
        lines.append(f"  - {tid}: {description}")
    return "\n".join(lines)


def _build_arg_parser() -> argparse.ArgumentParser:
    """コマンドライン引数パーサーを構築して返す。"""
    parser = argparse.ArgumentParser(
        description="qmd_to_pptx MCPサーバーを起動する。",
    )
    parser.add_argument(
        "--transport",
        choices=["stdio", "http"],
        default="stdio",
        help="トランスポート方式（デフォルト: stdio）",
    )
    parser.add_argument(
        "--host",
        default="0.0.0.0",
        help="HTTPモード時のバインドアドレス（デフォルト: 0.0.0.0）",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8000,
        help="HTTPモード時のポート番号（デフォルト: 8000）",
    )
    return parser


def main() -> None:
    """
    MCPサーバーのエントリーポイント関数。

    コマンドライン引数を解析してトランスポート方式に応じてサーバーを起動する。
    ログはすべて標準エラー出力（stderr）に出力する。
    """
    # ロギングハンドラーのストリームをstderrに設定する（stdout汚染防止）
    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logging.basicConfig(handlers=[handler], level=logging.INFO)

    parser = _build_arg_parser()
    args = parser.parse_args()

    if args.transport == "stdio":
        # stdioトランスポートでサーバーを起動する
        mcp.run(transport="stdio")
    else:
        # Streamable HTTPトランスポートでサーバーを起動する
        mcp.run(
            transport="streamable-http",
            host=args.host,
            port=args.port,
        )


if __name__ == "__main__":
    main()
