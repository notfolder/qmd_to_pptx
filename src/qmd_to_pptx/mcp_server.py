"""
MCPサーバーのエントリーポイントモジュール。

FastMCPを使用して markdown_to_pptx および list_templates ツールを公開する。
stdioトランスポートでは標準出力はMCPプロトコル専用のため、
ログはすべて標準エラー出力（stderr）に出力する。
"""

from __future__ import annotations

import argparse
import logging
import sys

from mcp.server.fastmcp import FastMCP

from . import render
from .template_registry import TemplateRegistry

# FastMCPサーバーインスタンスを生成する
mcp = FastMCP("qmd_to_pptx")


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
        config/templates.yaml に登録済みのテンプレートID。
        省略時はデフォルトレイアウトを使用する。

    Returns
    -------
    str
        処理結果を示す文字列メッセージ。
    """
    # template_id が指定された場合はレジストリからPPTXパスを解決する
    reference_doc: str | None = None
    if template_id is not None:
        registry = TemplateRegistry()
        try:
            reference_doc = registry.resolve(template_id)
        except ValueError as e:
            # 未登録のIDを指定した場合はエラーメッセージを返して変換しない
            return f"エラー: {e}"

    try:
        render(content, output, reference_doc)
        return f"PPTXファイルを生成しました: {output}"
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
