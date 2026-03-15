"""
MCPサーバーのエントリーポイントモジュール。

FastMCPを使用してmarkdown_to_pptxツールを公開する。
stdioトランスポートでは標準出力はMCPプロトコル専用のため、
ログはすべて標準エラー出力（stderr）に出力する。
"""

from __future__ import annotations

import argparse
import logging
import sys

from mcp.server.fastmcp import FastMCP

from . import render

# FastMCPサーバーインスタンスを生成する
mcp = FastMCP("qmd_to_pptx")


@mcp.tool()
def markdown_to_pptx(
    content: str,
    output: str,
    reference_doc: str | None = None,
) -> str:
    """
    QMDまたはMarkdownのテキストからPPTXファイルを生成する。

    Parameters
    ----------
    content : str
        QMDまたはMarkdownのテキスト文字列（ファイルパスではなくテキスト内容そのもの）。
    output : str
        出力先PPTXファイルパス（サーバー側ファイルシステム上のパス）。
    reference_doc : str | None
        ベースとなるPPTXテンプレートのパス。省略時はデフォルトレイアウトを使用する。

    Returns
    -------
    str
        処理結果を示す文字列メッセージ。
    """
    try:
        render(content, output, reference_doc)
        return f"PPTXファイルを生成しました: {output}"
    except Exception as e:
        logging.error("PPTX生成中にエラーが発生しました: %s", e, exc_info=True)
        return f"エラーが発生しました: {e}"


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
