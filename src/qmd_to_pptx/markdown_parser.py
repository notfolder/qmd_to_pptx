"""
Markdownパーサーモジュール。

スライド分割器が生成した各スライドのMarkdownテキストをHTMLに変換し、
ElementTree形式のDOMツリーを生成する。
"""

from __future__ import annotations

import xml.etree.ElementTree as ET

import markdown
from pymdownx import superfences as sf


def _mermaid_fence_format(
    source: str,
    language: str,
    css_class: str,
    options: dict,
    md: markdown.Markdown,
    **kwargs: object,
) -> str:
    """
    Mermaidコードブロックをクラス属性 language-mermaid を持つ <code> 要素に変換する
    カスタムフェンスフォーマッター。

    pymdownx.superfences の custom_fences に登録して使用する。

    Parameters
    ----------
    source : str
        コードブロックのソーステキスト。
    language : str
        コードブロックの言語名（"mermaid"）。
    css_class : str
        CSSクラス名。
    options : dict
        追加オプション。
    md : markdown.Markdown
        Markdownインスタンス。
    **kwargs : object
        その他のキーワード引数。

    Returns
    -------
    str
        変換後のHTML文字列（<code class="language-mermaid">...</code>）。
    """
    # HTMLエスケープして <code class="language-mermaid"> で囲む
    escaped = (
        source
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )
    return f'<code class="language-mermaid">{escaped}</code>'


class MarkdownParser:
    """
    Markdownパーサークラス。

    pymdownx.superfences・pymdownx.arithmatex・tables・fenced_code の各extensionを
    適用してMarkdownをHTMLに変換し、ElementTree形式のDOMツリーを返す。
    """

    # 使用するMarkdown extension設定
    _EXTENSIONS: list = [
        "pymdownx.superfences",
        "pymdownx.arithmatex",
        "tables",
        "fenced_code",
        "md_in_html",
    ]

    def __init__(self) -> None:
        """Markdownパーサーの初期化。"""
        # カスタムフェンス設定（Mermaidブロックを専用フォーマッターで変換）
        self._extension_configs = {
            "pymdownx.superfences": {
                "custom_fences": [
                    {
                        "name": "mermaid",
                        "class": "language-mermaid",
                        "format": _mermaid_fence_format,
                    }
                ]
            },
            "pymdownx.arithmatex": {
                "generic": True,
            },
        }

    def parse(self, text: str) -> ET.Element:
        """
        MarkdownテキストをHTMLに変換してDOMツリーのルートElementを返す。

        Parameters
        ----------
        text : str
            スライド本文のMarkdownテキスト。

        Returns
        -------
        ET.Element
            DOMツリーのルート要素（<div>）。
        """
        # Markdownをインスタンス化して変換する
        md = markdown.Markdown(
            extensions=self._EXTENSIONS,
            extension_configs=self._extension_configs,
        )
        html_body = md.convert(text)

        # ElementTreeでパースするためにルート要素で囲む
        # HTMLエンティティを含む可能性があるためXMLとして処理できるよう変換する
        wrapped = f"<div>{html_body}</div>"
        try:
            root = ET.fromstring(wrapped)
        except ET.ParseError:
            # パース失敗時は空のdiv要素を返す
            root = ET.Element("div")

        return root
