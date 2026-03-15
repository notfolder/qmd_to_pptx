"""
前処理器モジュール。

QMD/MD形式の入力テキストを受け取り、後段の共通パイプラインが期待する
QMD互換形式へ正規化して返す。
"""

from __future__ import annotations

import re


class Preprocessor:
    """
    前処理器クラス。

    QMDまたはMD形式のテキストを受け取り、YAMLフロントマター補完・
    Mermaid記法統一・コードブロック記法統一を適用して正規化する。
    """

    # Quartoのmermaid記法: ```{mermaid} を検出する正規表現
    _QUARTO_MERMAID_PATTERN: re.Pattern[str] = re.compile(
        r"^```\{mermaid\}",
        re.MULTILINE,
    )

    # Quartoのコードブロック記法: ```{言語名} または ```{.言語名} を検出する正規表現
    _QUARTO_CODE_PATTERN: re.Pattern[str] = re.compile(
        r"^```\{\.?([a-zA-Z0-9_\-]+)[^}]*\}",
        re.MULTILINE,
    )

    def normalize(self, text: str) -> str:
        """
        入力テキストをQMD互換形式に正規化して返す。

        処理順序:
        1. YAMLフロントマターが存在しない場合、空のフロントマターを付与する
        2. Quartoのmermaid記法（```{mermaid}）を```mermaid形式に置換する
        3. Quartoのコードブロック記法（```{言語名}）を```言語名形式に置換する

        Parameters
        ----------
        text : str
            入力QMD/MDテキスト。

        Returns
        -------
        str
            正規化済みQMDテキスト。
        """
        # 1. YAMLフロントマター補完
        text = self._ensure_frontmatter(text)
        # 2. Mermaid記法の統一
        text = self._normalize_mermaid(text)
        # 3. コードブロック記法の統一
        text = self._normalize_code_blocks(text)
        return text

    def _ensure_frontmatter(self, text: str) -> str:
        """
        テキスト先頭にYAMLフロントマターが存在しない場合、空のフロントマターを付与する。

        Parameters
        ----------
        text : str
            入力テキスト。

        Returns
        -------
        str
            YAMLフロントマターが補完されたテキスト。
        """
        if not text.startswith("---\n"):
            # title/author/date を空文字とした最小限のYAMLフロントマターを付与する
            frontmatter = "---\ntitle: ''\nauthor: ''\ndate: ''\n---\n"
            text = frontmatter + text
        return text

    def _normalize_mermaid(self, text: str) -> str:
        """
        Quartoのmermaid記法（```{mermaid}）を```mermaid形式に置換する。

        Parameters
        ----------
        text : str
            入力テキスト。

        Returns
        -------
        str
            Mermaid記法が統一されたテキスト。
        """
        return self._QUARTO_MERMAID_PATTERN.sub("```mermaid", text)

    def _normalize_code_blocks(self, text: str) -> str:
        """
        Quartoのコードブロック記法（```{言語名} または ```{.言語名}）を
        ```言語名形式に置換する。

        mermaidは既に変換済みのため、ここでは通常の言語名のみ対象とする。

        Parameters
        ----------
        text : str
            入力テキスト。

        Returns
        -------
        str
            コードブロック記法が統一されたテキスト。
        """
        return self._QUARTO_CODE_PATTERN.sub(r"```\1", text)
