"""
前処理器モジュール。

QMD/MD形式の入力テキストを受け取り、後段の共通パイプラインが期待する
QMD互換形式へ正規化して返す。
"""

from __future__ import annotations

import html
import re


class Preprocessor:
    """
    前処理器クラス。

    QMDまたはMD形式のテキストを受け取り、YAMLフロントマター補完・
    Mermaid記法統一・コードブロック記法統一・フェンスドdiv変換を適用して正規化する。
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

    # フェンスドdiv開始行: ::: {.class} または :::: {.class} を検出する正規表現
    _FENCED_DIV_OPEN_PATTERN: re.Pattern[str] = re.compile(
        r"^(:{3,})\s*\{([^}]*)\}\s*$"
    )

    # フェンスドdiv終了行: ::: のみの行を検出する正規表現（先頭コロン数で終端を判定）
    _FENCED_DIV_CLOSE_PATTERN: re.Pattern[str] = re.compile(
        r"^(:{3,})\s*$"
    )

    def normalize(self, text: str) -> str:
        """
        入力テキストをQMD互換形式に正規化して返す。

        処理順序:
        1. YAMLフロントマターが存在しない場合、空のフロントマターを付与する
        2. Quartoのmermaid記法（```{mermaid}）を```mermaid形式に置換する
        3. Quartoのコードブロック記法（```{言語名}）を```言語名形式に置換する
        4. Quartoのフェンスドdiv記法（::: {.class}）をHTML<div>タグに変換する

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
        # 4. フェンスドdiv記法をHTMLの<div>タグに変換する
        text = self._normalize_fenced_divs(text)
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

    def _normalize_fenced_divs(self, text: str) -> str:
        """
        Quartoのフェンスドdiv記法（::: {.class}）をHTMLの<div class="class">タグに変換する。

        コードブロック（```）内の `::: {.class}` 行は変換対象外とする。
        ネストしたフェンスドdiv（:::: 内の ::: など）も再帰的に正しく変換する。

        例:
          ::: {.notes}            →   <div class="notes">
          ノート内容              →   ノート内容
          :::                     →   </div>

        Parameters
        ----------
        text : str
            入力テキスト。

        Returns
        -------
        str
            フェンスドdivがHTMLに変換されたテキスト。
        """
        lines = text.split("\n")
        result, _ = self._parse_lines_with_fenced_divs(lines, 0, closing_colons=None)
        return "\n".join(result)

    def _parse_lines_with_fenced_divs(
        self,
        lines: list[str],
        start: int,
        closing_colons: str | None,
    ) -> tuple[list[str], int]:
        """
        行リストを走査してフェンスドdivをHTMLの<div>タグに変換する。

        コードブロック（```）内の `::: {.class}` 行は変換対象外とする。
        ネストしたdivは再帰的に処理する。

        Parameters
        ----------
        lines : list[str]
            入力テキストを改行で分割した行リスト。
        start : int
            走査開始インデックス。
        closing_colons : str | None
            現在処理中のdivブロックを閉じるコロン文字列。
            トップレベル呼び出し時は None。

        Returns
        -------
        tuple[list[str], int]
            (変換後の行リスト, 次の走査開始インデックス) のタプル。
        """
        result: list[str] = []
        i = start
        in_code_block = False

        while i < len(lines):
            line = lines[i]

            # コードブロック（```）の開始・終了を追跡してフェンスドdivの誤変換を防ぐ
            if line.startswith("```"):
                in_code_block = not in_code_block
                result.append(line)
                i += 1
                continue

            if not in_code_block:
                # 現在のdivブロックを閉じる行かどうかを確認する
                if closing_colons is not None:
                    close_m = self._FENCED_DIV_CLOSE_PATTERN.match(line)
                    if close_m and close_m.group(1) == closing_colons:
                        # 閉じタグを追加せずに呼び出し元に制御を返す
                        return result, i + 1

                # フェンスドdiv開始行かどうかを確認する
                open_m = self._FENCED_DIV_OPEN_PATTERN.match(line)
                if open_m:
                    colons = open_m.group(1)
                    attr_str = open_m.group(2)
                    class_names = self._extract_class_names(attr_str)
                    # markdown="1" 属性を付与して md_in_html 拡張が内部Markdownを処理できるようにする
                    # class_names をHTMLエスケープして不正な属性値を防ぐ
                    safe_class = html.escape(class_names, quote=True)
                    result.append(f'<div class="{safe_class}" markdown="1">')
                    i += 1
                    # ネストした内容を再帰的に処理する
                    inner_lines, i = self._parse_lines_with_fenced_divs(
                        lines, i, closing_colons=colons
                    )
                    result.extend(inner_lines)
                    result.append("</div>")
                    continue

            result.append(line)
            i += 1

        return result, i

    @staticmethod
    def _extract_class_names(attr_str: str) -> str:
        """
        Quartoの属性文字列（例: `.notes`, `.column width="50%"`）から
        CSSクラス名を抽出してスペース区切りの文字列で返す。

        Parameters
        ----------
        attr_str : str
            フェンスドdivの属性文字列（`{` `}` を除いた部分）。

        Returns
        -------
        str
            スペース区切りのCSSクラス名文字列。クラス名がない場合は属性文字列をトリムして返す。
        """
        classes = re.findall(r"\.([a-zA-Z0-9_-]+)", attr_str)
        return " ".join(classes) if classes else attr_str.strip()
