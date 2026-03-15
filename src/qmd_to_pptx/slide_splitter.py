"""
スライド分割器モジュール。

QMD本文テキストをスライド単位に分割し、各スライドのSlideContentリストを生成する。
"""

from __future__ import annotations

import re

from .models import SeparatorType, SlideContent


class SlideSplitter:
    """
    スライド分割器クラス。

    Quartoの仕様に従い、レベル1/2見出しおよび水平区切り線でテキストを分割する。
    """

    # YAMLフロントマターブロックを検出する正規表現（テキスト先頭の---から次の---まで）
    _FRONTMATTER_PATTERN: re.Pattern[str] = re.compile(
        r"^---\n.*?\n---\n",
        re.DOTALL,
    )

    # 見出し行に付与された属性ブロック（{...}）を検出する正規表現
    # 例: ## Slide Title {background-image="background.png"}
    _ATTR_PATTERN: re.Pattern[str] = re.compile(
        r"\{([^}]*)\}\s*$",
    )

    # background-image 属性の値を取得する正規表現
    _BG_IMAGE_PATTERN: re.Pattern[str] = re.compile(
        r'background-image\s*=\s*["\']([^"\']+)["\']',
    )

    def split(self, text: str, slide_level: int) -> list[SlideContent]:
        """
        テキストをスライド単位に分割し、SlideContentのリストを返す。

        Parameters
        ----------
        text : str
            正規化済みQMDテキスト（YAMLフロントマターを含む）。
        slide_level : int
            スライド区切りとして扱う見出しレベル（1または2）。

        Returns
        -------
        list[SlideContent]
            各スライドの内容を表すSlideContentのリスト。
        """
        # YAMLフロントマターブロックを除去して本文テキストを取得する
        body = self._FRONTMATTER_PATTERN.sub("", text, count=1)

        # slide_levelに基づいてスライド区切りパターンを決定する
        if slide_level == 1:
            # slide-level: 1 の場合、# 見出しのみがスライド区切り（##は本文の見出しとして扱う）
            separator_pattern = re.compile(
                r"^(# .+|---)\s*$",
                re.MULTILINE,
            )
        else:
            # slide-level: 2 (デフォルト) の場合、# と ## 見出し、および --- が区切り
            separator_pattern = re.compile(
                r"^(#{1,2} .+|---)\s*$",
                re.MULTILINE,
            )

        # テキストを区切り行で分割する
        parts: list[SlideContent] = []
        last_end = 0
        pending_sep_type: SeparatorType | None = None
        pending_title: str = ""
        pending_bg: str | None = None

        for match in separator_pattern.finditer(body):
            chunk = body[last_end:match.start()].strip()

            # 最初の区切り前のテキストは本文として保留しておく
            if pending_sep_type is not None:
                parts.append(SlideContent(
                    body_text=chunk,
                    separator_type=pending_sep_type,
                    title=pending_title,
                    background_image=pending_bg,
                ))

            sep_line = match.group(1)
            pending_sep_type, pending_title, pending_bg = self._parse_separator(
                sep_line, slide_level
            )
            last_end = match.end()

        # 最後のチャンクを追加する
        if pending_sep_type is not None:
            last_chunk = body[last_end:].strip()
            parts.append(SlideContent(
                body_text=last_chunk,
                separator_type=pending_sep_type,
                title=pending_title,
                background_image=pending_bg,
            ))

        return parts

    def _parse_separator(
        self,
        sep_line: str,
        slide_level: int,
    ) -> tuple[SeparatorType, str, str | None]:
        """
        区切り行を解析してSeparatorType・タイトル・背景画像パスを返す。

        Parameters
        ----------
        sep_line : str
            区切り行（見出し行または水平区切り線）。
        slide_level : int
            スライド区切りとして扱う見出しレベル。

        Returns
        -------
        tuple[SeparatorType, str, str | None]
            (SeparatorType, タイトル文字列, 背景画像パスまたはNone) のタプル。
        """
        if sep_line == "---":
            return SeparatorType.RULER, "", None

        # 見出し行の解析: ## Slide Title {background-image="background.png"}
        heading_match = re.match(r"^(#+)\s+(.*)", sep_line)
        if not heading_match:
            return SeparatorType.RULER, "", None

        level = len(heading_match.group(1))
        raw_title = heading_match.group(2).strip()

        # 属性ブロック（{...}）を検出して背景画像パスを取得し、タイトルから除去する
        background_image: str | None = None
        attr_match = self._ATTR_PATTERN.search(raw_title)
        if attr_match:
            attr_str = attr_match.group(1)
            bg_match = self._BG_IMAGE_PATTERN.search(attr_str)
            if bg_match:
                background_image = bg_match.group(1)
            # タイトルから属性ブロックを除去する
            raw_title = raw_title[:attr_match.start()].strip()

        # slide_levelに基づいてSeparatorTypeを決定する
        if slide_level == 1:
            # slide-level: 1 の場合、# 見出しのみがスライド区切り（正規表現で ## は除外済み）
            if level == 1:
                sep_type = SeparatorType.HEADING1
            else:
                sep_type = SeparatorType.HEADING2
        else:
            # slide-level: 2 (デフォルト) の場合
            if level == 1:
                sep_type = SeparatorType.HEADING1
            else:
                sep_type = SeparatorType.HEADING2

        return sep_type, raw_title, background_image
