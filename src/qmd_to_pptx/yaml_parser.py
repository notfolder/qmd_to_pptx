"""
YAMLパーサーモジュール。

QMDファイル先頭のYAMLフロントマターブロックを抽出し、
メタデータとしてSlideMetadataに変換する。
"""

from __future__ import annotations

import logging
import re

import yaml

from .models import SlideMetadata

# モジュールロガーを取得する
logger = logging.getLogger(__name__)


class YAMLParser:
    """
    YAMLパーサークラス。

    QMDテキストの先頭YAMLフロントマターブロック（---で囲まれた領域）を
    抽出し、SlideMetadataとして返す。
    """

    # YAMLフロントマターブロックを検出する正規表現（テキスト先頭の---から次の---まで）
    _FRONTMATTER_PATTERN: re.Pattern[str] = re.compile(
        r"^---\n(.*?)\n---\n",
        re.DOTALL,
    )

    def parse(self, text: str) -> SlideMetadata:
        """
        テキスト先頭のYAMLフロントマターブロックを抽出し、SlideMetadataを返す。

        Parameters
        ----------
        text : str
            正規化済みQMDテキスト（前処理器により必ずフロントマターが存在する）。

        Returns
        -------
        SlideMetadata
            解析したメタデータオブジェクト。フィールドが欠落している場合はデフォルト値を使用する。
        """
        match = self._FRONTMATTER_PATTERN.match(text)
        if not match:
            # 前処理器によりフロントマターは必ず存在するが、念のためデフォルト値を返す
            return SlideMetadata()

        yaml_text = match.group(1)
        try:
            data: dict = yaml.safe_load(yaml_text) or {}
        except yaml.YAMLError:
            data = {}

        # title / author / date の取得（欠落時は空文字）
        title = str(data.get("title", "") or "")
        author = str(data.get("author", "") or "")
        date = str(data.get("date", "") or "")
        theme = str(data.get("theme", "") or "")

        # format.pptx.reference-doc の取得（階層を辿る）
        reference_doc: str | None = None
        fmt = data.get("format")
        if isinstance(fmt, dict):
            pptx = fmt.get("pptx")
            if isinstance(pptx, dict):
                ref = pptx.get("reference-doc")
                if ref:
                    reference_doc = str(ref)

        # format.pptx.incremental の取得
        incremental: bool = False
        if isinstance(fmt, dict):
            pptx = fmt.get("pptx")
            if isinstance(pptx, dict):
                inc = pptx.get("incremental")
                if isinstance(inc, bool):
                    incremental = inc

        # slide-level の取得（省略時は2）
        slide_level_raw = data.get("slide-level", 2)
        try:
            slide_level = int(slide_level_raw)
        except (TypeError, ValueError):
            logger.warning(
                "slide-level の値 %r を整数に変換できませんでした。2 にフォールバックします。",
                slide_level_raw,
            )
            slide_level = 2
        if slide_level not in (1, 2):
            logger.warning(
                "slide-level の値 %r は有効範囲外（1 または 2）です。2 にフォールバックします。",
                slide_level,
            )
            slide_level = 2

        return SlideMetadata(
            title=title,
            author=author,
            date=date,
            theme=theme,
            reference_doc=reference_doc,
            incremental=incremental,
            slide_level=slide_level,
        )
