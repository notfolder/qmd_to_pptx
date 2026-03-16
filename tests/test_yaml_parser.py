"""
YAMLパーサー（YAMLParser）の単体テスト。
"""

import pytest
from qmd_to_pptx.yaml_parser import YAMLParser
from qmd_to_pptx.models import SlideMetadata


class TestYAMLParser:
    """YAMLParser クラスの単体テスト。"""

    def setup_method(self) -> None:
        """テスト前にYAMLParserインスタンスを生成する。"""
        self.parser = YAMLParser()

    def _make_text(self, yaml_body: str, body: str = "本文") -> str:
        """テスト用のフロントマター付きテキストを生成する。"""
        return f"---\n{yaml_body}\n---\n{body}"

    # --- title / author / date のテスト ---

    def test_parse_title(self) -> None:
        """titleフィールドを正しく解析する。"""
        text = self._make_text("title: テストタイトル")
        meta = self.parser.parse(text)
        assert meta.title == "テストタイトル"

    def test_parse_author(self) -> None:
        """authorフィールドを正しく解析する。"""
        text = self._make_text("author: テスト著者")
        meta = self.parser.parse(text)
        assert meta.author == "テスト著者"

    def test_parse_date(self) -> None:
        """dateフィールドを正しく解析する。"""
        text = self._make_text("date: 2026-03-15")
        meta = self.parser.parse(text)
        assert meta.date == "2026-03-15"

    def test_parse_empty_fields_use_defaults(self) -> None:
        """title/author/dateが欠落した場合は空文字をデフォルト値とする。"""
        text = self._make_text("")
        meta = self.parser.parse(text)
        assert meta.title == ""
        assert meta.author == ""
        assert meta.date == ""

    # --- theme のテスト ---

    def test_parse_theme(self) -> None:
        """themeフィールドを正しく解析する。"""
        text = self._make_text("theme: modern")
        meta = self.parser.parse(text)
        assert meta.theme == "modern"

    def test_parse_theme_default_empty(self) -> None:
        """themeが欠落した場合は空文字をデフォルト値とする。"""
        text = self._make_text("title: テスト")
        meta = self.parser.parse(text)
        assert meta.theme == ""

    # --- format.pptx.reference-doc のテスト ---

    def test_parse_reference_doc(self) -> None:
        """format.pptx.reference-doc を正しく解析する。"""
        text = self._make_text(
            "format:\n  pptx:\n    reference-doc: template.pptx"
        )
        meta = self.parser.parse(text)
        assert meta.reference_doc == "template.pptx"

    def test_parse_reference_doc_missing(self) -> None:
        """reference-docが欠落した場合はNoneを返す。"""
        text = self._make_text("title: テスト")
        meta = self.parser.parse(text)
        assert meta.reference_doc is None

    # --- format.pptx.incremental のテスト ---

    def test_parse_incremental_true(self) -> None:
        """format.pptx.incremental: true を正しく解析する。"""
        text = self._make_text("format:\n  pptx:\n    incremental: true")
        meta = self.parser.parse(text)
        assert meta.incremental is True

    def test_parse_incremental_false(self) -> None:
        """format.pptx.incremental: false を正しく解析する。"""
        text = self._make_text("format:\n  pptx:\n    incremental: false")
        meta = self.parser.parse(text)
        assert meta.incremental is False

    def test_parse_incremental_default_false(self) -> None:
        """incremental が欠落した場合は False をデフォルト値とする。"""
        text = self._make_text("title: テスト")
        meta = self.parser.parse(text)
        assert meta.incremental is False

    # --- slide-level のテスト ---

    def test_parse_slide_level_1(self) -> None:
        """slide-level: 1 を正しく解析する。"""
        text = self._make_text("slide-level: 1")
        meta = self.parser.parse(text)
        assert meta.slide_level == 1

    def test_parse_slide_level_2(self) -> None:
        """slide-level: 2 を正しく解析する。"""
        text = self._make_text("slide-level: 2")
        meta = self.parser.parse(text)
        assert meta.slide_level == 2

    def test_parse_slide_level_default_2(self) -> None:
        """slide-level が欠落した場合は 2 をデフォルト値とする。"""
        text = self._make_text("title: テスト")
        meta = self.parser.parse(text)
        assert meta.slide_level == 2

    def test_parse_slide_level_invalid_falls_back_to_2(self) -> None:
        """無効な slide-level 値は 2 にフォールバックする。"""
        text = self._make_text("slide-level: invalid")
        meta = self.parser.parse(text)
        assert meta.slide_level == 2

    def test_parse_slide_level_out_of_range_falls_back_to_2(self) -> None:
        """有効範囲外の slide-level 値は 2 にフォールバックする。"""
        text = self._make_text("slide-level: 3")
        meta = self.parser.parse(text)
        assert meta.slide_level == 2

    # --- フロントマターなし（前処理器で補完済み想定） ---

    def test_parse_returns_defaults_when_no_frontmatter(self) -> None:
        """フロントマターが存在しない場合はデフォルト値のSlideMetadataを返す。"""
        text = "本文のみのテキスト"
        meta = self.parser.parse(text)
        assert isinstance(meta, SlideMetadata)
        assert meta.title == ""
        assert meta.slide_level == 2
