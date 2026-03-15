"""
スライド分割器（SlideSplitter）の単体テスト。
"""

import pytest
from qmd_to_pptx.slide_splitter import SlideSplitter
from qmd_to_pptx.models import SeparatorType, SlideContent


# テスト用のYAMLフロントマター
_FRONTMATTER = "---\ntitle: テスト\nauthor: ''\ndate: ''\n---\n"


class TestSlideSplitter:
    """SlideSplitter クラスの単体テスト。"""

    def setup_method(self) -> None:
        """テスト前にSlideSplitterインスタンスを生成する。"""
        self.splitter = SlideSplitter()

    # --- slide-level: 2 (デフォルト) のテスト ---

    def test_split_heading2_creates_slide(self) -> None:
        """## 見出しでスライドを分割する（slide-level: 2）。"""
        text = _FRONTMATTER + "## スライドタイトル\n\nコンテンツ\n"
        slides = self.splitter.split(text, slide_level=2)
        assert len(slides) == 1
        assert slides[0].separator_type == SeparatorType.HEADING2
        assert slides[0].title == "スライドタイトル"
        assert "コンテンツ" in slides[0].body_text

    def test_split_heading1_creates_section_header_slide(self) -> None:
        """# 見出しで Section Header スライドを分割する（slide-level: 2）。"""
        text = _FRONTMATTER + "# セクション名\n\n## スライド\n\n本文\n"
        slides = self.splitter.split(text, slide_level=2)
        assert slides[0].separator_type == SeparatorType.HEADING1
        assert slides[0].title == "セクション名"

    def test_split_ruler_creates_slide(self) -> None:
        """--- 水平区切り線でタイトルなしスライドを生成する。"""
        text = _FRONTMATTER + "---\n\n本文テキスト\n"
        slides = self.splitter.split(text, slide_level=2)
        assert len(slides) == 1
        assert slides[0].separator_type == SeparatorType.RULER
        assert slides[0].title == ""

    def test_split_multiple_slides(self) -> None:
        """複数のスライドを正しく分割する。"""
        text = (
            _FRONTMATTER
            + "# セクション\n\n## スライド1\n\n内容1\n\n## スライド2\n\n内容2\n"
        )
        slides = self.splitter.split(text, slide_level=2)
        assert len(slides) == 3
        assert slides[0].separator_type == SeparatorType.HEADING1
        assert slides[1].separator_type == SeparatorType.HEADING2
        assert slides[2].separator_type == SeparatorType.HEADING2

    # --- slide-level: 1 のテスト ---

    def test_split_level1_heading1_creates_content_slide(self) -> None:
        """slide-level: 1 の場合、# 見出しがスライド区切りになる。"""
        text = _FRONTMATTER + "# スライドタイトル\n\nコンテンツ\n"
        slides = self.splitter.split(text, slide_level=1)
        assert len(slides) == 1
        assert slides[0].separator_type == SeparatorType.HEADING1
        assert slides[0].title == "スライドタイトル"

    def test_split_level1_heading2_is_not_separator(self) -> None:
        """slide-level: 1 の場合、## 見出しはスライド区切りにならず本文として扱われる。"""
        text = _FRONTMATTER + "# スライドタイトル\n\n## 本文の見出し\n\nコンテンツ\n"
        slides = self.splitter.split(text, slide_level=1)
        assert len(slides) == 1
        assert "## 本文の見出し" in slides[0].body_text

    def test_split_level1_multiple_slides(self) -> None:
        """slide-level: 1 の場合、# のみで複数スライドを分割する。"""
        text = _FRONTMATTER + "# スライド1\n\n内容1\n\n# スライド2\n\n内容2\n"
        slides = self.splitter.split(text, slide_level=1)
        assert len(slides) == 2
        assert slides[0].title == "スライド1"
        assert slides[1].title == "スライド2"

    # --- 背景画像属性のテスト ---

    def test_split_background_image_attribute(self) -> None:
        """見出しの {background-image="..."} 属性を正しく解析する。"""
        text = _FRONTMATTER + '## スライド {background-image="bg.png"}\n\n本文\n'
        slides = self.splitter.split(text, slide_level=2)
        assert len(slides) == 1
        assert slides[0].background_image == "bg.png"
        assert slides[0].title == "スライド"

    def test_split_no_background_image(self) -> None:
        """背景画像属性がない場合は None を返す。"""
        text = _FRONTMATTER + "## スライド\n\n本文\n"
        slides = self.splitter.split(text, slide_level=2)
        assert slides[0].background_image is None

    def test_split_background_image_removed_from_title(self) -> None:
        """タイトルから属性ブロックが除去される。"""
        text = _FRONTMATTER + '## タイトル {background-image="image.jpg"}\n\n本文\n'
        slides = self.splitter.split(text, slide_level=2)
        assert "{background-image" not in slides[0].title
        assert slides[0].title == "タイトル"

    # --- 空テキストのテスト ---

    def test_split_empty_body_returns_empty_list(self) -> None:
        """フロントマターのみのテキストはスライドを返さない。"""
        text = _FRONTMATTER
        slides = self.splitter.split(text, slide_level=2)
        assert slides == []

    def test_split_body_text_is_stripped(self) -> None:
        """スライドの本文テキストは前後の空白が除去される。"""
        text = _FRONTMATTER + "## スライド\n\n   内容   \n"
        slides = self.splitter.split(text, slide_level=2)
        assert slides[0].body_text == "内容"
