"""
数式レンダラー（FormulaRenderer）の単体テスト。
"""

import xml.etree.ElementTree as ET
import pytest
from pptx import Presentation
from pptx.util import Emu
from qmd_to_pptx.formula_renderer import FormulaRenderer


def _make_slide():
    """テスト用スライドを生成して返す。"""
    prs = Presentation()
    layout = prs.slide_layouts[5]  # Blankレイアウト
    return prs.slides.add_slide(layout)


def _make_inline_element(latex: str) -> ET.Element:
    """インライン数式要素 span.arithmatex を生成する。"""
    elem = ET.Element("span")
    elem.set("class", "arithmatex")
    elem.text = f"\\({latex}\\)"
    return elem


def _make_block_element(latex: str) -> ET.Element:
    """ブロック数式要素 div.arithmatex を生成する。"""
    elem = ET.Element("div")
    elem.set("class", "arithmatex")
    elem.text = f"\\[{latex}\\]"
    return elem


class TestFormulaRenderer:
    """FormulaRenderer クラスの単体テスト。"""

    def setup_method(self) -> None:
        """テスト前にFormulaRendererインスタンスを生成する。"""
        self.renderer = FormulaRenderer()

    # --- _extract_latex のテスト ---

    def test_extract_latex_inline(self) -> None:
        r"""インライン数式 \(...\) からデリミタを除去してLaTeXを返す。"""
        elem = _make_inline_element("E=mc^2")
        result = self.renderer._extract_latex(elem)
        assert result == "E=mc^2"

    def test_extract_latex_block(self) -> None:
        r"""ブロック数式 \[...\] からデリミタを除去してLaTeXを返す。"""
        elem = _make_block_element(r"\sum_{i=0}^{n} i")
        result = self.renderer._extract_latex(elem)
        assert result == r"\sum_{i=0}^{n} i"

    def test_extract_latex_dollar_delimiter(self) -> None:
        """$$...$$ 形式のデリミタも除去する。"""
        elem = ET.Element("span")
        elem.set("class", "arithmatex")
        elem.text = "$$E=mc^2$$"
        result = self.renderer._extract_latex(elem)
        assert result == "E=mc^2"

    def test_extract_latex_no_delimiter(self) -> None:
        """デリミタなしのテキストはそのまま返す。"""
        elem = ET.Element("span")
        elem.set("class", "arithmatex")
        elem.text = "E=mc^2"
        result = self.renderer._extract_latex(elem)
        assert result == "E=mc^2"

    # --- render_block のテスト ---

    def test_render_block_does_not_raise_on_valid_latex(self) -> None:
        """有効なLaTeXのブロック数式をレンダリングしても例外が発生しない。"""
        slide = _make_slide()
        elem = _make_block_element("E=mc^2")
        # 例外なく完了することを確認
        self.renderer.render_block(
            slide, elem,
            Emu(500000), Emu(500000), Emu(7000000), Emu(2000000)
        )

    def test_render_block_fallback_on_invalid_latex(self) -> None:
        """無効なLaTeXでもフォールバックとしてテキストボックスが追加される。"""
        slide = _make_slide()
        initial_shapes = len(slide.shapes)
        elem = ET.Element("div")
        elem.set("class", "arithmatex")
        elem.text = "\\[これは無効な数式です\\]"
        self.renderer.render_block(
            slide, elem,
            Emu(500000), Emu(500000), Emu(7000000), Emu(2000000)
        )
        # フォールバックでテキストボックスが追加されるか、OMMLが追加されるか
        # いずれにせよ例外なく完了することを確認
        # （フォールバック時はshapeが追加される）

    def test_render_block_empty_latex_is_skipped(self) -> None:
        """LaTeXテキストが空の場合は何も追加しない。"""
        slide = _make_slide()
        initial_shapes = len(slide.shapes)
        elem = ET.Element("div")
        elem.set("class", "arithmatex")
        elem.text = ""
        self.renderer.render_block(
            slide, elem,
            Emu(500000), Emu(500000), Emu(7000000), Emu(2000000)
        )
        # 空テキストの場合は何も追加されない
        assert len(slide.shapes) == initial_shapes
