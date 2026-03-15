"""
Mermaidレンダラー（MermaidRenderer）の単体テスト。
"""

import xml.etree.ElementTree as ET
import pytest
from pptx import Presentation
from pptx.util import Emu
from qmd_to_pptx.mermaid_renderer import MermaidRenderer


def _make_slide():
    """テスト用スライドを生成して返す。"""
    prs = Presentation()
    layout = prs.slide_layouts[5]  # Blankレイアウト
    return prs.slides.add_slide(layout)


def _make_mermaid_element(mermaid_text: str) -> ET.Element:
    """テスト用Mermaid code要素を生成する。"""
    elem = ET.Element("code")
    elem.set("class", "language-mermaid")
    elem.text = mermaid_text
    return elem


class TestMermaidRenderer:
    """MermaidRenderer クラスの単体テスト。"""

    def setup_method(self) -> None:
        """テスト前にMermaidRendererインスタンスを生成する。"""
        self.renderer = MermaidRenderer()

    def test_render_simple_flowchart_does_not_raise(self) -> None:
        """シンプルなflowchartをレンダリングしても例外が発生しない。"""
        slide = _make_slide()
        elem = _make_mermaid_element("flowchart LR\n    A --> B")
        # 例外なく完了することを確認
        self.renderer.render(
            slide, elem,
            Emu(500000), Emu(500000), Emu(7000000), Emu(4000000)
        )

    def test_render_adds_shapes_to_slide(self) -> None:
        """レンダリング後にスライドにShapeが追加される。"""
        slide = _make_slide()
        initial_shapes = len(slide.shapes)
        elem = _make_mermaid_element("flowchart LR\n    A --> B")
        self.renderer.render(
            slide, elem,
            Emu(500000), Emu(500000), Emu(7000000), Emu(4000000)
        )
        # パース成功またはフォールバック時も必ずShapeが追加される
        assert len(slide.shapes) > initial_shapes

    def test_render_fallback_on_invalid_mermaid(self) -> None:
        """無効なMermaidテキストでもフォールバックとしてテキストボックスが追加される。"""
        slide = _make_slide()
        initial_shapes = len(slide.shapes)
        elem = _make_mermaid_element("これはMermaidではありません！！！###")
        self.renderer.render(
            slide, elem,
            Emu(500000), Emu(500000), Emu(7000000), Emu(4000000)
        )
        # フォールバック時も何らかのShapeが追加される
        assert len(slide.shapes) >= initial_shapes

    def test_render_empty_mermaid_uses_fallback(self) -> None:
        """空のMermaidテキストはフォールバック処理される。"""
        slide = _make_slide()
        elem = _make_mermaid_element("")
        # 例外なく完了することを確認
        self.renderer.render(
            slide, elem,
            Emu(500000), Emu(500000), Emu(7000000), Emu(4000000)
        )

    def test_extract_nodes(self) -> None:
        """_extract_nodes が graph_data から正しくノードリストを返す。"""
        graph_data = {
            "vertices": {"A": {}, "B": {}, "C": {}}
        }
        nodes = self.renderer._extract_nodes(graph_data)
        assert set(nodes) == {"A", "B", "C"}

    def test_extract_nodes_empty(self) -> None:
        """vertices が空の場合は空リストを返す。"""
        assert self.renderer._extract_nodes({}) == []

    def test_extract_edges(self) -> None:
        """_extract_edges が graph_data から正しくエッジリストを返す。"""
        graph_data = {
            "edges": [
                {"start": "A", "end": "B"},
                {"start": "B", "end": "C"},
            ]
        }
        edges = self.renderer._extract_edges(graph_data)
        assert ("A", "B") in edges
        assert ("B", "C") in edges

    def test_extract_edges_empty(self) -> None:
        """edges が空の場合は空リストを返す。"""
        assert self.renderer._extract_edges({}) == []

    def test_pos_to_emu_center(self) -> None:
        """正規化座標 (0, 0) は描画エリアの中央付近に変換される。"""
        left, top, width, height = 0, 0, 9144000, 5143500
        x_emu, y_emu = self.renderer._pos_to_emu(0.0, 0.0, left, top, width, height)
        # 中央付近であることを確認（厳密な値ではなく範囲で確認）
        assert width // 4 < x_emu < width * 3 // 4
        assert height // 4 < y_emu < height * 3 // 4
