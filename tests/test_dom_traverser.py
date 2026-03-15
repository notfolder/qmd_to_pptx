"""
DOMトラバーサー（DOMTraverser）の単体テスト。
"""

import xml.etree.ElementTree as ET
import pytest
from qmd_to_pptx.dom_traverser import DOMTraverser
from qmd_to_pptx.models import DOMNodeType


def _make_root(*children_xml: str) -> ET.Element:
    """テスト用のルート <div> 要素を生成する。"""
    inner = "".join(children_xml)
    return ET.fromstring(f"<div>{inner}</div>")


class TestDOMTraverser:
    """DOMTraverser クラスの単体テスト。"""

    def setup_method(self) -> None:
        """テスト前にDOMTraverserインスタンスを生成する。"""
        self.traverser = DOMTraverser()

    def _types(self, xml_str: str) -> list[DOMNodeType]:
        """XMLからトラバース結果のノード種別リストを返す。"""
        root = _make_root(xml_str)
        nodes = self.traverser.traverse(root)
        return [n.node_type for n in nodes]

    # --- タグ種別判定のテスト ---

    def test_traverse_h1(self) -> None:
        """<h1> タグを H1 として識別する。"""
        assert DOMNodeType.H1 in self._types("<h1>見出し</h1>")

    def test_traverse_h2(self) -> None:
        """<h2> タグを H2 として識別する。"""
        assert DOMNodeType.H2 in self._types("<h2>見出し</h2>")

    def test_traverse_paragraph(self) -> None:
        """<p> タグを PARAGRAPH として識別する。"""
        assert DOMNodeType.PARAGRAPH in self._types("<p>段落テキスト</p>")

    def test_traverse_ul(self) -> None:
        """<ul> タグを UL として識別する。"""
        assert DOMNodeType.UL in self._types("<ul><li>アイテム</li></ul>")

    def test_traverse_ol(self) -> None:
        """<ol> タグを OL として識別する。"""
        assert DOMNodeType.OL in self._types("<ol><li>アイテム</li></ol>")

    def test_traverse_table(self) -> None:
        """<table> タグを TABLE として識別する。"""
        assert DOMNodeType.TABLE in self._types("<table><tr><td>セル</td></tr></table>")

    def test_traverse_code_non_mermaid(self) -> None:
        """クラスが language-mermaid でない <code> タグを CODE として識別する。"""
        assert DOMNodeType.CODE in self._types('<code class="language-python">print()</code>')

    def test_traverse_code_mermaid(self) -> None:
        """class='language-mermaid' の <code> タグを MERMAID として識別する。"""
        assert DOMNodeType.MERMAID in self._types('<code class="language-mermaid">A --> B</code>')

    def test_traverse_formula_inline(self) -> None:
        """class='arithmatex' の <span> タグを FORMULA_INLINE として識別する。"""
        assert DOMNodeType.FORMULA_INLINE in self._types('<span class="arithmatex">\\(E=mc^2\\)</span>')

    def test_traverse_formula_block(self) -> None:
        """class='arithmatex' の <div> タグを FORMULA_BLOCK として識別する。"""
        assert DOMNodeType.FORMULA_BLOCK in self._types('<div class="arithmatex">\\[E=mc^2\\]</div>')

    def test_traverse_notes(self) -> None:
        """class='notes' の <div> タグを NOTES として識別する。"""
        assert DOMNodeType.NOTES in self._types('<div class="notes">スピーカーノート</div>')

    def test_traverse_columns(self) -> None:
        """class='columns' の <div> タグを COLUMNS として識別する。"""
        assert DOMNodeType.COLUMNS in self._types('<div class="columns"><div class="column">左</div></div>')

    def test_traverse_incremental(self) -> None:
        """class='incremental' の <div> タグを INCREMENTAL として識別する。"""
        assert DOMNodeType.INCREMENTAL in self._types('<div class="incremental"><ul><li>アイテム</li></ul></div>')

    def test_traverse_non_incremental(self) -> None:
        """class='nonincremental' の <div> タグを NON_INCREMENTAL として識別する。"""
        assert DOMNodeType.NON_INCREMENTAL in self._types('<div class="nonincremental"><ul><li>アイテム</li></ul></div>')

    # --- traverse() の返り値のテスト ---

    def test_traverse_returns_list(self) -> None:
        """traverse() は DOMNodeInfo のリストを返す。"""
        root = _make_root("<p>テスト</p>")
        result = self.traverser.traverse(root)
        assert isinstance(result, list)

    def test_traverse_empty_root_returns_empty_list(self) -> None:
        """空のルート要素は空リストを返す。"""
        root = ET.Element("div")
        result = self.traverser.traverse(root)
        assert result == []

    def test_traverse_preserves_element(self) -> None:
        """DOMNodeInfo に元の Element 参照が保持される。"""
        root = _make_root("<p>テスト段落</p>")
        nodes = self.traverser.traverse(root)
        assert len(nodes) == 1
        assert "テスト段落" in "".join(nodes[0].element.itertext())

    def test_traverse_multiple_nodes_ordered(self) -> None:
        """複数ノードが出現順に返される。"""
        root = _make_root("<h2>見出し</h2><p>段落</p>")
        nodes = self.traverser.traverse(root)
        types = [n.node_type for n in nodes]
        assert types == [DOMNodeType.H2, DOMNodeType.PARAGRAPH]
