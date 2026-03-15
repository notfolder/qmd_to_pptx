"""
Markdownパーサー（MarkdownParser）の単体テスト。
"""

import xml.etree.ElementTree as ET
import pytest
from qmd_to_pptx.markdown_parser import MarkdownParser


class TestMarkdownParser:
    """MarkdownParser クラスの単体テスト。"""

    def setup_method(self) -> None:
        """テスト前にMarkdownParserインスタンスを生成する。"""
        self.parser = MarkdownParser()

    def _find_tags(self, root: ET.Element, tag: str) -> list[ET.Element]:
        """ルート要素から指定タグを再帰的に検索する。"""
        return list(root.iter(tag))

    # --- 基本タグ変換 ---

    def test_parse_paragraph(self) -> None:
        """段落テキストを <p> 要素として変換する。"""
        root = self.parser.parse("これはテスト段落です。")
        paragraphs = self._find_tags(root, "p")
        assert len(paragraphs) >= 1
        assert "テスト段落" in "".join(paragraphs[0].itertext())

    def test_parse_heading2(self) -> None:
        """## 見出しを <h2> 要素として変換する。"""
        root = self.parser.parse("## 見出し2")
        headings = self._find_tags(root, "h2")
        assert len(headings) == 1
        assert "見出し2" in "".join(headings[0].itertext())

    def test_parse_unordered_list(self) -> None:
        """箇条書きリストを <ul><li> 要素として変換する。"""
        root = self.parser.parse("- アイテム1\n- アイテム2\n- アイテム3")
        ul_elements = self._find_tags(root, "ul")
        assert len(ul_elements) >= 1
        li_elements = self._find_tags(root, "li")
        assert len(li_elements) == 3

    def test_parse_table(self) -> None:
        """Markdown表を <table> 要素として変換する。"""
        md = "| 列1 | 列2 |\n|-----|-----|\n| A   | B   |"
        root = self.parser.parse(md)
        tables = self._find_tags(root, "table")
        assert len(tables) == 1

    def test_parse_mermaid_fenced_code(self) -> None:
        """```mermaid ブロックを class='language-mermaid' の <code> 要素に変換する。"""
        md = "```mermaid\nflowchart LR\n    A --> B\n```"
        root = self.parser.parse(md)
        # language-mermaid クラスを持つ code 要素を探す
        code_elements = [
            elem for elem in root.iter("code")
            if "language-mermaid" in elem.get("class", "")
        ]
        assert len(code_elements) == 1
        assert "A --> B" in "".join(code_elements[0].itertext())

    def test_parse_returns_element(self) -> None:
        """parse() は ET.Element を返す。"""
        root = self.parser.parse("テスト")
        assert isinstance(root, ET.Element)

    def test_parse_empty_string(self) -> None:
        """空文字列は空の div 要素を返す。"""
        root = self.parser.parse("")
        assert isinstance(root, ET.Element)

    # --- リストネスト ---

    def test_parse_nested_ul_2spaces(self) -> None:
        """順序なしリストでスペース2つインデントがネスト <ul>構造として認識される。"""
        md = "- 親アイテム\n  - 子アイテム"
        root = self.parser.parse(md)
        # 外側 ul の直下 li 内にネストされた ul が存在する
        outer_ul = root.find(".//ul")
        assert outer_ul is not None, "外側 <ul> が見つからない"
        parent_li = outer_ul.find("li")
        assert parent_li is not None, "<li> が見つからない"
        nested_ul = parent_li.find(".//ul")
        assert nested_ul is not None, "スペース2つインデントのネスト <ul> が生成されていない"

    def test_parse_nested_ol_2spaces(self) -> None:
        """順序リストでスペース2つインデントがネスト <ol>構造として認識される。"""
        md = "1. 親アイテム\n  1. 子アイテム"
        root = self.parser.parse(md)
        outer_ol = root.find(".//ol")
        assert outer_ol is not None, "外側 <ol> が見つからない"
        parent_li = outer_ol.find("li")
        assert parent_li is not None, "<li> が見つからない"
        nested_ol = parent_li.find(".//ol")
        assert nested_ol is not None, "スペース2つインデントのネスト <ol> が生成されていない"
