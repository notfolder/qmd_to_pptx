"""
DOMトラバーサーモジュール。

Markdownパーサーが生成したElementTree形式のDOMツリーを走査し、
各ノードの種別を判定してDOMNodeInfoのリストを返す。
"""

from __future__ import annotations

import xml.etree.ElementTree as ET

from .models import DOMNodeInfo, DOMNodeType


class DOMTraverser:
    """
    DOMトラバーサークラス。

    DOMツリーのルートから深さ優先でノードを走査し、
    ノード種別とノード内容を DOMNodeInfo のリストとして返す。
    """

    def traverse(self, root: ET.Element) -> list[DOMNodeInfo]:
        """
        DOMツリーを深さ優先で走査してDOMNodeInfoのリストを返す。

        Parameters
        ----------
        root : ET.Element
            DOMツリーのルート要素。

        Returns
        -------
        list[DOMNodeInfo]
            ノード種別とノード内容を格納したDOMNodeInfoのリスト。
        """
        result: list[DOMNodeInfo] = []
        # ルートの直下の子ノードを走査する（ルート自体は<div>ラッパーのため除外）
        for child in root:
            self._traverse_node(child, result)
        return result

    def _traverse_node(
        self, element: ET.Element, result: list[DOMNodeInfo]
    ) -> None:
        """
        個々のノードを判定し、DOMNodeInfoとしてresultに追加する。

        子ノードを再帰的に走査する必要がある場合（columnsなど）は
        子ノードも処理する。

        Parameters
        ----------
        element : ET.Element
            処理対象のノード要素。
        result : list[DOMNodeInfo]
            結果を格納するリスト。
        """
        tag = element.tag
        css_class = element.get("class", "")

        # タグ名・クラス属性に基づいてノード種別を判定する
        node_type = self._classify_node(tag, css_class)

        if node_type is not None:
            result.append(DOMNodeInfo(node_type=node_type, element=element))
        else:
            # 未知のノードは子ノードを再帰的に走査する
            for child in element:
                self._traverse_node(child, result)

    def _classify_node(self, tag: str, css_class: str) -> DOMNodeType | None:
        """
        タグ名とクラス属性からDOMNodeTypeを判定して返す。

        判定できない場合はNoneを返す。

        Parameters
        ----------
        tag : str
            HTMLタグ名。
        css_class : str
            クラス属性値（複数クラスは空白区切り）。

        Returns
        -------
        DOMNodeType | None
            判定したノード種別。不明な場合はNone。
        """
        classes = set(css_class.split())

        # 見出し系
        if tag == "h1":
            return DOMNodeType.H1
        if tag == "h2":
            return DOMNodeType.H2

        # 段落
        if tag == "p":
            return DOMNodeType.PARAGRAPH

        # リスト
        if tag == "ul":
            return DOMNodeType.UL
        if tag == "ol":
            return DOMNodeType.OL

        # テーブル
        if tag == "table":
            return DOMNodeType.TABLE

        # コードブロック系（Mermaid / 通常）
        if tag == "code":
            if "language-mermaid" in classes:
                return DOMNodeType.MERMAID
            return DOMNodeType.CODE

        # arithmatex（数式）
        if tag == "span" and "arithmatex" in classes:
            return DOMNodeType.FORMULA_INLINE
        if tag == "div" and "arithmatex" in classes:
            return DOMNodeType.FORMULA_BLOCK

        # スピーカーノート
        if tag == "div" and "notes" in classes:
            return DOMNodeType.NOTES

        # 2カラムコンテナ
        if tag == "div" and "columns" in classes:
            return DOMNodeType.COLUMNS

        # インクリメンタルリスト
        if tag == "div" and "incremental" in classes:
            return DOMNodeType.INCREMENTAL

        # 非インクリメンタルリスト
        if tag == "div" and "nonincremental" in classes:
            return DOMNodeType.NON_INCREMENTAL

        return None
