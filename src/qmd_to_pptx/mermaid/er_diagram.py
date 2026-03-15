"""
ER図レンダラーモジュール。

erDiagram を entities 辞書と relationships リストから描画する。
"""

from __future__ import annotations

import networkx as nx
from pptx.slide import Slide

from .base import BaseDiagramRenderer


class ErDiagramRenderer(BaseDiagramRenderer):
    """
    erDiagram を描画するレンダラー。

    graph_data の "entities" 辞書と "relationships" リストを読み取り、
    spring_layout で配置した矩形ノードとコネクターを描画する。
    entities のキーはラベル（エンティティ名）、値の id フィールドが内部ID。
    relationships は内部IDで参照するため逆引きマップを使用する。
    """

    def render(
        self,
        slide: Slide,
        graph_data: dict,
        left: int,
        top: int,
        width: int,
        height: int,
    ) -> None:
        """
        erDiagram をスライドに描画する。

        Parameters
        ----------
        slide : Slide
            python-pptxのSlideオブジェクト。
        graph_data : dict
            graph_data辞書（"entities"と"relationships"キーを含む）。
        left, top, width, height : int
            描画エリアのEMU座標。
        """
        entities = graph_data.get("entities", {})
        relationships = graph_data.get("relationships", [])
        if not entities:
            return

        # 内部IDからエンティティ名への逆引きマップを構築する
        id_to_label: dict[str, str] = {
            ent.get("id", ""): label
            for label, ent in entities.items()
            if ent.get("id")
        }

        nodes = list(entities.keys())
        edges: list[tuple[str, str]] = [
            (id_to_label[rel["entityA"]], id_to_label[rel["entityB"]])
            for rel in relationships
            if rel.get("entityA") in id_to_label and rel.get("entityB") in id_to_label
        ]

        G = nx.DiGraph()
        G.add_nodes_from(nodes)
        for src, dst in edges:
            G.add_edge(src, dst)
        # kamada_kawai_layoutでノード配置の重なりを軽減する（失敗時はspring_layoutにフォールバック）
        try:
            pos: dict[str, tuple[float, float]] = nx.kamada_kawai_layout(G)
        except Exception:
            pos = nx.spring_layout(G, seed=42, k=2.0)

        node_shapes = self._draw_nodes(slide, nodes, pos, left, top, width, height)
        self._draw_edges(slide, edges, pos, node_shapes, left, top, width, height)
