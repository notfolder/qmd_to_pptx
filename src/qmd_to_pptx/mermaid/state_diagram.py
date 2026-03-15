"""
ステートダイアグラムレンダラーモジュール。

stateDiagram-v2 を nodes リストと edges リストから描画する。
"""

from __future__ import annotations

import networkx as nx
from pptx.slide import Slide

from .base import BaseDiagramRenderer


class StateDiagramRenderer(BaseDiagramRenderer):
    """
    stateDiagram-v2 を描画するレンダラー。

    graph_data の "nodes" リストと "edges" リストを読み取り、
    spring_layout で配置した矩形ノードとコネクターを描画する。
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
        stateDiagram-v2 をスライドに描画する。

        Parameters
        ----------
        slide : Slide
            python-pptxのSlideオブジェクト。
        graph_data : dict
            graph_data辞書（"nodes"と"edges"キーを含む）。
        left, top, width, height : int
            描画エリアのEMU座標。
        """
        raw_nodes = graph_data.get("nodes", [])
        raw_edges = graph_data.get("edges", [])
        if not raw_nodes:
            return

        # ノードIDリストとラベルマップを構築する
        nodes = [n["id"] for n in raw_nodes if "id" in n]
        label_map = {n["id"]: n.get("label", n["id"]) for n in raw_nodes}
        edges = [
            (e["start"], e["end"])
            for e in raw_edges
            if "start" in e and "end" in e
        ]

        G = nx.DiGraph()
        G.add_nodes_from(nodes)
        for src, dst in edges:
            G.add_edge(src, dst)
        pos: dict[str, tuple[float, float]] = nx.spring_layout(G, seed=42)

        node_shapes = self._draw_nodes(
            slide, nodes, pos, left, top, width, height, label_map=label_map
        )
        self._draw_edges(slide, edges, pos, node_shapes, left, top, width, height)
