"""
Mermaidレンダラーモジュール。

Mermaid記法のテキストをパースしてグラフ構造を取得し、
python-pptxのShapeとして現在のスライドに図を描画する。
"""

from __future__ import annotations

import xml.etree.ElementTree as ET

import networkx as nx
from mermaid_parser import MermaidParser
from pptx.slide import Slide
from pptx.util import Emu, Pt


# ノードのデフォルトサイズ（EMU）
_NODE_WIDTH_EMU: int = 1200000
_NODE_HEIGHT_EMU: int = 500000


class MermaidRenderer:
    """
    Mermaidレンダラークラス。

    Mermaidテキストをmermaid-parser-pyで解析してNetworkXグラフに変換し、
    spring_layoutで座標計算後、python-pptxのShape/Connectorとして描画する。
    """

    def render(
        self,
        slide: Slide,
        element: ET.Element,
        left: int,
        top: int,
        width: int,
        height: int,
    ) -> None:
        """
        elementからMermaidテキストを取り出し、スライドにグラフを描画する。

        Parameters
        ----------
        slide : Slide
            python-pptxのSlideオブジェクト。
        element : ET.Element
            Mermaidコード要素（code class="language-mermaid"）。
        left : int
            描画エリアの左端座標（EMU）。
        top : int
            描画エリアの上端座標（EMU）。
        width : int
            描画エリアの幅（EMU）。
        height : int
            描画エリアの高さ（EMU）。
        """
        mermaid_text = "".join(element.itertext()).strip()

        try:
            # mermaid-parser-pyでノードとエッジを取得する
            mp = MermaidParser()
            result = mp.parse(mermaid_text)
            graph_data = result.get("graph_data", {})
            nodes = self._extract_nodes(graph_data)
            edges = self._extract_edges(graph_data)
        except Exception:
            # パース失敗時はテキストボックスにそのまま表示する
            self._render_fallback(slide, mermaid_text, left, top, width, height)
            return

        if not nodes:
            self._render_fallback(slide, mermaid_text, left, top, width, height)
            return

        # NetworkXグラフを構築してspring_layoutで座標計算する
        G = nx.DiGraph()
        G.add_nodes_from(nodes)
        for src, dst in edges:
            G.add_edge(src, dst)

        # spring_layoutで正規化された座標（-1.0〜1.0）を計算する
        pos: dict[str, tuple[float, float]] = nx.spring_layout(G, seed=42)

        # 座標をEMUに変換してノードとエッジを描画する
        node_shapes = self._draw_nodes(slide, nodes, pos, left, top, width, height)
        self._draw_edges(slide, edges, pos, left, top, width, height)

    def _extract_nodes(self, graph_data: dict) -> list[str]:
        """
        mermaid-parser-pyの解析結果からノードIDのリストを取得する。

        Parameters
        ----------
        graph_data : dict
            graph_data辞書（"vertices"キーにノード情報を含む）。

        Returns
        -------
        list[str]
            ノードIDのリスト。
        """
        vertices = graph_data.get("vertices", {})
        if isinstance(vertices, dict):
            return list(vertices.keys())
        return []

    def _extract_edges(self, graph_data: dict) -> list[tuple[str, str]]:
        """
        mermaid-parser-pyの解析結果からエッジのリストを取得する。

        Parameters
        ----------
        graph_data : dict
            graph_data辞書（"edges"キーにエッジ情報を含む）。

        Returns
        -------
        list[tuple[str, str]]
            (始点ノードID, 終点ノードID) のタプルリスト。
        """
        raw_edges = graph_data.get("edges", [])
        edges: list[tuple[str, str]] = []
        for edge in raw_edges:
            if isinstance(edge, dict):
                src = edge.get("start")
                dst = edge.get("end")
                if src is not None and dst is not None:
                    edges.append((str(src), str(dst)))
        return edges

    def _pos_to_emu(
        self,
        x_norm: float,
        y_norm: float,
        left: int,
        top: int,
        width: int,
        height: int,
    ) -> tuple[int, int]:
        """
        spring_layoutで計算した正規化座標（-1.0〜1.0）をEMU座標に変換する。

        Parameters
        ----------
        x_norm : float
            正規化X座標（-1.0〜1.0）。
        y_norm : float
            正規化Y座標（-1.0〜1.0）。
        left : int
            描画エリアの左端座標（EMU）。
        top : int
            描画エリアの上端座標（EMU）。
        width : int
            描画エリアの幅（EMU）。
        height : int
            描画エリアの高さ（EMU）。

        Returns
        -------
        tuple[int, int]
            (x_emu, y_emu) のタプル。
        """
        margin_x = _NODE_WIDTH_EMU // 2
        margin_y = _NODE_HEIGHT_EMU // 2
        usable_w = width - _NODE_WIDTH_EMU
        usable_h = height - _NODE_HEIGHT_EMU

        # -1.0〜1.0 を 0〜1 に正規化する
        x_ratio = (x_norm + 1.0) / 2.0
        y_ratio = (y_norm + 1.0) / 2.0

        x_emu = left + margin_x + int(x_ratio * usable_w)
        y_emu = top + margin_y + int(y_ratio * usable_h)
        return x_emu, y_emu

    def _draw_nodes(
        self,
        slide: Slide,
        nodes: list[str],
        pos: dict[str, tuple[float, float]],
        left: int,
        top: int,
        width: int,
        height: int,
    ) -> dict[str, object]:
        """
        ノードを矩形Shapeとしてスライドに配置する。

        Parameters
        ----------
        slide : Slide
            python-pptxのSlideオブジェクト。
        nodes : list[str]
            ノードIDのリスト。
        pos : dict[str, tuple[float, float]]
            ノードIDをキー、正規化座標を値とする辞書。
        left : int
            描画エリアの左端座標（EMU）。
        top : int
            描画エリアの上端座標（EMU）。
        width : int
            描画エリアの幅（EMU）。
        height : int
            描画エリアの高さ（EMU）。

        Returns
        -------
        dict[str, object]
            ノードIDをキー、Shapeオブジェクトを値とする辞書。
        """
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        from pptx.util import Emu

        node_shapes: dict[str, object] = {}
        for node_id in nodes:
            if node_id not in pos:
                continue
            x_norm, y_norm = pos[node_id]
            cx, cy = self._pos_to_emu(x_norm, y_norm, left, top, width, height)
            shape_left = cx - _NODE_WIDTH_EMU // 2
            shape_top = cy - _NODE_HEIGHT_EMU // 2

            # add_shape で矩形を追加する
            from pptx.enum.shapes import MSO_SHAPE_TYPE
            from pptx.util import Emu
            shape = slide.shapes.add_shape(
                1,  # MSO_SHAPE_TYPE.RECTANGLE
                Emu(shape_left),
                Emu(shape_top),
                Emu(_NODE_WIDTH_EMU),
                Emu(_NODE_HEIGHT_EMU),
            )
            # ノードIDをテキストとして設定する
            shape.text = node_id
            tf = shape.text_frame
            for para in tf.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(12)

            node_shapes[node_id] = shape

        return node_shapes

    def _draw_edges(
        self,
        slide: Slide,
        edges: list[tuple[str, str]],
        pos: dict[str, tuple[float, float]],
        left: int,
        top: int,
        width: int,
        height: int,
    ) -> None:
        """
        エッジをConnectorShapeとしてスライドに描画する。

        Parameters
        ----------
        slide : Slide
            python-pptxのSlideオブジェクト。
        edges : list[tuple[str, str]]
            (始点ノードID, 終点ノードID) のタプルリスト。
        pos : dict[str, tuple[float, float]]
            ノードIDをキー、正規化座標を値とする辞書。
        left : int
            描画エリアの左端座標（EMU）。
        top : int
            描画エリアの上端座標（EMU）。
        width : int
            描画エリアの幅（EMU）。
        height : int
            描画エリアの高さ（EMU）。
        """
        from pptx.util import Emu

        for src, dst in edges:
            if src not in pos or dst not in pos:
                continue
            sx_norm, sy_norm = pos[src]
            dx_norm, dy_norm = pos[dst]
            sx, sy = self._pos_to_emu(sx_norm, sy_norm, left, top, width, height)
            dx, dy = self._pos_to_emu(dx_norm, dy_norm, left, top, width, height)

            # コネクターを追加する（直線コネクター: connector_type=1）
            connector = slide.shapes.add_connector(
                1,  # MSO_CONNECTOR_TYPE.STRAIGHT
                Emu(sx),
                Emu(sy),
                Emu(dx),
                Emu(dy),
            )

    def _render_fallback(
        self,
        slide: Slide,
        mermaid_text: str,
        left: int,
        top: int,
        width: int,
        height: int,
    ) -> None:
        """
        Mermaidパース失敗時のフォールバック処理。
        テキストボックスにMermaidテキストをそのまま表示する。

        Parameters
        ----------
        slide : Slide
            python-pptxのSlideオブジェクト。
        mermaid_text : str
            Mermaidテキスト。
        left : int
            左端座標（EMU）。
        top : int
            上端座標（EMU）。
        width : int
            幅（EMU）。
        height : int
            高さ（EMU）。
        """
        from pptx.util import Emu
        shape = slide.shapes.add_textbox(
            Emu(left), Emu(top), Emu(width), Emu(height)
        )
        tf = shape.text_frame
        tf.word_wrap = True
        para = tf.paragraphs[0]
        run = para.add_run()
        run.text = mermaid_text
        run.font.name = "Courier New"
        run.font.size = Pt(10)
