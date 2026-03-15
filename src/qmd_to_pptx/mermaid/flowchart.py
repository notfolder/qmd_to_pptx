"""
フローチャートレンダラーモジュール。

flowchart / graph 系のMermaid記法を解析し、14種類のノード形状・
7種類のエッジ矢印・4種類の線種・エッジラベルに対応してスライドに描画する。
描画方向（TD/TB/BT/LR/RL）を考慮したDAG階層レイアウトをサポートする。
"""

from __future__ import annotations

import re

import networkx as nx
from lxml import etree as lxml_etree
from pptx.oxml.ns import qn
from pptx.slide import Slide
from pptx.util import Emu, Pt

from .base import BaseDiagramRenderer, NODE_WIDTH_EMU, NODE_HEIGHT_EMU


# ---------------------- ノード形状マップ ----------------------
# Mermaid頂点タイプ → MSO_AUTO_SHAPE_TYPE整数値
# python-pptx add_shape() に渡す値（MSO_AUTO_SHAPE_TYPE enum）
_SHAPE_MAP: dict[str, int] = {
    "square":        1,   # RECTANGLE（デフォルト矩形 []）
    "round":         5,   # ROUNDED_RECTANGLE（丸角矩形 ()）
    "stadium":       5,   # ROUNDED_RECTANGLE（スタジアム型 ([])）
    "subroutine":    1,   # RECTANGLE（サブルーチン [[]] ※二重枠は省略）
    "cylinder":      13,  # CAN（シリンダー [()]）
    "circle":        9,   # OVAL（円 (())）
    "odd":           51,  # PENTAGON（非対称 >]）
    "diamond":       4,   # DIAMOND（ひし形 {}）
    "hexagon":       10,  # HEXAGON（六角形 {{}}）
    "lean_right":    2,   # PARALLELOGRAM（右傾斜 [/]）
    "lean_left":     2,   # PARALLELOGRAM + flipH（左傾斜 [\]）
    "trapezoid":     3,   # TRAPEZOID（台形 [/\]）
    "inv_trapezoid": 3,   # TRAPEZOID + flipV（逆台形 [\/]）
    "doublecircle":  9,   # OVAL（二重円 ((())) ※二重枠は省略）
}

# ---------------------- エッジ矢印マップ ----------------------
# mermaid-parser-py の edge["type"] → headEnd/tailEnd 設定辞書
# headEnd: コネクター終点（dst側）の矢印種別
# tailEnd: コネクター始点（src側）の矢印種別（両方向エッジのみ）
_EDGE_ARROW_MAP: dict[str, dict[str, str]] = {
    "arrow_open":          {},                                      # 矢印なし（開放線）
    "arrow_point":         {"headEnd": "arrow"},                    # →（通常矢印）
    "double_arrow_point":  {"headEnd": "arrow", "tailEnd": "arrow"}, # ←→（両方向）
    "arrow_circle":        {"headEnd": "oval"},                     # 丸端矢印
    "double_arrow_circle": {"headEnd": "oval", "tailEnd": "oval"},  # 両方向丸端
    "arrow_cross":         {"headEnd": "arrow"},                    # ×端（arrow で代替）
    "double_arrow_cross":  {"headEnd": "arrow", "tailEnd": "arrow"},# 両方向×端
}


class FlowchartRenderer(BaseDiagramRenderer):
    """
    flowchart / graph 系のMermaid記法を描画するレンダラー。

    mermaid-parser-py が返す graph_data（vertices/edges）を受け取り、
    ノード形状・エッジスタイル・エッジラベルに対応してスライドに描画する。
    """

    def render(
        self,
        slide: Slide,
        graph_data: dict,
        mermaid_text: str,
        left: int,
        top: int,
        width: int,
        height: int,
    ) -> None:
        """
        フローチャートをスライドに描画する。

        Parameters
        ----------
        slide : Slide
            python-pptxのSlideオブジェクト。
        graph_data : dict
            mermaid-parser-py が返す graph_data 辞書（vertices/edges を含む）。
        mermaid_text : str
            元のMermaidテキスト（フォールバック用）。
        left : int
            描画エリアの左端座標（EMU）。
        top : int
            描画エリアの上端座標（EMU）。
        width : int
            描画エリアの幅（EMU）。
        height : int
            描画エリアの高さ（EMU）。
        """
        vertices: dict = graph_data.get("vertices", {})
        raw_edges: list = graph_data.get("edges", [])

        if not vertices:
            self._render_fallback(slide, mermaid_text, left, top, width, height)
            return

        nodes = list(vertices.keys())

        # NetworkXグラフを構築してレイアウトを計算する
        G = nx.DiGraph()
        G.add_nodes_from(nodes)
        for edge in raw_edges:
            if isinstance(edge, dict):
                src = edge.get("start")
                dst = edge.get("end")
                if src is not None and dst is not None:
                    G.add_edge(str(src), str(dst))

        # mermaid_textから描画方向を取得して階層レイアウトを計算する
        direction = self._extract_direction(mermaid_text)
        pos: dict[str, tuple[float, float]] = self._hierarchical_layout(G, direction)

        # 頂点をシェイプ種別に応じて描画する
        node_shapes = self._draw_nodes_flowchart(
            slide, vertices, pos, left, top, width, height
        )

        # エッジを矢印・線種・ラベルに応じて描画する
        self._draw_edges_flowchart(
            slide, raw_edges, pos, node_shapes, left, top, width, height
        )

    def _draw_nodes_flowchart(
        self,
        slide: Slide,
        vertices: dict,
        pos: dict[str, tuple[float, float]],
        left: int,
        top: int,
        width: int,
        height: int,
    ) -> dict[str, object]:
        """
        フローチャートのノードをシェイプ種別に応じて描画する。

        vertices の各エントリから "type" フィールドを取得し、
        _SHAPE_MAP に基づいてAutoShapeTypeを選択する。
        lean_left は水平フリップ、inv_trapezoid は垂直フリップを追加で適用する。

        Parameters
        ----------
        slide : Slide
            python-pptxのSlideオブジェクト。
        vertices : dict
            ノードID → {"text": str, "type": str, ...} の辞書。
        pos : dict[str, tuple[float, float]]
            ノードIDをキー、正規化座標を値とする辞書。
        left, top, width, height : int
            描画エリアのEMU座標。

        Returns
        -------
        dict[str, object]
            ノードIDをキー、Shapeオブジェクトを値とする辞書。
        """
        node_shapes: dict[str, object] = {}

        for node_id, vertex in vertices.items():
            if node_id not in pos:
                continue

            x_norm, y_norm = pos[node_id]
            cx, cy = self._pos_to_emu(x_norm, y_norm, left, top, width, height)
            shape_left = cx - NODE_WIDTH_EMU // 2
            shape_top = cy - NODE_HEIGHT_EMU // 2

            # ノード種別に対応するAutoShapeTypeを取得する（未知の場合は矩形）
            node_type = (
                vertex.get("type", "square")
                if isinstance(vertex, dict)
                else "square"
            )
            shape_type = _SHAPE_MAP.get(node_type, 1)

            shape = slide.shapes.add_shape(
                shape_type,
                Emu(shape_left),
                Emu(shape_top),
                Emu(NODE_WIDTH_EMU),
                Emu(NODE_HEIGHT_EMU),
            )

            # lean_left（左傾き平行四辺形）は水平フリップを適用する
            if node_type == "lean_left":
                self._apply_flip(shape, flip_h=True, flip_v=False)

            # inv_trapezoid（逆台形）はprstGeomをinvertedTrapezoidに書き換える
            # flipVを使うとテキストフレームも反転するため、XMLで輪郭の形状のみ変更する
            if node_type == "inv_trapezoid":
                sp_el = shape._element
                spPr = sp_el.find(qn("p:spPr"))
                if spPr is not None:
                    prstGeom = spPr.find(qn("a:prstGeom"))
                    if prstGeom is not None:
                        prstGeom.set("prst", "invertedTrapezoid")

            # 表示ラベルを設定する（text フィールドを優先、なければノードID）
            label = (
                vertex.get("text", node_id)
                if isinstance(vertex, dict)
                else str(node_id)
            )
            shape.text = label
            tf = shape.text_frame
            for para in tf.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(12)

            node_shapes[node_id] = shape

        return node_shapes

    def _apply_flip(self, shape: object, flip_h: bool, flip_v: bool) -> None:
        """
        shapeのDrawingML XML要素にフリップ属性を設定する。

        Parameters
        ----------
        shape : object
            python-pptxのShapeオブジェクト。
        flip_h : bool
            水平フリップを適用する場合True。
        flip_v : bool
            垂直フリップを適用する場合True。
        """
        sp_el = shape._element
        spPr = sp_el.find(qn("p:spPr"))
        if spPr is None:
            return
        xfrm = spPr.find(qn("a:xfrm"))
        if xfrm is None:
            return
        if flip_h:
            xfrm.set("flipH", "1")
        if flip_v:
            xfrm.set("flipV", "1")

    def _draw_edges_flowchart(
        self,
        slide: Slide,
        raw_edges: list,
        pos: dict[str, tuple[float, float]],
        node_shapes: dict[str, object],
        left: int,
        top: int,
        width: int,
        height: int,
    ) -> None:
        """
        フローチャートのエッジを矢印種別・線種・ラベルに応じて描画する。

        stroke="invisible" のエッジはコネクター自体を生成しない。
        stroke="dotted" は破線、stroke="thick" は3pt実線を適用する。
        edge_type で headEnd/tailEnd の矢印形状を決定する。
        text が設定されている場合はコネクターXMLにラベルテキストを直接設定する。

        Parameters
        ----------
        slide : Slide
            python-pptxのSlideオブジェクト。
        raw_edges : list
            mermaid-parser-py の edges リスト。各要素に start/end/stroke/type/text を持つ。
        pos : dict[str, tuple[float, float]]
            ノードIDをキー、正規化座標を値とする辞書。
        node_shapes : dict[str, object]
            ノードIDをキー、Shapeオブジェクトを値とする辞書。
        left, top, width, height : int
            描画エリアのEMU座標。
        """
        for edge in raw_edges:
            if not isinstance(edge, dict):
                continue

            src = str(edge.get("start", ""))
            dst = str(edge.get("end", ""))
            if not src or not dst:
                continue
            if src not in pos or dst not in pos:
                continue

            stroke = edge.get("stroke", "normal")
            edge_type = edge.get("type", "arrow_point")
            edge_label: str = edge.get("text", "") or ""

            # invisible エッジはコネクターを生成しない
            if stroke == "invisible":
                continue

            sx_norm, sy_norm = pos[src]
            dx_norm, dy_norm = pos[dst]
            sx, sy = self._pos_to_emu(sx_norm, sy_norm, left, top, width, height)
            dx, dy = self._pos_to_emu(dx_norm, dy_norm, left, top, width, height)

            # 直線コネクターを追加する
            connector = slide.shapes.add_connector(
                1,  # MSO_CONNECTOR_TYPE.STRAIGHT
                Emu(sx), Emu(sy), Emu(dx), Emu(dy),
            )

            # 接続ポイントを決定してshapeに接続する
            src_cp, dst_cp = self._connection_indices(
                dx_norm - sx_norm, dy_norm - sy_norm
            )
            src_shape = node_shapes.get(src)
            dst_shape = node_shapes.get(dst)
            if src_shape is not None:
                connector.begin_connect(src_shape, src_cp)
            if dst_shape is not None:
                connector.end_connect(dst_shape, dst_cp)

            # コネクター内の <a:ln> XML要素を取得または作成する
            cxn_el = connector._element
            spPr = cxn_el.find(qn("p:spPr"))
            if spPr is None:
                continue
            ln = spPr.find(qn("a:ln"))
            if ln is None:
                ln = lxml_etree.SubElement(spPr, qn("a:ln"))

            # 線種を適用する
            if stroke == "dotted":
                # 破線スタイル
                prstDash = lxml_etree.SubElement(ln, qn("a:prstDash"))
                prstDash.set("val", "dash")
            elif stroke == "thick":
                # 太線（3pt = 38100 EMU）
                ln.set("w", "38100")

            # 矢印種別を適用する（未知の場合は通常矢印）
            arrow_conf = _EDGE_ARROW_MAP.get(edge_type, {"headEnd": "arrow"})
            if "headEnd" in arrow_conf:
                head = lxml_etree.SubElement(ln, qn("a:headEnd"))
                head.set("type", arrow_conf["headEnd"])
                head.set("w", "med")
                head.set("len", "med")
            if "tailEnd" in arrow_conf:
                tail = lxml_etree.SubElement(ln, qn("a:tailEnd"))
                tail.set("type", arrow_conf["tailEnd"])
                tail.set("w", "med")
                tail.set("len", "med")

            # ラベルがある場合はコネクターXMLにテキストを直接設定する（グループ化不要）
            if edge_label:
                self._set_connector_label(connector, edge_label)

    def _extract_direction(self, mermaid_text: str) -> str:
        """
        Mermaidテキストの宣言行から描画方向を抽出する。

        "flowchart LR" や "graph TD" のような宣言行から
        方向キーワードを取得する。見つからない場合は "TD"（上→下）を返す。

        Parameters
        ----------
        mermaid_text : str
            Mermaidフローチャートのテキスト。

        Returns
        -------
        str
            描画方向（"TD" / "TB" / "BT" / "LR" / "RL"）。デフォルトは "TD"。
        """
        match = re.search(
            r'(?:flowchart|graph)\s+(TD|TB|LR|RL|BT)\b',
            mermaid_text,
            re.IGNORECASE,
        )
        return match.group(1).upper() if match else "TD"

    def _hierarchical_layout(
        self,
        G: nx.DiGraph,
        direction: str,
    ) -> dict[str, tuple[float, float]]:
        """
        描画方向に基づいたDAG階層レイアウト座標を計算する。

        グラフがDAGである場合はトポロジカル世代でレベルを決定し、
        各レベル内のノードを均等配置した正規化座標（-1.0〜1.0）を返す。
        サイクルを含むグラフはkamada_kawai → spring_layoutにフォールバックする。

        Parameters
        ----------
        G : nx.DiGraph
            レイアウト計算対象の有向グラフ。
        direction : str
            描画方向（"TD" / "TB" / "BT" / "LR" / "RL"）。

        Returns
        -------
        dict[str, tuple[float, float]]
            ノードIDをキー、正規化座標(-1.0〜1.0)のタプルを値とする辞書。
        """
        # サイクルを含む場合はkamada_kawai → spring_layoutにフォールバック
        if not nx.is_directed_acyclic_graph(G):
            try:
                return nx.kamada_kawai_layout(G)
            except Exception:
                return nx.spring_layout(G, seed=42, k=2.0)

        # トポロジカル世代でノードをレベルに分類する
        generations = list(nx.topological_generations(G))
        n_levels = len(generations)
        pos: dict[str, tuple[float, float]] = {}

        for level_idx, level_nodes in enumerate(generations):
            n_in_level = len(level_nodes)
            # レベル軸の正規化座標（最初の世代=-1.0, 最後の世代=+1.0）
            level_norm = (
                -1.0 + 2.0 * level_idx / (n_levels - 1)
                if n_levels > 1
                else 0.0
            )
            for node_idx, node_id in enumerate(sorted(level_nodes)):
                # 同レベル内での均等配置座標
                span_norm = (
                    -1.0 + 2.0 * node_idx / (n_in_level - 1)
                    if n_in_level > 1
                    else 0.0
                )
                if direction in ("TD", "TB"):
                    # 上→下: level_normがy軸（-1.0=上端, +1.0=下端）
                    pos[node_id] = (span_norm, level_norm)
                elif direction == "BT":
                    # 下→上: y軸を反転
                    pos[node_id] = (span_norm, -level_norm)
                elif direction == "LR":
                    # 左→右: level_normがx軸（-1.0=左端, +1.0=右端）
                    pos[node_id] = (level_norm, span_norm)
                elif direction == "RL":
                    # 右→左: x軸を反転
                    pos[node_id] = (-level_norm, span_norm)
                else:
                    # 未知の方向はTD（上→下）として扱う
                    pos[node_id] = (span_norm, level_norm)

        return pos

    def _set_connector_label(
        self,
        connector: object,
        text: str,
        font_size_pt: int = 10,
    ) -> None:
        """
        コネクターのDrawingML XML要素にテキストラベルを直接設定する。

        <p:cxnSp> 要素に <p:txBody> を追加してラベルをコネクター自体に持たせる。
        グループ化による接続ポイント消失を回避する。

        Parameters
        ----------
        connector : object
            python-pptxのConnectorオブジェクト。
        text : str
            設定するラベルテキスト。
        font_size_pt : int
            フォントサイズ（ポイント）。デフォルト10pt。
        """
        cxn_el = connector._element
        # 既存のtxBodyを削除して重複追加を防ぐ
        for existing in cxn_el.findall(qn("p:txBody")):
            cxn_el.remove(existing)

        # <p:txBody> を構築してcxnSp要素に追加する
        txBody = lxml_etree.SubElement(cxn_el, qn("p:txBody"))

        # テキスト本文プロパティ: 垂直中央揃え・折り返しあり
        bodyPr = lxml_etree.SubElement(txBody, qn("a:bodyPr"))
        bodyPr.set("anchor", "ctr")
        bodyPr.set("wrap", "square")

        # リストスタイル（OOXML必須要素）
        lxml_etree.SubElement(txBody, qn("a:lstStyle"))

        # テキスト段落とランを追加する
        p_el = lxml_etree.SubElement(txBody, qn("a:p"))
        r_el = lxml_etree.SubElement(p_el, qn("a:r"))
        rPr = lxml_etree.SubElement(r_el, qn("a:rPr"))
        rPr.set("dirty", "0")
        rPr.set("sz", str(font_size_pt * 100))  # sz は 1/100ポイント単位
        t_el = lxml_etree.SubElement(r_el, qn("a:t"))
        t_el.text = text
