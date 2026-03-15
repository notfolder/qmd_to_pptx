"""
Mermaid図レンダラーの基底クラスモジュール。

各ダイアグラム種別レンダラーが継承する共通ユーティリティメソッドを提供する。
"""

from __future__ import annotations

from lxml import etree as lxml_etree
from pptx.oxml.ns import qn
from pptx.slide import Slide
from pptx.util import Emu, Pt


# ノードのデフォルトサイズ（EMU）
NODE_WIDTH_EMU: int = 1200000
NODE_HEIGHT_EMU: int = 500000


class BaseDiagramRenderer:
    """
    全Mermaidダイアグラムレンダラーの基底クラス。

    正規化座標変換・ノード描画・エッジ描画・フォールバック描画の
    共通ユーティリティメソッドを提供する。
    各サブクラスはこのクラスを継承して専用のrender()を実装する。
    """

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
        margin_x = NODE_WIDTH_EMU // 2
        margin_y = NODE_HEIGHT_EMU // 2
        usable_w = width - NODE_WIDTH_EMU
        usable_h = height - NODE_HEIGHT_EMU

        # -1.0〜1.0 を 0〜1 に正規化する
        x_ratio = (x_norm + 1.0) / 2.0
        y_ratio = (y_norm + 1.0) / 2.0

        x_emu = left + margin_x + int(x_ratio * usable_w)
        y_emu = top + margin_y + int(y_ratio * usable_h)
        return x_emu, y_emu

    def _connection_indices(
        self,
        dx: float,
        dy: float,
    ) -> tuple[int, int]:
        """
        方向ベクトル (dx, dy) から始点・終点の接続ポイントインデックスを決定する。

        DrawingML 標準矩形の接続ポイントは以下のとおり。
        - 0: 上辺中点
        - 1: 右辺中点
        - 2: 下辺中点
        - 3: 左辺中点

        Parameters
        ----------
        dx : float
            X方向の差（dst.x - src.x）。正規化座標系での値。
        dy : float
            Y方向の差（dst.y - src.y）。正規化座標系での値。

        Returns
        -------
        tuple[int, int]
            (始点接続ポイントインデックス, 終点接続ポイントインデックス)
        """
        if abs(dx) >= abs(dy):
            # 水平方向が支配的な場合
            if dx >= 0:
                # 左から右: src右辺 → dst左辺
                return (1, 3)
            else:
                # 右から左: src左辺 → dst右辺
                return (3, 1)
        else:
            # 垂直方向が支配的な場合
            if dy >= 0:
                # 上から下: src下辺 → dst上辺
                return (2, 0)
            else:
                # 下から上: src上辺 → dst下辺
                return (0, 2)

    def _draw_nodes(
        self,
        slide: Slide,
        nodes: list[str],
        pos: dict[str, tuple[float, float]],
        left: int,
        top: int,
        width: int,
        height: int,
        label_map: dict[str, str] | None = None,
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
        label_map : dict[str, str] | None
            ノードIDをキー、表示ラベル文字列を値とする辞書。省略時はノードIDを表示する。

        Returns
        -------
        dict[str, object]
            ノードIDをキー、Shapeオブジェクトを値とする辞書。
        """
        node_shapes: dict[str, object] = {}
        for node_id in nodes:
            if node_id not in pos:
                continue
            x_norm, y_norm = pos[node_id]
            cx, cy = self._pos_to_emu(x_norm, y_norm, left, top, width, height)
            shape_left = cx - NODE_WIDTH_EMU // 2
            shape_top = cy - NODE_HEIGHT_EMU // 2

            # 矩形Shapeを追加する
            shape = slide.shapes.add_shape(
                1,  # MSO_AUTO_SHAPE_TYPE.RECTANGLE
                Emu(shape_left),
                Emu(shape_top),
                Emu(NODE_WIDTH_EMU),
                Emu(NODE_HEIGHT_EMU),
            )
            # 表示ラベルを決定する（label_mapがあればそれを優先する）
            label = label_map.get(node_id, node_id) if label_map else node_id
            shape.text = label
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
        node_shapes: dict[str, object],
        left: int,
        top: int,
        width: int,
        height: int,
    ) -> None:
        """
        エッジをConnectorShapeとしてスライドに描画し、始点・終点のshapeに接続する。

        Parameters
        ----------
        slide : Slide
            python-pptxのSlideオブジェクト。
        edges : list[tuple[str, str]]
            (始点ノードID, 終点ノードID) のタプルリスト。
        pos : dict[str, tuple[float, float]]
            ノードIDをキー、正規化座標を値とする辞書。
        node_shapes : dict[str, object]
            ノードIDをキー、Shapeオブジェクトを値とする辞書。
        left : int
            描画エリアの左端座標（EMU）。
        top : int
            描画エリアの上端座標（EMU）。
        width : int
            描画エリアの幅（EMU）。
        height : int
            描画エリアの高さ（EMU）。
        """
        for src, dst in edges:
            if src not in pos or dst not in pos:
                continue
            sx_norm, sy_norm = pos[src]
            dx_norm, dy_norm = pos[dst]
            sx, sy = self._pos_to_emu(sx_norm, sy_norm, left, top, width, height)
            dx, dy = self._pos_to_emu(dx_norm, dy_norm, left, top, width, height)

            # 直線コネクターを追加する（connector_type=1 = STRAIGHT）
            connector = slide.shapes.add_connector(
                1,  # MSO_CONNECTOR_TYPE.STRAIGHT
                Emu(sx),
                Emu(sy),
                Emu(dx),
                Emu(dy),
            )

            # 方向ベクトルから接続ポイントインデックスを決定する
            src_cp, dst_cp = self._connection_indices(
                dx_norm - sx_norm, dy_norm - sy_norm
            )

            # shapeに begin_connect/end_connect でブロックに接続する
            src_shape = node_shapes.get(src)
            dst_shape = node_shapes.get(dst)
            if src_shape is not None:
                connector.begin_connect(src_shape, src_cp)
            if dst_shape is not None:
                connector.end_connect(dst_shape, dst_cp)

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
