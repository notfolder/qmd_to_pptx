"""
マインドマップレンダラーモジュール。

mindmap をツリーレイアウトで左→右方向に描画する。
"""

from __future__ import annotations

from pptx.slide import Slide

from .base import BaseDiagramRenderer


class MindmapRenderer(BaseDiagramRenderer):
    """
    mindmap を描画するレンダラー。

    graph_data の "nodes" リストからツリー構造を読み取り、
    ルートノードを左端に配置して子ノードを右側へ深さに応じて展開する。
    縦位置は葉ノードの数を基準に均等分配する。
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
        mindmap をスライドに描画する。

        Parameters
        ----------
        slide : Slide
            python-pptxのSlideオブジェクト。
        graph_data : dict
            graph_data辞書（"nodes"キーにツリー構造を含む）。
        left, top, width, height : int
            描画エリアのEMU座標。
        """
        nodes_data = graph_data.get("nodes", [])
        if not nodes_data:
            return

        # 最初の要素がルートノード（children を含む完全ツリー）
        root = nodes_data[0]

        # ツリーを後順走査して深さと縦位置を収集する
        depth_map: dict[str, int] = {}
        order_map: dict[str, float] = {}
        counter: list[int] = [0]
        self._collect_layout(root, depth_map, order_map, counter)

        if not depth_map:
            return

        # descr フィールドを表示ラベルとして使う
        label_map: dict[str, str] = {
            n.get("nodeId", ""): n.get("descr", n.get("nodeId", ""))
            for n in nodes_data
            if n.get("nodeId")
        }

        # 正規化座標（-1.0〜1.0）を計算する（深さ→X、縦位置→Y）
        max_depth = max(depth_map.values()) if depth_map else 1
        leaf_count = counter[0]  # 葉ノードの総数
        total = max(leaf_count - 1, 1)  # ゼロ除算を防ぐ

        pos: dict[str, tuple[float, float]] = {}
        for node_id, depth in depth_map.items():
            x_norm = (depth / max_depth) * 2.0 - 1.0
            y_val = order_map.get(node_id, 0.0)
            y_norm = (y_val / total) * 2.0 - 1.0
            pos[node_id] = (x_norm, y_norm)

        # 親子エッジを収集する
        edges: list[tuple[str, str]] = []
        self._collect_edges(root, edges)

        all_node_ids = list(depth_map.keys())
        node_shapes = self._draw_nodes(
            slide, all_node_ids, pos, left, top, width, height, label_map=label_map
        )
        self._draw_edges(slide, edges, pos, node_shapes, left, top, width, height)

    def _collect_layout(
        self,
        node: dict,
        depth_map: dict[str, int],
        order_map: dict[str, float],
        counter: list[int],
    ) -> None:
        """
        マインドマップのツリーを後順走査して各ノードの深さと縦位置を収集する。

        Parameters
        ----------
        node : dict
            現在のノード（nodeId, level, children を持つ）。
        depth_map : dict[str, int]
            ノードIDをキーにした深さのマップ（レベル4=depth1, 8=depth2, ...）。
        order_map : dict[str, float]
            ノードIDをキーにした縦位置（葉ノードは整数、内部ノードは子の中央）。
        counter : list[int]
            葉ノードの通し番号（リストで参照渡し）。
        """
        node_id = node.get("nodeId", "")
        if not node_id:
            return

        # レベル値は4の倍数（4=depth1, 8=depth2, 12=depth3）
        depth_map[node_id] = node.get("level", 4) // 4
        children = node.get("children", [])

        if not children:
            # 葉ノード: 縦位置としてカウンター値を割り当てる
            order_map[node_id] = float(counter[0])
            counter[0] += 1
        else:
            # 内部ノード: 先に子を走査してから子の縦位置の平均を取る
            for child in children:
                self._collect_layout(child, depth_map, order_map, counter)
            child_orders = [
                order_map[c["nodeId"]]
                for c in children
                if c.get("nodeId") in order_map
            ]
            order_map[node_id] = (
                sum(child_orders) / len(child_orders) if child_orders else 0.0
            )

    def _collect_edges(
        self,
        node: dict,
        edges: list[tuple[str, str]],
    ) -> None:
        """
        マインドマップのツリーを走査して親子エッジを収集する。

        Parameters
        ----------
        node : dict
            現在のノード。
        edges : list[tuple[str, str]]
            (親nodeId, 子nodeId) タプルを追加するリスト。
        """
        node_id = node.get("nodeId", "")
        for child in node.get("children", []):
            child_id = child.get("nodeId", "")
            if node_id and child_id:
                edges.append((node_id, child_id))
            self._collect_edges(child, edges)
