"""
MermaidRendererファサードモジュール。

各ダイアグラム種別レンダラーへ処理を委譲するファサードクラスを提供する。
"""

from __future__ import annotations

import io
import contextlib
import logging
import re
import xml.etree.ElementTree as ET

# mermaid パッケージが IPython 未インストール時に stdout へ警告を print() するため、
# インポート時のみ stdout を抑制する
with contextlib.redirect_stdout(io.StringIO()):
    from mermaid_parser import MermaidParser

from pptx.slide import Slide

# モジュールロガーを取得する
logger = logging.getLogger(__name__)

from .base import BaseDiagramRenderer
from .class_diagram import ClassDiagramRenderer
from .er_diagram import ErDiagramRenderer
from .flowchart import FlowchartRenderer
from .gantt_parser import parse_gantt
from .gantt_renderer import GanttRenderer
from .gitgraph_parser import parse_gitgraph
from .gitgraph_renderer import GitGraphRenderer
from .journey_parser import parse_journey
from .journey_renderer import JourneyRenderer
from .mindmap import MindmapRenderer
from .pie_parser import parse_pie
from .pie_renderer import PieChartRenderer
from .quadrant_parser import parse_quadrant
from .quadrant_renderer import QuadrantRenderer
from .requirement_parser import parse_requirement
from .requirement_renderer import RequirementRenderer
from .sequence_diagram import SequenceDiagramRenderer
from .state_diagram import StateDiagramRenderer
from .timeline_parser import parse_timeline
from .timeline_renderer import TimelineRenderer


class MermaidRenderer:
    """
    Mermaidレンダラーファサードクラス。

    elementからMermaidテキストを取り出し、graph_typeに応じて
    専用レンダラーへ処理を委譲する。
    パース失敗時やサポート外ダイアグラム種別はフォールバック描画を行う。
    """

    def __init__(self) -> None:
        """各ダイアグラム種別レンダラーのインスタンスを生成する。"""
        self._flowchart = FlowchartRenderer()
        self._class_diagram = ClassDiagramRenderer()
        self._state_diagram = StateDiagramRenderer()
        self._er_diagram = ErDiagramRenderer()
        self._mindmap = MindmapRenderer()
        self._sequence = SequenceDiagramRenderer()
        self._gantt = GanttRenderer()
        self._gitgraph = GitGraphRenderer()
        self._journey = JourneyRenderer()
        self._pie = PieChartRenderer()
        self._quadrant = QuadrantRenderer()
        self._requirement = RequirementRenderer()
        self._timeline = TimelineRenderer()
        # フォールバック描画は BaseDiagramRenderer に委ねる
        self._base = BaseDiagramRenderer()

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

        # JSエンジンが対応していないダイアグラムタイプはパーサーを呼ばずフォールバックする
        # これにより mermaid-parser-py の JS エンジンが stderr に出力するエラーを抑止する
        _UNSUPPORTED_PREFIXES = (
            "zenuml",
        )
        first_line = mermaid_text.splitlines()[0].strip().lower() if mermaid_text else ""
        if any(first_line.startswith(p) for p in _UNSUPPORTED_PREFIXES):
            self._base._render_fallback(slide, mermaid_text, left, top, width, height)
            return

        # gitGraph は mermaid-parser-py が graph_data をメソッド経由で保持するため
        # JSON.stringify で空になる。カスタムパーサーへ直接委譲する
        if first_line.startswith("gitgraph"):
            try:
                git_graph = parse_gitgraph(mermaid_text)
            except Exception:
                self._base._render_fallback(slide, mermaid_text, left, top, width, height)
                return
            self._gitgraph.render(slide, git_graph, left, top, width, height)
            return

        # pie は mermaid-parser-py が graph_data を空で返すため、
        # パーサー呼び出し前にカスタムパーサーへ直接委譲する
        if first_line.startswith("pie"):
            try:
                pie_chart = parse_pie(mermaid_text)
            except Exception:
                self._base._render_fallback(slide, mermaid_text, left, top, width, height)
                return
            self._pie.render(slide, pie_chart, left, top, width, height)
            return

        # journey は mermaid-parser-py が graph_data を空で返すため、
        # パーサー呼び出し前にカスタムパーサーへ直接委譲する
        if first_line.startswith("journey"):
            try:
                journey_chart = parse_journey(mermaid_text)
            except Exception:
                self._base._render_fallback(slide, mermaid_text, left, top, width, height)
                return
            self._journey.render(slide, journey_chart, left, top, width, height)
            return

        # requirementDiagram は mermaid-parser-py の Jison レキサーが
        # ASCII \w 限定のため日本語で SpiderMonkeyError を送出する。
        # カスタムパーサーへ直接委譲する
        if first_line.startswith("requirementdiagram"):
            try:
                req_diagram = parse_requirement(mermaid_text)
            except Exception:
                self._base._render_fallback(slide, mermaid_text, left, top, width, height)
                return
            self._requirement.render(slide, req_diagram, left, top, width, height)
            return

        # quadrantChart は mermaid-parser-py が graph_data を空で返し、
        # 日本語テキストでは SpiderMonkeyError を送出するため、
        # パーサー呼び出し前にカスタムパーサーへ直接委譲する
        if first_line.startswith("quadrantchart"):
            try:
                quadrant_chart = parse_quadrant(mermaid_text)
            except Exception:
                self._base._render_fallback(slide, mermaid_text, left, top, width, height)
                return
            self._quadrant.render(slide, quadrant_chart, left, top, width, height)
            return

        # timeline は mermaid-parser-py が graph_data を空で返すため、
        # パーサー呼び出し前にカスタムパーサーへ直接委譲する
        if first_line.startswith("timeline"):
            try:
                timeline_data = parse_timeline(mermaid_text)
            except Exception:
                self._base._render_fallback(slide, mermaid_text, left, top, width, height)
                return
            self._timeline.render(slide, timeline_data, left, top, width, height)
            return

        # gantt は mermaid-parser-py が graph_data を返せないため、
        # パーサー呼び出し前にカスタムパーサーへ直接委譲する
        if first_line.startswith("gantt"):
            try:
                gantt_chart = parse_gantt(mermaid_text)
            except Exception:
                self._base._render_fallback(slide, mermaid_text, left, top, width, height)
                return
            self._gantt.render(slide, gantt_chart, left, top, width, height)
            return

        # sequenceDiagram は graph_type="sequence" で返るが、
        # パーサー呼び出し前に first_line で早期検出して直接委譲することもできる
        # （パーサーが正常に対応していることを確認済みのため、パーサー経由で処理する）

        try:
            # mermaid-parser-pyでノードとエッジを取得する
            mp = MermaidParser()
            result = mp.parse(mermaid_text)
            graph_data = result.get("graph_data", {})
            graph_type = result.get("graph_type", "")
        except Exception:
            # パース失敗時はテキストボックスにそのまま表示する
            self._base._render_fallback(slide, mermaid_text, left, top, width, height)
            return

        # graph_typeに応じて専用レンダラーへ分岐する
        if graph_type == "sequence":
            self._sequence.render(slide, graph_data, left, top, width, height)
            return
        if graph_type == "stateDiagram":
            # stateDiagram-v2 の direction キーワードを解析して graph_data に注入する
            # （graph_data root には direction フィールドが含まれないため、
            #   Mermaidテキストから直接抽出する）
            _dir_m = re.search(
                r'^\s*direction\s+(TD|TB|LR|RL|BT)\b',
                mermaid_text,
                re.IGNORECASE | re.MULTILINE,
            )
            graph_data["_direction"] = _dir_m.group(1).upper() if _dir_m else "TB"
            self._state_diagram.render(slide, graph_data, left, top, width, height)
            return
        if graph_type == "class":
            self._class_diagram.render(slide, graph_data, left, top, width, height)
            return
        if graph_type == "er":
            self._er_diagram.render(slide, graph_data, left, top, width, height)
            return
        if graph_type == "mindmap":
            self._mindmap.render(slide, graph_data, left, top, width, height)
            return
        if graph_type == "gantt":
            # graph_data={} になるケースのフォールバック（上の early-return で処理済みのため通常不達）
            try:
                gantt_chart = parse_gantt(mermaid_text)
            except Exception:
                pass
            else:
                self._gantt.render(slide, gantt_chart, left, top, width, height)
                return

        # flowchart / graph 系: FlowchartRenderer に委譲する
        vertices = graph_data.get("vertices", {})
        if not vertices:
            self._base._render_fallback(slide, mermaid_text, left, top, width, height)
            return

        self._flowchart.render(slide, graph_data, mermaid_text, left, top, width, height)

    # ------------------------------------------------------------------
    # 後方互換メソッド（既存テスト・外部コードからの直接呼び出しに対応する）
    # ------------------------------------------------------------------

    def _extract_nodes(self, graph_data: dict) -> list[str]:
        """
        mermaid-parser-pyの解析結果からノードIDのリストを取得する（後方互換）。

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
        mermaid-parser-pyの解析結果からエッジのリストを取得する（後方互換）。

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
        正規化座標（-1.0〜1.0）をEMU座標に変換する（後方互換）。

        Parameters
        ----------
        x_norm : float
            正規化X座標（-1.0〜1.0）。
        y_norm : float
            正規化Y座標（-1.0〜1.0）。
        left, top, width, height : int
            描画エリアのEMU座標。

        Returns
        -------
        tuple[int, int]
            (x_emu, y_emu) のタプル。
        """
        return self._base._pos_to_emu(x_norm, y_norm, left, top, width, height)
