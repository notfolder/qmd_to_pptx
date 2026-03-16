"""
Mermaid 象限チャート（quadrantChart）カスタムパーサーモジュール。

mermaid-parser-py は quadrantChart 構文の graph_data を空で返し、
さらに日本語テキストを含む場合は SpiderMonkeyError を送出するため、
Mermaid テキストを直接正規表現で解析する独自パーサーを実装する。

サポートする構文（Mermaid 公式仕様準拠）:
    quadrantChart
        title <タイトルテキスト>
        x-axis <左ラベル> [-->  <右ラベル>]
        y-axis <下ラベル> [--> <上ラベル>]
        quadrant-1 <テキスト>      ← 右上の象限
        quadrant-2 <テキスト>      ← 左上の象限
        quadrant-3 <テキスト>      ← 左下の象限
        quadrant-4 <テキスト>      ← 右下の象限
        <ポイント名>[:::クラス名]: [x, y] [スタイル属性, ...]
        classDef <クラス名> <スタイル属性, ...>

ポイントのスタイル属性（インライン・classDef 共通）:
    color: #rrggbb          --- 塗りつぶし色
    radius: <整数>           --- 半径（ピクセル相当）
    stroke-width: <整数>px   --- 枠線幅（ピクセル相当）
    stroke-color: #rrggbb   --- 枠線色

スタイル適用優先順位（高い → 低い）:
    1. インライン直接指定
    2. classDef クラス定義
    3. デフォルト値
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field


# ---------------------------------------------------------------------------
# データクラス定義
# ---------------------------------------------------------------------------


@dataclass
class PointStyle:
    """ポイントのスタイル情報を保持するデータクラス。"""

    color: str | None = None             # 塗りつぶし色 (#rrggbb)
    radius: int | None = None            # 半径（ピクセル相当整数値）
    stroke_width: int | None = None      # 枠線幅（ピクセル相当整数値）
    stroke_color: str | None = None      # 枠線色 (#rrggbb)


@dataclass
class QuadrantPoint:
    """象限チャートの1データポイントを表すデータクラス。"""

    name: str                                   # ポイント名（例: "Campaign A"）
    x: float                                    # X座標 0.0〜1.0（クランプ済み）
    y: float                                    # Y座標 0.0〜1.0（クランプ済み）
    class_name: str | None = None               # :::className で指定されたクラス名
    inline_style: PointStyle = field(default_factory=PointStyle)  # インラインスタイル


@dataclass
class QuadrantChart:
    """
    Mermaid quadrantChart 全体を表すデータクラス。

    Attributes
    ----------
    title : str
        グラフタイトル。空文字列の場合はタイトルなし。
    x_label_left : str
        X軸左端ラベル（低い側）。省略時は空文字列。
    x_label_right : str
        X軸右端ラベル（高い側）。省略時は空文字列。
    y_label_bottom : str
        Y軸下端ラベル（低い側）。省略時は空文字列。
    y_label_top : str
        Y軸上端ラベル（高い側）。省略時は空文字列。
    quadrant_labels : dict[int, str]
        象限ラベル辞書。キーは象限番号（1〜4）、値はラベルテキスト。
        1=右上、2=左上、3=左下、4=右下。
    points : list[QuadrantPoint]
        データポイントリスト（出現順）。
    class_defs : dict[str, PointStyle]
        classDef 定義辞書。キーはクラス名、値は PointStyle。
    """

    title: str = ""
    x_label_left: str = ""
    x_label_right: str = ""
    y_label_bottom: str = ""
    y_label_top: str = ""
    quadrant_labels: dict[int, str] = field(default_factory=dict)
    points: list[QuadrantPoint] = field(default_factory=list)
    class_defs: dict[str, PointStyle] = field(default_factory=dict)


# ---------------------------------------------------------------------------
# 正規表現パターン
# ---------------------------------------------------------------------------

# "quadrantChart" ヘッダー行
_RE_HEADER = re.compile(r"^\s*quadrantChart\s*$", re.IGNORECASE)

# "title <text>" 行
_RE_TITLE = re.compile(r"^\s*title\s+(.+)", re.IGNORECASE)

# "x-axis <left> [-->  <right>]" 行
_RE_X_AXIS = re.compile(
    r"^\s*x-axis\s+(?P<left>.+?)(?:\s*-->\s*(?P<right>.+))?\s*$",
    re.IGNORECASE,
)

# "y-axis <bottom> [--> <top>]" 行
_RE_Y_AXIS = re.compile(
    r"^\s*y-axis\s+(?P<bottom>.+?)(?:\s*-->\s*(?P<top>.+))?\s*$",
    re.IGNORECASE,
)

# "quadrant-N <text>" 行（N は 1〜4）
_RE_QUADRANT = re.compile(
    r"^\s*quadrant-(?P<num>[1-4])\s+(?P<text>.+)\s*$",
    re.IGNORECASE,
)

# ポイント行: "<名前>[:::クラス名]: [x, y] [スタイル...]"
# Mermaid の classDef 参照は ":::" で始まる（3個のコロン）。
# re.compile では (?:::...) は non-capturing group + ":" × 2 なので、
# Mermaid の 3コロン ":::" に対応するためには (?::::...) と書く必要がある。
# name は ":" も "[" も含めないことで、:::class が続く位置で正確に停止する。
_RE_POINT = re.compile(
    r"^\s*(?P<name>[^:[]+)(?::::(?P<class>[A-Za-z_][A-Za-z0-9_]*))?"
    r"\s*:\s*\[(?P<x>-?[0-9]*\.?[0-9]+)\s*,\s*(?P<y>-?[0-9]*\.?[0-9]+)\]"
    r"(?P<style_str>.*)$",
)

# "classDef <name> <styles>" 行
_RE_CLASSDEF = re.compile(
    r"^\s*classDef\s+(?P<name>[A-Za-z_][A-Za-z0-9_]*)\s+(?P<styles>.+)\s*$",
    re.IGNORECASE,
)

# コメント行 ("%%"で始まる)
_RE_COMMENT = re.compile(r"^\s*%%")

# アクセシビリティ記述行 ("accTitle:" / "accDescr:")
_RE_ACC = re.compile(r"^\s*acc(?:Title|Descr)\b", re.IGNORECASE)

# スタイル属性の1エントリを解析するパターン: "key: value"
_RE_STYLE_ENTRY = re.compile(
    r"(?P<key>color|radius|stroke-width|stroke-color)\s*:\s*(?P<value>[^,]+)",
    re.IGNORECASE,
)

# 色文字列の正規化パターン (#rrggbb または rrggbb)
_RE_COLOR = re.compile(r"^#?([0-9a-fA-F]{6})$")


# ---------------------------------------------------------------------------
# ヘルパー関数
# ---------------------------------------------------------------------------


def _parse_color(raw: str) -> str | None:
    """
    色文字列を #rrggbb 形式に正規化して返す。

    Parameters
    ----------
    raw : str
        色文字列（"#ff3300" または "ff3300" 形式）。

    Returns
    -------
    str | None
        "#rrggbb" 形式の文字列。無効な値の場合は None を返す。
    """
    raw = raw.strip()
    m = _RE_COLOR.match(raw)
    if m:
        return f"#{m.group(1).lower()}"
    return None


def _parse_styles(style_str: str) -> PointStyle:
    """
    スタイル文字列を解析して PointStyle を返す。

    スタイル文字列はカンマ区切りの "key: value" の組で構成される。
    認識できない属性は無視する。

    Parameters
    ----------
    style_str : str
        スタイル文字列（例: "color: #ff3300, radius: 10"）。

    Returns
    -------
    PointStyle
        解析結果のスタイルデータクラス。
    """
    style = PointStyle()
    for m in _RE_STYLE_ENTRY.finditer(style_str):
        key = m.group("key").lower()
        value = m.group("value").strip()
        if key == "color":
            style.color = _parse_color(value)
        elif key == "radius":
            try:
                style.radius = int(float(value))
            except ValueError:
                pass
        elif key == "stroke-width":
            # "Npx" 形式にも対応する（px を除去して整数変換）
            try:
                style.stroke_width = int(float(value.lower().replace("px", "").strip()))
            except ValueError:
                pass
        elif key == "stroke-color":
            style.stroke_color = _parse_color(value)
    return style


def _clamp_coord(value: float) -> float:
    """
    座標値を 0.0〜1.0 の範囲にクランプして返す。

    Parameters
    ----------
    value : float
        変換前の座標値。

    Returns
    -------
    float
        0.0〜1.0 にクランプされた座標値。
    """
    return max(0.0, min(1.0, value))


# ---------------------------------------------------------------------------
# メインパーサー関数
# ---------------------------------------------------------------------------


def parse_quadrant(text: str) -> QuadrantChart:
    """
    Mermaid quadrantChart テキストを解析して QuadrantChart を返す。

    文法エラーの行は警告なくスキップする。
    ポイント座標は 0.0〜1.0 の範囲にクランプする。

    Parameters
    ----------
    text : str
        Mermaid quadrantChart 構文のテキスト（"quadrantChart" ヘッダー行を含む）。

    Returns
    -------
    QuadrantChart
        解析結果のデータクラス。
    """
    chart = QuadrantChart()

    for line in text.splitlines():
        stripped = line.strip()

        # 空行・コメント行・アクセシビリティ行はスキップする
        if not stripped or _RE_COMMENT.match(line) or _RE_ACC.match(line):
            continue

        # "quadrantChart" ヘッダー行はスキップする
        if _RE_HEADER.match(line):
            continue

        # title 行
        m = _RE_TITLE.match(line)
        if m:
            chart.title = m.group(1).strip()
            continue

        # x-axis 行
        m = _RE_X_AXIS.match(line)
        if m:
            chart.x_label_left = m.group("left").strip()
            right = m.group("right")
            chart.x_label_right = right.strip() if right else ""
            continue

        # y-axis 行
        m = _RE_Y_AXIS.match(line)
        if m:
            chart.y_label_bottom = m.group("bottom").strip()
            top = m.group("top")
            chart.y_label_top = top.strip() if top else ""
            continue

        # quadrant-N 行
        m = _RE_QUADRANT.match(line)
        if m:
            num = int(m.group("num"))
            chart.quadrant_labels[num] = m.group("text").strip()
            continue

        # classDef 行（ポイント行より先に判定する）
        m = _RE_CLASSDEF.match(line)
        if m:
            class_name = m.group("name")
            styles = _parse_styles(m.group("styles"))
            chart.class_defs[class_name] = styles
            continue

        # ポイント行
        m = _RE_POINT.match(line)
        if m:
            name = m.group("name").strip()
            # ポイント名のクォートを除去する（"..." 形式に対応）
            if name.startswith('"') and name.endswith('"'):
                name = name[1:-1]
            x = _clamp_coord(float(m.group("x")))
            y = _clamp_coord(float(m.group("y")))
            class_name = m.group("class")
            style_str = m.group("style_str") or ""
            inline_style = _parse_styles(style_str)
            chart.points.append(
                QuadrantPoint(
                    name=name,
                    x=x,
                    y=y,
                    class_name=class_name,
                    inline_style=inline_style,
                )
            )
            continue

    return chart
