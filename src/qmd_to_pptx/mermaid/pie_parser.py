"""
Mermaid 円グラフ（pie chart）カスタムパーサーモジュール。

mermaid-parser-py は pie 構文の graph_data を空で返すため、
Mermaid テキストを直接正規表現で解析する独自パーサーを実装する。

サポートする構文（Mermaid 公式仕様準拠）:
    pie [showData] [title <タイトルテキスト>]
        "<ラベル>" : <正の数値（小数点2桁まで）>
        ...

オプション構成:
    - showData   : セクションラベルに実数値を合わせて表示する
    - title      : グラフタイトル（省略可）
    - textPosition: YAML front-matter の config.pie.textPosition を参照する（デフォルト 0.75）
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field


@dataclass
class PieSection:
    """円グラフの1セクション（スライス）を表すデータクラス。"""

    label: str    # 凡例・データラベルに表示するテキスト
    value: float  # 正の数値（Mermaid は 0 より大きい値のみ許容する）


@dataclass
class PieChart:
    """
    Mermaid pie chart 全体を表すデータクラス。

    Attributes
    ----------
    title : str
        グラフタイトル。空文字列の場合はタイトルなし。
    show_data : bool
        True の場合、データラベルに実数値を合わせて表示する（showData キーワード）。
    text_position : float
        データラベルの半径方向位置。0.0（中心）〜 1.0（外縁）。デフォルト 0.75。
    sections : list[PieSection]
        セクションリスト。Mermaid の記述順（時計回り表示順）を保持する。
    """

    title: str = ""
    show_data: bool = False
    text_position: float = 0.75
    sections: list[PieSection] = field(default_factory=list)


# ---------------------------------------------------------------------------
# 正規表現パターン
# ---------------------------------------------------------------------------

# pie ヘッダー行: "pie [showData] [title ...]"
_RE_HEADER = re.compile(
    r"^\s*pie(?P<rest>.*)",
    re.IGNORECASE,
)

# showData キーワード（ヘッダー行の rest 部分から検索する）
_RE_SHOW_DATA = re.compile(r"\bshowData\b", re.IGNORECASE)

# title キーワード（ヘッダー行 rest 部またはそれ以降の行）
_RE_TITLE_INLINE = re.compile(r"\btitle\s+(.+)", re.IGNORECASE)

# セクション行: "  "ラベル" : 数値"
_RE_SECTION = re.compile(
    r'^\s*"(?P<label>[^"]+)"\s*:\s*(?P<value>\d+(?:\.\d+)?)\s*$'
)

# config.pie.textPosition（YAML front-matter 内から取得する）
_RE_TEXT_POSITION = re.compile(
    r"textPosition\s*:\s*(?P<val>[01](?:\.\d+)?)",
    re.IGNORECASE,
)


def parse_pie(mermaid_text: str) -> PieChart:
    """
    Mermaid の pie chart テキストを解析して PieChart データクラスを返す。

    バリデーション:
    - ヘッダー行が "pie" で始まらない場合は ValueError を送出する。
    - セクションの数値が 0 以下の場合はそのセクションをスキップする。

    Parameters
    ----------
    mermaid_text : str
        Mermaid テキスト（複数行）。YAML front-matter を含む場合も許容する。

    Returns
    -------
    PieChart
        解析結果を格納したデータクラス。
    """
    lines = mermaid_text.strip().splitlines()

    chart = PieChart()

    # YAML front-matter（--- で囲まれたブロック）から textPosition を取得する
    yaml_block = _extract_yaml_front_matter(mermaid_text)
    if yaml_block:
        m = _RE_TEXT_POSITION.search(yaml_block)
        if m:
            chart.text_position = float(m.group("val"))

    # "pie" で始まるヘッダー行を探す（YAML front-matter 後の行から探す）
    # front-matter が終わった後の行を処理する
    in_front_matter = False
    front_matter_done = False
    pie_header_found = False
    pie_lines: list[str] = []

    for raw_line in lines:
        stripped = raw_line.strip()

        # YAML front-matter の開始・終了を追跡する
        if not front_matter_done:
            if stripped == "---" and not in_front_matter:
                in_front_matter = True
                continue
            if stripped == "---" and in_front_matter:
                front_matter_done = True
                continue
            if in_front_matter:
                continue
            # front-matter なし
            front_matter_done = True

        if not pie_header_found:
            m = _RE_HEADER.match(raw_line)
            if m:
                pie_header_found = True
                rest = m.group("rest")

                # showData キーワードの検出
                if _RE_SHOW_DATA.search(rest):
                    chart.show_data = True

                # インラインタイトルの検出（"pie showData title ..." の形式）
                tm = _RE_TITLE_INLINE.search(rest)
                if tm:
                    chart.title = tm.group(1).strip()
        else:
            pie_lines.append(raw_line)

    if not pie_header_found:
        raise ValueError("pie ヘッダー行が見つかりません")

    # pie_lines を処理する（title 単独行、セクション行）
    for raw_line in pie_lines:
        stripped = raw_line.strip()
        if not stripped or stripped.startswith("%%"):
            # 空行・コメント行はスキップする
            continue

        # title 単独行（"title テキスト"）
        tm = _RE_TITLE_INLINE.match(stripped)
        if tm:
            # インラインタイトルより単独行を優先する（後入れ優先）
            chart.title = tm.group(1).strip()
            continue

        # セクション行
        sm = _RE_SECTION.match(raw_line)
        if sm:
            value = float(sm.group("value"))
            if value > 0:
                chart.sections.append(
                    PieSection(label=sm.group("label"), value=value)
                )

    return chart


def _extract_yaml_front_matter(text: str) -> str:
    """
    テキストの先頭にある YAML front-matter ブロック（--- で囲まれた部分）を抽出する。

    Parameters
    ----------
    text : str
        入力テキスト。

    Returns
    -------
    str
        YAML front-matter の中身。front-matter がない場合は空文字列。
    """
    stripped = text.lstrip()
    if not stripped.startswith("---"):
        return ""
    # 2 番目の "---" を探す
    end_idx = stripped.find("\n---", 3)
    if end_idx == -1:
        return ""
    return stripped[3:end_idx]
