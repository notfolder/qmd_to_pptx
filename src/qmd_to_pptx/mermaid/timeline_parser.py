"""
Mermaid タイムライン（timeline）カスタムパーサーモジュール。

mermaid-parser-py は timeline の graph_data を JavaScript メソッド経由で保持するため、
JSON.stringify で関数がすべて除去され、空の結果しか返らない。
本モジュールはその代替として Python で直接行ベース状態機械パーサーを実装する。

サポートする構文（Mermaid 公式仕様 timeline 準拠）:

    timeline
        title <タイトルテキスト>
        section <セクション名>
        <period> : <event1> [: <event2> ...]   # 同一行に複数イベント
                 : <event>                       # 継続行によるイベント追加
    %% コメント行

テキスト内の <br> / <br/> / <br /> タグは改行（\\n）に変換する。
accTitle / accDescr はアクセシビリティ専用のためスキップする。

section キーワードで以降の period をグループ化できる。
section が存在しない場合、各 period が独立して色割り当ての単位になる。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field


# ---------------------------------------------------------------------------
# データクラス定義
# ---------------------------------------------------------------------------

@dataclass
class TimelineEvent:
    """
    タイムラインの1イベントを表すデータクラス。

    Attributes
    ----------
    text : str
        イベントテキスト。<br> タグは \\n に変換済み。
    """

    text: str


@dataclass
class TimelinePeriod:
    """
    タイムラインの1時期（period）とそれに属するイベントを表すデータクラス。

    Attributes
    ----------
    label : str
        期間ラベル（例: "2004", "Industry 1.0"）。<br> タグは \\n に変換済み。
    events : list[TimelineEvent]
        この期間に属するイベントのリスト。
    section : str | None
        所属するセクション名。セクション未定義の場合は None。
    """

    label: str
    events: list[TimelineEvent] = field(default_factory=list)
    section: str | None = None


@dataclass
class TimelineData:
    """
    Mermaid timeline 全体を表すデータクラス。

    Attributes
    ----------
    title : str
        ダイアグラムタイトル。省略時は空文字列。
    sections : list[str]
        セクション名リスト（出現順・重複なし）。
    periods : list[TimelinePeriod]
        期間リスト（出現順）。
    """

    title: str = ""
    sections: list[str] = field(default_factory=list)
    periods: list[TimelinePeriod] = field(default_factory=list)


# ---------------------------------------------------------------------------
# 正規表現パターン
# ---------------------------------------------------------------------------

# "timeline" ヘッダー行
_RE_HEADER = re.compile(r"^\s*timeline\s*$", re.IGNORECASE)

# "title <text>" 行
_RE_TITLE = re.compile(r"^\s*title\s+(.+)", re.IGNORECASE)

# "section <name>" 行（名前にコロンを含まない）
_RE_SECTION = re.compile(r"^\s*section\s+([^:\n]+)", re.IGNORECASE)

# コメント行（"%%"で始まる）
_RE_COMMENT = re.compile(r"^\s*%%")

# アクセシビリティ記述行（"accTitle:" / "accDescr:"）をスキップ
_RE_ACC = re.compile(r"^\s*acc(?:Title|Descr)\b", re.IGNORECASE)

# <br> タグ（各種バリエーション）を改行に変換するパターン
_RE_BR = re.compile(r"<br\s*/?>", re.IGNORECASE)

# コロン区切りパターン: ` : ` または行頭の `: ` を検出する
# Mermaid 公式仕様では period はコロンを含めないため、最初のコロン+スペースが区切り
_RE_COLON_SEP = re.compile(r"\s*:\s+")


# ---------------------------------------------------------------------------
# ユーティリティ関数
# ---------------------------------------------------------------------------

def _process_br(text: str) -> str:
    """
    <br>/<br/>/<br /> タグを改行文字（\\n）に変換する。

    Parameters
    ----------
    text : str
        変換対象テキスト。

    Returns
    -------
    str
        変換済みテキスト。
    """
    return _RE_BR.sub("\n", text)


def _split_by_colon(text: str) -> list[str]:
    """
    テキストをコロン区切り（` : `）で分割し、各部分をトリムして返す。

    Mermaid の仕様では `: ` （コロン+スペース）が区切り文字である。
    コロンの前後どちらかにスペースがない場合（例: `a:b`）は区切りとみなさない。

    Parameters
    ----------
    text : str
        分割対象テキスト。

    Returns
    -------
    list[str]
        分割・トリム済みのパーツリスト（空文字列は除外）。
    """
    parts = _RE_COLON_SEP.split(text)
    return [p.strip() for p in parts if p.strip()]


# ---------------------------------------------------------------------------
# メインパーサー関数
# ---------------------------------------------------------------------------

def parse_timeline(text: str) -> TimelineData:
    """
    Mermaid timeline テキストを解析して TimelineData を返す。

    文法エラーの行は警告なくスキップする。
    テキスト内の <br> タグは \\n に変換する。

    Parameters
    ----------
    text : str
        Mermaid timeline 構文のテキスト（"timeline" ヘッダー行を含む）。

    Returns
    -------
    TimelineData
        解析結果のデータクラス。
    """
    title: str = ""
    sections: list[str] = []
    periods: list[TimelinePeriod] = []
    current_section: str | None = None
    # セクション名の重複排除用
    seen_sections: dict[str, None] = {}

    for line in text.splitlines():
        stripped = line.strip()

        # 空行・コメント行・アクセシビリティ行はスキップ
        if not stripped or _RE_COMMENT.match(line) or _RE_ACC.match(line):
            continue

        # "timeline" ヘッダー行はスキップ
        if _RE_HEADER.match(stripped):
            continue

        # title 行
        m = _RE_TITLE.match(stripped)
        if m:
            title = _process_br(m.group(1).strip())
            continue

        # section 行
        m = _RE_SECTION.match(stripped)
        if m:
            sec_name = _process_br(m.group(1).strip())
            current_section = sec_name
            if sec_name not in seen_sections:
                seen_sections[sec_name] = None
                sections.append(sec_name)
            continue

        # 継続イベント行: `: <event>` で始まる行（先頭にコロン）
        if stripped.startswith(":"):
            if not periods:
                # period が未定義の継続行は無視
                continue
            rest = stripped[1:].strip()
            if not rest:
                continue
            # 残りをさらにコロン区切りで分割してイベントを追加
            event_parts = _split_by_colon(rest)
            for ep in event_parts:
                processed = _process_br(ep)
                if processed:
                    periods[-1].events.append(TimelineEvent(text=processed))
            continue

        # period + event(s) 行: `<period> : <event> [: <event> ...]`
        # コロン区切りが存在するかチェック
        if ":" in stripped:
            # 最初のコロン+スペースの位置を探す
            m_colon = _RE_COLON_SEP.search(stripped)
            if m_colon:
                period_label_raw = stripped[: m_colon.start()].strip()
                rest = stripped[m_colon.end():]
                period_label = _process_br(period_label_raw)
                if not period_label:
                    # period ラベルが空の場合は継続行として扱う
                    if periods:
                        event_parts = _split_by_colon(rest)
                        for ep in event_parts:
                            processed = _process_br(ep)
                            if processed:
                                periods[-1].events.append(TimelineEvent(text=processed))
                    continue

                # 新しい period を作成
                new_period = TimelinePeriod(
                    label=period_label,
                    section=current_section,
                )
                # 残りのコロン区切りをイベントとして追加
                event_parts = _split_by_colon(rest)
                for ep in event_parts:
                    processed = _process_br(ep)
                    if processed:
                        new_period.events.append(TimelineEvent(text=processed))
                periods.append(new_period)
                continue

        # コロンなし行は period のみ（イベントなし）として扱う
        if stripped:
            period_label = _process_br(stripped)
            new_period = TimelinePeriod(
                label=period_label,
                section=current_section,
            )
            periods.append(new_period)

    return TimelineData(
        title=title,
        sections=sections,
        periods=periods,
    )
