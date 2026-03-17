"""
Mermaid ユーザージャーニー図カスタムパーサーモジュール。

mermaid-parser-py は journey 構文の graph_data を空で返すため、
Mermaid テキストを直接正規表現で解析する独自パーサーを実装する。

サポートする構文（Mermaid 公式仕様準拠）:
    journey
        title <タイトルテキスト>
        section <セクション名>
            <タスク名> : <スコア(1-5)> [: <アクター1>, <アクター2>, ...]

スコアは 1〜5 の整数値で、値が大きいほど感情が高い（満足・幸せ）。
アクターは省略可能。複数アクターはカンマ区切りで指定する。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field


@dataclass
class JourneyTask:
    """ユーザージャーニー図の1タスクを表すデータクラス。"""

    task: str            # タスク名
    score: int           # 感情スコア (1〜5)
    people: list[str]    # アクター名リスト（空リスト可）
    section: str         # 所属セクション名（セクション未定義の場合は空文字列）


@dataclass
class JourneyChart:
    """
    Mermaid journey 全体を表すデータクラス。

    Attributes
    ----------
    title : str
        ダイアグラムタイトル。省略時は空文字列。
    sections : list[str]
        セクション名リスト（出現順、重複なし）。
    tasks : list[JourneyTask]
        タスクリスト（出現順）。
    actors : list[str]
        全アクター名リスト（出現順、重複排除済み）。
    """

    title: str = ""
    sections: list[str] = field(default_factory=list)
    tasks: list[JourneyTask] = field(default_factory=list)
    actors: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# 正規表現パターン
# ---------------------------------------------------------------------------

# "journey" ヘッダー行: 先頭の "journey" キーワードのみ（以降は無視）
_RE_HEADER = re.compile(r"^\s*journey\s*$", re.IGNORECASE)

# "title <text>" 行: タイトルを取得する
_RE_TITLE = re.compile(r"^\s*title\s+(.+)", re.IGNORECASE)

# "section <name>" 行: セクション名を取得する
_RE_SECTION = re.compile(r"^\s*section\s+(.+)", re.IGNORECASE)

# タスク行: "<タスク名> : <スコア> [: <アクター1>, <アクター2>, ...]"
# タスク名はコロンを含まない（Mermaid 公式仕様）
# スコアは 1〜5 の整数（範囲外はクランプする）
_RE_TASK = re.compile(
    r"^\s*(?P<name>[^:]+?)\s*:\s*(?P<score>\d+)(?:\s*:\s*(?P<actors>.+))?\s*$"
)

# コメント行（"%%"で始まる、または "#" で始まる）
_RE_COMMENT = re.compile(r"^\s*(?:%%|#)")

# アクセシビリティ記述行（"accTitle:" / "accDescr:" / "accDescr{{"）をスキップ
_RE_ACC = re.compile(r"^\s*acc(?:Title|Descr)\b", re.IGNORECASE)


def parse_journey(text: str) -> JourneyChart:
    """
    Mermaid journey テキストを解析して JourneyChart を返す。

    文法エラーの行は警告なくスキップする。
    スコアは 1〜5 の範囲にクランプする。
    アクターは出現順で重複排除して JourneyChart.actors に格納する。

    Parameters
    ----------
    text : str
        Mermaid journey 構文のテキスト（"journey" ヘッダー行を含む）。

    Returns
    -------
    JourneyChart
        解析結果のデータクラス。
    """
    title: str = ""
    sections: list[str] = []
    tasks: list[JourneyTask] = []
    current_section: str = ""
    # 出現順を保持しながら重複排除するために dict をキューとして利用する
    seen_actors: dict[str, None] = {}

    for line in text.splitlines():
        stripped = line.strip()

        # 空行・コメント行・アクセシビリティ行はスキップする
        if not stripped or _RE_COMMENT.match(line) or _RE_ACC.match(line):
            continue

        # "journey" ヘッダー行はスキップする
        if _RE_HEADER.match(line):
            continue

        # title 行
        m = _RE_TITLE.match(line)
        if m:
            title = m.group(1).strip()
            continue

        # section 行
        m = _RE_SECTION.match(line)
        if m:
            sec_name = m.group(1).strip()
            current_section = sec_name
            if sec_name not in sections:
                sections.append(sec_name)
            continue

        # タスク行
        m = _RE_TASK.match(line)
        if m:
            name = m.group("name").strip()
            score = max(1, min(5, int(m.group("score"))))
            actors_raw = m.group("actors")
            if actors_raw:
                people = [p.strip() for p in actors_raw.split(",") if p.strip()]
            else:
                people = []
            tasks.append(
                JourneyTask(
                    task=name,
                    score=score,
                    people=people,
                    section=current_section,
                )
            )
            # 全アクターを出現順で重複排除する
            for person in people:
                seen_actors[person] = None
            continue

    # セクションなしでタスクがある場合: 空セクションを1つ登録する
    if not sections and tasks:
        sections = [""]

    actors = list(seen_actors.keys())

    return JourneyChart(
        title=title,
        sections=sections,
        tasks=tasks,
        actors=actors,
    )
