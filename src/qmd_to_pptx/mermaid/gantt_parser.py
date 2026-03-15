"""
Mermaidガントチャートカスタムパーサーモジュール。

mermaid-parser-py の ganttDb はJSメソッドのみのためJSON.stringifyでスキップされ
graph_data が空になる。そのためMermaidテキストを直接解析するカスタムパーサーを実装する。

サポート構文:
- title / dateFormat / excludes / axisFormat (無視) / tickInterval (無視) / todayMarker (無視)
- section
- タスク行: タスク名 :[タグ...,] [id,] start, end/duration
    - タグ: done / active / crit / milestone
    - 開始: 日付文字列(YYYY-MM-DD) / after taskId [taskId...]
    - 終了: 日付文字列(YYYY-MM-DD) / 期間(Nd/Nw/Nh) / until taskId
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from typing import Optional


@dataclass
class GanttTask:
    """ガントチャートの1タスクを表すデータクラス。"""

    title: str
    task_id: str
    section: str
    start_date: date
    end_date: date
    is_done: bool = False
    is_active: bool = False
    is_crit: bool = False
    is_milestone: bool = False


@dataclass
class GanttSection:
    """ガントチャートの1セクションを表すデータクラス。"""

    name: str
    tasks: list[GanttTask] = field(default_factory=list)


@dataclass
class GanttChart:
    """ガントチャート全体を表すデータクラス。"""

    title: str
    date_format: str
    excludes: list[str]
    sections: list[GanttSection]

    def all_tasks(self) -> list[GanttTask]:
        """全セクションのタスクをフラットなリストで返す。"""
        result: list[GanttTask] = []
        for section in self.sections:
            result.extend(section.tasks)
        return result


# タグ文字列セット（小文字で比較する）
_VALID_TAGS: frozenset[str] = frozenset({"done", "active", "crit", "milestone"})

# YYYY-MM-DD 形式の日付パターン
_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")

# 期間パターン: Nd / Nw / Nh（例: 7d, 2w, 24h）
_DURATION_RE = re.compile(r"^(\d+)(d|w|h)$", re.IGNORECASE)

# after参照パターン: "after taskId [taskId...]"
_AFTER_RE = re.compile(r"^after\s+(.+)$", re.IGNORECASE)

# until参照パターン: "until taskId"
_UNTIL_RE = re.compile(r"^until\s+(\S+)$", re.IGNORECASE)


def _parse_duration(s: str) -> timedelta:
    """
    期間文字列を timedelta に変換する。

    Parameters
    ----------
    s : str
        期間文字列（例: "7d", "2w", "24h"）。

    Returns
    -------
    timedelta
        期間に対応する timedelta オブジェクト。

    Raises
    ------
    ValueError
        期間文字列の形式が不正な場合。
    """
    m = _DURATION_RE.match(s.strip())
    if not m:
        raise ValueError(f"不正な期間文字列: {s!r}")
    n = int(m.group(1))
    unit = m.group(2).lower()
    if unit == "d":
        return timedelta(days=n)
    if unit == "w":
        return timedelta(weeks=n)
    # unit == "h": 時間を日数に切り上げ変換する
    return timedelta(days=max(1, (n + 23) // 24))


def _parse_date_str(s: str) -> date:
    """
    YYYY-MM-DD 形式の日付文字列を date オブジェクトに変換する。

    Parameters
    ----------
    s : str
        日付文字列。

    Returns
    -------
    date
        日付オブジェクト。

    Raises
    ------
    ValueError
        日付文字列の形式が不正な場合。
    """
    return datetime.strptime(s.strip(), "%Y-%m-%d").date()


def _resolve_start(
    value: str,
    task_end_map: dict[str, date],
) -> date:
    """
    start の値文字列を date に解決する。

    日付文字列か "after taskId [...]" 参照のみ受け付ける。

    Parameters
    ----------
    value : str
        解析対象の値文字列。
    task_end_map : dict[str, date]
        タスクIDから終了日への既存マッピング（参照解決用）。

    Returns
    -------
    date
        解決された日付。

    Raises
    ------
    ValueError
        解決できない場合。
    """
    v = value.strip()

    # 日付文字列
    if _DATE_RE.match(v):
        return _parse_date_str(v)

    # after taskId [taskId...] 参照: 最も遅い終了日を開始日として使用する
    m = _AFTER_RE.match(v)
    if m:
        ref_ids = m.group(1).split()
        resolved: list[date] = []
        for rid in ref_ids:
            if rid in task_end_map:
                resolved.append(task_end_map[rid])
        if resolved:
            return max(resolved)
        raise ValueError(f"after参照タスクID {ref_ids} が見つかりません")

    raise ValueError(f"start 日付を解決できません: {v!r}")


def _resolve_end(
    value: str,
    task_end_map: dict[str, date],
    start_date: date,
) -> date:
    """
    end の値文字列を date に解決する。

    日付文字列 / "until taskId" 参照 / "after taskId" 参照 / 期間文字列を受け付ける。

    Parameters
    ----------
    value : str
        解析対象の値文字列。
    task_end_map : dict[str, date]
        タスクIDから終了日への既存マッピング（参照解決用）。
    start_date : date
        期間文字列解決時の基準日。

    Returns
    -------
    date
        解決された日付。

    Raises
    ------
    ValueError
        解決できない場合。
    """
    v = value.strip()

    # 日付文字列
    if _DATE_RE.match(v):
        return _parse_date_str(v)

    # until taskId 参照
    m_until = _UNTIL_RE.match(v)
    if m_until:
        rid = m_until.group(1)
        if rid in task_end_map:
            return task_end_map[rid]
        raise ValueError(f"until参照タスクID '{rid}' が見つかりません")

    # after taskId 参照（end にも許可）
    m_after = _AFTER_RE.match(v)
    if m_after:
        ref_ids = m_after.group(1).split()
        resolved: list[date] = []
        for rid in ref_ids:
            if rid in task_end_map:
                resolved.append(task_end_map[rid])
        if resolved:
            return max(resolved)
        raise ValueError(f"after参照タスクID {ref_ids} が見つかりません")

    # 期間文字列
    if _DURATION_RE.match(v):
        return start_date + _parse_duration(v)

    raise ValueError(f"end 日付を解決できません: {v!r}")


def _parse_task_line(
    line: str,
    section_name: str,
    task_end_map: dict[str, date],
    date_format: str,
    task_counter: list[int],
) -> Optional[GanttTask]:
    """
    1行のタスク行を解析して GanttTask を返す。

    フォーマット: タスク名 :[タグ...,] [id,] start, end/duration

    Parameters
    ----------
    line : str
        タスク行（先頭/末尾の空白は除去済み）。
    section_name : str
        現在のセクション名。
    task_end_map : dict[str, date]
        タスクIDから終了日への既存マッピング（参照解決用・更新もする）。
    date_format : str
        dateFormat ディレクティブの値（現在はYYYY-MM-DDのみサポート）。
    task_counter : list[int]
        自動ID生成用カウンター（長さ1のリスト）。

    Returns
    -------
    GanttTask | None
        解析成功時はGanttTaskオブジェクト、失敗時はNone。
    """
    if ":" not in line:
        return None

    colon_idx = line.index(":")
    task_title = line[:colon_idx].strip()
    rest = line[colon_idx + 1:].strip()

    if not task_title:
        return None

    # カンマで分割してパーツを順次処理する
    parts = [p.strip() for p in rest.split(",")]

    is_done = False
    is_active = False
    is_crit = False
    is_milestone = False
    task_id: Optional[str] = None
    start_str: Optional[str] = None
    end_str: Optional[str] = None

    i = 0
    while i < len(parts):
        part = parts[i]

        # 空文字列はスキップする
        if not part:
            i += 1
            continue

        lower_part = part.lower()

        # タグの検出: done/active/crit/milestone
        if lower_part in _VALID_TAGS:
            if lower_part == "done":
                is_done = True
            elif lower_part == "active":
                is_active = True
            elif lower_part == "crit":
                is_crit = True
            elif lower_part == "milestone":
                is_milestone = True
            i += 1
            continue

        # 日付文字列の検出 → start に設定し、次パーツを end として取得する
        if _DATE_RE.match(part):
            start_str = part
            if i + 1 < len(parts):
                end_str = parts[i + 1]
            i += 2
            break

        # after参照の検出 → start に設定し、次パーツを end として取得する
        if _AFTER_RE.match(part):
            start_str = part
            if i + 1 < len(parts):
                end_str = parts[i + 1]
            i += 2
            break

        # タグでも日付でもafterでもない → タスクIDとして扱う
        if task_id is None:
            task_id = part
        i += 1

    # start が見つからない場合は解析失敗とする
    if start_str is None:
        return None

    # タスクIDが未設定の場合は自動生成する
    if task_id is None:
        task_id = f"task{task_counter[0]}"
    task_counter[0] += 1

    # start_date を解決する
    try:
        start_date = _resolve_start(start_str, task_end_map)
    except ValueError:
        return None

    # end_date を解決する
    if end_str is None:
        end_date = start_date + timedelta(days=1)
    else:
        try:
            end_date = _resolve_end(end_str, task_end_map, start_date)
        except ValueError:
            end_date = start_date + timedelta(days=1)

    # end が start 以前の場合は 1日後に設定する
    if end_date <= start_date:
        end_date = start_date + timedelta(days=1)

    # マイルストーンは期間を1日に正規化する
    if is_milestone:
        end_date = start_date + timedelta(days=1)

    # task_end_map に登録する（後続タスクの after参照解決用）
    task_end_map[task_id] = end_date

    return GanttTask(
        title=task_title,
        task_id=task_id,
        section=section_name,
        start_date=start_date,
        end_date=end_date,
        is_done=is_done,
        is_active=is_active,
        is_crit=is_crit,
        is_milestone=is_milestone,
    )


def parse_gantt(text: str) -> GanttChart:
    """
    Mermaidガントチャートテキストを解析して GanttChart オブジェクトを返す。

    Parameters
    ----------
    text : str
        Mermaidガントチャートのテキスト（gantt から始まる）。

    Returns
    -------
    GanttChart
        解析結果のガントチャートオブジェクト。
    """
    lines = text.splitlines()

    chart_title = ""
    date_format = "YYYY-MM-DD"
    excludes: list[str] = []
    sections: list[GanttSection] = []
    current_section: Optional[GanttSection] = None
    task_end_map: dict[str, date] = {}   # タスクIDから終了日のマッピング
    task_counter = [0]                   # 自動ID生成用カウンター（リストで可変参照）

    for raw_line in lines:
        line = raw_line.strip()

        # 空行・コメント行・gantt宣言行はスキップする
        if not line or line.startswith("%%") or line.startswith("//"):
            continue
        if line.lower() == "gantt":
            continue

        lower_line = line.lower()

        # ディレクティブ行の処理
        if lower_line.startswith("title "):
            chart_title = line[6:].strip()
            continue
        if lower_line.startswith("dateformat "):
            date_format = line[11:].strip()
            continue
        if lower_line.startswith("excludes "):
            excludes_str = line[9:].strip()
            excludes = [e.strip() for e in excludes_str.split(",") if e.strip()]
            continue

        # 表示用ディレクティブは描画に影響しないため無視する
        if (
            lower_line.startswith("axisformat ")
            or lower_line.startswith("tickinterval ")
            or lower_line.startswith("todaymarker ")
            or lower_line.startswith("weekday ")
        ):
            continue

        # section ディレクティブ
        if lower_line.startswith("section "):
            section_name = line[8:].strip()
            current_section = GanttSection(name=section_name, tasks=[])
            sections.append(current_section)
            continue

        # コロンを含む行 → タスク行の可能性
        if ":" in line:
            if current_section is None:
                # セクションが未定義の場合はデフォルトセクションを作成する
                current_section = GanttSection(name="", tasks=[])
                sections.append(current_section)
            task = _parse_task_line(
                line,
                current_section.name,
                task_end_map,
                date_format,
                task_counter,
            )
            if task is not None:
                current_section.tasks.append(task)

    return GanttChart(
        title=chart_title,
        date_format=date_format,
        excludes=excludes,
        sections=sections,
    )
