"""
Mermaidガントチャートカスタムパーサーモジュール。

mermaid-parser-py の ganttDb はJSメソッドのみのためJSON.stringifyでスキップされ
graph_data が空になる。そのためMermaidテキストを直接解析するカスタムパーサーを実装する。

サポート構文:
- title / dateFormat / excludes / axisFormat (無視) / tickInterval (無視) / todayMarker (無視)
- section
- タスク行: タスク名 :[タグ...,] [id,] start, end/duration
    - タグ: done / active / crit / milestone
    - 開始: 日付文字列 / after taskId [taskId...]
    - 終了: 日付文字列 / 期間(Nd/Nw/Nh/NM/Ny) / until taskId [taskId...]

サポートする dateFormat 値:
    YYYY-MM-DD, YYYY-MM-DD HH:mm, YYYY-MM-DD HH:mm:ss,
    YYYY-MM-DDTHH:mm, YYYY-MM-DDTHH:mm:ss,
    YYYY-MM, YYYY, YYYYMMDD,
    MM/DD/YYYY, DD/MM/YYYY, M/D/YYYY,
    X (Unixタイムスタンプ秒), x (Unixタイムスタンプミリ秒)
    ※ 上記以外のフォーマットは YYYY-MM-DD にフォールバックする

サポートする期間（duration）単位:
    d/D=日, w/W=週, h/H=時間(切り上げ), M=月(30日近似), y/Y=年(365日近似),
    m=分(最小1日), s=秒(最小1日), ms=ミリ秒(最小1日)
"""

from __future__ import annotations

import math
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

# Mermaid dateFormat 文字列 → (正規表現パターン文字列, strptime フォーマット文字列) のマッピング
# 日が省略されたフォーマット（YYYY-MM / YYYY 等）は月初/年初 (day=1) を補完する
# 時刻を含むフォーマットは datetime.strptime().date() で日付のみを返す
# "UNIX_SECONDS" / "UNIX_MILLISECONDS" は Unix タイムスタンプ用のセンチネル値
_DATEFORMAT_MAP: dict[str, tuple[str, str]] = {
    # --- 日付のみのフォーマット ---
    "YYYY-MM-DD":          (r"^\d{4}-\d{2}-\d{2}$",                      "%Y-%m-%d"),
    "YYYY-MM":             (r"^\d{4}-\d{2}$",                             "%Y-%m"),
    "YYYY":                (r"^\d{4}$",                                   "%Y"),
    "YYYYMMDD":            (r"^\d{8}$",                                   "%Y%m%d"),
    "MM/DD/YYYY":          (r"^\d{2}/\d{2}/\d{4}$",                      "%m/%d/%Y"),
    "DD/MM/YYYY":          (r"^\d{2}/\d{2}/\d{4}$",                      "%d/%m/%Y"),
    "M/D/YYYY":            (r"^\d{1,2}/\d{1,2}/\d{4}$",                  "%m/%d/%Y"),
    # --- 時刻付きフォーマット（dayjs 形式と対応; 時刻部は切り捨てて date のみ使用）---
    "YYYY-MM-DD HH:mm":    (r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$",         "%Y-%m-%d %H:%M"),
    "YYYY-MM-DD HH:mm:ss": (r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$",  "%Y-%m-%d %H:%M:%S"),
    "YYYY-MM-DDTHH:mm:ss": (r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$",  "%Y-%m-%dT%H:%M:%S"),
    "YYYY-MM-DDTHH:mm":    (r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$",         "%Y-%m-%dT%H:%M"),
    # --- Unix タイムスタンプ（dayjs: X=秒, x=ミリ秒）---
    # 両フォーマットとも数字のみのパターンだが、strptime_fmt のセンチネルで区別する。
    # パターン自体が同じでも、dateFormat ディレクティブで形式が明示されるため自動判別不要。
    "X":                   (r"^\d+$",                                     "UNIX_SECONDS"),
    "x":                   (r"^\d+$",                                     "UNIX_MILLISECONDS"),
}


def _get_date_helpers(date_format: str) -> tuple[re.Pattern, str]:
    """
    dateFormat 文字列から (日付判定正規表現, strptime フォーマット) を返す。

    Parameters
    ----------
    date_format : str
        Mermaid の dateFormat ディレクティブ値（例: "YYYY-MM-DD", "YYYY-MM"）。

    Returns
    -------
    tuple[re.Pattern, str]
        (コンパイル済み正規表現, strptime フォーマット文字列) のタプル。
        未知のフォーマットは YYYY-MM-DD にフォールバックする。
    """
    pat_str, strptime_fmt = _DATEFORMAT_MAP.get(
        date_format, _DATEFORMAT_MAP["YYYY-MM-DD"]
    )
    return re.compile(pat_str), strptime_fmt


# 期間パターン: mermaid.js ganttDb.js の parseDuration に対応する単位
# ms（ミリ秒）は M（月）より先に評価する必要があるため先頭に配置する
# M=月(大文字), d/D=日, h/H=時間, m=分, s/S=秒, w/W=週, y/Y=年, ms=ミリ秒
_DURATION_RE = re.compile(r"^(\d+(?:\.\d+)?)(ms|M|[dhmswyYDWH])$")

# after参照パターン: "after taskId [taskId...]"
_AFTER_RE = re.compile(r"^after\s+(.+)$", re.IGNORECASE)

# until参照パターン: "until taskId [taskId...]"（mermaid.js 準拠で複数ID対応）
_UNTIL_RE = re.compile(r"^until\s+(.+)$", re.IGNORECASE)


def _parse_duration(s: str) -> timedelta:
    """
    期間文字列を timedelta に変換する。

    mermaid.js ganttDb.js の parseDuration に対応する単位:
    - ms (ミリ秒): 最小1日
    - M  (月, 大文字): 30日近似
    - d / D (日): そのまま
    - h / H (時間): 24時間単位で切り上げて日数に変換
    - m  (分): 最小1日
    - s / S (秒): 最小1日
    - w / W (週): そのまま
    - y / Y (年): 365日近似

    Parameters
    ----------
    s : str
        期間文字列（例: "7d", "2w", "24h", "3M", "1y"）。

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
    n = float(m.group(1))
    unit = m.group(2)

    # 大文字 M = 月: 30日近似（小文字 m = 分 と区別するため厳密比較）
    # 小数値は切り上げ（1.2M → 36日）
    if unit == "M":
        return timedelta(days=math.ceil(n * 30))

    unit_lower = unit.lower()

    if unit_lower == "y":
        # 年: 365日近似、小数値は切り上げ
        return timedelta(days=math.ceil(n * 365))
    if unit_lower == "w":
        return timedelta(weeks=math.ceil(n))
    if unit_lower == "d":
        return timedelta(days=math.ceil(n))
    if unit_lower == "h":
        # 時間を日数に切り上げ変換する（浮動小数点対応: math.ceil を使用）
        return timedelta(days=max(1, math.ceil(n / 24)))
    # m (分), s (秒), ms (ミリ秒): 日未満の粒度はガントチャートに不要のため最小1日
    return timedelta(days=1)


def _parse_date_str(s: str, strptime_fmt: str = "%Y-%m-%d") -> date:
    """
    strptime_fmt に従って日付文字列を date オブジェクトに変換する。

    - YYYY-MM など日が省略されたフォーマットは月初 (day=1) を補完する。
    - 時刻を含むフォーマット（HH:mm 等）は時刻を切り捨てて date を返す。
    - "UNIX_SECONDS" / "UNIX_MILLISECONDS" センチネル値の場合は Unix タイムスタンプとして解析する。

    Parameters
    ----------
    s : str
        日付文字列。
    strptime_fmt : str
        Python の strptime フォーマット文字列（例: "%Y-%m-%d", "%Y-%m-%d %H:%M"）。
        または "UNIX_SECONDS" / "UNIX_MILLISECONDS" のセンチネル文字列。

    Returns
    -------
    date
        日付オブジェクト。

    Raises
    ------
    ValueError
        日付文字列の形式が不正な場合。
    """
    s = s.strip()
    if strptime_fmt == "UNIX_SECONDS":
        return date.fromtimestamp(int(s))
    if strptime_fmt == "UNIX_MILLISECONDS":
        return date.fromtimestamp(int(s) // 1000)
    # 時刻を含むフォーマットも .date() で日付のみを取得する
    return datetime.strptime(s, strptime_fmt).date()


def _resolve_start(
    value: str,
    task_end_map: dict[str, date],
    date_format: str = "YYYY-MM-DD",
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
    date_format : str
        Mermaid の dateFormat ディレクティブ値（例: "YYYY-MM-DD", "YYYY-MM"）。

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
    date_re, strptime_fmt = _get_date_helpers(date_format)

    # 日付文字列
    if date_re.match(v):
        return _parse_date_str(v, strptime_fmt)

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
    date_format: str = "YYYY-MM-DD",
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
    date_format : str
        Mermaid の dateFormat ディレクティブ値（例: "YYYY-MM-DD", "YYYY-MM"）。

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
    date_re, strptime_fmt = _get_date_helpers(date_format)

    # 日付文字列
    if date_re.match(v):
        return _parse_date_str(v, strptime_fmt)

    # until taskId [taskId...] 参照
    # mermaid.js 準拠: 参照タスクの開始日の最小値を終了日として使用する
    # 本実装では開始日マップが存在しないため終了日の最小値で代替する
    m_until = _UNTIL_RE.match(v)
    if m_until:
        ref_ids = m_until.group(1).split()
        resolved_ends: list[date] = []
        for rid in ref_ids:
            if rid in task_end_map:
                resolved_ends.append(task_end_map[rid])
        if resolved_ends:
            return min(resolved_ends)
        raise ValueError(f"until参照タスクID {ref_ids} が見つかりません")

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
        _date_re, _ = _get_date_helpers(date_format)
        if _date_re.match(part):
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
        start_date = _resolve_start(start_str, task_end_map, date_format)
    except ValueError:
        return None

    # end_date を解決する
    if end_str is None:
        end_date = start_date + timedelta(days=1)
    else:
        try:
            end_date = _resolve_end(end_str, task_end_map, start_date, date_format)
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
