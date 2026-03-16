"""
Mermaid Gitグラフ（gitGraph）カスタムパーサーモジュール。

mermaid-parser-py は gitGraph の graph_data を JavaScript メソッド経由で保持するため、
JSON.stringify で関数がすべて除去され、commitType 定数しか返らない。
本モジュールはその代替として Python で直接行ベース状態機械パーサーを実装する。

サポートする構文（Mermaid 公式仕様 gitGraph 準拠）:

    gitGraph [LR:|TB:|BT:]

    commit
    commit id: "..." [msg: "..."] [type: NORMAL|REVERSE|HIGHLIGHT] [tag: "..."]
    branch <name> [order: <n>]
    checkout <name>
    switch <name>
    merge <name> [id: "..."] [type: ...] [tag: "..."]
    cherry-pick id: "..." [parent: "..."]
    %% コメント行

config オプション（YAML front-matter での設定、パーサーが解釈する）:
    mainBranchName: <string>   デフォルトブランチ名（デフォルト "main"）
    mainBranchOrder: <int>     デフォルトブランチの表示順序（デフォルト 0）

コミットタイプ:
    NORMAL      : 通常コミット（塗りつぶし円で描画）
    REVERSE     : リバートコミット（✕付き円で描画）
    HIGHLIGHT   : ハイライトコミット（塗りつぶし矩形で描画）
    MERGE       : マージコミット（二重円で描画）- merge コマンドで自動設定
    CHERRY_PICK : チェリーピックコミット（紫円で描画）- cherry-pick コマンドで自動設定
"""

from __future__ import annotations

import re
import uuid
from dataclasses import dataclass, field


# ---------------------------------------------------------------------------
# データクラス定義
# ---------------------------------------------------------------------------

@dataclass
class GitCommit:
    """
    Gitコミットを表すデータクラス。

    Attributes
    ----------
    commit_id : str
        コミットの一意ID。id 属性が指定されない場合は自動生成する。
    commit_type : str
        コミットタイプ文字列。NORMAL / REVERSE / HIGHLIGHT / MERGE / CHERRY_PICK のいずれか。
    tag : str | None
        タグ文字列（例: "v1.0.0"）。指定がない場合は None。
    msg : str | None
        コミットメッセージ。表示ラベルは commit_id を優先し、なければ msg を使う。
    branch : str
        このコミットが属するブランチ名。
    parents : list[str]
        親コミットIDのリスト。通常コミットは1個、マージは2個。
    cherry_from : str | None
        cherry-pick の参照元コミットID。CHERRY_PICK タイプのみ設定される。
    """

    commit_id: str
    commit_type: str = "NORMAL"     # NORMAL / REVERSE / HIGHLIGHT / MERGE / CHERRY_PICK
    tag: str | None = None
    msg: str | None = None
    branch: str = "main"
    parents: list[str] = field(default_factory=list)
    cherry_from: str | None = None


@dataclass
class GitBranch:
    """
    Gitブランチを表すデータクラス。

    Attributes
    ----------
    name : str
        ブランチ名。
    order : int | None
        表示順序の指定。None の場合は定義順で表示する。
    """

    name: str
    order: int | None = None


@dataclass
class GitGraph:
    """
    Mermaid gitGraph 全体を表すデータクラス。

    Attributes
    ----------
    direction : str
        グラフの向き。LR（左→右、デフォルト）/ TB（上→下）/ BT（下→上）。
    branches : list[GitBranch]
        定義順のブランチリスト。
    commits : list[GitCommit]
        アクション発生順のコミットリスト。
    main_branch : str
        デフォルトブランチ名（通常 "main"）。
    """

    direction: str = "LR"
    branches: list[GitBranch] = field(default_factory=list)
    commits: list[GitCommit] = field(default_factory=list)
    main_branch: str = "main"


# ---------------------------------------------------------------------------
# 正規表現パターン
# ---------------------------------------------------------------------------

# 1行目: gitGraph [LR:|TB:|BT:]
_RE_HEADER = re.compile(
    r"^\s*gitGraph\s*(?P<dir>LR:|TB:|BT:)?\s*(?::|$)",
    re.IGNORECASE,
)

# commit 行から各属性を抽出する（id/msg/type/tag の任意組み合わせ）
_RE_COMMIT_ID   = re.compile(r'\bid\s*:\s*"(?P<v>[^"]*)"|\bid\s*:\s*\'(?P<v2>[^\']*)\'' )
_RE_COMMIT_MSG  = re.compile(r'\bmsg\s*:\s*"(?P<v>[^"]*)"|\bmsg\s*:\s*\'(?P<v2>[^\']*)\'' )
_RE_COMMIT_TYPE = re.compile(r'\btype\s*:\s*(?P<v>NORMAL|REVERSE|HIGHLIGHT)\b', re.IGNORECASE)
_RE_COMMIT_TAG  = re.compile(r'\btag\s*:\s*"(?P<v>[^"]*)"|\btag\s*:\s*\'(?P<v2>[^\']*)\'' )

# branch 行: "branch <name> [order: <n>]"
_RE_BRANCH = re.compile(
    r'^\s*branch\s+(?P<name>"[^"]+"|\'[^\']+\'|[\w/.\-]+)\s*(?:order\s*:\s*(?P<order>\d+))?\s*$',
    re.IGNORECASE,
)

# checkout/switch 行: "checkout <name>" または "switch <name>"
_RE_CHECKOUT = re.compile(
    r'^\s*(?:checkout|switch)\s+(?P<name>"[^"]+"|\'[^\']+\'|[\w/.\-]+)\s*$',
    re.IGNORECASE,
)

# merge 行: "merge <name> [id/type/tag 属性...]"
_RE_MERGE = re.compile(
    r'^\s*merge\s+(?P<name>"[^"]+"|\'[^\']+\'|[\w/.\-]+)(?P<rest>.*)$',
    re.IGNORECASE,
)

# cherry-pick 行: "cherry-pick id: "..." [parent: "..."]"
_RE_CHERRY_PICK = re.compile(
    r'^\s*cherry-pick\s+(?P<rest>.+)$',
    re.IGNORECASE,
)
_RE_CHERRY_PARENT = re.compile(
    r'\bparent\s*:\s*"(?P<v>[^"]*)"|\bparent\s*:\s*\'(?P<v2>[^\']*)\''
)

# mainBranchName の抽出（YAML front-matter の config.gitGraph セクション）
_RE_MAIN_BRANCH_NAME = re.compile(
    r'mainBranchName\s*:\s*["\']?(?P<name>[\w/.\-]+)["\']?',
    re.IGNORECASE,
)

# コメント行
_RE_COMMENT = re.compile(r'^\s*%%')


# ---------------------------------------------------------------------------
# ユーティリティ
# ---------------------------------------------------------------------------

def _strip_quotes(s: str) -> str:
    """引用符（ダブル・シングル）で囲まれた文字列から引用符を除去する。"""
    s = s.strip()
    if (s.startswith('"') and s.endswith('"')) or (s.startswith("'") and s.endswith("'")):
        return s[1:-1]
    return s


def _extract_quoted(pattern: re.Pattern[str], text: str) -> str | None:
    """
    正規表現パターンで文字列を検索し、グループ `v` または `v2` の値を返す。

    Parameters
    ----------
    pattern : re.Pattern[str]
        ダブルクォート用グループ `v` とシングルクォート用グループ `v2` を持つパターン。
    text : str
        検索対象テキスト。

    Returns
    -------
    str | None
        マッチしたグループの値。マッチなしの場合は None。
    """
    m = pattern.search(text)
    if not m:
        return None
    return m.group("v") if m.group("v") is not None else m.group("v2")


def _auto_commit_id(branch: str, seq: int) -> str:
    """
    id 属性が指定されていない場合の自動生成コミットID。

    Parameters
    ----------
    branch : str
        ブランチ名（IDの可読性向上に使用）。
    seq : int
        全コミット通し番号。

    Returns
    -------
    str
        "seq-branch-..." 形式の自動ID。
    """
    short = uuid.uuid4().hex[:7]
    return f"{seq}-{short}"


# ---------------------------------------------------------------------------
# パーサー本体
# ---------------------------------------------------------------------------

def parse_gitgraph(text: str, yaml_frontmatter: str = "") -> GitGraph:
    """
    Mermaid gitGraph テキストを解析して GitGraph データクラスに変換する。

    Parameters
    ----------
    text : str
        gitGraph の Mermaid テキスト（「gitGraph」から始まる行を含む）。
    yaml_frontmatter : str
        YAML front-matter テキスト（mainBranchName の抽出に使用）。
        省略時は空文字列。

    Returns
    -------
    GitGraph
        解析結果を格納した GitGraph データクラス。

    Raises
    ------
    ValueError
        パース不能なテキストが渡された場合（現在は例外を発生させず空グラフを返す）。
    """
    # --- mainBranchName の抽出 ---
    main_branch = "main"
    mb_match = _RE_MAIN_BRANCH_NAME.search(yaml_frontmatter + "\n" + text)
    if mb_match:
        main_branch = mb_match.group("name")

    lines = text.splitlines()

    # --- 1行目: ヘッダー（direction）解析 ---
    direction = "LR"
    header_idx = 0
    for i, line in enumerate(lines):
        stripped = line.strip()
        if _RE_COMMENT.match(stripped):
            continue
        m = _RE_HEADER.match(stripped)
        if m:
            dir_token = (m.group("dir") or "LR:").upper().rstrip(":")
            if dir_token in ("LR", "TB", "BT"):
                direction = dir_token
            header_idx = i
            break

    graph = GitGraph(direction=direction, main_branch=main_branch)

    # --- デフォルトブランチ登録 ---
    _branch_map: dict[str, GitBranch] = {}
    default_branch = GitBranch(name=main_branch, order=0)
    graph.branches.append(default_branch)
    _branch_map[main_branch] = default_branch

    # --- 状態変数 ---
    current_branch: str = main_branch              # カレントブランチ名
    branch_heads: dict[str, str | None] = {main_branch: None}  # ブランチ先頭コミットID
    commit_seq: int = 0                             # 全コミット通し番号

    # --- 行ごとの解析 ---
    for line in lines[header_idx + 1:]:
        stripped = line.strip()

        # 空行・コメントをスキップ
        if not stripped or _RE_COMMENT.match(stripped):
            continue

        lower = stripped.lower()

        # commit
        if lower.startswith("commit"):
            rest = stripped[6:]

            # 各属性を正規表現で抽出する
            cid = _extract_quoted(_RE_COMMIT_ID, rest)
            if cid is None:
                cid = _auto_commit_id(current_branch, commit_seq)

            msg = _extract_quoted(_RE_COMMIT_MSG, rest)

            tm = _RE_COMMIT_TYPE.search(rest)
            ctype = tm.group("v").upper() if tm else "NORMAL"

            tag = _extract_quoted(_RE_COMMIT_TAG, rest)

            # 親コミットの確定（カレントブランチの先頭）
            parent_id = branch_heads.get(current_branch)
            parents = [parent_id] if parent_id else []

            commit = GitCommit(
                commit_id=cid,
                commit_type=ctype,
                tag=tag,
                msg=msg,
                branch=current_branch,
                parents=parents,
            )
            graph.commits.append(commit)
            branch_heads[current_branch] = cid
            commit_seq += 1
            continue

        # branch
        mb = _RE_BRANCH.match(stripped)
        if mb:
            bname = _strip_quotes(mb.group("name"))
            border = int(mb.group("order")) if mb.group("order") else None
            if bname not in _branch_map:
                new_branch = GitBranch(name=bname, order=border)
                graph.branches.append(new_branch)
                _branch_map[bname] = new_branch
                # 新ブランチはカレントブランチの先頭から開始する
                branch_heads[bname] = branch_heads.get(current_branch)
            current_branch = bname
            continue

        # checkout / switch
        mc = _RE_CHECKOUT.match(stripped)
        if mc:
            bname = _strip_quotes(mc.group("name"))
            if bname in _branch_map:
                current_branch = bname
            # 不明ブランチへの checkout は無視する
            continue

        # merge
        mm = _RE_MERGE.match(stripped)
        if mm:
            src_branch = _strip_quotes(mm.group("name"))
            rest = mm.group("rest") or ""

            # merge 属性の抽出
            mid = _extract_quoted(_RE_COMMIT_ID, rest)
            if mid is None:
                mid = _auto_commit_id(current_branch, commit_seq)

            tm = _RE_COMMIT_TYPE.search(rest)
            mtype = tm.group("v").upper() if tm else "MERGE"

            mtag = _extract_quoted(_RE_COMMIT_TAG, rest)

            # 親: [カレントブランチ先頭, マージ元ブランチ先頭]
            src_head = branch_heads.get(src_branch)
            cur_head = branch_heads.get(current_branch)
            parents: list[str] = []
            if cur_head:
                parents.append(cur_head)
            if src_head:
                parents.append(src_head)

            commit = GitCommit(
                commit_id=mid,
                commit_type=mtype if mtype != "NORMAL" else "MERGE",
                tag=mtag,
                branch=current_branch,
                parents=parents,
            )
            graph.commits.append(commit)
            branch_heads[current_branch] = mid
            commit_seq += 1
            continue

        # cherry-pick
        mcp = _RE_CHERRY_PICK.match(stripped)
        if mcp:
            rest = mcp.group("rest")
            src_id = _extract_quoted(_RE_COMMIT_ID, rest)
            parent_id_cp = _extract_quoted(_RE_CHERRY_PARENT, rest)

            if src_id is None:
                # cherry-pick id が指定されていない場合はスキップする
                continue

            # cherry-pick コミットの ID は自動生成する
            cp_id = _auto_commit_id(current_branch, commit_seq)

            # 親はカレントブランチ先頭
            cur_head = branch_heads.get(current_branch)
            parents = [cur_head] if cur_head else []

            commit = GitCommit(
                commit_id=cp_id,
                commit_type="CHERRY_PICK",
                tag=src_id,            # タグとしてコピー元IDを表示する（Mermaid準拠）
                branch=current_branch,
                parents=parents,
                cherry_from=src_id,
            )
            graph.commits.append(commit)
            branch_heads[current_branch] = cp_id
            commit_seq += 1
            continue

        # 認識できない行は無視する

    return graph
