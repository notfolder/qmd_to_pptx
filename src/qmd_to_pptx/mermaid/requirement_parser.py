"""
Mermaid 要件図（requirementDiagram）カスタムパーサーモジュール。

mermaid-parser-py の requirementDiagram パーサーは Jison 生成のレキサーが
`\w` (ASCII 限定) でトークン化するため、名前・テキストに日本語等を使うと
SpiderMonkeyError が発生する。本モジュールはその代替として Python で直接
正規表現ベースの状態機械パーサーを実装する。

サポートする構文（Mermaid 公式仕様 requirementDiagram 準拠）:

    requirementDiagram

    [direction TB|BT|LR|RL]

    <requirementType> [nodeName][:::class] {
        id: <value>
        text: <value>
        risk: low|medium|high
        verifymethod: analysis|inspection|test|demonstration
    }

    element [nodeName][:::class] {
        type: <value>
        docref: <value>
    }

    <src> - <relType> -> <dst>    ← 順方向
    <dst> <- <relType> - <src>    ← 逆方向

    classDef <name> fill:#rrggbb, stroke:#rrggbb, ...
    class <name1>[,<name2>...] <className>
    style <name> fill:#rrggbb, stroke:#rrggbb, ...

    <name>:::className             ← ノード単体クラス適用行

名前・テキストは "..." でクォートすることで Unicode / 日本語・スペースを含めることができる。
クォートなしの場合は `[^:,\r\n\{\<\>\-\=]+` に相当するテキストとして解析する。

テキスト中の Markdown 書式 (**bold** / *italic*) は parse_inline_markdown() で
(text, bold, italic) タプルのリストに変換する。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field


# ---------------------------------------------------------------------------
# データクラス定義
# ---------------------------------------------------------------------------

# 要件タイプ → ステレオタイプ表示名のマッピング
REQUIREMENT_TYPES: dict[str, str] = {
    "requirement": "Requirement",
    "functionalrequirement": "Functional Requirement",
    "interfacerequirement": "Interface Requirement",
    "performancerequirement": "Performance Requirement",
    "physicalrequirement": "Physical Requirement",
    "designconstraint": "Design Constraint",
}

# リスクレベルの正規化
_RISK_NORM: dict[str, str] = {
    "low": "Low",
    "medium": "Medium",
    "high": "High",
}

# 検証方法の正規化
_VERIFY_NORM: dict[str, str] = {
    "analysis": "Analysis",
    "inspection": "Inspection",
    "test": "Test",
    "demonstration": "Demonstration",
}


@dataclass
class NodeStyle:
    """ノードのスタイル情報を保持するデータクラス。"""

    fill: str | None = None         # 塗りつぶし色 (#rrggbb)
    stroke: str | None = None       # 枠線色 (#rrggbb)
    stroke_width: str | None = None # 枠線幅文字列 (例: "2px")
    color: str | None = None        # テキスト色 (#rrggbb)
    font_weight: str | None = None  # フォントウェイト (例: "bold")


@dataclass
class RequirementNode:
    """
    要件ノードを表すデータクラス。

    Attributes
    ----------
    name : str
        ノード識別名（定義ブロックの名前）。
    req_type : str
        要件タイプキーワード（小文字正規化済み）。
        requirement / functionalrequirement / interfacerequirement /
        performancerequirement / physicalrequirement / designconstraint
    stereotype : str
        ステレオタイプ表示文字列（例: "Functional Requirement"）。
    req_id : str
        要件 ID（id: フィールド値）。
    text : str
        要件テキスト（text: フィールド値）。
    risk : str
        リスクレベル（正規化済み: Low / Medium / High / ""）。
    verify_method : str
        検証方法（正規化済み: Analysis / Inspection / Test / Demonstration / ""）。
    classes : list[str]
        適用されたクラス名のリスト。
    style : NodeStyle
        直接スタイル指定。
    """

    name: str
    req_type: str = "requirement"
    stereotype: str = "Requirement"
    req_id: str = ""
    text: str = ""
    risk: str = ""
    verify_method: str = ""
    classes: list[str] = field(default_factory=list)
    style: NodeStyle = field(default_factory=NodeStyle)


@dataclass
class ElementNode:
    """
    エレメントノードを表すデータクラス。

    Attributes
    ----------
    name : str
        ノード識別名。
    elem_type : str
        エレメントタイプ（type: フィールド値）。
    docref : str
        ドキュメント参照（docref: フィールド値）。
    classes : list[str]
        適用されたクラス名のリスト。
    style : NodeStyle
        直接スタイル指定。
    """

    name: str
    elem_type: str = ""
    docref: str = ""
    classes: list[str] = field(default_factory=list)
    style: NodeStyle = field(default_factory=NodeStyle)


@dataclass
class Relationship:
    """
    ノード間リレーションを表すデータクラス。

    Attributes
    ----------
    src : str
        始点ノード名。
    dst : str
        終点ノード名。
    rel_type : str
        リレーションタイプ（小文字正規化済み）。
        contains / copies / derives / satisfies / verifies / refines / traces
    """

    src: str
    dst: str
    rel_type: str


@dataclass
class RequirementDiagram:
    """
    Mermaid requirementDiagram 全体を表すデータクラス。

    Attributes
    ----------
    direction : str
        レイアウト方向 (TB | BT | LR | RL)。デフォルトは "TB"。
    requirements : dict[str, RequirementNode]
        ノード名 → RequirementNode の辞書（定義順保持）。
    elements : dict[str, ElementNode]
        ノード名 → ElementNode の辞書（定義順保持）。
    relations : list[Relationship]
        リレーションのリスト（定義順）。
    class_defs : dict[str, NodeStyle]
        classDef 名 → NodeStyle の辞書。
    """

    direction: str = "TB"
    requirements: dict[str, RequirementNode] = field(default_factory=dict)
    elements: dict[str, ElementNode] = field(default_factory=dict)
    relations: list[Relationship] = field(default_factory=list)
    class_defs: dict[str, NodeStyle] = field(default_factory=dict)


# ---------------------------------------------------------------------------
# 正規表現パターン
# ---------------------------------------------------------------------------

# ヘッダー行
_RE_HEADER = re.compile(r"^\s*requirementDiagram\s*$", re.IGNORECASE)

# direction 行
_RE_DIRECTION = re.compile(r"^\s*direction\s+(TB|BT|LR|RL)\s*$", re.IGNORECASE)

# 要件タイプキーワード
_REQUIREMENT_TYPE_PATTERN = "|".join(
    re.escape(k) for k in [
        "functionalRequirement", "interfaceRequirement",
        "performanceRequirement", "physicalRequirement",
        "designConstraint", "requirement",
    ]
)

# 要件ブロック開始行: "<type> [:::class] <name>[:::class] {"
# または: "<type> <name>[:::class] {"
_RE_REQ_START = re.compile(
    r"^\s*(?P<rtype>" + _REQUIREMENT_TYPE_PATTERN + r")"
    r"(?:::(?P<cls1>[A-Za-z_][A-Za-z0-9_]*))?"
    r"\s+(?P<name>" + r'"[^"]*"' + r"|[^\s{:::]+)"
    r"(?:::(?P<cls2>[A-Za-z_][A-Za-z0-9_]*))?"
    r"\s*\{",
    re.IGNORECASE,
)

# エレメントブロック開始行: "element [:::class] <name>[:::class] {"
_RE_ELEM_START = re.compile(
    r"^\s*element"
    r"(?:::(?P<cls1>[A-Za-z_][A-Za-z0-9_]*))?"
    r"\s+(?P<name>" + r'"[^"]*"' + r"|[^\s{:::]+)"
    r"(?:::(?P<cls2>[A-Za-z_][A-Za-z0-9_]*))?"
    r"\s*\{",
    re.IGNORECASE,
)

# ブロック終了行
_RE_BLOCK_END = re.compile(r"^\s*\}\s*$")

# フィールド行（ブロック内）: "key: value" または "key: "value""
_RE_FIELD = re.compile(
    r"""^\s*(?P<key>[A-Za-z]+)\s*:\s*(?P<val>"[^"]*"|.+?)\s*$"""
)

# 順方向リレーション: "<src> - <type> -> <dst>"
_RE_REL_FWD = re.compile(
    r"""^\s*(?P<src>"[^"]*"|[^\s\-<>]+)\s*-\s*(?P<rtype>[A-Za-z]+)\s*->\s*(?P<dst>"[^"]*"|[^\s\-<>]+)\s*$"""
)

# 逆方向リレーション: "<dst> <- <type> - <src>"
_RE_REL_REV = re.compile(
    r"""^\s*(?P<dst>"[^"]*"|[^\s\-<>]+)\s*<-\s*(?P<rtype>[A-Za-z]+)\s*-\s*(?P<src>"[^"]*"|[^\s\-<>]+)\s*$"""
)

# classDef 行: "classDef <name> <styles>"
_RE_CLASSDEF = re.compile(
    r"^\s*classDef\s+(?P<name>[A-Za-z_][A-Za-z0-9_]*)\s+(?P<styles>.+)\s*$",
    re.IGNORECASE,
)

# class 適用行: "class <name1>[,<name2>...] <className>"
_RE_CLASS_APPLY = re.compile(
    r"^\s*class\s+(?P<names>[A-Za-z_][A-Za-z0-9_,\s]*)\s+(?P<cls>[A-Za-z_][A-Za-z0-9_]*)\s*$",
    re.IGNORECASE,
)

# style 直接適用行: "style <name> <styles>"
_RE_STYLE_APPLY = re.compile(
    r"^\s*style\s+(?P<name>[A-Za-z_][A-Za-z0-9_\"]*)\s+(?P<styles>.+)\s*$",
    re.IGNORECASE,
)

# ブロック外での :::class 適用行: "<name>:::class"
# ※ Mermaid の :::className 構文は 3 コロン (:::)
_RE_CLASS_SHORTHAND = re.compile(
    r"""^\s*(?P<name>"[^"]*"|[A-Za-z_][A-Za-z0-9_]*):::(?P<cls>[A-Za-z_][A-Za-z0-9_]*)\s*$"""
)

# コメント行
_RE_COMMENT = re.compile(r"^\s*%%")

# アクセシビリティ行
_RE_ACC = re.compile(r"^\s*acc(?:Title|Descr)\b", re.IGNORECASE)

# CSS スタイルプロパティの1エントリ
_RE_STYLE_PROP = re.compile(
    r"(?P<key>fill|stroke-width|stroke|color|font-weight)\s*:\s*(?P<value>[^,;]+)",
    re.IGNORECASE,
)


# ---------------------------------------------------------------------------
# ユーティリティ関数
# ---------------------------------------------------------------------------


def _strip_quotes(s: str) -> str:
    """
    文字列の前後のダブルクォートを除去する。

    Parameters
    ----------
    s : str
        入力文字列。

    Returns
    -------
    str
        クォートを除去した文字列。
    """
    s = s.strip()
    if s.startswith('"') and s.endswith('"') and len(s) >= 2:
        return s[1:-1]
    return s


def _parse_node_style(styles_str: str) -> NodeStyle:
    """
    CSS スタイル文字列（カンマ/セミコロン区切り）を NodeStyle に変換する。

    Parameters
    ----------
    styles_str : str
        スタイル文字列（例: "fill:#ffa, stroke:#000, color: green"）。

    Returns
    -------
    NodeStyle
        解析結果のスタイルデータクラス。
    """
    style = NodeStyle()
    for m in _RE_STYLE_PROP.finditer(styles_str):
        key = m.group("key").lower()
        value = m.group("value").strip().rstrip(";")
        if key == "fill":
            style.fill = value
        elif key == "stroke":
            style.stroke = value
        elif key == "stroke-width":
            style.stroke_width = value
        elif key == "color":
            style.color = value
        elif key == "font-weight":
            style.font_weight = value
    return style


def _merge_styles(base: NodeStyle, override: NodeStyle) -> NodeStyle:
    """
    2つの NodeStyle をマージし、override が優先される NodeStyle を返す。

    Parameters
    ----------
    base : NodeStyle
        ベーススタイル（優先度低）。
    override : NodeStyle
        上書きスタイル（優先度高）。

    Returns
    -------
    NodeStyle
        マージ結果。
    """
    return NodeStyle(
        fill=override.fill if override.fill is not None else base.fill,
        stroke=override.stroke if override.stroke is not None else base.stroke,
        stroke_width=override.stroke_width if override.stroke_width is not None else base.stroke_width,
        color=override.color if override.color is not None else base.color,
        font_weight=override.font_weight if override.font_weight is not None else base.font_weight,
    )


def parse_inline_markdown(text: str) -> list[tuple[str, bool, bool]]:
    """
    テキスト中の Markdown 書式（**bold** / *italic*）を解析する。

    Parameters
    ----------
    text : str
        入力テキスト（ダブルクォート除去済み）。

    Returns
    -------
    list[tuple[str, bool, bool]]
        (テキスト, bold, italic) のタプルリスト。
        通常テキストは (text, False, False)、
        太字は (text, True, False)、
        斜体は (text, False, True)。
    """
    result: list[tuple[str, bool, bool]] = []
    # **bold** → (text, True, False)
    # *italic* → (text, False, True)
    # それ以外 → (text, False, False)
    pattern = re.compile(r"\*\*(.+?)\*\*|\*(.+?)\*|([^*]+)")
    for m in pattern.finditer(text):
        if m.group(1) is not None:
            result.append((m.group(1), True, False))
        elif m.group(2) is not None:
            result.append((m.group(2), False, True))
        elif m.group(3):
            result.append((m.group(3), False, False))
    return result if result else [(text, False, False)]


# ---------------------------------------------------------------------------
# ヘルパー関数
# ---------------------------------------------------------------------------


def _apply_field(
    key: str,
    val: str,
    current_req: "RequirementNode | None",
    current_elem: "ElementNode | None",
) -> None:
    """
    フィールド キー・値をカレントノードに適用する。

    Parameters
    ----------
    key : str
        フィールドキー（小文字正規化済み）。
    val : str
        フィールド値（クォート除去済み）。
    current_req : RequirementNode | None
        カレント要件ノード。
    current_elem : ElementNode | None
        カレントエレメントノード。
    """
    if current_req is not None:
        if key == "id":
            current_req.req_id = val
        elif key == "text":
            current_req.text = val
        elif key == "risk":
            current_req.risk = _RISK_NORM.get(val.lower(), val)
        elif key in ("verifymethod", "verificationmethod"):
            current_req.verify_method = _VERIFY_NORM.get(val.lower(), val)
    elif current_elem is not None:
        if key == "type":
            current_elem.elem_type = val
        elif key == "docref":
            current_elem.docref = val


def _parse_inline_fields(
    inline_text: str,
    current_req: "RequirementNode | None",
    current_elem: "ElementNode | None",
) -> None:
    """
    インラインブロック内のフィールド文字列を解析してノードに適用する。

    セミコロン区切りで各フィールドを分割し、``key: value`` パターンに
    マッチしたものを適用する。

    Parameters
    ----------
    inline_text : str
        ``{`` と ``}`` の間のテキスト。
    current_req : RequirementNode | None
        フィールドを適用する要件ノード。
    current_elem : ElementNode | None
        フィールドを適用するエレメントノード。
    """
    for segment in re.split(r"[;\n]", inline_text):
        m = _RE_FIELD.match(segment.strip())
        if m:
            key = m.group("key").lower().strip()
            val = _strip_quotes(m.group("val").strip())
            _apply_field(key, val, current_req, current_elem)


# ---------------------------------------------------------------------------
# メインパーサー関数
# ---------------------------------------------------------------------------


def parse_requirement(text: str) -> RequirementDiagram:
    """
    Mermaid requirementDiagram テキストを解析して RequirementDiagram を返す。

    名前・テキストは "..." クォートで日本語・スペースに対応する。
    文法エラーの行は警告なくスキップする。

    Parameters
    ----------
    text : str
        Mermaid requirementDiagram 構文のテキスト。

    Returns
    -------
    RequirementDiagram
        解析結果のデータクラス。
    """
    diagram = RequirementDiagram()

    # 状態管理
    # in_block: None | "requirement" | "element"
    in_block: str | None = None
    current_req: RequirementNode | None = None
    current_elem: ElementNode | None = None

    lines = text.splitlines()

    for line in lines:
        stripped = line.strip()

        # 空行・コメント行・アクセシビリティ行はスキップする
        if not stripped or _RE_COMMENT.match(line) or _RE_ACC.match(line):
            continue

        # ヘッダー行はスキップする
        if _RE_HEADER.match(line):
            continue

        # ブロック内の処理
        if in_block is not None:
            # ブロック終了
            if _RE_BLOCK_END.match(line):
                if in_block == "requirement" and current_req is not None:
                    diagram.requirements[current_req.name] = current_req
                    current_req = None
                elif in_block == "element" and current_elem is not None:
                    diagram.elements[current_elem.name] = current_elem
                    current_elem = None
                in_block = None
                continue

            # フィールド行の解析
            m = _RE_FIELD.match(line)
            if m:
                key = m.group("key").lower().strip()
                val = _strip_quotes(m.group("val").strip())
                _apply_field(key, val, current_req, current_elem)
            continue

        # ブロック外の処理

        # direction 行
        m = _RE_DIRECTION.match(line)
        if m:
            diagram.direction = m.group(1).upper()
            continue

        # 要件ブロック開始行
        m = _RE_REQ_START.match(line)
        if m:
            rtype_raw = m.group("rtype").lower()
            # 型のノーマル化キーは全小文字
            rtype_key = rtype_raw.replace(" ", "").lower()
            stereotype = REQUIREMENT_TYPES.get(rtype_key, "Requirement")
            name = _strip_quotes(m.group("name").strip())
            cls1 = m.group("cls1")
            cls2 = m.group("cls2")
            classes: list[str] = []
            if cls1:
                classes.append(cls1)
            if cls2:
                classes.append(cls2)
            current_req = RequirementNode(
                name=name,
                req_type=rtype_key,
                stereotype=stereotype,
                classes=classes,
            )
            # 同一行に } が含まれる場合はインラインブロックとして即時処理する
            brace_open = line.find("{")
            brace_close = line.rfind("}")
            if brace_open >= 0 and brace_close > brace_open:
                inline_text = line[brace_open + 1 : brace_close]
                _parse_inline_fields(inline_text, current_req, None)
                diagram.requirements[current_req.name] = current_req
                current_req = None
            else:
                in_block = "requirement"
            continue

        # エレメントブロック開始行
        m = _RE_ELEM_START.match(line)
        if m:
            name = _strip_quotes(m.group("name").strip())
            cls1 = m.group("cls1")
            cls2 = m.group("cls2")
            classes = []
            if cls1:
                classes.append(cls1)
            if cls2:
                classes.append(cls2)
            current_elem = ElementNode(name=name, classes=classes)
            # 同一行に } が含まれる場合はインラインブロックとして即時処理する
            brace_open = line.find("{")
            brace_close = line.rfind("}")
            if brace_open >= 0 and brace_close > brace_open:
                inline_text = line[brace_open + 1 : brace_close]
                _parse_inline_fields(inline_text, None, current_elem)
                diagram.elements[current_elem.name] = current_elem
                current_elem = None
            else:
                in_block = "element"
            continue

        # classDef 行
        m = _RE_CLASSDEF.match(line)
        if m:
            cls_name = m.group("name")
            style = _parse_node_style(m.group("styles"))
            diagram.class_defs[cls_name] = style
            continue

        # class 適用行
        m = _RE_CLASS_APPLY.match(line)
        if m:
            names_raw = m.group("names")
            cls_name = m.group("cls").strip()
            for n in names_raw.split(","):
                n = n.strip()
                if n in diagram.requirements:
                    if cls_name not in diagram.requirements[n].classes:
                        diagram.requirements[n].classes.append(cls_name)
                if n in diagram.elements:
                    if cls_name not in diagram.elements[n].classes:
                        diagram.elements[n].classes.append(cls_name)
            continue

        # style 直接適用行
        m = _RE_STYLE_APPLY.match(line)
        if m:
            name = _strip_quotes(m.group("name").strip())
            style = _parse_node_style(m.group("styles"))
            if name in diagram.requirements:
                diagram.requirements[name].style = style
            if name in diagram.elements:
                diagram.elements[name].style = style
            continue

        # 順方向リレーション行
        m = _RE_REL_FWD.match(line)
        if m:
            src = _strip_quotes(m.group("src").strip())
            dst = _strip_quotes(m.group("dst").strip())
            rel_type = m.group("rtype").lower()
            diagram.relations.append(Relationship(src=src, dst=dst, rel_type=rel_type))
            continue

        # 逆方向リレーション行
        m = _RE_REL_REV.match(line)
        if m:
            src = _strip_quotes(m.group("src").strip())
            dst = _strip_quotes(m.group("dst").strip())
            rel_type = m.group("rtype").lower()
            diagram.relations.append(Relationship(src=src, dst=dst, rel_type=rel_type))
            continue

        # ブロック外の :::class ショートハンド適用行（単独行）
        m = _RE_CLASS_SHORTHAND.match(line)
        if m and m.group("cls"):
            name = _strip_quotes(m.group("name").strip())
            cls_name = m.group("cls")
            if name in diagram.requirements:
                if cls_name not in diagram.requirements[name].classes:
                    diagram.requirements[name].classes.append(cls_name)
            if name in diagram.elements:
                if cls_name not in diagram.elements[name].classes:
                    diagram.elements[name].classes.append(cls_name)
            continue

        # その他の行（認識できない構文）はスキップする

    # ブロックが閉じられないまま終了した場合の処理
    if in_block == "requirement" and current_req is not None:
        diagram.requirements[current_req.name] = current_req
    elif in_block == "element" and current_elem is not None:
        diagram.elements[current_elem.name] = current_elem

    return diagram


def resolve_node_style(
    node_classes: list[str],
    direct_style: NodeStyle,
    class_defs: dict[str, NodeStyle],
) -> NodeStyle:
    """
    ノードの有効スタイルを classDef → class 適用 → 直接指定の優先順で解決する。

    Parameters
    ----------
    node_classes : list[str]
        ノードに適用されたクラス名のリスト（適用順）。
    direct_style : NodeStyle
        ノードへの直接スタイル指定（style キーワード）。
    class_defs : dict[str, NodeStyle]
        classDef 定義辞書。

    Returns
    -------
    NodeStyle
        解決済みの有効スタイル。
    """
    # デフォルトスタイルから始める
    resolved = NodeStyle()
    # classDef の適用（リスト順に重ねる）
    for cls_name in node_classes:
        if cls_name in class_defs:
            resolved = _merge_styles(resolved, class_defs[cls_name])
    # 直接スタイル指定で上書き
    resolved = _merge_styles(resolved, direct_style)
    return resolved
