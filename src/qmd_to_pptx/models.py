"""
全コンポーネントが共有するデータクラスおよびEnumの定義モジュール。

他モジュールへの依存を持たない。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum, auto
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    import xml.etree.ElementTree as ET


class SeparatorType(Enum):
    """スライド区切りの種別を表す列挙型。"""

    # `#` で始まるレベル1見出し。Section Headerスライドを生成する（slide-level: 2 の場合）
    HEADING1 = auto()
    # `##` で始まるレベル2見出し。通常のコンテンツスライドを生成する（slide-level: 2 の場合）
    HEADING2 = auto()
    # `---` 水平区切り線。タイトルなしスライドを生成する
    RULER = auto()


class DOMNodeType(Enum):
    """DOMトラバーサーが識別するノードの種別を表す列挙型。"""

    H1 = auto()              # <h1> タグ
    H2 = auto()              # <h2> タグ
    PARAGRAPH = auto()       # <p> タグ
    UL = auto()              # <ul> タグ
    OL = auto()              # <ol> タグ
    TABLE = auto()           # <table> タグ
    CODE = auto()            # <code> タグ（Mermaid以外）
    MERMAID = auto()         # <code class="language-mermaid"> タグ
    FORMULA_INLINE = auto()  # <span class="arithmatex"> タグ
    FORMULA_BLOCK = auto()   # <div class="arithmatex"> タグ
    NOTES = auto()           # <div class="notes"> タグ
    COLUMNS = auto()         # <div class="columns"> タグ
    INCREMENTAL = auto()     # <div class="incremental"> タグ
    NON_INCREMENTAL = auto() # <div class="nonincremental"> タグ


@dataclass
class SlideMetadata:
    """YAMLパーサーが生成し、スライドレンダラーへ渡すメタデータ。"""

    # プレゼンテーションのタイトル。未設定時は空文字
    title: str = ""
    # 作成者名。未設定時は空文字
    author: str = ""
    # 作成日。未設定時は空文字
    date: str = ""
    # スライドテーマ名。未設定時は空文字
    theme: str = ""
    # YAMLフロントマターの format.pptx.reference-doc の値。未設定時は None
    reference_doc: str | None = None
    # リストのデフォルト逐次表示設定。未設定時は False
    incremental: bool = False
    # スライド区切りとして扱う見出しレベル（1または2）。未設定時は 2
    slide_level: int = 2


@dataclass
class SlideContent:
    """スライド分割器が生成し、スライドレンダラーへ渡す各スライドの内容。"""

    # スライド本文のMarkdownテキスト
    body_text: str
    # このスライドを生成した区切りの種別
    separator_type: SeparatorType
    # 見出し行から抽出したスライドタイトル。水平区切り線由来の場合は空文字
    title: str = ""
    # {background-image="..."} 属性で指定された画像パス。未指定時は None
    background_image: str | None = None


@dataclass
class DOMNodeInfo:
    """DOMトラバーサーが返すノード情報。"""

    # ノードの種別
    node_type: DOMNodeType
    # ElementTree形式のノード要素
    element: "ET.Element"


@dataclass
class PlaceholderInfo:
    """レイアウトJSONの各プレースホルダー定義。"""

    # python-pptx の placeholder_format.idx に対応するプレースホルダー番号
    idx: int
    # コンテンツの役割（title / body / subtitle / left_content / right_content / left_header / right_header / caption）
    role: str
    # 左端座標（EMU）
    left: int
    # 上端座標（EMU）
    top: int
    # 幅（EMU）
    width: int
    # 高さ（EMU）
    height: int


@dataclass
class LayoutDef:
    """レイアウト1つ分のプレースホルダー集合。"""

    # そのレイアウトに属するプレースホルダーのリスト
    placeholders: list[PlaceholderInfo] = field(default_factory=list)


@dataclass
class LayoutJSON:
    """default_layout.json 全体を表すトップレベルオブジェクト。"""

    # スライド幅（EMU）
    slide_width_emu: int
    # スライド高さ（EMU）
    slide_height_emu: int
    # レイアウト名をキー、LayoutDef を値とする辞書
    layouts: dict[str, LayoutDef] = field(default_factory=dict)
