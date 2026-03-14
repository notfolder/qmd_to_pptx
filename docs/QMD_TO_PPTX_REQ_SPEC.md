# Pythonライブラリ設計

## qmd → PowerPoint 変換ライブラリ

------------------------------------------------------------------------

# 1 結論

Quarto Markdown（`.qmd`）を解析し、既存の PowerPoint テンプレートを基に
PowerPoint ファイルを生成する Python ライブラリを設計する。

特徴

-   YAML フロントマター解析
-   slide 分割
-   Markdown → DOM
-   Mermaid → PowerPoint 図
-   LaTeX 数式 → PowerPoint 数式
-   PowerPoint テンプレート対応

主要ライブラリ

  用途             ライブラリ
  ---------------- -------------------
  Markdown解析     Python-Markdown
  Mermaid解析      mermaid-parser-py
  グラフ           NetworkX
  PowerPoint生成   python-pptx
  LaTeX→MathML     latex2mathml
------------------------------------------------------------------------

# 2 全体アーキテクチャ

    qmd file
     │
     ├─ YAML parser
     │
     ├─ Slide splitter
     │
     └─ Markdown parser
            │
            ▼
       python-markdown
            │
            ▼
         DOM Tree
            │
            ▼
    DOM renderer
            │
            ├─ Text renderer
            │
            ├─ Mermaid renderer
            │     mermaid-parser
            │     ↓
            │     NetworkX
            │     ↓
            │     python-pptx
            │
            └─ Math renderer
                  arithmatex
                  ↓
                  latex2mathml
                  ↓
                  mathml2omml
                  ↓
                  python-pptx
------------------------------------------------------------------------

# 3 処理フロー

    QMD
     ↓
    YAML parse
     ↓
    slide split
     ↓
    Markdown parse
     ↓
    DOM
     ↓
    DOM traversal
     ↓
    slide rendering
     ↓
    PPTX

Mermaid

    Mermaid
     ↓
    mermaid-parser
     ↓
    NetworkX
     ↓
    PowerPoint shapes

数式

    LaTeX
     ↓
    MathML
     ↓
    OMML
     ↓
    PowerPoint equation

------------------------------------------------------------------------

# 4 YAML解析

目的

-   slide theme
-   layout
-   title
-   template

例

    ---
    title: Test
    format: pptx
    theme: default
    ---

------------------------------------------------------------------------

# 5 Slide 分割

Quartoでは

    ---

または

    ## heading

で slide 分割。

------------------------------------------------------------------------

# 6 Markdown解析

使用

-   Python-Markdown

extensions

    extensions = [
        "pymdownx.superfences",
        "pymdownx.arithmatex",
        "tables",
        "fenced_code"
    ]

DOM生成

    markdown
     ↓
    ElementTree

------------------------------------------------------------------------

# 6 DOMトラバーサ

    ElementTree
     ├─ h1
     ├─ p
     ├─ ul
     ├─ code
     └─ div.arithmatex

------------------------------------------------------------------------

# 7 Mermaidレンダラー

Mermaid検出

    code block
    class="language-mermaid"

処理

    Mermaid
     ↓
    mermaid-parser
     ↓
    Graph
     ↓
    NetworkX
     ↓
    layout
     ↓
    python-pptx shapes

------------------------------------------------------------------------

# 8 PowerPoint図生成

    slide.shapes.add_shape()
    slide.shapes.add_connector()

レイアウト

    networkx.spring_layout

------------------------------------------------------------------------

# 9 数式レンダラー

検出

    span.arithmatex
    div.arithmatex

処理

    LaTeX
     ↓
    latex2mathml
     ↓
    MathML
     ↓
    mathml2omml
     ↓
    OMML
     ↓
    python-pptx

------------------------------------------------------------------------

# 10 PowerPoint数式

PowerPointは

    OMML

形式。

例

    <m:oMath>
      <m:r>
        <m:t>E=mc^2</m:t>
      </m:r>
    </m:oMath>

------------------------------------------------------------------------

# 11 テンプレート対応

入力

    template.pptx

ロード

    prs = Presentation(template)

------------------------------------------------------------------------

# 12 Slide Renderer

責務

    DOM
     ↓
    PowerPoint shapes

処理

    render_heading
    render_text
    render_list
    render_mermaid
    render_math

------------------------------------------------------------------------

# 13 DOM→PPTマッピング

  DOM            PowerPoint
  -------------- ------------
  h1             title
  p              textbox
  ul             bullet
  table          table
  code mermaid   diagram
  arithmatex     equation

------------------------------------------------------------------------

# 14 API設計

メインAPI

convert(
    input="---¥n",
    template="template.pptx",
    output="slides.pptx"
)

