# 独自パーサーと mermaid.js パーサーの差異分析・修正方針

## 対象パーサー一覧

| ダイアグラム種 | 本プロジェクトのパーサー | mermaid.js 側のファイル | 文法方式 |
|---|---|---|---|
| Gantt | `gantt_parser.py` | `diagrams/gantt/parser/gantt.jison` | Jison |
| gitGraph | `gitgraph_parser.py` | `packages/parser/src/language/gitGraph/gitGraph.langium` | Langium |
| User Journey | `journey_parser.py` | `diagrams/user-journey/parser/journey.jison` | Jison |
| Pie | `pie_parser.py` | `packages/parser/src/language/pie/pie.langium` | Langium |
| Quadrant Chart | `quadrant_parser.py` | `diagrams/quadrant-chart/parser/quadrant.jison` | Jison |
| Requirement | `requirement_parser.py` | `diagrams/requirement/parser/requirementDiagram.jison` | Jison |
| Timeline | `timeline_parser.py` | `diagrams/timeline/parser/timeline.jison` | Jison |

---

## 各パーサーの差異分析

### 1. Gantt（`gantt_parser.py` vs `gantt.jison`）

#### mermaid.js が対応しているが本実装が未対応の機能

| 機能 | mermaid.js | 本実装 | 影響度 | 影響内容 |
|---|---|---|---|---|
| `inclusiveEndDates` | 終了日を含む（当日まで）として解釈 | 未実装（常に非包含） | 高 | タスク期間の表示日付がずれる |
| `includes` | 含める日付を明示指定 | 未実装 | 低 | `excludes` の逆指定で週末等を含める設定が無効 |
| `weekday` | 週の開始曜日の設定（monday〜sunday） | 無視 | 低 | ガントチャートの週末強調表示が固定される |
| `weekend` | 週末の定義（friday/saturday） | 無視 | 低 | ガントチャートの週末強調表示が固定される |
| `topAxis` | 軸を上部に表示 | 無視 | 低 | レイアウトのみに影響（上部軸が描画されない） |
| `click` コマンド | リンク・コールバックの設定 | 未実装 | なし | インタラクティビティのみで描画に影響なし |
| `accTitle` / `accDescr` | アクセシビリティ情報 | 無視 | なし | 描画には不要 |

#### 対応済みの差異（本修正で解消）

以下は当初未対応だったが、本 PR で修正・対応した。

| 機能 | mermaid.js (dayjs) | 修正前の本実装 | 修正後の本実装 |
|---|---|---|---|
| `dateFormat YYYY-MM-DD HH:mm` | 時刻付きフォーマットで日付を解析 | 未対応（YYYY-MM-DDへフォールバックして解析失敗） | 対応済み（時刻部を切り捨てて日付のみ使用） |
| `dateFormat YYYY-MM-DDTHH:mm` | ISO 8601 ～ T 区切り時刻 | 未対応 | 対応済み |
| `dateFormat YYYY-MM-DD HH:mm:ss` | 秒まで含む時刻付きフォーマット | 未対応 | 対応済み |
| `dateFormat YYYY-MM-DDTHH:mm:ss` | ISO 8601 秒まで | 未対応 | 対応済み |
| `dateFormat YYYYMMDD` | コンパクト日付フォーマット | 未対応 | 対応済み |
| `dateFormat X` | Unixタイムスタンプ（秒） | 未対応 | 対応済み |
| `dateFormat x` | Unixタイムスタンプ（ミリ秒） | 未対応 | 対応済み |
| 期間単位 `M`（月）| 月単位の期間指定（dayjs `M`） | 未対応 | 対応済み（30日近似） |
| 期間単位 `y/Y`（年）| 年単位の期間指定（dayjs `y`） | 未対応 | 対応済み（365日近似） |
| 期間単位 `m`（分）| 分単位の期間指定 | 未対応 | 対応済み（最小1日） |
| 期間単位 `s`（秒）| 秒単位の期間指定 | 未対応 | 対応済み（最小1日） |
| 期間単位 `ms`（ミリ秒）| ミリ秒単位の期間指定 | 未対応 | 対応済み（最小1日） |
| `until` 複数IDの参照 | `until id1 id2` の形式で複数タスクを参照 | 単一IDのみ対応 | 対応済み（複数IDの最小終了日を使用） |

**注意**: `until` は mermaid.js では参照タスクの **開始時刻** の最小値を終了日として使用するが、本実装では開始日マップを保持しないため**終了日の最小値**で代替している。非重複タスクでは実質的な影響はない。

#### 重要度の高い差異の詳細

**`inclusiveEndDates` の動作差異**

mermaid.js の `gantt.jison` では、`inclusiveEndDates` キーワードを宣言すると終了日を「含む」として処理される。これにより、たとえば `2024-01-01, 2024-01-03` と記述した場合、通常は 1/1〜1/2（3日を含まない）だが、`inclusiveEndDates` 有効時は 1/1〜1/3（3日を含む）となる。

現状の `gantt_parser.py` は `inclusiveEndDates` を行ごとスキップしており、終了日の解釈が常に非包含になっている。

---

### 2. gitGraph（`gitgraph_parser.py` vs `gitGraph.langium`）

#### mermaid.js が対応しているが本実装が未対応の機能

| 機能 | mermaid.js (Langium) | 本実装 | 影響度 | 影響内容 |
|---|---|---|---|---|
| `commit "message"` | `msg:` キーワードなしで文字列を直接指定可能 | 未対応（`msg:` 必須） | 中 | コミットメッセージが表示されない |
| 複数タグ | `tag:` を複数回指定可（`tags+=STRING`） | 単一タグのみ | 低 | 最初のタグのみ表示される |
| cherry-pick の `tag:` | `cherry-pick` コマンドで `tag:` 指定可 | 未対応 | 低 | タグが表示されない |

#### 差異の詳細

**`commit "message"` 形式の未対応**

Langium 文法の定義は以下の通りで、`'msg:'?` とオプショナルになっている。

```ebnf
Commit:
    'commit'
    (
        'id:' id=STRING
        | 'msg:'? message=STRING   ← msg: なしでも文字列を受け付ける
        | 'tag:' tags+=STRING      ← タグは複数指定可
        | 'type:' type=(...)
    )* EOL;
```

`gitgraph_parser.py` は `msg: "..."` 形式のみ対応しており、`commit "直接メッセージ"` の構文は無視される。

---

### 3. User Journey（`journey_parser.py` vs `journey.jison`）

#### mermaid.js が対応しているが本実装が未対応の機能

| 機能 | mermaid.js | 本実装 | 影響度 | 影響内容 |
|---|---|---|---|---|
| `#` コメント | `\#[^\n]*` で行中の `#` 以降をコメントとして処理 | 未対応（`%%` のみ） | 低 | `#` コメントが含まれると解析エラーの可能性 |
| `accTitle` / `accDescr` | アクセシビリティ情報 | スキップ処理あり | なし | — |

#### 差異の詳細

journey.jison では `\#[^\n]*` をコメントとして処理するルールがある。`journey_parser.py` は `%%` コメントのみ対応しており、`#` から行末をコメントとして扱わない。`#` 文字を含むタスク名・セクション名が誤って解析される可能性がある。

---

### 4. Pie（`pie_parser.py` vs `pie.langium`）

#### mermaid.js が対応しているが本実装が未対応の機能

| 機能 | mermaid.js (Langium) | 本実装 | 影響度 | 影響内容 |
|---|---|---|---|---|
| 負の数値 | `/-?[0-9]+\.[0-9]+/` 等、負数も文法上許容 | 未対応（正数のみ） | なし | 値 ≤ 0 はスキップするため実質影響なし |
| `accTitle` / `accDescr` | アクセシビリティ情報 | 未対応 | なし | — |

`pie.langium` では `NUMBER_PIE` ターミナルが負数を含む文法定義になっているが、パイチャートの性質上、負の値を持つデータは意味をなさない。本実装では正数のみ受け付けており、実際の描画結果への影響はない。

---

### 5. Quadrant Chart（`quadrant_parser.py` vs `quadrant.jison`）

#### mermaid.js が対応しているが本実装が未対応の機能

| 機能 | mermaid.js | 本実装 | 影響度 | 影響内容 |
|---|---|---|---|---|
| Markdown テキスト形式 | `` ["`...`"] `` のMarkdown埋め込み | 未対応 | 低 | ポイント名・軸ラベルでMarkdown未レンダリング |
| 軸デリミタのみ形式 | `x-axis Low -->` で右ラベルなしの場合、左ラベルに ` ⟶` を付加 | 未実装 | 低 | ラベル末尾の矢印が省略される |
| 複数ダッシュのデリミタ | `--->`、`---->`等も許容（`-{2,}>`） | 未対応（`-->` のみ） | 低 | 複数ダッシュ使用時にデリミタ認識失敗 |

#### 差異の詳細

**軸デリミタのみ形式**

quadrant.jison では `X-AXIS text AXIS-TEXT-DELIMITER` ルール（右ラベルなし）が存在し、左ラベルテキストに ` ⟶` を付加して表示する。`quadrant_parser.py` はこのパターンを実装しておらず、`-->` なしの場合は `x_label_right = ""` のまま処理される。

---

### 6. Requirement（`requirement_parser.py` vs `requirementDiagram.jison`）

#### mermaid.js が対応しているが本実装が未対応の機能

`requirement_parser.py` は非常に詳細に実装されており、`requirementDiagram.jison` の主要機能のほとんどをカバーしている。

| 機能 | mermaid.js | 本実装 | 影響度 | 影響内容 |
|---|---|---|---|---|
| `direction` | TB/BT/LR/RL サポート | 実装済み | — | — |
| 要件タイプ全種類 | requirement/functional/interface/performance/physical/designConstraint | 実装済み | — | — |
| リレーション全種類 | contains/copies/derives/satisfies/verifies/refines/traces | 実装済み | — | — |
| `classDef` / `class` / `style` | スタイル定義・適用 | 実装済み | — | — |
| `title` | タイトル設定 | 実装済み | — | — |
| `accTitle` / `accDescr` | アクセシビリティ | スキップ処理あり | なし | 描画には不要 |

**本パーサーは機能差異が最も小さく、修正は不要と判断できる。**

---

### 7. Timeline（`timeline_parser.py` vs `timeline.jison`）

#### mermaid.js が対応しているが本実装が未対応の機能

| 機能 | mermaid.js | 本実装 | 影響度 | 影響内容 |
|---|---|---|---|---|
| `direction LR` | `timeline LR` でタイムラインの向きを左→右に設定 | 未対応（常に縦方向） | 中 | レイアウト方向が固定される |
| `direction TD` | `timeline TD` でタイムラインの向きを上→下に設定 | 未対応 | 中 | レイアウト方向が固定される |
| `#` コメント | `\#[^\n]*` で行中の `#` 以降をコメントとして処理 | 未対応（`%%` のみ） | 低 | `#` コメントが含まれると誤解析の可能性 |

#### 差異の詳細

**direction 未対応**

timeline.jison では以下のようにヘッダーで方向を指定できる。

```jison
"timeline"[ \t]+LR       return 'timeline_lr';
"timeline"[ \t]+TD       return 'timeline_td';
```

`timeline_parser.py` は `timeline` キーワードのみを認識しており、`LR`/`TD` 指定を解析しない。`TimelineData` データクラスにも `direction` フィールドが存在しない。

---

## 差異の影響度サマリー

### 描画に影響する差異（対応優先度：高）

| パーサー | 差異内容 | 影響 |
|---|---|---|
| gantt | `inclusiveEndDates` 未対応 | タスク終了日の解釈が異なりバー長さがずれる |
| gitGraph | `commit "message"` 形式未対応 | msg: なしのコミットメッセージが無視される |

### 描画に部分的に影響する差異（対応優先度：中）

| パーサー | 差異内容 | 影響 |
|---|---|---|
| timeline | `direction LR/TD` 未対応 | レイアウト方向の指定が無視される |
| gitGraph | 複数タグ未対応 | 2番目以降のタグが表示されない |

### 軽微な差異（対応優先度：低）

| パーサー | 差異内容 | 影響 |
|---|---|---|
| journey | `#` コメント未対応 | `#` 文字が含まれる場合に誤解析の可能性 |
| timeline | `#` コメント未対応 | 同上 |
| gantt | `includes` / `weekday` / `weekend` 未対応 | 週末強調表示の設定が無視される |
| quadrant | 複数ダッシュのデリミタ未対応 | `---->` 等を使った場合にデリミタ認識失敗 |
| quadrant | 軸デリミタのみ形式での `⟶` 付加未実装 | 右ラベルなし時の矢印表示が省略される |
| quadrant | Markdown テキスト未対応 | `` ["`...`"] `` 形式のMarkdownが解析されない |
| pie | 負の数値（文法上） | 実質影響なし |

---

## 修正方針（現状実装の修正）

### 修正が必要な箇所と対応内容

#### gantt_parser.py

- `inclusiveEndDates` ディレクティブを検出し、`GanttChart` に `inclusive_end_dates: bool` フィールドを追加する
- `includes` ディレクティブを `GanttChart.includes: list[str]` として保持する
- `weekday` / `weekend` ディレクティブを `GanttChart.weekday` / `GanttChart.weekend_day` として保持する（レンダラーで週末強調表示に利用）

最優先：`inclusiveEndDates` は終了日の解釈に関わるため先に対応する。

#### gitgraph_parser.py

- `commit "message"` 形式（`msg:` キーワードなし）を `_RE_COMMIT_MSG` の正規表現拡張で対応する
- `tags+=STRING`（複数タグ）を `GitCommit.tags: list[str]` フィールドに変更して複数タグを保持する
- `cherry-pick` の `tag:` 属性を抽出できるよう `_RE_CHERRY_PICK` の解析処理を拡張する

#### journey_parser.py

- `#` コメントのスキップ処理を追加する（`_RE_COMMENT` に `#` 行を追加、もしくは行内コメント除去処理を追加）

#### timeline_parser.py

- `TimelineData` に `direction: str` フィールドを追加する（デフォルト `"TD"`）
- ヘッダー解析で `timeline LR` / `timeline TD` を認識して `direction` を設定する
- `#` コメントのスキップ処理を追加する

#### quadrant_parser.py

- 複数ダッシュのデリミタ（`-->`、`--->`等）を許容するよう `_RE_X_AXIS` / `_RE_Y_AXIS` の正規表現を修正する（`-->` → `-{2,}>`）
- `x-axis Low -->` 形式（右ラベルなし）で左ラベルに ` ⟶` を付加する処理を追加する

---

## Jison から lark への移行の検討

### Python lark と Jison の仕様比較

| 項目 | Python lark | Jison (JavaScript) |
|---|---|---|
| 解析アルゴリズム | Earley / LALR(1) / CYK | LALR(1) のみ |
| 文法形式 | EBNF（拡張BNF）・読みやすい宣言型 | Bison/Yacc 互換の BNF |
| 意味アクション | Python の Transformer/Visitor クラス | JavaScript インラインコード |
| AST 自動構築 | 自動生成 | 手動実装が必要 |
| Unicode 対応 | ネイティブ | 制限あり（`\w` が ASCII のみ） |
| 保守性 | 高（文法と処理を分離） | 中（文法とJS処理が混在） |
| 依存関係 | `lark` パッケージ（pip install） | Node.js 環境（ビルド必要） |
| 成熟度 | 活発に開発中 | メンテナンスが低活性 |

### Jison ファイルから lark 文法への変換アプローチ

Jison ファイルは BNF 形式の文法 + JavaScript 意味アクションで構成されているのに対し、lark は EBNF 形式の文法定義のみで、意味処理は Python の `Transformer`/`Visitor` クラスで実装する。

変換の主な対応関係は以下の通り。

| Jison | lark |
|---|---|
| `%lex` セクション（正規表現トークン定義） | 文法内の `terminal` 定義または `%import` |
| `%%` セクション（BNF 文法規則） | lark の EBNF 規則（`: ` と `\|` で記述） |
| `{ yy.setXxx($1); }` 意味アクション | Python `Transformer` クラスのメソッドに移植 |
| `%options case-insensitive` | 正規表現に `(?i)` フラグを付与 |
| `<<EOF>>` | 文法上は通常不要（lark が EOF を自動処理） |

変換作業の規模感（1パーサーあたり）：
- 文法定義の変換：1〜2日
- Transformer 実装とテスト：1〜2日
- 既存テストの通過確認：0.5日

全7パーサーを移行する場合、合計 2〜3週間程度の工数が見込まれる。

### 結論：現状修正 vs lark への移行

**推奨：現状の正規表現ベース実装を部分修正する**

理由は以下の通り。

1. **既存実装が安定して動作している**。全パーサーにテストが存在し、主要機能は正しく動作している。lark への移行は全実装の書き直しを意味し、デグレードリスクが高い。

2. **差異の大半は軽微かつ局所的**。影響度の高い差異（`inclusiveEndDates`、`commit "message"` 形式）は数十行の修正で対応可能であり、全体的な移行は費用対効果が低い。

3. **依存関係の増加を避けられる**。現在の実装は Python 標準ライブラリのみで動作している。`lark` を追加することは、依存関係の増加とバージョン管理の複雑化につながる。

4. **mermaid.js 側の文法は今後も変化する**。Langium への移行が進行中であり、いま lark 文法を整備しても追従コストが高い。正規表現ベースの実装の方が差分修正は容易である。

**lark への移行を検討すべきケース**（将来的な判断基準）：

- 解析できないケースが頻発し、正規表現での対応が困難になった場合
- 新しいダイアグラム種を多数追加する場合（文法定義が複雑化する場合）
- 文法ファイルを mermaid.js と共有・同期したい要件が生じた場合

現時点では **現状修正による対応を選択し、影響度の高い差異から優先的に修正することを推奨する**。
