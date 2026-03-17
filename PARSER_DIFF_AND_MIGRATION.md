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

## 全文法要素の対応状況詳細

各パーサーの全文法要素について、mermaid.js との対応状況を網羅的に整理する。

凡例：
- ✅ 対応済み
- ❌ 未対応（動作に影響あり）
- ⚠️ 部分対応（スキップ・無視のため描画には影響なし）
- 🚫 対応不可（構文上の制約により実装困難）

---

### 1. Gantt（`gantt_parser.py`）

#### ディレクティブ

| 文法要素 | mermaid.js 仕様 | 対応状況 | 備考 |
|---|---|---|---|
| `gantt` ヘッダー | 必須キーワード | ✅ | |
| `title <text>` | タイトル設定 | ✅ | |
| `dateFormat <fmt>` | 日付フォーマット指定 | ✅ | 以下の日付フォーマット一覧を参照 |
| `excludes <list>` | 除外日の指定（weekends, dayname, date） | ✅ | パース済み（レンダラーでの利用は未実装） |
| `axisFormat <fmt>` | 軸ラベルの表示フォーマット | ⚠️ | 行スキップ（軸表示なし） |
| `tickInterval <val>` | 軸目盛り間隔 | ⚠️ | 行スキップ |
| `todayMarker <on\|off\|css>` | 今日マーカー表示 | ⚠️ | 行スキップ |
| `inclusiveEndDates` | 終了日を包含して解釈 | ❌ | 常に非包含（終了日が1日ずれる） |
| `includes <list>` | 含める日付の明示指定 | ⚠️ | 行スキップ（`excludes` のレンダラー利用も未実装のため実質影響なし） |
| `weekday <day>` | 週の開始曜日設定 | ⚠️ | 行スキップ |
| `weekend <fri\|sat>` | 週末定義 | ⚠️ | 行スキップ |
| `topAxis` | 軸を上部に表示 | ⚠️ | 行スキップ |
| `click <id> <callback>` | インタラクション設定 | ⚠️ | 行スキップ（描画に影響なし） |
| `accTitle: <text>` | アクセシビリティタイトル | ⚠️ | 行スキップ（描画に影響なし） |
| `accDescr: <text>` | アクセシビリティ説明 | ⚠️ | 行スキップ（描画に影響なし） |
| `%% <text>` | コメント行 | ✅ | |

#### タスク構文

| 文法要素 | mermaid.js 仕様 | 対応状況 | 備考 |
|---|---|---|---|
| `section <name>` | セクション定義 | ✅ | |
| タスク名 | テキスト（コロンの前） | ✅ | |
| タグ `done` | 完了状態 | ✅ | |
| タグ `active` | 進行中状態 | ✅ | |
| タグ `crit` | クリティカル表示 | ✅ | |
| タグ `milestone` | マイルストーン表示 | ✅ | |
| `id, ` | タスクID明示指定 | ✅ | |
| 開始: 日付文字列 | dateFormat に準拠した日付 | ✅ | |
| 開始: `after <taskId>` | 参照タスク終了後 | ✅ | |
| 開始: `after <id1> <id2>` | 複数タスクの最大終了日後 | ✅ | |
| 終了: 日付文字列 | dateFormat に準拠した日付 | ✅ | |
| 終了: `<N>d` / `<N>D` | N日間 | ✅ | |
| 終了: `<N>w` / `<N>W` | N週間 | ✅ | |
| 終了: `<N>h` / `<N>H` | N時間（日単位切り上げ） | ✅ | |
| 終了: `<N>M` | N月（30日近似） | ✅ | |
| 終了: `<N>y` / `<N>Y` | N年（365日近似） | ✅ | |
| 終了: `<N>m` | N分（最小1日） | ✅ | |
| 終了: `<N>s` / `<N>S` | N秒（最小1日） | ✅ | |
| 終了: `<N>ms` | Nミリ秒（最小1日） | ✅ | |
| 終了: `until <taskId>` | 参照タスク開始まで | ✅ | 実装上は参照タスクの終了日最小値を使用 |
| 終了: `until <id1> <id2>` | 複数タスクの開始最小値まで | ✅ | |

#### dateFormat 対応一覧

| dateFormat 値 | 例 | 対応状況 | 備考 |
|---|---|---|---|
| `YYYY-MM-DD` | 2025-01-15 | ✅ | デフォルト形式 |
| `YYYY-MM` | 2025-01 | ✅ | 日は1日として補完 |
| `YYYY` | 2025 | ✅ | 月日は1月1日として補完 |
| `YYYYMMDD` | 20250115 | ✅ | |
| `YYYY-MM-DD HH:mm` | 2025-01-15 09:00 | ✅ | 時刻は切り捨て |
| `YYYY-MM-DD HH:mm:ss` | 2025-01-15 09:00:00 | ✅ | 時刻は切り捨て |
| `YYYY-MM-DDTHH:mm` | 2025-01-15T09:00 | ✅ | 時刻は切り捨て |
| `YYYY-MM-DDTHH:mm:ss` | 2025-01-15T09:00:00 | ✅ | 時刻は切り捨て |
| `YYYY-MM-DDTHH:mm:ssZ` | 2025-01-15T09:00:00+09:00 | ✅ | タイムゾーン付き ISO 8601 |
| `YYYY-MM-DDTHH:mmZ` | 2025-01-15T09:00+09:00 | ✅ | タイムゾーン付き ISO 8601 |
| `YYYY/MM/DD` | 2025/01/15 | ✅ | スラッシュ形式（アジア圏） |
| `YYYY/MM/DD HH:mm` | 2025/01/15 09:00 | ✅ | 時刻は切り捨て |
| `YYYY/MM/DD HH:mm:ss` | 2025/01/15 09:00:00 | ✅ | 時刻は切り捨て |
| `MM/DD/YYYY` | 01/15/2025 | ✅ | 米国スラッシュ形式 |
| `DD/MM/YYYY` | 15/01/2025 | ✅ | 欧州スラッシュ形式 |
| `M/D/YYYY` | 1/15/2025 | ✅ | 月日1桁許容 |
| `D/M/YYYY` | 15/1/2025 | ✅ | 日月1桁許容 |
| `DD.MM.YYYY` | 15.01.2025 | ✅ | ドット区切り（欧州） |
| `MM-DD-YYYY` | 01-15-2025 | ✅ | 米国ハイフン形式 |
| `YY-MM-DD` | 25-01-15 | ✅ | 2桁年（ハイフン） |
| `YY/MM/DD` | 25/01/15 | ✅ | 2桁年（スラッシュ） |
| `DD MMM YYYY` | 15 Jan 2025 | ✅ | 欧州英語月名（カンマなし） |
| `D MMM YYYY` | 5 Jan 2025 | ✅ | 欧州英語月名（1桁日、カンマなし） |
| `MMM D, YYYY` | Jan 15, 2025 | 🚫 | カンマがタスク行の区切り文字と衝突するため対応不可（下記補足参照） |
| `MMMM D, YYYY` | January 15, 2025 | 🚫 | カンマがタスク行の区切り文字と衝突するため対応不可（下記補足参照） |
| `X` | 1736899200 | ✅ | Unixタイムスタンプ（秒） |
| `x` | 1736899200000 | ✅ | Unixタイムスタンプ（ミリ秒） |

> **補足**: `MMM D, YYYY` と `MMMM D, YYYY` はカンマをタスク行の属性区切り文字として使用するため、パーサー構造上の制約により実現が困難。mermaid.js 側では dayjs がパース後の文字列全体を処理するため問題にならない。

---

### 2. gitGraph（`gitgraph_parser.py`）

#### コマンド・属性

| 文法要素 | mermaid.js 仕様 | 対応状況 | 備考 |
|---|---|---|---|
| `gitGraph` ヘッダー | 必須キーワード | ✅ | |
| `gitGraph LR:` | 左→右方向 | ✅ | |
| `gitGraph TB:` | 上→下方向 | ✅ | |
| `gitGraph BT:` | 下→上方向 | ✅ | |
| `commit` | コミット（ID自動生成） | ✅ | |
| `commit id: "..."` | コミットID明示指定 | ✅ | |
| `commit msg: "..."` | コミットメッセージ（`msg:` あり） | ✅ | |
| `commit "message"` | コミットメッセージ（`msg:` なし） | ❌ | `msg:` キーワードが必須 |
| `commit type: NORMAL` | 通常コミット | ✅ | |
| `commit type: REVERSE` | リバートコミット | ✅ | |
| `commit type: HIGHLIGHT` | ハイライトコミット | ✅ | |
| `commit tag: "..."` | タグ指定（1個） | ✅ | |
| `commit tag: "..." tag: "..."` | タグ複数指定 | ❌ | 1個目のみ取得 |
| `branch <name>` | ブランチ作成 | ✅ | |
| `branch <name> order: <n>` | 表示順序指定 | ✅ | |
| `checkout <name>` | ブランチ切り替え | ✅ | |
| `switch <name>` | ブランチ切り替え（別名） | ✅ | |
| `merge <name>` | マージコミット | ✅ | |
| `merge <name> id: "..."` | マージコミットID指定 | ✅ | |
| `merge <name> type: ...` | マージコミットタイプ | ✅ | |
| `merge <name> tag: "..."` | マージコミットタグ | ✅ | |
| `cherry-pick id: "..."` | チェリーピック | ✅ | |
| `cherry-pick id: "..." parent: "..."` | チェリーピック親指定 | ✅ | |
| `cherry-pick id: "..." tag: "..."` | チェリーピックタグ指定 | ❌ | `tag:` を抽出せず（コピー元IDがタグに設定される） |
| `%% <text>` | コメント行 | ✅ | |
| `accTitle: <text>` | アクセシビリティタイトル | ⚠️ | 行スキップ |
| `accDescr: <text>` | アクセシビリティ説明 | ⚠️ | 行スキップ |
| YAML `mainBranchName:` | デフォルトブランチ名 | ✅ | |
| YAML `mainBranchOrder:` | デフォルトブランチ順序 | ⚠️ | 未使用（order=0固定） |

---

### 3. User Journey（`journey_parser.py`）

#### コマンド・属性

| 文法要素 | mermaid.js 仕様 | 対応状況 | 備考 |
|---|---|---|---|
| `journey` ヘッダー | 必須キーワード | ✅ | |
| `title <text>` | タイトル設定 | ✅ | |
| `section <name>` | セクション定義 | ✅ | |
| タスク `name : score` | タスク名とスコア（1〜5） | ✅ | 範囲外はクランプ |
| タスク `name : score : actor1` | アクター1名 | ✅ | |
| タスク `name : score : actor1, actor2` | アクター複数 | ✅ | |
| スコア範囲（1〜5） | 1が最低、5が最高 | ✅ | 範囲外はクランプ処理 |
| `%% <text>` | コメント行 | ✅ | |
| `#` コメント | 行中の `#` 以降をコメント | ❌ | `%%` のみ対応 |
| `accTitle: <text>` | アクセシビリティタイトル | ⚠️ | 行スキップ |
| `accDescr: <text>` | アクセシビリティ説明 | ⚠️ | 行スキップ |
| `accDescr {{ ... }}` | 複数行アクセシビリティ説明 | ⚠️ | 行スキップ |

---

### 4. Pie（`pie_parser.py`）

#### コマンド・属性

| 文法要素 | mermaid.js 仕様 | 対応状況 | 備考 |
|---|---|---|---|
| `pie` ヘッダー | 必須キーワード | ✅ | |
| `pie showData` | データ値を表示 | ✅ | |
| `pie title "..."` | ヘッダー行にインラインタイトル | ✅ | |
| `title <text>` | タイトル（単独行） | ✅ | インラインより優先 |
| `"label" : <value>` | セクション定義（正の数値） | ✅ | |
| 整数値 | 整数のセクション値 | ✅ | |
| 小数値 | 小数点付きセクション値 | ✅ | |
| 負の数値 | 文法上許容（意味なし） | ❌ | 正数のみ受け付ける（スキップ） |
| `%% <text>` | コメント行 | ✅ | |
| YAML `textPosition:` | ラベル半径位置（0.0〜1.0） | ✅ | |
| `accTitle: <text>` | アクセシビリティタイトル | ⚠️ | 行スキップ |
| `accDescr: <text>` | アクセシビリティ説明 | ⚠️ | 行スキップ |

---

### 5. Quadrant Chart（`quadrant_parser.py`）

#### コマンド・属性

| 文法要素 | mermaid.js 仕様 | 対応状況 | 備考 |
|---|---|---|---|
| `quadrantChart` ヘッダー | 必須キーワード | ✅ | |
| `title <text>` | タイトル設定 | ✅ | |
| `x-axis <left>` | X軸左ラベルのみ | ✅ | |
| `x-axis <left> --> <right>` | X軸左右ラベル（`-->` 2ダッシュ） | ✅ | |
| `x-axis <left> --->` 等 | X軸（3ダッシュ以上） | ❌ | 2ダッシュ `-->` のみ対応 |
| `x-axis <left> -->` | デリミタのみ（右ラベルなし） | ❌ | 左ラベルに `⟶` 付加されない |
| `y-axis <bottom>` | Y軸下ラベルのみ | ✅ | |
| `y-axis <bottom> --> <top>` | Y軸上下ラベル（`-->` 2ダッシュ） | ✅ | |
| `y-axis <bottom> --->` 等 | Y軸（3ダッシュ以上） | ❌ | 2ダッシュのみ対応 |
| `quadrant-1 <text>` | 象限1（右上）ラベル | ✅ | |
| `quadrant-2 <text>` | 象限2（左上）ラベル | ✅ | |
| `quadrant-3 <text>` | 象限3（左下）ラベル | ✅ | |
| `quadrant-4 <text>` | 象限4（右下）ラベル | ✅ | |
| ポイント `name : [x, y]` | データポイント | ✅ | |
| ポイント `name:::class : [x, y]` | classDef 参照付きポイント | ✅ | |
| ポイントスタイル `color:` | 塗りつぶし色 | ✅ | |
| ポイントスタイル `radius:` | 半径 | ✅ | |
| ポイントスタイル `stroke-width:` | 枠線幅 | ✅ | |
| ポイントスタイル `stroke-color:` | 枠線色 | ✅ | |
| `classDef <name> <styles>` | クラス定義 | ✅ | |
| Markdown テキスト `` ["`...`"] `` | ラベルの Markdown 埋め込み | ❌ | プレーンテキストとして扱われる |
| `%% <text>` | コメント行 | ✅ | |
| `accTitle: <text>` | アクセシビリティタイトル | ⚠️ | 行スキップ |
| `accDescr: <text>` | アクセシビリティ説明 | ⚠️ | 行スキップ |

---

### 6. Requirement（`requirement_parser.py`）

#### コマンド・属性

| 文法要素 | mermaid.js 仕様 | 対応状況 | 備考 |
|---|---|---|---|
| `requirementDiagram` ヘッダー | 必須キーワード | ✅ | |
| `direction TB` | 上→下レイアウト | ✅ | |
| `direction BT` | 下→上レイアウト | ✅ | |
| `direction LR` | 左→右レイアウト | ✅ | |
| `direction RL` | 右→左レイアウト | ✅ | |
| `requirement <name> { ... }` | 要件ノード | ✅ | |
| `functionalRequirement <name> { ... }` | 機能要件ノード | ✅ | |
| `interfaceRequirement <name> { ... }` | インターフェース要件ノード | ✅ | |
| `performanceRequirement <name> { ... }` | 性能要件ノード | ✅ | |
| `physicalRequirement <name> { ... }` | 物理要件ノード | ✅ | |
| `designConstraint <name> { ... }` | 設計制約ノード | ✅ | |
| `element <name> { ... }` | エレメントノード | ✅ | |
| ノード属性 `id: <value>` | 要件ID | ✅ | |
| ノード属性 `text: <value>` | 要件テキスト | ✅ | |
| ノード属性 `risk: low\|medium\|high` | リスクレベル | ✅ | |
| ノード属性 `verifymethod:` | 検証方法 | ✅ | |
| エレメント属性 `type: <value>` | エレメントタイプ | ✅ | |
| エレメント属性 `docref: <value>` | ドキュメント参照 | ✅ | |
| リレーション `src - contains -> dst` | contains 関係 | ✅ | |
| リレーション `src - copies -> dst` | copies 関係 | ✅ | |
| リレーション `src - derives -> dst` | derives 関係 | ✅ | |
| リレーション `src - satisfies -> dst` | satisfies 関係 | ✅ | |
| リレーション `src - verifies -> dst` | verifies 関係 | ✅ | |
| リレーション `src - refines -> dst` | refines 関係 | ✅ | |
| リレーション `src - traces -> dst` | traces 関係 | ✅ | |
| 逆方向リレーション `dst <- relType - src` | 逆方向記法 | ✅ | |
| `classDef <name> fill:, stroke:, ...` | クラス定義 | ✅ | |
| `class <name1>,<name2> <className>` | クラス適用 | ✅ | |
| `style <name> fill:, stroke:, ...` | 直接スタイル指定 | ✅ | |
| `name:::className` | ノードへのクラス適用 | ✅ | |
| クォートなし名前 | ASCII ワード文字 | ✅ | |
| クォートあり名前 `"..."` | Unicode・スペース含む名前 | ✅ | |
| `%% <text>` | コメント行 | ✅ | |
| `accTitle: <text>` | アクセシビリティタイトル | ⚠️ | 行スキップ |
| `accDescr: <text>` | アクセシビリティ説明 | ⚠️ | 行スキップ |

---

### 7. Timeline（`timeline_parser.py`）

#### コマンド・属性

| 文法要素 | mermaid.js 仕様 | 対応状況 | 備考 |
|---|---|---|---|
| `timeline` ヘッダー | 必須キーワード | ✅ | |
| `timeline LR` | 左→右レイアウト | ❌ | `direction` フィールドなし |
| `timeline TD` | 上→下レイアウト | ❌ | `direction` フィールドなし |
| `title <text>` | タイトル設定 | ✅ | |
| `section <name>` | セクション定義 | ✅ | |
| `<period> : <event>` | 期間とイベント | ✅ | |
| `<period> : <event1> : <event2>` | 複数イベント（同一行） | ✅ | |
| `: <event>` | 継続行（イベント追加） | ✅ | |
| `<period>` のみ | イベントなしの期間 | ✅ | |
| `<br>` / `<br/>` / `<br />` | テキスト内改行タグ | ✅ | `\n` に変換 |
| `%% <text>` | コメント行 | ✅ | |
| `#` コメント | 行中の `#` 以降をコメント | ❌ | `%%` のみ対応 |
| `accTitle: <text>` | アクセシビリティタイトル | ⚠️ | 行スキップ |
| `accDescr: <text>` | アクセシビリティ説明 | ⚠️ | 行スキップ |

---

## 対応状況サマリー（全パーサー）

| パーサー | ✅ 主な対応項目 | ❌ 未対応（影響あり） | ⚠️ 部分対応 | 🚫 対応不可 |
|---|---|---|---|---|
| gantt | title / dateFormat（全主要形式）/ excludes / section / タスク属性全種 / 期間単位全種 | inclusiveEndDates | axisFormat / tickInterval / todayMarker / includes / weekday / weekend / topAxis / click / acc系 | MMM D, YYYY / MMMM D, YYYY（カンマ衝突） |
| gitGraph | 全コマンド基本構文 / LR・TB・BT / branch order / cherry-pick parent | commit "msg"（msg: なし）/ 複数 tag / cherry-pick tag | mainBranchOrder / acc系 | — |
| journey | title / section / タスク全構文 / アクター | # コメント | acc系 | — |
| pie | pie / showData / title / セクション / textPosition | — | 負の数値（実質影響なし）/ acc系 | — |
| quadrant | title / x-axis / y-axis / 象限ラベル / ポイント / スタイル / classDef | 複数ダッシュデリミタ / デリミタのみ形式 / Markdown埋め込み | acc系 | — |
| requirement | 全機能（direction / 全ノードタイプ / 全リレーション / 全スタイル） | — | acc系 | — |
| timeline | title / section / period / event / 継続行 / br変換 | direction LR・TD / # コメント | acc系 | — |

---

