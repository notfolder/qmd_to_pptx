# PyPIリリース手順

## 前提条件

- [PyPI](https://pypi.org/) アカウントを持っていること
- PyPI の API トークンを取得済みであること（Account settings → API tokens）
- `uv` がインストールされていること（`uv --version` で確認）

## リリースフロー

```mermaid
flowchart LR
    A[バージョン更新] --> B[テスト実行]
    B --> C[ビルド]
    C --> D[TestPyPI確認]
    D --> E[本番PyPIリリース]
```

## 手順

### 1. バージョンを更新する

`pyproject.toml` の `version` フィールドをセマンティックバージョニング（`MAJOR.MINOR.PATCH`）に従って更新する。

```toml
version = "0.1.4"
```

### 2. テストを実行する

すべてのテストが通過することを確認する。

```bash
uv run pytest
```

### 3. ビルドする

`dist/` ディレクトリに `.tar.gz`（sdist）と `.whl`（wheel）が生成される。

```bash
rm -rf dist/  # 古いビルド成果物を削除
uv build
```

### 4. TestPyPI で動作確認する（推奨）

本番リリース前に [TestPyPI](https://test.pypi.org/) でパッケージの見た目・インストールを確認する。

```bash
uv publish --index testpypi
```

確認後、TestPyPI からインストールして動作を検証する。

```bash
pip install --index-url https://test.pypi.org/simple/ qmd-to-pptx
```

### 5. 本番 PyPI へリリースする

```bash
uv publish
```

トークンを環境変数で渡す場合は以下のとおり。

```bash
UV_PUBLISH_TOKEN=<PyPI_API_TOKEN> uv publish
```

### 6. リリースを確認する

- PyPI のパッケージページ（`https://pypi.org/project/qmd-to-pptx/`）でバージョンが反映されていることを確認する
- `pip install qmd-to-pptx==<バージョン>` でインストールできることを確認する

## API トークンの管理

### ローカル開発：~/.pypirc に保存する

`~/.pypirc` にトークンを記述しておくと、`uv publish` 実行時に `--token` オプションを省略できる。

```ini
[distutils]
index-servers =
    pypi
    testpypi

[pypi]
username = __token__
password = pypi-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

[testpypi]
repository = https://test.pypi.org/legacy/
username = __token__
password = pypi-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
```

ファイルを作成したら、他ユーザーから読めないようにパーミッションを制限する。

```bash
chmod 600 ~/.pypirc
```

この設定後は、`--token` オプションなしでそのままリリースできる。  
`--index` にセクション名を指定することで `~/.pypirc` の認証情報が読み込まれる。

```bash
uv publish                       # 本番PyPI（[pypi] セクションを使用）
uv publish --index testpypi      # TestPyPI（[testpypi] セクションを使用）
```

> **注意**: `--publish-url` でURLを直接指定すると `~/.pypirc` のセクションが特定できず、  
> トークンの入力を求められる。TestPyPI には必ず `--index testpypi` を使うこと。

### CI/CD

リポジトリのシークレット変数に登録し、環境変数 `UV_PUBLISH_TOKEN` として参照する。

API トークンはソースコードやコミットログに含めないこと。
