# PyPIリリース手順

## 前提条件

- [PyPI](https://pypi.org/) アカウントを持っていること
- PyPI の API トークンを取得済みであること（Account settings → API tokens）
- `uv` がインストールされていること（`uv --version` で確認）
- `twine` が dev 依存としてインストールされていること（`uv sync` で導入される）

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
uv run twine upload --repository testpypi dist/*
```

確認後、TestPyPI からインストールして動作を検証する。

```bash
uv pip install --index-url https://test.pypi.org/simple/ qmd-to-pptx
```

### 5. 本番 PyPI へリリースする

```bash
uv run twine upload dist/*
```

### 6. リリースを確認する

- PyPI のパッケージページ（`https://pypi.org/project/qmd-to-pptx/`）でバージョンが反映されていることを確認する
- `uv pip install qmd-to-pptx==<バージョン>` でインストールできることを確認する

## API トークンの管理

`twine` は `~/.pypirc` を自動的に参照するため、トークンを都度入力する必要がない。

### ローカル開発：~/.pypirc に保存する

**1. `~/.pypirc` を作成する**

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

**2. パーミッションを制限する**

```bash
chmod 600 ~/.pypirc
```

これ以降、`uv run twine upload dist/*` を実行するだけでトークン入力なしにリリースできる。

### CI/CD

リポジトリのシークレット変数に `TWINE_USERNAME`（値: `__token__`）と `TWINE_PASSWORD`（値: トークン）を登録し、以下のように参照する。

```bash
TWINE_USERNAME=__token__ TWINE_PASSWORD=<PyPI_API_TOKEN> uv run twine upload dist/*
```

API トークンはソースコードやコミットログに含めないこと。
