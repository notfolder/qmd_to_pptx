"""
前処理器（Preprocessor）の単体テスト。
"""

import pytest
from qmd_to_pptx.preprocessor import Preprocessor


class TestPreprocessor:
    """Preprocessor クラスの単体テスト。"""

    def setup_method(self) -> None:
        """テスト前にPreprocessorインスタンスを生成する。"""
        self.preprocessor = Preprocessor()

    # --- _ensure_frontmatter のテスト ---

    def test_ensure_frontmatter_adds_when_missing(self) -> None:
        """YAMLフロントマターが存在しない場合、空のフロントマターを付与する。"""
        text = "## スライドタイトル\n\nコンテンツ"
        result = self.preprocessor.normalize(text)
        assert result.startswith("---\ntitle: ''\nauthor: ''\ndate: ''\n---\n")

    def test_ensure_frontmatter_keeps_existing(self) -> None:
        """既存のYAMLフロントマターはそのまま維持する。"""
        text = "---\ntitle: テスト\n---\n## スライド\n"
        result = self.preprocessor.normalize(text)
        assert result.startswith("---\ntitle: テスト\n---\n")
        # 余分なフロントマターが付与されないことを確認
        assert result.count("---") == 2

    def test_ensure_frontmatter_not_added_when_starts_with_dash(self) -> None:
        """先頭が ---\\n で始まる場合はフロントマターを付与しない。"""
        text = "---\ntitle: 既存タイトル\nauthor: 著者\ndate: 2026-01-01\n---\n本文"
        result = self.preprocessor.normalize(text)
        assert result.count("---\ntitle:") == 1

    # --- _normalize_mermaid のテスト ---

    def test_normalize_mermaid_converts_quarto_syntax(self) -> None:
        """Quartoの```{mermaid}形式を```mermaid形式に置換する。"""
        text = "---\ntitle: ''\nauthor: ''\ndate: ''\n---\n```{mermaid}\nflowchart LR\n  A --> B\n```\n"
        result = self.preprocessor.normalize(text)
        assert "```mermaid" in result
        assert "```{mermaid}" not in result

    def test_normalize_mermaid_keeps_standard_syntax(self) -> None:
        """既に```mermaid形式になっているブロックはそのまま維持する。"""
        text = "---\ntitle: ''\nauthor: ''\ndate: ''\n---\n```mermaid\nflowchart LR\n  A --> B\n```\n"
        result = self.preprocessor.normalize(text)
        assert result.count("```mermaid") == 1

    # --- _normalize_code_blocks のテスト ---

    def test_normalize_code_blocks_converts_braces_syntax(self) -> None:
        """```{python}形式を```python形式に置換する。"""
        text = "---\ntitle: ''\nauthor: ''\ndate: ''\n---\n```{python}\nx = 1\n```\n"
        result = self.preprocessor.normalize(text)
        assert "```python" in result
        assert "```{python}" not in result

    def test_normalize_code_blocks_converts_dotted_syntax(self) -> None:
        """```{.r}形式を```r形式に置換する。"""
        text = "---\ntitle: ''\nauthor: ''\ndate: ''\n---\n```{.r}\nplot(1)\n```\n"
        result = self.preprocessor.normalize(text)
        assert "```r" in result
        assert "```{.r}" not in result

    def test_normalize_code_blocks_does_not_affect_mermaid(self) -> None:
        """```{mermaid}はmermaid変換ステップで処理済みのため、コードブロック変換後も正しい形式を維持する。"""
        text = "---\ntitle: ''\nauthor: ''\ndate: ''\n---\n```{mermaid}\nflowchart LR\n  A --> B\n```\n"
        result = self.preprocessor.normalize(text)
        assert "```mermaid" in result
        assert "```{mermaid}" not in result

    def test_normalize_full_pipeline(self) -> None:
        """フロントマター補完・Mermaid変換・コードブロック変換を一括で実行する。"""
        text = "## タイトル\n\n```{python}\nx = 1\n```\n\n```{mermaid}\nA --> B\n```\n"
        result = self.preprocessor.normalize(text)
        assert result.startswith("---\n")
        assert "```python" in result
        assert "```mermaid" in result
        assert "```{python}" not in result
        assert "```{mermaid}" not in result
