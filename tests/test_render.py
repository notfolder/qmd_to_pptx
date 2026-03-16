"""
render() 関数の統合テスト。
"""

import os
import tempfile
import pytest
from pathlib import Path
from qmd_to_pptx import render


class TestRender:
    """render() 関数の統合テスト。"""

    def _render_and_check(self, text: str, **kwargs) -> Path:
        """テキストをレンダリングして出力ファイルのPathを返す。"""
        with tempfile.TemporaryDirectory() as tmpdir:
            output = Path(tmpdir) / "output.pptx"
            render(text, str(output), **kwargs)
            assert output.exists(), "PPTXファイルが生成されなかった"
            assert output.stat().st_size > 0, "PPTXファイルが空"
            return output

    # --- 基本動作のテスト ---

    def test_render_simple_markdown(self) -> None:
        """シンプルなMarkdownテキストをレンダリングしてPPTXを生成する。"""
        with tempfile.TemporaryDirectory() as tmpdir:
            output = Path(tmpdir) / "output.pptx"
            render("## スライドタイトル\n\nコンテンツ", str(output))
            assert output.exists()
            assert output.stat().st_size > 0

    def test_render_qmd_with_frontmatter(self) -> None:
        """YAMLフロントマター付きQMDテキストをレンダリングする。"""
        qmd = (
            "---\ntitle: テスト\nauthor: 著者\ndate: 2026-01-01\n---\n"
            "## スライド1\n\nコンテンツ1\n"
        )
        with tempfile.TemporaryDirectory() as tmpdir:
            output = Path(tmpdir) / "output.pptx"
            render(qmd, str(output))
            assert output.exists()

    def test_render_from_file(self) -> None:
        """ファイルパスを入力として受け取りレンダリングする。"""
        with tempfile.TemporaryDirectory() as tmpdir:
            input_file = Path(tmpdir) / "input.md"
            input_file.write_text(
                "## スライドタイトル\n\nファイルからのコンテンツ", encoding="utf-8"
            )
            output = Path(tmpdir) / "output.pptx"
            render(str(input_file), str(output))
            assert output.exists()

    def test_render_text_is_not_treated_as_file(self) -> None:
        """長いテキスト文字列がファイルパスとして扱われずに正常処理される。"""
        long_text = "## スライド\n\n" + "コンテンツ " * 100
        with tempfile.TemporaryDirectory() as tmpdir:
            output = Path(tmpdir) / "output.pptx"
            render(long_text, str(output))
            assert output.exists()

    # --- スライドレイアウトのテスト ---

    def test_render_with_section_header(self) -> None:
        """# 見出しで Section Header スライドが生成される（slide-level: 2）。"""
        qmd = "---\ntitle: テスト\n---\n# セクション\n\n## スライド\n\nコンテンツ\n"
        with tempfile.TemporaryDirectory() as tmpdir:
            output = Path(tmpdir) / "output.pptx"
            render(qmd, str(output))
            assert output.exists()

    def test_render_slide_level_1(self) -> None:
        """slide-level: 1 で # のみスライド区切りとなる。"""
        qmd = (
            "---\ntitle: テスト\nslide-level: 1\n---\n"
            "# スライド1\n\n## 見出し（本文）\n\nコンテンツ\n"
            "# スライド2\n\nコンテンツ2\n"
        )
        with tempfile.TemporaryDirectory() as tmpdir:
            output = Path(tmpdir) / "output.pptx"
            render(qmd, str(output))
            assert output.exists()

    def test_render_with_list(self) -> None:
        """箇条書きリストを含むスライドを正しくレンダリングする。"""
        text = "## スライド\n\n- アイテム1\n- アイテム2\n  - サブアイテム\n"
        with tempfile.TemporaryDirectory() as tmpdir:
            output = Path(tmpdir) / "output.pptx"
            render(text, str(output))
            assert output.exists()

    def test_render_with_table(self) -> None:
        """表を含むスライドを正しくレンダリングする。"""
        text = "## スライド\n\n| 列1 | 列2 |\n|-----|-----|\n| A   | B   |\n"
        with tempfile.TemporaryDirectory() as tmpdir:
            output = Path(tmpdir) / "output.pptx"
            render(text, str(output))
            assert output.exists()

    def test_render_with_mermaid(self) -> None:
        """Mermaid図を含むスライドを正しくレンダリングする。"""
        text = "## スライド\n\n```mermaid\nflowchart LR\n    A --> B\n```\n"
        with tempfile.TemporaryDirectory() as tmpdir:
            output = Path(tmpdir) / "output.pptx"
            render(text, str(output))
            assert output.exists()

    def test_render_with_speaker_notes(self) -> None:
        """スピーカーノートを含むスライドを正しくレンダリングする。"""
        text = (
            "## スライド\n\nコンテンツ\n\n"
            '::: {.notes}\nスピーカーノートです。\n:::\n'
        )
        with tempfile.TemporaryDirectory() as tmpdir:
            output = Path(tmpdir) / "output.pptx"
            render(text, str(output))
            assert output.exists()

    def test_render_blank_slide(self) -> None:
        """コンテンツなし（Blank）スライドを正しくレンダリングする。"""
        text = "## スライド\n"
        with tempfile.TemporaryDirectory() as tmpdir:
            output = Path(tmpdir) / "output.pptx"
            render(text, str(output))
            assert output.exists()

    def test_render_with_horizontal_rule_slide(self) -> None:
        """水平区切り線でのスライド分割を正しく処理する。"""
        text = (
            "---\ntitle: テスト\n---\n"
            "## スライド1\n\nコンテンツ1\n\n---\n\nタイトルなしコンテンツ\n"
        )
        with tempfile.TemporaryDirectory() as tmpdir:
            output = Path(tmpdir) / "output.pptx"
            render(text, str(output))
            assert output.exists()

    def test_render_quarto_mermaid_syntax(self) -> None:
        """Quartoのネイティブ Mermaid 記法（```{mermaid}）を正しく処理する。"""
        text = (
            "## スライド\n\n"
            "```{mermaid}\nflowchart LR\n    A --> B\n```\n"
        )
        with tempfile.TemporaryDirectory() as tmpdir:
            output = Path(tmpdir) / "output.pptx"
            render(text, str(output))
            assert output.exists()

    def test_render_quarto_code_block_syntax(self) -> None:
        """Quartoのコードブロック記法（```{python}）を正しく処理する。"""
        text = (
            "## スライド\n\n"
            "```{python}\nprint('hello')\n```\n"
        )
        with tempfile.TemporaryDirectory() as tmpdir:
            output = Path(tmpdir) / "output.pptx"
            render(text, str(output))
            assert output.exists()

    def test_render_incremental_true(self) -> None:
        """incremental: true を指定したときに例外なくレンダリングが完了する。"""
        qmd = (
            "---\ntitle: テスト\nformat:\n  pptx:\n    incremental: true\n---\n"
            "## スライド\n\n- アイテム1\n- アイテム2\n- アイテム3\n"
        )
        with tempfile.TemporaryDirectory() as tmpdir:
            output = Path(tmpdir) / "output.pptx"
            render(qmd, str(output))
            assert output.exists()
            assert output.stat().st_size > 0

    def test_render_ignores_nonexistent_yaml_reference_doc(self) -> None:
        """YAMLのreference-docに存在しないパスを指定した場合、無視してデフォルトで生成する。"""
        # YAMLに存在しないファイルパスを reference-doc として設定する
        qmd = (
            "---\ntitle: テスト\nformat:\n  pptx:\n    reference-doc: 'nonexistent.pptx'\n---\n"
            "## スライド\n\nコンテンツ\n"
        )
        with tempfile.TemporaryDirectory() as tmpdir:
            output = Path(tmpdir) / "output.pptx"
            # 存在しないパスは内部で無効化されデフォルトPresentationを使用する
            render(qmd, str(output), reference_doc=None)
            assert output.exists()
            assert output.stat().st_size > 0
