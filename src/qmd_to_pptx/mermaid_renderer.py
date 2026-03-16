"""
Mermaidレンダラーモジュール（後方互換のための再エクスポート）。

実装は qmd_to_pptx.mermaid サブパッケージに移動した。
このモジュールは slide_renderer.py 等からの既存 import を維持するために残す。
"""

from qmd_to_pptx.mermaid import MermaidRenderer

__all__ = ["MermaidRenderer"]
