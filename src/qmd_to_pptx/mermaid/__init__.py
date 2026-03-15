"""
qmd_to_pptx.mermaid サブパッケージ。

各ダイアグラム種別レンダラーとファサードクラスを提供する。
"""

from .renderer import MermaidRenderer

__all__ = ["MermaidRenderer"]
