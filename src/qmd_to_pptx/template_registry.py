"""
テンプレートレジストリモジュール。

config/templates.yaml（または環境変数 QMD_TO_PPTX_TEMPLATES で指定したパス）を
読み込み、テンプレートIDをPPTXファイルパスに解決する。
"""

from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Any

import yaml

# デフォルトのテンプレート設定ファイルパス（カレントディレクトリ相対）
_DEFAULT_TEMPLATES_PATH = "config/templates.yaml"
# 環境変数名
_ENV_VAR_NAME = "QMD_TO_PPTX_TEMPLATES"

logger = logging.getLogger(__name__)


class TemplateRegistry:
    """
    テンプレートレジストリクラス。

    MCPサーバー起動時にテンプレート設定を読み込み、
    テンプレートIDからPPTXファイルの絶対パスを解決する。

    テンプレート設定ファイルの読み込み順序:
        1. 環境変数 QMD_TO_PPTX_TEMPLATES に指定されたパス
        2. カレントディレクトリの config/templates.yaml

    どちらも存在しない場合は空レジストリとして動作し、エラーは発生しない。
    """

    def __init__(self) -> None:
        """
        テンプレートレジストリを初期化する。

        環境変数または既定パスからテンプレート設定を読み込む。
        """
        # 設定ファイルのパスを決定する
        config_path = self._resolve_config_path()
        # テンプレートデータを読み込む（{id: {path, description}}）
        self._templates: dict[str, dict[str, str]] = self._load(config_path)

    def resolve(self, template_id: str) -> str:
        """
        テンプレートIDに対応するPPTXファイルの絶対パスを返す。

        Parameters
        ----------
        template_id : str
            テンプレートの識別子。templates.yaml のトップレベルキー。

        Returns
        -------
        str
            テンプレートPPTXファイルの絶対パス文字列。

        Raises
        ------
        ValueError
            指定された template_id が登録されていない場合。
        """
        if template_id not in self._templates:
            available = ", ".join(self._templates.keys()) if self._templates else "（登録なし）"
            raise ValueError(
                f"テンプレートID '{template_id}' は登録されていません。"
                f"利用可能なID: {available}"
            )
        return self._templates[template_id]["path"]

    def list_templates(self) -> dict[str, str]:
        """
        登録済みテンプレートの一覧を返す。

        Returns
        -------
        dict[str, str]
            {テンプレートID: 説明文} の辞書。登録がない場合は空の辞書。
        """
        return {
            tid: entry.get("description", "（説明なし）")
            for tid, entry in self._templates.items()
        }

    def default_path(self) -> tuple[str, str] | None:
        """
        登録済みテンプレートが存在する場合、先頭エントリーの (ID, パス) を返す。

        template_id が未指定のとき、自動的に使用するデフォルトテンプレートを
        選択するために使用する。登録がない場合は None を返す。

        Returns
        -------
        tuple[str, str] | None
            (テンプレートID, PPTXファイルパス) のタプル。
            登録がない場合は None。
        """
        for tid, entry in self._templates.items():
            return (tid, entry["path"])
        return None

    # ------------------------------------------------------------------
    # 内部メソッド
    # ------------------------------------------------------------------

    @staticmethod
    def _resolve_config_path() -> Path:
        """
        テンプレート設定ファイルのパスを解決して返す。

        環境変数 QMD_TO_PPTX_TEMPLATES が設定されていればそのパス、
        未設定ならカレントディレクトリの config/templates.yaml を返す。

        Returns
        -------
        Path
            設定ファイルのパス（ファイルが存在しない場合もPathオブジェクトを返す）。
        """
        env_path = os.environ.get(_ENV_VAR_NAME)
        if env_path:
            return Path(env_path)
        return Path(_DEFAULT_TEMPLATES_PATH)

    @staticmethod
    def _load(config_path: Path) -> dict[str, dict[str, str]]:
        """
        設定ファイルを読み込んでテンプレートデータを返す。

        ファイルが存在しない場合や読み込みエラーの場合は、
        空の辞書を返してエラーログを出力する。

        Parameters
        ----------
        config_path : Path
            YAMLファイルのパス。

        Returns
        -------
        dict[str, dict[str, str]]
            {テンプレートID: {"path": ..., "description": ...}} の辞書。
        """
        if not config_path.exists():
            logger.debug(
                "テンプレート設定ファイルが見つかりません（空レジストリで動作）: %s",
                config_path,
            )
            return {}

        try:
            raw: Any = yaml.safe_load(config_path.read_text(encoding="utf-8"))
        except yaml.YAMLError as e:
            logger.warning("テンプレート設定ファイルのYAML解析に失敗しました: %s", e)
            return {}

        # templates キーがない、またはNullの場合は空レジストリとして扱う
        if not isinstance(raw, dict):
            return {}
        templates_raw = raw.get("templates")
        if not isinstance(templates_raw, dict):
            return {}

        result: dict[str, dict[str, str]] = {}
        for tid, entry in templates_raw.items():
            # path フィールドが必須
            if not isinstance(entry, dict) or "path" not in entry:
                logger.warning(
                    "テンプレート '%s' に path フィールドがありません。スキップします。",
                    tid,
                )
                continue
            result[tid] = {
                "path": str(entry["path"]),
                "description": str(entry.get("description", "（説明なし）")),
            }

        logger.info(
            "テンプレート設定を読み込みました: %s 件 (%s)",
            len(result),
            config_path,
        )
        return result
