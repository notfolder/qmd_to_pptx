"""
テンプレートPPTXチェッカー。

指定したPPTXテンプレートに qmd_to_pptx が利用するレイアウトが
揃っているかを検査し、不足している場合は対処方法を案内する。

使い方:
    python check_template.py <template.pptx>

例:
    python check_template.py my_template.pptx
"""

from __future__ import annotations

import sys
from pathlib import Path


# qmd_to_pptx が使用するレイアウト名と、各レイアウトに必要なプレースホルダー idx の定義
_REQUIRED_LAYOUTS: dict[str, dict] = {
    "Title Slide": {
        "description": "タイトルスライド（表紙）",
        "required_idx": [0, 1],
        "idx_roles": {0: "タイトル", 1: "サブタイトル"},
    },
    "Title and Content": {
        "description": "本文スライド（最も多用）",
        "required_idx": [0, 1],
        "idx_roles": {0: "タイトル", 1: "コンテンツ（本文）"},
    },
    "Section Header": {
        "description": "セクション区切りスライド（# 見出し）",
        "required_idx": [0],
        "idx_roles": {0: "タイトル"},
    },
    "Two Content": {
        "description": "2カラムレイアウト（:::: {.columns} ブロック）",
        "required_idx": [0, 1, 2],
        "idx_roles": {0: "タイトル", 1: "左コンテンツ", 2: "右コンテンツ"},
    },
    "Comparison": {
        "description": "図・表を含む2カラムレイアウト",
        "required_idx": [0, 1, 2],
        "idx_roles": {0: "タイトル", 1: "左コンテンツ", 2: "右コンテンツ"},
    },
    "Content with Caption": {
        "description": "テキスト＋図・表の混在スライド（左: キャプション, 右: 図）",
        "required_idx": [0, 1, 2],
        "idx_roles": {0: "タイトル", 1: "コンテンツ（図・表用）", 2: "キャプション（テキスト用）"},
    },
    "Blank": {
        "description": "--- 区切りで生成される空白スライド",
        "required_idx": [],
        "idx_roles": {},
    },
}

# フォールバックチェーン（該当レイアウトがない場合の代替）
_FALLBACK: dict[str, str] = {
    "Two Content": "Title and Content",
    "Comparison": "Title and Content",
    "Content with Caption": "Title and Content",
}


def _check_template(pptx_path: str) -> None:
    """
    PPTXテンプレートを検査して結果と対処方法を標準出力に表示する。

    Parameters
    ----------
    pptx_path : str
        テンプレートPPTXファイルのパス。
    """
    try:
        from pptx import Presentation
    except ImportError:
        print("[ERROR] python-pptx がインストールされていません。")
        print("        pip install python-pptx または uv pip install python-pptx を実行してください。")
        sys.exit(1)

    path = Path(pptx_path)
    if not path.exists():
        print(f"[ERROR] ファイルが見つかりません: {pptx_path}")
        sys.exit(1)
    if path.suffix.lower() != ".pptx":
        print(f"[ERROR] PPTX ファイルを指定してください: {pptx_path}")
        sys.exit(1)

    try:
        prs = Presentation(str(path))
    except Exception as exc:
        print(f"[ERROR] ファイルを開けませんでした: {exc}")
        sys.exit(1)

    # テンプレートのレイアウト名 → idx セットを収集する
    template_layouts: dict[str, set[int]] = {}
    for layout in prs.slide_layouts:
        idx_set = {ph.placeholder_format.idx for ph in layout.placeholders}
        template_layouts[layout.name] = idx_set

    print("=" * 60)
    print(f"テンプレート検査: {path.name}")
    print("=" * 60)
    print()

    ok_layouts: list[str] = []
    warn_layouts: list[str] = []  # フォールバックで代替可能
    ng_layouts: list[str] = []   # フォールバックも不可
    detail_msgs: list[str] = []

    for layout_name, spec in _REQUIRED_LAYOUTS.items():
        present = layout_name in template_layouts

        if present:
            # 必要な idx が揃っているか確認する
            existing_idx = template_layouts[layout_name]
            missing_idx = [
                idx for idx in spec["required_idx"]
                if idx not in existing_idx
            ]
            if not missing_idx:
                ok_layouts.append(layout_name)
            else:
                # レイアウトはあるが必要な idx が不足している
                missing_desc = ", ".join(
                    f"idx={i}（{spec['idx_roles'].get(i, '不明')}）"
                    for i in missing_idx
                )
                fallback = _FALLBACK.get(layout_name)
                if fallback and fallback in template_layouts:
                    warn_layouts.append(layout_name)
                    detail_msgs.append(
                        f"  [WARN] \"{layout_name}\"\n"
                        f"         不足プレースホルダー: {missing_desc}\n"
                        f"         → \"{fallback}\" にフォールバックして動作しますが、\n"
                        f"           レイアウト本来の配置にはなりません。\n"
                        f"           PowerPoint でプレースホルダーを追加することを推奨します。"
                    )
                else:
                    ng_layouts.append(layout_name)
                    detail_msgs.append(
                        f"  [NG]   \"{layout_name}\"\n"
                        f"         不足プレースホルダー: {missing_desc}\n"
                        f"         → PowerPoint でこのレイアウトにプレースホルダーを追加してください。"
                    )
        else:
            # レイアウト自体が存在しない
            fallback = _FALLBACK.get(layout_name)
            fallback_exists = fallback and fallback in template_layouts
            if not spec["required_idx"]:
                # Blank はプレースホルダー不要のため存在しなくてもフォールバック可
                warn_layouts.append(layout_name)
                detail_msgs.append(
                    f"  [WARN] \"{layout_name}\" が存在しません。\n"
                    f"         → スライドレイアウト[0]（{list(template_layouts.keys())[0]}）で代替されます。"
                )
            elif fallback_exists:
                warn_layouts.append(layout_name)
                detail_msgs.append(
                    f"  [WARN] \"{layout_name}\" が存在しません。\n"
                    f"         → \"{fallback}\" にフォールバックして動作しますが、\n"
                    f"           専用レイアウトを追加することを推奨します。\n"
                    f"           PowerPoint の「スライドマスター」でレイアウトを追加し、\n"
                    f"           名前を \"{layout_name}\" に設定してください。"
                )
            else:
                ng_layouts.append(layout_name)
                detail_msgs.append(
                    f"  [NG]   \"{layout_name}\" が存在しません（フォールバックも不可）。\n"
                    f"         → PowerPoint の「スライドマスター」でレイアウトを追加し、\n"
                    f"           名前を \"{layout_name}\" に設定してください。\n"
                    f"           必要プレースホルダー: "
                    + ", ".join(
                        f"idx={i}（{spec['idx_roles'].get(i, '不明')}）"
                        for i in spec["required_idx"]
                    )
                )

    # --- サマリー表示 ---
    total = len(_REQUIRED_LAYOUTS)
    print(f"チェック対象レイアウト数: {total}")
    print(f"  OK   : {len(ok_layouts)} 件")
    print(f"  WARN : {len(warn_layouts)} 件（フォールバックで動作するが推奨設定あり）")
    print(f"  NG   : {len(ng_layouts)} 件（動作に問題が生じる可能性あり）")
    print()

    # --- OK レイアウト ---
    if ok_layouts:
        print("【OK】以下のレイアウトは正常に設定されています:")
        for name in ok_layouts:
            desc = _REQUIRED_LAYOUTS[name]["description"]
            print(f"  ✓ \"{name}\"  — {desc}")
        print()

    # --- 詳細（WARN / NG）---
    if detail_msgs:
        print("【要確認】")
        for msg in detail_msgs:
            print(msg)
            print()

    # --- テンプレート内の全レイアウト一覧 ---
    print("【テンプレート内のレイアウト一覧】")
    for name, idx_set in template_layouts.items():
        sorted_idx = sorted(idx_set)
        print(f"  \"{name}\"  プレースホルダー idx: {sorted_idx}")
    print()

    # --- 総合判定 ---
    print("=" * 60)
    if ng_layouts:
        print("総合判定: NG — テンプレートを修正することを強く推奨します。")
        print("  対処方法:")
        print("  1. PowerPoint でテンプレートファイルを開く")
        print("  2. [表示] → [スライドマスター] を選択する")
        print("  3. 上記の [NG] レイアウトを追加し、名前と必要なプレースホルダーを設定する")
        print("  4. 上書き保存して閉じ、再度このスクリプトで確認する")
    elif warn_layouts:
        print("総合判定: WARN — 一部のレイアウトが最適でありません。")
        print("  フォールバックにより動作しますが、専用レイアウトを追加することで")
        print("  より意図した配置でスライドを生成できます。")
    else:
        print("総合判定: OK — すべてのレイアウトが正常に設定されています！")
    print("=" * 60)


def main() -> None:
    """エントリーポイント。コマンドライン引数を解析して検査を実行する。"""
    if len(sys.argv) != 2 or sys.argv[1] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0 if sys.argv[1:] and sys.argv[1] in ("-h", "--help") else 1)

    _check_template(sys.argv[1])


if __name__ == "__main__":
    main()
