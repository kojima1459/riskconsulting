#!/usr/bin/env python3
"""tools/dossier_check.py

docs/08_ドシエ収集プロンプト集.md の D-N コードフェンス（build/build_rpn.py の
read_dossier_prompts が配布ブックへ逐語で焼き込む唯一の値源）が、15章§2.0
ルール5b (a)〜(d) の4点を満たしているかを検査する（W14裁定書37 B-07）。

検査語の正は docs/spec/15_プロンプトとJSONスキーマ.md のルール5b (a)〜(d) の
逐語から**この関数の中だけで**抽出する。別ファイルに同じ一覧を重複して書かない
（15章§2.0 の NT_FOOTER_WORDS と同じやり方。裁定書37 班5 タスク9）。

判定:
  (a) 対象企業の本社所在地または証券コードの併記 ―― フェンスが対象企業を
      名指ししている（{{企業名}} を含む）場合のみ課す。企業を特定しない
      フェンス（業種・マクロ横断のもの）は対象外とする。
  (b) 「見当たらない」「取得できず」の両方の明記指示
  (c) 「出典URL」の明記指示
  (d) 「まとめサイト」（就活情報サイト・個人ブログを含む禁止指示）の明記

0件で exit 0。欠落があれば、フェンスごとの欠落点を一覧にして exit 1。
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SPEC15 = ROOT / "docs" / "spec" / "15_プロンプトとJSONスキーマ.md"
DOSSIER = ROOT / "docs" / "08_ドシエ収集プロンプト集.md"


class CheckError(Exception):
    pass


def load_required_words() -> dict[str, list[str]]:
    """15章ルール5b (a)〜(d) の逐語から検査語を抽出する（値源はこの1関数のみ）。"""
    if not SPEC15.exists():
        raise CheckError(f"[dossier_check] 15章が見つかりません: {SPEC15}")
    text = SPEC15.read_text(encoding="utf-8")

    m = re.search(r"5b\.(.*?)(?=\n\d[a-z]?\.\s|\Z)", text, re.S)
    if not m:
        raise CheckError("[dossier_check] 15章にルール5bが見つかりません")
    block = m.group(1)

    lines: dict[str, str] = {}
    for tag in "abcd":
        mm = re.search(rf"\({tag}\)\s*(.+?)(?=\n\s*\([a-d]\)|\Z)", block, re.S)
        if not mm:
            raise CheckError(f"[dossier_check] 15章ルール5bに({tag})が見つかりません")
        lines[tag] = re.sub(r"\s+", "", mm.group(1))

    required: dict[str, list[str]] = {}

    # (a) 「本社所在地または証券コード」→ どちらか一方の言及で足りる(OR)
    a_terms = [t for t in ("本社所在地", "証券コード") if t in lines["a"]]
    if not a_terms:
        raise CheckError("[dossier_check] 15章(a)から本社所在地/証券コードの語を抽出できません")
    required["a"] = a_terms

    # (b) 『見当たらない』『取得できず』の2語(内側の鍵括弧の中身)。両方必須(AND)
    b_terms = re.findall(r"『(.+?)』", lines["b"])
    if len(b_terms) < 2:
        raise CheckError("[dossier_check] 15章(b)から2語を抽出できません")
    required["b"] = b_terms

    # (c) 「各項目に出典URLを付けること」から、末尾の助詞を落として核となる語幹
    #     (漢字列 + URL) を1つ取り出す。
    c_quote = re.search(r"「(.+?)」", lines["c"])
    if not c_quote:
        raise CheckError("[dossier_check] 15章(c)の指示文（鍵括弧）が見つかりません")
    c_core = re.search(r"[^\s、。「」]{1,6}URL", c_quote.group(1))
    if not c_core:
        raise CheckError("[dossier_check] 15章(c)からURLを含む語を抽出できません")
    required["c"] = [c_core.group(0)]

    # (d) 「まとめサイト・就活情報サイト・個人ブログは情報源に使わないこと」の
    #     先頭項目(「・」区切りの1つ目)を1語取り出す。
    d_quote = re.search(r"「(.+?)」", lines["d"])
    if not d_quote:
        raise CheckError("[dossier_check] 15章(d)の指示文（鍵括弧）が見つかりません")
    d_first = d_quote.group(1).split("・")[0]
    if not d_first:
        raise CheckError("[dossier_check] 15章(d)から語を抽出できません")
    required["d"] = [d_first]

    return required


def load_dossier_fences() -> dict[str, str]:
    """docs/08 の `## D-N. ...` 見出し直下の最初のコードフェンスを逐語で返す
    （build/build_rpn.py の read_dossier_prompts と同一の抽出規則）。"""
    if not DOSSIER.exists():
        raise CheckError(f"[dossier_check] docs/08 が見つかりません: {DOSSIER}")
    text = DOSSIER.read_text(encoding="utf-8")

    found: dict[str, str] = {}
    heads = list(re.finditer(r"^##\s+(D-\d)\.", text, re.M))
    if not heads:
        raise CheckError("[dossier_check] docs/08 に `## D-N.` 見出しが1つもありません")
    for i, m in enumerate(heads):
        tail = heads[i + 1].start() if i + 1 < len(heads) else len(text)
        fence = re.search(r"^```[^\n]*\n(.*?)\n```", text[m.end():tail], re.S | re.M)
        if fence:
            found[m.group(1)] = fence.group(1)
    return found


def check_fence(fence_id: str, body: str, required: dict[str, list[str]]) -> list[str]:
    """1本のフェンスについて欠落点を返す（空リストなら合格）。"""
    missing: list[str] = []

    # (a) は対象企業を名指ししているフェンスにのみ課す
    if "{{企業名}}" in body:
        if not any(term in body for term in required["a"]):
            missing.append(
                "(a) 本社所在地/証券コードの併記が無い（{}のいずれか）".format(
                    "・".join(required["a"])
                )
            )

    # (b) 両方必須
    for term in required["b"]:
        if term not in body:
            missing.append(f"(b) 「{term}」の明記が無い")

    # (c) いずれか1つで可（抽出語は1つのみだが将来複数化に耐える形にしておく）
    if not any(term in body for term in required["c"]):
        missing.append("(c) 「{}」の明記が無い".format("・".join(required["c"])))

    # (d) 同上
    if not any(term in body for term in required["d"]):
        missing.append("(d) 「{}」（まとめサイト等の禁止指示）が無い".format("・".join(required["d"])))

    return missing


def main() -> int:
    try:
        required = load_required_words()
        fences = load_dossier_fences()
    except CheckError as e:
        print(str(e))
        return 1

    print("[dossier_check] 検査語（15章ルール5bから抽出）:")
    for tag in "abcd":
        print(f"  ({tag}) {required[tag]}")

    ng: dict[str, list[str]] = {}
    for fence_id in sorted(fences, key=lambda k: int(k.split("-")[1])):
        problems = check_fence(fence_id, fences[fence_id], required)
        if problems:
            ng[fence_id] = problems

    if not ng:
        print(f"[dossier_check] OK: 全{len(fences)}本のフェンスが15章ルール5bの4点を満たしています")
        return 0

    print(f"[dossier_check] NG: {len(ng)}本のフェンスで欠落があります")
    for fence_id, problems in ng.items():
        print(f"  {fence_id}:")
        for p in problems:
            print(f"    - {p}")
    return 1


if __name__ == "__main__":
    sys.exit(main())
