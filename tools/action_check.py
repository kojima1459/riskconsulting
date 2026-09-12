#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""action_check.py - action名の3点照合(裁定書38 班D §1(4)。ui_check の2点一致の拡張)

================================================================================
なぜ要るか:
    `tools/ui_check.py` は「JSが送るaction名」⇔「modNaviHost.IsAllowed が
    許可する集合」の2点一致だけを見ている。本ツールはそこへ3点目
    ――**11章§0.0b の action 一覧**――を足す。11章はHTML画面の設計の正
    であり、ここに載っている action が実装(Dispatch)からもJSからも
    無くなっていたら、それは「設計は書いたのに実装が追随していない」
    (文書の嘘の裏返し=実装の嘘)である。

3点の関係(非対称。設計上重要):
    (1) JS ⇔ Dispatch: 完全一致(対称)。片方にしかない action は ERROR。
        例外は VBA_ONLY_ALLOWED に理由つきで書いたものだけ
        (`tools/ui_check.py` の同名リストと同じ思想・同じ1件
        `save_step_edit` = 第2段の予約口。用意はあるが今は断る)。
    (2) 11章§0.0b ⇔ (JS ∩ Dispatch): **部分集合**関係だけを見る(非対称)。
        11章の表自身が「下表は**主な** action」と明記しており、上級区画の
        action(受信箱・判断台帳・商談記録・エクスポート等 約18本)は
        個別に列挙されていない。これは文書の不備ではなく明示された設計
        (主要導線だけを表にする)なので、「表に無い」ことをERRORにすると
        誤検知になる(伝書鳩1-3: 検査が正しい書き方を叱ると検査は無視される)。
        ERRORにするのは逆方向だけ: **11章が載せている action が実装に
        存在しない**(=設計だけ書いて実装しなかった、または実装が
        リネームされて追随しなかった)場合。

対象:
    JS   : ui/app.js・ui/views.js・ui/index.html の send()/data-action
    VBA  : src/ui/navi/modNaviActions.bas の Dispatch と
           src/ui/navi/modNaviActions2.bas の DispatchMore の Case 文字列
    仕様 : docs/spec/11_画面設計.md の `## 0.0b` 表の「action」列
           (バッククォート区切りで複数個/セルの形を分解する)

使い方:
    python3 tools/action_check.py
    python3 tools/action_check.py --verbose
    exit code: 0 = ERROR 0件 / 1 = ERROR 1件以上 / 2 = 自己テスト失敗
================================================================================
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

TOOLS_DIR = Path(__file__).resolve().parent
REPO_ROOT = TOOLS_DIR.parent
sys.path.insert(0, str(TOOLS_DIR))

import ui_check  # noqa: E402 (既存のJS action抽出ヘルパを再利用)

UI_DIR = REPO_ROOT / "ui"
NAVI_DIR = REPO_ROOT / "src" / "ui" / "navi"
DISPATCH_PATH = NAVI_DIR / "modNaviActions.bas"
DISPATCH_MORE_PATH = NAVI_DIR / "modNaviActions2.bas"
SPEC11_PATH = REPO_ROOT / "docs" / "spec" / "11_画面設計.md"

CASE_LITERAL = re.compile(r'"([a-z][a-z0-9_]*)"')
DISPATCH_FUNC = re.compile(
    r"Public\s+Function\s+Dispatch\b.*?End\s+Function", re.DOTALL | re.IGNORECASE)
DISPATCH_MORE_FUNC = re.compile(
    r"Public\s+Function\s+DispatchMore\b.*?End\s+Function", re.DOTALL | re.IGNORECASE)

# ui_check.py の VBA_ONLY_ALLOWED と同じ思想・同じ1件(save_step_edit=第2段の
# 予約口。JSから送られなくても Dispatch にはあってよい)。
VBA_ONLY_ALLOWED = set(ui_check.VBA_ONLY_ALLOWED)

# 0.0b 表の「action」列だけを対象にする(見出し行の位置で区切る)。
SPEC11_SECTION = re.compile(
    r"^##\s*0\.0b\b.*?(?=^##\s|\Z)", re.DOTALL | re.MULTILINE)
TABLE_ROW = re.compile(r"^\|(.+?)\|(.+?)\|(.+?)\|\s*$", re.MULTILINE)
BACKTICK_TOKEN = re.compile(r"`([a-z][a-z0-9_]*)`")


def read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def dispatch_actions() -> set[str]:
    out: set[str] = set()
    if DISPATCH_PATH.exists():
        m = DISPATCH_FUNC.search(read(DISPATCH_PATH))
        if m:
            for line in m.group(0).split("\n"):
                if re.match(r"\s*Case\b", line, re.IGNORECASE):
                    out |= set(CASE_LITERAL.findall(line))
    if DISPATCH_MORE_PATH.exists():
        m = DISPATCH_MORE_FUNC.search(read(DISPATCH_MORE_PATH))
        if m:
            for line in m.group(0).split("\n"):
                if re.match(r"\s*Case\b", line, re.IGNORECASE):
                    out |= set(CASE_LITERAL.findall(line))
    return out


def spec11_actions(text: str) -> set[str]:
    """11章 `## 0.0b` の表の「action」列(1列目)からだけ action 名を拾う。

    表の2列目(画面部品)・3列目(既存関数)にも `関数名` のバッククォートが
    大量にあるため、**1列目に限定しないと画面部品名や関数名を action と
    誤認する**(伝書鳩1-3の「部分一致の罠」と同型の事故)。
    """
    sec = SPEC11_SECTION.search(text)
    if not sec:
        return set()
    out: set[str] = set()
    for m in TABLE_ROW.finditer(sec.group(0)):
        col1 = m.group(1)
        stripped = col1.strip()
        # ヘッダ行("action")・区切り行("---")だけを除く。**部分一致にしない**
        # (実測で発見した誤検知: action名 `totally_fake_action` のように
        # 語尾に "action" を含む行が、ヘッダ行の部分一致で丸ごと無視され、
        # 検査が骨抜きになった。伝書鳩1-3の「部分一致の罠」と同型の事故)。
        if stripped == "action" or set(stripped) <= set("-"):
            continue
        out |= set(BACKTICK_TOKEN.findall(col1))
    return out


def run_checks(verbose: bool) -> int:
    errors = 0

    def err(msg: str) -> None:
        nonlocal errors
        errors += 1
        print("ERROR " + msg)

    if not UI_DIR.exists():
        print("ERROR ui/ がありません(検査不能)")
        return 1
    js_set: set[str] = set()
    for name in ("app.js", "views.js", "index.html"):
        path = UI_DIR / name
        if path.exists():
            js_set |= ui_check.js_actions(read(path))

    if not DISPATCH_PATH.exists() or not DISPATCH_MORE_PATH.exists():
        err("(1) Dispatch/DispatchMore のファイルが見つかりません")
        disp_set: set[str] = set()
    else:
        disp_set = dispatch_actions()
        if not disp_set:
            err("(1) modNaviActions.Dispatch の Case を読み取れませんでした")

    only_js = sorted(js_set - disp_set)
    only_disp = sorted(disp_set - js_set - VBA_ONLY_ALLOWED)
    for a in only_js:
        err('(1) 画面が送るのに Dispatch が受けない action: "%s"'
            "(modNaviActions.Dispatch へ Case を足すか、画面側の綴りを直す)" % a)
    for a in only_disp:
        err('(1) Dispatch が受けるのに画面が送らない action: "%s"'
            "(画面へ配線するか、tools/action_check.py の VBA_ONLY_ALLOWED へ"
            "理由つきで書く)" % a)
    print("  (1) JS ⇔ Dispatch               JS %d / Dispatch %d / 例外 %d / "
          "不一致 %d" % (len(js_set), len(disp_set), len(VBA_ONLY_ALLOWED),
                       len(only_js) + len(only_disp)))

    if not SPEC11_PATH.exists():
        err("(2) %s がありません" % SPEC11_PATH.relative_to(REPO_ROOT))
        spec_set: set[str] = set()
    else:
        spec_set = spec11_actions(read(SPEC11_PATH))
        if not spec_set:
            err("(2) 11章 `## 0.0b` の action 表を読み取れませんでした"
                "(見出しが動いた可能性。SPEC11_SECTION を確認してください)")
    implemented = (js_set | VBA_ONLY_ALLOWED) & disp_set
    doc_only = sorted(spec_set - implemented)
    for a in doc_only:
        err('(2) 11章 `## 0.0b` に載っている action "%s" が実装(JS/Dispatch)に'
            "存在しません(設計だけ書いて実装が追随していないか、実装側が"
            "リネームされました)" % a)
    print("  (2) 11章⊆(JS∩Dispatch)           11章 %d件(主な action の"
          "抜粋。網羅ではないため逆方向〔実装にあって11章に無い〕は"
          "検査しない) / 不一致 %d" % (len(spec_set), len(doc_only)))
    if verbose:
        print("      JS      : " + ", ".join(sorted(js_set)))
        print("      Dispatch: " + ", ".join(sorted(disp_set)))
        print("      11章    : " + ", ".join(sorted(spec_set)))

    return errors


# ---------------------------------------------------------------------------
# 自己テスト(骨抜き防止)
# ---------------------------------------------------------------------------
def self_test() -> bool:
    cases = []

    disp = dispatch_actions()
    cases.append(("Dispatch抽出 非空", len(disp) > 10))
    cases.append(("Dispatch抽出に initialize を含む", "initialize" in disp))
    cases.append(("Dispatch抽出に inbox_judge を含む(2ファイル目)",
                  "inbox_judge" in disp))

    sample = (
        "## 0.0b HTML画面の action と画面部品の対応(v4.0)\n"
        "\n"
        "| action | 画面部品 | 対応する既存関数(仮置き) |\n"
        "|---|---|---|\n"
        "| `initialize` | 画面起動時 | `ResolveTransport` |\n"
        "| `open_case` / `new_case` | 案件セレクト | `ReadCaseCtx` |\n"
        "\n"
        "---\n"
        "\n"
        "## 0. 体験原則\n"
        "| `not_an_action` | 別の節の表 | `SomeFunc` |\n"
    )
    parsed = spec11_actions(sample)
    cases.append(("11章表抽出 1列目のみ", parsed == {"initialize", "open_case", "new_case"}))
    cases.append(("11章表抽出 2列目3列目は拾わない",
                  "ResolveTransport" not in parsed and "ReadCaseCtx" not in parsed))
    cases.append(("11章表抽出 セクション境界を超えない",
                  "not_an_action" not in parsed))
    # 部分一致の罠(実測で発見): action名の語尾に "action" が含まれる行を
    # ヘッダ行と誤認して丸ごと落とさないこと。
    trap = (
        "## 0.0b HTML画面の action と画面部品の対応(v4.0)\n"
        "\n"
        "| action | 画面部品 | 対応する既存関数(仮置き) |\n"
        "|---|---|---|\n"
        "| `resize` / `close` / `weird_action` | X | Y |\n"
    )
    cases.append(("11章表抽出 部分一致の罠を踏まない",
                  {"resize", "close", "weird_action"} <= spec11_actions(trap)))

    bad = [name for name, ok in cases if not ok]
    for name in bad:
        print("  自己テスト NG: %s" % name)
    print("  自己テスト: %d/%d" % (len(cases) - len(bad), len(cases)))
    return not bad


def main() -> int:
    ap = argparse.ArgumentParser(description="action名の3点照合(裁定書38 班D)")
    ap.add_argument("--verbose", action="store_true",
                    help="照合したaction名の集合を全部出す")
    args = ap.parse_args()

    print("action_check: action名の3点照合(裁定書38 班D §1(4))")
    errors = run_checks(args.verbose)

    if not self_test():
        print("結果: 自己テスト失敗(検出器が壊れています)")
        return 2

    if errors:
        print("ERROR: %d 件" % errors)
        print("結果: NG")
        return 1
    print("結果: OK 2条件(JS⇔Dispatch完全一致 / 11章⊆実装)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
