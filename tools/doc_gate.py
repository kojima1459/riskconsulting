#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""doc_gate.py - 文書検問(裁定書38 班D §1(3)。伝書鳩20260912 Part1-3の移植)

================================================================================
なぜ要るか(伝書鳩1-3):
    検問はコードだけを守り、文書を1文字も検査していなかった。配布文書が
    古いままだと、テスターの報告に「本当のバグ」と「文書が古いだけ」が
    混ざり、実機テストのループが終わらない(司令塔が毎回人力で切り分ける
    羽目になる)。**見つかった文書の嘘はdocs側を直す**(コード側の疑いは
    concerns)。

検査対象: docs/24・25・26・README.md のみ。
対象外(誤検知の温床。伝書鳩1-3「適用範囲の設計」がそのまま正):
    - `docs/spec_*` / `spec/` 配下の仕様章(正はspecであり、本ツールが
      仕様を裁定してはいけない)
    - `audit_*`(監査記録)
    - 裁定書(ファイル名・見出しに「裁定書」を含む)
    - ファイル名に日付(8桁の数字)を持つ記録文書
    - `docs/受領/` 配下(社外から受け取った原文書)

4検査:
    (1) 「既定N」形式の既定値 vs sheets_main.json の config 既定値
        (1行1キー限定。1行に複数キーが並ぶ行は総当たりせずSKIPする=
        伝書鳩1-3が挙げた「1行複数キーの罠」の回避)
    (2) `[ボタン名]` 表記が ui/index.html・ui/views.js・ui/app.js・
        build/sheets_main.json のキャプション文字列・src/ui/modUI*.bas の
        文字列リテラルのいずれかに**完全一致**で存在する(部分一致にしない。
        伝書鳩1-3が挙げた「正しい長い名前を禁止語の部分一致で赤くする」を
        避けるため、判定は「存在するか」だけで、逆方向〔存在しない語を
        禁止するブロックリスト〕は持たない)
    (3) src が吐くエラーコード `E0\\d{3}` ⊆ 16章§1の表 ⊆ docs/25 が
        本文中に書くエラーコード(一方向の部分集合。docs/25は全コードを
        網羅する文書ではないため、逆方向〔16章にあってdocs/25に無い〕は
        検査しない=誤検知の温床)
    (4) 対象外ファイルが誤って検査対象へ混ざっていないことの自己確認

使い方:
    python3 tools/doc_gate.py
    python3 tools/doc_gate.py --verbose
    exit code: 0 = ERROR 0件 / 1 = ERROR 1件以上 / 2 = 自己テスト失敗
================================================================================
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

TOOLS_DIR = Path(__file__).resolve().parent
REPO_ROOT = TOOLS_DIR.parent

TARGET_DOCS = ("docs/24_実機テスト手順書_Windows.md",
               "docs/25_利用ガイド.md",
               "docs/26_はじめてガイド_調査からレポートまで.md",
               "README.md")

SHEETS_JSON = REPO_ROOT / "build" / "sheets_main.json"
UI_DIR = REPO_ROOT / "ui"
SPEC16 = REPO_ROOT / "docs" / "spec" / "16_エッジケースとセキュリティ設計.md"

# ---------------------------------------------------------------------------
# 対象外の判定(伝書鳩1-3「適用範囲の設計」)。
# ---------------------------------------------------------------------------
DATED_FILENAME = re.compile(r"\d{8}")


def is_excluded(rel_path: str) -> bool:
    name = Path(rel_path).name
    if rel_path.startswith("docs/spec/") or rel_path.startswith("docs/受領/"):
        return True
    if name.startswith("spec_") or name.startswith("audit_"):
        return True
    if "裁定書" in name:
        return True
    if DATED_FILENAME.search(name):
        return True
    return False


# ---------------------------------------------------------------------------
# (1) 既定値
# ---------------------------------------------------------------------------
TABLE_ROW = re.compile(r"^\|(.+?)\|(.*)\|\s*$", re.MULTILINE)
BACKTICK_TOKEN = re.compile(r"`([a-zA-Z_][a-zA-Z0-9_]*)`")
DEFAULT_VALUE = re.compile(
    r"既定(?:は空(?!\S)|\"([^\"]*)\"|([A-Z]{2,})(?![a-zA-Z])|([0-9]+))")


def load_config_defaults() -> dict[str, object]:
    d = json.loads(SHEETS_JSON.read_text(encoding="utf-8"))
    for s in d.get("sheets", []):
        if s.get("name") == "config":
            return {row["name"]: row["value"] for row in s["defaults"]}
    return {}


def check_default_values(text: str, config_defaults: dict[str, object],
                         known_keys: set[str]) -> list[tuple[str, str, str]]:
    """戻り値: (キー, 文書記載, build実値) の不一致リスト。"""
    mismatches = []
    for row in TABLE_ROW.finditer(text):
        col1, rest = row.group(1), row.group(2)
        keys = BACKTICK_TOKEN.findall(col1)
        keys = [k for k in keys if k in known_keys]
        if len(keys) != 1:
            continue  # 1行複数キーの罠を避ける(0件・複数件はSKIP)
        key = keys[0]
        m = DEFAULT_VALUE.search(rest)
        if not m:
            continue
        if "既定は空" in rest:
            doc_val = ""
        else:
            doc_val = m.group(1) if m.group(1) is not None else (
                m.group(2) if m.group(2) is not None else m.group(3))
        real = config_defaults.get(key)
        real_str = "" if real is None else (
            "TRUE" if real is True else "FALSE" if real is False else str(real))
        if doc_val != real_str:
            mismatches.append((key, doc_val, real_str))
    return mismatches


# ---------------------------------------------------------------------------
# (2) ボタン名・タブ名
# ---------------------------------------------------------------------------
# `[名前]`(マークダウンリンク `[text](url)` は除外)。太字強調 `**[名前]**` の
# `**` は名前に含めない。
BRACKET_NAME = re.compile(r"(?<!\!)\[([^\[\]\n]{1,40})\](?!\()")
TAB_NAME = re.compile(r"「([^「」\n]{1,20})」タブ")

# 実測で判明した誤検知1: Windows/Excel自身のメニュー・タブ名(このアプリの
# ボタンではない)。docs/24(実機テスト手順書)はOSやExcel本体の操作手順も
# `[…]`/「…」タブ 表記で書くため、機械的な括弧一致だけでは区別できない。
# 該当箇所を1件ずつ確認して明示した(伝書鳩1-3「誤検知は1件も残さない」)。
#   - `[再表示]`: シート見出し右クリックメニュー(docs/24:146,584)
#   - `[フィルター]`: Excel [データ]タブの機能(docs/24:267)
#   - `[名前を付けて保存]`: Excelの[ファイル]メニュー(docs/24:440・554)
#   - `[オプション]`: Excel本体の設定画面(docs/24:566)
#   - 「全般」タブ: Windowsのファイルのプロパティダイアログ(docs/24:64・107)
NATIVE_OS_UI_ALLOWLIST = {"再表示", "フィルター", "名前を付けて保存", "オプション"}
NATIVE_OS_TAB_ALLOWLIST = {"全般"}

# 実測で判明した誤検知2: README.md冒頭の「現況」段落は
# 「以下の段落は当時の記録＝履歴であり」と明記された**過去の実装状態の
# 記録**である(裁定書の「旧値が書いてあるのが正しい」と同じ型。ファイル
# 単位の除外規則では拾えないため、段落単位でこの目印を見る)。
#
# 実測で判明した誤検知3: docs/24 冒頭の「vX.Y変更概要」段落も同型
# (例: 「v4.0変更概要(...): 本書は旧タブ・旧ボタン名([① 調べる指示文を
# 出す]〜…)の記述を全面削除し、実装から1字単位で引用し直した」)。
# 変更概要は**「何を何から何へ変えたか」を書く場**であり、旧い名前が
# 出てくるのは正しい記録である(裁定書の「旧値」と同じ理由。
# 伝書鳩1-3「適用範囲の設計」)。段落の先頭が `vX.Y変更概要` で始まる形を
# 汎用パターンとして併せて除外する。
HISTORICAL_PARAGRAPH_MARKER = "当時の記録"
CHANGELOG_PARAGRAPH_RE = re.compile(r"^v[\d.]+\s*(?:変更概要|系)")


def extract_ui_names(texts: dict) -> str:
    """突合対象の全テキストを1本に連結する。"""
    chunks = list(texts.values())
    d = json.loads(SHEETS_JSON.read_text(encoding="utf-8"))
    chunks.append(json.dumps(d, ensure_ascii=False))
    for path in sorted((REPO_ROOT / "src" / "ui").glob("modUI*.bas")):
        chunks.append(path.read_text(encoding="utf-8", errors="ignore"))
    return "\n".join(chunks)


def _strip_historical_paragraphs(text: str) -> str:
    """履歴段落(空行区切り)を検査対象から外す。

    2形: (a) `当時の記録` を含む段落、(b) `vX.Y変更概要` で始まる段落
    (changelogは「何を何から何へ変えたか」を書く場であり、旧名が出るのが
    正しい記録=裁定書の「旧値」と同じ理由)。
    """
    paras = re.split(r"\n\s*\n", text)
    out = []
    for p in paras:
        if HISTORICAL_PARAGRAPH_MARKER in p:
            continue
        if CHANGELOG_PARAGRAPH_RE.match(p.strip()):
            continue
        out.append(p)
    return "\n\n".join(out)


def check_bracket_names(text: str, haystack: str) -> list[str]:
    text = _strip_historical_paragraphs(text)
    missing = []
    for m in BRACKET_NAME.finditer(text):
        name = m.group(1).strip("*").strip()
        if not name or name.isdigit():
            continue
        if name in NATIVE_OS_UI_ALLOWLIST:
            continue
        if name not in haystack:
            missing.append(name)
    return missing


def check_tab_names(text: str, haystack: str) -> list[str]:
    text = _strip_historical_paragraphs(text)
    missing = []
    for m in TAB_NAME.finditer(text):
        name = m.group(1)
        if name in NATIVE_OS_TAB_ALLOWLIST:
            continue
        if name not in haystack:
            missing.append(name)
    return missing


# ---------------------------------------------------------------------------
# (3) エラーコード
# ---------------------------------------------------------------------------
ERROR_CODE = re.compile(r"E0\d{3}")


def codes_in(text: str) -> set[str]:
    return set(ERROR_CODE.findall(text))


def spec16_section1(text: str) -> str:
    m = re.search(r"^## 1\..*?(?=^## 2\.)", text, re.DOTALL | re.MULTILINE)
    return m.group(0) if m else text


# ---------------------------------------------------------------------------
# 本検査
# ---------------------------------------------------------------------------
def run_checks(verbose: bool) -> int:
    errors = 0

    def err(msg: str) -> None:
        nonlocal errors
        errors += 1
        print("ERROR " + msg)

    docs_texts: dict[str, str] = {}
    for rel in TARGET_DOCS:
        path = REPO_ROOT / rel
        if is_excluded(rel):
            err("(4) 対象外指定のはずのファイルが対象一覧に混ざっています: %s" % rel)
            continue
        if not path.exists():
            err("対象文書がありません: %s" % rel)
            continue
        docs_texts[rel] = path.read_text(encoding="utf-8")

    # 自己確認(4): 除外規則が実際に効くか(README以下は対象外にならないこと・
    # docs/spec配下は対象外になること)を空撃ちで見る。
    n4 = 0
    if not is_excluded("docs/spec/13_データ設計.md"):
        err("(4) docs/spec/ 配下が除外されていません(設計)")
        n4 += 1
    if not is_excluded("docs/受領/髙橋_実行フロー解説資料_v1.0.md"):
        err("(4) docs/受領/ 配下が除外されていません(設計)")
        n4 += 1
    if not is_excluded("docs/22_実装前監査記録_20260829.md"):
        err("(4) 日付付きファイル名が除外されていません(設計)")
        n4 += 1
    if is_excluded("docs/25_利用ガイド.md"):
        err("(4) 対象文書 docs/25 が誤って除外されています(設計)")
        n4 += 1
    print("  (4) 対象外判定の自己確認         不一致 %d件" % n4)

    if not SHEETS_JSON.exists():
        err("build/sheets_main.json がありません")
        return errors
    config_defaults = load_config_defaults()
    known_keys = set(config_defaults)

    n1 = 0
    for rel, text in docs_texts.items():
        for key, doc_val, real_val in check_default_values(text, config_defaults, known_keys):
            err('(1) %s: `%s` の「既定%s」が build 実値「%s」と食い違います'
                % (rel, key, doc_val, real_val))
            n1 += 1
    print("  (1) 既定値の食い違い             %d件" % n1)

    haystack = extract_ui_names(docs_texts if False else {
        "index.html": (UI_DIR / "index.html").read_text(encoding="utf-8")
        if (UI_DIR / "index.html").exists() else "",
        "views.js": (UI_DIR / "views.js").read_text(encoding="utf-8")
        if (UI_DIR / "views.js").exists() else "",
        "app.js": (UI_DIR / "app.js").read_text(encoding="utf-8")
        if (UI_DIR / "app.js").exists() else "",
    })
    n2 = 0
    for rel, text in docs_texts.items():
        for name in check_bracket_names(text, haystack):
            err('(2) %s: `[%s]` が画面(ui/・sheets_main.json・modUI*.bas)の'
                "どこにも見つかりません" % (rel, name))
            n2 += 1
        for name in check_tab_names(text, haystack):
            err('(2) %s: 「%s」タブ が画面のどこにも見つかりません' % (rel, name))
            n2 += 1
    print("  (2) ボタン名・タブ名の不在        %d件" % n2)

    src_codes: set[str] = set()
    for path in sorted((REPO_ROOT / "src").rglob("*.bas")):
        src_codes |= codes_in(path.read_text(encoding="utf-8", errors="ignore"))
    spec16_codes: set[str] = set()
    if SPEC16.exists():
        spec16_codes = codes_in(spec16_section1(SPEC16.read_text(encoding="utf-8")))
    n3 = 0
    only_src = sorted(src_codes - spec16_codes)
    for code in only_src:
        err("(3) src が使うエラーコード %s が 16章§1の表に載っていません" % code)
        n3 += 1
    for rel, text in docs_texts.items():
        for code in sorted(codes_in(text) - spec16_codes):
            err("(3) %s が書くエラーコード %s が 16章§1の表に載っていません"
                "(docs側の誤記の疑い)" % (rel, code))
            n3 += 1
    print("  (3) エラーコード ⊆16章§1         src %d件・16章 %d件 / 不一致 %d件"
          % (len(src_codes), len(spec16_codes), n3))
    if verbose:
        print("      src codes  : " + ", ".join(sorted(src_codes)))
        print("      16章 codes : " + ", ".join(sorted(spec16_codes)))

    return errors


# ---------------------------------------------------------------------------
# 自己テスト(骨抜き防止)。伝書鳩1-3が指摘した2つの誤検知パターンを
# 具体的に踏まないことを検証する。
# ---------------------------------------------------------------------------
def self_test() -> bool:
    cases = []

    # (1) 1行複数キーの罠: 複数キーが並ぶ行は既定値照合をSKIPする。
    multi_row_text = "| `direct_model` / `direct_api_base` | gpt-4.1 / url | 既定gpt-4.1 |\n"
    cases.append(("1行複数キーの罠回避",
                  check_default_values(multi_row_text, {"direct_model": "gpt-4.1"},
                                       {"direct_model", "direct_api_base"}) == []))
    # 正例: 1行1キーなら検出できる。
    single_row_ok = "| `llm_wait_sec` | 必要なら | 既定1200 |\n"
    cases.append(("既定値 正例(一致)",
                  check_default_values(single_row_ok, {"llm_wait_sec": 1200},
                                       {"llm_wait_sec"}) == []))
    single_row_bad = "| `llm_wait_sec` | 必要なら | 既定900 |\n"
    cases.append(("既定値 負例(不一致を検出)",
                  check_default_values(single_row_bad, {"llm_wait_sec": 1200},
                                       {"llm_wait_sec"}) == [("llm_wait_sec", "900", "1200")]))
    cases.append(("既定値 TRUE/FALSE型",
                  check_default_values("| `anonymize_default` | x | 既定TRUE |\n",
                                       {"anonymize_default": True},
                                       {"anonymize_default"}) == []))
    cases.append(("既定値 空文字型",
                  check_default_values("| `data_dir` | x | 既定は空 |\n",
                                       {"data_dir": ""}, {"data_dir"}) == []))

    # (2) 部分一致の罠: 禁止語の部分一致で「正しい長い名前」を赤くしない。
    #     本ツールはブロックリスト方式を取らず「存在するか」だけを見るため、
    #     長い名前がhaystackに存在すれば常に緑になることを確認する。
    cases.append(("部分一致の罠回避(存在すれば緑)",
                  check_bracket_names("[診断を開く]", "ボタン[診断を開く]です") == []))
    cases.append(("ボタン名 負例(存在しなければ赤)",
                  check_bracket_names("[存在しないボタン]", "他のテキスト") ==
                  ["存在しないボタン"]))
    cases.append(("マークダウンリンクは対象外",
                  check_bracket_names("[16章](docs/spec/16.md)", "") == []))
    cases.append(("タブ名 正例", check_tab_names("「使い方」タブ", "使い方タブがある")
                  == []))
    cases.append(("タブ名 負例", check_tab_names("「操作ガイド」タブ", "使い方タブ")
                  == ["操作ガイド"]))
    # 実測で発見した誤検知の再発防止。
    cases.append(("OS純正メニューは誤検知しない",
                  check_bracket_names("Excelの[フィルター]を押す", "") == []))
    cases.append(("OS純正タブ名は誤検知しない",
                  check_tab_names("「全般」タブのチェックを入れる", "") == []))
    cases.append(("履歴段落は検査しない",
                  check_bracket_names(
                      "現況（当時の記録＝履歴であり、[① 案件を作る]と呼ばれていた）",
                      "") == []))
    cases.append(("履歴段落でない箇所は通常どおり検査する",
                  check_bracket_names(
                      "現況（当時の記録＝履歴）\n\n[① 案件を作る]を押す。",
                      "") == ["① 案件を作る"]))
    cases.append(("vX.Y変更概要 段落は検査しない",
                  check_bracket_names(
                      "v4.0変更概要: 旧ボタン名（[① 調べる指示文を出す]〜）"
                      "を全面削除した。",
                      "") == []))

    # 対象外判定
    cases.append(("spec除外", is_excluded("docs/spec/13_データ設計.md")))
    cases.append(("受領除外", is_excluded("docs/受領/x.md")))
    cases.append(("裁定書除外", is_excluded("docs/裁定書38.md")))
    cases.append(("日付ファイル除外", is_excluded("docs/22_実装前監査記録_20260829.md")))
    cases.append(("対象文書は除外されない", not is_excluded("docs/25_利用ガイド.md")))
    cases.append(("README除外されない", not is_excluded("README.md")))

    bad = [name for name, ok in cases if not ok]
    for name in bad:
        print("  自己テスト NG: %s" % name)
    print("  自己テスト: %d/%d" % (len(cases) - len(bad), len(cases)))
    return not bad


def main() -> int:
    ap = argparse.ArgumentParser(description="文書検問(裁定書38 班D §1(3))")
    ap.add_argument("--verbose", action="store_true", help="エラーコード集合を列挙する")
    args = ap.parse_args()

    print("doc_gate: 文書検問(裁定書38 班D §1(3)。伝書鳩Part1-3の移植)")
    errors = run_checks(args.verbose)

    if not self_test():
        print("結果: 自己テスト失敗(検出器が壊れています)")
        return 2

    if errors:
        print("ERROR: %d 件" % errors)
        print("結果: NG")
        return 1
    print("結果: OK 4条件(既定値/ボタン名・タブ名/エラーコード/対象外判定)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
