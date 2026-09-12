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
        **ナビ画面(HTML画面=正)の区画ボタンの文脈では `ui/` だけを照合する**
        (W15 Round2 R2-10 → T-M2)。旧シート画面には同じ位置に別名のボタン
        (例: 出力の区画)が今も実装として残っているため、手順書がHTML画面の
        区画を説明しながら旧シート画面の名前を書いていても緑になっていた。
        Round2 の Fix 波は `src/ui/modUI*.bas` だけを外したが、旧ボタン名の
        多くは `build/sheets_main.json` にもキャプション・案内文として載って
        いるため4語が素通りした(実測)。区画の正は `ui/` の実物だけなので、
        照合先も `ui/` だけにする。**シート画面専用の手順**(docs/24 §5・§8、
        docs/25 第1部〔予備: `ui_mode=sheet`〕)は従来どおり全集合を見る。
    (3) src が吐くエラーコード `E0\\d{3}` ⊆ 16章§1の表 ⊆ docs/25 が
        本文中に書くエラーコード(一方向の部分集合。docs/25は全コードを
        網羅する文書ではないため、逆方向〔16章にあってdocs/25に無い〕は
        検査しない=誤検知の温床)
    (4) 対象外ファイルが誤って検査対象へ混ざっていないことの自己確認

使い方:
    python3 tools/doc_gate.py
    python3 tools/doc_gate.py --verbose
    python3 tools/doc_gate.py --selftest   # 回帰網だけ(docs は検査しない)
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
#   - `[コンテンツの有効化]`: Excel のセキュリティ警告の黄色い帯のボタン
#     (docs/25:28・89・420、docs/26:120、docs/24:123・136・291)。このアプリの
#     ボタンではないので ui/ にも modUI* にも無い。これまでは
#     build/sheets_main.json の案内文に同じ語があるせいで**たまたま**緑だった。
#     区画の文脈の照合先を ui/ だけにした(T-M2)ことで実体が出たため、
#     他のOS純正ボタンと同じくここへ明示する。
NATIVE_OS_UI_ALLOWLIST = {"再表示", "フィルター", "名前を付けて保存", "オプション",
                          "コンテンツの有効化",
                          "デバッグ", "VBAProject のコンパイル", "ツール", "参照設定"}  # VBE のメニュー(docs/24 §8.1-6)
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


def extract_ui_names(texts: dict, include_sheet_ui: bool = True,
                     include_sheets_json: bool = True) -> str:
    """突合対象の全テキストを1本に連結する。

    既定(両方True)は「画面のどこかにある名前か」を見る従来の集合
    = ui/ + build/sheets_main.json + src/ui/modUI*.bas。

    **ナビ画面の区画ボタンの文脈では `ui/` だけを見る**(W15 Round2 T-M2)。
    include_sheet_ui=False で旧シート画面の描画(`src/ui/modUI*.bas`)を、
    include_sheets_json=False でシート定義(`build/sheets_main.json`)を外す。
    Round2 の Fix 波は modUI* だけを外したが、旧シート画面のボタン名の多くは
    **sheets_main.json 側にも**キャプション・案内文として載っているため、
    R2 が名指しした旧名のうち4語(実測)が区画文脈で書き戻しても緑のままだった。
    「区画N の [ボタン名]」はHTML画面の実物(`ui/`)が唯一の正なので、
    照合先も `ui/` だけにする。
    """
    chunks = list(texts.values())
    if include_sheets_json:
        d = json.loads(SHEETS_JSON.read_text(encoding="utf-8"))
        chunks.append(json.dumps(d, ensure_ascii=False))
    if include_sheet_ui:
        for path in sorted((REPO_ROOT / "src" / "ui").glob("modUI*.bas")):
            chunks.append(path.read_text(encoding="utf-8", errors="ignore"))
    return "\n".join(chunks)


def build_haystacks(ui_texts: dict) -> tuple[str, str]:
    """(通常の照合集合, 区画の文脈の照合集合) を作る。

    **配線をここ1箇所に閉じる**。run_checks に直書きしていたときは、区画側の
    旗を True へ戻す変異を自己テストが1件も捕まえられなかった(実測)。
    自己テストはこの関数の戻り値を見るので、戻し変異はその場で赤くなる。
    """
    return (extract_ui_names(ui_texts),
            extract_ui_names(ui_texts, include_sheet_ui=False,
                             include_sheets_json=False))


# ナビ画面(HTML画面=正)の説明をしている文書・区間。ここだけ modUI* を外す。
# 値は「この行以降は対象外」を表す正規表現(None なら文書の最後まで対象)。
#   docs/25 は第0部(HTML画面)だけが対象。第1部は `ui_mode=sheet` の予備画面の
#   手順なので、旧シート画面のボタン名が書いてあるのが**正しい**。
#   docs/24 は §5・§8 が旧シート画面の手順のままなので対象にしない
#   (本波の担当は docs/25・26。docs/24 の整理は司令塔の裁定待ち)。
NAVI_DOC_REGIONS = {
    "docs/25_利用ガイド.md": re.compile(r"^# 第1部"),
    "docs/26_はじめてガイド_調査からレポートまで.md": None,
}

KUKAKU_WORD = "区画"


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


def check_bracket_names(text: str, haystack: str, navi_haystack: str | None = None,
                        navi_end: "re.Pattern | None" = None) -> list[tuple[str, bool]]:
    """戻り値: (見つからなかった名前, ナビ画面の区画の文脈か) のリスト。

    navi_haystack を渡すと、**ナビ画面の区画の文脈**の行だけはそちら
    (modUI* を外した集合)で照合する(R2-10)。文脈の判定は2形:
      (a) その行に「区画」が出てくる
      (b) 見出し行に「区画」を持つ表の中の行(docs/25 の区画一覧表)
    navi_end にマッチする行が来たら、そこから先は通常の照合へ戻す
    (docs/25 の第1部=シート画面の予備手順)。
    """
    text = _strip_historical_paragraphs(text)
    missing = []
    in_navi = navi_haystack is not None
    in_table = False
    table_is_navi = False
    for line in text.split("\n"):
        if in_navi and navi_end is not None and navi_end.match(line):
            in_navi = False
        is_row = line.lstrip().startswith("|")
        if is_row:
            if not in_table:            # 表の1行目 = 見出し行だけが表全体の文脈を決める
                table_is_navi = KUKAKU_WORD in line
                in_table = True
        else:
            in_table = False
            table_is_navi = False
        navi_ctx = in_navi and (KUKAKU_WORD in line or (is_row and table_is_navi))
        hay = navi_haystack if navi_ctx else haystack
        for m in BRACKET_NAME.finditer(line):
            name = m.group(1).strip("*").strip()
            if not name or name.isdigit():
                continue
            if name in NATIVE_OS_UI_ALLOWLIST:
                continue
            if name not in hay:
                missing.append((name, navi_ctx))
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

    ui_texts = {
        "index.html": (UI_DIR / "index.html").read_text(encoding="utf-8")
        if (UI_DIR / "index.html").exists() else "",
        "views.js": (UI_DIR / "views.js").read_text(encoding="utf-8")
        if (UI_DIR / "views.js").exists() else "",
        "app.js": (UI_DIR / "app.js").read_text(encoding="utf-8")
        if (UI_DIR / "app.js").exists() else "",
    }
    # 区画の文脈は ui/ だけを照合する(T-M2)。配線は build_haystacks が持つ。
    haystack, navi_haystack = build_haystacks(ui_texts)
    n2 = 0
    for rel, text in docs_texts.items():
        is_navi_doc = rel in NAVI_DOC_REGIONS
        for name, navi_ctx in check_bracket_names(
                text, haystack,
                navi_haystack if is_navi_doc else None,
                NAVI_DOC_REGIONS.get(rel)):
            if navi_ctx:
                err('(2) %s: 区画の説明にある `[%s]` が **ui/** に見つかりません'
                    "(旧シート画面 modUI*.bas の名前のままの疑い。HTML画面の"
                    "実物の逐語へ直してください)" % (rel, name))
            else:
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
                  [("存在しないボタン", False)]))
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
                      "") == [("① 案件を作る", False)]))
    cases.append(("vX.Y変更概要 段落は検査しない",
                  check_bracket_names(
                      "v4.0変更概要: 旧ボタン名（[① 調べる指示文を出す]〜）"
                      "を全面削除した。",
                      "") == []))

    # W15 Round2 R2-10: 区画の文脈では旧シート画面の名前で緑にしない。
    full = "旧シート画面の[レポートを出す]\nHTML画面の[レポート出力]"
    navi = "HTML画面の[レポート出力]"
    cases.append(("区画の文脈は modUI* の名前で緑にしない",
                  check_bracket_names("区画4の [レポートを出す] を押す。", full, navi)
                  == [("レポートを出す", True)]))
    cases.append(("区画の文脈でも ui/ にあれば緑",
                  check_bracket_names("区画4の [レポート出力] を押す。", full, navi)
                  == []))
    cases.append(("区画の文脈でない行は従来どおり(シート画面専用の手順)",
                  check_bracket_names("`S4_骨子` の [レポートを出す] を押す。",
                                      full, navi) == []))
    cases.append(("見出しに区画を持つ表の行は区画の文脈",
                  check_bracket_names(
                      "| 区画 | 見出し | 中身 |\n|---|---|---|\n"
                      "| 4 | 出力 | [レポートを出す] |", full, navi)
                  == [("レポートを出す", True)]))
    cases.append(("区画を持たない表の行は区画の文脈ではない",
                  check_bracket_names(
                      "| 名前 | 中身 |\n|---|---|\n| 出力 | [レポートを出す] |",
                      full, navi) == []))
    # 表の文脈を決めるのは**見出し行だけ**(途中の行の「区画」で表全体を
    # 巻き込まない=誤検知を作らない)。
    cases.append(("表の途中行の区画は次の行まで広げない",
                  check_bracket_names(
                      "| 名前 | 中身 |\n|---|---|\n| 区画4 | [レポート出力] |\n"
                      "| 下書き | [レポートを出す] |", full, navi) == []))
    cases.append(("navi_end 以降は従来どおり(docs/25 第1部)",
                  check_bracket_names(
                      "区画4の [レポート出力]\n\n# 第1部（予備: `ui_mode=sheet`）\n\n"
                      "区画4の [レポートを出す] を押す。",
                      full, navi, re.compile(r"^# 第1部")) == []))
    cases.append(("navi_haystack を渡さなければ従来の挙動",
                  check_bracket_names("区画4の [レポートを出す] を押す。", full)
                  == []))

    # W15 Round2 T-M2: 区画の文脈の照合先は **ui/ だけ**。
    # modUI* だけでなく build/sheets_main.json 由来の名前でも緑にしない
    # (Round2 の Fix はここが抜けており、旧シート名4語が素通りしていた)。
    full_j = "HTML画面の[提案骨子を見る]\nシート定義の案内文にある[ここに貼る]"
    navi_j = "HTML画面の[提案骨子を見る]"
    cases.append(("区画の文脈は sheets_main.json の名前でも緑にしない",
                  check_bracket_names("区画2の [ここに貼る] に貼る。", full_j, navi_j)
                  == [("ここに貼る", True)]))
    cases.append(("区画の文脈でなければ sheets_main.json の名前は従来どおり緑",
                  check_bracket_names("`S1_調べる` の [ここに貼る] に貼る。",
                                      full_j, navi_j) == []))
    cases.append(("区画の文脈でも ui/ にある名前は緑(誤検知0)",
                  check_bracket_names("区画3の [提案骨子を見る] を押す。",
                                      full_j, navi_j) == []))
    # 照合集合の組み立てそのものを固定する(どちらの旗も効くこと)。
    only_ui = extract_ui_names({"x": "ZzDocGateProbe"}, include_sheet_ui=False,
                               include_sheets_json=False)
    with_json = extract_ui_names({"x": "ZzDocGateProbe"}, include_sheet_ui=False,
                                 include_sheets_json=True)
    cases.append(("区画用の集合は ui/ のテキストだけ", only_ui == "ZzDocGateProbe"))
    cases.append(("include_sheets_json=True なら sheets_main.json が入る",
                  len(with_json) > len(only_ui) and "ZzDocGateProbe" in with_json))
    # **配線**の固定(run_checks が実際に渡す2本)。旗を戻す変異をここで捕まえる。
    hay, navi_hay = build_haystacks({"x": "ZzDocGateProbe"})
    cases.append(("配線: 区画用は ui/ だけ", navi_hay == "ZzDocGateProbe"))
    cases.append(("配線: 通常用は sheets_main.json を含む", '"sheets"' in hay))
    cases.append(("配線: 通常用は modUI*.bas を含む", "modUI" in hay))
    cases.append(("配線: 区画用は modUI*.bas を含まない", "modUI" not in navi_hay))

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
    ap.add_argument("--selftest", action="store_true",
                    help="自己テスト(回帰網)だけを回す(docs は検査しない)")
    args = ap.parse_args()

    print("doc_gate: 文書検問(裁定書38 班D §1(3)。伝書鳩Part1-3の移植)")
    if args.selftest:
        if not self_test():
            print("結果: 自己テスト失敗(検出器が壊れています)")
            return 2
        print("結果: 自己テストOK")
        return 0
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
