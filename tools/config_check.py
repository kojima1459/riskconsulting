#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""config_check.py - configキーの5点一致検査(裁定書38 班D §1(2))

================================================================================
なぜ要るか:
    config キーは (1) 配布ブックへ焼く既定値表 (2) 起動時にコードが登録する
    既定値 (3) 実際にコードが読む値 (4) 13章§2.3の仕様表 (5) 19章§4の一覧、
    という5つの場所に同じ情報が分散している。どれか1つだけ更新して残りを
    直し忘れる事故を機械で数える。

5点:
    A. build/sheets_main.json の `config` シートの `defaults[].name`
    B. modBoot.bas / modBootNavi.bas の `modConfig.RegisterDefault "key"`
    C. src/**/*.bas の `modConfig.Get{Str,Long,Bool,Double}("key"` /
       `HasKey("key"` での読取(1件以上)。動的連結("接頭辞" & 式)で
       組み立てるキー(`"dr_url_" & kind` 等)は、その接頭辞で始まる
       キーを読取ありとして救済する(orphan_check.py と同じ設計)。
    D. docs/spec/13_データ設計.md ### 2.3 の表(1列目。`/`区切りの複数キー
       セルを分解する)
    E. docs/spec/19_用語集とレジストリ.md の `**configキー**:` 一覧

ビルドが焼く例外(明示。5点一致から外す):
    - `tests_expected`: 13章§2.3が明記するとおり「値源はビルド入力であって
      既定値表ではない」。RegisterDefaultで登録せず、`wintest/tests_expected.txt`
      からビルドが直接焼く。読取は `modTestsRunnerUi` が定数
      `TR_EXPECTED_KEY` 経由で行うため、C(直接の文字列リテラル読取)には
      機械的に出現しない。A/D/E には出現してよい。

使い方:
    python3 tools/config_check.py
    python3 tools/config_check.py --verbose
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

SHEETS_JSON = REPO_ROOT / "build" / "sheets_main.json"
BOOT_FILES = (REPO_ROOT / "src" / "ui" / "modBoot.bas",
              REPO_ROOT / "src" / "ui" / "modBootNavi.bas")
SRC_ROOT = REPO_ROOT / "src"
SPEC13 = REPO_ROOT / "docs" / "spec" / "13_データ設計.md"
SPEC19 = REPO_ROOT / "docs" / "spec" / "19_用語集とレジストリ.md"

# ビルドが焼く例外(13章§2.3本文が明記)。RegisterDefaultと直接読取の
# 2点だけ免除する(A/D/Eには出現するのが正)。
BUILD_BAKED_EXCEPTIONS = {"tests_expected"}

# 意図的に「存在しないキー」を読ませるテスト用の名前(modConfig の
# fail-closed 契約を検証するための合成キー)。実在するconfigキーではない
# ので、読取集合Cの計算から除外する。
TEST_SENTINEL_RE = re.compile(r"^rpn_no_such_key")

# Step別上書きキーの動的合成(13章§2.3・裁定書34 §1.2): 実装は
# `stepKey = stepName & "_" & kind` で組み立て、`stepName` は呼び出し時の
# 実引数(実行時の値)なので静的に解決できない
# (`src/core/modGatewayRPN.bas` の `ResolveTuning`)。`ch_effort`/
# `ch_verbosity` はこの経路の唯一の**トップレベル登録キー**(stepName="ch"
# のとき)で、動的連結の接頭辞/接尾辞救済では追えないため、読取検査 C だけ
# 明示的に除外する(RegisterDefault・13章・19章・sheets_main.jsonの4点は
# 通常どおり照合する)。
DYNAMIC_STEP_OVERRIDE_KEYS = {"ch_effort", "ch_verbosity"}

GET_CALL = re.compile(
    r"\b(?:GetStr|GetLong|GetBool|GetDouble|HasKey)\s*\(\s*\"([a-zA-Z0-9_]+)\"",
    re.IGNORECASE)
# 定数名を第1引数に渡す間接呼び出し(例: `GetBool(VP_CFG_KEY, True)`)。
GET_CALL_IDENT = re.compile(
    r"\b(?:GetStr|GetLong|GetBool|GetDouble|HasKey)\s*\(\s*([A-Za-z_]\w*)\s*[,)]",
    re.IGNORECASE)
CONST_STR_DEF = re.compile(
    r"\bConst\s+(\w+)\s+As\s+String\s*=\s*\"([^\"]+)\"", re.IGNORECASE)
REGISTER_DEFAULT = re.compile(r'RegisterDefault\s+"([a-zA-Z0-9_]+)"', re.IGNORECASE)
CONCAT_PREFIX = re.compile(r'"([a-zA-Z0-9_]+_)"\s*&')

# ラッパー関数の1段だけの転送解決(例: `CapCfg(maxRows, cfgKey, dfltRows)` の
# 本体が `modConfig.GetLong(cfgKey, dfltRows)` を呼ぶ場合、呼び出し側の
# `CapCfg(x, "kb_case_rows", 5)` の "kb_case_rows" を読取ありとして拾う)。
FUNC_DEF_HEAD = re.compile(
    r"(?:Private|Public)\s+Function\s+(\w+)\s*\(([^)]*)\)", re.IGNORECASE)
PARAM_NAME = re.compile(r"(?:ByVal\s+|ByRef\s+)?(\w+)\s+As\s+", re.IGNORECASE)


def load_json_keys() -> list[dict]:
    d = json.loads(SHEETS_JSON.read_text(encoding="utf-8"))
    for s in d.get("sheets", []):
        if s.get("name") == "config":
            return s.get("defaults", [])
    return []


def load_regdef_keys() -> set[str]:
    out: set[str] = set()
    for f in BOOT_FILES:
        if f.exists():
            out |= set(REGISTER_DEFAULT.findall(f.read_text(encoding="utf-8")))
    return out


def _find_wrapper_key_params(text: str) -> dict[str, int]:
    """`modConfig.GetXxx(paramName, ...)` を素通しするラッパー関数を探す。

    戻り値: 関数名 -> configキーを受け取る引数の位置(0始まり)。
    """
    wrappers: dict[str, int] = {}
    for m in FUNC_DEF_HEAD.finditer(text):
        fname = m.group(1)
        params = [p.strip() for p in m.group(2).split(",") if p.strip()]
        param_names = []
        for p in params:
            pm = PARAM_NAME.search(p + " As ")  # 型注釈が無い形にも耐える
            param_names.append(pm.group(1) if pm else p.split()[-1])
        end = text.find("End Function", m.end())
        if end < 0:
            continue
        body = text[m.end():end]
        fwd = GET_CALL_IDENT.search(body)
        if fwd and fwd.group(1) in param_names:
            wrappers[fname] = param_names.index(fwd.group(1))
    return wrappers


def load_read_keys_and_prefixes_from_text(text: str) -> tuple[set[str], set[str]]:
    reads: set[str] = set()
    prefixes: set[str] = set()

    # (a) 直接の文字列リテラル読取。
    for m in GET_CALL.finditer(text):
        name = m.group(1)
        if TEST_SENTINEL_RE.match(name):
            continue
        reads.add(name)

    # (b) 定数名を介した間接読取(`Const X As String = "key"` -> `GetXxx(X,`)。
    const_map = dict(CONST_STR_DEF.findall(text))
    for m in GET_CALL_IDENT.finditer(text):
        ident = m.group(1)
        if ident in const_map:
            reads.add(const_map[ident])

    # (c) 1段ラッパー関数経由の読取(`CapCfg(n, "kb_case_rows", 5)` 型)。
    wrappers = _find_wrapper_key_params(text)
    for fname, key_idx in wrappers.items():
        for call in re.finditer(r"\b%s\s*\(([^()]*)\)" % re.escape(fname), text):
            args = [a.strip() for a in call.group(1).split(",")]
            if key_idx < len(args):
                lit = re.match(r'^"([a-zA-Z0-9_]+)"$', args[key_idx])
                if lit:
                    reads.add(lit.group(1))

    for m in CONCAT_PREFIX.finditer(text):
        prefixes.add(m.group(1))
    return reads, prefixes


def load_read_keys_and_prefixes(src_root: Path) -> tuple[set[str], set[str]]:
    reads: set[str] = set()
    prefixes: set[str] = set()
    for path in sorted(src_root.rglob("*.bas")):
        text = path.read_text(encoding="utf-8", errors="ignore")
        r, p = load_read_keys_and_prefixes_from_text(text)
        reads |= r
        prefixes |= p
    return reads, prefixes


def parse_spec13(text: str) -> set[str]:
    m = re.search(r"### 2\.3.*?(?=\n### 2\.4)", text, re.DOTALL)
    if not m:
        return set()
    out: set[str] = set()
    for row in re.finditer(r"^\|(.+?)\|(.+?)\|(.+?)\|\s*$", m.group(0), re.MULTILINE):
        col1 = row.group(1).strip()
        if col1 == "name" or set(col1) <= set("- "):
            continue
        for part in col1.split("/"):
            part = part.strip().strip("`").strip()
            if part:
                out.add(part)
    return out


def parse_spec19(text: str) -> set[str]:
    m = re.search(r"\*\*configキー\*\*:\s*13章§2\.3の表が正[（(](.+?)[）)]",
                  text, re.DOTALL)
    if not m:
        return set()
    return set(re.findall(r"([a-z][a-z0-9_]*)", m.group(1)))


def run_checks(verbose: bool) -> int:
    errors = 0

    def err(msg: str) -> None:
        nonlocal errors
        errors += 1
        print("ERROR " + msg)

    if not SHEETS_JSON.exists():
        print("ERROR %s がありません" % SHEETS_JSON.relative_to(REPO_ROOT))
        return 1
    defaults = load_json_keys()
    if not defaults:
        err("build/sheets_main.json に config シートの defaults が見つかりません")
    json_keys = {row["name"] for row in defaults}

    regdef_keys = load_regdef_keys()
    read_keys, dyn_prefixes = load_read_keys_and_prefixes(SRC_ROOT)
    spec13_keys = parse_spec13(SPEC13.read_text(encoding="utf-8")) if SPEC13.exists() else set()
    spec19_keys = parse_spec19(SPEC19.read_text(encoding="utf-8")) if SPEC19.exists() else set()

    if not spec13_keys:
        err("docs/spec/13_データ設計.md の ### 2.3 表を読み取れませんでした")
    if not spec19_keys:
        err("docs/spec/19_用語集とレジストリ.md の configキー一覧を読み取れませんでした")

    def is_read(name: str) -> bool:
        if name in read_keys:
            return True
        return any(name.startswith(p) for p in dyn_prefixes if len(p) >= 3)

    all_names = json_keys | regdef_keys | spec13_keys | spec19_keys

    n1 = n2 = n3 = n4 = 0
    for name in sorted(all_names):
        in_json = name in json_keys
        in_regdef = name in regdef_keys
        in_read = is_read(name)
        in_13 = name in spec13_keys
        in_19 = name in spec19_keys
        is_exception = name in BUILD_BAKED_EXCEPTIONS

        if not in_json:
            err('config "%s" が build/sheets_main.json の config.defaults に'
                "ありません" % name)
            n1 += 1
        if not in_regdef and not is_exception:
            err('config "%s" が RegisterDefault(modBoot/modBootNavi)で'
                "登録されていません(13章§2.3の既定値は起動時に"
                "RegisterDefaultで登録する契約=modConfig.bas冒頭のコメント)"
                % name)
            n2 += 1
        if not in_read and not is_exception and name not in DYNAMIC_STEP_OVERRIDE_KEYS:
            err('config "%s" を読むコードが src/**/*.bas に見つかりません'
                "(未読キー)" % name)
            n3 += 1
        if not in_13:
            err('config "%s" が docs/spec/13_データ設計.md ### 2.3 の表に'
                "ありません" % name)
            n4 += 1
        if not in_19:
            err('config "%s" が docs/spec/19_用語集とレジストリ.md の'
                "configキー一覧にありません" % name)
            n4 += 1

    print("  A(sheets_main.json)   %d件" % len(json_keys))
    print("  B(RegisterDefault)    %d件 / 不一致 %d" % (len(regdef_keys), n2))
    print("  C(読取。動的接頭辞%d件で救済) %d件 / 不一致 %d" %
          (len(dyn_prefixes), len(read_keys), n3))
    print("  D(13章§2.3)           %d件" % len(spec13_keys))
    print("  E(19章§4)             %d件 / D・Eの不一致 %d" %
          (len(spec19_keys), n4))
    print("  A自体の不一致(欠落)    %d件" % n1)
    print("  ビルドが焼く例外: %s" % ", ".join(sorted(BUILD_BAKED_EXCEPTIONS)))
    print("  Step別上書きの動的合成で読取検査Cのみ除外: %s" %
          ", ".join(sorted(DYNAMIC_STEP_OVERRIDE_KEYS)))
    if verbose:
        print("      動的接頭辞: " + ", ".join(sorted(dyn_prefixes)))

    return errors


# ---------------------------------------------------------------------------
# 自己テスト
# ---------------------------------------------------------------------------
def self_test() -> bool:
    cases = []

    cases.append(("RegisterDefault抽出",
                  load_regdef_keys() >= {"app_version", "llm_transport"}))
    reads, prefixes = load_read_keys_and_prefixes(SRC_ROOT)
    cases.append(("読取抽出 非空", len(reads) > 10))
    cases.append(("読取抽出 テストsentinelを除外",
                  "rpn_no_such_key" not in reads))
    cases.append(("動的接頭辞抽出(dr_url_)", "dr_url_" in prefixes))
    cases.append(("定数名間接読取(ui_fullscreen)", "ui_fullscreen" in reads))
    cases.append(("定数名間接読取(guide_tour_done)", "guide_tour_done" in reads))
    cases.append(("ラッパー関数1段解決(kb_case_rows)", "kb_case_rows" in reads))
    cases.append(("ラッパー関数1段解決(kb_incident_rows)",
                  "kb_incident_rows" in reads))

    # 合成データでの正例/負例(実ファイルの偶然に依存しない検証)。
    sample_const = (
        'Private Const VP_CFG_KEY As String = "ui_fullscreen"\n'
        "Sub X()\n"
        "    If Not modConfig.GetBool(VP_CFG_KEY, True) Then Exit Sub\n"
        "End Sub\n"
    )
    r2, _ = load_read_keys_and_prefixes_from_text(sample_const)
    cases.append(("定数名間接読取(合成)", "ui_fullscreen" in r2))

    sample_wrapper = (
        "Private Function CapCfg(ByVal maxRows As Long, ByVal cfgKey As String, "
        "ByVal dfltRows As Long) As Long\n"
        "    CapCfg = modConfig.GetLong(cfgKey, dfltRows)\n"
        "End Function\n"
        "Sub Y()\n"
        '    n = CapCfg(0, "kb_case_rows", 5)\n'
        "End Sub\n"
    )
    r3, _ = load_read_keys_and_prefixes_from_text(sample_wrapper)
    cases.append(("ラッパー関数1段解決(合成)", "kb_case_rows" in r3))

    sample13 = (
        "### 2.3 `config`\n"
        "| name | 既定値 | 説明 |\n"
        "|---|---|---|\n"
        "| app_version | 2.0.0 | |\n"
        "| direct_model / direct_api_base | gpt-4.1 / url | direct用 |\n"
        "\n"
        "### 2.4 ログ\n"
        "| name | x |\n"
    )
    parsed13 = parse_spec13(sample13)
    cases.append(("13章表 複数キーセルの分解",
                  parsed13 == {"app_version", "direct_model", "direct_api_base"}))

    sample19 = "**configキー**: 13章§2.3の表が正（app_version / llm_transport / mock_llm）他"
    parsed19 = parse_spec19(sample19)
    cases.append(("19章一覧の抽出",
                  parsed19 == {"app_version", "llm_transport", "mock_llm"}))

    bad = [name for name, ok in cases if not ok]
    for name in bad:
        print("  自己テスト NG: %s" % name)
    print("  自己テスト: %d/%d" % (len(cases) - len(bad), len(cases)))
    return not bad


def main() -> int:
    ap = argparse.ArgumentParser(description="configキーの5点一致検査(裁定書38 班D)")
    ap.add_argument("--verbose", action="store_true", help="動的接頭辞も列挙する")
    args = ap.parse_args()

    print("config_check: configキーの5点一致検査(裁定書38 班D §1(2))")
    errors = run_checks(args.verbose)

    if not self_test():
        print("結果: 自己テスト失敗(検出器が壊れています)")
        return 2

    if errors:
        print("ERROR: %d 件" % errors)
        print("結果: NG")
        return 1
    print("結果: OK 5点一致(sheets_main.json/RegisterDefault/読取/13章/19章)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
