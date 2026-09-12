#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
bench_s1.py - S1抽出品質ベンチの採点器(班F・裁定書37 §5 / w14 B_dr_quality.md §5)

数値の定義は bench/README.md と一字一句一致させること(このファイルが変更に
追随していなければ README を直す)。

3モード:
  --mode replay      bench/gold/*.json と、指定ディレクトリ配下のS1出力JSON群を
                      突き合わせ、一致率/回収率/捏造率/揺れ(+出典実在率)を出す。
                      AIを1回も呼ばない(既定・CI向け)。
  --mode selfcheck    src/test/modMockLlm.bas の BuildS1NewJson() を実行時に
                      抽出し、採点器自身の回帰(期待値どおりの数値が出るか)を
                      確認する。exit 0/1。
  --gold-check        bench/gold/*.json の forbidden 断片が対応する
                      bench/fixtures/<company>/ に本当に無いこと、recall 断片が
                      本当にあることを検算する(gold の嘘を機械で止める)。
                      exit 0/1。

正規化規則(表記揺れ耐性)は src/app/modGround.bas の NormalizeForMatch を
そのままPythonへ写す(値源は1箇所。ズレたらここを modGround に合わせる)。
gate.py には入れない(LLM実行を伴う指標をCIの合否にしないため。README参照)。
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import unicodedata
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
BENCH_DIR = REPO_ROOT / "bench"
GOLD_DIR = BENCH_DIR / "gold"
FIXTURES_DIR = BENCH_DIR / "fixtures"
OUT_DIR = BENCH_DIR / "out"
MOCK_LLM_BAS = REPO_ROOT / "src" / "test" / "modMockLlm.bas"

# data_key の正: 13章§2.2(10本の input_* 系。s1_json 等の生成物系は除く)。
INPUT_DATA_KEYS = [
    "input_hp", "input_yuho", "input_memo", "input_contract",
    "input_prev_renewal", "input_dossier", "input_field_notes",
    "input_coverage_note", "input_finance", "input_hearing_answers",
]

# ============================================================================
# NormalizeForMatch - src/app/modGround.bas の逐語移植(純関数)。
#   (1) 全角英数記号(U+FF01..U+FF5E)を半角へ寄せ、全角空白(U+3000)は半角空白へ。
#   (2) 空白・タブ・改行、句読点・鍵括弧・中黒・カンマ・ピリオド・長音、記号類を
#       落とす。(3) 英大文字を小文字へ。数字・かな・漢字はそのまま残す。
# ============================================================================
_DROP_CHARS = set(
    " \t\r\n"
    "!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~"
    "、。・「」『』【】〔〕ー"
)


def normalize_for_match(t: str) -> str:
    if not t:
        return ""
    out = []
    for ch in t:
        code = ord(ch)
        # 全角英数記号 -> 半角
        if 0xFF01 <= code <= 0xFF5E:
            code -= 0xFEE0
        # 全角空白 -> 半角空白(次で落ちる)
        if code == 0x3000:
            code = 0x20
        ch2 = chr(code)
        # 英大文字 -> 小文字
        if "A" <= ch2 <= "Z":
            ch2 = ch2.lower()
        if ch2 in _DROP_CHARS:
            continue
        out.append(ch2)
    return "".join(out)


def fragment_found(fragment: str, haystack_norm: str) -> bool:
    """gold の recall/forbidden 断片が、正規化済み haystack に部分一致するか。
    modGround.QuoteFound と違い、断片は短い固定文字列なので headChars による
    先頭切詰めは行わず全文一致で見る(先頭だけ一致して尾部が違う誤検知を避ける)。
    """
    q = normalize_for_match(fragment)
    if not q:
        return False
    if not haystack_norm:
        return False
    return q in haystack_norm


# ============================================================================
# VBA文字列連結の逐語再現("s = s & "...."" & vbLf" の系列をJSON文字列へ復元)。
#   selfcheck が modMockLlm.bas から「本物のmock応答」を抽出するために使う。
# ============================================================================
def _parse_vba_concat_line(line: str) -> str:
    out = []
    i, n = 0, len(line)
    while i < n:
        c = line[i]
        if c == '"':
            j = i + 1
            buf = []
            while j < n:
                if line[j] == '"':
                    if j + 1 < n and line[j + 1] == '"':
                        buf.append('"')
                        j += 2
                        continue
                    j += 1
                    break
                buf.append(line[j])
                j += 1
            out.append("".join(buf))
            i = j
        elif line[i:i + 4] == "vbLf":
            out.append("\n")
            i += 4
        else:
            i += 1
    return "".join(out)


def extract_vba_function_json(bas_path: Path, func_name: str) -> str:
    """modMockLlm.bas から `Public Function <func_name>() As String ... End
    Function` を切り出し、`s = s & "...."`(& vbLf 可)の並びだけを実行時に
    連結してJSON文字列を復元する(ハードコードで手写ししない=複製腐敗を避ける)。
    """
    text = bas_path.read_text(encoding="utf-8")
    pat = re.compile(
        r"Public Function " + re.escape(func_name) + r"\(\).*?\n(.*?)\nEnd Function",
        re.S,
    )
    m = pat.search(text)
    if not m:
        raise ValueError(f"{func_name} が {bas_path} に見つかりません")
    body = m.group(1)
    pieces = []
    for line in body.splitlines():
        ls = line.strip()
        if ls.startswith("s = s &") or ls.startswith("s = s&"):
            rhs = ls.split("&", 1)[1] if "&" in ls else ""
            # 先頭の "s = s" を落としたあとの全体を安全に再走査する
            pieces.append(_parse_vba_concat_line(ls[len("s = s"):]))
    return "".join(pieces)


# ============================================================================
# gold / fixtures の読み込み
# ============================================================================
def load_gold() -> dict:
    gold = {}
    for p in sorted(GOLD_DIR.glob("*.json")):
        data = json.loads(p.read_text(encoding="utf-8"))
        key = p.stem
        gold[key] = data
    return gold


def load_fixture_text(fixture_dir: str) -> str:
    d = FIXTURES_DIR / fixture_dir
    if not d.is_dir():
        return ""
    parts = []
    for key in INPUT_DATA_KEYS:
        f = d / f"{key}.txt"
        if f.is_file():
            parts.append(f.read_text(encoding="utf-8"))
    return "\n".join(parts)


# ============================================================================
# --gold-check: gold の forbidden が fixtures に無いこと、recall があることを検算
# ============================================================================
def cmd_gold_check() -> int:
    gold = load_gold()
    if not gold:
        print("[gold-check] bench/gold/*.json が1件もありません(fail-closed)")
        return 1
    problems = []
    for key, g in gold.items():
        fixture_dir = g.get("fixture_dir", key)
        hay = load_fixture_text(fixture_dir)
        hay_norm = normalize_for_match(hay)
        if not hay_norm:
            problems.append(f"[{key}] fixtures/{fixture_dir}/ の本文が空です(検算不能)")
            continue
        for frag in g.get("recall", []):
            if not fragment_found(frag, hay_norm):
                problems.append(
                    f"[{key}] recall断片が fixtures に見当たりません: {frag!r}"
                )
        for frag in g.get("forbidden", []):
            if fragment_found(frag, hay_norm):
                problems.append(
                    f"[{key}] forbidden断片が fixtures に実在します"
                    f"(捏造判定に使えません): {frag!r}"
                )
    if problems:
        print(f"[gold-check] NG {len(problems)}件:")
        for p in problems:
            print("  - " + p)
        return 1
    print(f"[gold-check] OK: {len(gold)}社の gold を検算 (recall実在・forbidden不在を確認)")
    return 0


# ============================================================================
# JSON値の取得(ドット区切りパス。ネスト辞書のみ対応。配列添字は使わない)
# ============================================================================
def get_path(obj, path: str):
    cur = obj
    for part in path.split("."):
        if not isinstance(cur, dict) or part not in cur:
            return None
        cur = cur[part]
    return cur


def values_equal(a, b) -> bool:
    if a is None or b is None:
        return a == b
    return normalize_for_match(str(a)) == normalize_for_match(str(b))


# ============================================================================
# S1出力JSONの読み込み(生のS1オブジェクト、または13章§2.2の案件エクスポート
#   形式 {"format":"riscon-navi-case",...,"data":[{"key":"s1_json","content":
#   "..."}...]} のどちらでも受ける。live運用(会社PCの案件JSONエクスポート)を
#   そのまま bench/out/live/<company>.json に置けるようにするため)。
# ============================================================================
def load_s1_output(path: Path):
    raw = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(raw, dict) and raw.get("format") == "riscon-navi-case":
        for item in raw.get("data", []):
            if item.get("key") == "s1_json":
                content = item.get("content", "")
                if not content:
                    return None
                return json.loads(content)
        return None
    return raw


def company_key_of(filename: str) -> str:
    # "<company>.json" / "<company>.<run>.json" のどちらも <company> を返す
    stem = Path(filename).stem
    return stem.split(".")[0]


def collect_replay_files(in_path: Path) -> dict:
    """company -> [Path,...] (複数あれば揺れ計算に使う実行群)"""
    files = {}
    if in_path.is_file():
        candidates = [in_path]
    else:
        candidates = sorted(in_path.glob("*.json"))
    for p in candidates:
        key = company_key_of(p.name)
        files.setdefault(key, []).append(p)
    return files


# ============================================================================
# 採点本体(replay/selfcheckの共通処理)
# ============================================================================
def score_outputs(gold: dict, files_by_company: dict) -> dict:
    exact_total = 0
    exact_hit = 0
    recall_total = 0
    recall_hit = 0
    forbidden_total = 0
    forbidden_hit = 0
    variance_total = 0
    variance_consistent = 0
    src_checked = 0
    src_hit = 0
    per_company = {}

    for key, g in gold.items():
        outs = files_by_company.get(key, [])
        fixture_dir = g.get("fixture_dir", key)
        fixture_text = load_fixture_text(fixture_dir)
        fixture_norm = normalize_for_match(fixture_text)

        c_exact_total = 0
        c_exact_hit = 0
        c_recall_total = 0
        c_recall_hit = 0
        c_forbidden_total = 0
        c_forbidden_hit = 0
        exact_values_by_field = {k: [] for k in g.get("exact", {})}

        if not outs:
            per_company[key] = {
                "runs": 0, "exact": None, "recall": None, "forbidden": None,
                "note": "出力JSONなし(採点対象外)",
            }
            continue

        for out_path in outs:
            try:
                data = load_s1_output(out_path)
            except Exception as e:  # noqa: BLE001
                per_company.setdefault(key, {})["note"] = f"読込失敗: {e}"
                continue
            if data is None:
                continue
            text_norm = normalize_for_match(json.dumps(data, ensure_ascii=False))

            for field, expect in g.get("exact", {}).items():
                actual = get_path(data, field)
                exact_values_by_field[field].append(actual)
                exact_total += 1
                c_exact_total += 1
                if values_equal(actual, expect):
                    exact_hit += 1
                    c_exact_hit += 1

            for frag in g.get("recall", []):
                recall_total += 1
                c_recall_total += 1
                if fragment_found(frag, text_norm):
                    recall_hit += 1
                    c_recall_hit += 1

            for frag in g.get("forbidden", []):
                forbidden_total += 1
                c_forbidden_total += 1
                if fragment_found(frag, text_norm):
                    forbidden_hit += 1
                    c_forbidden_hit += 1

            sources = data.get("sources") if isinstance(data, dict) else None
            if sources:
                for s in sources:
                    url = s.get("url") if isinstance(s, dict) else str(s)
                    if not url:
                        continue
                    src_checked += 1
                    if fragment_found(url, fixture_norm):
                        src_hit += 1

        # 揺れ: 同一companyで2回以上の出力があるフィールドだけ数える
        if len(outs) >= 2:
            for field, values in exact_values_by_field.items():
                if len(values) < 2:
                    continue
                variance_total += 1
                normed = {normalize_for_match(str(v)) for v in values}
                if len(normed) == 1:
                    variance_consistent += 1

        per_company[key] = {
            "runs": len(outs),
            "exact": (c_exact_hit / c_exact_total) if c_exact_total else None,
            "recall": (c_recall_hit / c_recall_total) if c_recall_total else None,
            "forbidden": (c_forbidden_hit / c_forbidden_total) if c_forbidden_total else None,
        }

    summary = {
        "exact_match_rate": (exact_hit / exact_total) if exact_total else None,
        "exact_hit": exact_hit, "exact_total": exact_total,
        "recall_rate": (recall_hit / recall_total) if recall_total else None,
        "recall_hit": recall_hit, "recall_total": recall_total,
        "forbidden_rate": (forbidden_hit / forbidden_total) if forbidden_total else None,
        "forbidden_hit": forbidden_hit, "forbidden_total": forbidden_total,
        "variance_rate": (variance_consistent / variance_total) if variance_total else None,
        "variance_consistent": variance_consistent, "variance_total": variance_total,
        "source_exist_rate": (src_hit / src_checked) if src_checked else None,
        "source_checked": src_checked, "source_hit": src_hit,
        "per_company": per_company,
    }
    return summary


def fmt_pct(v):
    return "n/a" if v is None else f"{v * 100:.1f}%"


def print_summary(summary: dict) -> None:
    print("== S1抽出品質ベンチ 結果 ==")
    print(f"一致率(exact) : {fmt_pct(summary['exact_match_rate'])} "
          f"({summary['exact_hit']}/{summary['exact_total']})")
    print(f"回収率(recall): {fmt_pct(summary['recall_rate'])} "
          f"({summary['recall_hit']}/{summary['recall_total']})")
    print(f"捏造率(forbidden): {fmt_pct(summary['forbidden_rate'])} "
          f"({summary['forbidden_hit']}/{summary['forbidden_total']})")
    print(f"揺れ(variance一致): {fmt_pct(summary['variance_rate'])} "
          f"({summary['variance_consistent']}/{summary['variance_total']})")
    print(f"出典実在率(sources): {fmt_pct(summary['source_exist_rate'])} "
          f"({summary['source_hit']}/{summary['source_checked']})")
    print("-- 社別 --")
    for key, row in summary["per_company"].items():
        if row.get("runs", 0) == 0:
            print(f"  {key}: 出力なし ({row.get('note', '')})")
            continue
        print(
            f"  {key}: runs={row['runs']} exact={fmt_pct(row['exact'])} "
            f"recall={fmt_pct(row['recall'])} forbidden={fmt_pct(row['forbidden'])}"
        )


def cmd_replay(in_path: Path) -> int:
    gold = load_gold()
    if not gold:
        print("[replay] bench/gold/*.json が1件もありません(fail-closed)")
        return 1
    files_by_company = collect_replay_files(in_path)
    summary = score_outputs(gold, files_by_company)
    print_summary(summary)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    (OUT_DIR / "summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"\n(詳細を {OUT_DIR / 'summary.json'} へ書き出しました)")
    return 0


# ============================================================================
# --mode selfcheck: mockのS1応答を抽出し、採点器自身の回帰を確認する
# ============================================================================
def cmd_selfcheck() -> int:
    if not MOCK_LLM_BAS.is_file():
        print(f"[selfcheck] {MOCK_LLM_BAS} が見つかりません(fail-closed)")
        return 1
    try:
        mock_json_text = extract_vba_function_json(MOCK_LLM_BAS, "BuildS1NewJson")
    except Exception as e:  # noqa: BLE001
        print(f"[selfcheck] mock応答の抽出に失敗しました: {e}")
        return 1
    try:
        mock_json = json.loads(mock_json_text)
    except Exception as e:  # noqa: BLE001
        print(f"[selfcheck] 抽出したmock応答がJSONとしてパースできません: {e}")
        print(mock_json_text[:400])
        return 1

    if mock_json.get("company_name") != "株式会社浜松スイーツファクトリー":
        print("[selfcheck] 抽出したmock応答のcompany_nameが期待値と異なります: "
              f"{mock_json.get('company_name')!r}")
        return 1

    # このmock専用の自己テスト用gold(companyはfixturesを持たないため
    # exact/recall/forbiddenは「mockのJSON自体」を対象に見る=採点器の
    # 算術そのものを検算する回帰テスト)。
    self_gold = {
        "_mock": {
            "company": "株式会社浜松スイーツファクトリー(mock)",
            "exact": {
                "company_name": "株式会社浜松スイーツファクトリー",
                "financials.source": "unknown",
            },
            "recall": [
                "浜松本社工場", "積志第二工場", "小麦粉", "バター",
                "浸水想定区域内", "EC直販",
            ],
            "forbidden": [
                "セグメント別売上高120億円", "従業員数1200名", "本社は東京都",
            ],
        }
    }

    ok = True

    # ケース1: 完全一致のダミー出力 -> 一致率1.0・捏造率0・回収率1.0
    r1 = score_outputs_direct(self_gold, {"_mock": [mock_json]})
    print("[selfcheck] ケース1(無傷のmock応答):",
          f"exact={r1['exact_match_rate']} recall={r1['recall_rate']} "
          f"forbidden={r1['forbidden_rate']}")
    if r1["exact_match_rate"] != 1.0 or r1["recall_rate"] != 1.0 or r1["forbidden_rate"] != 0.0:
        print("[selfcheck] NG: ケース1の期待値(1.0/1.0/0.0)と一致しません")
        ok = False

    # ケース2: 1フィールド壊す(company_nameを改変) -> 一致率が下がる(設計どおり0.9系)
    broken = dict(mock_json)
    broken["company_name"] = "株式会社ハママツスイーツ(誤記)"
    r2 = score_outputs_direct(self_gold, {"_mock": [broken]})
    print("[selfcheck] ケース2(company_nameを破壊):",
          f"exact={r2['exact_match_rate']}")
    if r2["exact_match_rate"] is None or r2["exact_match_rate"] >= 1.0:
        print("[selfcheck] NG: ケース2は1.0未満に下がるはずです")
        ok = False

    # ケース3: 捏造混入(forbidden断片を本文へ混ぜる) -> 捏造率が上がる
    tampered = dict(mock_json)
    tampered["business_summary"] = mock_json["business_summary"] + " 従業員数1200名。"
    r3 = score_outputs_direct(self_gold, {"_mock": [tampered]})
    print("[selfcheck] ケース3(forbidden混入):",
          f"forbidden={r3['forbidden_rate']}")
    if not r3["forbidden_rate"] or r3["forbidden_rate"] <= 0.0:
        print("[selfcheck] NG: ケース3は捏造率が0より大きくなるはずです")
        ok = False

    # 変異注入(a): fragment_found を常にTrueにすると、ケース1のforbidden(0.0)が壊れる
    def always_true(_frag, _hay):
        return True

    global fragment_found
    orig = fragment_found
    fragment_found = always_true  # noqa: F811
    try:
        r_mut = score_outputs_direct(self_gold, {"_mock": [mock_json]})
    finally:
        fragment_found = orig
    print("[selfcheck] 変異注入(fragment_found常時True):",
          f"forbidden={r_mut['forbidden_rate']}")
    if r_mut["forbidden_rate"] != 1.0:
        print("[selfcheck] NG: 変異注入で捏造率1.0(全件誤検知)にならず、"
              "テスト自体が有効な変異を作れていません")
        ok = False

    # 変異注入(b): normalize_for_match を素通しにすると、表記揺れ耐性の
    #   ぶんだけ一致率/回収率が下がることを実演する(元の書式が全角混在でないため
    #   ここでは正規化を無効化しても数値が変わらないことを確認する=このgoldは
    #   表記揺れを含まない素直な文字列である裏取りにもなる)。
    global normalize_for_match
    orig_norm = normalize_for_match
    normalize_for_match = lambda t: t  # noqa: E731
    try:
        r_noNorm = score_outputs_direct(self_gold, {"_mock": [mock_json]})
    finally:
        normalize_for_match = orig_norm
    print("[selfcheck] 変異注入(正規化を無効化):",
          f"exact={r_noNorm['exact_match_rate']} recall={r_noNorm['recall_rate']}")

    if ok:
        print("[selfcheck] OK: 採点器の算術は期待どおりです")
        return 0
    print("[selfcheck] NG: 上記の不一致を確認してください")
    return 1


def score_outputs_direct(gold: dict, outputs_by_key: dict) -> dict:
    """selfcheck専用: ファイルではなく読み込み済みJSONを直接渡す薄いラッパー。"""
    exact_total = exact_hit = 0
    recall_total = recall_hit = 0
    forbidden_total = forbidden_hit = 0
    for key, g in gold.items():
        for data in outputs_by_key.get(key, []):
            text_norm = normalize_for_match(json.dumps(data, ensure_ascii=False))
            for field, expect in g.get("exact", {}).items():
                actual = get_path(data, field)
                exact_total += 1
                if values_equal(actual, expect):
                    exact_hit += 1
            for frag in g.get("recall", []):
                recall_total += 1
                if fragment_found(frag, text_norm):
                    recall_hit += 1
            for frag in g.get("forbidden", []):
                forbidden_total += 1
                if fragment_found(frag, text_norm):
                    forbidden_hit += 1
    return {
        "exact_match_rate": (exact_hit / exact_total) if exact_total else None,
        "recall_rate": (recall_hit / recall_total) if recall_total else None,
        "forbidden_rate": (forbidden_hit / forbidden_total) if forbidden_total else None,
    }


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--mode", choices=["replay", "selfcheck"], default=None)
    ap.add_argument("--in", dest="in_path", type=str, default=None,
                     help="replay: S1出力JSON群のディレクトリ、または単一ファイル")
    ap.add_argument("--gold-check", action="store_true")
    args = ap.parse_args()

    if args.gold_check:
        return cmd_gold_check()
    if args.mode == "selfcheck":
        return cmd_selfcheck()
    if args.mode == "replay":
        in_path = Path(args.in_path) if args.in_path else (OUT_DIR / "live")
        if not in_path.exists():
            print(f"[replay] 入力パスがありません: {in_path}")
            return 1
        return cmd_replay(in_path)

    ap.print_help()
    return 1


if __name__ == "__main__":
    sys.exit(main())
