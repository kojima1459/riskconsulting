#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""kb_check.py - ナレッジブック実データの純テスト(D-5. 班D)

================================================================================
何を見るか(裁定書「実装班D」依頼 D-5):
    build/sheets_kb.json の seed_rows(業種マスタ・リスクライブラリ・事故事例・
    判断基準)について、以下4条件を機械で照合する。

    (a) 全 industry_code が業種マスタに実在する
        (リスクライブラリ.industry_code / 事故事例.industry_code)
    (b) enum内であること
        (リスクライブラリ.category / typical_freq / typical_impact、
         事故事例.category、判断基準.rule_class / workaround)
    (c) ID の重複なし(risk_lib_id / inc_id / rule_id、各シート内)
    (d) 文字数上限
        (リスクライブラリ.typical_scenario<=300 / check_points<=200、
         事故事例.headline<=60 / cause<=150)
    (e) メニュー一覧: target_categories の各値が enum `category` 内 /
        menu_id 重複なし(裁定書46 F-8)
    (f) メニュー種目対応: menu_id がメニュー一覧に、line_id が種目マスタに
        実在 / 同じ対(menu_id, line_id)の重複なし(裁定書46 F-8。対応表は
        多対多なので同じ menu_id・line_id 単体の重複は許すが、同一の対の
        重複登録だけを見る)
    (g) 種目マスタ: line_id 重複なし(裁定書46 F-8)

    enum の正は build/sheets_kb.json の `enums`(19章§3)。空欄は「値源が無い」
    ことの明示であり、enum検査の対象外とする(typical_freq/typical_impact/
    priority/loss_scale は空欄を許容する列)。

gate_count.py の契約(W15 §3 X3-3):
    実際に見た行数だけを `Checked.record` に積む。0行の検査項目は要点行に
    現れず、`--selftest` はその場で1行を壊して赤になることを確認する
    (敵対的検証。件数は実測でリテラル定数を含まない)。

使い方:
    python3 tools/kb_check.py                 # 実データを検査
    python3 tools/kb_check.py --selftest      # 検出器の自己テスト(変異注入)
    exit code: 0 = OK / 1 = 違反あり / 2 = 自己テスト失敗
================================================================================
"""
from __future__ import annotations

import argparse
import copy
import json
import sys
from pathlib import Path

TOOLS_DIR = Path(__file__).resolve().parent
REPO_ROOT = TOOLS_DIR.parent
KB_JSON = REPO_ROOT / "build" / "sheets_kb.json"

sys.path.insert(0, str(TOOLS_DIR))
import gate_count  # noqa: E402  (要点行の契約。W15 §3 X3-3)


def _sheet(data: dict, name: str) -> dict:
    for s in data["sheets"]:
        if s.get("name") == name:
            return s
    raise KeyError(name)


def _col_index(sheet: dict, col: str) -> int:
    for i, c in enumerate(sheet["columns"]):
        if c["name"] == col:
            return i
    raise KeyError(col)


def _rows_by_col(sheet: dict, cols: list) -> list:
    """seed_rows を [{col: value, ...}, ...] へ変換する。"""
    idx = {c: _col_index(sheet, c) for c in cols}
    out = []
    for row in sheet.get("seed_rows") or []:
        out.append({c: row[i] for c, i in idx.items()})
    return out


def run_checks(data: dict, checked: "gate_count.Checked") -> list:
    """4条件を照合し、違反メッセージのリストを返す(空なら合格)。"""
    errors: list = []
    enums = data["enums"]
    by = {s["name"]: s for s in data["sheets"]}

    industry_codes = {row[0] for row in by["業種マスタ"].get("seed_rows") or []}
    checked.record("業種マスタ", len(industry_codes))

    risk_rows = _rows_by_col(
        by["リスクライブラリ"],
        ["risk_lib_id", "industry_code", "category", "typical_freq",
         "typical_impact", "typical_scenario", "check_points"],
    )
    inc_rows = _rows_by_col(
        by["事故事例"],
        ["inc_id", "industry_code", "category", "headline", "cause"],
    )
    rule_rows = _rows_by_col(
        by["判断基準"], ["rule_id", "rule_class", "workaround"],
    )

    # --- (a) industry_code が業種マスタに実在 -------------------------------
    n_a = 0
    for r in risk_rows:
        n_a += 1
        if r["industry_code"] not in industry_codes:
            errors.append(
                "(a) リスクライブラリ %s: industry_code %r が業種マスタに無い"
                % (r["risk_lib_id"], r["industry_code"]))
    for r in inc_rows:
        n_a += 1
        if r["industry_code"] not in industry_codes:
            errors.append(
                "(a) 事故事例 %s: industry_code %r が業種マスタに無い"
                % (r["inc_id"], r["industry_code"]))
    checked.record("industry_code実在", n_a)

    # --- (b) enum内 ----------------------------------------------------------
    cat_set = set(enums["category"])
    freq_set = set(enums["frequency"])
    impact_set = set(enums["impact"])
    rule_class_set = set(enums["rule_class"])
    workaround_set = set(enums["rule_workaround"])

    n_b = 0
    for r in risk_rows:
        n_b += 1
        if r["category"] not in cat_set:
            errors.append("(b) リスクライブラリ %s: category %r が enum外"
                          % (r["risk_lib_id"], r["category"]))
        n_b += 1
        if r["typical_freq"] and r["typical_freq"] not in freq_set:
            errors.append("(b) リスクライブラリ %s: typical_freq %r が enum外"
                          % (r["risk_lib_id"], r["typical_freq"]))
        n_b += 1
        if r["typical_impact"] and r["typical_impact"] not in impact_set:
            errors.append("(b) リスクライブラリ %s: typical_impact %r が enum外"
                          % (r["risk_lib_id"], r["typical_impact"]))
    for r in inc_rows:
        n_b += 1
        if r["category"] not in cat_set:
            errors.append("(b) 事故事例 %s: category %r が enum外"
                          % (r["inc_id"], r["category"]))
    for r in rule_rows:
        n_b += 1
        if r["rule_class"] not in rule_class_set:
            errors.append("(b) 判断基準 %s: rule_class %r が enum外"
                          % (r["rule_id"], r["rule_class"]))
        n_b += 1
        if r["workaround"] not in workaround_set:
            errors.append("(b) 判断基準 %s: workaround %r が enum外"
                          % (r["rule_id"], r["workaround"]))
    checked.record("enum照合", n_b)

    # --- (c) ID重複なし --------------------------------------------------------
    n_c = 0
    for label, rows, key in (
        ("risk_lib_id", risk_rows, "risk_lib_id"),
        ("inc_id", inc_rows, "inc_id"),
        ("rule_id", rule_rows, "rule_id"),
    ):
        seen: dict = {}
        for r in rows:
            n_c += 1
            seen[r[key]] = seen.get(r[key], 0) + 1
        dups = [k for k, n in seen.items() if n > 1]
        for k in dups:
            errors.append("(c) %s %r が重複(%d件)" % (label, k, seen[k]))
    checked.record("ID重複検査", n_c)

    # --- (d) 文字数上限 ----------------------------------------------------
    LIMITS = (
        ("リスクライブラリ", "typical_scenario", 300),
        ("リスクライブラリ", "check_points", 200),
        ("事故事例", "headline", 60),
        ("事故事例", "cause", 150),
    )
    n_d = 0
    for sheet_name, col, limit in LIMITS:
        rows = risk_rows if sheet_name == "リスクライブラリ" else inc_rows
        id_col = "risk_lib_id" if sheet_name == "リスクライブラリ" else "inc_id"
        for r in rows:
            n_d += 1
            val = r[col] or ""
            if len(val) > limit:
                errors.append(
                    "(d) %s %s: %s が%d字を超過(%d字)"
                    % (sheet_name, r[id_col], col, limit, len(val)))
    checked.record("文字数上限", n_d)

    # --- (e) メニュー一覧: target_categories が enum内 / menu_id重複なし ----
    menu_rows = _rows_by_col(
        by["メニュー一覧"], ["menu_id", "target_categories"],
    )
    n_e = 0
    for r in menu_rows:
        n_e += 1
        cats = [c for c in (r["target_categories"] or "").split(";") if c]
        for c in cats:
            n_e += 1
            if c not in cat_set:
                errors.append(
                    "(e) メニュー一覧 %s: target_categories %r が enum外"
                    % (r["menu_id"], c))
    seen_menu: dict = {}
    for r in menu_rows:
        n_e += 1
        seen_menu[r["menu_id"]] = seen_menu.get(r["menu_id"], 0) + 1
    for k, n in seen_menu.items():
        if n > 1:
            errors.append("(e) メニュー一覧 menu_id %r が重複(%d件)" % (k, n))
    checked.record("メニュー一覧照合", n_e)

    # --- (f) メニュー種目対応: menu_id/line_id実在 / 対の重複なし ----------
    menu_ids = {r["menu_id"] for r in menu_rows}
    line_ids = {row[0] for row in by["種目マスタ"].get("seed_rows") or []}
    map_rows = _rows_by_col(by["メニュー種目対応"], ["menu_id", "line_id"])
    n_f = 0
    for r in map_rows:
        n_f += 1
        if r["menu_id"] not in menu_ids:
            errors.append(
                "(f) メニュー種目対応: menu_id %r がメニュー一覧に無い"
                % r["menu_id"])
        n_f += 1
        if r["line_id"] not in line_ids:
            errors.append(
                "(f) メニュー種目対応: line_id %r が種目マスタに無い"
                % r["line_id"])
    seen_pair: dict = {}
    for r in map_rows:
        n_f += 1
        pair = (r["menu_id"], r["line_id"])
        seen_pair[pair] = seen_pair.get(pair, 0) + 1
    for (mid, lid), n in seen_pair.items():
        if n > 1:
            errors.append(
                "(f) メニュー種目対応 (%s, %s) が重複(%d件)" % (mid, lid, n))
    checked.record("メニュー種目対応照合", n_f)

    # --- (g) 種目マスタ: line_id重複なし -------------------------------
    line_rows_all = [row[0] for row in by["種目マスタ"].get("seed_rows") or []]
    n_g = 0
    seen_line: dict = {}
    for lid in line_rows_all:
        n_g += 1
        seen_line[lid] = seen_line.get(lid, 0) + 1
    for k, n in seen_line.items():
        if n > 1:
            errors.append("(g) 種目マスタ line_id %r が重複(%d件)" % (k, n))
    checked.record("種目マスタ照合", n_g)

    return errors


# ---------------------------------------------------------------------------
# 敵対的検証(--selftest): 各条件を1件だけ壊して赤になることを実演する。
# ---------------------------------------------------------------------------
def _mutate_and_expect_fail(data: dict, mutate, label: str) -> bool:
    mutated = copy.deepcopy(data)
    mutate(mutated)
    checked = gate_count.Checked()
    errors = run_checks(mutated, checked)
    ok = len(errors) > 0
    print("  [mutate:%s] %s -> %s" %
          (label, "違反を検出" if ok else "検出できず(NG)",
           errors[0] if errors else "(検出なし)"))
    return ok


def self_test() -> bool:
    data = json.loads(KB_JSON.read_text(encoding="utf-8"))
    by = {s["name"]: s for s in data["sheets"]}
    results = []

    # (a) industry_code を存在しないコードへ壊す
    def m_a(d):
        s = _sheet(d, "リスクライブラリ")
        s["seed_rows"][0][_col_index(s, "industry_code")] = "99-存在しない"
    results.append(_mutate_and_expect_fail(data, m_a, "industry_code不在"))

    # (b) category を enum外の値へ壊す
    def m_b(d):
        s = _sheet(d, "リスクライブラリ")
        s["seed_rows"][0][_col_index(s, "category")] = "no_such_category"
    results.append(_mutate_and_expect_fail(data, m_b, "category enum外"))

    # (b') rule_class を enum外の値へ壊す
    def m_b2(d):
        s = _sheet(d, "判断基準")
        s["seed_rows"][0][_col_index(s, "rule_class")] = "no_such_class"
    results.append(_mutate_and_expect_fail(data, m_b2, "rule_class enum外"))

    # (c) risk_lib_id を重複させる
    def m_c(d):
        s = _sheet(d, "リスクライブラリ")
        idx = _col_index(s, "risk_lib_id")
        s["seed_rows"][1][idx] = s["seed_rows"][0][idx]
    results.append(_mutate_and_expect_fail(data, m_c, "risk_lib_id重複"))

    # (d) headline を61字へ壊す(60字上限)
    def m_d(d):
        s = _sheet(d, "事故事例")
        idx = _col_index(s, "headline")
        s["seed_rows"][0][idx] = "あ" * 61
    results.append(_mutate_and_expect_fail(data, m_d, "headline上限超過"))

    # (e) メニュー一覧の target_categories を enum外の値へ壊す
    def m_e1(d):
        s = _sheet(d, "メニュー一覧")
        idx = _col_index(s, "target_categories")
        s["seed_rows"][0][idx] = "no_such_category"
    results.append(_mutate_and_expect_fail(data, m_e1, "target_categories enum外"))

    # (e') メニュー一覧の menu_id を重複させる
    def m_e2(d):
        s = _sheet(d, "メニュー一覧")
        idx = _col_index(s, "menu_id")
        if len(s["seed_rows"]) < 2:
            s["seed_rows"].append(list(s["seed_rows"][0]))
        s["seed_rows"][1][idx] = s["seed_rows"][0][idx]
    results.append(_mutate_and_expect_fail(data, m_e2, "menu_id重複"))

    # (f) メニュー種目対応の menu_id を存在しないIDへ壊す
    def m_f1(d):
        s = _sheet(d, "メニュー種目対応")
        idx = _col_index(s, "menu_id")
        s["seed_rows"][0][idx] = "M-9999"
    results.append(_mutate_and_expect_fail(data, m_f1, "メニュー種目対応menu_id不在"))

    # (f') メニュー種目対応の対を重複させる
    def m_f2(d):
        s = _sheet(d, "メニュー種目対応")
        if len(s["seed_rows"]) < 2:
            s["seed_rows"].append(list(s["seed_rows"][0]))
        s["seed_rows"][1] = list(s["seed_rows"][0])
    results.append(_mutate_and_expect_fail(data, m_f2, "メニュー種目対応の対の重複"))

    # (g) 種目マスタの line_id を重複させる
    def m_g(d):
        s = _sheet(d, "種目マスタ")
        idx = _col_index(s, "line_id")
        if len(s["seed_rows"]) < 2:
            s["seed_rows"].append(list(s["seed_rows"][0]))
        s["seed_rows"][1][idx] = s["seed_rows"][0][idx]
    results.append(_mutate_and_expect_fail(data, m_g, "種目マスタline_id重複"))

    # 参考: 正規データ自体は0件の違反であること(誤検出がないこと)も見る。
    checked0 = gate_count.Checked()
    baseline_errors = run_checks(data, checked0)
    baseline_ok = len(baseline_errors) == 0
    print("  [baseline] 正規データの違反件数: %d件 -> %s"
          % (len(baseline_errors), "OK" if baseline_ok else "NG(誤検出)"))
    if not baseline_ok:
        for e in baseline_errors[:10]:
            print("    -", e)

    return all(results) and baseline_ok


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--selftest", action="store_true",
                    help="検出器の自己テスト(変異注入)だけを行う")
    args = ap.parse_args()

    print("kb_check: ナレッジブック実データ純テスト(D-5)")

    if args.selftest:
        ok = self_test()
        print("結果: %s" % ("自己テストOK(敵対的検証4件+baseline)" if ok
                            else "自己テスト失敗(検出器が壊れています)"))
        return 0 if ok else 2

    if not KB_JSON.exists():
        print("[kb_check] 見つかりません: %s" % KB_JSON)
        return 1

    data = json.loads(KB_JSON.read_text(encoding="utf-8"))
    checked = gate_count.Checked()
    errors = run_checks(data, checked)

    if gate_count.report(checked, required=(
            "業種マスタ", "industry_code実在", "enum照合",
            "ID重複検査", "文字数上限",
            "メニュー一覧照合", "メニュー種目対応照合", "種目マスタ照合")):
        return 1

    if errors:
        print("結果: NG (違反 %d件)" % len(errors))
        for e in errors[:50]:
            print("  -", e)
        if len(errors) > 50:
            print("  ...ほか %d件" % (len(errors) - 50))
        return 1

    print("結果: OK (4条件とも違反0件)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
