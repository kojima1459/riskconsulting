#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
validate_check.py - 15章§11の検証ルール表 <-> modValidate <-> modTestsPure の照合(17章§4-2)

================================================================================
役割(17章§4-2の4本目。prompt_diff.py / sheet_check.py / 19章enum照合に続く検問):

    15章§11「検証ルール ケースID一覧」を**唯一の正**としてパースし、そこから
    導いた全ケースIDに対して、

      (a) `src/app/modValidate*.bas` に**そのケースの実装がある**
          - 判定が「不合格」「警告」のケース: `"[ケースID] ` で始まる文字列
            リテラルが実装に**必ず1本以上**ある(戻り値の各行はこの形で始まる。
            15章§0 原則10)。
          - 判定が「合格」のケース(V-S2C-05 / V-S3C-05): エラー文を持たない
            ので、逆に `[ケースID] ` で始まるリテラルが**あってはならない**。
            ただしケースIDそのものは実装(コメント)に現れていること
            = 「読み落としていない」ことの最低限の証拠を求める。
      (b) `src/test/modTestsPure*.bas` に**1テスト1ケース**でテストがある
          - テスト名(Check系呼び出しの第1引数の文字列リテラル)にケースIDが
            含まれ、**1ケースにつきちょうど1本**であること。
          - 1つのテスト名が2つ以上のケースIDを含まないこと(1テスト1ケース)。
      (c) 件数の一致
          - §11の各行の「(n件)」と展開したID数
          - §11の「合計64件(不合格n件/警告n件/合格判定n件)」と実際の内訳
          - 各Check節の検証ルール表に載っているID集合と§11の展開結果
          - ケースIDを含むテスト名の総数 = ケース総数
      (d) エラー文テンプレの一字一句(既定でON。--no-templates で外す)
          - 各Check節の表の `エラー文テンプレ` を `{...}` で切った**固定部**が、
            実装のソースにそのまま現れること。文言が仕様から漂流したら落ちる。
          - 照合の範囲は**当該ケースIDの近傍に限る**(裁定書7 C-11)。実装を
            1本の文字列に連結して素の部分文字列検索を掛けると、別ケースが
            偶然もつ同一の固定部が肩代わりして文言の漂流を検出できない
            (実測: V-S3-01「件です(3件固定)」を「(3件確定)」へ壊しても
             V-S3C-01 の同一文言が吸収して緑のまま通った)。近傍＝当該IDが
            現れた行から**次にケースIDが現れる行の直前まで**(行継続 `_` で
            次行へ続くエラー文もこの区間に入る)。同じIDの区間は連結する。
      (e) 自己整合(裁定書7 C-12。照合器自身の骨抜きを検出する)
          - (a)(b)(d) の各ループが**64件すべてを検査したこと**を件数で確認する
            (特定ケースを `continue` で飛ばす改変が入ると落ちる)。
          - (d) の照合片が0件なら不合格(照合しなかったことを緑にしない)。

    未統合段階(modValidate は出来たがテストはまだ、の並行作業)では
    `--no-tests` で (b) と (b)由来の件数照合を省ける。**出荷前の検問では
    必ず --no-tests 無しで回すこと**(17章§4-2はテスト側の存在まで求めている)。

使い方:
    python3 tools/validate_check.py
    python3 tools/validate_check.py --no-tests       # (b)を省く(実装先行時)
    python3 tools/validate_check.py --no-templates   # (d)を省く
    exit code: 0 = 全一致 / 1 = 不一致あり
================================================================================
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
SPEC_15 = REPO_ROOT / "docs" / "spec" / "15_プロンプトとJSONスキーマ.md"
IMPL_GLOB = "modValidate*.bas"
IMPL_DIR = REPO_ROOT / "src" / "app"
TEST_GLOB = "modTestsPure*.bas"
TEST_DIR = REPO_ROOT / "src" / "test"

SECTION_11_HEAD = "## 11. 検証ルール ケースID一覧"

# §11の表の行: | CheckS1 | V-S1-01 ～ V-S1-11(11件) | 01/02/... | 04/... | - |
ROW_11 = re.compile(
    r"^\|\s*(Check\w+)\s*\|\s*(V-[A-Z0-9]+)-(\d+)\s*[～~]\s*(V-[A-Z0-9]+)-(\d+)\s*"
    r"[（(](\d+)件[)）]\s*\|([^|]*)\|([^|]*)\|([^|]*)\|"
)
# 合計行: **合計64件**(不合格52件 / 警告10件 / 合格判定2件)
TOTAL_11 = re.compile(
    r"合計(\d+)件\*{0,2}[（(]不合格(\d+)件\s*/\s*警告(\d+)件\s*/\s*合格判定(\d+)件"
)
# 各Check節の表の行: | V-S1-01 | 対象キー | 条件 | 不合格 | `[V-S1-01] ...` |
ROW_RULE = re.compile(
    r"^\|\s*(V-[A-Z0-9]+-\d+)\s*\|([^|]*)\|(.*)\|\s*(不合格|警告|合格)\s*\|(.*)\|\s*$"
)
CASE_ID = re.compile(r"V-(?:S1|S2|S3|S4|PF|S2C|S3C)-\d{2}")
# テスト名 = Check / Chk* 呼び出しの第1引数の文字列リテラル
TEST_NAME = re.compile(r"\b(?:Check|Chk[A-Za-z]*)\s*\(?\s*\"((?:[^\"]|\"\")*)\"")


class Report:
    def __init__(self) -> None:
        self.errors: list[str] = []
        self.notes: list[str] = []

    def err(self, msg: str) -> None:
        self.errors.append(msg)

    def note(self, msg: str) -> None:
        self.notes.append(msg)


def split_nums(cell: str) -> list[str]:
    """判定列のセルから 2桁の番号だけを拾う(注記の「(...)」以降は見ない)。"""
    head = re.split(r"[（(]", cell.strip(), maxsplit=1)[0]
    if head.strip() in ("-", ""):
        return []
    return re.findall(r"\b(\d{2})\b", head)


def parse_section_11(text: str, rep: Report):
    """§11の表を展開する。戻り値: (順序つきID一覧, ID->判定, Check関数->ID一覧)"""
    pos = text.find(SECTION_11_HEAD)
    if pos < 0:
        rep.err(f"15章に「{SECTION_11_HEAD}」が見つかりません(章の見出しが変わった?)")
        return [], {}, {}
    body = text[pos:]

    ids: list[str] = []
    verdict: dict[str, str] = {}
    by_check: dict[str, list[str]] = {}

    for line in body.splitlines():
        m = ROW_11.match(line)
        if not m:
            continue
        check, pre1, lo, pre2, hi, declared = (m.group(1), m.group(2), m.group(3),
                                               m.group(4), m.group(5), int(m.group(6)))
        fail_c, warn_c, pass_c = m.group(7), m.group(8), m.group(9)
        if pre1 != pre2:
            rep.err(f"§11 {check}: 範囲の接頭辞が食い違っています({pre1} / {pre2})")
            continue
        width = len(lo)
        row_ids = [f"{pre1}-{str(n).zfill(width)}" for n in range(int(lo), int(hi) + 1)]
        if len(row_ids) != declared:
            rep.err(f"§11 {check}: 範囲 {pre1}-{lo}～{hi} は{len(row_ids)}件ですが"
                    f"表の宣言は{declared}件です")

        seen: dict[str, str] = {}
        for label, cell in (("不合格", fail_c), ("警告", warn_c), ("合格", pass_c)):
            for nn in split_nums(cell):
                cid = f"{pre1}-{nn}"
                if cid in seen:
                    rep.err(f"§11 {check}: {cid} が「{seen[cid]}」と「{label}」に重複しています")
                seen[cid] = label
        missing = [c for c in row_ids if c not in seen]
        extra = [c for c in seen if c not in row_ids]
        if missing:
            rep.err(f"§11 {check}: 判定列に現れないケースID: {', '.join(missing)}"
                    "(判定は3値。どれか1列に必ず載ること)")
        if extra:
            rep.err(f"§11 {check}: 範囲外のケースIDが判定列にあります: {', '.join(sorted(extra))}")

        for cid in row_ids:
            ids.append(cid)
            verdict[cid] = seen.get(cid, "?")
        by_check[check] = row_ids

    if not ids:
        rep.err("§11の表を1行も読み取れませんでした(表の書式が変わった?)")
        return ids, verdict, by_check

    m = TOTAL_11.search(body)
    if not m:
        rep.err("§11の合計行(「合計n件(不合格n件 / 警告n件 / 合格判定n件)」)が読めません")
    else:
        total, n_fail, n_warn, n_pass = (int(m.group(1)), int(m.group(2)),
                                         int(m.group(3)), int(m.group(4)))
        got = {
            "total": len(ids),
            "不合格": sum(1 for v in verdict.values() if v == "不合格"),
            "警告": sum(1 for v in verdict.values() if v == "警告"),
            "合格": sum(1 for v in verdict.values() if v == "合格"),
        }
        if got["total"] != total:
            rep.err(f"(c) §11 合計{total}件の宣言に対し、展開したケースIDは{got['total']}件です")
        if (got["不合格"], got["警告"], got["合格"]) != (n_fail, n_warn, n_pass):
            rep.err(f"(c) §11 内訳の宣言(不合格{n_fail}/警告{n_warn}/合格{n_pass})に対し、"
                    f"実際は(不合格{got['不合格']}/警告{got['警告']}/合格{got['合格']})です")
    return ids, verdict, by_check


def parse_rule_tables(text: str, rep: Report):
    """各Check節の検証ルール表から ID -> (判定, エラー文テンプレ) を拾う。"""
    end = text.find(SECTION_11_HEAD)
    body = text[:end] if end > 0 else text
    out: dict[str, tuple[str, str]] = {}
    for line in body.splitlines():
        m = ROW_RULE.match(line)
        if not m:
            continue
        cid, verdict, tmpl_cell = m.group(1), m.group(4), m.group(5)
        tm = re.search(r"`([^`]*)`", tmpl_cell)
        tmpl = tm.group(1) if tm else ""
        if cid in out:
            rep.err(f"各Check節: {cid} の行が2回定義されています")
        out[cid] = (verdict, tmpl)
    return out


def read_sources(directory: Path, pattern: str) -> dict[str, str]:
    return {p.name: p.read_text(encoding="utf-8", errors="replace")
            for p in sorted(directory.glob(pattern))}


def case_regions(srcs: dict[str, str]) -> dict[str, str]:
    """ケースIDごとの「近傍」区間を作る((d)の照合範囲。裁定書7 C-11)。

    各ファイルを行単位で見て、ケースIDが現れた行から**次にケースIDが現れる行の
    直前まで**をそのIDの区間とする。同じIDの区間は連結する。行継続(`_`)で
    次行へ続くエラー文は次のケースID行までに入るので同じ区間に収まる。
    """
    regions: dict[str, list[str]] = {}
    for text in srcs.values():
        lines = text.split("\n")
        hits = [(i, sorted(set(CASE_ID.findall(line)))) for i, line in enumerate(lines)]
        anchors = [(i, found) for i, found in hits if found]
        for pos, (i, found) in enumerate(anchors):
            end = anchors[pos + 1][0] if pos + 1 < len(anchors) else len(lines)
            seg = "\n".join(lines[i:end])
            for cid in found:
                regions.setdefault(cid, []).append(seg)
    return {cid: "\n".join(parts) for cid, parts in regions.items()}


def check_impl(ids: list[str], verdict: dict[str, str], rep: Report) -> None:
    srcs = read_sources(IMPL_DIR, IMPL_GLOB)
    if not srcs:
        rep.err(f"(a) {IMPL_DIR}/{IMPL_GLOB} が1本もありません(modValidate 未実装)")
        return
    rep.note(f"(a) 実装: {', '.join(srcs)}")
    joined = "\n".join(srcs.values())
    # 実装中の "[V-xxx] " で始まる文字列リテラルを集める。
    lit_ids = set(re.findall(r'"\[(V-(?:S1|S2|S3|S4|PF|S2C|S3C)-\d{2})\]\s', joined))
    examined = 0
    for cid in ids:
        examined += 1
        has_lit = cid in lit_ids
        mentioned = cid in joined
        if verdict.get(cid) == "合格":
            if has_lit:
                rep.err(f"(a) {cid} は判定「合格」でエラー文を持たないのに、実装に "
                        f'"[{cid}] " で始まるリテラルがあります(戻り値に現れてはいけない)')
            if not mentioned:
                rep.err(f"(a) {cid}(判定「合格」)が実装に1度も現れません"
                        "(スキップ判定の実装とコメントで存在を示すこと)")
        else:
            if not has_lit:
                rep.err(f"(a) {cid} のエラー文が実装にありません"
                        f'("[{cid}] " で始まる文字列リテラルが1本も無い)')
    stray = sorted(lit_ids - set(ids))
    if stray:
        rep.err(f"(a) §11に無いケースIDのエラー文が実装にあります: {', '.join(stray)}")
    if examined != len(ids):
        rep.err(f"(e) 自己整合: (a)の照合ループが{examined}件しか回っていません"
                f"(§11の展開は{len(ids)}件)。ケースを飛ばす改変が入っています")


def check_templates(rules: dict[str, tuple[str, str]], ids: list[str],
                    verdict: dict[str, str], rep: Report) -> None:
    srcs = read_sources(IMPL_DIR, IMPL_GLOB)
    regions = case_regions(srcs)
    checked = 0
    visited = 0
    with_tmpl = 0
    for cid in ids:
        visited += 1
        entry = rules.get(cid)
        if entry is None:
            continue
        tmpl = entry[1]
        if not tmpl:
            continue
        with_tmpl += 1
        # (d)は当該ケースIDの近傍だけを見る(裁定書7 C-11)。全ソース連結に対する
        # 素の部分文字列検索だと、別ケースの同一文言が肩代わりして漂流を見逃す。
        near = regions.get(cid, "")
        for frag in re.split(r"\{[^}]*\}", tmpl):
            if len(frag.strip()) < 2:
                continue
            if frag not in near:
                rep.err(f"(d) {cid} のエラー文テンプレが実装と一致しません。"
                        f"15章の固定部「{frag}」が modValidate*.bas の "
                        f"{cid} の近傍にありません")
            checked += 1
    rep.note(f"(d) エラー文テンプレの固定部 {checked} 片を {len(regions)} 件の"
             "ケースID近傍で照合")

    # --- (e) 自己整合(裁定書7 C-12) ---
    if visited != len(ids):
        rep.err(f"(e) 自己整合: (d)の照合ループが{visited}件しか回っていません"
                f"(§11の展開は{len(ids)}件)。ケースを飛ばす改変が入っています")
    if checked == 0:
        rep.err("(e) 自己整合: (d)の照合片が0件です"
                "(エラー文テンプレを1片も照合していない状態を緑にしない)")
    expect_tmpl = sum(1 for c in ids if verdict.get(c) != "合格")
    if with_tmpl != expect_tmpl:
        rep.err(f"(e) 自己整合: エラー文テンプレを持つケースが{with_tmpl}件ですが、"
                f"§11で「不合格・警告」のケースは{expect_tmpl}件です"
                "(判定3値のうちエラー文を持つのはこの2値。各Check節の表の"
                "テンプレ欄が欠けているか、判定列が食い違っています)")


def check_tests(ids: list[str], rep: Report) -> None:
    srcs = read_sources(TEST_DIR, TEST_GLOB)
    if not srcs:
        rep.err(f"(b) {TEST_DIR}/{TEST_GLOB} が1本もありません")
        return
    rep.note(f"(b) テスト: {', '.join(srcs)}")

    names: list[tuple[str, str]] = []   # (ファイル名, テスト名)
    for fname, text in srcs.items():
        for m in TEST_NAME.finditer(text):
            names.append((fname, m.group(1).replace('""', '"')))

    hit: dict[str, list[str]] = {cid: [] for cid in ids}
    tagged = 0
    for fname, name in names:
        found = sorted(set(CASE_ID.findall(name)))
        if not found:
            continue
        tagged += 1
        if len(found) > 1:
            rep.err(f"(b) 1テスト1ケースに反します({fname}): テスト名に "
                    f"{', '.join(found)} が同居「{name}」")
        for cid in found:
            if cid in hit:
                hit[cid].append(f"{fname}: {name}")
            else:
                rep.err(f"(b) §11に無いケースID {cid} のテストがあります({fname}: {name})")

    if tagged == 0:
        raw = sum(len(CASE_ID.findall(t)) for t in srcs.values())
        if raw > 0:
            rep.note(f"(b) 参考: テストソースにケースIDは{raw}箇所ありますが、"
                     "テスト名(Check / Chk* 呼び出しの第1引数の文字列リテラル)としては"
                     "1件も拾えませんでした。判定ヘルパの名前が Check / Chk* 以外なら"
                     "本スクリプトの TEST_NAME を合わせてください")

    examined = 0
    for cid in ids:
        examined += 1
        n = len(hit[cid])
        if n == 0:
            rep.err(f"(b) {cid} のテストがありません"
                    "(modTestsPure* にテストを1本置き、テスト名にケースIDを含めること)")
        elif n > 1:
            rep.err(f"(b) {cid} のテストが{n}本あります(1ケース1テスト): "
                    + " / ".join(hit[cid]))
    if tagged != len(ids):
        rep.err(f"(c) ケースIDを含むテスト名は{tagged}本ですが、ケース総数は{len(ids)}件です")
    if examined != len(ids):
        rep.err(f"(e) 自己整合: (b)の照合ループが{examined}件しか回っていません"
                f"(§11の展開は{len(ids)}件)。ケースを飛ばす改変が入っています")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="15章§11 <-> modValidate <-> modTestsPure の照合(17章§4-2)")
    parser.add_argument("--no-tests", action="store_true",
                        help="(b)テスト側の照合を省く(modValidate 実装先行時のみ)")
    parser.add_argument("--no-templates", action="store_true",
                        help="(d)エラー文テンプレの一字一句照合を省く")
    args = parser.parse_args()

    rep = Report()
    if not SPEC_15.exists():
        print(f"[validate_check] 15章が見つかりません: {SPEC_15}")
        return 1
    text = SPEC_15.read_text(encoding="utf-8")

    ids, verdict, by_check = parse_section_11(text, rep)
    rules = parse_rule_tables(text, rep)

    if ids:
        # (c) 各Check節の表 <-> §11 の突合。
        rule_ids = set(rules)
        expect_err = {c for c in ids if verdict.get(c) != "合格"}
        for cid in sorted(expect_err - rule_ids):
            rep.err(f"(c) {cid} が§11にありますが、各Check節の検証ルール表に行がありません")
        for cid in sorted(rule_ids - set(ids)):
            rep.err(f"(c) {cid} が各Check節の表にありますが、§11の一覧にありません")
        for cid in sorted(rule_ids & set(ids)):
            if rules[cid][0] != verdict.get(cid):
                rep.err(f"(c) {cid} の判定が §11「{verdict.get(cid)}」と"
                        f"各Check節「{rules[cid][0]}」で食い違います")

        check_impl(ids, verdict, rep)
        if not args.no_templates:
            check_templates(rules, ids, verdict, rep)
        if args.no_tests:
            rep.note("(b) --no-tests のためテスト側の照合をスキップしました"
                     "(出荷前の検問では必ず外して回すこと)")
        else:
            check_tests(ids, rep)

    print("=" * 78)
    print("validate_check レポート - 15章§11 <-> modValidate <-> modTestsPure")
    print("=" * 78)
    if ids:
        summary = " / ".join(f"{k}:{len(v)}件" for k, v in by_check.items())
        print(f"§11から展開したケース: 計{len(ids)}件  ({summary})")
    for note in rep.notes:
        print(f"  {note}")
    if rep.errors:
        print("\n[不一致]")
        for e in rep.errors:
            print(f"  ERROR {e}")
    print("\n" + "-" * 78)
    if rep.errors:
        print(f"結果: NG(exit code 1) - {len(rep.errors)}件の不一致")
    else:
        print("結果: OK(exit code 0)")
    print("-" * 78)
    return 1 if rep.errors else 0


if __name__ == "__main__":
    sys.exit(main())
