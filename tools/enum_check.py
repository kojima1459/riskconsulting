#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
enum_check.py - 19章§3 enumレジストリ と modUICase の変換表の一致検査(17章§4-2)

================================================================================
位置づけ:
    17章§4-2「2. 一致検査」は4本の照合を求めている。
      (1) 15章 <-> modPrompts*/modSchemas    -> tools/prompt_diff.py
      (2) **19章 <-> enum変換表**            -> 本ツール(2本目の実体化)
      (3) 13章 <-> シート実体                -> tools/sheet_check.py
      (4) 15章§11 <-> modValidate のケース    -> tools/validate_check.py

    11章§5 は「日本語ラベル⇔enumの変換は modUICase の共通変換表(19章と一致必須)
    のみで行う」と定めている。**一致必須**が人の目視で担保されている限り、19章を
    1行直したときに実装が置き去りになったことは誰にも分からない。本ツールは
    その突合を機械化する。

検査するもの:
    [1] 19章§3の全行を「変換表に載せる行(REQUIRED)」と「載せない行(EXCLUDED)」
        へ**漏れなく**分類できること。19章にenum行が増えたのに本ツールの分類表へ
        足していなければ落ちる(=新しいenumが黙って変換表から漏れるのを防ぐ)。
    [2] REQUIRED の各行について、modUICase.EnumPairsCsv() が返すペア一覧が
        19章の機械値・日本語ラベル・**その並び順**と完全一致すること。
    [3] EnumPairsCsv に REQUIRED 以外のグループが混ざっていないこと。
    [4] 1グループ内で機械値・日本語ラベルがそれぞれ一意であること
        (重複すると EnumEn / EnumJa の逆引きが一意に決まらない)。
    [5] 値・ラベルに CSV の区切り(",")や改行が含まれないこと(表形式が壊れる)。

なぜ .bas を静的評価するのか:
    LibreOffice を起動して EnumPairsCsv を実行する手もあるが、それでは
    「実行できる環境がある」ことに検査が依存する。本ツールは 15章の突合
    (prompt_diff.py)と同じく `s = s & "..." & vbLf` 方式の関数本体を静的に
    評価する。評価できない書き方(条件分岐・ループ・未対応の項)は差分として
    数える=表を関数の中で組み立て直して検査を骨抜きにできない。

使い方:
    python3 tools/enum_check.py            # 照合(exit code: 0=一致 / 1=不一致)
    python3 tools/enum_check.py --dump     # 19章から期待されるCSV本文を出力
    python3 tools/enum_check.py --dump-bas # 上を .bas の連結文へ整形して出力
================================================================================
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

TOOLS_DIR = Path(__file__).resolve().parent
REPO_ROOT = TOOLS_DIR.parent
DEFAULT_SPEC = REPO_ROOT / "docs" / "spec" / "19_用語集とレジストリ.md"
DEFAULT_SRC = REPO_ROOT / "src" / "ui" / "modUICase.bas"
PAIRS_FUNC = "EnumPairsCsv"

# ==============================================================================
# [1] 19章§3の行 -> 変換表のグループ名
# ------------------------------------------------------------------------------
# 変換表に載せるのは「画面が日本語ラベルで見せる列」だけである(13章§2.10-§2.17
# が `日本語表示` と明記した列＋案件入力の属性欄＋HOMEの状態表示)。受信箱・判断
# 台帳・フィードバックは 13章§2.5-§2.7 が機械値のまま持つ列(入力規則も機械値)
# なので変換しない。ここが「載せる/載せない」の唯一の宣言であり、19章に行が
# 増えたら必ずどちらかへ足さないと本ツールが落ちる。
#
# 左辺=19章§3の1列目(項目キー。括弧書きの注記は落として突合する)。
# 右辺=modUICase の変換表グループ名。
# ==============================================================================
REQUIRED: dict[str, str] = {
    # --- 案件系(案件入力の属性欄・HOMEの状態表示。13章§2.10/§2.11) ---
    "case_type": "case_type",
    "dossier_tier": "dossier_tier",
    "channel": "channel",
    "kanji": "kanji",
    "bid": "bid",
    "reins": "reins",
    "status": "case_status",          # 19章の `status(案件)` 行
    "s4_variant": "s4_variant",
    "quality_mode": "quality_mode",
    # --- 分析系(S1-S3シートの「日本語表示」列。13章§2.12/§2.13/§2.14) ---
    "category": "risk_category",
    "insurability.transferability": "transferability",
    "risk.status": "risk_status",
    "emerging.horizon": "horizon",
    "frequency": "frequency",
    "impact": "impact",
    "source": "evidence_source",
    "gap_type": "gap_type",
    "proposal_kind": "proposal_kind",
    "locations.type": "location_type",
    "input_quality.aspect": "input_quality_aspect",
    "input_quality.status": "input_quality_status",
    "input_quality.overall": "input_quality_overall",
    "field_insights.tag": "field_insight_tag",
    # --- 壁打ち(13章§2.17 role: 日本語表示 自分/AI) ---
    "sparring.role": "sparring_role",
    # --- 受信箱の判定入力列(v2.5・裁定書9 B2/N5。13章§2.6) ---
    #     受信箱の他列は機械値のまま(EXCLUDED)だが、judge_to は利用者が選ぶ
    #     入力列であり19章§3が日本語ラベル(採択/条件付き保留/却下)を持つため載せる。
    "inbox.judge_to": "inbox_judge_to",
    # --- 受信箱の投函下書き行の入力列(v2.5.2・裁定書11 Q4。13章§2.6) ---
    #     source_kind も judge_to と同じく利用者が下書き行で選ぶ入力列であり、
    #     19章§3が日本語ラベル(部内投稿/現場の声/ウォッチ)を持つため載せる。
    #     受信箱のデータ行は従来どおり機械値のまま(store が書く)。
    "inbox.source_kind": "inbox_source_kind",
    # --- 判断台帳の入力列(v2.5.3・裁定書12 V7。13章§2.7) ---
    #     decision / result も利用者が起票・結果記録のときに選ぶ入力列であり、
    #     19章§3が日本語ラベルを持つため載せる(受信箱の judge_to と同作法)。
    #     store が書く確定行の値は従来どおり機械値のまま。
    "judgement.decision": "judge_decision",
    "judgement.result": "judge_result",
}

# 変換表に載せない行と、その理由(1行ずつ書く。理由の無い除外を作らない)。
EXCLUDED: dict[str, str] = {
    "data_key": "内部キー。19章§3が『表示なし』と明記(13章§2.2のcase_data列)",
    "failed_step": "案件一覧の内部列。13章§2.10のHOMEに表示欄が無い",
    "llm_transport": "configの値。HOMEの赤帯は文言で出す(13章§2.10 hm_transport_banner)",
    "mock_fault": "config の内部値。19章§3が『表示なし』と明記",
    "validate_result": "run_log の内部値(13章§2.4)。画面に列が無い",
    "s2c.issue_type": "批判JSONの内部enum。13章に対応する画面列が無い(入念モードの中間生成物)",
    "s3c.issue_type": "同上",
    "fb.event": "13章§2.5 フィードバックの event 列は機械値のまま持つ(入力規則も機械値)",
    "inbox.status": "同上(status 列)",
    "drop_type": "同上(drop_type 列)。19章§3のenum欄が `T0～T10` の範囲表記で個別値を列挙しない",
    "revive_tag": "同上(revive_tag 列)",
    "pf.survival": "同上(pf_survival 列。診断結果の要約列)",
    "pf.relation": "PF診断JSONの中の日本語enum。受信箱シートに列が無い",
    "pf.approach": "同上(3手)",
    "scheme.status": "ナレッジブック側の列(13章§3.4)。本体UIは描かない",
    "mech.layer": "ナレッジブック側の列(13章§3.5)",
    "watch.source_kind": "ナレッジブック側の列(13章§3.8)",
    "watch.classification": "同上",
    "rule.rule_class": "ナレッジブック側の列(13章§3.2)",
    "rule.workaround": "同上。19章§3のenum欄(4値)と日本語欄(2つの説明)が対応しない書式",
    "target.status": "ナレッジブック側の列(13章§3.6)",
    "rt.status": "ナレッジブック側の列(13章§3.10)",
    "fg.grade": "Phase 1.5(15章§9 SchemaFG)。本体UIは描かない",
}


class CheckError(Exception):
    pass


# ==============================================================================
# 19章§3 の表をパースする
# ==============================================================================
_EMPH = re.compile(r"\*\*|`")
_PARENS = re.compile(r"[（(][^（()）]*[)）]")


def _strip_key(cell: str) -> str:
    """項目キーのセルから括弧書きの注記を落として素のキーにする。"""
    s = _EMPH.sub("", cell).strip()
    prev = None
    while prev != s:                      # 入れ子の括弧書きも落とす
        prev = s
        s = _PARENS.sub("", s).strip()
    return s


def _split_values(cell: str) -> list[str]:
    s = _EMPH.sub("", cell).strip()
    s = re.sub(r"^[（(]日本語enum[)）]\s*", "", s)   # locations.type の前置き
    return [v.strip() for v in s.split("/") if v.strip()]


def parse_registry(spec_path: Path) -> dict[str, tuple[list[str], list[str]]]:
    """19章§3の `項目 | enum | 日本語` を {キー: (機械値列, ラベル列)} で返す。"""
    if not spec_path.exists():
        raise CheckError(f"19章が見つかりません: {spec_path}")

    lines = spec_path.read_text(encoding="utf-8").split("\n")
    start = None
    end = len(lines)
    for i, line in enumerate(lines):
        if line.startswith("## ") and "enumレジストリ" in line:
            start = i
            continue
        if start is not None and line.startswith("## ") and i > start:
            end = i
            break
    if start is None:
        raise CheckError("19章に「## 3. enumレジストリ」節が見つかりません")

    out: dict[str, tuple[list[str], list[str]]] = {}
    for line in lines[start:end]:
        s = line.strip()
        if not s.startswith("|"):
            continue
        cells = [c.strip() for c in s.strip("|").split("|")]
        if len(cells) < 3:
            continue
        key = _strip_key(cells[0])
        if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", key):
            continue                       # 見出し行・区切り行
        values = _split_values(cells[1])
        labels = _split_values(cells[2])
        if labels and labels[0].startswith("同左"):
            labels = list(values)          # 日本語enum(値そのものがラベル)
        if key in out:
            raise CheckError(f"19章§3に同じ項目キーが2行あります: {key}")
        out[key] = (values, labels)
    if not out:
        raise CheckError("19章§3からenum行を1件も読み取れませんでした")
    return out


def expected_pairs(reg: dict[str, tuple[list[str], list[str]]]) -> list[tuple[str, str, str]]:
    """REQUIRED の並び順どおりに (グループ, 機械値, ラベル) を並べる。"""
    pairs: list[tuple[str, str, str]] = []
    problems: list[str] = []

    unknown = sorted(set(reg) - set(REQUIRED) - set(EXCLUDED))
    if unknown:
        problems.append(
            "19章§3の次の行が enum_check.py の分類表(REQUIRED / EXCLUDED)に"
            "ありません。変換表へ載せるか、載せない理由を EXCLUDED へ書いて"
            f"ください: {unknown}")

    missing = sorted(set(REQUIRED) - set(reg))
    if missing:
        problems.append(f"REQUIRED に挙げた項目が19章§3にありません(改名?): {missing}")

    for key, group in REQUIRED.items():
        if key not in reg:
            continue
        values, labels = reg[key]
        if len(values) != len(labels):
            problems.append(
                f"[{key}] 19章§3の機械値({len(values)}個)と日本語({len(labels)}個)の"
                f"数が一致しません: {values} / {labels}")
            continue
        if len(set(values)) != len(values):
            problems.append(f"[{key}] 機械値に重複があります: {values}")
        if len(set(labels)) != len(labels):
            problems.append(f"[{key}] 日本語ラベルに重複があります: {labels}")
        for v, la in zip(values, labels):
            if "," in v or "," in la or "\n" in v or "\n" in la:
                problems.append(f"[{key}] 値・ラベルに ',' か改行が含まれます: {v} / {la}")
            pairs.append((group, v, la))

    if problems:
        raise CheckError("\n".join(problems))
    return pairs


# ==============================================================================
# modUICase.EnumPairsCsv() を静的に評価する
# ------------------------------------------------------------------------------
# prompt_diff.py と同じ考え方(`s = s & "..." & vbLf` 方式の連結だけを評価する)。
# 制御構文・未対応の項があれば評価失敗として不一致に数える。
# ==============================================================================
VBA_CONSTS = {"vblf": "\n", "vbcrlf": "\r\n", "vbcr": "\r", "vbtab": "\t",
              "vbnullstring": ""}
FUNC_HEAD_RE = re.compile(r"^\s*Public\s+Function\s+(\w+)\s*\(", re.IGNORECASE)
FUNC_END_RE = re.compile(r"^\s*End\s+Function\b", re.IGNORECASE)


def _merge_continuations(raw_lines: list[str]) -> list[str]:
    out: list[str] = []
    acc = ""
    for raw in raw_lines:
        line = raw.rstrip("\n\r")
        if line.lstrip().startswith("'"):
            if acc:
                out.append(acc)
                acc = ""
            out.append(line)
            continue
        rstripped = line.rstrip()
        if rstripped.endswith(" _") or rstripped == "_":
            acc = (acc + " " if acc else "") + rstripped[:-1].rstrip()
            continue
        acc = (acc + " " if acc else "") + line
        out.append(acc)
        acc = ""
    if acc:
        out.append(acc)
    return out


def _strip_comment(line: str) -> str:
    in_str = False
    for i, c in enumerate(line):
        if c == '"':
            in_str = not in_str
        elif c == "'" and not in_str:
            return line[:i]
    return line


def _split_amp(expr: str) -> list[str]:
    parts, cur, in_str = [], [], False
    for c in expr:
        if c == '"':
            in_str = not in_str
            cur.append(c)
        elif c == "&" and not in_str:
            parts.append("".join(cur))
            cur = []
        else:
            cur.append(c)
    parts.append("".join(cur))
    return parts


def _assign_index(stmt: str) -> int:
    in_str = False
    for i, c in enumerate(stmt):
        if c == '"':
            in_str = not in_str
        elif c == "=" and not in_str:
            return i
    return -1


def _eval_term(term: str, env: dict[str, str]) -> str:
    t = term.strip()
    if not t:
        raise CheckError("空の項があります")
    if t.startswith('"') and t.endswith('"') and len(t) >= 2:
        return t[1:-1].replace('""', '"')
    low = t.lower()
    if low in VBA_CONSTS:
        return VBA_CONSTS[low]
    if low in env:
        return env[low]
    raise CheckError(
        f"{PAIRS_FUNC} に評価できない項があります: {t[:40]!r}"
        "(文字列リテラル・vbLf等の組込定数・同一関数内の変数の連結だけを評価します)")


def eval_pairs_func(src_path: Path) -> str:
    if not src_path.exists():
        raise CheckError(f"変換表の実装が見つかりません: {src_path}")
    merged = _merge_continuations(src_path.read_text(encoding="utf-8").split("\n"))

    body: list[str] | None = None
    cur: list[str] = []
    name = None
    for line in merged:
        code = _strip_comment(line)
        if name is None:
            m = FUNC_HEAD_RE.match(code)
            if m:
                name = m.group(1)
                cur = []
            continue
        if FUNC_END_RE.match(code):
            if name.lower() == PAIRS_FUNC.lower():
                body = cur
                break
            name = None
            continue
        cur.append(line)

    if body is None:
        raise CheckError(
            f"{src_path.name} に Public Function {PAIRS_FUNC}() が見つかりません"
            "(11章§5の変換表はこの1本が唯一の値源です)")

    env: dict[str, str] = {}
    for raw in body:
        stmt = _strip_comment(raw).strip()
        if not stmt:
            continue
        low = stmt.lower()
        if low.startswith(("dim ", "static ")):
            # `Dim s As String` は空文字で初期化される。1行目の `s = s & "..."`
            # を評価できるようにするため、宣言を環境へ写しておく。
            m = re.match(r"(?:dim|static)\s+([A-Za-z_]\w*)\s+as\s+string\b", low)
            if m:
                env.setdefault(m.group(1), "")
            continue
        if low.startswith(("const ", "exit ", "on error")):
            continue
        if low.startswith(("if ", "elseif ", "else", "end if", "for ", "next",
                           "do ", "loop", "select ", "case ", "with ", "end with",
                           "while ", "wend")):
            raise CheckError(f"{PAIRS_FUNC} に制御構文があります: {stmt[:60]}")
        eq = _assign_index(stmt)
        if eq < 0:
            raise CheckError(f"{PAIRS_FUNC} に代入以外の文があります: {stmt[:60]}")
        lhs = stmt[:eq].strip()
        if not re.fullmatch(r"[A-Za-z_]\w*", lhs):
            raise CheckError(f"{PAIRS_FUNC} に単純変数以外への代入があります: {stmt[:60]}")
        env[lhs.lower()] = "".join(_eval_term(p, env) for p in _split_amp(stmt[eq + 1:]))

    if PAIRS_FUNC.lower() not in env:
        raise CheckError(f"{PAIRS_FUNC} への戻り値の代入が見つかりません")
    return env[PAIRS_FUNC.lower()]


def parse_pairs_csv(text: str) -> list[tuple[str, str, str]]:
    out: list[tuple[str, str, str]] = []
    for lineno, line in enumerate(text.split("\n"), 1):
        s = line.strip()
        if not s:
            continue
        cells = s.split(",")
        if len(cells) != 3:
            raise CheckError(
                f"{PAIRS_FUNC} の {lineno} 行目が `グループ,機械値,日本語` の3列では"
                f"ありません: {s!r}")
        out.append((cells[0].strip(), cells[1].strip(), cells[2].strip()))
    return out


# ==============================================================================
# 照合
# ==============================================================================
def compare(want: list[tuple[str, str, str]],
            got: list[tuple[str, str, str]]) -> list[str]:
    problems: list[str] = []

    want_groups = []
    for g, _, _ in want:
        if g not in want_groups:
            want_groups.append(g)
    got_groups = []
    for g, _, _ in got:
        if g not in got_groups:
            got_groups.append(g)

    extra = [g for g in got_groups if g not in want_groups]
    if extra:
        problems.append(
            f"変換表に19章§3と対応しないグループがあります(REQUIRED外): {extra}")
    lost = [g for g in want_groups if g not in got_groups]
    if lost:
        problems.append(f"変換表に載っていないグループがあります: {lost}")

    for g in want_groups:
        w = [(v, la) for gg, v, la in want if gg == g]
        o = [(v, la) for gg, v, la in got if gg == g]
        if w == o:
            continue
        problems.append(f"[{g}] 19章§3と一致しません")
        problems.append(f"    19章 : {w}")
        problems.append(f"    実装 : {o}")

    if not problems and want != got:
        problems.append("グループの並び順が19章§3(REQUIRED の宣言順)と違います")
    return problems


def to_bas(pairs: list[tuple[str, str, str]]) -> str:
    """.bas の連結文へ整形(--dump-bas。実装を手で書き写す事故を防ぐため)。"""
    out = []
    for g, v, la in pairs:
        out.append(f'    s = s & "{g},{v},{la}" & vbLf')
    return "\n".join(out)


def main() -> int:
    ap = argparse.ArgumentParser(
        description="19章§3 enumレジストリ と modUICase の変換表の一致検査(17章§4-2)")
    ap.add_argument("--spec", default=str(DEFAULT_SPEC))
    ap.add_argument("--src", default=str(DEFAULT_SRC))
    ap.add_argument("--dump", action="store_true", help="期待されるCSV本文を出力")
    ap.add_argument("--dump-bas", action="store_true", help="上を .bas の連結文で出力")
    args = ap.parse_args()

    try:
        reg = parse_registry(Path(args.spec))
        want = expected_pairs(reg)
    except CheckError as e:
        print("[enum_check] 19章§3の読み取りに失敗しました:")
        print(str(e))
        return 1

    if args.dump or args.dump_bas:
        if args.dump_bas:
            print(to_bas(want))
        else:
            for g, v, la in want:
                print(f"{g},{v},{la}")
        return 0

    print("=" * 78)
    print("enum_check: 19章§3 <-> modUICase の変換表(17章§4-2の一致検査)")
    print("=" * 78)
    print(f"  19章の行数     : {len(reg)}"
          f"(変換対象 {len(REQUIRED)} / 対象外 {len(EXCLUDED)})")
    print(f"  期待するペア数 : {len(want)}")

    try:
        csv_text = eval_pairs_func(Path(args.src))
        got = parse_pairs_csv(csv_text)
    except CheckError as e:
        print("  FAIL: 変換表を読み取れませんでした")
        print(f"    {e}")
        return 1

    print(f"  実装のペア数   : {len(got)}  ({Path(args.src).name}.{PAIRS_FUNC})")

    problems = compare(want, got)
    if problems:
        print("-" * 78)
        for p in problems:
            print(f"  {p}")
        print("-" * 78)
        print(f"結果: 不一致 {len([p for p in problems if not p.startswith('    ')])} 件"
              "(exit code 1)")
        return 1

    print("-" * 78)
    print("結果: 一致(exit code 0)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
