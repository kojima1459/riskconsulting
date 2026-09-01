#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
t48_check.py - ブック内テスト実行(T-48)の合否判定4条件の検査(裁定書16 F1)

================================================================================
位置づけ:
    17章 T-48 は「ブック内テスト実行(modTestsRunnerUi.RunAllTestsFromBook)は
    ps1と**同一の4条件**で合否を出す」と定めている。
        (1) FAIL 0件
        (2) SKIP 0件
        (3) 純層の実行本数 = 期待本数
        (4) 層(b) 1本以上
    この4条件は**利用者が社内PCで回せる唯一の検問**の判定式そのものであり、
    ここから1条件でも落ちれば「落ちているのに全PASSと出る」検問になる。
    ところが従来、この判定式を守る機械検査は1本も無かった(=判定を書き換えても
    どのゲートも赤くならない)。本ツールはその穴を塞ぐ。

検査するもの:
    [1] 対象モジュール src/test/modTestsRunnerUi.bas が存在すること。
    [2] その中に**合否判定の関数**(4条件を And で連ねた If を持つ手続き)が
        ちょうど1つ在ること。
    [3] その If の条件式が **And で連なる4項**であり、各項が変数の代入元まで
        遡って次の4役へ**過不足なく**割り当たること。
          fail-zero  : FailCount() 由来の変数 = 0
          skip-zero  : SkipCount() 由来の変数 = 0
          pure-equals: ExecutedCount() 由来(引き算なし=純層)の変数
                       = ExpectedCount() 由来(期待本数)の変数
          excel-ge-1 : ExecutedCount() の**差し引き**由来(層(b))の変数 >= 1
        1条件でも消えれば役が欠け、条件を足しても割り当たらない項が出て赤になる。
    [4] 判定の分岐が fail-open していないこと(真枝=全PASS表示 / 偽枝=NG表示が
        両方在ること)。Else を消して常に全PASSにする改変を止める。
    [5] 17章 T-48 の**規定文**に4条件がそのまま書かれていること
        (仕様側の規定文を薄めてから実装を薄める、という順序の骨抜きを止める)。

骨抜き防止(fail-closed):
    対象モジュール不在・判定関数不在・条件式の抽出0件・規定文の行不在は、
    いずれも**合格側へ倒さず** NG(exit 1)にする。「検査対象が見つからないので
    何も検査せず緑」を作らない。

なぜ .bas を静的評価するのか:
    enum_check.py / prompt_diff.py と同じ理由である。Excel も LibreOffice も
    要らない静的読取なので、検査が「実行できる環境がある」ことに依存しない。
    また、条件を関数呼び出しの奥へ隠して式の形だけ保つ改変も、変数の代入元を
    たどる本ツールでは「解決できない項」として差分に出る。

使い方:
    python3 tools/t48_check.py          # 検査(exit code: 0=OK / 1=NG)
    python3 tools/t48_check.py --dump   # 抽出した判定式と役の割当を表示
================================================================================
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

TOOLS_DIR = Path(__file__).resolve().parent
REPO = TOOLS_DIR.parent
BAS_PATH = REPO / "src" / "test" / "modTestsRunnerUi.bas"
CH17_PATH = REPO / "docs" / "spec" / "17_実装計画とタスク分解.md"

# 17章 T-48 の規定文に必ず含まれていなければならない語(裁定書16 F1)。
CH17_REQUIRED_PHRASES = [
    "ps1と同一の4条件",
    "FAIL 0件",
    "SKIP 0件",
    "純層の実行本数=期待本数",
    "層(b) 1本以上",
]

# 真枝/偽枝に必ず現れる表示語(fail-open 改変の検出)。
PASS_LITERAL = "全PASS"
NG_LITERAL = "NG"

ROLES = ("fail-zero", "skip-zero", "pure-equals", "excel-ge-1")


class CheckFailed(Exception):
    pass


# ---------------------------------------------------------------------------
# .bas の素朴な前処理(行継続の連結・コメントと文字列の除去)
# ---------------------------------------------------------------------------
def strip_comment(line: str) -> str:
    """行末コメントを落とす(VBAの文字列は二重引用符のみ。'' は無い)。"""
    out = []
    in_str = False
    for ch in line:
        if ch == '"':
            in_str = not in_str
            out.append(ch)
            continue
        if ch == "'" and not in_str:
            break
        out.append(ch)
    return "".join(out)


def logical_lines(text: str):
    """行継続( _ )を連結し、コメントを落とした論理行を (行番号, 本文) で返す。"""
    raw = text.replace("\r\n", "\n").split("\n")
    result = []
    buf = ""
    start = 0
    for i, line in enumerate(raw, start=1):
        body = strip_comment(line).rstrip()
        if not buf:
            start = i
        if body.endswith("_") and (len(body) == 1 or body[-2].isspace()):
            buf += body[:-1]
            continue
        buf += body
        if buf.strip():
            result.append((start, re.sub(r"\s+", " ", buf.strip())))
        buf = ""
    if buf.strip():
        result.append((start, re.sub(r"\s+", " ", buf.strip())))
    return result


PROC_RE = re.compile(
    r"^(?:Public |Private |Friend )?(?:Static )?(Sub|Function)\s+([A-Za-z_]\w*)")
END_PROC_RE = re.compile(r"^End (Sub|Function)$")
ASSIGN_RE = re.compile(r"^([A-Za-z_]\w*)\s*=\s*(.+)$")


def procedures(lines):
    """[(手続き名, 開始行, [(行番号, 本文), ...])] を返す。"""
    procs = []
    cur = None
    for no, body in lines:
        m = PROC_RE.match(body)
        if m and cur is None:
            cur = [m.group(2), no, []]
            continue
        if cur is not None:
            if END_PROC_RE.match(body):
                procs.append(tuple(cur))
                cur = None
            else:
                cur[2].append((no, body))
    return procs


def assignment_map(lines):
    """変数名 -> 代入右辺の一覧(同名に複数代入があれば複数入る)。"""
    amap = {}
    for _, body in lines:
        if body.startswith(("If ", "ElseIf ", "For ", "Do ", "While ", "Case ")):
            continue
        m = ASSIGN_RE.match(body)
        if not m:
            continue
        amap.setdefault(m.group(1), []).append(m.group(2).strip())
    return amap


def resolve(expr: str, amap, depth: int = 0) -> str:
    """式中の変数を「代入が1つだけの変数」に限って代入元へ展開する。

    代入が2つ以上ある変数は展開しない(どちらの値か機械には決まらないため)。
    展開できない=役に割り当たらない=赤、という方向へ倒す。"""
    if depth > 6:
        return expr
    def repl(m):
        name = m.group(0)
        rhs = amap.get(name)
        if not rhs or len(rhs) != 1:
            return name
        if re.fullmatch(r"[A-Za-z_]\w*", rhs[0]) and rhs[0] == name:
            return name
        return "( " + resolve(rhs[0], amap, depth + 1) + " )"
    return re.sub(r"(?<![\w.])[A-Za-z_]\w*(?!\s*\()", repl, expr)


# ---------------------------------------------------------------------------
# 判定式の抽出と役割の割当
# ---------------------------------------------------------------------------
def split_and(cond: str):
    """トップレベルの And で条件式を割る(括弧の中の And は割らない)。"""
    parts, buf, depth = [], "", 0
    tokens = re.split(r"(\(|\)|\bAnd\b|\bOr\b)", cond)
    for tok in tokens:
        if tok == "(":
            depth += 1
            buf += tok
        elif tok == ")":
            depth -= 1
            buf += tok
        elif tok == "And" and depth == 0:
            parts.append(buf.strip())
            buf = ""
        elif tok == "Or" and depth == 0:
            # Or が入ると「どれか1つ満たせば合格」に化ける。役に割り当てず赤にする。
            parts.append(buf.strip())
            parts.append("__OR__")
            buf = ""
        else:
            buf += tok
    parts.append(buf.strip())
    return [p for p in parts if p]


def classify(conj: str, amap):
    """条件1項を4役のいずれかへ割り当てる。割り当たらなければ None。"""
    m = re.match(r"^(.+?)\s*(>=|<=|<>|=|>|<)\s*(.+?)$", conj)
    if not m:
        return None
    lhs, op, rhs = m.group(1).strip(), m.group(2), m.group(3).strip()
    rl, rr = resolve(lhs, amap), resolve(rhs, amap)

    def has(src, name):
        return name in src

    def is_pure(src):
        return has(src, "ExecutedCount") and "-" not in src

    def is_excel(src):
        return has(src, "ExecutedCount") and "-" in src

    if op == "=" and rhs == "0" and has(rl, "FailCount"):
        return "fail-zero"
    if op == "=" and rhs == "0" and has(rl, "SkipCount"):
        return "skip-zero"
    if op == "=" and ((is_pure(rl) and has(rr, "ExpectedCount"))
                      or (is_pure(rr) and has(rl, "ExpectedCount"))):
        return "pure-equals"
    if op == ">=" and rhs == "1" and is_excel(rl):
        return "excel-ge-1"
    return None


def find_decision(procs):
    """4条件を And で連ねた If を持つ手続きを探す。ちょうど1つでなければ例外。"""
    hits = []
    for name, start, body in procs:
        for no, line in body:
            m = re.match(r"^If (.+) Then$", line)
            if not m:
                continue
            cond = m.group(1)
            if " And " not in cond:
                continue
            hits.append((name, no, cond, body))
    if not hits:
        raise CheckFailed(
            "合否判定の関数が見当たらない(And で連なる If が1つも無い)。"
            "判定を関数呼び出しの奥へ隠すと本検査は通らない=fail-closed。")
    if len(hits) > 1:
        raise CheckFailed(
            "And で連なる If が %d 箇所ある。合否判定がどれか機械に決まらない: %s"
            % (len(hits), ", ".join("%s:%d" % (h[0], h[1]) for h in hits)))
    return hits[0]


def check_branches(body, cond_line_no):
    """真枝=全PASS表示 / 偽枝=NG表示 が両方在ることを確かめる(fail-open検出)。"""
    tail = [ln for no, ln in body if no >= cond_line_no]
    try:
        end_at = next(i for i, ln in enumerate(tail) if ln == "End If")
    except StopIteration:
        raise CheckFailed("合否判定の If に対応する End If が見つからない。")
    block = tail[1:end_at]
    if not any(ln == "Else" for ln in block):
        raise CheckFailed(
            "合否判定に Else が無い。落ちても全PASSと表示する fail-open になる。")
    idx = [i for i, ln in enumerate(block) if ln == "Else"][0]
    true_arm, false_arm = block[:idx], block[idx + 1:]
    if not any(PASS_LITERAL in ln for ln in true_arm):
        raise CheckFailed("合否判定の真枝に「%s」の表示が無い。" % PASS_LITERAL)
    if not any(NG_LITERAL in ln for ln in false_arm):
        raise CheckFailed("合否判定の偽枝に「%s」の表示が無い。" % NG_LITERAL)


def check_ch17():
    """17章 T-48 の規定文に4条件がそのまま書かれていること。"""
    if not CH17_PATH.exists():
        raise CheckFailed("17章が見つからない: %s" % CH17_PATH)
    rows = [ln for ln in CH17_PATH.read_text(encoding="utf-8").split("\n")
            if ln.startswith("| T-48 ")]
    if len(rows) != 1:
        raise CheckFailed(
            "17章の T-48 行が %d 本(1本であること)。規定文の所在が定まらない。"
            % len(rows))
    row = rows[0]
    missing = [p for p in CH17_REQUIRED_PHRASES if p not in row]
    if missing:
        raise CheckFailed(
            "17章 T-48 の規定文から4条件の語が消えている: %s" % " / ".join(missing))
    return row


def run(dump=False):
    if not BAS_PATH.exists():
        raise CheckFailed("対象モジュールが見つからない: %s" % BAS_PATH)

    lines = logical_lines(BAS_PATH.read_text(encoding="utf-8"))
    procs = procedures(lines)
    if not procs:
        raise CheckFailed("対象モジュールから手続きを1つも抽出できなかった。")

    amap = assignment_map(lines)
    name, no, cond, body = find_decision(procs)
    conjuncts = split_and(cond)
    if not conjuncts:
        raise CheckFailed("合否判定の条件式を1項も抽出できなかった。")

    assigned = {}
    unmatched = []
    for c in conjuncts:
        role = classify(c, amap) if c != "__OR__" else None
        if role is None:
            unmatched.append(c)
        elif role in assigned:
            raise CheckFailed(
                "条件「%s」が既出の役 %s と重複している(4条件が揃わない)。"
                % (c, role))
        else:
            assigned[role] = c

    if dump:
        print("判定関数: %s (%s:%d)" % (name, BAS_PATH.name, no))
        print("条件式  : %s" % cond)
        for r in ROLES:
            print("  %-12s <- %s" % (r, assigned.get(r, "(欠落)")))
        for c in unmatched:
            print("  %-12s <- %s" % ("(未分類)", c))

    lack = [r for r in ROLES if r not in assigned]
    if lack:
        raise CheckFailed(
            "合否判定から条件が欠けている: %s (条件式: %s)" % (" / ".join(lack), cond))
    if unmatched:
        raise CheckFailed(
            "合否判定に4条件以外の項がある(判定の意味が変わる): %s"
            % " / ".join(unmatched))

    check_branches(body, no)
    check_ch17()
    return name, cond


def main():
    ap = argparse.ArgumentParser(
        description="T-48(ブック内テスト実行)の合否判定4条件の検査")
    ap.add_argument("--dump", action="store_true",
                    help="抽出した判定式と役の割当を表示する")
    args = ap.parse_args()
    try:
        name, cond = run(dump=args.dump)
    except CheckFailed as exc:
        print("NG: %s" % exc)
        print("結果: NG (T-48の合否判定4条件・17章規定文)")
        return 1
    print("判定関数 %s / 条件式: %s" % (name, cond))
    print("OK: 4条件(FAIL0 / SKIP0 / 純層=期待 / 層(b)>=1)と17章T-48の規定文が一致")
    return 0


if __name__ == "__main__":
    sys.exit(main())
