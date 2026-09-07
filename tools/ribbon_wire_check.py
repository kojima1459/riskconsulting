#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ribbon_wire_check.py - リボンの抽出切断を招く `"},` の検問(裁定書33 C-2・16章 E-63)

================================================================================
なぜ要るか(相手側の確定事実。14章§2 の 7.・16章 E-63):
    社内AIアドイン「リボンちゃん」は 200応答の本文を
      始点 `content":"`(無ければ `text":"`)から
      終点 `","` / `"},` / `"`+LF の **最も手前** まで
    で切り出す。当方が返させるJSONの中に `"},` があると、HTTPボディ上は
    `\\"},\\"` になって **`"},` が部分一致し、そこで本文が切られる**。
    `","` は本文中では `\\",\\"`、LF は `\\n` になるので当たらない。
    **危ないのは `"},` だけ**である。

    したがって「1行に詰めたJSON」は実機で途中から切られる。整形(閉じ括弧の
    直前で改行)すれば `\\"\\n}` になり当たらない。本ツールは、その約束が
    ソース側で守られているかを毎回機械で見る。

走査対象(裁定書33 C-2):
    (a) src/test/modMockLlm*.bas の文字列リテラル             -> ERROR
    (b) docs/spec/15_プロンプトとJSONスキーマ.md のコードフェンス
        (`{` で始まる行を含むもの=JSONの例)                   -> 既定は WARN
    (c) src/app/modSchemas.bas の文字列リテラル                -> 既定は WARN

    (b)(c) を既定で WARN にしている理由(**司令塔の裁定待ちの保留**。裁定書33 に
    無い判断なので勝手に本文を書き換えないための扱い):
      ・(a) は **モデルが返す本文**の模擬であり、切断が実際に起きる側である。
      ・(b)(c) は **こちらがモデルへ送る側**のテキスト(スキーマの型注記
        `{"type": "string"},` など)で、リボンの終点規則は応答にしか掛からない。
        送信側で `"},` が出ても切断そのものは起きない。
      ・一方で「送る側の見本が1行詰め」だと、モデルが1行詰めを真似る余地は残る。
        そこを詰めるには 15章のスキーマ節と modSchemas を全面的に整形し直す
        (実測 15章 71件・modSchemas 65件)ことになり、本波の外である。
    `--strict-docs` を付けると (b)(c) も ERROR に昇格する(裁定が出たら
    tools/gate.py の引数へ足すだけで発効する)。

使い方:
    python3 tools/ribbon_wire_check.py
    python3 tools/ribbon_wire_check.py --strict-docs
    exit code: 0 = ERROR 0件 / 1 = ERROR 1件以上 / 2 = 自己テスト失敗

自己テスト(骨抜き防止):
    毎回、末尾で負例(`{"a":"x"},{"b":"y"}` を含むダミーの .bas 本文)と
    正例(改行済みの同じJSON)を走らせ、**負例で検出できなければ exit 2** で
    止める(検出器が壊れたまま緑になるのを防ぐ)。
================================================================================
"""
from __future__ import annotations

import argparse
import glob
import os
import re
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 切断を招く3文字。ここだけが唯一の値源。
CUT_MARK = '"},'

# VBAの文字列リテラル("" は1つの " )。
LITERAL = re.compile(r'"((?:[^"]|"")*)"')
# 組込の改行/タブ定数(リテラルとリテラルの間にあれば実体を挟む)。
CONST_TOKEN = {"vbLf": "\n", "vbCrLf": "\r\n", "vbCr": "\r", "vbTab": "\t",
               "vbNullString": ""}
# 連結だけでできている繋ぎ(空白・& ・行継続)。ここだけなら literal は隣接する。
GLUE_ONLY = re.compile(r"^[\s&_]*$")

# 隣接していないことを表す番人(この文字はソースに現れない)。
SEP = "\x00"


def decode_bas(path: str) -> tuple[str, list[tuple[int, int]]]:
    """.bas を「文字列リテラルの中身をつないだ1本の文字列」へ落とす。

    戻り値 = (本文, [(本文中の開始オフセット, 行番号), ...])。
    ・`s = s & ...` のように**連続する連結文**は前の本文へ continue して足す
      (行またぎで `"}` と `,` が隣り合う形を見落とさないため。実際に
      modMockLlm.BuildS1NewJson で1件見つかった)。
    ・リテラル以外の式(関数呼び出し・変数)が挟まったところには番人を置き、
      本当は隣接していない文字どうしが繋がって誤検出になるのを防ぐ。
    """
    with open(path, encoding="utf-8") as fh:
        raw_lines = fh.read().split("\n")

    parts: list[str] = []
    index: list[tuple[int, int]] = []
    size = 0
    for lineno, raw in enumerate(raw_lines, 1):
        line = raw.strip()
        if line.startswith("'") or line.startswith("Attribute "):
            parts.append(SEP)
            size += len(SEP)
            continue
        pieces, joined = decode_line(line)
        if not pieces:
            parts.append(SEP)
            size += len(SEP)
            continue
        index.append((size, lineno))
        parts.append(joined)
        size += len(joined)
    return "".join(parts), index


def decode_line(line: str) -> tuple[int, str]:
    """1行を復元する。戻り値 = (リテラル数, 復元文字列)。"""
    out: list[str] = []
    count = 0
    pos = 0
    prev_end = None
    for m in LITERAL.finditer(line):
        gap = line[prev_end:m.start()] if prev_end is not None else line[pos:m.start()]
        if prev_end is not None:
            out.append(gap_text(gap))
        else:
            # 代入の左辺など、最初のリテラルより前は「行の頭」なので番人を置く。
            out.append(SEP if not is_continuation(gap) else "")
        out.append(m.group(1).replace('""', '"'))
        count += 1
        prev_end = m.end()
    if prev_end is not None:
        out.append(gap_text(line[prev_end:]))
    return count, "".join(out)


def is_continuation(gap: str) -> bool:
    """`s = s & ` の形か(=直前の行の本文へ続いている)。"""
    return re.match(r"^\s*(\w+)\s*=\s*\1\s*&\s*$", gap) is not None


def gap_text(gap: str) -> str:
    """リテラルとリテラルの間の式を、本文としてどう扱うか。"""
    if GLUE_ONLY.match(gap):
        return ""
    stripped = gap.replace("&", " ").replace("_", " ").strip()
    if stripped in CONST_TOKEN:
        return CONST_TOKEN[stripped]
    return SEP


def find_hits(text: str, index: list[tuple[int, int]]) -> list[int]:
    """本文中の `"},` の位置を行番号へ写す。"""
    hits = []
    start = 0
    while True:
        k = text.find(CUT_MARK, start)
        if k < 0:
            break
        hits.append(lineno_of(index, k))
        start = k + 1
    return hits


def lineno_of(index: list[tuple[int, int]], offset: int) -> int:
    found = 0
    for off, lineno in index:
        if off <= offset:
            found = lineno
        else:
            break
    return found


def scan_bas(path: str) -> list[int]:
    text, index = decode_bas(path)
    return find_hits(text, index)


def scan_md_fences(path: str) -> list[int]:
    """コードフェンスのうち `{` で始まる行を含むもの(=JSONの例)だけを見る。"""
    with open(path, encoding="utf-8") as fh:
        lines = fh.read().split("\n")
    hits: list[int] = []
    inside = False
    buf: list[tuple[int, str]] = []
    for lineno, line in enumerate(lines, 1):
        if line.startswith("```"):
            if inside:
                if any(t.lstrip().startswith("{") for _, t in buf):
                    hits.extend(n for n, t in buf if CUT_MARK in t)
                inside = False
                buf = []
            else:
                inside = True
                buf = []
            continue
        if inside:
            buf.append((lineno, line))
    return hits


# ==============================================================================
# 自己テスト(負例で検出できなければ exit 2)
# ==============================================================================
NEG_SRC = '''Attribute VB_Name = "selftestNeg"
Public Function Body() As String
    Dim s As String
    s = ""
    s = s & "{""a"":""x""},{""b"":""y""}"
    Body = s
End Function
'''
POS_SRC = '''Attribute VB_Name = "selftestPos"
Public Function Body() As String
    Dim s As String
    s = ""
    s = s & "{""a"":""x""" & vbLf
    s = s & "},{""b"":""y""" & vbLf
    s = s & "}"
    Body = s
End Function
'''
# 行またぎ(前の行の末尾 `"}` と次の行の頭 `,` が隣り合う)の負例。
NEG_SPLIT_SRC = '''Attribute VB_Name = "selftestNegSplit"
Public Function Body() As String
    Dim s As String
    s = ""
    s = s & "{""a"":{""b"":""x""}"
    s = s & ",""c"":1}"
    Body = s
End Function
'''


def self_test(tmpdir: str) -> bool:
    ok = True
    cases = [("neg", NEG_SRC, True), ("pos", POS_SRC, False),
             ("neg_split", NEG_SPLIT_SRC, True)]
    for name, src, want_hit in cases:
        path = os.path.join(tmpdir, "selftest_%s.bas" % name)
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(src)
        hits = scan_bas(path)
        got = len(hits) > 0
        mark = "OK" if got == want_hit else "NG"
        if got != want_hit:
            ok = False
        print("  自己テスト %-9s 期待=%s 実際=%d件 ... %s"
              % (name, "検出" if want_hit else "検出なし", len(hits), mark))
        os.remove(path)
    return ok


def main() -> int:
    ap = argparse.ArgumentParser(description="`\"},` の検問(裁定書33 C-2)")
    ap.add_argument("--strict-docs", action="store_true",
                    help="(b)15章フェンスと(c)modSchemas も ERROR にする")
    args = ap.parse_args()

    print("== ribbon_wire_check: リボンの抽出切断を招く `\"},` の走査 ==")

    errors = 0
    warns = 0

    # (a) mock 応答(モデルが返す本文の模擬)= ERROR
    mocks = sorted(glob.glob(os.path.join(REPO, "src", "test", "modMockLlm*.bas")))
    if not mocks:
        print("ERROR: 走査対象 src/test/modMockLlm*.bas が1本もありません(検査不能)")
        return 1
    for path in mocks:
        hits = scan_bas(path)
        rel = os.path.relpath(path, REPO)
        for lineno in hits:
            print('ERROR %s:%d 文字列に `"},` があります(実機で本文がここで切られる)'
                  % (rel, lineno))
        errors += len(hits)
        print("  (a) %-32s %d件" % (rel, len(hits)))

    # (b) 15章のコードフェンス(JSONの例を含むもの)
    md = os.path.join(REPO, "docs", "spec", "15_プロンプトとJSONスキーマ.md")
    md_hits = scan_md_fences(md) if os.path.exists(md) else []
    level_b = "ERROR" if args.strict_docs else "WARN "
    for lineno in md_hits[:5]:
        print('%s docs/spec/15_プロンプトとJSONスキーマ.md:%d フェンス内に `"},`'
              % (level_b.strip(), lineno))
    if len(md_hits) > 5:
        print("  ... 他 %d件(先頭5件のみ表示)" % (len(md_hits) - 5))
    print("  (b) 15章のJSONフェンス%19s %d件" % ("", len(md_hits)))

    # (c) modSchemas(送信側のスキーマ本文)
    schemas = sorted(glob.glob(os.path.join(REPO, "src", "app", "modSchemas*.bas")))
    sc_hits = 0
    for path in schemas:
        hits = scan_bas(path)
        sc_hits += len(hits)
        print("  (c) %-32s %d件" % (os.path.relpath(path, REPO), len(hits)))

    if args.strict_docs:
        errors += len(md_hits) + sc_hits
    else:
        warns += len(md_hits) + sc_hits

    tmpdir = os.path.join(REPO, "tools")
    if not self_test(tmpdir):
        print("結果: 自己テスト失敗(検出器が壊れています)")
        return 2

    if errors:
        print("ERROR: %d 件" % errors)
        print("結果: NG")
        return 1
    print("結果: OK (ERROR 0件 / WARN %d件=(b)(c)は裁定待ちの保留)" % warns)
    return 0


if __name__ == "__main__":
    sys.exit(main())
