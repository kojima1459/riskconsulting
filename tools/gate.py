#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""gate.py - 全検問の一括実行ラッパー(トークン効率版)。

目的(発注者指示 2026-08-30):
  テストが数百～数千本規模になると、検問が緑でも生ログが数千行出て
  レビュー側(司令塔・エージェント)のトークンを浪費する。本ラッパーは
    - 成功したゲート: 1行(ゲート名+要点の数値)だけを出す
    - 失敗したゲート: 要点行+末尾ログ(既定60行)+全文ログのパスを出す
  という非対称出力で「緑は静かに・赤は雄弁に」を機械で強制する。

使い方:
  python3 tools/gate.py               # 全ゲート
  python3 tools/gate.py --only lint,lo-pure
  python3 tools/gate.py --tail 120    # 失敗時の表示行数
  python3 tools/gate.py --list        # ゲート一覧
  python3 tools/gate.py --verbose     # 従来どおり全出力(デバッグ用)

全文ログは常に <tempdir>/rpn_gate_logs/<gate>.log へ保存される(成功時も)。
"""
import argparse
import os
import re
import subprocess
import sys
import tempfile
import time

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOGDIR = os.path.join(tempfile.gettempdir(), "rpn_gate_logs")

# (名前, コマンド, 成功時に抜き出す要点パターン(最後にマッチした行を表示))
GATES = [
    ("lint",         [sys.executable, "tools/vba_lint.py"],
     r"ERROR: \d+ 件"),
    ("lo-compile",   [sys.executable, "tools/run_lo_tests.py", "--mode", "compile"],
     r"対象モジュール数: \d+|結果: OK"),
    ("lo-pure",      [sys.executable, "tools/run_lo_tests.py", "--mode", "pure"],
     r"PASS \d+ / FAIL \d+ / SKIP \d+"),
    ("build-dev",    [sys.executable, "build/build_rpn.py", "--dev"],
     r"自己検証 OK|Done"),
    ("build-prod",   [sys.executable, "build/build_rpn.py", "--prod"],
     r"自己検証 OK|Done"),
    ("build-kb",     [sys.executable, "build/build_rpn.py", "--kb"],
     r"自己検証 OK|Done"),
    ("sheet",        [sys.executable, "tools/sheet_check.py"],
     r"OK: 全\d+項目一致"),
    ("sheet-kb",     [sys.executable, "tools/sheet_check.py", "--kb"],
     r"OK: 全\d+項目一致"),
    ("ship",         [sys.executable, "tools/ship_check.py"],
     r"結果: .*PASS"),
    # 裁定書27 W9-A: 配布 vbaProject.bin を「別実装で読み戻す」2本。
    #   bin-roundtrip = olevba で解凍して src/ とバイト比較(中身の検問)
    #   lo-xlsm       = LibreOffice に配布xlsmを開かせる(入れ物の検問)
    ("bin-roundtrip", [sys.executable, "tools/bin_roundtrip.py"],
     r"結果: OK .*5条件"),
    ("lo-xlsm",      [sys.executable, "tools/lo_xlsm.py"],
     r"結果: OK 3条件"),
    ("prompt-diff",  [sys.executable, "tools/prompt_diff.py", "--strict"],
     r"一致: \d+件"),
    ("validate",     [sys.executable, "tools/validate_check.py"],
     r"結果: OK|ケース: 計\d+件"),
    ("enum",         [sys.executable, "tools/enum_check.py"],
     r"ペア|OK|一致"),
    ("caption",      [sys.executable, "tools/caption_check.py"],
     r"OK: 全\d+本"),
    ("t48",          [sys.executable, "tools/t48_check.py"],
     r"OK: 4条件"),
    ("render",       [sys.executable, "tools/render_report.py"],
     r"OK: .*確認しました|生成:"),
    ("render-f",     [sys.executable, "tools/render_report.py", "--faithful"],
     r"OK: .*確認しました|生成:"),
]


def pick_summary(text, pattern):
    hits = re.findall(r"^.*(?:%s).*$" % pattern, text, re.M)
    return hits[-1].strip() if hits else "(要点行なし)"


def main():
    ap = argparse.ArgumentParser(description="全検問の一括実行(緑は静かに・赤は雄弁に)")
    ap.add_argument("--only", help="カンマ区切りのゲート名で絞る")
    ap.add_argument("--tail", type=int, default=60, help="失敗時に表示する末尾行数(既定60)")
    ap.add_argument("--list", action="store_true", help="ゲート一覧を表示して終了")
    ap.add_argument("--verbose", action="store_true", help="成功時も全出力(デバッグ用)")
    args = ap.parse_args()

    if args.list:
        for name, cmd, _ in GATES:
            print("%-12s %s" % (name, " ".join(cmd[1:])))
        return 0

    selected = GATES
    if args.only:
        wanted = [w.strip() for w in args.only.split(",") if w.strip()]
        unknown = [w for w in wanted if w not in {n for n, _, _ in GATES}]
        if unknown:
            print("未知のゲート名: %s (--list で一覧)" % ", ".join(unknown))
            return 2
        selected = [g for g in GATES if g[0] in wanted]

    os.makedirs(LOGDIR, exist_ok=True)
    reds = []
    t0 = time.time()
    for name, cmd, pattern in selected:
        started = time.time()
        proc = subprocess.run(cmd, cwd=REPO, capture_output=True, text=True,
                              encoding="utf-8", errors="replace")
        out = (proc.stdout or "") + (("\n" + proc.stderr) if proc.stderr else "")
        logpath = os.path.join(LOGDIR, name + ".log")
        with open(logpath, "w", encoding="utf-8") as fp:
            fp.write(out)
        secs = time.time() - started
        if proc.returncode == 0 and not args.verbose:
            print("OK  %-12s %5.1fs  %s" % (name, secs, pick_summary(out, pattern)))
        elif proc.returncode == 0:
            print("OK  %-12s %5.1fs" % (name, secs))
            print(out)
        else:
            reds.append(name)
            print("NG  %-12s %5.1fs  exit=%d  全文: %s" % (name, secs, proc.returncode, logpath))
            lines = out.rstrip().split("\n")
            shown = lines[-args.tail:]
            if len(lines) > args.tail:
                print("  ...(先頭%d行は全文ログ参照)..." % (len(lines) - args.tail))
            for ln in shown:
                print("  " + ln)
    total = time.time() - t0
    if reds:
        print("結果: NG %d/%d ゲート赤 [%s] (%.0fs)" % (len(reds), len(selected), ", ".join(reds), total))
        return 1
    print("結果: 全%dゲート緑 (%.0fs)" % (len(selected), total))
    return 0


if __name__ == "__main__":
    sys.exit(main())
