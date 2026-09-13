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
    # 裁定書30 裁定1(e): 純層は配布集合(ship:true)と dev専用で別々に照合する。
    #   lo-pure     = 配布集合だけを走らせ tests_expected の prod と一致
    #   lo-pure-dev = dev専用モジュールだけを走らせ dev_only と一致
    ("lo-pure-dev",  [sys.executable, "tools/run_lo_tests.py", "--mode", "pure",
                      "--pure-set", "dev-only"],
     r"PASS \d+ / FAIL \d+ / SKIP \d+"),
    ("build-dev",    [sys.executable, "build/build_rpn.py", "--dev"],
     r"自己検証 OK|Done"),
    ("build-prod",   [sys.executable, "build/build_rpn.py", "--prod"],
     r"自己検証 OK|Done"),
    ("build-kb",     [sys.executable, "build/build_rpn.py", "--kb"],
     r"自己検証 OK|Done"),
    ("build-final",  [sys.executable, "build/build_rpn.py", "--final"],
     r"自己検証 OK|Done"),  # 一段化: frm/frx と参照設定を bin へ(裁定書38 Z-42)。build-prod の後
    ("sheet",        [sys.executable, "tools/sheet_check.py"],
     r"OK: 全\d+項目一致"),
    ("sheet-kb",     [sys.executable, "tools/sheet_check.py", "--kb"],
     r"OK: 全\d+項目一致"),
    # 裁定書34 §1.1/§0.3(W12-A): --final を付けると、第2段のスクリプト
    #   (build/win/import_navi_modules.ps1)の存在と必須文字列も見る。第2段の
    #   産物(dist/final/)は CI には無いので SKIP になり、赤にはならない。
    ("ship",         [sys.executable, "tools/ship_check.py", "--final"],
     r"結果: .*PASS"),
    # 裁定書27 W9-A: 配布 vbaProject.bin を「別実装で読み戻す」2本。
    #   bin-roundtrip = olevba で解凍して src/ とバイト比較(中身の検問)
    #   lo-xlsm       = LibreOffice に配布xlsmを開かせる(入れ物の検問)
    ("bin-roundtrip", [sys.executable, "tools/bin_roundtrip.py"],
     r"結果: OK .*6条件"),
    ("lo-xlsm",      [sys.executable, "tools/lo_xlsm.py"],
     r"結果: OK 3条件"),
    ("bin-rt-final", [sys.executable, "tools/bin_roundtrip.py", "--final"],
     r"結果: OK"),  # final の参照設定と designer 4本(裁定書38 Z-42)
    ("lo-xlsm-final", [sys.executable, "tools/lo_xlsm.py", "--book",
                       "dist/final/リスク提案ナビ.xlsm"],
     r"結果: OK 3条件"),  # LO が frmNaviHtml 込みで読めてコンパイルできる
    ("prompt-diff",  [sys.executable, "tools/prompt_diff.py", "--strict"],
     r"一致: \d+件"),
    ("dossier",      [sys.executable, "tools/dossier_check.py"],
     r"\[dossier_check\] OK: 全\d+本"),  # docs/08 が15章ルール5bの4点を含む(裁定書37 B-07)
    ("config",       [sys.executable, "tools/config_check.py"],
     r"結果: OK"),  # config の5点一致(sheets_main/RegisterDefault/読取/13章/19章。裁定書38 班D・班H)
    ("orphan",       [sys.executable, "tools/orphan_check.py"],
     r"結果: OK \(孤児Public 0件"),  # 呼ばれない Public を機械で数える(裁定書38 班D)
    ("doc-gate",     [sys.executable, "tools/doc_gate.py"],
     r"結果: OK 4条件"),  # 文書の既定値・ボタン名・エラーコード(裁定書38 班D)
    ("action",       [sys.executable, "tools/action_check.py"],
     r"結果: OK 2条件"),  # ui action ⇔ Dispatch ⇔ 11章(裁定書38 班D)
    ("validate",     [sys.executable, "tools/validate_check.py"],
     r"結果: OK|ケース: 計\d+件"),
    ("enum",         [sys.executable, "tools/enum_check.py"],
     r"ペア|OK|一致"),
    ("caption",      [sys.executable, "tools/caption_check.py"],
     r"OK: 全\d+本"),
    ("t48",          [sys.executable, "tools/t48_check.py"],
     r"OK: 4条件"),
    # 裁定書40 Q-M2 / Q-m2 で新設したが**どのゲートにも載っていなかった**ので、
    #   裁定書41 §2 で登録した(載せるまで、SEC-09「原文未照合」の判定を反転
    #   させる変異が全ゲートを素通りする状態が続いていた)。検査①は LibreOffice
    #   で実物のHTMLを組み、node の最小DOMスタブでページ内JSを実際に走らせる。
    ("notice",       [sys.executable, "tools/notice_check.py"],
     r"\[notice_check\] OK"),
    # 裁定書41 §2: 要点行から `|生成:` を落とした。旧パターンは OK 行が無くても
    #   「生成: …」を要点として表示できたので、DOM検査を飛ばした回でも
    #   ゲート一覧が成功したように読めた(render_report.py 側の fail-open も同時に
    #   塞いだ。node が無ければ exit 2 で赤)。
    ("render",       [sys.executable, "tools/render_report.py"],
     r"OK: .*確認しました"),
    ("render-f",     [sys.executable, "tools/render_report.py", "--faithful"],
     r"OK: .*確認しました"),
    # 裁定書33 C-2(W11-c): リボンの抽出切断を招く `"},` の走査。mock 応答
    #   (モデルが返す本文の模擬)に1件でもあれば赤。15章のJSONフェンスと
    #   modSchemas は既定 WARN(--strict-docs で昇格。裁定待ちの保留)。
    ("render-p",     [sys.executable, "tools/render_proposal.py"],
     r"\[render_proposal\] OK"),  # 提案書 Wide 22枚(20章。裁定書38 班C)
    ("render-pf",    [sys.executable, "tools/render_proposal.py", "--faithful"],
     r"\[render_proposal\] OK"),
    ("wire",         [sys.executable, "tools/ribbon_wire_check.py"],
     r"結果: OK"),
    # 裁定書34 §1.4(W12-A): HTML画面(ui/ と src/ui/navi/)の配線と閉じ込め。
    #   lint はVBAしか読まず、LO はフォームを知らないので、その隙間だけを見る。
    #   HTML画面が実際に描かれるかは Windows 実機でしか確認できない(Z-43)。
    ("ui",           [sys.executable, "tools/ui_check.py"],
     r"結果: OK 6条件"),
    # 裁定書43 §2(司令塔の最終検問): 「作ったのに繋いでいない」検問を0にする。
    #   gate_count.py は GATES の全実行ファイルが要点行の契約に入っているか、
    #   件数がリテラル定数でないかを見る**検問の検問**なのに、GATES に載って
    #   おらず一括実行では一度も回っていなかった(自分自身を数える 33 本目)。
    ("gate-count",   [sys.executable, "tools/gate_count.py"],
     r"検査実施: [1-9]\d*件"),
    # 同上。S1 抽出品質ベンチの採点器(bench_s1.py)も一括実行から外れていた。
    #   selfcheck=採点の算術と変異注入、gold-check=5社の gold の検算。
    ("bench",        [sys.executable, "tools/bench_s1.py", "--mode", "selfcheck"],
     r"\[selfcheck\] OK"),
    ("bench-gold",   [sys.executable, "tools/bench_s1.py", "--gold-check"],
     r"\[gold-check\] OK"),
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
