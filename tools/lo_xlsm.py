#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
lo_xlsm.py - 配布 .xlsm を LibreOffice に開かせて VBA を読み込ませる検問
             (裁定書27 W9-A 検証② / 新ゲート `lo-xlsm`)

================================================================================
何を守るゲートか:
    配布方式B(裁定書27 W9-A)では、ビルドが**完成品の vbaProject.bin** を
    ゼロから書き出す。バイト列としては tools/bin_roundtrip.py が「解凍したら
    src/ と一致する」ことを見るが、それは**入れ物(CFB/dir/PROJECT)が
    正しく組めているか**を何も保証しない。ヘッダが1バイト壊れていても
    自前のリーダーなら読めてしまう。
    そこで**別実装のOffice実装(LibreOffice)に実際に開かせる**。LOは xlsm の
    vbaProject.bin を取り込んでドキュメントのBasicライブラリを作るので、
      (1) LOがそのbinを解釈できたか
      (2) 全モジュールが名前どおり載っているか
    が外から確かめられる。

3段の検査(どれか1つでも欠ければ NG。骨抜き禁止):
    [1] LibreOffice が配布 .xlsm を読み込み、ドキュメントのBasicライブラリを
        列挙できること(=binが「Officeの実装で」解釈できた証跡)。
        ライブラリが1本も出てこない/ドキュメントが開けない、は**失格**であって
        「確認できなかったのでスキップ」にはしない。
    [2] 列挙されたモジュール名の集合が、配布 bin の dir が申告する集合と
        完全一致すること(順不同)。LOが黙って落としたモジュールを見つける。
    [3] **配布 bin から取り出したソース**で、既存 lo-compile(モード2)と同じ
        隔離ライブラリ方式のコンパイル確認を全モジュールに対して行うこと。
        src/ ではなく**出荷するバイト列側**を入力にする点だけが違う。

なぜ [3] を「ドキュメントのライブラリ全体を1回コンパイル」で済ませないのか:
    LibreOffice Basic には実Excelと違う既知の制約がある(tools/run_lo_tests.py
    の技術メモ 2/6/7)。`As String()` を返す関数宣言や、モジュールを跨ぐ
    Public Type の解決がその例で、**実Excelでは正しいコードがLOでだけ落ちる**。
    run_lo_tests はLO用の一時コピーにだけ機械的な書き換えを当ててこれを回避
    している。同じ回避を当てるには1モジュールずつ隔離ライブラリへ入れる必要が
    あり、ドキュメント内のライブラリをそのまま一括コンパイルする方式は
    「実Excelでは通るのに毎回赤」になって検問として使えない。
    したがって [3] は run_lo_tests の実装をそのまま呼ぶ(二重実装を作らない)。

使い方:
    python3 tools/lo_xlsm.py                     # dist/ の配布ブック
    python3 tools/lo_xlsm.py --book dist/リスク提案ナビ_dev.xlsm
    python3 tools/lo_xlsm.py --skip-compile      # [1][2] のみ(デバッグ用)
    exit code: 0 = 全PASS / 1 = いずれか失格 / 2 = 環境不備(soffice不在等)
================================================================================
"""

from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape as xml_escape

TOOLS_DIR = Path(__file__).resolve().parent
REPO_ROOT = TOOLS_DIR.parent
sys.path.insert(0, str(TOOLS_DIR))
sys.path.insert(0, str(REPO_ROOT / "build"))

import run_lo_tests as R      # noqa: E402  ([3] の実装を共有する)
import ovba_write             # noqa: E402  (配布binの読み戻し)

DEFAULT_BOOKS = (
    REPO_ROOT / "dist" / "リスク提案ナビ.xlsm",
    REPO_ROOT / "dist" / "リスク提案ナビ_dev.xlsm",
)

# LO が xlsm から取り込んだVBAを置くドキュメント内ライブラリ名(=VBAプロジェクト名)。
DOC_LIB_NAME = "VBAProject"

# ドキュメントを開いてBasicライブラリを列挙し、結果をファイルへ書く駆動マクロ。
# 例外は握りつぶさず "ERR ..." として書き出す(呼び出し側が失格にする)。
DRIVER_SRC = '''Sub Main
  Dim oDoc As Object, oLibs As Object, oLib As Object
  Dim sOut As String, i As Integer, j As Integer, nF As Integer
  Dim aArgs(1) As New com.sun.star.beans.PropertyValue
  On Error Goto Fail
  aArgs(0).Name = "Hidden" : aArgs(0).Value = True
  ' MacroExecMode.NEVER_EXECUTE(=0)。VBAは**取り込ませるが実行はさせない**。
  ' (4=USE_CONFIG_REJECT_CONFIRMATION は本ブックの読み込み自体を拒否して
  '  loadComponentFromURL が null を返す。実測 2026-09-03)
  aArgs(1).Name = "MacroExecutionMode" : aArgs(1).Value = 0
  oDoc = StarDesktop.loadComponentFromURL("__URL__", "_blank", 0, aArgs())
  If IsNull(oDoc) Then
    sOut = "ERR loadComponentFromURL returned null"
    Goto Emit
  End If
  oLibs = oDoc.BasicLibraries
  For i = 0 To UBound(oLibs.ElementNames)
    sOut = sOut & "LIB" & Chr(9) & oLibs.ElementNames(i) & Chr(10)
    If Not oLibs.isLibraryLoaded(oLibs.ElementNames(i)) Then
      oLibs.loadLibrary(oLibs.ElementNames(i))
    End If
    oLib = oLibs.getByName(oLibs.ElementNames(i))
    For j = 0 To UBound(oLib.ElementNames)
      sOut = sOut & "MOD" & Chr(9) & oLibs.ElementNames(i) & Chr(9) & _
             oLib.ElementNames(j) & Chr(9) & _
             CStr(Len(oLib.getByName(oLib.ElementNames(j)))) & Chr(10)
    Next j
  Next i
  sOut = sOut & "OK" & Chr(10)
  Goto Emit
Fail:
  sOut = sOut & "ERR " & Err & " " & Error$ & " at line " & Erl & Chr(10)
Emit:
  nF = FreeFile
  Open "__RES__" For Output As #nF
  Print #nF, sOut
  Close #nF
  If Not IsNull(oDoc) Then oDoc.close(False)
End Sub
'''


def read_bin_modules(book: Path) -> dict:
    """配布 .xlsm の vbaProject.bin から {モジュール名: {source,type}} を読む。"""
    with zipfile.ZipFile(book) as z:
        if "xl/vbaProject.bin" not in z.namelist():
            raise SystemExit(f"ERROR: {book} に xl/vbaProject.bin がありません")
        return ovba_write.read_modules(z.read("xl/vbaProject.bin"))


def run_lo_open(book: Path, work: Path, timeout_sec: int, verbose: bool):
    """[1][2] 用: LOでブックを開き、(ライブラリ名, モジュール名) を回収する。"""
    soffice = R.find_soffice()
    template = R.ensure_template_profile(soffice, verbose)
    profile = work / "profile_open"
    R.fresh_profile_copy(template, profile)

    res = work / "lo_xlsm_result.txt"
    # 日本語を含むパスの file:// URL を LO Basic の loadComponentFromURL に渡すと
    # 黙って null が返る(実測 2026-09-03)。検問の対象はブックの中身であって
    # ファイル名ではないので、ASCIIの一時パスへ複製してから開かせる。
    ascii_book = work / "book_under_test.xlsm"
    shutil.copyfile(book, ascii_book)
    src = DRIVER_SRC.replace("__URL__", ascii_book.resolve().as_uri()) \
                    .replace("__RES__", str(res))
    lib = profile / "user" / "basic" / "LoXlsmDrv"
    lib.mkdir(parents=True, exist_ok=True)
    (lib / "Drv.xba").write_text(
        R.XBA_TEMPLATE.format(name="Drv", body=xml_escape(src)), encoding="utf-8")
    (lib / "script.xlb").write_text(
        R.XLB_TEMPLATE.format(libname="LoXlsmDrv",
                              elements=' <library:element library:name="Drv"/>'),
        encoding="utf-8")
    R.register_libraries(profile, ["LoXlsmDrv"])

    uri = ("vnd.sun.star.script:LoXlsmDrv.Drv.Main"
           "?language=Basic&location=application")
    rc, out, err = R.run_uri(soffice, profile, uri, timeout_sec)
    text = res.read_text(encoding="utf-8", errors="replace") if res.exists() else ""
    if verbose:
        print(f"    soffice exit={rc}")
        if out.strip():
            print("    stdout:", out.strip()[:2000])
        if err.strip():
            print("    stderr:", err.strip()[:2000])
    return rc, text


def main() -> int:
    ap = argparse.ArgumentParser(
        description="配布xlsmをLibreOfficeに開かせるVBA検問(裁定書27 W9-A)")
    ap.add_argument("--book", help="検査するブック(既定: dist/ の prod -> dev)")
    ap.add_argument("--timeout", type=int, default=180, help="1呼び出しの上限秒")
    ap.add_argument("--skip-compile", action="store_true",
                    help="[3] コンパイル確認を省く(デバッグ用。ゲートでは使わない)")
    ap.add_argument("--verbose", action="store_true")
    args = ap.parse_args()

    if args.book:
        book = Path(args.book)
        if not book.is_absolute():
            book = REPO_ROOT / book
    else:
        book = next((p for p in DEFAULT_BOOKS if p.exists()), None)
        if book is None:
            print("ERROR: dist/ に配布ブックがありません。"
                  "先に python3 build/build_rpn.py --prod を実行してください。",
                  file=sys.stderr)
            return 2
    if not book.exists():
        print(f"ERROR: ブックが見つかりません: {book}", file=sys.stderr)
        return 2

    print("=== lo_xlsm.py (配布binをLibreOfficeに読み込ませる検問) ===")
    print(f"ブック: {book}")

    bin_mods = read_bin_modules(book)
    print(f"配布binのモジュール: {len(bin_mods)}本"
          f"(うち document module "
          f"{sum(1 for v in bin_mods.values() if v['type'] == 'document')}本)")

    work = Path(tempfile.mkdtemp(prefix="rpn_lo_xlsm_"))
    failures: list[str] = []
    try:
        # --- [1] LOで開けるか -------------------------------------------------
        print("\n[1] LibreOffice が配布 .xlsm を読み込み、Basicライブラリを列挙できること")
        rc, text = run_lo_open(book, work, args.timeout, args.verbose)
        if rc != 0:
            failures.append(f"soffice が異常終了しました(exit={rc})")
        if not text.strip():
            failures.append("駆動マクロが結果を1行も書きませんでした"
                            "(LOがドキュメントを開けていない可能性)")
        for line in text.splitlines():
            if line.startswith("ERR"):
                failures.append(f"LO側の例外: {line}")
        libs, got_mods = [], {}
        for line in text.splitlines():
            parts = line.split("\t")
            if parts[0] == "LIB" and len(parts) >= 2:
                libs.append(parts[1])
            elif parts[0] == "MOD" and len(parts) >= 4:
                got_mods[parts[2]] = (parts[1], int(parts[3]))
        print(f"    ライブラリ: {libs}")
        if DOC_LIB_NAME not in libs:
            failures.append(
                f"ドキュメントに '{DOC_LIB_NAME}' ライブラリがありません"
                f"(実際: {libs})。LOが vbaProject.bin を取り込めていません")
        if "OK" not in text.split():
            failures.append("駆動マクロが最後まで到達しませんでした(OK行なし)")

        # --- [2] モジュール集合の一致 ----------------------------------------
        print("\n[2] LOが列挙したモジュール集合 ⇔ 配布binの dir が申告する集合")
        want = sorted(bin_mods)
        got = sorted(n for n, (lb, _ln) in got_mods.items() if lb == DOC_LIB_NAME)
        print(f"    bin={len(want)}本 / LO={len(got)}本")
        if want != got:
            only_bin = [n for n in want if n not in got]
            only_lo = [n for n in got if n not in want]
            failures.append(
                f"モジュール集合が不一致(binのみ={only_bin} / LOのみ={only_lo})")
        empty = sorted(n for n, (lb, ln) in got_mods.items()
                       if lb == DOC_LIB_NAME and ln == 0
                       and len(bin_mods.get(n, {}).get("source", b"")) > 0)
        if empty:
            failures.append(f"LO側で本文が空になったモジュール: {empty}")

        # --- [3] 配布binのソースでコンパイル確認 -------------------------------
        if args.skip_compile:
            print("\n[3] コンパイル確認: --skip-compile のため省略(ゲートでは使わない)")
            failures.append("--skip-compile が指定されました"
                            "(ゲートとしては未実施=失格)")
        else:
            print("\n[3] 配布binから取り出したソースで全モジュールのコンパイル確認"
                  "(lo-compile と同じ隔離ライブラリ方式)")
            src_dir = work / "bin_src"
            src_dir.mkdir(parents=True, exist_ok=True)
            all_modules: dict[str, Path] = {}
            for name, info in sorted(bin_mods.items()):
                if info["type"] == "document":
                    # document module(ThisWorkbook / Sheet1)は隔離ライブラリでは
                    # Workbook_Open などのイベント宣言が解決できない。実Excel側の
                    # 契約は最小ThisWorkbookで、中身は bin_roundtrip がバイト比較
                    # しているため、ここでは対象外にする(理由を明示して除外)。
                    continue
                p = src_dir / f"{name}.bas"
                p.write_text(info["source"].decode("cp932", errors="replace"),
                             encoding="utf-8")
                all_modules[name] = p
            print(f"    対象モジュール数: {len(all_modules)}"
                  f"(document module は除外)")
            soffice = R.find_soffice()
            template = R.ensure_template_profile(soffice, args.verbose)
            type_blocks = R.collect_public_type_blocks(all_modules)
            ok, results = R.run_compile_mode(
                soffice, template, all_modules, work, args.timeout,
                args.verbose, type_blocks)
            if not ok:
                bad = [n for n, o, _d in results if not o]
                failures.append(f"コンパイルに失敗したモジュール: {bad}")
    finally:
        shutil.rmtree(work, ignore_errors=True)

    print("\n" + "-" * 78)
    if failures:
        print(f"結果: NG {len(failures)}件")
        for f in failures:
            print(f"  - {f}")
        return 1
    print("結果: OK 3条件(LOが配布binを読み込めた / モジュール集合一致 / "
          "全モジュールのコンパイル成功)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
