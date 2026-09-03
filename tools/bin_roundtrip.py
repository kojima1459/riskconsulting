#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
bin_roundtrip.py - 配布 vbaProject.bin の読み戻し検問(裁定書27 W9-A 検証① /
                   新ゲート `bin-roundtrip`)

================================================================================
何を守るゲートか:
    配布方式B(裁定書27 W9-A)は「モジュールが最初から入った正規の
    vbaProject.bin をビルドが書き出す」方式である。書き出す側(build/ovba_write.py)
    が間違っても、書いた側の理屈で読み返しては何も検査したことにならない。
    本ツールは**配布物(dist/*.xlsm)の側から**bin を開き、

      [1] 各モジュールのソースが src/ の .bas/.cls(ヘッダ除去・CRLF・CP932)と
          **バイト一致**すること
      [2] モジュール数と集合が build/modules.json(その配布に載る分)と一致すること
      [3] 隠しシート vba_src が存在しないこと
      [4] bin に配布禁止の文字列("VBProject" / "AddFromString" /
          "ExecuteExcel4Macro" / "WScript.Shell" / "new:{")が現れないこと
          (裁定書27 W9-B 6。**圧縮を解いた本文で**検査する)
      [5] 全モジュールの MODULEOFFSET が 0(p-code キャッシュを持たない)こと

    を確かめる。読めない・数えられない・比較できないは**すべて失格**にする
    (「対象が見つからないので検査せず緑」を作らない)。

なぜ olevba を使うのか:
    [1] の解凍は oletools.olevba(第三者実装)で行う。ビルドが使った
    build/ovba.py の解凍器で読み返すと「自分の圧縮器のバグを自分の解凍器が
    帳消しにする」ため、往復検査としての意味が消える。
    olevba が入っていない場合は**緑にせず** exit 2(環境不備)で止める。

整形規則の値源:
    .bas → モジュールソースの整形(Attribute行/.clsヘッダの除去・改行正規化)は
    build/build_rpn.py の `_vba_src_text` **1実装だけ**を呼ぶ。ここで書き写すと
    ビルドと検問が別々に緩められるため、必ず import して使う。

使い方:
    python3 tools/bin_roundtrip.py                 # dist/ の dev と prod 両方
    python3 tools/bin_roundtrip.py --book dist/リスク提案ナビ.xlsm
    exit code: 0 = 全PASS / 1 = 失格 / 2 = 環境不備(oletools不在・ブック不在)
================================================================================
"""

from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path

TOOLS_DIR = Path(__file__).resolve().parent
REPO_ROOT = TOOLS_DIR.parent
sys.path.insert(0, str(REPO_ROOT / "build"))

import build_rpn      # noqa: E402  (_vba_src_text / 禁止文字列表 / 台帳読み)
import ovba_write     # noqa: E402  (dir の申告を読む)

try:
    from oletools.olevba import VBA_Parser
except ImportError:   # 第三者実装が無いと往復検査にならないので緑にしない
    VBA_Parser = None

DEFAULT_BOOKS = (
    REPO_ROOT / "dist" / "リスク提案ナビ_dev.xlsm",
    REPO_ROOT / "dist" / "リスク提案ナビ.xlsm",
)
DEV_SUFFIX = "_dev.xlsm"


def expected_sources(is_dev: bool) -> dict:
    """台帳から {モジュール名: 期待するモジュール本文(bytes)} を作る。"""
    modules = build_rpn.load_manifest(str(REPO_ROOT / "build" / "modules.json"))
    present, _missing = [], []
    for m in modules:
        if (REPO_ROOT / m["path"]).exists():
            present.append(m)
    shipped = build_rpn._shipped_modules(present, is_dev)
    out = {}
    for m in shipped:
        body = build_rpn._vba_src_text(str(REPO_ROOT), m)
        want = body.replace("\n", "\r\n").encode("cp932")
        if not want.endswith(b"\r\n"):
            want += b"\r\n"
        out[m["name"]] = want
    return out


def check_book(book: Path) -> list[str]:
    errors: list[str] = []
    is_dev = book.name.endswith(DEV_SUFFIX)
    print("=" * 78)
    print(f"ブック: {book}  ({'dev' if is_dev else 'prod'})")
    print("=" * 78)

    with zipfile.ZipFile(book) as z:
        names = z.namelist()
        if "xl/vbaProject.bin" not in names:
            return [f"{book.name}: xl/vbaProject.bin がありません"]
        vba_bin = z.read("xl/vbaProject.bin")

    # --- [3] vba_src シートの不在 -------------------------------------------
    import openpyxl
    wb = openpyxl.load_workbook(book, read_only=True, keep_links=False)
    sheetnames = list(wb.sheetnames)
    wb.close()
    if "vba_src" in sheetnames:
        errors.append(f"{book.name}: 隠しシート vba_src が残っています"
                      "(配布方式Bでは存在してはいけません)")
    print(f"[3] vba_src シート: {'あり(失格)' if 'vba_src' in sheetnames else 'なし'}"
          f" / シート{len(sheetnames)}枚")

    # --- [1] olevba で解凍して .bas とバイト比較 ------------------------------
    got: dict[str, bytes] = {}
    parser = VBA_Parser(str(book))
    try:
        for (_fn, stream, _vba_fn, code) in parser.extract_macros():
            name = stream.split("/")[-1]
            got[name] = code.encode("cp932", errors="replace") \
                if isinstance(code, str) else code
    finally:
        parser.close()
    print(f"[1] olevba が解凍したモジュール: {len(got)}本")

    want = expected_sources(is_dev)
    doc_names = {n for n, info in ovba_write.read_modules(vba_bin).items()
                 if info["type"] == "document"}
    # 配布方式Bで焼く document module は ThisWorkbook のみ(build_rpn.py
    # build_baked_vba_project)。Sheet1 等の他の document module は焼かない
    # (openpyxl 製の成果物ワークシートに codeName="Sheet1" が無く、焼くと
    # 名前だけの孤児モジュールになるため)。ここで document 集合が
    # {"ThisWorkbook"} ちょうどであることを確かめる(集合一致条件は緩めない)。
    if doc_names != {"ThisWorkbook"}:
        errors.append(
            f"{book.name}: document module 集合が {{'ThisWorkbook'}} と不一致"
            f"(実際: {sorted(doc_names)})。配布方式Bは ThisWorkbook 以外の "
            "document module を焼かない設計です。")
    std_got = {n: v for n, v in got.items() if n not in doc_names}

    for name, want_src in sorted(want.items()):
        if name not in std_got:
            errors.append(f"{book.name}: '{name}' が配布binにありません")
            continue
        actual = ovba_write.strip_attribute_lines(std_got[name])
        if actual != want_src:
            errors.append(
                f"{book.name}: '{name}' の本文が src/ と不一致"
                f"(期待{len(want_src)}バイト / 実際{len(actual)}バイト)")
    extra = sorted(set(std_got) - set(want))
    if extra:
        errors.append(f"{book.name}: 台帳に無いモジュールが載っています: {extra}")

    # --- [2] モジュール数と集合 ----------------------------------------------
    print(f"[2] モジュール集合: 台帳{len(want)}本 / bin(document除く){len(std_got)}本"
          f" / document module {sorted(doc_names)}")
    if len(std_got) != len(want):
        errors.append(
            f"{book.name}: モジュール数が台帳と不一致(台帳{len(want)} / bin{len(std_got)})")

    # --- [4] 禁止文字列 -------------------------------------------------------
    hits = build_rpn.forbidden_strings_in_bin(vba_bin)
    print(f"[4] 配布禁止文字列: {hits if hits else 'なし'}")
    if hits:
        errors.append(
            f"{book.name}: vbaProject.bin に配布禁止の文字列があります"
            f"(裁定書27 W9-B 6): {', '.join(hits)}")

    # --- [5] MODULEOFFSET=0 ---------------------------------------------------
    import struct
    import ovba
    dir_dec = ovba.ovba_decompress(ovba.CFBReader(vba_bin).read("dir"))
    offsets = [struct.unpack("<I", body)[0]
               for _o, rid, _s, body in ovba_write.iter_dir_records(dir_dec)
               if rid == ovba_write.REC_MODULEOFFSET]
    bad = [o for o in offsets if o != 0]
    print(f"[5] MODULEOFFSET: {len(offsets)}件すべて{'0' if not bad else '0ではない'}"
          "(p-codeキャッシュ無し)")
    if not offsets:
        errors.append(f"{book.name}: dir に MODULEOFFSET が1件もありません"
                      "(検査が成立していません)")
    if bad:
        errors.append(f"{book.name}: MODULEOFFSET≠0 のモジュールが{len(bad)}件"
                      "(p-codeキャッシュが混入しています)")
    return errors


def main() -> int:
    ap = argparse.ArgumentParser(
        description="配布 vbaProject.bin の読み戻し検問(裁定書27 W9-A)")
    ap.add_argument("--book", help="検査するブック(既定: dist/ の dev と prod 両方)")
    args = ap.parse_args()

    if VBA_Parser is None:
        print("ERROR: oletools が import できません(pip install oletools)。"
              "第三者実装で解凍できないと往復検査になりません。", file=sys.stderr)
        return 2

    if args.book:
        p = Path(args.book)
        books = [p if p.is_absolute() else REPO_ROOT / p]
    else:
        books = [p for p in DEFAULT_BOOKS if p.exists()]
    if not books:
        print("ERROR: dist/ にビルド済みブックがありません。"
              "先に python3 build/build_rpn.py --dev を実行してください。",
              file=sys.stderr)
        return 2

    print("=== bin_roundtrip.py (配布binを解凍して src/ とバイト比較) ===")
    errors: list[str] = []
    for b in books:
        if not b.exists():
            print(f"ERROR: ブックが見つかりません: {b}", file=sys.stderr)
            return 2
        errors.extend(check_book(b))
        print()

    print("-" * 78)
    if errors:
        print(f"結果: NG {len(errors)}件")
        for e in errors:
            print(f"  - {e}")
        return 1
    print(f"結果: OK 検査したブック{len(books)}冊 / 5条件"
          "(本文バイト一致・モジュール数と集合・vba_src不在・禁止文字列不在・"
          "MODULEOFFSET=0)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
