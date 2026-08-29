#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_rpn.py - 「リスク提案ナビ」ビルドスクリプト

dist/リスク提案ナビ.xlsm(または dist/リスク提案ナビ_dev.xlsm)を生成する。

--------------------------------------------------------------------------
移植元: PoC「マイ本棚AI」 build/build_mybookshelf.py。
    modules.json 方式(台帳に列挙した .bas のソースを vba_src シートへ1行
    1モジュールで焼き込み、起動時に VBE へ注入する)と、ビルド後自己検証
    (シート存在・モジュール数一致・30,000字/32,000字上限・ソース一致)は維持した。

**移植にあたって完全に削除したもの(16章 NFR-S2)**:
    ・キー難読化(OBF1)の埋め込み処理 obfuscate_secret() と OBF1 接頭辞
    ・環境変数からAPIキー/URLを読んで config へ焼き込む処理
    ・発行キー(publish_key)関連の分岐
  本製品はキーをブックに一切入れない。難読化は「漏れてないように見せる」だけで、
  鍵とアルゴリズムが同居すれば復元可能であり、本PJのキーはローテーション不可
  (借用の共用キー1本)なので一度露出したら恒久被害になる。
  **このスクリプトはキーに触れない**。api_key.txt も環境変数も読まない
  (17章 T-46② の検問がその不在を毎リリース grep で検査する)。

アーキテクチャ:
  1. build/sheets_main.json(シート台帳)と build/modules.json(モジュール台帳)
     を読む。シート構成の正はこの2ファイルだけで、コード側に散らさない。
  2. openpyxl でブックを組み立てる。build/template_skeleton.xlsm があれば
     keep_vba=True で読み込み、その vbaProject.bin をそのまま引き継ぐ。
     無ければ新規ブックとして組み立て、保存後に [Content_Types].xml の
     ワークブックパートだけを macroEnabled 用へ書き換える(拡張子と中身の
     不整合で Excel が警告を出すのを防ぐ)。
  3. vba_src シートへ modules.json のモジュールソースを格納する。
  4. 一時パスへ書き出し、**自己検証に合格してから**正規パスへ確定する
     (検証に落ちた不良品は *.failed.xlsm へ退避し、前回の良品には触れない)。

このスクリプトが書き込むのは <repo>/build と <repo>/dist の配下のみ。
--------------------------------------------------------------------------
"""

from __future__ import annotations

import argparse
import io
import json
import os
import re
import struct
import sys
import tempfile
import zipfile

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill, Protection
from openpyxl.utils import get_column_letter, quote_sheetname
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

try:
    import olefile
except ImportError:      # 検証の一部が使えなくなるだけでビルド自体は続けられる
    olefile = None

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# build/ovba.py(自己完結のOVBA圧縮/解凍・CFBリーダー)を import できるようにする。
# 通常は `python3 build/build_rpn.py` 実行時に sys.path[0] が build/ になるが、
# 別ディレクトリからの import 実行でも解決できるよう明示的に足す。
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)
import ovba  # noqa: E402  (自己完結モジュール。ThisWorkbook外科パッチに使う)

DEFAULT_ROOT = os.path.dirname(SCRIPT_DIR)
DEFAULT_MODULES_JSON = os.path.join(SCRIPT_DIR, "modules.json")
DEFAULT_SHEETS_JSON = os.path.join(SCRIPT_DIR, "sheets_main.json")
DEFAULT_TEMPLATE = os.path.join(SCRIPT_DIR, "template_skeleton.xlsm")
DEFAULT_SHEETS_KB_JSON = os.path.join(SCRIPT_DIR, "sheets_kb.json")
KB_DEFAULT_RESERVE_ROWS = 30

APP_TITLE = "リスク提案ナビ"
EXCEL_CELL_LIMIT = 32000        # Excelの技術上限(セル1個あたりの文字数)
MODULE_CONTRACT_LIMIT = 30000   # 12章§2 の契約上限(1モジュールあたり)

VALID_ROLES = ("core", "app", "ui", "test")

# 入力規則(データの入力規則リスト)をセル内リテラルで書くときの上限。Excelの仕様で
# formula1 は 255 字まで。超える集合(case_data の data_key 28値など)は静的リストでは
# 表現できないため、ビルドは「省略した事実」を標準出力へ出して先送りする
# (最終形は 11章§5 の隠しレンジ方式。W3 T-31/T-32)。
DV_INLINE_LIMIT = 255

# 表の既定確保行数(雛形として styling / 入力規則を敷いておく行数)。
DEFAULT_FLAT_RESERVE_ROWS = 50

HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
TITLE_FONT = Font(bold=True, size=12)
HEADER_FONT = Font(bold=True)
NOTE_FONT = Font(size=9, color="808080")
SHEET_TITLE_FONT = Font(bold=True, size=14)
LOCKED = Protection(locked=True)
UNLOCKED = Protection(locked=False)

_ILLEGAL_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

WORKBOOK_CT_XLSX = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
WORKBOOK_CT_XLSM = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"


def _clean(s):
    """XMLに書けない制御文字を除去する(PoC実証済みの防御処理を踏襲)。"""
    return _ILLEGAL_XML.sub("", s) if isinstance(s, str) else s


class BuildError(Exception):
    """ビルド契約違反(モジュール欠落・文字数超過など)。呼び出し側でexit 1にする。"""


# ===========================================================================
# 自己インストーラ(VBA/ThisWorkbook ストリームの中身)と vbaProject.bin 外科パッチ
# ---------------------------------------------------------------------------
# 移植元: PoC「マイ本棚AI」 build/build_mybookshelf.py の自己インストーラ機構
#   (ThisWorkbook ストリームの外科的差し替え + dir MOFFSET=0 + _VBA_PROJECT 無害化)。
#   PoC固有の起動先(modViewport / modInstallCheck / RunFirstRunPromptEarly)は
#   RPNには存在しないため落とし、起動先を **modBoot.Boot** に読み替えた
#   (12章§2.1: modBoot が Workbook_Open から呼ぶ唯一の起動入口)。低レベルの
#   OVBA圧縮/解凍・CFBリーダーは build/ovba.py(自己完結・実証済み)を使う。
#
# 建付け(12章§2 の自己インストール機構):
#   ・ビルドは template_skeleton.xlsm の「本物の vbaProject.bin」をそのまま成果物へ
#     持ち込む(どのExcelでも文句なく開けるのはこのため)。
#   ・その ThisWorkbook ストリームだけを下記の自己インストーラソースへ差し替える。
#     起動時(Workbook_Open)に vba_src シートの各モジュールをVBEへ注入し、
#     以後は通常モジュールとして modBoot.Boot(§2.1の起動シーケンス)へ渡す。
#
# 注意: このVBAソースは ThisWorkbook ストリームへ「元と同じ圧縮後バイト長」で
#   差し込む(in-place外科パッチ。ストリームを伸ばすとCFBのFATを組み直す必要が
#   あり、Excelが読めなくなる)。ソースは必ずASCIIのみ(CP932安全)。長い説明を
#   足すと圧縮後サイズ上限(テンプレ実測1,148B)を超えてビルドが落ちるため、
#   意図の説明はコメント側(ここ)へ書く。
#
#   ・モジュール追加(Add/Name/DeleteLines/AddFromString)の失敗は1本ずつローカルに
#     握って f を加算し、f>0 のときは Save しない。半端な注入状態をファイルへ
#     焼き付けないため。あわせて ThisWorkbook.Saved=True を立て、閉じる際の
#     「保存しますか?」の反射押しで半端状態が保存されるのを防ぐ(PoC実機由来)。
#   ・本文があるはず(LenB(s)>0)なのに注入後 CountOfLines<1 の無言破損も f 加算。
#   ・Boot は同期呼び出しだと注入直後に 1004 になることがあるため OnTime で1秒
#     後ろへ切り離し、予約時刻を vba_src!E1 に置く(modBoot 側が任意で取り消す)。
#     OnTime予約自体が失敗したら同期で modBoot.Boot を呼ぶフォールバックを持つ。
# ===========================================================================
_INSTALLER_SRC_TEXT = '''Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Workbook_Open()
  Install
End Sub
Public Sub Install()
  Dim p As Object, w As Worksheet, c As Object, e As Object
  Dim r As Long, n As String, s As String, l As Long, f As Long
  On Error GoTo Trust
  Set p = ThisWorkbook.VBProject
  On Error GoTo Done
  Set w = ThisWorkbook.Worksheets("vba_src")
  l = w.Cells(w.Rows.Count, 1).End(-4162).Row
  For r = 2 To l
    n = CStr(w.Cells(r, 1).Value)
    s = CStr(w.Cells(r, 3).Value)
    If LenB(n) > 0 Then
      On Error Resume Next
      Set e = Nothing: Set e = p.VBComponents(n)
      If Not e Is Nothing Then p.VBComponents.Remove e
      Set c = Nothing
      Set c = p.VBComponents.Add(1)
      If c Is Nothing Then
        f = f + 1
      Else
        Err.Clear
        c.Name = n
        If c.CodeModule.CountOfLines > 0 Then c.CodeModule.DeleteLines 1, c.CodeModule.CountOfLines
        If LenB(s) > 0 Then c.CodeModule.AddFromString s
        If Err.Number <> 0 Then
          f = f + 1
        ElseIf LenB(s) > 0 Then
          If c.CodeModule.CountOfLines < 1 Then f = f + 1
          If Err.Number <> 0 Then f = f + 1
        End If
      End If
      Err.Clear
      On Error GoTo Done
    End If
  Next r
  On Error Resume Next
  If f > 0 Then
    MsgBox "Setup NG(" & f & "). Close WITHOUT saving, then reopen.", vbCritical
    ThisWorkbook.Saved = True
    Exit Sub
  End If
  ThisWorkbook.Save
  Err.Clear
  Dim bt As Date
  bt = Now + TimeSerial(0, 0, 1)
  Application.OnTime bt, "'" & ThisWorkbook.Name & "'!modBoot.Boot"
  If Err.Number <> 0 Then
    Err.Clear
    Application.Run "modBoot.Boot"
  Else
    w.Cells(1, 5).Value = CDbl(bt)
    Err.Clear
  End If
  Exit Sub
Trust:
  MsgBox "Trust the VBA project, then reopen.", vbCritical
  Exit Sub
Done:
End Sub
'''


def build_installer_src() -> bytes:
    """自己インストーラソースをCP932(実体はASCIIのみ)・CRLFのバイト列にする。"""
    try:
        _INSTALLER_SRC_TEXT.encode("ascii")
    except UnicodeEncodeError as e:
        raise BuildError(
            f"自己インストーラソースはASCIIのみで書いてください(CP932安全): {e}")
    return _INSTALLER_SRC_TEXT.replace("\n", "\r\n").encode("cp932")


def _pad_to_exact_or_die(compressed: bytes, target: int, stream_name: str) -> bytes:
    """ovba.pad_to_exact の到達不能差分(1/2/4/7バイト)を BuildError化して誘導する。
    OVBA空チャンクは3/5バイト単位でしか長さを埋められないため、これらの差分は
    原理的に到達不能。ovba.py はプロダクト固有ロジックを持たない自己完結モジュール
    という設計方針のため、誘導文はここ(呼び出し側)で出す。"""
    try:
        return ovba.pad_to_exact(compressed, target)
    except ValueError as e:
        raise BuildError(
            f"{stream_name}ストリームのpaddingが目標バイト数に到達できません({e})。"
            "OVBA空チャンクは3バイト単位/5バイト単位の組合せでしか長さを埋められず、"
            "元サイズとの差分が1/2/4/7バイトのときは到達不能です。"
            "_INSTALLER_SRC_TEXT のコメント・変数名を1～2バイト増減してから再実行してください。")


# ---------------------------------------------------------------------------
# _VBA_PROJECTストリームの無害化(幽霊コンパイルエラー根治。PoC R23c-F1由来)
#   配布xlsmの vbaProject.bin は template_skeleton.xlsm 由来で、その
#   _VBA_PROJECT ストリーム(3,061B)が当時のOfficeビルドのVersionスタンプと
#   PerformanceCacheを保持したまま。MS-OVBAは「Versionが開き手のOfficeと一致
#   すると PerformanceCache がソースより優先される」「書き手は Version=0xFFFF と
#   し PerformanceCache を含めてはならない(MUST)」と定める。ユーザーのExcel
#   ビルドと一致すると、注入した新ソースでなく古いキャッシュ側の名前解決が
#   信用され、実在するはずのプロシージャが見つからないコンパイルエラーになる。
#   対策(二重防御): Version を 0xFFFF へ書き換え、PerformanceCache 全体をゼロ埋め。
#   ストリーム長は 3,061B のまま変えない(olefile.write_stream は同サイズ書込のみ)。
# ---------------------------------------------------------------------------
_VBA_PROJECT_STREAM_SIZE = 3061       # template_skeleton.xlsm 由来の固定長
_VBA_PROJECT_RESERVED1 = 0x61CC       # MS-OVBA 2.3.4.1 Reserved1(固定値)
_VBA_PROJECT_VERSION_IGNORE = 0xFFFF  # 「キャッシュを使うな」の相互運用値
_VBA_PROJECT_CACHE_OFFSET = 7         # PerformanceCache の開始オフセット


def _neutralize_vba_project(stream: bytes) -> bytes:
    """_VBA_PROJECTストリームを同サイズのままキャッシュ無効化する。"""
    if len(stream) != _VBA_PROJECT_STREAM_SIZE:
        raise BuildError(
            f"_VBA_PROJECTストリームのサイズが想定外です: "
            f"期待={_VBA_PROJECT_STREAM_SIZE}バイト 実際={len(stream)}バイト "
            "(template_skeleton.xlsm が差し替わった可能性。同サイズ書換の前提が崩れます)")
    reserved1 = struct.unpack("<H", stream[0:2])[0]
    if reserved1 != _VBA_PROJECT_RESERVED1:
        raise BuildError(
            f"_VBA_PROJECTストリームの先頭2バイトが 0x{_VBA_PROJECT_RESERVED1:04X} "
            f"ではありません(実際=0x{reserved1:04X})。MS-OVBA の _VBA_PROJECT ヘッダ"
            "として解釈できないため中断します。")
    out = bytearray(stream)
    struct.pack_into("<H", out, 2, _VBA_PROJECT_VERSION_IGNORE)
    for i in range(_VBA_PROJECT_CACHE_OFFSET, len(out)):
        out[i] = 0
    return bytes(out)


def patch_installer(vba_bin: bytes, installer_src: bytes) -> bytes:
    """template由来の vbaProject.bin に外科パッチを当てる:
    (1) VBA/ThisWorkbook を自己インストーラソースへ差し替え(元と同じ圧縮後バイト長)
    (2) VBA/dir の ThisWorkbook.MOFFSET を 0 に書き換え
    (3) VBA/_VBA_PROJECT の PerformanceCache 無害化
    バイナリ全体のバイト長は不変(in-place)。"""
    if olefile is None:
        raise BuildError(
            "olefile が import できないため vbaProject.bin の外科パッチを実施できません"
            "(pip install olefile が必要。テンプレートを使うビルドには必須です)")
    skel = ovba.CFBReader(vba_bin)

    # (2) dir ストリームの ThisWorkbook.MOFFSET を 0 に。
    dir_dec = ovba.ovba_decompress(skel.read("dir"))
    needle = struct.pack("<HI", 0x0019, len("ThisWorkbook")) + b"ThisWorkbook"
    idx = dir_dec.find(needle)
    if idx < 0:
        raise BuildError("dir stream: ThisWorkbook MNAME レコードが見つかりません"
                         "(template_skeleton.xlsm が非互換の可能性)")
    i = idx
    found = False
    while i < len(dir_dec):
        rid = struct.unpack("<H", dir_dec[i:i + 2])[0]
        sz = struct.unpack("<I", dir_dec[i + 2:i + 6])[0]
        if rid == 0x0031:  # MOFFSET
            patched = bytearray(dir_dec)
            struct.pack_into("<I", patched, i + 6, 0)
            dir_dec = bytes(patched)
            found = True
            break
        i += 6 + sz
    if not found:
        raise BuildError("dir stream: ThisWorkbook MOFFSET レコードが見つかりません")

    orig_dir_size = skel.entries["dir"]["size"]
    orig_tw_size = skel.entries["ThisWorkbook"]["size"]
    new_dir = _pad_to_exact_or_die(ovba.ovba_compress(dir_dec), orig_dir_size, "dir")

    # (1) ThisWorkbook ストリームは「元と同じバイト数」でしか差し替えられない。
    tw_compressed = ovba.ovba_compress(installer_src)
    if len(tw_compressed) > orig_tw_size:
        raise BuildError(
            f"自己インストーラ(ThisWorkbookストリーム)が圧縮後{len(tw_compressed)}バイトで、"
            f"差し替え可能な上限{orig_tw_size}バイトを{len(tw_compressed) - orig_tw_size}"
            "バイト超過しました。_INSTALLER_SRC_TEXT のコメント/変数名を削るか、処理を "
            "modBoot 側(vba_srcから注入される標準モジュール。サイズ上限が緩い)へ移してください。")
    new_tw = _pad_to_exact_or_die(tw_compressed, orig_tw_size, "ThisWorkbook")

    # (3) PerformanceCache 無害化。同サイズ書き込みのみ。
    new_vbaproj = _neutralize_vba_project(skel.read("_VBA_PROJECT"))

    buf = io.BytesIO(vba_bin)
    ole = olefile.OleFileIO(buf, write_mode=True)
    ole.write_stream("VBA/dir", new_dir)
    ole.write_stream("VBA/ThisWorkbook", new_tw)
    ole.write_stream("VBA/_VBA_PROJECT", new_vbaproj)
    ole.close()
    buf.seek(0)
    return buf.read()


# ---------------------------------------------------------------------------
# 台帳の読み込みと検証
# ---------------------------------------------------------------------------
def load_manifest(path):
    with open(path, encoding="utf-8") as fp:
        data = json.load(fp)
    modules = data["modules"]
    # name/path の重複検出。インストーラは行順に Remove -> Add するため、
    # 重複があると「後の行のソースが無言で勝つ」。
    seen_names, seen_paths = {}, {}
    for m in modules:
        for key in ("name", "path", "role"):
            if key not in m:
                raise BuildError(f"modules.json: エントリに必須キー'{key}'がありません: {m}")
        if m["role"] not in VALID_ROLES:
            raise BuildError(
                f"modules.json: {m['name']} の role が不正です: {m['role']}"
                f"(12章§2の層 {VALID_ROLES} のいずれか)")
        nm, pth = m["name"], m["path"]
        if nm in seen_names:
            raise BuildError(
                f"modules.json: name '{nm}' が重複しています"
                f"(既存: {seen_names[nm]['path']} / 重複: {pth})。"
                "後の行が無言で勝つ事故を防ぐため重複を解消してください。")
        seen_names[nm] = m
        if pth in seen_paths:
            raise BuildError(
                f"modules.json: path '{pth}' が重複しています"
                f"(name={seen_paths[pth]['name']!r} と name={nm!r})。")
        seen_paths[pth] = m
    return modules


def load_sheets(path):
    with open(path, encoding="utf-8") as fp:
        data = json.load(fp)
    sheets = data.get("sheets") or []
    if not sheets:
        raise BuildError("sheets_main.json: sheets が空です")
    names = [s["name"] for s in sheets]
    if len(names) != len(set(names)):
        raise BuildError(f"sheets_main.json: シート名が重複しています: {names}")
    roles = [s.get("role") for s in sheets]
    if roles.count("guard") != 1:
        raise BuildError("sheets_main.json: role=guard のシートはちょうど1枚必要です")
    if roles.count("vba_src") != 1:
        raise BuildError("sheets_main.json: role=vba_src のシートはちょうど1枚必要です")
    if sheets[0].get("role") != "guard":
        raise BuildError(
            "sheets_main.json: role=guard のシートは配列の先頭に置いてください"
            "(マクロ無効で開かれたとき最初に見える必要があるため)")
    guard = data.get("guard_sheet")
    if guard and guard != sheets[0]["name"]:
        raise BuildError(
            f"sheets_main.json: guard_sheet({guard})と先頭シート({sheets[0]['name']})が不一致")
    return data, sheets


# src/ 配下に実在するのに modules.json へ未登録の .bas を検出してビルドを止める。
# 台帳との突き合わせだけでは、台帳に無いものは永久に検出できない。
# 意図的にビルドへ含めないファイルは EXCLUDE に明記する。
UNREGISTERED_EXCLUDE: set[str] = set()


def check_unregistered(modules, root):
    import glob as _glob
    registered = {m["path"].replace("\\", "/") for m in modules}
    found = set()
    for p in _glob.glob(os.path.join(root, "src", "**", "*.bas"), recursive=True):
        found.add(os.path.relpath(p, root).replace("\\", "/"))
    orphans = sorted(found - registered - UNREGISTERED_EXCLUDE)
    if orphans:
        print("modules.json に未登録の .bas があります(実機でコンパイルエラーになります):",
              file=sys.stderr)
        for o in orphans:
            print(f"  {o}", file=sys.stderr)
        print("  -> build/modules.json に追加するか、意図的に除外するなら "
              "UNREGISTERED_EXCLUDE へ明記してください。", file=sys.stderr)
        sys.exit(1)


def validate_modules(modules, root, allow_missing):
    check_unregistered(modules, root)
    present, missing = [], []
    for m in modules:
        p = os.path.join(root, m["path"])
        (present if os.path.exists(p) else missing).append(m)

    if missing:
        print(f"モジュールファイルが見つかりません({len(missing)}件):", file=sys.stderr)
        for m in missing:
            print(f"  [{m['role']:<4}] {m['name']:<20} -> {m['path']}", file=sys.stderr)
        if not allow_missing:
            print("  (--allow-missing を付けると警告に緩和して続行できます)", file=sys.stderr)
            sys.exit(1)
        print("  --allow-missing 指定のため警告として続行します。", file=sys.stderr)
    return present, missing


# ---------------------------------------------------------------------------
# vba_src への焼き込み規則(唯一の実装)
# ---------------------------------------------------------------------------
def _vba_src_modules(present_modules):
    """vba_src シートへ載せる対象モジュールだけを返す。
    クラスモジュール(VBComponents.Add(1)で追加できない)と、台帳で
    vba_src=false にされたものは対象外。ビルドと自己検証で同じ判定を使う。"""
    return [m for m in present_modules
            if m.get("type") != "class" and m.get("vba_src") is not False]


def _vba_src_text(root, m):
    """src/ の .bas から vba_src セルへ格納する文字列を作る(ビルド規則の唯一の実装)。
    Attribute 行を落とし、_clean() で XML に書けない制御文字を除去し、改行をLFへ
    正規化する。**整形ロジックはここ1箇所に閉じる**。自己検証はこの関数の戻り値と
    成果物のC列を突き合わせるので、二重実装があると検査が無意味になる。"""
    path = os.path.join(root, m["path"])
    with open(path, encoding="utf-8-sig") as fp:
        txt = fp.read()
    txt = txt.replace("\r\n", "\n").replace("\r", "\n")
    out_lines = []
    for line in txt.split("\n"):
        stripped = line.lstrip("\ufeff")   # 先頭のBOM(U+FEFF)を落とす
        if stripped.lstrip().startswith("Attribute "):
            continue
        out_lines.append(stripped)
    cleaned = _clean("\n".join(out_lines))

    # 整形後ソースの先頭行が空白のみならビルドを止める。行数の期待値計算
    # (_expected_line_count)が落とすのは【末尾】の空行だけで、【先頭】の空行は
    # 1行として数える(=非対称)。先頭空行のモジュールを許すと、将来
    # 「先頭空行トリムだけが期待/実測のどちらかにだけ紛れ込む」実装変更で
    # D列突合が恒久的な偽陽性になりかねない。
    if cleaned.split("\n", 1)[0].strip() == "":
        raise BuildError(
            f"{m['name']}.bas: 整形後ソースの先頭行が空白のみです"
            "(先頭空行は期待/実測の行数計算が非対称になるため禁止)")
    return cleaned


def _expected_line_count(src):
    """ソース文字列の「末尾空行を落とした行数」を返す(純関数)。
    起動時の注入結果を行数で検算するための期待値。空文字列は0行。"""
    s = src.replace("\r\n", "\n").replace("\r", "\n")
    if s == "":
        return 0
    parts = s.split("\n")
    n = len(parts)
    while n > 0 and parts[n - 1].replace("\t", " ").strip() == "":
        n -= 1
    return n


# ---------------------------------------------------------------------------
# シート生成
# ---------------------------------------------------------------------------
class BuildCtx:
    """シート生成中に持ち回る文脈。生成しながら「後で検証すべき期待値」を溜める。
    期待値をビルド時に組み立てておくことで、自己検証が台帳を二度解釈せずに済む。"""

    def __init__(self, wb, data, mock_llm, app_version, present_modules, root):
        self.wb = wb
        self.enums = data.get("enums") or {}
        self.policy = data.get("protection_policy") or {}
        self.text_fmt = data.get("text_number_format", "@")
        self.mock_llm = mock_llm
        self.app_version = app_version
        self.present_modules = present_modules
        self.root = root
        # 期待値(自己検証がそのまま使う)
        self.names = []          # [(name, sheet, coord)]
        self.flat_headers = {}   # sheet -> [列名]
        self.block_headers = {}  # sheet -> [(block名, [列名])]
        self.dv_skipped = []     # [(sheet, target, enumキー, 長さ)]
        self.injected = []


def _add_name(ctx, ws, name, col, row):
    """ブック内で一意の名前付きレンジを1セルに定義する(13章§2.9)。"""
    coord = f"{get_column_letter(col)}{row}"
    ref = f"{quote_sheetname(ws.title)}!${get_column_letter(col)}${row}"
    if any(n == name for n, _, _ in ctx.names):
        raise BuildError(f"名前付きレンジ '{name}' が重複しています(ブック内で一意。13章§2.9)")
    ctx.wb.defined_names.add(DefinedName(name, attr_text=ref))
    ctx.names.append((name, ws.title, coord))


def _apply_dv(ctx, ws, enum_key, target):
    """静的リストの入力規則を張る。W0は静的リスト(11章§5の隠しレンジ方式はW3)。"""
    values = ctx.enums.get(enum_key)
    if not values:
        raise BuildError(
            f"sheets_main.json: 入力規則 '{enum_key}' が enums に定義されていません"
            f"(シート {ws.title})")
    formula = '"' + ",".join(values) + '"'
    if len(formula) > DV_INLINE_LIMIT:
        ctx.dv_skipped.append((ws.title, target, enum_key, len(formula)))
        return False
    dv = DataValidation(type="list", formula1=formula, allow_blank=True,
                        showErrorMessage=True,
                        errorTitle="入力できない値です",
                        error="一覧から選んでください(19章§3のenumレジストリが正)。")
    ws.add_data_validation(dv)
    dv.add(target)
    return True


def _style_header_cell(cell, note):
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(vertical="center", wrap_text=True)
    cell.protection = LOCKED
    if note:
        c = Comment(_clean(note), "sheets_main.json")
        c.width, c.height = 320, 100
        cell.comment = c


def _write_block(ctx, ws, columns, header_row, reserve_rows, data_locked,
                 cell_level, prefill=None):
    """ヘッダ行1本＋確保行を書き、確保行の末尾行番号を返す。

    cell_level=True: 1シートに複数ブロックが縦積みされるため、列書式を列単位で
    掛けられない(同じ物理列がブロックごとに別の意味を持つ)。確保行のセルへ個別に
    書式・保護を書く。
    cell_level=False: 1シート1テーブルなので列単位で書式を掛けてよい(安価)。
    どちらの場合もヘッダセルは Locked=True(11章§5の入力対象外セル)。"""
    n_rows = max(int(reserve_rows or 0), 1)
    first, last = header_row + 1, header_row + n_rows
    for i, col in enumerate(columns, start=1):
        letter = get_column_letter(i)
        is_text = col.get("type", "String") == "String"
        cell = ws.cell(row=header_row, column=i, value=col["name"])
        _style_header_cell(cell, col.get("header_note"))

        width = col.get("width")
        if width:
            cur = ws.column_dimensions[letter].width
            # ブロック縦積みでは同じ物理列を複数ブロックが共有するため広い方を採る。
            ws.column_dimensions[letter].width = max(cur or 0, width)

        if cell_level:
            for r in range(first, last + 1):
                dc = ws.cell(row=r, column=i)
                if is_text:
                    dc.number_format = ctx.text_fmt
                dc.protection = LOCKED if data_locked else UNLOCKED
                dc.alignment = Alignment(vertical="top", wrap_text=True)
        else:
            cd = ws.column_dimensions[letter]
            if is_text:
                # 16章 E-46/NFR-S7の二重化: 先頭 = + - @ を格納型数式として
                # 評価させないためテキスト書式へ固定する(実行時の SetCellSafe と二重)。
                cd.number_format = ctx.text_fmt
            cd.protection = LOCKED if data_locked else UNLOCKED
            # 列書式より後に書いたセル書式が勝つ。ヘッダだけは常に Locked。
            cell.protection = LOCKED

        if col.get("dv"):
            _apply_dv(ctx, ws, col["dv"], f"{letter}{first}:{letter}{last}")

    for cname, values in (prefill or {}).items():
        idx = next((i for i, c in enumerate(columns, start=1) if c["name"] == cname), None)
        if idx is None:
            raise BuildError(f"sheets_main.json: prefill の列 '{cname}' が {ws.title} にありません")
        for k, v in enumerate(values):
            if header_row + 1 + k > last:
                raise BuildError(
                    f"sheets_main.json: {ws.title}/{cname} の prefill が reserve_rows を超えています")
            ws.cell(row=header_row + 1 + k, column=idx, value=_clean(v))
    return last


def _make_guard(wb, spec, ctx):
    ws = wb.create_sheet(spec["name"], 0)
    for i, text in enumerate(spec.get("lines") or [], start=1):
        cell = ws.cell(row=i, column=1, value=_clean(text))
        if i == 1:
            cell.font = Font(bold=True, size=16)
        cell.alignment = Alignment(vertical="center")
    ws.column_dimensions["A"].width = 100
    # 13章§2.9: ガードシートは配布ビルドでのみ可視・先頭・**保護なし**。
    ws.protection.sheet = False
    ws.sheet_state = spec.get("state", "visible")
    return ws


def _make_config(wb, spec, ctx):
    ws = wb.create_sheet(spec["name"])
    ws.protection.sheet = False
    headers = spec.get("headers") or ["name", "value", "説明"]
    for c, h in enumerate(headers, 1):
        _style_header_cell(ws.cell(row=1, column=c, value=h), None)
    for i, w in enumerate(spec.get("widths") or [], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    row = 2
    for item in spec.get("defaults") or []:
        value = item.get("value")
        if item["name"] == "mock_llm":
            value = bool(ctx.mock_llm)
        if item["name"] == "app_version":
            value = ctx.app_version
        ws.cell(row=row, column=1, value=item["name"]).protection = LOCKED
        ws.cell(row=row, column=2, value=value).protection = LOCKED
        ws.cell(row=row, column=3, value=_clean(item.get("note", ""))).protection = LOCKED
        row += 1
    ctx.flat_headers[spec["name"]] = list(headers)
    ws.sheet_state = spec.get("state", "hidden")
    return ws


def _make_form(wb, spec, ctx):
    """帳票型(13章§2.9/§2.10/§2.11): 1行目ヘッダを持たず、全ての値を名前付きレンジで
    読み書きする。A=見出し / B=値(名前付きレンジ) / C=注記 の3列で組む。
    見た目を変えても名前付きレンジ名は変えない(セル番地をコードに書かせないため)。"""
    ws = wb.create_sheet(spec["name"])
    ws.protection.sheet = False
    ws.column_dimensions["A"].width = spec.get("label_width", 24)
    ws.column_dimensions["B"].width = spec.get("value_width", 60)
    ws.column_dimensions["C"].width = spec.get("note_width", 50)

    ws.cell(row=1, column=1, value=_clean(spec["name"])).font = SHEET_TITLE_FONT
    row = 3
    for sec in spec.get("sections") or []:
        ws.cell(row=row, column=1, value=_clean(sec.get("title", ""))).font = TITLE_FONT
        row += 1
        tall = bool(sec.get("tall"))
        for fld in sec.get("fields") or []:
            lab = ws.cell(row=row, column=1, value=_clean(fld.get("label", fld["range"])))
            lab.protection = LOCKED
            lab.alignment = Alignment(vertical="top")

            val = ws.cell(row=row, column=2)
            if fld.get("type", "String") == "String":
                val.number_format = ctx.text_fmt
            val.protection = UNLOCKED if fld.get("input") else LOCKED
            val.alignment = Alignment(vertical="top", wrap_text=True)
            _add_name(ctx, ws, fld["range"], 2, row)
            if fld.get("dv"):
                _apply_dv(ctx, ws, fld["dv"], f"B{row}")

            if fld.get("note"):
                nt = ws.cell(row=row, column=3, value=_clean(fld["note"]))
                nt.font = NOTE_FONT
                nt.protection = LOCKED
                nt.alignment = Alignment(vertical="top", wrap_text=True)
            if tall:
                ws.row_dimensions[row].height = 48
            row += 1

        at = sec.get("anchor_table")
        if at:
            for i, h in enumerate(at.get("headers") or [], start=1):
                _style_header_cell(ws.cell(row=row, column=i, value=h), at.get("note"))
            row += 1
            _add_name(ctx, ws, at["range"], 1, row)
            for r in range(row, row + max(int(at.get("reserve_rows") or 1), 1)):
                for i in range(1, len(at.get("headers") or []) + 1):
                    dc = ws.cell(row=r, column=i)
                    dc.number_format = ctx.text_fmt
                    dc.protection = LOCKED
                    dc.alignment = Alignment(vertical="top", wrap_text=True)
            row += max(int(at.get("reserve_rows") or 1), 1)
        row += 1
    ws.sheet_state = spec.get("state", "visible")
    return ws


def _make_table(wb, spec, ctx):
    """テーブル型(13章§2.9)。

    (a) 1シート1テーブル: 1行目=物理名ヘッダ。
    (b) 複数ブロックを縦積み(S1-S4): ブロック名と同名の名前付きレンジ(ヘッダ行の
        左端セル)をアンカーにし、ブロック間に空行を1行以上置く(空行が表の終端)。
    (c) 見出し用の名前付きレンジを持つ2シート(ヒアリングシート hs_ / 壁打ち sp_):
        先頭に見出しブロックを置いてから (b) と同じ構成にする。"""
    ws = wb.create_sheet(spec["name"])
    ws.protection.sheet = False
    data_locked = bool(spec.get("data_locked"))

    if spec.get("sheet_note"):
        c = Comment(_clean(spec["sheet_note"]), "sheets_main.json")
        c.width, c.height = 360, 110
        ws.cell(row=1, column=1).comment = c

    columns = spec.get("columns")
    if columns:
        _write_block(ctx, ws, columns, 1,
                     spec.get("reserve_rows", DEFAULT_FLAT_RESERVE_ROWS),
                     data_locked, cell_level=False)
        ctx.flat_headers[spec["name"]] = [c["name"] for c in columns]
        ws.sheet_state = spec.get("state", "visible")
        return ws

    row = 1
    if spec.get("header_fields"):
        ws.column_dimensions["A"].width = 22
        ws.cell(row=row, column=1,
                value=_clean(spec.get("header_title", spec["name"]))).font = SHEET_TITLE_FONT
        row += 2
        for fld in spec["header_fields"]:
            lab = ws.cell(row=row, column=1, value=_clean(fld.get("label", fld["range"])))
            lab.protection = LOCKED
            val = ws.cell(row=row, column=2)
            if fld.get("type", "String") == "String":
                val.number_format = ctx.text_fmt
            val.protection = UNLOCKED if fld.get("input") else LOCKED
            val.alignment = Alignment(vertical="top", wrap_text=True)
            _add_name(ctx, ws, fld["range"], 2, row)
            if fld.get("note"):
                nt = ws.cell(row=row, column=3, value=_clean(fld["note"]))
                nt.font = NOTE_FONT
                nt.protection = LOCKED
            row += 1
        row += 1

    blocks = spec.get("blocks") or []
    if not blocks:
        raise BuildError(f"sheets_main.json: シート '{spec['name']}' に columns も blocks もありません")
    seen = []
    for blk in blocks:
        ws.cell(row=row, column=1, value=_clean(blk["title"])).font = TITLE_FONT
        row += 1
        header_row = row
        last = _write_block(ctx, ws, blk["columns"], header_row,
                            blk.get("reserve_rows"), data_locked,
                            cell_level=True, prefill=blk.get("prefill"))
        _add_name(ctx, ws, blk["name"], 1, header_row)
        seen.append((blk["name"], [c["name"] for c in blk["columns"]]))
        # ブロックとブロックの間には空行を1行以上置く(13章§2.9・§2.2 逆シリアライズ規約5)
        row = last + 2
    ctx.block_headers[spec["name"]] = seen
    ws.sheet_state = spec.get("state", "visible")
    return ws


def _make_vba_src(wb, spec, ctx):
    """vba_src シート: 標準モジュールのソースを1行1モジュールで格納する。
    A=module_name / B=type / C=source / D=期待行数。
    クラスモジュール(type=class)は VBComponents.Add(1) で追加できないため対象外。"""
    ws = wb.create_sheet(spec["name"])
    ws.protection.sheet = False
    for c, h in enumerate(["module_name", "type", "source", "expected_lines"], 1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)

    injected = []
    row = 2
    for m in _vba_src_modules(ctx.present_modules):
        cleaned = _vba_src_text(ctx.root, m)
        if len(cleaned) > MODULE_CONTRACT_LIMIT:
            raise BuildError(
                f"{m['name']}.bas は{len(cleaned)}字で12章§2のモジュール契約上限"
                f"({MODULE_CONTRACT_LIMIT}字)を超過しています")
        if len(cleaned) >= EXCEL_CELL_LIMIT:
            raise BuildError(
                f"{m['name']}.bas は{len(cleaned)}字でExcelセルの技術上限"
                f"({EXCEL_CELL_LIMIT}字)を超過しています")
        ws.cell(row=row, column=1, value=m["name"])
        ws.cell(row=row, column=2, value="std")
        ws.cell(row=row, column=3, value=cleaned)
        ws.cell(row=row, column=4, value=_expected_line_count(cleaned))
        injected.append(m["name"])
        row += 1

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 80
    ws.column_dimensions["D"].width = 14
    ws.sheet_state = spec.get("state", "veryHidden")
    ctx.injected = injected
    return ws


SHEET_BUILDERS = {
    "guard": _make_guard,
    "config": _make_config,
    "form": _make_form,
    "table": _make_table,
    "vba_src": _make_vba_src,
}


# ---------------------------------------------------------------------------
# ナレッジブック(ナレッジブック.xlsx)ビルド(13章§3・17章 T-03)
# ---------------------------------------------------------------------------
# 本体ブックと違い、ナレッジブックは**マクロ無し.xlsx**(guard/vba_src/config
# シートを持たない。13章§3冒頭・11章§1)であり、全15シートが「1シート1テーブル」
# (13章§3.1-§3.10。ブロック縦積みは無い)なので、本体の BuildCtx/SHEET_BUILDERS
# より軽量な専用の生成路を持つ。ヘッダ書式(_style_header_cell)・入力規則の
# 255字上限判定・段階保存からの自己検証という「型」だけを本体ビルドと共有する。


def load_kb_sheets(path):
    with open(path, encoding="utf-8") as fp:
        data = json.load(fp)
    sheets = data.get("sheets") or []
    if not sheets:
        raise BuildError("sheets_kb.json: sheets が空です")
    names = [s["name"] for s in sheets]
    if len(names) != len(set(names)):
        raise BuildError(f"sheets_kb.json: シート名が重複しています: {names}")
    for s in sheets:
        if not s.get("columns"):
            raise BuildError(
                f"sheets_kb.json: シート'{s['name']}'に columns がありません"
                "(13章§3のナレッジブックは全シートが1シート1テーブル)")
    return data, sheets


class KbBuildCtx:
    """KBシート生成中に持ち回る文脈(本体ブックの BuildCtx のKB版・軽量)。"""

    def __init__(self, data):
        self.enums = data.get("enums") or {}
        self.text_fmt = data.get("text_number_format", "@")
        self.flat_headers = {}   # sheet -> [列名](自己検証がそのまま使う)
        self.seed_counts = {}    # sheet -> 書き込んだ seed_rows 件数
        self.dv_skipped = []


def _kb_apply_dv(ctx, ws, enum_key, target):
    """静的リストの入力規則を張る(本体の _apply_dv と同じ255字上限規約)。"""
    values = ctx.enums.get(enum_key)
    if not values:
        raise BuildError(
            f"sheets_kb.json: 入力規則 '{enum_key}' が enums に定義されていません"
            f"(シート {ws.title})")
    formula = '"' + ",".join(values) + '"'
    if len(formula) > DV_INLINE_LIMIT:
        ctx.dv_skipped.append((ws.title, target, enum_key, len(formula)))
        return False
    dv = DataValidation(type="list", formula1=formula, allow_blank=True,
                        showErrorMessage=True,
                        errorTitle="入力できない値です",
                        error="一覧から選んでください(19章§3のenumレジストリが正)。")
    ws.add_data_validation(dv)
    dv.add(target)
    return True


def _kb_make_sheet(wb, spec, ctx):
    """1シート1テーブルのKBシートを1枚生成する。ヘッダ+seed_rows+確保行。"""
    ws = wb.create_sheet(spec["name"])
    columns = spec["columns"]
    reserve_rows = max(int(spec.get("reserve_rows") or KB_DEFAULT_RESERVE_ROWS), 1)
    seed_rows = spec.get("seed_rows") or []
    if len(seed_rows) > reserve_rows:
        reserve_rows = len(seed_rows)

    for i, col in enumerate(columns, start=1):
        letter = get_column_letter(i)
        cell = ws.cell(row=1, column=i, value=col["name"])
        _style_header_cell(cell, col.get("header_note"))
        width = col.get("width")
        if width:
            ws.column_dimensions[letter].width = width
        if col.get("type", "String") == "String":
            ws.column_dimensions[letter].number_format = ctx.text_fmt
        if col.get("dv"):
            target = f"{letter}2:{letter}{1 + reserve_rows}"
            _kb_apply_dv(ctx, ws, col["dv"], target)

    for r, row in enumerate(seed_rows, start=2):
        if len(row) != len(columns):
            raise BuildError(
                f"sheets_kb.json: シート'{spec['name']}'の seed_rows 行の列数が"
                f"columns定義({len(columns)}列)と不一致です: {row!r}")
        for c, val in enumerate(row, start=1):
            ws.cell(row=r, column=c, value=_clean(val) if isinstance(val, str) else val)

    ws.freeze_panes = "A2"
    ctx.flat_headers[spec["name"]] = [c["name"] for c in columns]
    ctx.seed_counts[spec["name"]] = len(seed_rows)
    return ws


# mock用ナレッジ行(15章§8.1「mock応答に現れるIDは…M-0012/L-03/S-0004/K-0003/
# P9/MC-0107のみ」)。17章T-14が「mock実行前にこれらの行がナレッジ雛形(T-03)へ
# 投入済みであること」を前提にしているため、本ビルドの自己検証で存在を担保する。
KB_MOCK_ROW_IDS = {
    "メニュー一覧": "M-0012",
    "種目マスタ": "L-03",
    "型ライブラリ": "S-0004",
    "成功事例": "K-0003",
    "パターンマスタ": "P9",
    "機構ライブラリ": "MC-0107",
}


def _kb_verify(out_path, sheets, ctx):
    errors = []
    try:
        wb2 = openpyxl.load_workbook(out_path)
    except Exception as e:
        return [f"再オープン失敗: {e}"]

    want_names = [s["name"] for s in sheets]
    if wb2.sheetnames != want_names:
        errors.append(
            f"シート構成が sheets_kb.json と不一致(順序含む): "
            f"期待={want_names} 実際={wb2.sheetnames}")

    for spec in sheets:
        name = spec["name"]
        if name not in wb2.sheetnames:
            continue
        ws = wb2[name]
        cols = ctx.flat_headers[name]
        got = [ws.cell(row=1, column=i).value for i in range(1, len(cols) + 1)]
        if got != cols:
            errors.append(f"'{name}' の1行目ヘッダ不一致: 期待={cols} 実際={got}")
        if ws.cell(row=1, column=len(cols) + 1).value not in (None, ""):
            errors.append(
                f"'{name}' の1行目に台帳外の余分な列があります: "
                f"{ws.cell(row=1, column=len(cols) + 1).value!r}")

    # 17章T-03のDoD「業種マスタ30行存在」。
    if "業種マスタ" in wb2.sheetnames:
        ws = wb2["業種マスタ"]
        n, r = 0, 2
        while ws.cell(row=r, column=1).value not in (None, ""):
            n += 1
            r += 1
        if n < 30:
            errors.append(f"業種マスタの行数が30行未満です(実際{n}行。17章T-03のDoD)")

    # 17章T-03のDoD「mock用ナレッジ行6件存在」(15章§8.1・17章T-14前提)。
    for sheet, mid in KB_MOCK_ROW_IDS.items():
        if sheet not in wb2.sheetnames:
            errors.append(f"mock用ナレッジ行の検査: シート'{sheet}'がありません({mid})")
            continue
        ws = wb2[sheet]
        found, r = False, 2
        while ws.cell(row=r, column=1).value not in (None, ""):
            if str(ws.cell(row=r, column=1).value) == mid:
                found = True
                break
            r += 1
        if not found:
            errors.append(
                f"mock用ナレッジ行 '{mid}' が '{sheet}' に見つかりません"
                "(15章§8.1・17章T-14前提)")

    return errors


def build_kb(kb_json_path, out_path):
    """ナレッジブック.xlsx を生成する(13章§3・17章 T-03)。マクロ無し.xlsx。"""
    print("=== build_rpn.py --kb ===")
    print(f"kb:     {kb_json_path}")
    print(f"out:    {out_path}")

    if not os.path.exists(kb_json_path):
        sys.exit(f"ERROR: sheets_kb.json が見つかりません: {kb_json_path}")
    try:
        data, sheets = load_kb_sheets(kb_json_path)
    except BuildError as e:
        sys.exit(f"ERROR: {e}")

    print(f"kb_version (sheets_kb.json): {data.get('kb_version', '0.0.0')}")
    print(f"\nStage 1: ブック生成(全{len(sheets)}シート、マクロ無し.xlsx)...")

    wb = openpyxl.Workbook()
    for name in list(wb.sheetnames):
        del wb[name]

    ctx = KbBuildCtx(data)
    try:
        for spec in sheets:
            _kb_make_sheet(wb, spec, ctx)
    except BuildError as e:
        sys.exit(f"ERROR: {e}")

    if ctx.dv_skipped:
        print(f"  入力規則を省略した列({len(ctx.dv_skipped)}件・"
              f"Excelのリスト上限{DV_INLINE_LIMIT}字超):")
        for sh, tgt, key, ln in ctx.dv_skipped:
            print(f"    {sh}!{tgt} enums.{key} ({ln}字)")

    seed_total = sum(ctx.seed_counts.values())
    print(f"  シート最終構成({len(wb.sheetnames)}件): {wb.sheetnames}")
    print(f"  seed_rows合計: {seed_total}行 "
          f"(業種マスタ={ctx.seed_counts.get('業種マスタ', 0)}行)")
    if wb.sheetnames != [s["name"] for s in sheets]:
        sys.exit(f"ERROR: シート構成が sheets_kb.json と不一致: {wb.sheetnames}")

    wb.active = 0

    print("Stage 2: 出力(一時パスへ)...")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    _out_base, _out_ext = os.path.splitext(out_path)
    staging_path = f"{_out_base}.building{_out_ext}"
    failed_path = f"{_out_base}.failed{_out_ext}"
    wb.save(staging_path)
    print(f"  一時出力: {staging_path} ({os.path.getsize(staging_path):,} bytes)")

    print("\nStage 3: ビルド後自己検証...")
    errors = _kb_verify(staging_path, sheets, ctx)
    if errors:
        print("自己検証 失敗:")
        for e in errors:
            print(f"  - {e}")
        try:
            if os.path.exists(failed_path):
                os.remove(failed_path)
            os.replace(staging_path, failed_path)
            print(f"  不良な出力を退避しました(前回の良品は無傷): {failed_path}")
        except OSError as e2:
            print(f"  不良な出力の退避にも失敗しました: {e2}")
        sys.exit(1)

    os.replace(staging_path, out_path)
    if os.path.exists(failed_path):
        try:
            os.remove(failed_path)
        except OSError:
            pass
    print(f"  出力: {out_path} ({os.path.getsize(out_path):,} bytes)")
    print("自己検証 OK: シート集合と順序・1行目ヘッダ・業種マスタ30行以上・"
          "mock用ナレッジ行6件(M-0012/L-03/S-0004/K-0003/P9/MC-0107)")
    print("  ※13章§3との突合(シート名・列名・順序)は tools/sheet_check.py --kb が行う。")
    print("\nDone.")


# ---------------------------------------------------------------------------
# ビルド後自己検証
# ---------------------------------------------------------------------------
def _first_diff_pos(a, b):
    n = min(len(a), len(b))
    for i in range(n):
        if a[i] != b[i]:
            return i
    return n


def _verify_vba_src_bodies(ws, got_names, present_modules, root):
    """成果物 vba_src のC列本文が src/ の対応ファイル(ビルド規則適用後)と完全一致
    することを全モジュールで検査する。あわせてD列(期待行数)も全数突合する
    (名前は揃っているのに中身が欠けている配布物を出荷段階で止めるため)。"""
    errors = []
    expected = {}
    for m in _vba_src_modules(present_modules):
        try:
            expected[m["name"]] = _vba_src_text(root, m)
        except OSError as e:
            errors.append(f"vba_src本文検査: '{m['name']}' のソースを読めません: {e}")

    seen = set()
    for i, nm in enumerate(got_names):
        actual = ws.cell(row=i + 2, column=3).value or ""
        seen.add(nm)
        if nm not in expected:
            errors.append(f"vba_src本文検査: '{nm}' に対応する src/ のモジュールがありません")
            continue
        want = expected[nm]
        if actual != want:
            pos = _first_diff_pos(want, actual)
            errors.append(
                f"vba_src本文が src/ と不一致: '{nm}' "
                f"(先頭差分位置={pos}, 期待{len(want)}字/実際{len(actual)}字, "
                f"期待={want[pos:pos + 40]!r} 実際={actual[pos:pos + 40]!r})")

        want_lines = _expected_line_count(want)
        got_d = ws.cell(row=i + 2, column=4).value
        got_d_num = (int(got_d) if isinstance(got_d, (int, float))
                     and not isinstance(got_d, bool) else None)
        if got_d_num != want_lines:
            errors.append(
                f"vba_src D列(期待行数)が src由来の計算値と不一致: '{nm}' "
                f"D列={got_d!r} 期待={want_lines}")

    for nm in expected:
        if nm not in seen:
            errors.append(f"vba_src本文検査: '{nm}' の行が成果物にありません")
    return errors


def _verify_layout(wb2, ctx):
    """生成時に溜めた期待値(名前付きレンジ・ヘッダ行)を成果物側で読み直して突合する。
    ここで見るのは「ブックの中身が台帳どおりか」だけで、台帳が13章どおりかは
    tools/sheet_check.py(13章のMarkdownを直接パースする)の担当。"""
    errors = []

    # 名前付きレンジ: 全て定義され、宣言どおりのシート・セルを指すこと。
    try:
        defined = dict(wb2.defined_names)
    except Exception as e:                                   # pragma: no cover
        return [f"名前付きレンジを読めません: {e}"]
    for name, sheet, coord in ctx.names:
        dn = defined.get(name)
        if dn is None:
            errors.append(f"名前付きレンジ '{name}' が成果物にありません(期待 {sheet}!{coord})")
            continue
        dests = list(dn.destinations)
        if len(dests) != 1:
            errors.append(f"名前付きレンジ '{name}' の参照先が1件ではありません: {dests}")
            continue
        got_sheet, got_coord = dests[0]
        if got_sheet != sheet or got_coord.replace("$", "") != coord:
            errors.append(
                f"名前付きレンジ '{name}' の参照先不一致: "
                f"期待={sheet}!{coord} 実際={got_sheet}!{got_coord}")
    extra = sorted(set(defined) - {n for n, _, _ in ctx.names})
    if extra:
        errors.append(f"台帳にない名前付きレンジが成果物にあります: {extra}")

    # 1行目=物理名ヘッダのシート。
    for sheet, cols in ctx.flat_headers.items():
        if sheet not in wb2.sheetnames:
            continue
        ws = wb2[sheet]
        got = [ws.cell(row=1, column=i).value for i in range(1, len(cols) + 1)]
        if got != cols:
            errors.append(f"'{sheet}' の1行目ヘッダ不一致: 期待={cols} 実際={got}")
        if ws.cell(row=1, column=len(cols) + 1).value not in (None, ""):
            errors.append(
                f"'{sheet}' の1行目に台帳外の余分な列があります: "
                f"{ws.cell(row=1, column=len(cols) + 1).value!r}")

    # ブロック縦積みのシート: アンカーの名前付きレンジからヘッダ行を特定して突合する
    # (19章§5の照合手順と同じ道筋を、ビルド直後にもう一度通す)。
    for sheet, blocks in ctx.block_headers.items():
        if sheet not in wb2.sheetnames:
            continue
        ws = wb2[sheet]
        for bname, cols in blocks:
            dn = defined.get(bname)
            if dn is None:
                errors.append(f"ブロックアンカー '{bname}' が成果物にありません({sheet})")
                continue
            got_sheet, got_coord = list(dn.destinations)[0]
            if got_sheet != sheet:
                errors.append(f"ブロックアンカー '{bname}' が別シートを指しています: {got_sheet}")
                continue
            hrow = int(re.sub(r"[^0-9]", "", got_coord))
            got = [ws.cell(row=hrow, column=i).value for i in range(1, len(cols) + 1)]
            if got != cols:
                errors.append(
                    f"'{sheet}'/{bname} のヘッダ行({hrow}行目)不一致: 期待={cols} 実際={got}")
            nxt = ws.cell(row=hrow, column=len(cols) + 1).value
            if nxt not in (None, ""):
                errors.append(f"'{sheet}'/{bname} のヘッダ行に台帳外の列があります: {nxt!r}")
    return errors


def _verify_installer_patch(vba_bin, installer_src):
    """成果物の vbaProject.bin に自己インストーラ外科パッチが正しく当たっているかを
    別実装(ovba.CFBReader)で読み戻して検証する3項目(PoC由来):
    (1) ThisWorkbookストリームの復元確認(先頭一致+末尾NULパディングのみ)
    (2) dir ストリームの ThisWorkbook.MOFFSET が 0
    (3) _VBA_PROJECT の無害化(サイズ不変・Version=0xFFFF・PerformanceCacheゼロ埋め)"""
    errors = []
    try:
        cfb = ovba.CFBReader(vba_bin)

        # (1) ThisWorkbook復元確認。空チャンクpaddingは解凍すると末尾に数バイトの
        #     NULが付くため「先頭が完全一致し、余剰があるならNULのみ」で確認する。
        tw_dec = ovba.ovba_decompress(cfb.read("ThisWorkbook"))
        tail = tw_dec[len(installer_src):]
        if not tw_dec.startswith(installer_src) or any(b != 0 for b in tail):
            errors.append(
                "ThisWorkbookストリームの復元結果が自己インストーラソースと不一致"
                "(先頭一致+末尾NULパディングという想定パターンから外れています)")

        # (2) dir MOFFSET=0確認。
        dir_dec = ovba.ovba_decompress(cfb.read("dir"))
        needle = struct.pack("<HI", 0x0019, len("ThisWorkbook")) + b"ThisWorkbook"
        idx = dir_dec.find(needle)
        moffset_ok = False
        if idx >= 0:
            i = idx
            while i < len(dir_dec):
                rid = struct.unpack("<H", dir_dec[i:i + 2])[0]
                sz = struct.unpack("<I", dir_dec[i + 2:i + 6])[0]
                if rid == 0x0031:
                    moffset_ok = (struct.unpack("<I", dir_dec[i + 6:i + 10])[0] == 0)
                    break
                i += 6 + sz
        if not moffset_ok:
            errors.append("dirストリームのThisWorkbook.MOFFSETが0になっていません")

        # (3) _VBA_PROJECT無害化確認。
        vp = cfb.read("_VBA_PROJECT")
        if len(vp) != _VBA_PROJECT_STREAM_SIZE:
            errors.append(
                f"_VBA_PROJECTストリームのサイズが変化しています: "
                f"期待={_VBA_PROJECT_STREAM_SIZE}バイト 実際={len(vp)}バイト")
        else:
            ver = struct.unpack("<H", vp[2:4])[0]
            if ver != _VBA_PROJECT_VERSION_IGNORE:
                errors.append(
                    f"_VBA_PROJECTのVersionが0x{_VBA_PROJECT_VERSION_IGNORE:04X}では"
                    f"ありません(実際=0x{ver:04X})。開き手のOfficeビルドと一致すると"
                    "PerformanceCacheがソースより優先され幽霊コンパイルエラーの原因になります")
            nonzero = sum(1 for b in vp[_VBA_PROJECT_CACHE_OFFSET:] if b != 0)
            if nonzero:
                errors.append(
                    f"_VBA_PROJECTのPerformanceCacheがゼロ埋めされていません(非ゼロ={nonzero}バイト)")
    except Exception as e:                                          # pragma: no cover
        errors.append(f"自己インストーラ外科パッチの検証中に例外: {e}")
    return errors


def verify_build(out_path, expected_vba_src_names, sheets, mock_llm_expected,
                 app_version, present_modules, root, has_vba_project, ctx=None,
                 installer_src=None):
    errors = []
    try:
        wb2 = openpyxl.load_workbook(out_path, keep_vba=True)
    except Exception as e:
        return [f"再オープン失敗: {e}"]

    want_names = [s["name"] for s in sheets]
    if wb2.sheetnames != want_names:
        errors.append(
            f"シート構成が sheets_main.json と不一致(順序含む): "
            f"期待={want_names} 実際={wb2.sheetnames}")

    guard_name = want_names[0]
    if wb2.sheetnames and wb2.sheetnames[0] != guard_name:
        errors.append(
            f"'{guard_name}' が先頭シートになっていません: 実際={wb2.sheetnames[0]!r}")
    active_title = wb2.active.title if wb2.active is not None else None
    if active_title != guard_name:
        errors.append(
            f"アクティブシートが'{guard_name}'ではありません: 実際={active_title!r}")

    for spec in sheets:
        name = spec["name"]
        if name not in wb2.sheetnames:
            continue
        want_state = spec.get("state", "visible")
        actual = wb2[name].sheet_state
        if actual != want_state:
            errors.append(f"シート'{name}'の可視性不一致: 期待={want_state} 実際={actual}")

    vba_spec = next((s for s in sheets if s.get("role") == "vba_src"), None)
    if vba_spec and vba_spec["name"] in wb2.sheetnames:
        ws = wb2[vba_spec["name"]]
        got_names = []
        r = 2
        while ws.cell(row=r, column=1).value:
            nm = ws.cell(row=r, column=1).value
            src = ws.cell(row=r, column=3).value or ""
            got_names.append(nm)
            if len(src) > EXCEL_CELL_LIMIT:
                errors.append(
                    f"vba_src '{nm}' のソースが{len(src)}字でExcelセル上限"
                    f"({EXCEL_CELL_LIMIT})を超過")
            if len(src) > MODULE_CONTRACT_LIMIT:
                errors.append(
                    f"vba_src '{nm}' のソースが{len(src)}字で30,000字契約を超過")
            r += 1
        if sorted(got_names) != sorted(expected_vba_src_names):
            errors.append(
                f"vba_srcモジュール集合が不一致: 期待={sorted(expected_vba_src_names)} "
                f"実際={sorted(got_names)}")
        errors.extend(_verify_vba_src_bodies(ws, got_names, present_modules, root))
    else:
        errors.append("vba_src シートが存在しない")

    cfg_spec = next((s for s in sheets if s.get("role") == "config"), None)
    if cfg_spec and cfg_spec["name"] in wb2.sheetnames:
        ws = wb2[cfg_spec["name"]]
        r = 2
        got_values, got_keys = {}, []
        while ws.cell(row=r, column=1).value:
            key = ws.cell(row=r, column=1).value
            got_keys.append(key)
            got_values[key] = ws.cell(row=r, column=2).value
            if not (ws.cell(row=r, column=3).value or "").strip():
                errors.append(f"config '{key}' の説明列が空です(13章§2.3は全キーに説明を持つ)")
            r += 1
        want_keys = [d["name"] for d in cfg_spec.get("defaults") or []]
        if got_keys != want_keys:
            errors.append(
                f"config のキー列が台帳と不一致(順序含む): 期待={want_keys} 実際={got_keys}")
        if "mock_llm" in got_values and bool(got_values["mock_llm"]) != mock_llm_expected:
            errors.append(
                f"config!mock_llm が期待値と不一致: 期待={mock_llm_expected} "
                f"実際={got_values['mock_llm']}")
        if "app_version" in got_values and got_values["app_version"] != app_version:
            errors.append(
                f"config!app_version が台帳と不一致: 期待={app_version} "
                f"実際={got_values['app_version']}")

    if ctx is not None:
        errors.extend(_verify_layout(wb2, ctx))

    # 成果物のパッケージ整合(拡張子と中身の一致 / vbaProject.bin の持ち込み)。
    try:
        with zipfile.ZipFile(out_path) as z:
            names = set(z.namelist())
            ct = z.read("[Content_Types].xml").decode("utf-8", errors="replace")
            if WORKBOOK_CT_XLSM not in ct:
                errors.append(
                    "[Content_Types].xml のワークブックパートが macroEnabled では"
                    "ありません(.xlsm 拡張子と中身が不整合になります)")
            if has_vba_project:
                if "xl/vbaProject.bin" not in names:
                    errors.append("テンプレート由来の xl/vbaProject.bin が成果物にありません")
                else:
                    vba_bin = z.read("xl/vbaProject.bin")
                    if olefile is not None:
                        ole = olefile.OleFileIO(io.BytesIO(vba_bin))
                        if not ole.exists("VBA/dir") or not ole.exists("VBA/ThisWorkbook"):
                            errors.append(
                                "vbaProject.bin に VBA/dir または VBA/ThisWorkbook "
                                "ストリームがありません")
                        ole.close()
                    # 自己インストーラ外科パッチの読み戻し検証(3項目。PoC由来)。
                    if installer_src is not None:
                        errors.extend(_verify_installer_patch(vba_bin, installer_src))
            elif "xl/vbaProject.bin" in names:
                errors.append(
                    "テンプレート無しのビルドなのに xl/vbaProject.bin が混入しています")
    except Exception as e:
        errors.append(f"パッケージ検証中に例外: {e}")

    return errors


# ---------------------------------------------------------------------------
# メイン
# ---------------------------------------------------------------------------
_BUILD_LEFTOVER_SUFFIXES = (
    ".building.xlsm", ".failed.xlsm",   # 本体ブック(--dev/--prod)
    ".building.xlsx", ".failed.xlsx",   # ナレッジブック(--kb)
)


def _sweep_build_leftovers(dist_dir: str) -> None:
    """dist/ に残った *.building.xlsm(x) / *.failed.xlsm(x) を消す。
    消せなくてもビルドは続ける(掃除の失敗で配布物を作れなくしない)。"""
    if not os.path.isdir(dist_dir):
        return
    for name in os.listdir(dist_dir):
        if name.endswith(_BUILD_LEFTOVER_SUFFIXES):
            path = os.path.join(dist_dir, name)
            try:
                os.remove(path)
                print(f"  前回の中間生成物を削除: {name}")
            except OSError as e:
                print(f"  注意: 中間生成物を削除できませんでした({name}: {e})")


def _patch_content_types(parts: dict) -> None:
    """[Content_Types].xml のワークブックパートを macroEnabled 用へ書き換える。
    テンプレート(本物の.xlsm)を使わずに openpyxl だけで組み立てた場合、
    ワークブックの content type が xlsx 用のままになり、Excel が
    「拡張子が実際の形式と一致しません」と警告する。"""
    key = "[Content_Types].xml"
    if key not in parts:
        return
    xml = parts[key].decode("utf-8")
    if WORKBOOK_CT_XLSM in xml:
        return
    parts[key] = xml.replace(WORKBOOK_CT_XLSX, WORKBOOK_CT_XLSM).encode("utf-8")


def main():
    ap = argparse.ArgumentParser(description=f"{APP_TITLE} ビルドスクリプト")
    mode = ap.add_mutually_exclusive_group(required=True)
    mode.add_argument("--dev", action="store_true",
                      help="開発ビルド(config mock_llm=TRUE)。既定出力: リスク提案ナビ_dev.xlsm")
    mode.add_argument("--prod", action="store_true",
                      help="本番ビルド(config mock_llm=FALSE)。既定出力: リスク提案ナビ.xlsm")
    mode.add_argument("--kb", action="store_true",
                      help="ナレッジブック(マクロ無し.xlsx)をビルドする(13章§3・17章T-03)。"
                           "既定出力: <root>/dist/ナレッジブック.xlsx。"
                           "--dev/--prod系の引数(--modules/--template/--allow-missing)は無関係")
    ap.add_argument("--root", default=DEFAULT_ROOT,
                    help="modules.json内パスの解決基準ディレクトリ(既定: <repo>)")
    ap.add_argument("--modules", default=DEFAULT_MODULES_JSON, help="modules.json のパス")
    ap.add_argument("--sheets", default=DEFAULT_SHEETS_JSON, help="sheets_main.json のパス")
    ap.add_argument("--sheets-kb", default=DEFAULT_SHEETS_KB_JSON,
                    help="sheets_kb.json のパス(--kb 用)")
    ap.add_argument("--template", default=DEFAULT_TEMPLATE,
                    help="template_skeleton.xlsm のパス(存在すれば vbaProject.bin を引き継ぐ)")
    ap.add_argument("--out", default=None,
                    help="出力先パス(既定: <root>/dist/リスク提案ナビ[_dev].xlsm。"
                         "--kb 指定時は <root>/dist/ナレッジブック.xlsx)")
    ap.add_argument("--allow-missing", action="store_true",
                    help="modules.jsonに列挙されたファイルの欠落をエラーでなく警告にする")
    args = ap.parse_args()

    root = os.path.abspath(args.root)

    if args.kb:
        kb_json_path = os.path.abspath(args.sheets_kb)
        kb_out_path = os.path.abspath(args.out) if args.out else os.path.join(
            root, "dist", "ナレッジブック.xlsx")
        _sweep_build_leftovers(os.path.join(root, "dist"))
        build_kb(kb_json_path, kb_out_path)
        return

    is_dev = bool(args.dev)
    mock_llm = is_dev

    out_path = os.path.abspath(args.out) if args.out else os.path.join(
        root, "dist", "リスク提案ナビ_dev.xlsm" if is_dev else "リスク提案ナビ.xlsm")

    _sweep_build_leftovers(os.path.join(root, "dist"))

    print(f"=== build_rpn.py ({'dev' if is_dev else 'prod'}) ===")
    print(f"root:     {root}")
    print(f"modules:  {args.modules}")
    print(f"sheets:   {args.sheets}")
    print(f"out:      {out_path}")

    if not os.path.exists(args.modules):
        sys.exit(f"ERROR: modules.json が見つかりません: {args.modules}")
    if not os.path.exists(args.sheets):
        sys.exit(f"ERROR: sheets_main.json が見つかりません: {args.sheets}")

    try:
        modules = load_manifest(args.modules)
        sheets_data, sheets = load_sheets(args.sheets)
    except BuildError as e:
        sys.exit(f"ERROR: {e}")

    app_version = sheets_data.get("app_version", "0.0.0")
    print(f"app_version (sheets_main.json): {app_version}")
    print()

    present, missing = validate_modules(modules, root, args.allow_missing)
    print(f"モジュール台帳: 全{len(modules)}件 / 実在{len(present)}件 / 欠落{len(missing)}件")
    for role in VALID_ROLES:
        n = sum(1 for m in present if m["role"] == role)
        if n:
            print(f"  role={role}: {n}件")

    has_template = os.path.exists(args.template)
    print(f"\nStage 1: ブック生成 (template={'あり' if has_template else 'なし'})...")
    if has_template:
        wb = openpyxl.load_workbook(args.template, keep_vba=True)
        for name in list(wb.sheetnames):
            del wb[name]
        print(f"  テンプレートの vbaProject.bin を引き継ぎます: {args.template}")
        print("  自己インストーラ(ThisWorkbookストリームの外科パッチ)は Stage 4 で当てます"
              "(build/ovba.py 経由)。")
    else:
        wb = openpyxl.Workbook()
        for name in list(wb.sheetnames):
            del wb[name]
        print("  build/template_skeleton.xlsm が無いため、VBAプロジェクトを持たない"
              ".xlsm として生成します(シート雛形の検査は成立します)。")

    print(f"Stage 2: シート生成 (sheets_main.json 全{len(sheets)}シート)...")
    ctx = BuildCtx(wb, sheets_data, mock_llm, app_version, present, root)
    try:
        for spec in sheets:
            role = spec.get("role")
            builder = SHEET_BUILDERS.get(role)
            if builder is None:
                raise BuildError(f"sheets_main.json: 未知の role '{role}'(シート {spec['name']})")
            builder(wb, spec, ctx)
    except BuildError as e:
        sys.exit(f"ERROR: {e}")

    injected = ctx.injected
    n_blocks = sum(len(v) for v in ctx.block_headers.values())
    print(f"  名前付きレンジ: {len(ctx.names)}本"
          f"(ブロックアンカー{n_blocks}本を含む)")
    print(f"  1行目=物理名ヘッダのシート: {len(ctx.flat_headers)}枚 / "
          f"ブロック縦積みのシート: {len(ctx.block_headers)}枚")
    if ctx.dv_skipped:
        print(f"  入力規則を省略した列({len(ctx.dv_skipped)}件・"
              f"Excelのリスト上限{DV_INLINE_LIMIT}字超。最終形はW3の隠しレンジ方式):")
        for sh, tgt, key, ln in ctx.dv_skipped:
            print(f"    {sh}!{tgt} enums.{key} ({ln}字)")
    print(f"  vba_src: {len(injected)}モジュールを格納 ({injected})")
    print(f"  シート最終構成({len(wb.sheetnames)}件): {wb.sheetnames}")
    if wb.sheetnames != [s["name"] for s in sheets]:
        sys.exit(f"ERROR: シート構成が sheets_main.json と不一致: {wb.sheetnames}")

    print("Stage 3: 一時保存...")
    wb.active = 0
    with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as tmp:
        tmp_path = tmp.name
    wb.save(tmp_path)
    print(f"  一時保存: {tmp_path} ({os.path.getsize(tmp_path):,} bytes)")

    print("Stage 4: パッケージ整形(.xlsm の content type / 自己インストーラ注入)...")
    with zipfile.ZipFile(tmp_path) as zin:
        parts = {n: zin.read(n) for n in zin.namelist()}
    _patch_content_types(parts)
    has_vba_project = "xl/vbaProject.bin" in parts
    print(f"  vbaProject.bin: {'あり' if has_vba_project else 'なし'}")

    installer_src = None
    if has_vba_project:
        # テンプレート由来の本物の vbaProject.bin に、ThisWorkbook自己インストーラを
        # 外科パッチする(dir MOFFSET=0・_VBA_PROJECT無害化を含む)。バイト長は不変。
        try:
            installer_src = build_installer_src()
            skel_bin = parts["xl/vbaProject.bin"]
            patched_bin = patch_installer(skel_bin, installer_src)
        except BuildError as e:
            sys.exit(f"ERROR: {e}")
        if len(skel_bin) != len(patched_bin):
            sys.exit("ERROR: vbaProject.bin のバイト長が外科パッチで変化しました"
                     "(バイナリ整合性エラー)")
        parts["xl/vbaProject.bin"] = patched_bin
        print(f"  自己インストーラ注入: ThisWorkbook差替 + dir MOFFSET=0 + "
              f"_VBA_PROJECT無害化(vbaProject.bin {len(skel_bin):,} bytes 不変)")

    print("Stage 5: 最終.xlsm書き出し(一時パスへ)...")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    # 拡張子を .xlsm のまま保つ(自己検証内の openpyxl.load_workbook が拡張子で
    # フォーマットを判定するため、".building" のような接尾辞を単純に足すと
    # 「サポートしていない形式」として自己検証自体が失敗する)。
    _out_base, _out_ext = os.path.splitext(out_path)
    staging_path = f"{_out_base}.building{_out_ext}"
    failed_path = f"{_out_base}.failed{_out_ext}"
    with zipfile.ZipFile(staging_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for n, data in parts.items():
            zout.writestr(n, data)
    os.unlink(tmp_path)
    print(f"  一時出力: {staging_path} ({os.path.getsize(staging_path):,} bytes)")

    print("\nStage 6: ビルド後自己検証...")
    errors = verify_build(staging_path, injected, sheets, mock_llm, app_version,
                          present, root, has_vba_project, ctx, installer_src)
    if errors:
        print("自己検証 失敗:")
        for e in errors:
            print(f"  - {e}")
        try:
            if os.path.exists(failed_path):
                os.remove(failed_path)
            os.replace(staging_path, failed_path)
            print(f"  不良な出力を退避しました(前回の良品は無傷): {failed_path}")
        except OSError as e2:
            print(f"  不良な出力の退避にも失敗しました: {e2}")
        sys.exit(1)

    # 検証OKになって初めて正規パスへ確定する(os.replace は同一ファイルシステム
    # 内であれば原子的)。ここまで前回の良品は無傷のまま。
    os.replace(staging_path, out_path)
    if os.path.exists(failed_path):
        try:
            os.remove(failed_path)
        except OSError:
            pass
    print(f"  出力: {out_path} ({os.path.getsize(out_path):,} bytes)")
    print("自己検証 OK: シート集合と順序・可視性 / ガードシートが先頭かつアクティブ / "
          "vba_srcモジュール集合一致 / 各ソース<=30,000字かつ<32,000字 / "
          "vba_src本文がsrc/と完全一致 / vba_src D列(期待行数)がsrc由来の計算値と一致 / "
          "configキー列(順序含む)と全キーの説明・mock_llm・app_version / "
          "名前付きレンジの本数と参照先 / 1行目ヘッダとブロックアンカー先ヘッダ行 / "
          "パッケージ content type")
    if has_vba_project:
        print("  自己インストーラ外科パッチ検証(3項目): "
              "ThisWorkbook復元確認 / dir MOFFSET=0 / "
              "_VBA_PROJECT無害化(Version=0xFFFF・PerformanceCacheゼロ埋め・3,061B不変)")
    print("  ※13章との突合(シート名・列名・順序・configキー・名前付きレンジ)は "
          "tools/sheet_check.py が行う。")
    print("\nDone.")


if __name__ == "__main__":
    main()
