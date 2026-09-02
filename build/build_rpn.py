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
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.worksheet.hyperlink import Hyperlink
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
# 1物理行のCP932バイト長上限(W5.3 H7 / 裁定書19)。VBEの1物理行上限1,023は
# 文字数ではなくCP932エンコード後のバイト数で数える。超えるとVBEが注入時に行を
# 分断し、実機で「SubまたはFunctionが定義されていません」になる。tools/vba_lint.py
# の MAX_LINE_CP932_BYTES と同値の二重化検問(裁定書19 H9の代替措置)。
MAX_LINE_CP932_BYTES = 1000     # 焼き込み前に超過を BuildError で止める

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
# 11章§3.3.3: 6欄の例文は「薄い灰色 RGB(150,150,150)」で本文と見分けさせる。
GRAY_FONT = Font(size=9, color="969696")
# 区画の中の小見出し(「── 1 調べた結果 ──」等)。区画見出し(TITLE_FONT)より弱く、
# 注記(NOTE_FONT)より強い1段を作る。
SUBHEAD_FONT = Font(bold=True, size=10)
SHEET_TITLE_FONT = Font(bold=True, size=14)
LOCKED = Protection(locked=True)
UNLOCKED = Protection(locked=False)

# ---------------------------------------------------------------------------
# ビルド時装飾の配色(裁定書14 裁定7)
# ---------------------------------------------------------------------------
# 白抜き文字を置く面は WCAG 2.x のコントラスト比 4.5:1 以上であることを
# _assert_contrast() がビルド時に機械検算する(目視で「読めるつもり」の配色を
# 焼き込まないため)。テーブル型ヘッダの HEADER_FILL(黒字)は現状維持。
BRAND_DARK = "014D44"        # RGB(1,77,68)  シート1行目のタイトル帯
BRAND_MID = "0B7D6E"         # 操作ガイドの章内見出し・手順番号バッジ
SECTION_TINT = "E7F3EE"      # 淡緑。セクション見出し(TITLE_FONT行)の地
WHITE = "FFFFFF"
CONTRAST_MIN = 4.5

TITLE_BAND_FILL = PatternFill("solid", fgColor=BRAND_DARK)
TITLE_BAND_FONT = Font(bold=True, size=14, color=WHITE)
SECTION_FILL = PatternFill("solid", fgColor=SECTION_TINT)
TITLE_BAND_COLS = 3          # A:C を帯にする(帳票型の A=見出し / B=値 / C=注記)

# タイトル帯の**文字**を置く列(既定=A列)。帯の塗り(A:C)は変えない。
# 裁定書16 F3: S1-S4は1行目のボタンが左端(A・B列)に生えるため、A1へ表題を書くと
# 文字がボタンの下に隠れて読めない。この4枚だけ表題をC1へ寄せる(帯は A1:C1 のまま)。
TITLE_BAND_TEXT_COL = {
    "S1_企業プロファイル": 3,
    "S2_リスク仮説": 3,
    "S3_提案": 3,
    "S4_骨子": 3,
}


def _rel_luminance(hex_rgb):
    """WCAG 2.x の相対輝度(0.0-1.0)。modSkin.bas のコントラスト検算と同じ式。"""
    v = []
    for i in (0, 2, 4):
        c = int(hex_rgb[i:i + 2], 16) / 255.0
        v.append(c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4)
    return 0.2126 * v[0] + 0.7152 * v[1] + 0.0722 * v[2]


def _contrast_ratio(fg_hex, bg_hex):
    a, b = _rel_luminance(fg_hex), _rel_luminance(bg_hex)
    hi, lo = max(a, b), min(a, b)
    return (hi + 0.05) / (lo + 0.05)


def _assert_contrast(pairs):
    """(前景, 背景, 用途) の全組で 4.5:1 以上を機械検算する。落ちたらビルド中止。"""
    bad = []
    for fg, bg, where in pairs:
        r = _contrast_ratio(fg, bg)
        if r < CONTRAST_MIN:
            bad.append(f"{where}: #{fg} on #{bg} = {r:.2f}:1 (< {CONTRAST_MIN}:1)")
    if bad:
        raise BuildError(
            "配色のコントラスト比が基準(4.5:1)に届きません(裁定書14 裁定7):\n  "
            + "\n  ".join(bad))
    return [(fg, bg, where, _contrast_ratio(fg, bg)) for fg, bg, where in pairs]

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
#   ・焼き付け(bake。裁定書23 C-1): 注入が全件成功した回の Save 直前に
#     vba_src!E3 へ "baked" を書く。次回以降の起動は Install 冒頭で E3 を読み、
#     "baked" なら VBProject に一切触らずそのまま OnTime(失敗時 Application.Run)
#     で modBoot.Boot へ渡る。VBProject に触らないので利用者に「VBAプロジェクト
#     オブジェクト モデルへのアクセスを信頼」(VBOM)を求めない。焼き付けは
#     VBOM有効の開発PC(Windows)で1回開いて閉じるだけで済む。
#     baked なのに Application.Run が失敗した(=モジュールが焼き付いていない)
#     ときだけ "Setup NG: modules missing. Ask the developer." の1文で止める。
#     E3 が空なら従来どおりの注入経路(VBOM必須)。ビルドは E3 を書かない
#     (成果物の E3 が空であることは verify_build が確認する)。
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
  Dim b As Boolean
  On Error GoTo Done
  Set w = ThisWorkbook.Worksheets("vba_src")
  If CStr(w.Cells(3, 5).Value) = "baked" Then
    b = True
    GoTo Boot
  End If
  On Error GoTo Trust
  Set p = ThisWorkbook.VBProject
  On Error GoTo Done
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
  w.Cells(3, 5).Value = "baked"
  ThisWorkbook.Save
Boot:
  On Error Resume Next
  Err.Clear
  Dim bt As Date
  bt = Now + TimeSerial(0, 0, 1)
  Application.OnTime bt, "'" & ThisWorkbook.Name & "'!modBoot.Boot"
  If Err.Number <> 0 Then
    Err.Clear
    Application.Run "modBoot.Boot"
    If Err.Number <> 0 And b Then
      MsgBox "Setup NG: modules missing. Ask the developer.", vbCritical
    End If
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

    # 1物理行のCP932バイト長検問(裁定書19 H7(4)/H9代替)。vba_src へ焼く前に止める。
    # tools/vba_lint.py と同じ基準をビルド側にも置いて二重化する(lint を通さずに
    # ビルドだけ回しても、行分断で壊れる配布物は出荷できない)。
    for i, line in enumerate(cleaned.split("\n"), 1):
        nbytes = len(line.encode("cp932", errors="replace"))
        if nbytes > MAX_LINE_CP932_BYTES:
            raise BuildError(
                f"{m['name']}.bas:{i} の1物理行が{nbytes}バイト(CP932)で上限"
                f"{MAX_LINE_CP932_BYTES}バイトを超過しています(文字数は{len(line)}字)。"
                "VBEの1行上限1,023は文字数ではなくバイト数です。VBEが行を分断し、"
                "実機で「SubまたはFunctionが定義されていません」になります。"
                "文字列リテラルの境界で複数行へ分けてください")
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
        self.tests_expected = read_tests_expected(root)


def _add_name(ctx, ws, name, col, row):
    """ブック内で一意の名前付きレンジを1セルに定義する(13章§2.9)。"""
    coord = f"{get_column_letter(col)}{row}"
    ref = f"{quote_sheetname(ws.title)}!${get_column_letter(col)}${row}"
    if any(n == name for n, _, _ in ctx.names):
        raise BuildError(f"名前付きレンジ '{name}' が重複しています(ブック内で一意。13章§2.9)")
    ctx.wb.defined_names.add(DefinedName(name, attr_text=ref))
    ctx.names.append((name, ws.title, coord))


def _add_name_block(ctx, ws, name, col, row, height):
    """縦N行×1列の範囲へ名前を定義する(13章§2.18 gd_test_result)。
    _add_name と同じ台帳(ctx.names)へ積むので、自己検証がそのまま参照先を突き合わせる。"""
    letter = get_column_letter(col)
    last = row + max(int(height), 1) - 1
    coord = f"{letter}{row}:{letter}{last}"
    ref = f"{quote_sheetname(ws.title)}!${letter}${row}:${letter}${last}"
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
                        error="一覧から選んでください。")
    ws.add_data_validation(dv)
    dv.add(target)
    return True


def _title_band(ws, row, text, text_col=None):
    """シート1行目のタイトル帯(裁定書14 裁定7)。BRAND_DARK 地＋白太字を A:C の
    3列へ敷く(結合はしない。結合セルは実行時の図形・行操作と相性が悪いため、
    同じ値の見た目だけを3列ぶん塗る)。

    text_col: 表題**文字**を置く列(既定は TITLE_BAND_TEXT_COL、無ければA列)。
    帯の塗り範囲は text_col によらず A:C で不変(裁定書16 F3)。"""
    if text_col is None:
        text_col = TITLE_BAND_TEXT_COL.get(ws.title, 1)
    if not 1 <= int(text_col) <= TITLE_BAND_COLS:
        raise BuildError(
            f"タイトル帯の文字列配置列が帯の範囲外です: シート '{ws.title}' col={text_col}")
    cell = ws.cell(row=row, column=int(text_col), value=_clean(text))
    cell.font = TITLE_BAND_FONT
    cell.alignment = Alignment(vertical="center", indent=1)
    for i in range(1, TITLE_BAND_COLS + 1):
        c = ws.cell(row=row, column=i)
        c.fill = TITLE_BAND_FILL
        c.protection = LOCKED
    ws.row_dimensions[row].height = 26
    return cell


def _style_section_title(ws, row, text):
    """セクション見出し(TITLE_FONT行)。淡緑地を A:C へ敷く(裁定書14 裁定7)。"""
    cell = ws.cell(row=row, column=1, value=_clean(text))
    cell.font = TITLE_FONT
    cell.alignment = Alignment(vertical="center", indent=1)
    for i in range(1, TITLE_BAND_COLS + 1):
        c = ws.cell(row=row, column=i)
        c.fill = SECTION_FILL
        c.protection = LOCKED
    return cell


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
    if spec.get("guide"):
        # 帳票型のうち案内専用シート(13章§2.18)。1行目ヘッダを持たず名前付き
        # レンジで参照する点は帳票型と同じだが、A=見出し/B=値/C=注記の3列組では
        # なく章立ての読み物なので、生成路だけを分ける(台帳の role は form)。
        return _make_guide(wb, spec, ctx)
    ws = wb.create_sheet(spec["name"])
    ws.protection.sheet = False
    ws.column_dimensions["A"].width = spec.get("label_width", 24)
    ws.column_dimensions["B"].width = spec.get("value_width", 60)
    ws.column_dimensions["C"].width = spec.get("note_width", 50)

    _title_band(ws, 1, spec["name"])
    row = 3
    for sec in spec.get("sections") or []:
        _style_section_title(ws, row, sec.get("title", ""))
        if sec.get("anchor"):
            # 区画の見出し行そのものを名前付きレンジで持つ(11章§3.1: [次へ]の
            # 移動先と強調枠の位置決めに使う。行番号をコードへ書かない)。
            _add_name(ctx, ws, sec["anchor"], 1, row)
        row += 1
        tall = bool(sec.get("tall"))
        for fld in sec.get("fields") or []:
            # (1) 名前付きレンジを持たない「説明1行」「例文」「見出し」の行。
            #     11章§3.3.3 が求める常設の案内文で、値ではないので A列へ直に書く。
            if not fld.get("range"):
                txt = ws.cell(row=row, column=1, value=_clean(fld.get("text", "")))
                txt.protection = LOCKED
                txt.alignment = Alignment(vertical="top", wrap_text=True)
                if fld.get("gray"):
                    txt.font = GRAY_FONT
                elif fld.get("head"):
                    txt.font = SUBHEAD_FONT
                else:
                    txt.font = NOTE_FONT
                row += 1
                continue

            lab = ws.cell(row=row, column=1, value=_clean(fld.get("label", fld["range"])))
            lab.protection = LOCKED
            lab.alignment = Alignment(vertical="top")

            # (2) 縦N行×1列の名前付きブロック(プレビュー5行・現場メモ60行・
            #     直貼り枠300行)。予約行方式ではなく、表示と代替入力の枠である
            #     (13章§2.10(c)・11章§3.3.1)。
            height = int(fld.get("block_rows") or 0)
            if height > 0:
                _add_name_block(ctx, ws, fld["range"], 2, row, height)
                for r in range(row, row + height):
                    bc = ws.cell(row=r, column=2)
                    bc.number_format = ctx.text_fmt
                    bc.protection = UNLOCKED if fld.get("input") else LOCKED
                    bc.alignment = Alignment(vertical="top", wrap_text=True)
                if fld.get("note"):
                    nt = ws.cell(row=row, column=3, value=_clean(fld["note"]))
                    nt.font = NOTE_FONT
                    nt.protection = LOCKED
                    nt.alignment = Alignment(vertical="top", wrap_text=True)
                row += height
                continue

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
    # 11章§0.1 原則⑥・§3.1.1: コーチ帯はスクロールしても消えない。
    # ウィンドウ枠固定はビルドが焼く(実行時に窓の設定を触らない)。
    if spec.get("freeze_panes"):
        ws.freeze_panes = spec["freeze_panes"]
    for rng in spec.get("hidden_rows") or []:
        for r in range(int(rng[0]), int(rng[1]) + 1):
            ws.row_dimensions[r].hidden = True
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
        _title_band(ws, row, spec.get("header_title", spec["name"]))
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
        _style_section_title(ws, row, blk["title"])
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


# ---------------------------------------------------------------------------
# 操作ガイド(13章§2.18・裁定書14 裁定3)
# ---------------------------------------------------------------------------
# 移植元: PoC「マイ本棚AI」 build/build_mybookshelf.py の _make_howto。
#   section / step / kv / note / back_to_toc の5ヘルパー、目次を先に「行だけ
#   予約」しておいて全章を書き終えてからハイパーリンクを書き戻す2パス方式、
#   および「目次の予約行数と実際の章数が食い違ったらビルドを止める」検問を
#   そのまま持ち込んだ(予約行数と章数がずれると以降の章が目次へ食い込む=
#   気づきにくい表示崩れになるため)。
#
# 本シートの約束(裁定書14 裁定3):
#   ・図形を1つも作らない(セルとシート内リンクだけ)。ボタンは実行時に
#     modUISheet.EnsureButton が gd_btn_test の位置へ生やす。
#   ・文言は専門用語を使わず、失敗の案内は「何が起きたか。どうすればよいか。」
#     の2文で書く。ボタン名・画面の文言は実装の文字列から引く(推測で書かない)。
GUIDE_HEAD_FILL = PatternFill("solid", fgColor=BRAND_DARK)
GUIDE_SUB_FILL = PatternFill("solid", fgColor=BRAND_MID)
GUIDE_HEAD_FONT = Font(bold=True, size=12, color=WHITE)
GUIDE_STEP_FONT = Font(bold=True, size=10, color=WHITE)
GUIDE_KEY_FONT = Font(bold=True, size=10, color=BRAND_DARK)
GUIDE_BODY_FONT = Font(size=10)
GUIDE_NOTE_FONT = Font(size=9, italic=True, color="6B7280")
GUIDE_LINK_FONT = Font(bold=True, size=10.5, color="1155CC", underline="single")
GUIDE_BACK_FONT = Font(size=9, bold=True, color="1155CC", underline="single")
GUIDE_WARN_FILL = PatternFill("solid", fgColor="FFF4D6")
GUIDE_ZEBRA_FILL = PatternFill("solid", fgColor="F4F6FB")
GUIDE_RESULT_FILL = PatternFill("solid", fgColor="F7F9FA")
GUIDE_THIN = Side(style="thin", color="D9E1EC")

# 指示文セル(B:C結合)の実効幅。列幅28+82の合計で、全角1字を2カウントとして
# 折返し行数を見積もる(行高が足りないと指示文が見切れる=H3と同じ事故になる)。
GUIDE_PROMPT_COLS = 108


def _wrapped_height(text, cols, line_pt=13.5, pad_pt=8):
    """折返しを織り込んだ行高(pt)。全角は2桁ぶんとして数える。"""
    lines = 0
    for raw in text.split("\n"):
        width = sum(2 if ord(ch) > 0x2E80 else 1 for ch in raw)
        lines += max(1, -(-width // cols))
    return max(24, line_pt * lines + pad_pt)

# HOMEの19ボタン。**modUIHome.EnsureHomeButtons のキャプションと逐語一致**
# させること(実装の文字列から引く。推測で書かない)。[S1][S2][S3][S4] の4つは
# 1行にまとめて説明するため、早見表の行数は実装のボタン本数と一致しない。
# HOMEのボタン早見表(13章§2.10のボタン一覧と逐語一致。上4本=番号つきの主要動線)。
# 使い方タブのボタン早見(11章§3.6・13章§2.10(f))。**キャプションの値源はここではなく
# `modUINav` の配置表定数**であり、tools/caption_check.py が実装・13章§2.10・初回
# ツアーの3系統と逐語照合する。第3要素はどの章へ出すか(①～④と帯)。
# 繰り返しキャプションのボタン([コピー]8本・区画②の18本)は、同じ語が複数出ると
# 逐語照合の突合が壊れるため配置表にも本表にも載せない(13章§2.10(f))。
GUIDE_NAV_BUTTONS = [
    ("← 戻る", "1つ上の区画へ画面を動かします。STEPの番号は変わりません。", "帯"),
    ("次へ →", "いまのSTEPの区画へ画面を動かし、黄色い枠をそこへ移します。", "帯"),
    ("ナレッジを読み直す", "社内ナレッジ(ナレッジブック.xlsx)をもう一度読み込みます。", "帯"),
    ("使い方を開く", "この「使い方」のタブを開きます。", "帯"),
    ("＋ もっと調べる（あと5本）", "調べる文の残り5本を出したり隠したりします。", "①"),
    ("貼ったものを保存する", "6つの欄の中身を、いまの案件として保存します。", "②"),
    ("まとめて作る", "下書きを4枚作ります。10～20分かかります。", "③"),
    ("レポートを出す", "お客様に見せるレポートを作り、ブラウザで開きます。", "④"),
    ("ヒアリングシートを出す", "訪問で聞くことの紙を作り、そのタブを出します。", "④"),
    ("下書きを見る", "下書きのタブを4枚出して、いちばん上の①へ移ります。", "④"),
]

# 使い方タブ⑦「上級」の5行(11章§3.6・§1.1 の1行目の説明文と**同一文**)。
GUIDE_ADVANCED = [
    ("壁打ち", "下書きが③まで出来てから、提案の前の予行演習をしたいときに使います。ふだんの案件では使いません。"),
    ("受信箱", "思いついた新しいサービスの案を、部で持ち寄って診断したいときに使います。日々の案件づくりには使いません。"),
    ("商談の記録", "商談のあとで、起きたことだけを記録したいときに使います。次アポ・見積依頼・成約などを選ぶだけです。"),
    ("判断台帳", "引受の判断を記録する台帳です。商品部で使います。営業所では使いません。"),
    ("案件一覧", "これまでに作った案件を一覧で見たいときに使います。ふだんは見なくて大丈夫です。"),
]

# 使い方タブ⑦「上級」の**動作ボタン**3本(11章§3.6・§3.7・13章§2.18)。
# 1件 = (キャプション, いつ使うか(1行), OnAction)。tools/caption_check.py が
# modUIGuide の配線と 13章§2.18 の表へ逐語照合する唯一の値源。
# [表示する](タブを1枚出すだけ)と違い、この3本は**押すと処理が走る**ため、
# キャプションを1語でも変えると押した人の期待が外れる。
GUIDE_ADVANCED_ACTIONS = [
    ("第2ラウンドを始める",
     "訪問して聞いてきたことを②へ貼り、もう一度作り直したいときに使います。",
     "modUIHome2.HomeFreezeRound"),
    ("企業ファイルへ保存",
     "同じ会社を次に担当する人へ、この案件の内容を引き継ぎたいときに使います。",
     "modUIHome2.HomeCompanySave"),
    ("企業ファイルを開く",
     "前の担当者が残した企業ファイルを、いまの案件へ取り込みたいときに使います。",
     "modUIHome2.HomeCompanyOpen"),
]

# 困ったとき(docs/25章§12「困ったときの1行対処」の要約。平易語)。
GUIDE_TROUBLES = [
    ("画面が白いまま動かない", "処理の最中です。触らずに待ってください。"),
    ("「対象案件が選ばれていません」と出た", "案件が選ばれていません。HOMEの対象案件を選び直してください。"),
    ("「先にStep1を実行してください。」と出た", "前の段がまだ終わっていません。[S1]を押してから出力してください。"),
    ("「本日のAI利用枠の上限です」と出た", "きょうの利用枠を使い切りました。翌日に続きの段から押し直してください。"),
    ("「画面の案件と保存先が一致しません。再描画してください」と出た",
     "別の案件の画面が残っています。HOMEで対象案件を選び直し、その段のシートを開き直してください。"),
    ("「ナレッジブック.xlsx を本体と同じフォルダに置いて[ナレッジ再読込]を押してください」と出た",
     "社内ナレッジにつながっていません。ナレッジブック.xlsx を本体ブックと同じフォルダに置き、"
     "[ナレッジ再読込]を押してください(起動時は自動でも探します)。"
     "同じフォルダに置いてあるのに出るときは、そのファイルを右クリック→プロパティ→"
     "[許可する]にチェックを入れて開き直してください(本体とナレッジブックの2ファイルとも)。"
     "それでも直らなければ管理者へ連絡してください。"),
    ("ボタンを押しても何も起きない", "ほかの処理が動いています。終わるまで待ってください。"),
    ("英語で Trust... というメッセージが出た", "VBAの信頼設定がまだ終わっていません。①の手順3と4をやり直してください。"),
    ("出力の内容が明らかにおかしい", "AIの答えがずれています。シートを直して、その段から作り直してください。直らなければ管理者へ連絡してください。"),
    ("警告に「（err_logタブの最後の行を開発担当へ送ってください）」と出た",
     "うまくいかなかった記録が残っています。画面いちばん下の err_log タブを開き、いちばん下の行をコピーして開発担当へ送ってください。"),
]

# 操作ガイド⑦「AIに調べさせる指示文」(裁定書17 H5)。
# ------------------------------------------------------------------------------
# 実機フィードバック第2報: 指示文集(docs/08)は利用者へ届いていなかった。
# **プロンプトはアプリの中に無ければならない**ので、docs/08 のコードフェンスを
# **逐語で**この章へ焼き込む。値源を二重に持たない(コピーを本ファイルへ書き写す
# と必ず腐る)ため、ビルド時に docs/08 を読んで取り出す。
DOSSIER_DOC_PATH = os.path.join("docs", "08_ドシエ収集プロンプト集.md")

# 標準3本(まずこの3本)と補助5本(必要なときだけ)。表示名は利用者の言葉にする
# (docs/08 の見出しは開発側の呼び方なのでそのままにしない)。
GUIDE_PROMPTS_STD = [
    ("D-1", "D-1 会社の基礎調査"),
    ("D-3", "D-3 リスクの兆候さがし"),
    ("D-9", "D-9 調達・仕入れの調べ"),
]
GUIDE_PROMPTS_OPT = [
    ("D-2", "D-2 業界・競合"),
    ("D-4", "D-4 世の中の動き(為替・人手・サイバー)"),
    ("D-6", "D-6 更新案件の変化さがし"),
    ("D-7", "D-7 拠点の地域概況(ハザード)"),
    ("D-8", "D-8 決算ハイライト(上場企業)"),
]
# 冒頭の運用ルール(直列運用。docs/08 1x の実測とdocs/26【1-5】に合わせる)。
GUIDE_PROMPT_RULES = [
    ("投げ方",
     "社内ディープリサーチは1本ずつしか動きません。1本ずつ順に投げ、各10分ほど待ちます。"),
    ("企業名の書き方",
     "企業名には必ず本社所在地か証券コードを添えてください(似た名前の別会社の情報が混ざった実例があります)。"),
    ("長さの上限",
     "指示文は2,000字以内です(下の指示文は全部収まっています)。"),
]


def read_dossier_prompts(root):
    """docs/08 の `## D-N. ...` 見出し直下の最初のコードフェンスを逐語で返す。

    操作ガイド⑦(裁定書17 H5)が1本1セルへそのまま入れる。欠落・空はビルドを
    止める(指示文の入っていないブックを配ると、利用者はどこにも辿り着けない)。
    """
    path = os.path.join(root, DOSSIER_DOC_PATH)
    if not os.path.exists(path):
        raise BuildError(
            f"{DOSSIER_DOC_PATH} が見つかりません({path})。"
            "操作ガイド⑦へ指示文を逐語転載できないためビルドを中止します。")
    with open(path, encoding="utf-8") as fp:
        text = fp.read()

    found = {}
    heads = list(re.finditer(r"^##\s+(D-\d)\.", text, re.M))
    for i, m in enumerate(heads):
        tail = heads[i + 1].start() if i + 1 < len(heads) else len(text)
        fence = re.search(r"^```[^\n]*\n(.*?)\n```", text[m.end():tail], re.S | re.M)
        if fence:
            found[m.group(1)] = fence.group(1)

    wanted = [k for k, _ in GUIDE_PROMPTS_STD] + [k for k, _ in GUIDE_PROMPTS_OPT]
    missing = [k for k in wanted if not (found.get(k) or "").strip()]
    if missing:
        raise BuildError(
            f"{DOSSIER_DOC_PATH} から指示文を取り出せませんでした: {missing}"
            "(`## D-N. ...` の直後のコードフェンスが値源です)")
    over = [k for k in wanted if len(found[k]) > 2000]
    if over:
        raise BuildError(
            f"指示文が2,000字を超えています: {over}"
            "(社内ディープリサーチの入力上限。docs/08 1x の実測)")
    return {k: found[k] for k in wanted}


GUIDE_GLOSSARY = [
    ("かんたん調査", "いちばん軽い調べ方。ホームページと営業メモくらいで進めます。"),
    ("しっかり調査", "重要な案件向け。AIに追加で調べさせた結果も貼って進めます。"),
    ("壁打ち", "AIと相談しながら座組を考える画面のこと。使い始めると調査の深さが自動で切り替わります。"),
    ("受信箱", "投稿・現場の声・ウォッチ結果をいったん受ける一覧のこと。"),
    ("ナレッジ", "社内にためた事例・メニュー・型のこと。別ファイル(ナレッジブック)にあります。"),
    ("判断台帳", "引受の判断と、その後どうなったかを残す記録のこと。"),
    ("段(S1からS4)", "S1=会社を知る、S2=リスクを出す、S3=提案を作る、S4=提案書の骨子を作る、の4段階のこと。"),
    ("入念", "AIに自分の答えを批判させて直させる作り方。時間はかかりますが質が上がります。"),
]


def _make_guide(wb, spec, ctx):
    """使い方: ブックの中だけで操作を学べる案内シート(セルのみ・図形なし)。
    11章§3.6(v3.2)の7章立て。①～④はナビの4区画と同じ順・同じ見出し語。"""
    prompts = read_dossier_prompts(ctx.root)
    ws = wb.create_sheet(spec["name"])
    ws.protection.sheet = False
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 82

    _title_band(ws, 1, f"{APP_TITLE} 使い方")

    row = [3]

    def anchor(cell, target_row, tip):
        cell.hyperlink = Hyperlink(
            ref="", location=f"{quote_sheetname(ws.title)}!A{target_row}", tooltip=tip)

    def section(title, fill=BRAND_DARK):
        if row[0] > 3:
            row[0] += 1
        r = row[0]
        c = ws.cell(row=r, column=1, value=_clean(title))
        c.font = GUIDE_HEAD_FONT
        c.alignment = Alignment(vertical="center", indent=1)
        for i in range(1, 4):
            ws.cell(row=r, column=i).fill = PatternFill("solid", fgColor=fill)
        ws.row_dimensions[r].height = 24
        row[0] = r + 2
        return r

    def back_to_toc(target_row):
        r = row[0]
        c = ws.cell(row=r, column=2, value="▲ 目次へ戻る")
        c.font = GUIDE_BACK_FONT
        c.alignment = Alignment(vertical="center", indent=1)
        anchor(c, target_row, "クリックで目次に戻ります")
        ws.row_dimensions[r].height = 18
        row[0] = r + 2

    def step(n, text, warn=False):
        r = row[0]
        c0 = ws.cell(row=r, column=1, value=n)
        c0.font = GUIDE_STEP_FONT
        c0.fill = GUIDE_SUB_FILL
        c0.alignment = Alignment(horizontal="center", vertical="center")
        c = ws.cell(row=r, column=2, value=_clean(text))
        c.font = GUIDE_BODY_FONT
        c.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
        if warn:
            c.fill = GUIDE_WARN_FILL
        ws.row_dimensions[r].height = max(20, 15 * (text.count("\n") + 1) + 6)
        row[0] = r + 1

    def kv(k, v, i=0):
        r = row[0]
        ck = ws.cell(row=r, column=2, value=_clean(k))
        ck.font = GUIDE_KEY_FONT
        ck.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
        cv = ws.cell(row=r, column=3, value=_clean(v))
        cv.font = GUIDE_BODY_FONT
        cv.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
        if i % 2 == 1:
            ck.fill = GUIDE_ZEBRA_FILL
            cv.fill = GUIDE_ZEBRA_FILL
        for c in (ck, cv):
            c.border = Border(bottom=GUIDE_THIN)
        ws.row_dimensions[r].height = max(20, 15 * (v.count("\n") + 1) + 5)
        row[0] = r + 1

    def note(text):
        r = row[0]
        c = ws.cell(row=r, column=2, value=_clean(text))
        c.font = GUIDE_NOTE_FONT
        c.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
        ws.row_dimensions[r].height = max(18, 14 * (text.count("\n") + 1) + 5)
        row[0] = r + 2

    def prompt(title, body):
        """指示文1本(見出し1行＋全文1セル)。全文は B:C 結合・折返しで、
        セルを1つ選んで Ctrl+C すればそのまま投げられる形にする(裁定書17 H5)。"""
        r = row[0]
        ck = ws.cell(row=r, column=2, value=_clean(title))
        ck.font = GUIDE_KEY_FONT
        ck.alignment = Alignment(vertical="center", indent=1)
        ws.row_dimensions[r].height = 20
        r += 1
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        cv = ws.cell(row=r, column=2, value=_clean(body))
        cv.font = GUIDE_BODY_FONT
        cv.fill = GUIDE_RESULT_FILL
        cv.number_format = ctx.text_fmt
        cv.alignment = Alignment(vertical="top", wrap_text=True, indent=1)
        ws.row_dimensions[r].height = _wrapped_height(body, GUIDE_PROMPT_COLS)
        row[0] = r + 2

    # ---- 目次(行だけ予約しておき、全章を書き終えてから書き戻す) ----------
    # 11章§3.6(v3.2): 7章立て。①～④は**ナビの4区画と同じ順・同じ見出し語**にする
    # (2箇所で違う呼び方をしない)。各章に書くのは「何をどこへ」「押すと何が起きる」
    # 「うまくいかないとき」の3つだけで、設計語・開発語は1語も持ち込まない。
    toc_n = 7
    toc_header_row = row[0]
    row[0] = toc_header_row + 2
    toc_first_row = row[0]
    row[0] += toc_n
    row[0] += 1

    def nav_buttons(chapter):
        """その章に属するボタンの早見(値源は GUIDE_NAV_BUTTONS)。
        tools/caption_check.py がここのキャプションを実装・13章§2.10と逐語照合する。"""
        rows = [(c, d) for c, d, ch in GUIDE_NAV_BUTTONS if ch == chapter]
        for i, (cap, desc) in enumerate(rows):
            kv(f"[{cap}]", desc, i)

    # ==== ① 会社のこと =====================================================
    ch1 = section("① 会社のこと")
    kv("何をどこへ",
       "ナビの①へ、種別・会社名・業種名・本社の場所の4つを入れます。\n"
       "4つを入れると、社内のディープリサーチに貼る文がすぐ下に出ます。", 0)
    kv("押すと何が起きる",
       "文の右の[コピー]を押すと、その1本ぶんが写ります。社内のディープリサーチに貼って投げます。\n"
       "1本ずつしか投げられません。返ってきてから次を投げます(各10分ほど)。", 1)
    kv("うまくいかないとき",
       "「文を写せませんでした」と出たら、枠の中の文をマウスで選んで Ctrl+C で写してください。\n"
       "業種の一覧が出ないときは、手で入力しても先へ進めます。", 0)
    nav_buttons("①")
    note("いちばん上の帯にあるボタンは、どの区画にいても使えます。")
    nav_buttons("帯")

    # はじめて開いたときに出る案内(ガイドツアー)をもう一度見るためのボタン置き場。
    # 図形は起動時に modUIGuide.EnsureGuideButtons が生やす(ビルドは作らない)。
    r = row[0]
    lab = ws.cell(row=r, column=2, value="はじめの案内")
    lab.font = GUIDE_KEY_FONT
    lab.alignment = Alignment(vertical="center", indent=1)
    ws.cell(row=r, column=3).alignment = Alignment(vertical="center", indent=1)
    _add_name(ctx, ws, "gd_btn_tour", 3, r)
    ws.row_dimensions[r].height = 28
    row[0] = r + 2
    back_to_toc(toc_header_row)

    # ==== ② 貼る ===========================================================
    ch2 = section("② 貼る")
    kv("何をどこへ",
       "ナビの②の6つの枠へ、返ってきた文章・会社のホームページ・有価証券報告書のリスクの章・\n"
       "いまの契約・ヒアリング回答を貼ります。現場メモだけは見出しの下に手で書きます。", 0)
    kv("押すと何が起きる",
       "[ここに貼る]を押すと、表の線や画像が入らずに文だけが入り、字数と貼った時刻がその場で出ます。\n"
       "枠には先頭の5行だけが出ます(途中で切れて保存されたわけではありません)。", 1)
    kv("うまくいかないとき",
       "「枠に入りきりませんでした」と出たら、[ここに貼る]で貼り直してください。長さは気にしなくて大丈夫です。\n"
       "「貼り付けられませんでした」と出たら、もう一度コピーしてから押してください。", 0)
    nav_buttons("②")
    note("[中身を見る]は読むだけの画面です。直したいときは、直した文章を[ここに貼る]で貼り直してください。\n"
         "[消す]は押す前に必ず確認が出ます。消すと元に戻せません。")
    back_to_toc(toc_header_row)

    # ==== ③ 作る ===========================================================
    ch3 = section("③ 作る")
    kv("何をどこへ", "貼り終えたら、ナビの③のボタンを押すだけです。入れるものはありません。", 0)
    kv("押すと何が起きる",
       "下書きを4枚作ります。10～20分かかります。\n"
       "画面が白くなっても動いていますので、閉じずにお待ちください。", 1)
    kv("うまくいかないとき",
       "「途中で止まりました」と出たら、もう一度押してください。\n"
       "それでも止まるときは、下の「困ったとき」の[記録を見る]を押してください。", 0)
    nav_buttons("③")
    back_to_toc(toc_header_row)

    # ==== ④ 出す ===========================================================
    ch4 = section("④ 出す")
    kv("何をどこへ", "ナビの④の3つのボタンを押すだけです。入れるものはありません。", 0)
    kv("押すと何が起きる",
       "レポートはブラウザで開きます。ヒアリングシートはタブが出ますので、そのまま印刷します。\n"
       "下書きは4枚のタブが出ます。違うところは手で直してください。", 1)
    kv("うまくいかないとき",
       "「先に③の[まとめて作る]を押してください」と出たら、③へ戻って押してください。", 0)
    nav_buttons("④")
    note("お客様に見せる前に、レポートの中身を必ずご確認ください。")
    back_to_toc(toc_header_row)

    # ==== ⑤ 困ったとき =====================================================
    ch5 = section("⑤ 困ったとき")
    note("直らないときは、下の[記録を見る]を押して、いちばん下の行をコピーして開発担当へ送ってください。")
    r = row[0]
    lab = ws.cell(row=r, column=2, value="記録を見る")
    lab.font = GUIDE_KEY_FONT
    lab.alignment = Alignment(vertical="center", indent=1)
    ws.cell(row=r, column=3).alignment = Alignment(vertical="center", indent=1)
    _add_name(ctx, ws, "gd_btn_logs", 3, r)
    ws.row_dimensions[r].height = 28
    row[0] = r + 2
    for i, (sym, fix) in enumerate(GUIDE_TROUBLES):
        kv(sym, fix, i)
    back_to_toc(toc_header_row)

    # ==== ⑥ 自己テスト =====================================================
    ch6 = section("⑥ 自己テスト(このブックが正しく動くかを自分で確かめる)")
    note("配ったファイルが途中で壊れていないかを、このブックの中だけで確かめられます。\n"
         "新しい版を受け取ったときと、動きがおかしいと感じたときに1回押してください。")
    step("1", "下の[テストを実行]を押します。確認の画面が出たら[はい]を押します。")
    step("2", "数分かかります。終わると合否が出て、下の結果欄に全文が入ります。")
    step("3", "合格でなければ、結果欄をそのままコピーして管理者へ送ってください。")

    r = row[0]
    lab = ws.cell(row=r, column=2, value="テストの実行")
    lab.font = GUIDE_KEY_FONT
    lab.alignment = Alignment(vertical="center", indent=1)
    btn_cell = ws.cell(row=r, column=3)
    btn_cell.alignment = Alignment(vertical="center", indent=1)
    _add_name(ctx, ws, "gd_btn_test", 3, r)
    ws.row_dimensions[r].height = 28
    row[0] = r + 2

    r = row[0]
    lab = ws.cell(row=r, column=2, value="テストの結果")
    lab.font = GUIDE_KEY_FONT
    lab.alignment = Alignment(vertical="top", indent=1)
    result_rows = int(spec.get("test_result_rows") or 20)
    for k in range(result_rows):
        c = ws.cell(row=r + k, column=3)
        c.font = GUIDE_BODY_FONT
        c.fill = GUIDE_RESULT_FILL
        c.number_format = ctx.text_fmt
        c.alignment = Alignment(vertical="top", wrap_text=True, indent=1)
        c.protection = LOCKED
    _add_name_block(ctx, ws, "gd_test_result", 3, r, result_rows)
    row[0] = r + result_rows + 1
    note("結果欄は実行するたびに上書きされます(前回ぶんは残りません)。")
    back_to_toc(toc_header_row)

    # ==== ⑦ 上級 ===========================================================
    ch7 = section("⑦ 上級(ふだんは使いません)")
    # ナビの[使い方を開く]と、上級の飛び先(13章§2.18)。見出し行を名前付きレンジで
    # 指し、実行時は文字列検索をしない。
    _add_name(ctx, ws, "gd_ch7_head", 1, ch7)
    note("下の[表示する]を押すと、そのタブが1枚だけ出ます。ふだんの案件づくりでは使いません。")
    for i, (sheet_name, when) in enumerate(GUIDE_ADVANCED):
        kv(sheet_name, when, i)
    r = row[0]
    lab = ws.cell(row=r, column=2, value="上級のタブを出す")
    lab.font = GUIDE_KEY_FONT
    lab.alignment = Alignment(vertical="top", indent=1)
    for k in range(len(GUIDE_ADVANCED)):
        ws.cell(row=r + k, column=3).alignment = Alignment(vertical="center", indent=1)
        ws.row_dimensions[r + k].height = 28
    _add_name_block(ctx, ws, "gd_btn_adv", 3, r, len(GUIDE_ADVANCED))
    row[0] = r + len(GUIDE_ADVANCED) + 1

    # 上級の動作ボタン3本(11章§3.6・裁定書22)。押すと処理が走るので[表示する]と
    # 行を分け、1行ずつ「いつ使うか」を左に書く。ボタンは実行時に
    # modUIGuide.EnsureGuideButtons が gd_btn_adv_act を起点に生やす。
    note("下の3つは押すと処理が走ります。ふだんの案件づくりでは使いません。")
    for i, (cap, when, _act) in enumerate(GUIDE_ADVANCED_ACTIONS):
        kv(cap, when, i)
    r = row[0]
    lab = ws.cell(row=r, column=2, value="上級の操作")
    lab.font = GUIDE_KEY_FONT
    lab.alignment = Alignment(vertical="top", indent=1)
    for k in range(len(GUIDE_ADVANCED_ACTIONS)):
        ws.cell(row=r + k, column=3).alignment = Alignment(vertical="center", indent=1)
        ws.row_dimensions[r + k].height = 28
    _add_name_block(ctx, ws, "gd_btn_adv_act", 3, r, len(GUIDE_ADVANCED_ACTIONS))
    row[0] = r + len(GUIDE_ADVANCED_ACTIONS) + 1
    back_to_toc(toc_header_row)

    # ==== 非表示行: 調べる文8本の雛形 ======================================
    # 11章§3.2・13章§2.18(v3.2): 利用者が読む面は**ナビの区画①**であり、本文を
    # ここで章として見せない。ただし値の在処は1箇所でなければならないので、
    # 本文は本シートの非表示行へ焼き、ナビはこのセルを読んで {{ }} を置換する。
    r = row[0]
    hdr = ws.cell(row=r, column=2, value="（ここから下は、ナビが読むための行です。利用者が読む面はナビの①です）")
    hdr.font = GUIDE_NOTE_FONT
    ws.row_dimensions[r].hidden = True
    r += 1
    for key, title in (GUIDE_PROMPTS_STD + GUIDE_PROMPTS_OPT):
        idx = (GUIDE_PROMPTS_STD + GUIDE_PROMPTS_OPT).index((key, title)) + 1
        ck = ws.cell(row=r, column=2, value=_clean(title))
        ck.font = GUIDE_KEY_FONT
        cv = ws.cell(row=r, column=3, value=_clean(prompts[key]))
        cv.font = GUIDE_BODY_FONT
        cv.number_format = ctx.text_fmt
        cv.alignment = Alignment(vertical="top", wrap_text=True, indent=1)
        cv.protection = LOCKED
        _add_name(ctx, ws, "gd_prompt_%02d" % idx, 3, r)
        ws.row_dimensions[r].hidden = True
        r += 1
    row[0] = r + 1

    # ---- 目次の書き戻し ---------------------------------------------------
    hc = ws.cell(row=toc_header_row, column=1, value="目次(クリックすると各章へ移動します)")
    hc.font = GUIDE_HEAD_FONT
    hc.alignment = Alignment(vertical="center", indent=1)
    for i in range(1, 4):
        ws.cell(row=toc_header_row, column=i).fill = GUIDE_HEAD_FILL
    ws.row_dimensions[toc_header_row].height = 24

    toc_entries = [
        (ch1, "① 会社のこと"),
        (ch2, "② 貼る"),
        (ch3, "③ 作る"),
        (ch4, "④ 出す"),
        (ch5, "⑤ 困ったとき"),
        (ch6, "⑥ 自己テスト"),
        (ch7, "⑦ 上級(ふだんは使いません)"),
    ]
    if len(toc_entries) != toc_n:
        raise BuildError(
            f"_make_guide: 目次の予約行数({toc_n})と実際の章数({len(toc_entries)})が"
            "不一致です(以降の章が目次へ食い込みます)")
    for i, (target_row, title) in enumerate(toc_entries):
        r = toc_first_row + i
        c = ws.cell(row=r, column=2, value=title)
        c.font = GUIDE_LINK_FONT
        c.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
        if i % 2 == 1:
            c.fill = GUIDE_ZEBRA_FILL
        anchor(c, target_row, "クリックでこの章へ移動します")
        ws.row_dimensions[r].height = 20

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

    # 固定セル(13章§2.9に位置を1行記録): E1=起動直後の modBoot.Boot 予約時刻
    # (自己インストーラが Application.OnTime で置く。**ビルドは触らない**)。
    # E2=ブック内テストの期待本数(裁定書14 裁定5)。wintest/tests_expected.txt の
    # 1行目を焼き込み、実行時に modTestsRunnerUi が読む。読めなければ実行しない
    # (fail-closed)ため、txtの欠落・非整数はここでビルドを止める。
    # E3=焼き付けマーカー "baked"(裁定書23 C-1)。**ビルドは触らない**。実機の
    # Excelが1回目の起動で注入に全件成功したときだけ書き、以後の起動は
    # VBProject に触らない(利用者にVBOM不要)。
    ws.cell(row=2, column=5, value=ctx.tests_expected)
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 80
    ws.column_dimensions["D"].width = 14
    ws.sheet_state = spec.get("state", "veryHidden")
    ctx.injected = injected
    return ws


TESTS_EXPECTED_PATH = os.path.join("wintest", "tests_expected.txt")


def read_tests_expected(root):
    """wintest/tests_expected.txt の1行目の整数(ブック内テストの期待本数)を返す。
    欠落・非整数はビルドを止める(裁定書14 裁定5。実行時は E2 を読めなければ
    テストを実行しない fail-closed なので、焼き込み側で必ず担保する)。"""
    path = os.path.join(root, TESTS_EXPECTED_PATH)
    if not os.path.exists(path):
        raise BuildError(
            f"{TESTS_EXPECTED_PATH} が見つかりません({path})。"
            "ブック内テストの期待本数(vba_src!E2)を焼き込めないためビルドを中止します。")
    with open(path, encoding="utf-8-sig") as fp:
        first = fp.readline().strip()
    if not re.fullmatch(r"[0-9]+", first):
        raise BuildError(
            f"{TESTS_EXPECTED_PATH} の1行目が整数ではありません: {first!r}。"
            "ブック内テストの期待本数(vba_src!E2)として使えないためビルドを中止します。")
    return int(first)


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
                        error="一覧から選んでください。")
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
    (3) _VBA_PROJECT の無害化(サイズ不変・Version=0xFFFF・PerformanceCacheゼロ埋め)
    (4) 焼き付け(baked)経路の存在(裁定書23 C-1)"""
    errors = []

    # (4) 焼き付け経路の存在検査。E3="baked" の読み書き・VBProjectを避ける分岐・
    #     Run失敗時の1文が揃っていなければ、利用者にVBOMを求めない配布が成立しない。
    src_text = installer_src.decode("ascii", errors="replace") if installer_src else ""
    for needle, why in (
            ('If CStr(w.Cells(3, 5).Value) = "baked" Then',
             "起動時に vba_src!E3 を読む baked 分岐"),
            ('w.Cells(3, 5).Value = "baked"',
             "注入成功時に vba_src!E3 へ baked を書く行"),
            ('Application.Run "modBoot.Boot"',
             "OnTime失敗時の Application.Run フォールバック"),
            ('MsgBox "Setup NG: modules missing. Ask the developer."',
             "baked なのに Run が失敗したときの1文"),
    ):
        if needle not in src_text:
            errors.append(
                f"自己インストーラに{why}がありません(裁定書23 C-1): {needle!r}")

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
        want_tab = (spec.get("tab_color") or "").upper()
        if want_tab:
            got_tab = getattr(wb2[name].sheet_properties.tabColor, "rgb", None) or ""
            if not str(got_tab).upper().endswith(want_tab):
                errors.append(
                    f"シート'{name}'のタブ色不一致: 期待={want_tab} 実際={got_tab!r}")

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
        # 裁定書14 裁定5: E2=ブック内テストの期待本数(wintest/tests_expected.txt)。
        if ctx is not None:
            got_e2 = ws.cell(row=2, column=5).value
            ok_e2 = (isinstance(got_e2, (int, float)) and not isinstance(got_e2, bool)
                     and int(got_e2) == ctx.tests_expected)
            if not ok_e2:
                errors.append(
                    f"vba_src!E2(ブック内テストの期待本数)が {TESTS_EXPECTED_PATH} と"
                    f"不一致: 期待={ctx.tests_expected} 実際={got_e2!r}")
        # 裁定書23 C-1: E3=焼き付けマーカー。ビルドは書かない(焼き付けは実機の
        # Excelが1回目の起動でだけ書く)。成果物に baked が立っていると、まだ
        # モジュールが焼き付いていないファイルが注入をスキップしてしまう。
        got_e3 = ws.cell(row=3, column=5).value
        if got_e3 not in (None, ""):
            errors.append(
                f"vba_src!E3(焼き付けマーカー)が成果物で空ではありません: 実際={got_e3!r}"
                "(ビルドはE3を書きません。焼き付け前に baked が立っていると"
                "モジュール注入がスキップされます)")
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
    try:
        # 白字を置く面のコントラスト比を機械検算してから塗り始める(裁定書14 裁定7)。
        checked = _assert_contrast([
            (WHITE, BRAND_DARK, "シート1行目のタイトル帯"),
            (WHITE, BRAND_MID, "操作ガイドの章内見出し・手順番号バッジ"),
        ])
        ctx = BuildCtx(wb, sheets_data, mock_llm, app_version, present, root)
        for spec in sheets:
            role = spec.get("role")
            builder = SHEET_BUILDERS.get(role)
            if builder is None:
                raise BuildError(f"sheets_main.json: 未知の role '{role}'(シート {spec['name']})")
            builder(wb, spec, ctx)
            tab = spec.get("tab_color")
            if tab:
                wb[spec["name"]].sheet_properties.tabColor = tab
    except BuildError as e:
        sys.exit(f"ERROR: {e}")
    print("  白字面のコントラスト比(4.5:1以上を機械検算): "
          + " / ".join(f"{w} {r:.1f}:1" for _, _, w, r in checked))
    print(f"  ブック内テストの期待本数(vba_src!E2 へ焼込): {ctx.tests_expected} "
          f"({TESTS_EXPECTED_PATH})")

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
          "各物理行<=1,000バイト(CP932) / "
          "vba_src本文がsrc/と完全一致 / vba_src D列(期待行数)がsrc由来の計算値と一致 / "
          "configキー列(順序含む)と全キーの説明・mock_llm・app_version / "
          "名前付きレンジの本数と参照先 / 1行目ヘッダとブロックアンカー先ヘッダ行 / "
          "タブ色 / vba_src!E2(ブック内テストの期待本数) / vba_src!E3(焼き付けマーカーが空) / "
          "パッケージ content type")
    if has_vba_project:
        print("  自己インストーラ外科パッチ検証(4項目): "
              "ThisWorkbook復元確認 / dir MOFFSET=0 / "
              "_VBA_PROJECT無害化(Version=0xFFFF・PerformanceCacheゼロ埋め・3,061B不変) / "
              "焼き付け(baked)経路の存在")
    print("  ※13章との突合(シート名・列名・順序・configキー・名前付きレンジ)は "
          "tools/sheet_check.py が行う。")
    print("\nDone.")


if __name__ == "__main__":
    main()
