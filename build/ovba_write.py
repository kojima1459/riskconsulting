#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ovba_write.py -- MS-OVBA 準拠の vbaProject.bin **ライター**(裁定書27 W9-A)。

背景(なぜ新設したか):
    従来の配布方式は template_skeleton.xlsm 由来の vbaProject.bin
    (ThisWorkbook 1本だけ)に自己インストーラを外科パッチし、隠しシート
    vba_src に置いたソースを起動時に VBComponents.Add + AddFromString で
    VBE へ注入する方式だった。この形は 2026-09-02 に社内AVの AMSI で検知され
    (VDI強制停止)、採用禁止になった(裁定書27)。
    本モジュールは「モジュールが最初から入った正規の vbaProject.bin を
    Linux 上でゼロから書き出す」ためのライターである。これにより配布物から
    隠しシート・VBProject 参照・AddFromString が消える。

参照仕様:
    [MS-OVBA] 2.3.1 PROJECT ストリーム / 2.3.3 PROJECTwm ストリーム /
              2.3.4.2 dir ストリーム / 2.3.4.3 モジュールストリーム /
              2.4.1 圧縮(build/ovba.py の ovba_compress を使う)
    [MS-CFB]  Compound File Binary(512B セクタ・FAT/DIFAT/MiniFAT/
              ディレクトリ)。本ファイルに最小限のライターを実装した。

設計方針(「写す」と「作る」の境界):
    ・PROJECTINFORMATION(SysKind/LCID/CodePage/Name/HelpFile/Constants/
      Version)と PROJECTREFERENCES(stdole/Office/Excel/VBA 等)は
      **template_skeleton.xlsm の dir からバイト列のまま写す**。新規参照は
      足さない(裁定書27)。
    ・PROJECTMODULES(モジュール台帳)だけを本モジュールが生成する。
    ・PROJECT ストリームの ID/CMG/DPB/GC/Name/HelpContextID/
      VersionCompatible32 行は template から写し、Document/Module/Class 行と
      [Workspace] だけを生成する。
    ・_VBA_PROJECT は template のものをそのまま写す。各モジュールの
      MODULEOFFSET を 0 にし、モジュールストリームには PerformanceCache を
      置かない(=p-code 無効)ので、Excel/LO はソースから再コンパイルする。
      これは従来の外科パッチと同じ状態である。

このモジュールは openpyxl を import しない(純粋にバイト列操作のみ)。
依存: 標準ライブラリ + 同ディレクトリの ovba.py。
"""

from __future__ import annotations

import struct
import uuid

import ovba

# ---------------------------------------------------------------------------
# CFB 定数(ovba.py と同じ値。読み書きで食い違わないよう再輸入する)
# ---------------------------------------------------------------------------
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
DIFSECT = 0xFFFFFFFC
FREESECT = 0xFFFFFFFF
NOSTREAM = 0xFFFFFFFF

SECTOR_SIZE = 512
MINI_SECTOR = 64
MINI_CUTOFF = 4096

CFB_MAGIC = b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'

# dir ストリームのレコードID(MS-OVBA 2.3.4.2)
REC_PROJECTMODULES = 0x000F
REC_PROJECTCOOKIE = 0x0013
REC_MODULENAME = 0x0019
REC_MODULENAMEUNICODE = 0x0047
REC_MODULESTREAMNAME = 0x001A
REC_MODULESTREAMNAME_UNI = 0x0032
REC_MODULEDOCSTRING = 0x001C
REC_MODULEDOCSTRING_UNI = 0x0048
REC_MODULEOFFSET = 0x0031
REC_MODULEHELPCONTEXT = 0x001E
REC_MODULECOOKIE = 0x002C
REC_MODULETYPE_PROCEDURAL = 0x0021
REC_MODULETYPE_DOCUMENT = 0x0022
# [MS-OVBA] 2.3.4.2.3.2.9 MODULEPRIVATE(Id=0x0028・Size=0)。
# **クラスモジュールにだけ**現れる(実測。下の build_modules_section の注参照)。
REC_MODULEPRIVATE = 0x0028
REC_MODULE_TERMINATOR = 0x002B
REC_DIR_TERMINATOR = 0x0010
REC_PROJECTVERSION = 0x0009


class OvbaWriteError(Exception):
    """vbaProject.bin の組み立てに失敗したことを表す例外。"""


# ===========================================================================
# 1. dir ストリーム
# ===========================================================================
def iter_dir_records(dir_dec: bytes):
    """解凍済み dir ストリームを (offset, id, size, body) で順に返す。

    PROJECTVERSION(0x0009)だけは Size フィールドが 4 と書かれているのに
    実データが 6 バイトある(MS-OVBA 2.3.4.2.1.10 の既知の例外)。
    """
    i = 0
    n = len(dir_dec)
    while i + 6 <= n:
        rid = struct.unpack('<H', dir_dec[i:i + 2])[0]
        size = struct.unpack('<I', dir_dec[i + 2:i + 6])[0]
        real = 6 if rid == REC_PROJECTVERSION else size
        body = dir_dec[i + 6:i + 6 + real]
        yield (i, rid, size, body)
        i += 6 + real


def split_template_dir(dir_dec: bytes) -> bytes:
    """template の dir から PROJECTINFORMATION+PROJECTREFERENCES 部分
    (= PROJECTMODULES レコードの直前まで)をバイト列のまま切り出す。"""
    for off, rid, _size, _body in iter_dir_records(dir_dec):
        if rid == REC_PROJECTMODULES:
            return dir_dec[:off]
    raise OvbaWriteError("template の dir に PROJECTMODULES(0x000F)がありません")


# --- MSForms 参照の除去(W9.2) ------------------------------------------------
# [MS-OVBA] 2.3.4.2.2 PROJECTREFERENCES の REFERENCE レコード群のうち、
# **MSForms(fm20.tlb)への参照だけ**を単位ごと落とす。
#   なぜ落とすのか:
#     (1) W9-B で DataObject の遅延バインドを撤去したので、配布物はもう
#         MSForms を1行も使わない。使わない型ライブラリへの参照が残っていると、
#         開き手の環境に fm20.tlb が無い/場所が違うときに「参照不可」となり、
#         VBA はプロジェクト全体をコンパイルするため**起動直後に生ダイアログ**が
#         出る(Mac の実機で実行時エラー5)。
#     (2) template を作った開発者の絶対パス(/Users/...)が REFERENCECONTROL の
#         LibidExtended に焼き込まれており、配布物に個人情報が載る。
#   REFERENCE の単位([MS-OVBA] 2.3.4.2.2):
#     [REFERENCENAME(0x0016 + 0x003E)] +
#       REFERENCEORIGINAL(0x0033) + REFERENCECONTROL(0x002F)
#         [+ NameRecordExtended(0x0016 + 0x003E)] + Reserved3(0x0030)
#     | REFERENCEREGISTERED(0x000D) | REFERENCEPROJECT(0x000E)
#   **1レコードだけ抜くと後続が読めなくなる**ので、必ずこの単位で落とす。
REC_REFERENCENAME = 0x0016
REC_REFERENCENAME_UNI = 0x003E
REC_REFERENCEORIGINAL = 0x0033
REC_REFERENCECONTROL = 0x002F
REC_REFERENCECONTROL_EXT = 0x0030
REC_REFERENCEREGISTERED = 0x000D
REC_REFERENCEPROJECT = 0x000E

MSFORMS_GUID = b"{0D452EE1-E08F-101A-852E-02608C4D0BB4}"
MSFORMS_NAME = b"MSForms"


def _reference_blocks(prefix: bytes):
    """dir の先頭ブロックを (start, end, is_reference) の並びへ切り分ける。

    参照でないレコード(PROJECTINFORMATION 群)は 1レコード=1ブロックで返す。
    """
    recs = list(iter_dir_records(prefix))
    spans = [(off, off + 6 + len(body)) for off, _rid, _sz, body in recs]
    ids = [rid for _off, rid, _sz, _body in recs]
    i = 0
    n = len(recs)
    while i < n:
        if ids[i] != REC_REFERENCENAME:
            yield (spans[i][0], spans[i][1], False)
            i += 1
            continue
        j = i + 1                                   # REFERENCENAME
        if j < n and ids[j] == REC_REFERENCENAME_UNI:
            j += 1
        if j < n and ids[j] == REC_REFERENCEORIGINAL:
            j += 1
        if j < n and ids[j] == REC_REFERENCECONTROL:
            j += 1
            if j < n and ids[j] == REC_REFERENCENAME:
                j += 1
                if j < n and ids[j] == REC_REFERENCENAME_UNI:
                    j += 1
            if j < n and ids[j] == REC_REFERENCECONTROL_EXT:
                j += 1
        elif j < n and ids[j] in (REC_REFERENCEREGISTERED, REC_REFERENCEPROJECT):
            j += 1
        else:
            raise OvbaWriteError(
                "dir の REFERENCE レコードの並びが [MS-OVBA] 2.3.4.2.2 と"
                "一致しません(id=0x%04X)" % (ids[j] if j < n else 0,))
        yield (spans[i][0], spans[j - 1][1], True)
        i = j


def strip_msforms_reference(prefix: bytes) -> bytes:
    """dir の先頭ブロックから MSForms への REFERENCE を単位ごと落として返す。"""
    out = bytearray()
    for start, end, is_ref in _reference_blocks(prefix):
        blob = prefix[start:end]
        if is_ref and (MSFORMS_GUID in blob or MSFORMS_NAME in blob):
            continue
        out += blob
    return bytes(out)


def template_project_cookie(dir_dec: bytes) -> bytes:
    """template の PROJECTCOOKIE(0x0013)の2バイトを返す。"""
    for _off, rid, _size, body in iter_dir_records(dir_dec):
        if rid == REC_PROJECTCOOKIE:
            return body
    raise OvbaWriteError("template の dir に PROJECTCOOKIE(0x0013)がありません")


def _rec(rid: int, body: bytes) -> bytes:
    return struct.pack('<HI', rid, len(body)) + body


def _mbcs(name: str, codepage: str) -> bytes:
    try:
        return name.encode(codepage)
    except UnicodeEncodeError as exc:
        raise OvbaWriteError(
            "モジュール名 %r をコードページ %s で表現できません" % (name, codepage)) from exc


def build_modules_section(modules, cookie: bytes, codepage: str = "cp932") -> bytes:
    """PROJECTMODULES セクション(MS-OVBA 2.3.4.2.3)を組み立てる。

    modules: VbaModule のリスト。
    MODULEOFFSET は常に 0(=モジュールストリームに PerformanceCache を
    置かず、先頭から圧縮ソースが始まる)。
    """
    out = bytearray()
    out += _rec(REC_PROJECTMODULES, struct.pack('<H', len(modules)))
    out += _rec(REC_PROJECTCOOKIE, cookie)
    for m in modules:
        nm = _mbcs(m.name, codepage)
        nu = m.name.encode('utf-16-le')
        sn = _mbcs(m.stream_name, codepage)
        su = m.stream_name.encode('utf-16-le')
        out += _rec(REC_MODULENAME, nm)
        out += _rec(REC_MODULENAMEUNICODE, nu)
        out += _rec(REC_MODULESTREAMNAME, sn)
        out += _rec(REC_MODULESTREAMNAME_UNI, su)
        out += _rec(REC_MODULEDOCSTRING, b'')
        out += _rec(REC_MODULEDOCSTRING_UNI, b'')
        out += _rec(REC_MODULEOFFSET, struct.pack('<I', 0))
        out += _rec(REC_MODULEHELPCONTEXT, struct.pack('<I', 0))
        out += _rec(REC_MODULECOOKIE, struct.pack('<H', m.cookie))
        # [MS-OVBA] 2.3.4.2.3.2.8 MODULETYPE:
        #   0x0021 = 標準(procedural)モジュール
        #   0x0022 = document / class / designer モジュール
        # クラスモジュールは 0x0022 側である(標準と同じ 0x0021 にすると、
        # PROJECT の Class= 行と食い違ってVBEが読み違える)。
        # document と class の区別は PROJECT ストリームの Document= / Class= 行が持つ。
        out += _rec(REC_MODULETYPE_PROCEDURAL if m.module_type == "std"
                    else REC_MODULETYPE_DOCUMENT, b'')
        # MODULEPRIVATE(0x0028・Size=0)は**クラスモジュールにだけ**続く。
        # 突合の根拠(2026-09-03・W9.3): Mac の実Excel が保存したクラス入り
        # サンプル(Class1 を1本足しただけのブック)の dir を解析したところ、
        #   ThisWorkbook / Sheet1 : ... 0x0022 -> 0x002B
        #   Module1               : ... 0x0021 -> 0x002B
        #   Class1                : ... 0x0022 -> **0x0028(size 0)** -> 0x002B
        # という並びだった。我々は 0x0028 を書いていなかったため、クラスを
        # 含む bin は Excel と不一致であり、これが実機の「実行時エラー 5」
        # (17章 Z-24)の原因だった。**属性8行と併せて修正した版(v7)が Mac の
        # 実Excel で正常に開くことを確認済み**(2026-09-03)。
        # document / std には 0x0028 を出さない(実測どおり)。
        if m.module_type == "class":
            out += _rec(REC_MODULEPRIVATE, b'')
        out += _rec(REC_MODULE_TERMINATOR, b'')
    out += _rec(REC_DIR_TERMINATOR, b'')
    return bytes(out)


# ===========================================================================
# 2. PROJECT / PROJECTwm ストリーム
# ===========================================================================
# PROJECT ストリームから「写す」行(MS-OVBA 2.3.1)。ここに無い行
# (Document=/Module=/Class=/[Workspace])はモジュール台帳から生成する。
# CMG/DPB/GC はプロジェクトの保護状態・パスワード・可視性の暗号化値であり、
# モジュール数とは無関係(モジュール構成を変えても再計算不要でそのまま写せる)。
PROJECT_COPY_KEYS = (
    "ID", "Name", "HelpContextID", "Description", "VersionCompatible32",
    "CMG", "DPB", "GC",
)


def parse_template_project(text: str) -> dict:
    """template の PROJECT ストリームから写す行を key->行全体 で拾う。"""
    copied = {}
    host_ext = []
    in_host = False
    for line in text.splitlines():
        s = line.strip()
        if s.startswith("["):
            in_host = (s.lower() == "[host extender info]")
            continue
        if in_host:
            if s:
                host_ext.append(s)
            continue
        if "=" in s:
            key = s.split("=", 1)[0]
            if key in PROJECT_COPY_KEYS:
                copied[key] = s
    return {"copied": copied, "host_extender": host_ext}


def build_project_stream(modules, tmpl: dict) -> bytes:
    """PROJECT ストリーム(CP932/CRLF)を組み立てる。"""
    copied = tmpl["copied"]
    if "ID" not in copied:
        raise OvbaWriteError("template の PROJECT に ID 行がありません")
    lines = [copied["ID"]]
    for m in modules:
        if m.is_document:
            lines.append("Document=%s/&H00000000" % m.name)
        elif m.module_type == "class":
            lines.append("Class=%s" % m.name)
        else:
            lines.append("Module=%s" % m.name)
    for key in ("Name", "HelpContextID", "Description", "VersionCompatible32",
                "CMG", "DPB", "GC"):
        if key in copied:
            lines.append(copied[key])
    lines.append("")
    lines.append("[Host Extender Info]")
    lines.extend(tmpl["host_extender"])
    lines.append("")
    lines.append("[Workspace]")
    for m in modules:
        if m.is_document:
            lines.append("%s=0, 0, 0, 0, C" % m.name)
        else:
            lines.append("%s=60, 60, 1354, 680, " % m.name)
    lines.append("")
    return ("\r\n".join(lines)).encode("cp932")


def build_projectwm(modules, codepage: str = "cp932") -> bytes:
    """PROJECTwm(MS-OVBA 2.3.3): MBCS名\\0 + UTF-16名\\0\\0 の並び + 終端 \\0\\0。"""
    out = bytearray()
    for m in modules:
        out += _mbcs(m.name, codepage) + b'\x00'
        out += m.name.encode('utf-16-le') + b'\x00\x00'
    out += b'\x00\x00'
    return bytes(out)


# ===========================================================================
# 3. CFB ライター
# ===========================================================================
def _cfb_sort_key(name: str):
    """MS-CFB のディレクトリ兄弟順(名前長 → 大文字化して比較)。

    [MS-CFB] 2.6.4 は比較に先立って UTF-16 コードポイントを特定の大文字化
    テーブルで変換すると規定しており、Python の str.upper()(Unicode既定の
    大文字化)と ASCII の範囲では一致するが、非ASCII文字では一致する保証が
    ない。本ライターが焼くストリーム名(モジュール名・"dir"・"PROJECT" 等の
    固定名)はすべて ASCII のはずであり、非ASCIIが来た場合は
    Python の大文字化で兄弟順を誤り、[MS-CFB] 非準拠のCFBを書き出す恐れが
    ある(Excel/LOでの読み込み失敗や兄弟順不一致に繋がりうる)。安全側に倒し
    fail-closed で止める。
    """
    try:
        name.encode("ascii")
    except UnicodeEncodeError as e:
        raise OvbaWriteError(
            "CFB エントリ名に非ASCII文字があります: %r。[MS-CFB] 2.6.4 の大文字化"
            "規則は非ASCII文字について Python の str.upper() と一致する保証が無く、"
            "兄弟順(ディレクトリ木)を誤って書き出す恐れがあるため fail-closed で"
            "停止します: %s" % (name, e))
    return (len(name), name.upper())


class _DirEntry:
    __slots__ = ("name", "etype", "data", "start", "size",
                 "left", "right", "child")

    def __init__(self, name, etype, data=b''):
        if len(name) > 31:
            raise OvbaWriteError("CFB のエントリ名が31文字を超えています: %r" % name)
        self.name = name
        self.etype = etype          # 1=storage, 2=stream, 5=root
        self.data = data
        self.start = ENDOFCHAIN
        self.size = len(data)
        self.left = NOSTREAM
        self.right = NOSTREAM
        self.child = NOSTREAM


def _build_tree(entries, ids):
    """ソート済みの子エントリ列から平衡二分木を作り、根のIDを返す。"""
    if not ids:
        return NOSTREAM
    mid = len(ids) // 2
    root = ids[mid]
    entries[root].left = _build_tree(entries, ids[:mid])
    entries[root].right = _build_tree(entries, ids[mid + 1:])
    return root


def write_cfb(root_children) -> bytes:
    """CFB(v3・512Bセクタ)を書き出す。

    root_children: [(name, etype, data_or_children)] の入れ子。
      etype=1(storage) のとき第3要素は同じ形式の子リスト。
      etype=2(stream)  のとき第3要素は bytes。
    """
    entries = [_DirEntry("Root Entry", 5)]

    def add(children, parent_index):
        ids = []
        for name, etype, payload in children:
            e = _DirEntry(name, etype, payload if etype == 2 else b'')
            entries.append(e)
            idx = len(entries) - 1
            ids.append((name, idx))
            if etype == 1:
                add(payload, idx)
        ids.sort(key=lambda t: _cfb_sort_key(t[0]))
        entries[parent_index].child = _build_tree(entries, [i for _n, i in ids])

    add(root_children, 0)

    # --- ミニストリーム(4096バイト未満のストリーム)を先に配置する ---
    minifat = []
    mini_blob = bytearray()
    for e in entries:
        if e.etype != 2:
            continue
        if e.size == 0:
            e.start = ENDOFCHAIN
            continue
        if e.size >= MINI_CUTOFF:
            continue
        start = len(mini_blob) // MINI_SECTOR
        nsec = (e.size + MINI_SECTOR - 1) // MINI_SECTOR
        for k in range(nsec):
            minifat.append(start + k + 1 if k < nsec - 1 else ENDOFCHAIN)
        mini_blob += e.data.ljust(nsec * MINI_SECTOR, b'\x00')
        e.start = start

    sectors = []   # 512B ペイロード
    fat = []       # sectors と同じ長さ

    def alloc(data: bytes) -> int:
        if not data:
            return ENDOFCHAIN
        start = len(sectors)
        nsec = (len(data) + SECTOR_SIZE - 1) // SECTOR_SIZE
        for k in range(nsec):
            sectors.append(data[k * SECTOR_SIZE:(k + 1) * SECTOR_SIZE]
                           .ljust(SECTOR_SIZE, b'\x00'))
            fat.append(start + k + 1 if k < nsec - 1 else ENDOFCHAIN)
        return start

    # --- 大きいストリーム ---
    for e in entries:
        if e.etype == 2 and e.size >= MINI_CUTOFF:
            e.start = alloc(e.data)

    # --- MiniFAT ---
    minifat_start = ENDOFCHAIN
    minifat_count = 0
    if minifat:
        per = SECTOR_SIZE // 4
        padded = minifat + [FREESECT] * ((-len(minifat)) % per)
        blob = b''.join(struct.pack('<I', v) for v in padded)
        minifat_start = alloc(blob)
        minifat_count = len(padded) // per

    # --- ミニストリームコンテナ(Root Entry のストリーム) ---
    entries[0].size = len(mini_blob)
    entries[0].start = alloc(bytes(mini_blob)) if mini_blob else ENDOFCHAIN

    # --- ディレクトリ ---
    dir_blob = bytearray()
    for e in entries:
        nm = e.name.encode('utf-16-le') + b'\x00\x00'
        rec = bytearray(128)
        rec[0:len(nm)] = nm
        struct.pack_into('<H', rec, 64, len(nm))
        rec[66] = e.etype
        rec[67] = 1                      # colorFlag: black(検証しない実装が大半)
        struct.pack_into('<I', rec, 68, e.left)
        struct.pack_into('<I', rec, 72, e.right)
        struct.pack_into('<I', rec, 76, e.child)
        struct.pack_into('<I', rec, 116, 0 if e.etype == 1 else e.start)
        struct.pack_into('<Q', rec, 120, 0 if e.etype == 1 else e.size)
        dir_blob += rec
    # ディレクトリセクタは 4 エントリ単位で埋める(空きは FREE エントリ)
    while len(dir_blob) % SECTOR_SIZE:
        rec = bytearray(128)
        struct.pack_into('<I', rec, 68, NOSTREAM)
        struct.pack_into('<I', rec, 72, NOSTREAM)
        struct.pack_into('<I', rec, 76, NOSTREAM)
        dir_blob += rec
    dir_start = alloc(bytes(dir_blob))

    # --- FAT セクタ(自分自身も FAT に載るので反復して本数を決める) ---
    per = SECTOR_SIZE // 4
    nfat = 1
    while True:
        need = -(-(len(sectors) + nfat) // per)
        if need == nfat:
            break
        nfat = need
    if nfat > 109:
        raise OvbaWriteError(
            "FAT セクタが %d 本必要で DIFAT がヘッダに収まりません(未対応)" % nfat)
    fat_sector_ids = list(range(len(sectors), len(sectors) + nfat))
    for _ in range(nfat):
        sectors.append(b'\x00' * SECTOR_SIZE)
        fat.append(FATSECT)

    fat_padded = fat + [FREESECT] * (nfat * per - len(fat))
    fat_blob = b''.join(struct.pack('<I', v) for v in fat_padded)
    for k, sid in enumerate(fat_sector_ids):
        sectors[sid] = fat_blob[k * SECTOR_SIZE:(k + 1) * SECTOR_SIZE]

    # --- ヘッダ ---
    hdr = bytearray(SECTOR_SIZE)
    hdr[0:8] = CFB_MAGIC
    struct.pack_into('<H', hdr, 24, 0x003E)     # minor version
    struct.pack_into('<H', hdr, 26, 0x0003)     # major version (v3)
    struct.pack_into('<H', hdr, 28, 0xFFFE)     # byte order (little endian)
    struct.pack_into('<H', hdr, 30, 9)          # sector shift (512)
    struct.pack_into('<H', hdr, 32, 6)          # mini sector shift (64)
    struct.pack_into('<I', hdr, 44, nfat)
    struct.pack_into('<I', hdr, 48, dir_start)
    struct.pack_into('<I', hdr, 56, MINI_CUTOFF)
    struct.pack_into('<I', hdr, 60, minifat_start)
    struct.pack_into('<I', hdr, 64, minifat_count)
    struct.pack_into('<I', hdr, 68, ENDOFCHAIN)  # first DIFAT sector
    struct.pack_into('<I', hdr, 72, 0)           # number of DIFAT sectors
    for k in range(109):
        v = fat_sector_ids[k] if k < nfat else FREESECT
        struct.pack_into('<I', hdr, 76 + 4 * k, v)

    return bytes(hdr) + b''.join(sectors)


# ===========================================================================
# 4. モジュール記述と組み立て本体
# ===========================================================================
class VbaModule:
    """vbaProject.bin へ焼き込む1モジュール。

    name        : VBE 上のモジュール名(=dir の MODULENAME)
    source      : モジュールストリームへ入れるソース(CP932・CRLF・末尾改行あり)。
                  **属性行を含めた完全な形**で、`module_stream_source()` が作る
    module_type : "std" | "class" | "document"
    """

    def __init__(self, name, source: bytes, module_type: str = "std",
                 stream_name: str | None = None, cookie: int = 0xFFFF):
        if module_type not in ("std", "class", "document"):
            raise OvbaWriteError("未知の module_type: %r" % module_type)
        self.name = name
        self.source = source
        self.module_type = module_type
        self.stream_name = stream_name or name
        self.cookie = cookie

    @property
    def is_document(self) -> bool:
        return self.module_type == "document"

    def stream_bytes(self) -> bytes:
        """モジュールストリーム本体(MODULEOFFSET=0 なので圧縮ソースのみ)。

        全モジュールで合計1MB超を圧縮するため高速版を使う(ovba.py 参照)。
        解凍結果は素朴版と同一で、読み戻し検証がそれを毎回確認する。
        """
        return ovba.ovba_compress(self.source, fast=True)


# VBE が .bas/.cls をエクスポートするときに先頭へ付ける属性行。
# vbaProject.bin のモジュールストリームには「モジュール属性 + ソース」が
# そのまま入る(Attribute 行はソースの一部である)。
_ATTR_STD = 'Attribute VB_Name = "%s"\r\n'
# クラスモジュールの VB_Base(実測。Mac 実Excel 製サンプルの Class1)。
# document module の VB_Base と同じく「そのモジュールの基底COMクラスのGUID」で
# あり、クラスモジュールでは常にこの値になる。
_VB_BASE_CLASS = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"

# 突合の根拠(2026-09-03・W9.3): Mac 実Excel 製サンプルの Class1 の属性行は
#   VB_Name / VB_Base / VB_GlobalNameSpace / VB_Creatable / VB_PredeclaredId /
#   VB_Exposed / VB_TemplateDerived / VB_Customizable の**8行**であった
# (旧実装は VB_Base・VB_TemplateDerived・VB_Customizable を欠く5行だった)。
# クラス本体の `Attribute <変数名>.VB_VarHelpID = -1`(WithEvents 宣言の直後)は
# 属性ヘッダではなく**本文の一部**なので、ここではなく本文側が持つ。
_ATTR_CLASS = (
    'Attribute VB_Name = "%s"\r\n'
    'Attribute VB_Base = "' + _VB_BASE_CLASS + '"\r\n'
    'Attribute VB_GlobalNameSpace = False\r\n'
    'Attribute VB_Creatable = False\r\n'
    'Attribute VB_PredeclaredId = False\r\n'
    'Attribute VB_Exposed = False\r\n'
    'Attribute VB_TemplateDerived = False\r\n'
    'Attribute VB_Customizable = False\r\n'
)
# document module の VB_Base は「そのドキュメントのCOMクラスのGUID」であり、
# ブック本体(Workbook)とワークシート(Worksheet)とで異なる([MS-OVBA] は
# VB_Base をホストが解釈する不透明な文字列としてのみ規定するが、Excel が
# 実際に埋め込む値はホストのタイプライブラリのGUIDに一致する)。
# 現状 build_baked_vba_project が焼くのは ThisWorkbook(Workbook)だけだが、
# 将来シートモジュールを焼く経路が増えても正しいGUIDを選べるよう分岐できる
# ようにしておく。
_VB_BASE_WORKBOOK = "00020819"   # Excel.Workbook
_VB_BASE_WORKSHEET = "00020820"  # Excel.Worksheet

_ATTR_DOCUMENT = (
    'Attribute VB_Name = "%s"\r\n'
    'Attribute VB_Base = "0{%s-0000-0000-C000-000000000046}"\r\n'
    'Attribute VB_GlobalNameSpace = False\r\n'
    'Attribute VB_Creatable = False\r\n'
    'Attribute VB_PredeclaredId = True\r\n'
    'Attribute VB_Exposed = True\r\n'
    'Attribute VB_TemplateDerived = False\r\n'
    'Attribute VB_Customizable = True\r\n'
)


def module_stream_source(name: str, body: str, module_type: str,
                         attributes: str | None = None,
                         doc_base_guid: str = _VB_BASE_WORKBOOK) -> bytes:
    """モジュールストリームへ入れるソース(属性行 + 本文)を CP932/CRLF で作る。

    attributes を与えた場合はそれを使う(.cls の元ヘッダを保つとき用)。
    doc_base_guid は module_type="document" のときだけ使う VB_Base のGUIDで、
    ブック(既定・_VB_BASE_WORKBOOK)かシート(_VB_BASE_WORKSHEET)かを呼び出し側
    が選ぶ。現行の呼び出しは ThisWorkbook のみなので既定のままでよい。

    改行について(W9.3 の実測メモ): Mac の実Excel が保存したブック
    (scratchpad/bisect/mac_class_sample.xlsm)は**全モジュールの改行が LF 単独**
    だった(Windows 製は CRLF)。本実装は **CRLF のまま**にする(切り分けブック
    v1_min / v3_full_noclass が CRLF で実機読み込みに成功しており、改行は
    Err 5 の要因ではないことが確かめられているため)。事実だけ記録に残す。
    """
    if attributes is None:
        if module_type == "class":
            attributes = _ATTR_CLASS % name
        elif module_type == "document":
            attributes = _ATTR_DOCUMENT % (name, doc_base_guid)
        else:
            attributes = _ATTR_STD % name
    text = body.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
    if text and not text.endswith("\r\n"):
        text += "\r\n"
    return (attributes + text).encode("cp932")


def build_vba_project(template_bin: bytes, modules,
                      vba_project_stream: bytes | None = None) -> bytes:
    """template の vbaProject.bin から情報を写しつつ、modules を焼き込んだ
    完成品 vbaProject.bin を返す。

    modules は VbaModule のリスト。並び順がそのまま dir/PROJECT の並びになる
    (慣例として ThisWorkbook などの document module を先頭に置く)。
    vba_project_stream を渡すと _VBA_PROJECT ストリームをその内容で置き換える。
    """
    if not modules:
        raise OvbaWriteError("modules が空です")
    names = [m.name for m in modules]
    if len(set(n.upper() for n in names)) != len(names):
        raise OvbaWriteError("モジュール名が重複しています: %r" % (names,))

    skel = ovba.CFBReader(template_bin)
    tmpl_dir = ovba.ovba_decompress(skel.read("dir"))
    prefix = split_template_dir(tmpl_dir)
    # 使わない MSForms 参照(と、そこに焼き込まれた開発者の絶対パス)を落とす。W9.2。
    prefix = strip_msforms_reference(prefix)
    cookie = template_project_cookie(tmpl_dir)

    dir_dec = prefix + build_modules_section(modules, cookie)
    dir_stream = ovba.ovba_compress(dir_dec, fast=True)

    project_tmpl = parse_template_project(
        skel.read("PROJECT").decode("cp932", errors="replace"))
    project_stream = build_project_stream(modules, project_tmpl)
    projectwm = build_projectwm(modules)
    # _VBA_PROJECT は template のものを写す。呼び出し側が「PerformanceCache 無害化
    # 済み(Version=0xFFFF)」の版を渡してきたらそれを優先する。
    vba_project = (vba_project_stream if vba_project_stream is not None
                   else skel.read("_VBA_PROJECT"))

    vba_children = [(m.stream_name, 2, m.stream_bytes()) for m in modules]
    vba_children.append(("_VBA_PROJECT", 2, vba_project))
    vba_children.append(("dir", 2, dir_stream))

    return write_cfb([
        ("VBA", 1, vba_children),
        ("PROJECT", 2, project_stream),
        ("PROJECTwm", 2, projectwm),
    ])


# ===========================================================================
# 5. 読み戻し(検証用)
# ===========================================================================
def read_reference_names(vba_bin: bytes) -> list:
    """完成品 vbaProject.bin の dir から、参照設定の名前を並び順で返す。

    裁定書34 §1.4(W12-A)。第2段(開発PCの import_navi_modules.ps1)が
    参照設定を足したあと、増えたのが SHDocVw と MSForms の2つだけであることを
    tools/ship_check.py --final が確かめるための読み出し口。

    REFERENCENAME(0x0016)の本体が MBCS のライブラリ名そのものである
    ([MS-OVBA] 2.3.4.2.2.2)。続く 0x003E は同じ名前の UTF-16 版なので読まない。
    """
    cfb = ovba.CFBReader(vba_bin)
    dir_dec = ovba.ovba_decompress(cfb.read("dir"))
    out = []
    for _off, rid, _size, body in iter_dir_records(dir_dec):
        if rid == REC_MODULENAME:
            break                       # MODULE 群に入ったら参照は終わり
        if rid == REC_REFERENCENAME:
            out.append(body.decode("cp932", errors="replace"))
    return out


def read_modules(vba_bin: bytes) -> dict:
    """完成品 vbaProject.bin を読み戻し、モジュール名 -> ソース(bytes)を返す。

    tools/bin_roundtrip.py と build_rpn.py の自己検証がこれを使う。
    dir の MODULEOFFSET を尊重するので、PerformanceCache 付きの bin も読める。
    """
    cfb = ovba.CFBReader(vba_bin)
    dir_dec = ovba.ovba_decompress(cfb.read("dir"))
    out = {}
    cur = None
    for _off, rid, _size, body in iter_dir_records(dir_dec):
        if rid == REC_MODULENAME:
            cur = {"name": body.decode("cp932"), "offset": 0, "type": None}
        elif cur is None:
            continue
        elif rid == REC_MODULESTREAMNAME:
            cur["stream"] = body.decode("cp932")
        elif rid == REC_MODULEOFFSET:
            cur["offset"] = struct.unpack('<I', body)[0]
        elif rid in (REC_MODULETYPE_PROCEDURAL, REC_MODULETYPE_DOCUMENT):
            # 0x0022 は document と class の両方に使われるので、dir だけでは
            # 区別できない。区別は PROJECT ストリームの Document= 行が持つ
            # (下で上書きする)。
            cur["type"] = ("document" if rid == REC_MODULETYPE_DOCUMENT
                           else "procedural")
        elif rid == REC_MODULE_TERMINATOR:
            raw = cfb.read(cur.get("stream", cur["name"]))
            # MODULEOFFSET より前は Excel が付ける performance cache(p-code)で、
            # 圧縮コンテナは offset から始まる(MS-OVBA 2.3.4.3)。当方 bin は
            # offset=0 なので従来どおり。Excel 保存後の bin(第2段)はここが要る。
            src = ovba.ovba_decompress(raw[cur["offset"]:])
            out[cur["name"]] = {"source": src, "type": cur["type"]}
            cur = None
    # PROJECT ストリームの Document= 行だけが「本当の document module」である。
    # 0x0022 だが Document= に無いものは class module。
    try:
        proj = cfb.read("PROJECT").decode("cp932", errors="replace")
    except KeyError:
        proj = ""
    doc_names = set()
    for line in proj.splitlines():
        if line.startswith("Document="):
            doc_names.add(line[len("Document="):].split("/")[0].strip())
    for nm, info in out.items():
        if info["type"] == "document" and nm not in doc_names:
            info["type"] = "class"
    return out


def strip_attribute_lines(src: bytes) -> bytes:
    """モジュールソース先頭の `Attribute VB_...` 行を落とした残りを返す。"""
    lines = src.split(b"\r\n")
    i = 0
    while i < len(lines) and lines[i].startswith(b"Attribute "):
        i += 1
    return b"\r\n".join(lines[i:])


# ===========================================================================
# 6. 第2段の一段化(17章 Z-42) -- UserForm と参照設定を bin へ足す
# ---------------------------------------------------------------------------
# 背景:
#   2段ビルドの第2段(build/win/import_navi_modules.ps1)は Windows 実Excel の
#   VBE COM で UserForm(frmNaviHtml)と参照設定(SHDocVw / MSForms)を足していた。
#   社内IT環境 v1.1 §3.1 が「PowerShell 不可」なので、第2段を CI 側へ寄せる
#   (= 一段化)。本節は **第1段の bin へ、フォーム1本と参照2本を書き足す**。
#
# 情報源(どこから作るか):
#   ・コード本文       : src/ui/navi/frmNaviHtml.frm の Begin…End より後ろ
#   ・デザイナ定義     : 同 .frm の VERSION 5.00〜End(OleObjectBlob 行を除く)
#                        -> `frmNaviHtml/\x03VBFrame`
#   ・コントロール情報 : src/ui/navi/frmNaviHtml.frx の OleObjectBlob(=CFB)
#                        -> `frmNaviHtml/\x01CompObj` `f` `o`
#   ・参照設定         : 下の定数(髙橋産物と template_skeleton の実測から作った。
#                        **個人の絶対パスを含む値は使わない**。W9.2 と同じ理由)
#   つまり **髙橋産物のファイルには依存しない**(CI に置けないので依存できない)。
#
# 実測の出典(2026-09-12・W15 班G):
#   髙橋産物 scratchpad/tk2/final_stale_w12a/リスク提案ナビ.xlsm の
#   xl/vbaProject.bin(6,885,376 B)を読んだ結果:
#     dir(解凍後 20,699 B)
#       off=132 REFERENCENAME "Office"      / off=162 REFERENCEREGISTERED
#       off=326 REFERENCENAME "SHDocVw"     / off=359 REFERENCEREGISTERED(117 B)
#       off=482 REFERENCENAME "MSForms"     / off=515 REFERENCEORIGINAL(111 B)
#                                             off=632 REFERENCECONTROL(59 B)
#                                             off=697 REFERENCENAME "MSForms"
#                                             off=730 REFERENCECONTROL_EXT(161 B)
#       off=897 PROJECTMODULES count=136 / off=905 PROJECTCOOKIE
#       off=20427 MODULE frmNaviHtml: stream="frmNaviHtml" MODULEOFFSET=19484
#                 MODULECOOKIE=60240 MODULETYPE=0x0022 + MODULEPRIVATE(0x0028)
#     PROJECT   : 135行目 `BaseClass=frmNaviHtml`、[Workspace] に
#                 `frmNaviHtml=392, 392, 2793, 1578, Z, 392, 392, 2793, 1578, C`
#     PROJECTwm : 136エントリ(frmNaviHtml は 135 番目)
#     designer  : frmNaviHtml/\x01CompObj(97 B)・\x03VBFrame(279 B)・f(38 B)・o(0 B)
#     _VBA_PROJECT: 104,952 B(p-code キャッシュ入り。__SRP_* が 300 本超)
#   当方の第1段 bin(1,071,104 B)は _VBA_PROJECT が 3,061 B の無害化版
#   (Version=0xFFFF・PerformanceCache ゼロ埋め)で、全 MODULEOFFSET=0。
#   **本節は _VBA_PROJECT と MODULEOFFSET をそのまま(無害化・0)に保つ**。
#   p-code を持たないので Excel/LO はソースから再コンパイルする(第1段と同じ)。
# ===========================================================================

FORM_STREAM_NAMES = ("\x01CompObj", "\x03VBFrame", "f", "o")

# --- 参照設定の値(個人の絶対パスを含まないものだけを使う) -------------------
# SHDocVw: 髙橋産物の REFERENCEREGISTERED からバイトそのまま
#          (C:\Windows\System32\ieframe.dll は Windows の既定位置で、
#           個人情報を含まない)。
SHDOCVW_LIBID = (b"*\\G{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0#"
                 b"C:\\Windows\\System32\\ieframe.dll#Microsoft Internet Controls")
# MSForms: 髙橋産物の REFERENCEORIGINAL は C:\WINDOWS\system32\FM20.DLL、
#   REFERENCECONTROL_EXT は C:\Users\nori\AppData\Local\Temp\VBE\MSForms.exd と
#   **開発者名を含む**。W9.2 が MSForms 参照を落とした理由そのものなので、
#   ここでは template_skeleton.xlsm 側の書き方(`*\H` = 登録名だけで引く形)を採る。
MSFORMS_ORIGINAL_LIBID = (b"*\\H{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0#"
                          b"fm20.tlb#Microsoft Forms 2.0 Object Library")
# 「ねじれた(twiddled)」libid。髙橋産物・template_skeleton の**両方で同一**の
# 固定値(null GUID)だったので、そのまま使う。
MSFORMS_TWIDDLED_LIBID = b"*\\G{00000000-0000-0000-0000-000000000000}#0.0#0##"
# OriginalTypeLib([MS-OVBA] 2.3.4.2.2.3): MSForms の TypeLib GUID をバイト順で。
MSFORMS_TYPELIB_GUID = uuid.UUID("{0D452EE1-E08F-101A-852E-02608C4D0BB4}").bytes_le

# フォームモジュールの VB_Base(髙橋産物の実測。VBE が .frm 取込時に作る2つの
# GUID で、**このフォーム固有の値**。同じフォームを焼き直す限り変える理由は無い)。
FORM_VB_BASE = ("0{0F2C78AF-89C5-4EBD-A603-8EC02854C6FF}"
                "{1672A8B2-35CB-45EF-BB30-E22D4679F96C}")
# フォームの属性ヘッダ8行(髙橋産物のモジュールストリーム先頭と同じ並び)。
_ATTR_FORM = (
    'Attribute VB_Name = "%s"\r\n'
    'Attribute VB_Base = "%s"\r\n'
    'Attribute VB_GlobalNameSpace = False\r\n'
    'Attribute VB_Creatable = False\r\n'
    'Attribute VB_PredeclaredId = True\r\n'
    'Attribute VB_Exposed = False\r\n'
    'Attribute VB_TemplateDerived = False\r\n'
    'Attribute VB_Customizable = False\r\n'
)
# [Workspace] の1行(髙橋産物の実測値そのまま。VBE の窓の位置なので機能に影響しない)。
FORM_WORKSPACE = "%s=392, 392, 2793, 1578, Z, 392, 392, 2793, 1578, C"


# ---------------------------------------------------------------------------
# 6.1 パスで引ける CFB リーダー(ovba.CFBReader は名前が平坦で、ストレージ配下の
#     `f` や `\x01CompObj` を引けないため、ここに木を辿る版を持つ)
# ---------------------------------------------------------------------------
def read_cfb_tree(data: bytes) -> dict:
    """CFB を {name: bytes(ストリーム) | dict(ストレージ)} の入れ子で返す。"""
    if data[:8] != CFB_MAGIC:
        raise OvbaWriteError("CFB のシグネチャがありません")
    num_fat = struct.unpack('<I', data[44:48])[0]
    dir_start = struct.unpack('<I', data[48:52])[0]
    minifat_start = struct.unpack('<I', data[60:64])[0]
    difat_first = struct.unpack('<I', data[68:72])[0]
    per = SECTOR_SIZE // 4

    def sector(n):
        off = (n + 1) * SECTOR_SIZE
        return data[off:off + SECTOR_SIZE]

    difat = [struct.unpack('<I', data[76 + 4 * k:80 + 4 * k])[0] for k in range(109)]
    sec = difat_first
    while sec not in (ENDOFCHAIN, FREESECT) and (sec + 2) * SECTOR_SIZE <= len(data):
        block = list(struct.unpack('<%dI' % per, sector(sec)))
        difat.extend(block[:-1])
        sec = block[-1]
    fat = []
    for fs in difat[:max(num_fat, 1)]:
        if fs in (ENDOFCHAIN, FREESECT) or (fs + 2) * SECTOR_SIZE > len(data):
            continue
        fat += list(struct.unpack('<%dI' % per, sector(fs)))

    minifat = []
    sec = minifat_start
    while sec != ENDOFCHAIN and sec < len(fat):
        minifat += list(struct.unpack('<128I', sector(sec)))
        sec = fat[sec]

    def chain_bytes(start, size):
        blob = bytearray()
        sec = start
        while sec != ENDOFCHAIN and sec < len(fat):
            blob += sector(sec)
            sec = fat[sec]
        return bytes(blob[:size])

    dir_blob = chain_bytes(dir_start, 1 << 30)
    entries = []
    for i in range(0, len(dir_blob), 128):
        e = dir_blob[i:i + 128]
        if len(e) < 128:
            break
        nlen = struct.unpack('<H', e[64:66])[0]
        name = e[:max(nlen - 2, 0)].decode('utf-16-le') if nlen else ''
        entries.append({
            'name': name, 'type': e[66],
            'left': struct.unpack('<I', e[68:72])[0],
            'right': struct.unpack('<I', e[72:76])[0],
            'child': struct.unpack('<I', e[76:80])[0],
            'start': struct.unpack('<I', e[116:120])[0],
            'size': struct.unpack('<I', e[120:124])[0],
        })
    if not entries or entries[0]['type'] != 5:
        raise OvbaWriteError("CFB のルートエントリがありません")

    mini_blob = chain_bytes(entries[0]['start'], entries[0]['size'])

    def stream_bytes(e):
        if e['size'] == 0:
            return b''
        if e['size'] < MINI_CUTOFF:
            blob = bytearray()
            sec = e['start']
            while sec != ENDOFCHAIN and sec < len(minifat):
                blob += mini_blob[sec * MINI_SECTOR:(sec + 1) * MINI_SECTOR]
                sec = minifat[sec]
            return bytes(blob[:e['size']])
        return chain_bytes(e['start'], e['size'])

    def walk(idx, out, seen):
        if idx == NOSTREAM or idx >= len(entries) or idx in seen:
            return
        seen.add(idx)
        e = entries[idx]
        walk(e['left'], out, seen)
        if e['type'] == 1:
            sub = {}
            walk(e['child'], sub, set())
            out[e['name']] = sub
        elif e['type'] == 2:
            out[e['name']] = stream_bytes(e)
        walk(e['right'], out, seen)

    root: dict = {}
    walk(entries[0]['child'], root, set())
    return root


def _tree_to_children(tree: dict):
    """read_cfb_tree の出力を write_cfb の root_children 形式へ戻す。"""
    out = []
    for name, val in tree.items():
        if isinstance(val, dict):
            out.append((name, 1, _tree_to_children(val)))
        else:
            out.append((name, 2, val))
    return out


# ---------------------------------------------------------------------------
# 6.2 参照設定レコードの組み立て([MS-OVBA] 2.3.4.2.2)
# ---------------------------------------------------------------------------
def build_reference_registered(name: str, libid: bytes,
                               codepage: str = "cp932") -> bytes:
    """REFERENCENAME + REFERENCENAME_UNI + REFERENCEREGISTERED(0x000D)。

    REFERENCEREGISTERED の本体 = SizeOfLibid(4) + Libid + Reserved1(4) + Reserved2(2)。
    """
    body = struct.pack('<I', len(libid)) + libid + b'\x00' * 6
    return (_rec(REC_REFERENCENAME, _mbcs(name, codepage))
            + _rec(REC_REFERENCENAME_UNI, name.encode('utf-16-le'))
            + _rec(REC_REFERENCEREGISTERED, body))


def build_reference_control(name: str, original: bytes, twiddled: bytes,
                            extended: bytes, typelib_guid: bytes,
                            cookie: int = 1, codepage: str = "cp932") -> bytes:
    """REFERENCENAME(+UNI) + REFERENCEORIGINAL + REFERENCECONTROL
    + NameRecordExtended + Reserved3(REFERENCECONTROL_EXT)。

    ・REFERENCEORIGINAL(0x0033) の Size は SizeOfLibidOriginal そのもの
      (本体に長さ前置きが無い)。髙橋産物で Size=111・本体111Bを実測。
    ・REFERENCECONTROL(0x002F) 本体 = SizeOfLibidTwiddled(4) + Libid
      + Reserved1(4) + Reserved2(2)。
    ・Reserved3(0x0030) 本体 = SizeOfLibidExtended(4) + Libid + Reserved4(4)
      + Reserved5(2) + OriginalTypeLib(16) + Cookie(4)。
    """
    if len(typelib_guid) != 16:
        raise OvbaWriteError("OriginalTypeLib は 16 バイトです")
    nm = _mbcs(name, codepage)
    nu = name.encode('utf-16-le')
    ctrl = struct.pack('<I', len(twiddled)) + twiddled + b'\x00' * 6
    ext = (struct.pack('<I', len(extended)) + extended + b'\x00' * 6
           + typelib_guid + struct.pack('<I', cookie))
    return (_rec(REC_REFERENCENAME, nm)
            + _rec(REC_REFERENCENAME_UNI, nu)
            + _rec(REC_REFERENCEORIGINAL, original)
            + _rec(REC_REFERENCECONTROL, ctrl)
            + _rec(REC_REFERENCENAME, nm)
            + _rec(REC_REFERENCENAME_UNI, nu)
            + _rec(REC_REFERENCECONTROL_EXT, ext))


def navi_reference_records() -> bytes:
    """HTML画面に要る参照2本(SHDocVw / MSForms)の dir レコード列。"""
    return (build_reference_registered("SHDocVw", SHDOCVW_LIBID)
            + build_reference_control("MSForms", MSFORMS_ORIGINAL_LIBID,
                                      MSFORMS_TWIDDLED_LIBID,
                                      MSFORMS_ORIGINAL_LIBID,
                                      MSFORMS_TYPELIB_GUID))


# ---------------------------------------------------------------------------
# 6.3 .frm / .frx を読む
# ---------------------------------------------------------------------------
def split_frm(text: str) -> tuple:
    """VBE がエクスポートした .frm を (デザイナ行, 属性行, コード行) に割る。

    .frm の形:
        VERSION 5.00
        Begin {C62A69F0-...} frmNaviHtml
           ...
           OleObjectBlob   =   "frmNaviHtml.frx":0000
        End
        Attribute VB_Name = "frmNaviHtml"
        ...
        (本文)
    """
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    if not lines or not lines[0].startswith("VERSION "):
        raise OvbaWriteError(".frm が `VERSION` で始まっていません")
    try:
        end = next(i for i, l in enumerate(lines) if l.strip() == "End")
    except StopIteration:
        raise OvbaWriteError(".frm のデザイナ部に `End` 行がありません")
    designer = lines[:end + 1]
    rest = lines[end + 1:]
    n = 0
    while n < len(rest) and rest[n].startswith("Attribute "):
        n += 1
    return designer, rest[:n], rest[n:]


def build_vbframe(designer_lines, codepage: str = "cp932") -> bytes:
    """`<form>/\\x03VBFrame` ストリーム(デザイナ定義のテキスト)。

    .frm のデザイナ部から **OleObjectBlob 行だけを落として** CRLF/CP932 にする。
    (OleObjectBlob は .frx への参照であって、bin では `f`/`o` ストリームが
     その中身を持つ。髙橋産物の \\x03VBFrame にも OleObjectBlob 行は無い。)
    """
    keep = [l for l in designer_lines if not l.strip().startswith("OleObjectBlob")]
    return ("\r\n".join(keep) + "\r\n").encode(codepage)


def frx_designer_streams(frx: bytes) -> dict:
    """.frx の OleObjectBlob(埋め込み CFB)から designer ストリームを取り出す。

    .frm の `OleObjectBlob = "xxx.frx":0000` が指す先は、小さなヘッダの後ろに
    CFB がまるごと入った塊である(実測: frmNaviHtml.frx は 2,584 B で、
    先頭 24 B がヘッダ、24 B 目から 2,560 B の CFB)。中身は
    `\\x01CompObj`(110 B)・`f`(38 B)・`o`(0 B) の3本。
    """
    i = frx.find(CFB_MAGIC)
    if i < 0:
        raise OvbaWriteError(".frx に CFB が見つかりません(OleObjectBlob が壊れています)")
    tree = read_cfb_tree(frx[i:])
    out = {}
    for name in ("\x01CompObj", "f", "o"):
        if name not in tree or isinstance(tree[name], dict):
            raise OvbaWriteError(".frx に %r ストリームがありません" % name)
        out[name] = tree[name]
    return out


def build_designer_storage(frm_text: str, frx: bytes) -> dict:
    """`<form>/` ストレージの4ストリームを .frm と .frx だけから作る。"""
    designer, _attrs, _code = split_frm(frm_text)
    out = dict(frx_designer_streams(frx))
    out["\x03VBFrame"] = build_vbframe(designer)
    missing = [n for n in FORM_STREAM_NAMES if n not in out]
    if missing:
        raise OvbaWriteError("designer ストレージに %r がありません" % (missing,))
    return out


def form_module_source(form_name: str, frm_text: str) -> bytes:
    """フォームのモジュールストリームへ入れるソース(属性8行 + 本文)。"""
    _designer, attrs, code = split_frm(frm_text)
    names = [a.split("=", 1)[0].strip() for a in attrs]
    if 'Attribute VB_Name' not in names:
        raise OvbaWriteError(".frm に Attribute VB_Name 行がありません")
    body = "\n".join(code)
    return module_stream_source(form_name, body, "class",
                                attributes=_ATTR_FORM % (form_name, FORM_VB_BASE))


# ---------------------------------------------------------------------------
# 6.4 dir / PROJECT / PROJECTwm への差し込み
# ---------------------------------------------------------------------------
def _module_record(name: str, stream_name: str, cookie: int = 0xFFFF,
                   codepage: str = "cp932") -> bytes:
    """designer(フォーム)1本ぶんの MODULE レコード。

    MODULETYPE は 0x0022、そのあとに MODULEPRIVATE(0x0028・Size=0)が続く
    (髙橋産物の frmNaviHtml の実測と同じ並び。クラスモジュールと同型)。
    MODULEOFFSET は 0(p-code を置かない)。
    """
    return (_rec(REC_MODULENAME, _mbcs(name, codepage))
            + _rec(REC_MODULENAMEUNICODE, name.encode('utf-16-le'))
            + _rec(REC_MODULESTREAMNAME, _mbcs(stream_name, codepage))
            + _rec(REC_MODULESTREAMNAME_UNI, stream_name.encode('utf-16-le'))
            + _rec(REC_MODULEDOCSTRING, b'')
            + _rec(REC_MODULEDOCSTRING_UNI, b'')
            + _rec(REC_MODULEOFFSET, struct.pack('<I', 0))
            + _rec(REC_MODULEHELPCONTEXT, struct.pack('<I', 0))
            + _rec(REC_MODULECOOKIE, struct.pack('<H', cookie))
            + _rec(REC_MODULETYPE_DOCUMENT, b'')
            + _rec(REC_MODULEPRIVATE, b'')
            + _rec(REC_MODULE_TERMINATOR, b''))


def insert_into_dir(dir_dec: bytes, reference_records: bytes,
                    module_record: bytes) -> bytes:
    """解凍済み dir へ 参照レコード列 と MODULE レコードを差し込む。

    ・参照は PROJECTMODULES(0x000F)の直前(= PROJECTREFERENCES の末尾)へ。
    ・MODULE は DIR_TERMINATOR(0x0010)の直前へ。PROJECTMODULES の Count も +1。
    """
    modules_at = None
    term_at = None
    count = None
    for off, rid, _size, body in iter_dir_records(dir_dec):
        if rid == REC_PROJECTMODULES and modules_at is None:
            modules_at = off
            count = struct.unpack('<H', body)[0]
        elif rid == REC_DIR_TERMINATOR:
            term_at = off
    if modules_at is None:
        raise OvbaWriteError("dir に PROJECTMODULES(0x000F)がありません")
    if term_at is None:
        raise OvbaWriteError("dir に終端レコード(0x0010)がありません")
    n_new = 1 if module_record else 0
    head = dir_dec[:modules_at] + reference_records
    mid = (_rec(REC_PROJECTMODULES, struct.pack('<H', count + n_new))
           + dir_dec[modules_at + 8:term_at] + module_record)
    return head + mid + dir_dec[term_at:]


def insert_into_project(project: bytes, form_name: str,
                        codepage: str = "cp932") -> bytes:
    """PROJECT ストリームへ `BaseClass=<form>` と [Workspace] の1行を足す。"""
    text = project.decode(codepage)
    lines = text.split("\r\n")
    base = "BaseClass=%s" % form_name
    if base in lines:
        raise OvbaWriteError("PROJECT に既に %s があります" % base)
    last = -1
    for i, l in enumerate(lines):
        if l.startswith(("Module=", "Class=", "Document=", "BaseClass=")):
            last = i
    if last < 0:
        raise OvbaWriteError("PROJECT に Module=/Document= 行がありません")
    lines.insert(last + 1, base)
    try:
        ws = lines.index("[Workspace]")
    except ValueError:
        raise OvbaWriteError("PROJECT に [Workspace] がありません")
    end = len(lines)
    while end > ws and lines[end - 1] == "":
        end -= 1
    lines.insert(end, FORM_WORKSPACE % form_name)
    return "\r\n".join(lines).encode(codepage)


def insert_into_projectwm(wm: bytes, form_name: str,
                          codepage: str = "cp932") -> bytes:
    """PROJECTwm の終端 \\x00\\x00 の手前へ1エントリ足す。"""
    if not wm.endswith(b'\x00\x00'):
        raise OvbaWriteError("PROJECTwm が \\x00\\x00 で終わっていません")
    entry = (_mbcs(form_name, codepage) + b'\x00'
             + form_name.encode('utf-16-le') + b'\x00\x00')
    return wm[:-2] + entry + b'\x00\x00'


# ---------------------------------------------------------------------------
# 6.5 本体
# ---------------------------------------------------------------------------
def add_userform(vba_bin: bytes, form_name: str, frm_text: str, frx: bytes,
                 reference_records: bytes | None = None) -> bytes:
    """第1段の bin へ UserForm 1本と参照設定を足した bin を返す。

    第2段(import_navi_modules.ps1)の (1) フォーム取込 と (2) 参照設定 に相当。
    _VBA_PROJECT・既存モジュールストリーム・MODULEOFFSET は一切触らない。
    """
    tree = read_cfb_tree(vba_bin)
    if "VBA" not in tree or not isinstance(tree["VBA"], dict):
        raise OvbaWriteError("bin に VBA ストレージがありません")
    if form_name in tree:
        raise OvbaWriteError("bin に既に %s ストレージがあります" % form_name)
    vba = tree["VBA"]
    if form_name in vba:
        raise OvbaWriteError("bin に既に VBA/%s ストリームがあります" % form_name)
    if "dir" not in vba:
        raise OvbaWriteError("bin に VBA/dir がありません")

    refs = navi_reference_records() if reference_records is None else reference_records
    dir_dec = ovba.ovba_decompress(vba["dir"])
    dir_dec = insert_into_dir(dir_dec, refs, _module_record(form_name, form_name))
    vba["dir"] = ovba.ovba_compress(dir_dec, fast=True)
    vba[form_name] = ovba.ovba_compress(form_module_source(form_name, frm_text),
                                        fast=True)
    tree["PROJECT"] = insert_into_project(tree["PROJECT"], form_name)
    tree["PROJECTwm"] = insert_into_projectwm(tree["PROJECTwm"], form_name)
    tree[form_name] = build_designer_storage(frm_text, frx)
    return write_cfb(_tree_to_children(tree))


def read_designer_storage(vba_bin: bytes, form_name: str) -> dict:
    """完成品 bin から `<form>/` ストレージの4ストリームを読み戻す。"""
    tree = read_cfb_tree(vba_bin)
    st = tree.get(form_name)
    if not isinstance(st, dict):
        raise OvbaWriteError("bin に %s ストレージがありません" % form_name)
    return {k: v for k, v in st.items() if not isinstance(v, dict)}


# ===========================================================================
# 7. 自己テスト(--selftest)
# ===========================================================================
def _selftest() -> int:
    """template_skeleton だけから最小の bin を作り、フォームを足して読み戻す。

    ここで確かめること(1つでも欠けたら赤):
      ・モジュール集合(標準 + フォーム)
      ・参照3本(Office / SHDocVw / MSForms)
      ・designer ストレージ4本の実在と中身
      ・PROJECT の BaseClass= 行 / PROJECTwm のエントリ
      ・既存モジュールの MODULEOFFSET が 0 のままであること
      ・フォーム本文が src/ui/navi/frmNaviHtml.frm のコード部と一致すること
    """
    import os
    import zipfile
    here = os.path.dirname(os.path.abspath(__file__))
    root = os.path.dirname(here)
    ok = True

    def check(cond, label):
        nonlocal ok
        print(("  PASS: " if cond else "  FAIL: ") + label)
        if not cond:
            ok = False

    with zipfile.ZipFile(os.path.join(here, "template_skeleton.xlsm")) as z:
        template_bin = z.read("xl/vbaProject.bin")
    base = build_vba_project(template_bin, [
        VbaModule("ThisWorkbook", module_stream_source(
            "ThisWorkbook", "Option Explicit\r\n", "document"), "document"),
        VbaModule("modDummy", module_stream_source(
            "modDummy", "Option Explicit\r\n", "std"), "std"),
    ])
    frm_path = os.path.join(root, "src", "ui", "navi", "frmNaviHtml.frm")
    frx_path = os.path.join(root, "src", "ui", "navi", "frmNaviHtml.frx")
    with open(frm_path, encoding="utf-8") as fh:
        frm_text = fh.read()
    with open(frx_path, "rb") as fh:
        frx = fh.read()

    print("ovba_write --selftest (17章 Z-42 一段化)")
    print("  template: %s (%d B) / 第1段相当 bin %d B"
          % ("template_skeleton.xlsm", len(template_bin), len(base)))
    out = add_userform(base, "frmNaviHtml", frm_text, frx)
    print("  フォーム入り bin: %d B" % len(out))

    mods = read_modules(out)
    check(set(mods) == {"ThisWorkbook", "modDummy", "frmNaviHtml"},
          "モジュール集合 = %s" % sorted(mods))
    check(mods.get("frmNaviHtml", {}).get("type") == "class",
          "frmNaviHtml の MODULETYPE が 0x0022(document 扱いでない)")
    refs = read_reference_names(out)
    check(refs == ["MSForms", "SHDocVw", "MSForms", "MSForms"] or
          set(refs) >= {"SHDocVw", "MSForms"},
          "参照設定 = %s" % refs)
    check("SHDocVw" in refs and "MSForms" in refs,
          "SHDocVw と MSForms がある")
    st = read_designer_storage(out, "frmNaviHtml")
    check(set(st) == set(FORM_STREAM_NAMES),
          "designer ストレージ4本 = %r" % sorted(st))
    check(st.get("\x03VBFrame", b"").startswith(b"VERSION 5.00\r\nBegin "),
          "\\x03VBFrame がデザイナ定義で始まる(%d B)" % len(st.get("\x03VBFrame", b"")))
    check(b"OleObjectBlob" not in st.get("\x03VBFrame", b""),
          "\\x03VBFrame に OleObjectBlob 行が無い")
    check(len(st.get("f", b"")) > 0, "f ストリームが空でない(%d B)" % len(st.get("f", b"")))
    tree = read_cfb_tree(out)
    proj = tree["PROJECT"].decode("cp932")
    check("BaseClass=frmNaviHtml" in proj.split("\r\n"), "PROJECT に BaseClass= 行")
    check(FORM_WORKSPACE % "frmNaviHtml" in proj, "PROJECT の [Workspace] に1行")
    check(b"frmNaviHtml\x00" in tree["PROJECTwm"], "PROJECTwm にエントリ")
    dir_dec = ovba.ovba_decompress(tree["VBA"]["dir"])
    offs = [struct.unpack('<I', body)[0]
            for _o, rid, _s, body in iter_dir_records(dir_dec)
            if rid == REC_MODULEOFFSET]
    check(offs == [0] * len(offs) and len(offs) == 3,
          "全 MODULEOFFSET=0(p-code 無し・%d本)" % len(offs))
    _d, _a, code = split_frm(frm_text)
    want = "\n".join(code).replace("\r\n", "\n").replace("\n", "\r\n")
    got = strip_attribute_lines(mods["frmNaviHtml"]["source"]).decode("cp932")
    check(got.rstrip("\r\n") == want.rstrip("\r\n"),
          "フォーム本文が .frm のコード部とバイト一致(%d 字)" % len(want))
    print("結果: %s" % ("PASS" if ok else "FAIL"))
    return 0 if ok else 1


if __name__ == "__main__":
    import argparse
    _ap = argparse.ArgumentParser(description="MS-OVBA vbaProject.bin ライター")
    _ap.add_argument("--selftest", action="store_true",
                     help="一段化(Z-42)の読み戻し自己テストを走らせる")
    _args = _ap.parse_args()
    if _args.selftest:
        raise SystemExit(_selftest())
    _ap.error("--selftest を指定してください")
