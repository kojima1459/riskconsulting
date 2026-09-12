#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ship_check.py - 出荷前検問(17章 T-46)のうち機械実行できる①②③を走らせる

================================================================================
17章 T-46 は毎リリース必須の5項目からなる。本スクリプトはそのうち Linux 側で
機械実行できる①②③を実装する(④⑤は Windows 実機の wintest が要るため、
本スクリプトは「未実施」であることを明示して申告するだけで、合格扱いにしない)。

  ① 外部由来テキストの直書き検査(16章 NFR-S7)
     対象 src/**/*.bas から modUtilText(防御関数の定義本体)を除外し、
     `\\.Value2?\\s*=` / `\\.Formula\\w*\\s*=` / `(Cells|Range|Offset)\\([^)]*\\)\\s*=`
     でヒットを取る。行末に `' SAFE:const` の注記が無いヒットのみFAIL。
     HTML経路(差し込む値が HtmlSafe / JsStringSafe を経由しない埋め込み)も同検問。
     **実装は tools/vba_lint.py の同名ルールを import して共有する**
     (二重実装すると片方だけ緩んでも誰も気付けないため)。
  ② キー走査(16章 NFR-S2 の走査仕様4項目)
     対象 = 全trackedファイル ＋ 成果物 *.xlsm / *.xlsx を全パート展開したテキスト。
     検出パターン = 語頭の sk-[A-Za-z0-9_-]{20,} ／ 連続32文字以上の [0-9a-f] ／
     連続40文字以上のBase64様文字列 ／ 文字列リテラルに OBF・難読・_KEY を含む
     定数定義。build/ 配下がキーを読む行が無いことも grep で確認する。
     **検出時はファイル名と行番号のみを出力し、キー値そのものは出力しない**
     (検問ログ自体が漏洩経路にならないようにする)。
  ③ `git check-ignore` で dist/ と成果物(*.xlsm / *.xlsx)が実際に除外されている
     ことの確認。あわせて tracked の .xlsm がビルド入力の許可枠
     (TRACKED_XLSM_ALLOWED = build/template_skeleton.xlsm。成果物ではない)
     だけであること、tracked の .xlsx が仕様入力(docs/ 配下)だけであることを確認する。

  ④ wintest(実Excel・COM経由)全PASS  ... Windows実機。本スクリプトでは未実施
  ⑤ modTestsPure が FAIL 0・SKIP 0・実行本数=tests_expected ... 同上
     (Linux側の同等確認は tools/run_lo_tests.py が行うが、17章§1のとおり
      **(c)緑は出荷条件ではない**。④⑤は実機で確認すること)

使い方:
    python3 tools/ship_check.py
    exit code: 0 = ①②③すべてPASS / 1 = いずれか失格
================================================================================
"""

from __future__ import annotations

import argparse
import math
import re
import subprocess
import sys
import zipfile
from collections import Counter
from pathlib import Path

TOOLS_DIR = Path(__file__).resolve().parent
REPO_ROOT = TOOLS_DIR.parent
sys.path.insert(0, str(TOOLS_DIR))

import vba_lint   # noqa: E402  (①の実装を共有するため)

ARTIFACT_SUFFIXES = (".xlsm", ".xlsx")
# 仕様入力として tracked を許す .xlsx の置き場(成果物ではない)。
TRACKED_XLSX_ALLOWED_PREFIXES = ("docs/",)
# tracked を許す .xlsm(成果物ではなくビルド入力)。build/template_skeleton.xlsm は
# 「本物の vbaProject.bin の供給元」で、自己インストール機構(12章§2・裁定書4 項目12)
# の土台。これだけは意図的に tracked にする(.gitignore の !build/template_skeleton.xlsm)。
# 中身は下の ② が全パート展開してキー走査する(除外ではなく検査対象に残す)。
TRACKED_XLSM_ALLOWED = ("build/template_skeleton.xlsm",)

# ------------------------------------------------------------------------------
# ② 検出パターン(16章 NFR-S2 の逐語)
# ------------------------------------------------------------------------------
# 直前が英数字なら語の途中("Risk-consulting-Navi_HTML_v0.3" の sk-consulting…)なので除外(W12-A)
PAT_SK = re.compile(r"(?<![A-Za-z0-9])sk-[A-Za-z0-9_-]{20,}")
PAT_HEX32 = re.compile(r"(?<![0-9A-Za-z])[0-9a-f]{32,}(?![0-9A-Za-z])")
PAT_B64 = re.compile(r"[A-Za-z0-9+/]{40,}={0,2}")
# 定数定義の中の OBF / 難読 / _KEY(VBA の Const と Python 等の大文字定数の両形)。
PAT_CONST_OBF = re.compile(
    r"(?:\bConst\s+\w+(?:\s+As\s+\w+)?\s*=|^\s*_?[A-Z][A-Za-z0-9_]*\s*=)\s*"
    r"[\"'][^\"']*(?:[Oo][Bb][Ff]|難読|_KEY)[^\"']*[\"']")
# ③のbuildスクリプト検査(「ビルド処理はキーに触れない」の機械的担保)。
PAT_BUILD_KEY_READ = re.compile(
    r"api_key\.txt"
    r"|(?:os\.environ|getenv|environ\.get)\s*[\(\[]\s*[\"'][^\"']*(?:KEY|SECRET|TOKEN)")


WORDLIKE_SEGMENT_RE = re.compile(r"^(?:[A-Za-z]+|[0-9]+)$")
# 識別子の形(先頭が英字・以降は英数字と _ ・24文字以内)。VBAの関数名や
# モジュール名(AsmS1User / modTestsPure4 / PptMaxSlidesT2 …)がこれに当たる。
IDENTIFIER_SEGMENT_RE = re.compile(r"^[A-Za-z][A-Za-z0-9_]{0,23}$")
# 識別子列とみなしてよいシャノンエントロピーの上限(bit/文字)。下の実測を参照。
IDENTIFIER_LIST_MAX_ENTROPY = 4.4


def _shannon_entropy(s: str) -> float:
    """文字出現頻度のシャノンエントロピー(bit/文字)。"""
    if not s:
        return 0.0
    n = len(s)
    return -sum((c / n) * math.log2(c / n) for c in Counter(s).values())


def _b64_looks_like_secret(s: str) -> bool:
    """Base64様の40文字以上のうち、鍵素材らしいものだけを拾う精度ガード。

    素の「40文字以上の [A-Za-z0-9+/] の連なり」は、スラッシュ区切りの語の列にも
    当たる。実測で本リポジトリでは9件、成果物のOOXMLでは名前空間URLの断片
    (`org/officeDocument/2006/relationships/worksheet` 等)が当たった。
    そこで3つの条件を課す。どれもランダムな鍵素材はまず満たさない。
      (1) 数字・英大文字・英小文字の3種すべてを含むこと
          (40文字でいずれか1種も含まない確率は 1e-9 未満。全小文字16進の鍵は
           PAT_HEX32 が別途拾うので検出力は落ちない)
      (2) "/" で切った断片が「全部が純アルファベットか純数字」ではないこと
          (=語の列。名前空間URLの断片はここで落ちる)
      (3) 「断片が全部"識別子の形"」かつ「エントロピーが低い」のではないこと
          (=関数名をスラッシュで並べた列。本リポジトリは14章§6の関数名一覧を
           `Fill/AsmS1User/AsmS2User/AsmS4System/AsmS4User` の書き癖で書くため、
           (2)だけでは AsmS1User のような英数混在の識別子を素通しできず、
           コメント1行が毎回赤になっていた)

    (3)の閾値の根拠(実測。乱数鍵 3,000,000 本 = os.urandom を base64 化した
    長さ40～92文字の列で計測):
      ・(3)の「識別子の形の列」に偶然当たった鍵は 38,015 本(1.27%)。
        その最小エントロピーは 4.234 bit/文字。
      ・閾値 4.4 で取りこぼす鍵は 23 本 = 1/130,000(0.00077%)。
        既に運用している(1)単体の取りこぼし率 0.019%(=1/5,300)より 25倍小さい
        ので、ガードを足したことで検出力の桁は下がっていない。
      ・一方、本リポジトリに実在する「識別子の形の列」23件のエントロピーは
        3.68～4.51 で、(1)を通過して(3)まで届く唯一の実在ヒット
        (modTestsPure4 の関数名一覧)は 3.68。閾値まで 0.72 の余裕がある。
      ・エントロピーが 4.4 以上の「識別子の形の列」は今まで通り赤のままにする
        (=見逃しではなく人が1回見る)。
    """
    if not (any(c.isdigit() for c in s)
            and any(c.isupper() for c in s)
            and any(c.islower() for c in s)):
        return False
    segs = [x for x in s.split("/") if x]
    if len(segs) >= 2 and all(WORDLIKE_SEGMENT_RE.match(x) for x in segs):
        return False
    if (len(segs) >= 2
            and all(IDENTIFIER_SEGMENT_RE.match(x) for x in segs)
            and _shannon_entropy(s) < IDENTIFIER_LIST_MAX_ENTROPY):
        return False
    return True


def _executable_lines(path: Path) -> list[tuple[int, str]]:
    """行コメントと(.py の)三重引用符ブロックを除いた「実行される行」を返す。

    build/ のキー読取検査で、docstring に書いた「キーを読まない」という説明文を
    違反として拾わないため。三重引用符の中は実行されないので、そこに書かれた
    api_key.txt は「読んでいる」ことにならない。
    """
    out: list[tuple[int, str]] = []
    in_doc = ""
    for i, line in enumerate(path.read_text(encoding="utf-8", errors="replace").splitlines(), 1):
        rest = line
        if path.suffix == ".py":
            while rest:
                if in_doc:
                    idx = rest.find(in_doc)
                    if idx < 0:
                        rest = ""
                        break
                    rest = rest[idx + 3:]
                    in_doc = ""
                    continue
                m = re.search(r'"""|\'\'\'', rest)
                if not m:
                    break
                head, rest = rest[:m.start()], rest[m.end():]
                out.append((i, head))
                in_doc = m.group(0)
            if in_doc:
                continue
        stripped = rest.lstrip()
        if stripped.startswith("#") or stripped.startswith("'"):
            continue
        out.append((i, rest))
    return out


def scan_text(name: str, text: str) -> list[str]:
    """1つのテキストを走査し「ファイル名:行番号 種別」だけの行を返す(値は出さない)。"""
    hits: list[str] = []
    for i, line in enumerate(text.splitlines(), 1):
        if PAT_SK.search(line):
            hits.append(f"{name}:{i} APIキー形状(sk-...)")
        if PAT_HEX32.search(line):
            hits.append(f"{name}:{i} 連続32文字以上の16進")
        m = PAT_B64.search(line)
        if m and _b64_looks_like_secret(m.group(0)):
            hits.append(f"{name}:{i} 連続40文字以上のBase64様文字列")
        if PAT_CONST_OBF.search(line):
            hits.append(f"{name}:{i} 難読化/キーらしき定数定義(OBF・難読・_KEY)")
    return hits


def tracked_files() -> list[str]:
    out = subprocess.run(["git", "-C", str(REPO_ROOT), "ls-files", "-z"],
                         capture_output=True)
    return [f.decode("utf-8") for f in out.stdout.split(b"\0") if f]


# 未trackedでも走査する作業ツリーのディレクトリ(裁定書5 項目9)。
# 理由: ②のキー走査が tracked だけを見ていると、実装中の新規ファイル(コミット
# 前)がまるごと素通りする。W1では新設11本の .bas が untracked のままで、
# 成果物 dist/ の vba_src に載っていたおかげで偶然拾えていただけだった。
# ビルド前に ship_check を回す運用や、vba_src に載らないファイル(tools/ の
# スクリプト等)では救済経路が無い。source 側の作業ディレクトリを直接足す。
UNTRACKED_SCAN_DIRS = ("src", "build", "tools", "wintest")


def untracked_worktree_files() -> list[str]:
    """UNTRACKED_SCAN_DIRS 配下の未trackedファイル(.gitignore 対象は除く)。"""
    out = subprocess.run(
        ["git", "-C", str(REPO_ROOT), "ls-files", "-z", "--others",
         "--exclude-standard", "--"] + list(UNTRACKED_SCAN_DIRS),
        capture_output=True,
    )
    return [f.decode("utf-8") for f in out.stdout.split(b"\0") if f]


def is_ignored(rel: str) -> bool:
    r = subprocess.run(["git", "-C", str(REPO_ROOT), "check-ignore", "-q", rel],
                       capture_output=True)
    return r.returncode == 0


# ==============================================================================
# ① 外部由来テキストの直書き検査
# ==============================================================================
def check_item1(src_root: Path) -> tuple[bool, list[str]]:
    print("=" * 78)
    print("T-46 ① 外部由来テキストの直書き検査(16章 NFR-S7)")
    print("=" * 78)
    if not src_root.exists():
        print("  src/ がありません。検査対象0件。")
        return True, []

    findings: list[str] = []
    files = vba_lint.discover_module_files(src_root)
    for path in files:
        info = vba_lint.load_module(path, src_root)
        vba_lint.check_cell_write_guard(info)
        vba_lint.check_html_embed_guard(info)
        for f in info.findings:
            if f.level == "ERROR":
                findings.append(f"{info.relpath.as_posix()}:{f.line} {f.message}")

    print(f"  検査対象 .bas/.cls: {len(files)}本"
          f"(除外: {sorted(vba_lint.CELL_WRITE_EXEMPT_MODULES)})")
    if findings:
        print(f"  FAIL: 注記の無いヒットが {len(findings)} 件")
        for f in findings:
            print(f"    - {f}")
        return False, findings
    print("  PASS: SetCellSafe / HtmlSafe / JsStringSafe を迂回する書込はありません")
    return True, []


# ==============================================================================
# ② キー走査
# ==============================================================================
def check_item2(dist_dir: Path) -> tuple[bool, list[str]]:
    print("\n" + "=" * 78)
    print("T-46 ② キー走査(16章 NFR-S2 の走査仕様4項目)")
    print("=" * 78)

    hits: list[str] = []
    files = tracked_files()
    tracked_count = len(files)
    # 作業ツリーの未trackedファイル(src/build/tools/wintest 配下)も対象に加える。
    extra = [f for f in untracked_worktree_files() if f not in set(files)]
    files = files + extra
    scanned = 0
    for rel in files:
        p = REPO_ROOT / rel
        if not p.exists():
            continue
        if p.suffix.lower() in ARTIFACT_SUFFIXES:
            hits.extend(scan_archive(p, rel))
            scanned += 1
            continue
        try:
            text = p.read_text(encoding="utf-8", errors="replace")
        except OSError:
            continue
        hits.extend(scan_text(rel, text))
        scanned += 1
    print(f"  走査(tracked {tracked_count}件 + 未tracked {len(extra)}件"
          f"[{'/'.join(UNTRACKED_SCAN_DIRS)}]): {scanned}ファイル")

    # 成果物(未trackedでも必ず走査する。配布物こそが本命)。
    art = 0
    if dist_dir.exists():
        for p in sorted(dist_dir.rglob("*")):
            if p.suffix.lower() in ARTIFACT_SUFFIXES:
                hits.extend(scan_archive(p, str(p.relative_to(REPO_ROOT))))
                art += 1
    print(f"  走査(成果物 {dist_dir.name}/ の *.xlsm *.xlsx): {art}ファイル(全パート展開)")

    # buildスクリプトがキーを読んでいないことの確認。
    build_hits: list[str] = []
    for p in sorted((REPO_ROOT / "build").glob("**/*")):
        if p.suffix not in (".py", ".ps1", ".sh", ".bat"):
            continue
        rel = str(p.relative_to(REPO_ROOT))
        # 注釈・docstring に「キーを読まない」と書くのは違反ではないので、
        # 実行される行だけを見る。
        for i, line in _executable_lines(p):
            if PAT_BUILD_KEY_READ.search(line):
                build_hits.append(f"{rel}:{i} build配下がキーを読んでいます")
    print(f"  build/ のキー読取検査: {'NG' if build_hits else 'OK'}")

    hits.extend(build_hits)
    if hits:
        print(f"  FAIL: 検出 {len(hits)} 件(値は出力しません)")
        for h in hits:
            print(f"    - {h}")
        return False, hits
    print("  PASS: キー形状・難読化定数の検出はゼロ、build/ はキーを読みません")
    return True, []


def scan_archive(path: Path, rel: str) -> list[str]:
    """xlsx/xlsm(ZIP)を全パート展開して走査する。素のgrepではXMLパート内が
    見えないため展開が必須(姉妹PJの実漏洩は sheet23.xml に埋まっていた)。"""
    hits: list[str] = []
    try:
        with zipfile.ZipFile(path) as z:
            for name in z.namelist():
                try:
                    data = z.read(name)
                except Exception:
                    continue
                hits.extend(scan_text(f"{rel}!{name}",
                                      data.decode("utf-8", errors="replace")))
    except zipfile.BadZipFile:
        hits.append(f"{rel}:0 ZIPとして開けません(成果物の検査が実施できていません)")
    return hits


# ==============================================================================
# ③ 除外の実効確認
# ==============================================================================
def check_item3() -> tuple[bool, list[str]]:
    print("\n" + "=" * 78)
    print("T-46 ③ dist/ と成果物(*.xlsm / *.xlsx)の除外の実効確認")
    print("=" * 78)

    problems: list[str] = []
    probes = ["dist/", "dist/リスク提案ナビ.xlsm", "dist/リスク提案ナビ.xlsx",
              "リスク提案ナビ.xlsm"]
    for probe in probes:
        ok = is_ignored(probe)
        print(f"  git check-ignore {probe:<32} -> {'除外OK' if ok else '除外されていない'}")
        if not ok:
            problems.append(f"{probe} が .gitignore で除外されていません")

    allowed_xlsm = 0
    for rel in tracked_files():
        low = rel.lower()
        if low.endswith(".xlsm"):
            if rel in TRACKED_XLSM_ALLOWED:
                allowed_xlsm += 1
                continue
            problems.append(f"tracked に .xlsm があります: {rel}")
        elif low.endswith(".xlsx"):
            if not any(rel.startswith(p) for p in TRACKED_XLSX_ALLOWED_PREFIXES):
                problems.append(
                    f"tracked の .xlsx が仕様入力の置き場"
                    f"{TRACKED_XLSX_ALLOWED_PREFIXES} の外にあります: {rel}")
    if allowed_xlsm:
        print(f"  tracked のビルド入力 .xlsm(成果物ではない・②が全パート走査): "
              f"{list(TRACKED_XLSM_ALLOWED)}")

    if problems:
        print(f"  FAIL: {len(problems)} 件")
        for p in problems:
            print(f"    - {p}")
        return False, problems
    print("  PASS: dist/ と成果物は除外され、tracked に成果物はありません")
    return True, []


# ==============================================================================
# ⑥ 配布物のAV表面積と経路の固定(裁定書27 W9-B 6)
# ------------------------------------------------------------------------------
# 2026-09-02 に社内AVのAMSIが自己インストーラ(VBProject/AddFromString)を検知
# して VDI が強制停止した。配布方式Bへ切り替えた以上、「配布物にその形が本当に
# 残っていないか」を毎リリース機械で確かめないと、いつか戻る。
# あわせて「配布物では direct 経路も mock も使えない」ことを config の実体で
# 固定する(コード側の分岐だけでは、configを1行足せば経路が生き返るため)。
# 検査は **prod の配布ブック**(dist/リスク提案ナビ.xlsm)に対して行う。
# ブックが無い/読めない/binが無いは**すべて失格**(fail-closed)。
# ==============================================================================
PROD_BOOK_NAME = "リスク提案ナビ.xlsm"


def _read_config_pairs(book: Path) -> dict:
    import openpyxl
    wb = openpyxl.load_workbook(book, read_only=True, keep_links=False)
    try:
        if "config" not in wb.sheetnames:
            return {}
        ws = wb["config"]
        out = {}
        for row in ws.iter_rows(min_row=2, max_col=2, values_only=True):
            if not row or row[0] in (None, ""):
                continue
            out[str(row[0]).strip()] = row[1]
        return out
    finally:
        wb.close()


def check_item6(dist_dir: Path) -> tuple[bool, list[str]]:
    print("\n" + "=" * 78)
    print("裁定書27 W9-B 6 ⑥ 配布物のAV表面積と経路の固定(prod ブック)")
    print("=" * 78)

    problems: list[str] = []
    book = dist_dir / PROD_BOOK_NAME
    if not book.exists():
        print(f"  FAIL: 配布ブックがありません: {book}")
        return False, [f"{book} がありません(python3 build/build_rpn.py --prod)"]

    # --- config の実体 --------------------------------------------------------
    cfg = _read_config_pairs(book)
    if not cfg:
        problems.append("config シートを読めません(または空です)")
    transport = str(cfg.get("llm_transport", "")).strip().lower()
    print(f"  llm_transport      : {transport!r}")
    if transport != "ribbon":
        problems.append(f"config!llm_transport が 'ribbon' ではありません: {transport!r}")
    print(f"  direct_api_base    : "
          f"{'あり(失格)' if 'direct_api_base' in cfg else 'なし'}")
    if "direct_api_base" in cfg:
        problems.append("config に direct_api_base が載っています"
                        "(prod では direct 経路の入口を置かない)")
    mock = cfg.get("mock_llm")
    print(f"  mock_llm           : {mock!r}")
    if mock is None:
        problems.append("config に mock_llm がありません(mockが無効である証跡が無い)")
    elif bool(mock):
        problems.append("config!mock_llm が TRUE です(配布物でmockを有効にしない)")
    if transport == "mock":
        problems.append("config!llm_transport が 'mock' です")

    # --- vba_src シートの不在 -------------------------------------------------
    import openpyxl
    wb = openpyxl.load_workbook(book, read_only=True, keep_links=False)
    sheetnames = list(wb.sheetnames)
    wb.close()
    print(f"  vba_src シート     : {'あり(失格)' if 'vba_src' in sheetnames else 'なし'}")
    if "vba_src" in sheetnames:
        problems.append("配布物に隠しシート vba_src が残っています"
                        "(自己インストール機構は撤去済みのはず)")

    # --- vbaProject.bin の禁止文字列 ------------------------------------------
    sys.path.insert(0, str(REPO_ROOT / "build"))
    import build_rpn      # 禁止文字列の表と判定は build 側の1実装を共有する
    vba_bin = b""
    with zipfile.ZipFile(book) as z:
        if "xl/vbaProject.bin" not in z.namelist():
            problems.append("配布物に xl/vbaProject.bin がありません")
            hits = []
        else:
            vba_bin = z.read("xl/vbaProject.bin")
            # prod=True で direct経路の痕跡(ServerXMLHTTP / MSXML2 / XMLHTTP)も
            # 見る(裁定書30 裁定1(f))。dev ビルドには載ってよいので、この3語は
            # prod ブックを見るこの⑥だけの禁止語である。
            hits = build_rpn.forbidden_strings_in_bin(vba_bin, prod=True)
    print(f"  bin の禁止文字列   : {hits if hits else 'なし'}"
          f"  (表: {list(build_rpn.FORBIDDEN_BIN_STRINGS)}"
          f" + prod限定 {list(build_rpn.FORBIDDEN_BIN_STRINGS_PROD)})")
    if hits:
        problems.append("vbaProject.bin に配布禁止の文字列があります: "
                        + ", ".join(hits))

    # --- dir ストリームの中身(W9.2) -------------------------------------------
    # (1) 開発者の絶対パス(/Users/...)が焼き込まれていないこと
    # (2) 使っていない MSForms(fm20.tlb)への参照が残っていないこと
    #     どちらも template の PROJECTREFERENCES を丸写ししていた副作用であり、
    #     build/ovba_write.strip_msforms_reference が落とす。落とし忘れを毎回ここで
    #     止める(bin は圧縮されているので、必ず解凍してから見る)。
    dir_hits: list[str] = []
    if vba_bin:
        import ovba       # noqa: E402  (build/ を sys.path へ入れた後に読む)
        try:
            dir_dec = ovba.ovba_decompress(ovba.CFBReader(vba_bin).read("dir"))
        except Exception as e:                                   # pragma: no cover
            problems.append(f"vbaProject.bin の dir を読めません: {e}")
            dir_dec = b""
        if b"/Users/" in dir_dec:
            dir_hits.append("/Users/(個人の絶対パス)")
        if b"MSForms" in dir_dec or b"{0D452EE1-E08F-101A-852E-02608C4D0BB4}" in dir_dec:
            dir_hits.append("MSForms 参照")
    print(f"  dir の残留物       : {dir_hits if dir_hits else 'なし'}")
    if dir_hits:
        problems.append("vbaProject.bin の dir に残ってはいけないものがあります: "
                        + ", ".join(dir_hits))

    if problems:
        print(f"  FAIL: {len(problems)} 件")
        for p in problems:
            print(f"    - {p}")
        return False, problems
    print("  PASS: llm_transport=ribbon / direct_api_base 不在 / mock 不在 / "
          "vba_src 不在 / 禁止文字列 不在 / dir に /Users/ とMSForms参照 不在")
    return True, []


# ------------------------------------------------------------------------------
# ⑦ ランチャー .bat の形(裁定書28「裁定の確定」)
# ------------------------------------------------------------------------------
# なぜ形を検問するのか: 利用者の手順は「OneDrive の『リスク提案ナビを起動』を
#   ダブルクリック」の1つだけである(docs/24 §1)。この .bat が cmd.exe に
#   読めない形(UTF-8・LF)で出ると、利用者の側では「何も起きない」だけになり、
#   起動できない理由が誰にも見えない。MyBookshelf の実績(裁定書28 追補)で
#   確かめられている形を、毎ビルド機械で固定する。
LAUNCHER_NAME = "リスク提案ナビを起動.bat"


def check_item7(dist_dir: Path) -> tuple[bool, list[str]]:
    print("\n" + "=" * 78)
    print("裁定書28 ⑦ 起動ランチャー(リスク提案ナビを起動.bat)の形")
    print("=" * 78)

    problems: list[str] = []
    bat = dist_dir / LAUNCHER_NAME
    if not bat.exists():
        print(f"  FAIL: ランチャーがありません: {bat}")
        return False, [f"{bat} がありません(python3 build/build_rpn.py --prod)"]

    raw = bat.read_bytes()
    print(f"  ファイル           : {bat} ({len(raw):,} bytes)")

    # (1) 改行は CRLF だけ(裸のLFが1つでもあると cmd.exe が行を読み違える)
    crlf = raw.count(b"\r\n")
    lf = raw.count(b"\n")
    print(f"  改行               : CRLF {crlf} / LF 合計 {lf}")
    if lf == 0 or crlf != lf:
        problems.append(f"改行が CRLF だけではありません(CRLF={crlf} / LF={lf})")

    # (2) CP932 で復号できること(UTF-8 で書くと日本語のパスが化ける)
    text = ""
    try:
        text = raw.decode("cp932")
    except UnicodeDecodeError as e:
        problems.append(f"CP932 で復号できません: {e}")
    if raw.startswith(b"\xef\xbb\xbf"):
        problems.append("先頭に UTF-8 BOM があります(cmd.exe が1行目を読み違えます)")

    # (3) `start "" excel.exe` の形(第1引数の "" が無いとパスが窓題名になる)
    has_start = 'start "" excel.exe' in text
    print(f"  start \"\" 形式      : {'あり' if has_start else 'なし(失格)'}")
    if not has_start:
        problems.append('start "" excel.exe の形がありません(裁定書28 追補)')

    # (4) excel.exe をフルパスで書いていないこと(版差・言語差で壊れる)
    full = re.findall(r"[A-Za-z]:\\[^\r\n\"]*excel\.exe", text, re.IGNORECASE)
    print(f"  excel のフルパス   : {full if full else 'なし'}")
    if full:
        problems.append("excel.exe をフルパスで起動しています: " + ", ".join(full))

    # (5) 裁定の4手順が揃っていること(SRC/DST・mkdir・xcopy /D /Y・data_dir.txt)
    # rem 行は除いて照合する(コメントに語が残っているだけでは合格させない。
    # 司令塔の変異注入で data_dir.txt の書き出し行を消しても rem 行の語で
    # 通ってしまったため、コマンド行の実形で照合する)。
    cmd_lines = [ln for ln in text.splitlines()
                 if ln.strip() and not ln.strip().lower().startswith("rem")]
    cmd_text = "\n".join(cmd_lines)
    for needed, why in (
        ("set \"SRC=%~dp0\"", "SRC=配布フォルダ"),
        ("D:\\リスク提案ナビ", "DST=D:"),
        ("%TEMP%\\リスク提案ナビ", "D:が無いときの退避先"),
        ("mkdir \"%DST%\"", "DSTの作成"),
        ("xcopy /D /Y \"%SRC%リスク提案ナビ.xlsm\"", "本体を新しければコピー"),
        ("xcopy /D /Y \"%SRC%ナレッジブック.xlsx\"", "ナレッジブックを新しければコピー"),
        ("> \"%DST%\\data_dir.txt\" echo %SRC%データ", "data_dir ポインタの書き出し(コマンド行)"),
        ("start \"\" excel.exe /x \"%DST%\\リスク提案ナビ.xlsm\"", "D: の本体を Excel で開く"),
    ):
        if needed not in cmd_text:
            problems.append(f"ランチャーに {why} のコマンド行がありません: {needed!r}")

    if problems:
        print(f"  FAIL: {len(problems)} 件")
        for p in problems:
            print(f"    - {p}")
        return False, problems
    print("  PASS: 存在 / CRLF / CP932 / start \"\" 形式 / excelフルパス無し / 4手順")
    return True, []


# ------------------------------------------------------------------------------
# ⑧ 手順書のコードブロック == dist の .bat(裁定書31 裁定2)
# ------------------------------------------------------------------------------
# なぜ照合するのか: 会社のメールは `.bat` も `.bat.txt` も受信時に削除し、zip は
#   Gmail 側が送信を拒否する(2026-09-05 実測)。したがって発行者が会社PCへ最初の
#   1本を持ち込む唯一の手段は「**手順書のコードブロックをメモ帳へ貼って保存する**」
#   になった(docs/24 §8.1c-1)。手順書の枠が配布物の .bat と1行でもずれると、
#   利用者の側では「押しても何も起きない」だけになり、原因が誰にも見えない。
#   改行(CRLF/LF)だけは書き方の違いなので LF へ正規化してから**バイト比較**する。
LAUNCHER_DOC = REPO_ROOT / "docs" / "24_実機テスト手順書_Windows.md"
# 手順書側のコードブロック(```bat ... ```)。1つだけ在ることも条件にする
# (増えるとどちらが正か分からなくなる)。
DOC_BAT_BLOCK_RE = re.compile(r"^```bat[ \t]*\r?\n(.*?)^```", re.M | re.S)


def _normalize_launcher(text: str) -> str:
    """CRLF・CR を LF へ正規化する(改行の書き方だけを無視する)。"""
    return text.replace("\r\n", "\n").replace("\r", "\n")


def compare_launcher(doc_text: str, bat_text: str) -> list[str]:
    """手順書のコードブロックと .bat の中身を突き合わせる(⑧の判定本体)。
    自己テストからも呼ぶので、ファイルI/Oを持たない純関数にしてある。"""
    blocks = DOC_BAT_BLOCK_RE.findall(doc_text)
    if len(blocks) != 1:
        return [f"docs/24 の ```bat コードブロックが {len(blocks)} 個です(1個であること)"]

    doc = _normalize_launcher(blocks[0])
    bat = _normalize_launcher(bat_text)
    if doc == bat:
        return []

    doc_lines = doc.split("\n")
    bat_lines = bat.split("\n")
    problems = []
    if len(doc_lines) != len(bat_lines):
        problems.append(f"行数が違います(手順書={len(doc_lines)} / bat={len(bat_lines)})")
    for i in range(min(len(doc_lines), len(bat_lines))):
        if doc_lines[i] != bat_lines[i]:
            problems.append(f"{i + 1}行目が違います: 手順書={doc_lines[i]!r} / bat={bat_lines[i]!r}")
    if not problems:
        problems.append("末尾が違います(手順書と bat の中身が一致しません)")
    return problems


def selftest_item8() -> list[str]:
    """⑧が空振りしていないことの自己テスト(骨抜き防止)。
    正例=同じ本文なら問題なし / 負例=1行変える・1行消すと必ず問題が出る。"""
    body = '@echo off\nset "DST=D:\\x"\nstart "" excel.exe /x "%DST%\\a.xlsm"\n'
    doc_ok = "前書き\n\n```bat\n" + body + "```\n後書き\n"
    out = []
    if compare_launcher(doc_ok, body.replace("\n", "\r\n")):
        out.append("自己テスト(正例): 同じ本文(CRLF違いだけ)を⑧が不一致と判定しました")
    changed = body.replace('set "DST=D:\\x"', 'set "DST=D:\\y"')
    if not compare_launcher("```bat\n" + changed + "```\n", body):
        out.append("自己テスト(負例1): 手順書側の1行を変えても⑧が通りました")
    dropped = body.replace('set "DST=D:\\x"\n', "")
    if not compare_launcher("```bat\n" + dropped + "```\n", body):
        out.append("自己テスト(負例2): 手順書側の1行を消しても⑧が通りました")
    if not compare_launcher("本文にコードブロックが無い\n", body):
        out.append("自己テスト(負例3): コードブロックが無くても⑧が通りました")
    return out


def check_item8(dist_dir: Path) -> tuple[bool, list[str]]:
    print("\n" + "=" * 78)
    print("裁定書31 ⑧ 手順書(docs/24 §8.1c-1)のコードブロック == dist の .bat")
    print("=" * 78)

    problems = selftest_item8()
    print(f"  自己テスト(正例1/負例3): {'OK' if not problems else 'NG'}")

    bat = dist_dir / LAUNCHER_NAME
    if not bat.exists():
        print(f"  FAIL: ランチャーがありません: {bat}")
        return False, problems + [f"{bat} がありません(python3 build/build_rpn.py --prod)"]
    if not LAUNCHER_DOC.exists():
        print(f"  FAIL: 手順書がありません: {LAUNCHER_DOC}")
        return False, problems + [f"{LAUNCHER_DOC} がありません"]

    try:
        bat_text = bat.read_bytes().decode("cp932")
    except UnicodeDecodeError as e:
        return False, problems + [f".bat を CP932 で復号できません: {e}"]
    doc_text = LAUNCHER_DOC.read_text(encoding="utf-8")

    problems += compare_launcher(doc_text, bat_text)
    print(f"  手順書             : {LAUNCHER_DOC}")
    print(f"  照合対象           : {bat}")

    if problems:
        print(f"  FAIL: {len(problems)} 件")
        for p in problems:
            print(f"    - {p}")
        return False, problems
    print("  PASS: 手順書のコードブロックと dist の .bat が一致(CRLF→LF 正規化のうえ)")
    return True, []



# ==============================================================================
# 第2段(開発PC)の産物の検査: --final(裁定書34 §0.3・§1.4・W12-A)
# ------------------------------------------------------------------------------
# 2段ビルドのおさらい:
#   第1段 = 当方CI(build/build_rpn.py)。標準モジュールだけの
#           dist/リスク提案ナビ.xlsm。ui_mode=sheet で単体でも動く。
#   第2段 = Windows 実Excel を持つ開発PC(build/win/import_navi_modules.ps1)。
#           UserForm(frmNaviHtml)と参照設定を第1段の産物へ組み込み、
#           dist/final/ へ書き出す。**当方CIでは実行できない**(実Excelが要る)。
# ここは「第2段の産物が返ってきたときに、当方の目で見る」ための検査である。
# CI では dist/final/ が無いので SKIP(赤にしない)。
#
# 見るもの(4条件):
#   (1) モジュール集合 = 台帳の ship 集合(標準モジュール) + navi 8本 + フォーム1
#       -> 第2段が「入れ忘れた」「余計なものを入れた」を止める。
#   (2) 参照設定が VBA / Excel / stdole / Office / SHDocVw / MSForms の6つだけ
#       -> 参照が増えると配布先の端末で「参照不可」になって全部動かなくなる。
#   (3) prod の禁止文字列が bin に無い(⑥と同じ表を共有する)
#   (4) ui/ の5本が本体と同じフォルダに同梱されている
# ==============================================================================
FINAL_DIR_NAME = "final"
# UserForm の器と、その中で使う HTML 画面の資産。
FINAL_FORM_NAME = "frmNaviHtml"
FINAL_UI_FILES = ("index.html", "style.css", "markdown.js", "views.js", "app.js")
# 第2段が足してよい参照設定(これ以外が1つでもあれば失格)。
#   VBA/Excel/stdole/Office = どの xlsm にも既定で入っている4つ
#   SHDocVw  = Microsoft Internet Controls(WebBrowser)
#   MSForms  = Microsoft Forms 2.0(UserForm)
FINAL_ALLOWED_REFS = ("VBA", "Excel", "stdole", "Office", "SHDocVw", "MSForms")


def _final_expected_modules() -> tuple[set, set]:
    """(標準モジュール名の集合, フォーム名の集合) を build/modules.json から。"""
    import json
    data = json.loads((REPO_ROOT / "build" / "modules.json").read_text(encoding="utf-8"))
    std, forms = set(), set()
    for m in data["modules"]:
        if m.get("ship") is False:
            continue                      # prod の配布物から外す印(裁定書27)
        if m.get("type") == "form":
            forms.add(m["name"])
        else:
            std.add(m["name"])
    return std, forms


# 第2段のスクリプト(開発PCでしか動かないので、当方CIは**在ることと形**だけ見る)。
STAGE2_SCRIPT = REPO_ROOT / "build" / "win" / "import_navi_modules.ps1"
# 落としてはいけない要素。ここが消えると第2段の産物が静かに欠ける。
STAGE2_REQUIRED = (
    ("開発PC専用", "開発PC専用である旨の注記"),
    ("frmNaviHtml.frm", "UserForm の取り込み"),
    ("frmNaviHtml.frx", ".frx の対の確認"),
    ("{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}", "Microsoft Internet Controls の参照設定"),
    ("archived_at", "案件一覧 26列目の追加"),
    ("tests_expected", "config の期待本数"),
    ("wintest\\tests_expected.txt", "期待本数を台帳から読むこと(直書き禁止)"),
    ("BN_DATA_KEYS", "data_key 33値を値源から読むこと(直書き禁止)"),
    ("'s1,s2,s3,s4,pf,sp,wt,fg,s2c,s3c,s2r,s3r,ch'", "run_log!step に ch を足すこと"),
    ("リスク提案ナビ.xlsm", "出力の名前を変えないこと"),
)
# 第2段のスクリプトに**書いてはいけない**もの(直書きの期待本数など)。
STAGE2_FORBIDDEN = (
    ("tests_expected='8", "期待本数の直書き(wintest/tests_expected.txt から読むこと)"),
    ("Risk-consulting-Navi", "髙橋さん側の製品名(当方の名前は「リスク提案ナビ」)"),
)


def check_stage2_script() -> tuple[bool, list[str]]:
    """第2段のスクリプトの存在と必須文字列(裁定書34 §1.1)。実行はしない。"""
    print("\n" + "=" * 78)
    print("裁定書34 §1.1 第2段のスクリプト(build/win/import_navi_modules.ps1)")
    print("=" * 78)
    problems: list[str] = []
    if not STAGE2_SCRIPT.exists():
        print(f"  FAIL: {STAGE2_SCRIPT.relative_to(REPO_ROOT)} がありません")
        return False, ["第2段のスクリプトがありません"]
    text = STAGE2_SCRIPT.read_text(encoding="utf-8", errors="replace")
    for needle, why in STAGE2_REQUIRED:
        if needle not in text:
            problems.append(f"第2段のスクリプトに {why} がありません(探した語: {needle})")
    for needle, why in STAGE2_FORBIDDEN:
        if needle in text:
            problems.append(f"第2段のスクリプトに書いてはいけないもの: {why}")
    print(f"  必須 {len(STAGE2_REQUIRED)}項目 / 禁止 {len(STAGE2_FORBIDDEN)}項目 : "
          f"{'OK' if not problems else f'{len(problems)}件 NG'}")
    for p in problems:
        print(f"  FAIL: {p}")
    return not problems, problems


def check_final(book: Path) -> tuple[bool, list[str]]:
    print("\n" + "=" * 78)
    print("裁定書34 §0.3 --final 第2段(開発PC)の産物の検査")
    print("=" * 78)

    problems: list[str] = []
    if not book.exists():
        print(f"  SKIP: {book} がありません"
              "(第2段は Windows 実Excel を持つ開発PCで作ります。CIでは作れません)")
        return True, []

    sys.path.insert(0, str(REPO_ROOT / "build"))
    import ovba_write   # noqa: E402
    import build_rpn    # noqa: E402  (禁止文字列の表を共有する)

    with zipfile.ZipFile(book) as z:
        names = z.namelist()
        if "xl/vbaProject.bin" not in names:
            print("  FAIL: xl/vbaProject.bin がありません")
            return False, [f"{book.name}: xl/vbaProject.bin がありません"]
        vba_bin = z.read("xl/vbaProject.bin")

    # --- (1) モジュール集合 ---------------------------------------------------
    got_info = ovba_write.read_modules(vba_bin)
    got = set(got_info.keys()) - {"ThisWorkbook"}
    # シートの文書モジュール(Sheet1 等)は Excel が勝手に作るので数えない。
    got = {n for n in got
           if got_info[n].get("type") != "document" or n == FINAL_FORM_NAME}
    want_std, want_forms = _final_expected_modules()
    want = want_std | want_forms
    missing = sorted(want - got)
    extra = sorted(got - want)
    print(f"  モジュール: 期待 {len(want)} / 実際 {len(got)}")
    for n in missing:
        problems.append(f"第2段の産物にモジュールがありません: {n}")
    for n in extra:
        problems.append(f"第2段の産物に台帳に無いモジュールがあります: {n}")
    if missing:
        print(f"    不足: {', '.join(missing)}")
    if extra:
        print(f"    余分: {', '.join(extra)}")

    # --- (2) 参照設定 ---------------------------------------------------------
    # dir ストリームの REFERENCENAME レコードから名前を拾う。ライブラリ名は
    # ASCII なので、bin から直接探すより dir を読むほうが誤検出が少ない。
    refs = sorted(set(ovba_write.read_reference_names(vba_bin))) \
        if hasattr(ovba_write, "read_reference_names") else None
    if refs is None:
        print("  参照設定  : 読み取り口がありません(ovba_write.read_reference_names 未実装)")
        problems.append("参照設定を読み取れません"
                        "(build/ovba_write.py に read_reference_names が要ります)")
    else:
        print(f"  参照設定  : {', '.join(refs) if refs else '(なし)'}")
        for r in refs:
            if r not in FINAL_ALLOWED_REFS:
                problems.append(f"許可していない参照設定があります: {r}"
                                f"(許可: {', '.join(FINAL_ALLOWED_REFS)})")
        for r in ("SHDocVw", "MSForms"):
            if r not in refs:
                problems.append(f"必要な参照設定がありません: {r}"
                                "(HTML画面が起動しません)")

    # --- (3) 禁止文字列(prod) ------------------------------------------------
    hits = build_rpn.forbidden_strings_in_bin(vba_bin, prod=True)
    print(f"  禁止文字列: {hits if hits else 'なし'}")
    for h in hits:
        problems.append(f"第2段の産物に配布禁止の文字列があります: {h}")

    # --- (4) ui/ の同梱 -------------------------------------------------------
    ui_dir = book.parent / "ui"
    missing_ui = [n for n in FINAL_UI_FILES if not (ui_dir / n).exists()]
    print(f"  ui/       : {len(FINAL_UI_FILES) - len(missing_ui)}/{len(FINAL_UI_FILES)} 本")
    for n in missing_ui:
        problems.append(f"第2段の産物と同じフォルダに ui/{n} がありません")

    for p in problems:
        print(f"  FAIL: {p}")
    if not problems:
        print("  PASS: 4条件(モジュール集合 / 参照設定6つ / 禁止文字列不在 / ui 5本)")
    return not problems, problems


def main() -> int:
    ap = argparse.ArgumentParser(
        description="リスク提案ナビ 出荷前検問(T-46 ①②③ + 裁定書27 ⑥ + 裁定書28 ⑦ + 裁定書31 ⑧)")
    ap.add_argument("--src", default=str(REPO_ROOT / "src"))
    ap.add_argument("--dist", default=str(REPO_ROOT / "dist"))
    ap.add_argument("--final", nargs="?", const="", default=None,
                    help="第2段(開発PC)の産物を検査する(裁定書34 §0.3)。"
                         "パスを省くと <dist>/final/リスク提案ナビ.xlsm。"
                         "ファイルが無ければ SKIP(CIを赤にしない)")
    args = ap.parse_args()

    ok1, _ = check_item1(Path(args.src).resolve())
    ok2, _ = check_item2(Path(args.dist).resolve())
    ok3, _ = check_item3()
    ok6, _ = check_item6(Path(args.dist).resolve())
    ok7, _ = check_item7(Path(args.dist).resolve())
    ok8, _ = check_item8(Path(args.dist).resolve())

    ok_final = None
    if args.final is not None:
        ok_script, _ = check_stage2_script()
        final_book = (Path(args.final).resolve() if args.final
                      else Path(args.dist).resolve() / FINAL_DIR_NAME / PROD_BOOK_NAME)
        ok_book, _ = check_final(final_book)
        ok_final = ok_script and ok_book

    print("\n" + "-" * 78)
    print(f"① 外部由来テキストの直書き検査 : {'PASS' if ok1 else 'FAIL'}")
    print(f"② キー走査(NFR-S2 4項目)      : {'PASS' if ok2 else 'FAIL'}")
    print(f"③ dist/成果物の除外の実効確認  : {'PASS' if ok3 else 'FAIL'}")
    print("④ wintest(実Excel)全PASS      : 未実施(Windows実機が必要)")
    print("⑤ modTestsPure 本数条件        : 未実施(実機。Linux側の同等確認は "
          "tools/run_lo_tests.py)")
    print(f"⑥ 配布物のAV表面積と経路の固定 : {'PASS' if ok6 else 'FAIL'}")
    print(f"⑦ 起動ランチャー(.bat)の形    : {'PASS' if ok7 else 'FAIL'}")
    print(f"⑧ 手順書のbatと配布batの一致   : {'PASS' if ok8 else 'FAIL'}")
    if ok_final is not None:
        print(f"⑨ 第2段の産物(--final)        : {'PASS/SKIP' if ok_final else 'FAIL'}")
    print("-" * 78)
    if ok_final is False:
        print("結果: NG(exit code 1) - 1つでも落ちたら出荷しない")
        return 1
    if ok1 and ok2 and ok3 and ok6 and ok7 and ok8:
        print("結果: ①②③⑥⑦⑧ PASS(exit code 0)。**出荷には④⑤の実機確認が別途必要です**")
        return 0
    print("結果: NG(exit code 1) - 1つでも落ちたら出荷しない")
    return 1


if __name__ == "__main__":
    sys.exit(main())
