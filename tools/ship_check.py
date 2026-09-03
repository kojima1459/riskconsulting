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
     検出パターン = sk-[A-Za-z0-9_-]{20,} ／ 連続32文字以上の [0-9a-f] ／
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
PAT_SK = re.compile(r"sk-[A-Za-z0-9_-]{20,}")
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
    with zipfile.ZipFile(book) as z:
        if "xl/vbaProject.bin" not in z.namelist():
            problems.append("配布物に xl/vbaProject.bin がありません")
            hits = []
        else:
            hits = build_rpn.forbidden_strings_in_bin(z.read("xl/vbaProject.bin"))
    print(f"  bin の禁止文字列   : {hits if hits else 'なし'}"
          f"  (表: {list(build_rpn.FORBIDDEN_BIN_STRINGS)})")
    if hits:
        problems.append("vbaProject.bin に配布禁止の文字列があります: "
                        + ", ".join(hits))

    if problems:
        print(f"  FAIL: {len(problems)} 件")
        for p in problems:
            print(f"    - {p}")
        return False, problems
    print("  PASS: llm_transport=ribbon / direct_api_base 不在 / mock 不在 / "
          "vba_src 不在 / 禁止文字列 不在")
    return True, []


def main() -> int:
    ap = argparse.ArgumentParser(description="リスク提案ナビ 出荷前検問(T-46 ①②③ + 裁定書27 ⑥)")
    ap.add_argument("--src", default=str(REPO_ROOT / "src"))
    ap.add_argument("--dist", default=str(REPO_ROOT / "dist"))
    args = ap.parse_args()

    ok1, _ = check_item1(Path(args.src).resolve())
    ok2, _ = check_item2(Path(args.dist).resolve())
    ok3, _ = check_item3()
    ok6, _ = check_item6(Path(args.dist).resolve())

    print("\n" + "-" * 78)
    print(f"① 外部由来テキストの直書き検査 : {'PASS' if ok1 else 'FAIL'}")
    print(f"② キー走査(NFR-S2 4項目)      : {'PASS' if ok2 else 'FAIL'}")
    print(f"③ dist/成果物の除外の実効確認  : {'PASS' if ok3 else 'FAIL'}")
    print("④ wintest(実Excel)全PASS      : 未実施(Windows実機が必要)")
    print("⑤ modTestsPure 本数条件        : 未実施(実機。Linux側の同等確認は "
          "tools/run_lo_tests.py)")
    print(f"⑥ 配布物のAV表面積と経路の固定 : {'PASS' if ok6 else 'FAIL'}")
    print("-" * 78)
    if ok1 and ok2 and ok3 and ok6:
        print("結果: ①②③⑥ PASS(exit code 0)。**出荷には④⑤の実機確認が別途必要です**")
        return 0
    print("結果: NG(exit code 1) - 1つでも落ちたら出荷しない")
    return 1


if __name__ == "__main__":
    sys.exit(main())
