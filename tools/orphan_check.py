#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""orphan_check.py - 孤児Public検査(裁定書38 班D §1(1))

================================================================================
なぜ要るか(伝書鳩20260912 Part1-2・C_evidence.md C-6 A-1):
    「作ったのに繋いでいない」は人の注意力では止まらない。実装した本人が
    コミットに「直した」と書いた状態で気付かなかった実例がある。
    src/**/*.bas の全Public(Sub/Function/Property/Const/変数)について、
    どこからも参照されていないものを機械で数える。

参照集合(すべて1つでもヒットすればセーフ):
    (a) 自モジュール外のコード、および全モジュールの文字列リテラル
        (OnAction/OnTime/Run の宛先文字列を含む。実装は「定義行」と
        「戻り値代入行」を除く全statementに対する \\bNAME\\b 走査で、
        コードと文字列リテラルの両方を一度にカバーする)
    (b) build/ tools/ docs/ 配下のテキストファイル(手順書・ビルド入力の
        文字列から呼ばれる入口を救済する)
    (c) 動的連結の救済(C_evidence A-1で確認した2パターン):
        - `"接頭辞" & 式` 型: 接頭辞文字列(モジュール修飾があれば末尾の
          ローカル名も)を「有効な接頭辞」として集め、その接頭辞で始まる
          識別子は救済する(modUIResearch.CopyPrompt / modUIGuide.ShowAdvanced)
        - `HandlerName("接頭辞", ...)` 型: 名前組立関数への第1引数の
          文字列リテラルも同様に「有効な接頭辞」として集める(modUICase6)

例外: 定義直前5行以内に `' @unused:理由` があれば SKIP(ERRORにしない)。

使い方:
    python3 tools/orphan_check.py
    python3 tools/orphan_check.py --path <dir>   # 検査対象を変える(テスト用)
    python3 tools/orphan_check.py --verbose       # 救済された候補も列挙
    exit code: 0 = 孤児(SKIP以外)0件 / 1 = 1件以上 / 2 = 自己テスト失敗
================================================================================
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

TOOLS_DIR = Path(__file__).resolve().parent
REPO_ROOT = TOOLS_DIR.parent
sys.path.insert(0, str(TOOLS_DIR))

import vba_lint  # noqa: E402  (既存の解析ヘルパを再利用)

DEFAULT_SRC_ROOT = REPO_ROOT / "src"

# 参照集合(b)を走査する外部ディレクトリ。テキストとして読めるものだけ。
EXTERNAL_DIRS = ("build", "tools", "docs")
EXTERNAL_EXTS = {".py", ".md", ".json", ".ps1", ".txt", ".bat", ".cfg",
                  ".yml", ".yaml", ".ini", ".csv"}

UNUSED_MARK = "@unused:"

# ---------------------------------------------------------------------------
# 宣言抽出
# ---------------------------------------------------------------------------
SUB_DEF = re.compile(r"^Public\s+Sub\s+(\w+)", re.IGNORECASE)
FUNC_DEF = re.compile(r"^Public\s+Function\s+(\w+)", re.IGNORECASE)
PROP_DEF = re.compile(r"^Public\s+Property\s+(?:Get|Let|Set)\s+(\w+)", re.IGNORECASE)
CONST_DEF_HEAD = re.compile(r"^Public\s+Const\s+(.+)$", re.IGNORECASE)
CONST_NAME = re.compile(r"(\w+)\s*(?:As\s+\w+\s*)?=")
TYPE_DEF = re.compile(r"^Public\s+Type\s+(\w+)", re.IGNORECASE)
# モジュールレベル変数(Sub/Function/Property/Const/Type/Enum を除く)。
VAR_DEF_HEAD = re.compile(
    r"^Public\s+(?!Sub\b|Function\b|Property\b|Const\b|Type\b|Enum\b|Declare\b)"
    r"(.+)$", re.IGNORECASE)
VAR_NAME = re.compile(r"(\w+)\s*(?:\([^)]*\))?\s*(?:As\s+[\w.]+)?")

KIND_LABELS = {"sub": "Sub", "func": "Function", "prop": "Property",
               "const": "Const", "type": "Type", "var": "変数"}


def extract_decls(stmt: str) -> list[tuple[str, str]]:
    """1statementから (kind, name) のリストを返す(空なら宣言でない)。"""
    m = SUB_DEF.match(stmt)
    if m:
        return [("sub", m.group(1))]
    m = FUNC_DEF.match(stmt)
    if m:
        return [("func", m.group(1))]
    m = PROP_DEF.match(stmt)
    if m:
        return [("prop", m.group(1))]
    m = CONST_DEF_HEAD.match(stmt)
    if m:
        return [("const", n) for n in CONST_NAME.findall(m.group(1))]
    m = TYPE_DEF.match(stmt)
    if m:
        return [("type", m.group(1))]
    m = VAR_DEF_HEAD.match(stmt)
    if m:
        rest = m.group(1)
        # `WithEvents X As Y` にも対応。
        rest = re.sub(r"^WithEvents\s+", "", rest, flags=re.IGNORECASE)
        names = [n for n in VAR_NAME.findall(rest) if n.lower() != "as"]
        return [("var", n) for n in names]
    return []


RETURN_ASSIGN_TMPL = r"^{name}\s*=(?!=)"


class Decl:
    __slots__ = ("module", "name", "kinds", "lineno", "unused_reason")

    def __init__(self, module, name, lineno):
        self.module = module
        self.name = name
        self.kinds = set()
        self.lineno = lineno
        self.unused_reason = None


def find_unused_reason(raw_lines: list[str], lineno: int) -> str | None:
    """定義行の直前5行以内に `' @unused:理由` があればその理由文字列を返す。

    直前の宣言の本体(`End Function` 等)やコードに行き当たったら、そこで
    走査を止める(関数どうしが近接していると、隣の宣言の @unused 注記を
    誤って拾ってしまうため。実測で発見した誤検知: modConfig.bas の孤児
    候補の1つの直前へ @unused を足したら、4行後方にある無関係な別の
    Public 関数まで SKIP になった)。

    注記: この docstring 自身が tools/ のテキストとして orphan_check の
    外部テキスト参照集合に含まれるため、既知の孤児候補の識別子を**この
    ファイル中に直接書いてはいけない**(書くと、その識別子はこの
    docstring 自身によって「参照されている」ことになり、二度と孤児として
    検出できなくなる。実測で発見した自己言及の罠)。
    """
    scanned = 0
    for i in range(lineno - 2, -1, -1):  # lineno は1始まり。直前行から遡る
        if scanned >= 5:
            break
        if i < 0 or i >= len(raw_lines):
            break
        line = raw_lines[i]
        stripped = line.strip()
        if stripped == "":
            scanned += 1
            continue
        if not stripped.startswith("'"):
            break  # コード行(前の宣言の End Function 等)に当たったら打ち切り
        scanned += 1
        idx = line.find(UNUSED_MARK)
        if idx >= 0:
            return line[idx + len(UNUSED_MARK):].strip()
    return None


# ---------------------------------------------------------------------------
# 動的連結の救済(2パターン)
# ---------------------------------------------------------------------------
CONCAT_PREFIX = re.compile(r'"([^"]+)"\s*&')
HANDLER_NAME_CALL = re.compile(r"\bHandlerName\s*\(\s*\"([^\"]+)\"", re.IGNORECASE)


def collect_dynamic_prefixes(all_statements: list[str]) -> set[str]:
    prefixes: set[str] = set()
    for stmt in all_statements:
        for lit in CONCAT_PREFIX.findall(stmt):
            prefixes.add(lit)
            if "." in lit:
                prefixes.add(lit.rsplit(".", 1)[-1])
        for lit in HANDLER_NAME_CALL.findall(stmt):
            prefixes.add(lit)
    # 空文字列や短すぎる接頭辞は誤救済(全部を救済してしまう)のもとなので除外。
    return {p for p in prefixes if len(p) >= 3}


# ---------------------------------------------------------------------------
# 外部テキスト(build/tools/docs)の読み込み
# ---------------------------------------------------------------------------
def load_external_text(repo_root: Path) -> str:
    chunks = []
    for d in EXTERNAL_DIRS:
        base = repo_root / d
        if not base.exists():
            continue
        for path in base.rglob("*"):
            if not path.is_file():
                continue
            if path.suffix.lower() not in EXTERNAL_EXTS:
                continue
            try:
                chunks.append(path.read_text(encoding="utf-8", errors="ignore"))
            except OSError:
                continue
    return "\n".join(chunks)


# ---------------------------------------------------------------------------
# 本検査
# ---------------------------------------------------------------------------
def run_checks(src_root: Path, verbose: bool) -> tuple[int, int, int]:
    """戻り値: (ERROR件数, SKIP件数, 救済件数)"""
    files = vba_lint.discover_module_files(src_root)
    if not files:
        print("[orphan_check] 対象ファイルがありません: %s" % src_root)
        return (1, 0, 0)

    modules = [vba_lint.load_module(f, src_root) for f in files]

    # module.vb_name -> raw_lines(定義行の @unused 探索用)
    raw_lines_by_module = {m.vb_name: m.raw_text.split("\n") for m in modules}

    # 宣言収集(モジュール内で同名が複数回出る = Property Get/Let/Set のペア)
    decls: dict[tuple[str, str], Decl] = {}
    # モジュールごとの statements(検索対象)
    module_stmts: dict[str, list[tuple[int, str]]] = {}
    all_stmt_texts: list[str] = []

    for info in modules:
        stmts = vba_lint.iter_statements(info.raw_text.split("\n"))
        module_stmts[info.vb_name] = stmts
        for lineno, stmt in stmts:
            all_stmt_texts.append(stmt)
            for kind, name in extract_decls(stmt):
                key = (info.vb_name, name)
                d = decls.get(key)
                if d is None:
                    d = Decl(info.vb_name, name, lineno)
                    decls[key] = d
                d.kinds.add(kind)

    # @unused 判定
    for (module, name), d in decls.items():
        d.unused_reason = find_unused_reason(raw_lines_by_module[module], d.lineno)

    # 動的連結の救済接頭辞
    dyn_prefixes = collect_dynamic_prefixes(all_stmt_texts)

    # 外部テキスト
    ext_text = load_external_text(REPO_ROOT)

    errors: list[Decl] = []
    skips: list[Decl] = []
    rescued: list[tuple[Decl, str]] = []

    for (module, name), d in decls.items():
        if d.unused_reason is not None:
            skips.append(d)
            continue

        # (a) 自モジュール外のコード + 全モジュールの文字列リテラル。
        #     ただし自モジュールの「定義行」と「戻り値代入行」は使用に数えない。
        pat = re.compile(r"\b%s\b" % re.escape(name), re.IGNORECASE)
        ret_pat = re.compile(RETURN_ASSIGN_TMPL.format(name=re.escape(name)),
                             re.IGNORECASE)
        used = False
        for mod_name, stmts in module_stmts.items():
            for lineno, stmt in stmts:
                if mod_name == module and lineno == d.lineno:
                    continue  # 定義行自身
                if mod_name == module and ret_pat.match(stmt.strip()):
                    continue  # 戻り値代入行
                if pat.search(stmt):
                    used = True
                    break
            if used:
                break

        if used:
            continue

        # (b) build/ tools/ docs/ のテキスト
        if pat.search(ext_text):
            continue

        # (c) 動的連結の救済
        rescue_hit = next((p for p in dyn_prefixes if name.startswith(p)), None)
        if rescue_hit is not None:
            rescued.append((d, rescue_hit))
            continue

        errors.append(d)

    errors.sort(key=lambda d: (d.module, d.lineno))
    skips.sort(key=lambda d: (d.module, d.lineno))
    rescued.sort(key=lambda t: (t[0].module, t[0].lineno))

    for d in skips:
        kind = "/".join(sorted(KIND_LABELS[k] for k in d.kinds))
        print("SKIP  %s:%d %s %s (@unused: %s)" %
              (d.module, d.lineno, kind, d.name, d.unused_reason))

    if verbose:
        for d, prefix in rescued:
            kind = "/".join(sorted(KIND_LABELS[k] for k in d.kinds))
            print("RESCUED %s:%d %s %s (動的連結の接頭辞: %r)" %
                  (d.module, d.lineno, kind, d.name, prefix))

    for d in errors:
        kind = "/".join(sorted(KIND_LABELS[k] for k in d.kinds))
        rel = _module_relpath(files, d.module, src_root)
        print("ERROR %s:%d %s %s (呼び出し元・文字列リテラル・build/tools/docs "
              "のいずれにも出現しません)" % (rel, d.lineno, kind, d.name))

    print("孤児Public候補(機械検出): %d件 / 動的連結で救済: %d件 / "
          "@unused でSKIP: %d件" % (len(errors) + len(rescued), len(rescued),
                                   len(skips)))
    return (len(errors), len(skips), len(rescued))


def _module_relpath(files: list[Path], vb_name: str, src_root: Path) -> str:
    for f in files:
        if f.stem == vb_name:
            try:
                return str(f.relative_to(REPO_ROOT))
            except ValueError:
                return str(f)
    return vb_name


# ---------------------------------------------------------------------------
# 自己テスト(骨抜き防止): 負例(未参照Publicを合成)で検出できるか。
# ---------------------------------------------------------------------------
def self_test() -> bool:
    cases = []

    # 宣言抽出
    cases.append(("Sub抽出", extract_decls("Public Sub Foo(ByVal x As Long)")
                  == [("sub", "Foo")]))
    cases.append(("Function抽出", extract_decls("Public Function Bar() As String")
                  == [("func", "Bar")]))
    cases.append(("Property抽出", extract_decls("Public Property Get Baz() As Long")
                  == [("prop", "Baz")]))
    cases.append(("Const抽出", extract_decls("Public Const A As Long = 1")
                  == [("const", "A")]))
    cases.append(("複数Const抽出",
                  sorted(extract_decls("Public Const A = 1, B = 2"))
                  == sorted([("const", "A"), ("const", "B")])))
    cases.append(("変数抽出", extract_decls("Public gCount As Long")
                  == [("var", "gCount")]))
    cases.append(("Sub誤検出しない(Function定義)",
                  ("sub", "Bar") not in extract_decls("Public Function Bar() As String")))

    # 動的連結の救済
    prefixes = collect_dynamic_prefixes([
        '"modUIResearch.CopyPrompt" & CStr(n2)',
        '"modUIGuide.ShowAdvanced" & CStr(i + 1)',
        'x = modUICase6.HandlerName("PasteInto", areaKey) & ";96"',
    ])
    cases.append(("動的連結 接頭辞(モジュール修飾込み)",
                  "modUIResearch.CopyPrompt" in prefixes))
    cases.append(("動的連結 接頭辞(ローカル名)", "CopyPrompt" in prefixes))
    cases.append(("動的連結 接頭辞(ShowAdvanced)", "ShowAdvanced" in prefixes))
    cases.append(("HandlerName第1引数の救済", "PasteInto" in prefixes))
    cases.append(("動的連結 短すぎる接頭辞は救済しない(誤爆防止)",
                  collect_dynamic_prefixes(['"ab" & x']) == set()))

    # @unused
    lines = ["' 何か", "' @unused: Phase2 予約(裁定書38)",
             "Public Function Reserved() As String", "End Function"]
    cases.append(("@unused検出", find_unused_reason(lines, 3) ==
                  "Phase2 予約(裁定書38)"))
    cases.append(("@unused無しはNone",
                  find_unused_reason(["Public Function X() As String"], 1) is None))
    # 近接する2宣言間での誤爆防止(実測で発見: 前の宣言の@unusedが次の宣言まで
    # 滲む事故があった)。
    bleed_lines = [
        "' @unused: reason A",
        "Public Function A() As Long",
        "    A = 1",
        "End Function",
        "",
        "Public Function B() As Long",  # lineno=6。直前3行は空行+End Function+実装行
        "    B = 2",
        "End Function",
    ]
    cases.append(("@unused 近接誤爆防止",
                  find_unused_reason(bleed_lines, 6) is None))
    cases.append(("@unused 正例(Aは検出できる)",
                  find_unused_reason(bleed_lines, 2) == "reason A"))

    bad = [name for name, ok in cases if not ok]
    for name in bad:
        print("  自己テスト NG: %s" % name)
    print("  自己テスト: %d/%d" % (len(cases) - len(bad), len(cases)))
    return not bad


def main() -> int:
    ap = argparse.ArgumentParser(description="孤児Public検査(裁定書38 班D)")
    ap.add_argument("--path", type=str, default=str(DEFAULT_SRC_ROOT),
                    help="検査対象ディレクトリ(既定: <repo>/src)")
    ap.add_argument("--verbose", action="store_true",
                    help="動的連結で救済した候補も列挙する")
    args = ap.parse_args()

    print("orphan_check: 孤児Public検査(裁定書38 班D §1(1))")
    src_root = Path(args.path).resolve()
    if not src_root.exists():
        print("[orphan_check] 対象ディレクトリが存在しません: %s" % src_root)
        return 1

    n_error, n_skip, n_rescued = run_checks(src_root, args.verbose)

    if not self_test():
        print("結果: 自己テスト失敗(検出器が壊れています)")
        return 2

    if n_error:
        print("結果: NG (孤児Public %d件)" % n_error)
        return 1
    print("結果: OK (孤児Public 0件 / SKIP %d件 / 動的連結救済 %d件)" %
          (n_skip, n_rescued))
    return 0


if __name__ == "__main__":
    sys.exit(main())
