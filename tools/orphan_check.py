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
        **自モジュール内の「自分の名札」は使用に数えない**(W15 Round2
        R2-16): `SRC & ".名前"` のような**ドットで始まる文字列リテラル**は
        ログの発生元表示であって呼び出しではない。これを使用に数えていた
        ため、「本番の呼出元が0件の新設関数が、自分のエラーログ行のおかげで
        緑になる」という穴があった(実測で発見)。OnAction/OnTime の宛先は
        `"モジュール名.名前"` という**完全な**文字列なので、この規則では
        落ちない(同一モジュール内の宛先登録は従来どおり救済される)。
    (b) build/ tools/ 配下のテキストファイル(手順書・ビルド入力の
        文字列から呼ばれる入口を救済する)
        **docs/ は救済しない**(W15 Round2 R1-01): 仕様書に名前を書くことは
        「使っている」ことではない。仕様を先に書く本PJの手順では、
        docs/ を救済集合に入れると**仕様書に書いた瞬間に検出不能**になる
        (fail-open)。docs/ にしか名前が無いものは `ERROR(docs-only)` として
        一覧に出す(救済はするが緑にはしない、ではなく**赤にする**)。
        **登記は「使用」ではない**(W15 Round2 T-M1): build/ tools/ の中にも
        「名前を表へ載せているだけ」の**台帳**がある。台帳への登記で救済して
        しまうと、docs/ を外したのと同じ fail-open が build/ tools/ 側に残る
        (実測: 呼出0件の Public 19本が `tools/vba_lint.py` の登記だけで緑に
        なっていた)。そこで救済する/しないの線を次のとおり引く:
          救済する = **そのテキストが実行時に VBA を呼ぶ**もの。
            - ビルドが焼き込む VBA ソース片(`build/build_rpn.py` の
              ThisWorkbook。ブックのイベントから ui 層を直接呼ぶ行がある。
              **関数名はここへ書かない**: 書くとこのファイル自身がその名前を
              救済してしまう=冒頭の自己言及の罠)
            - ツールが実行時に流し込む VBA ドライバ(`tools/render_report.py`
              `tools/render_proposal.py` `tools/run_lo_tests.py`)
            - 実行コマンドの宛先文字列(OnAction/OnTime/Run の宛先)
          救済しない = **名前を表へ載せるだけ**の登記。`REGISTRY_ONLY_TEXT`
            に列挙する(`tools/vba_lint.py` の `MODULE_REGISTRY` / `CONTRACT`
            と、モジュール台帳 `build/modules.json` 全体)。
        この線引きは「呼ばれているか」を見る本検査の目的そのものであり、
        登記の有無は 14章§6 との突合(vba_lint の契約検査)が別に見る。
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
    python3 tools/orphan_check.py --selftest      # 自己テスト+回帰網だけ(src は見ない)
    exit code: 0 = 孤児(SKIP以外)0件 / 1 = 1件以上 / 2 = 自己テスト失敗
================================================================================
"""
from __future__ import annotations

import argparse
import ast
import contextlib
import io
import re
import shutil
import sys
import tempfile
from pathlib import Path

TOOLS_DIR = Path(__file__).resolve().parent
REPO_ROOT = TOOLS_DIR.parent
sys.path.insert(0, str(TOOLS_DIR))

import vba_lint  # noqa: E402  (既存の解析ヘルパを再利用)

DEFAULT_SRC_ROOT = REPO_ROOT / "src"

# 参照集合(b)を走査する外部ディレクトリ。テキストとして読めるものだけ。
# docs/ は**救済しない**(W15 Round2 R1-01)。仕様書は「呼び出し」ではない。
EXTERNAL_DIRS = ("build", "tools")
# 救済はしないが、「仕様書にだけ名前がある」ことを一覧で言い当てるために読む。
DOC_DIRS = ("docs",)
EXTERNAL_EXTS = {".py", ".md", ".json", ".ps1", ".txt", ".bat", ".cfg",
                  ".yml", ".yaml", ".ini", ".csv"}

# 「登記(名前を表へ載せるだけ)」であって「使用」ではない外部テキスト
# (W15 Round2 T-M1)。キーは<repo>からの相対パス(区切りは "/")。
#   値 None       = そのファイル全体を救済集合から外す
#   値 (変数名,…) = その Python 代入ブロック(直前の見出しコメントを含む)だけを外す
# ここに載せる根拠は「実行時にその名前で VBA を呼ばない」こと。呼ぶもの
# (build_rpn.py の ThisWorkbook・render_*.py / run_lo_tests.py の VBA ドライバ)は
# 載せない。載せると本当の配線まで赤くなる(誤検知)。
REGISTRY_ONLY_TEXT: dict[str, tuple[str, ...] | None] = {
    # モジュール台帳。モジュール名とパスの登録簿であり、`_comment` / `_note` の
    # 散文も「どう作ったか」の記録であって呼び出しではない。
    "build/modules.json": None,
    # 公開契約表(14章§6の写し)とモジュール一覧。**登記の本丸**。
    "tools/vba_lint.py": ("MODULE_REGISTRY", "CONTRACT"),
}

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
# 自モジュール内の「自分の名札」(W15 Round2 R2-16)
# ---------------------------------------------------------------------------
STRING_LITERAL = re.compile(r'"[^"]*"')


def strip_self_labels(stmt: str, name: str) -> str:
    """自モジュール内の `SRC & ".名前"` を statement から落とす。

    落とすのは「**ドットで始まり、その名前だけで終わる**文字列リテラル」
    だけ(`".名前"`)。これはモジュール名定数と連結してログの発生元を作る
    書き方であり、呼び出しではない。OnAction/OnTime/Run の宛先は
    `"モジュール名.名前"` という完全な文字列なので、この規則には当たらず
    従来どおり使用に数える(同一モジュール内で宛先を登録する画面モジュールを
    赤くしないこと自体を自己テストで固定してある)。
    """
    pat = re.compile(r'^"\.%s"$' % re.escape(name), re.IGNORECASE)
    return STRING_LITERAL.sub(
        lambda m: '""' if pat.match(m.group(0)) else m.group(0), stmt)


# ---------------------------------------------------------------------------
# 登記ブロックの除去(W15 Round2 T-M1)
# ---------------------------------------------------------------------------
def blank_assignment_block(text: str, var_name: str) -> str:
    """Python ソースの `var_name = …` 代入ブロックを空行へ置き換える。

    代入の範囲は ast が返す実体(lineno..end_lineno)に、**直前の連続した
    コメント行**(その表の見出しコメント)を足した範囲。見出しコメントまで
    落とすのは、契約表の見出しが「required に入れない関数」の名前を列挙して
    いるためで、そこを残すと表本体だけ落としても救済が残る。

    行数は変えない(空行に置き換える)ので、他の照合の行番号は動かない。
    構文として読めないときは**ファイル全体を空にする**(fail-closed。
    読めないものを「使用あり」と見なす方向へ倒さない)。
    """
    try:
        tree = ast.parse(text)
    except SyntaxError:
        return ""
    lines = text.split("\n")
    for node in tree.body:
        targets: list[str] = []
        if isinstance(node, ast.Assign):
            targets = [t.id for t in node.targets if isinstance(t, ast.Name)]
        elif isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name):
            targets = [node.target.id]
        if var_name not in targets:
            continue
        start = node.lineno - 1
        end = (node.end_lineno or node.lineno) - 1
        while start > 0 and lines[start - 1].lstrip().startswith("#"):
            start -= 1
        for i in range(start, min(end + 1, len(lines))):
            lines[i] = ""
    return "\n".join(lines)


def strip_registry_text(rel_path: str, text: str) -> str | None:
    """登記だけのテキストを救済集合から落とす。None = ファイルごと落とす。"""
    if rel_path not in REGISTRY_ONLY_TEXT:
        return text
    names = REGISTRY_ONLY_TEXT[rel_path]
    if names is None:
        return None
    for name in names:
        text = blank_assignment_block(text, name)
    return text


# ---------------------------------------------------------------------------
# 外部テキスト(build/tools と docs)の読み込み
# ---------------------------------------------------------------------------
def load_external_text(repo_root: Path, dirs: tuple = EXTERNAL_DIRS) -> str:
    chunks = []
    for d in dirs:
        base = repo_root / d
        if not base.exists():
            continue
        for path in base.rglob("*"):
            if not path.is_file():
                continue
            if path.suffix.lower() not in EXTERNAL_EXTS:
                continue
            try:
                text = path.read_text(encoding="utf-8", errors="ignore")
            except OSError:
                continue
            try:
                rel = str(path.relative_to(repo_root)).replace("\\", "/")
            except ValueError:
                rel = path.name
            text = strip_registry_text(rel, text)
            if text is None:
                continue  # 登記だけのファイル(救済しない)
            chunks.append(text)
    return "\n".join(chunks)


# ---------------------------------------------------------------------------
# 本検査
# ---------------------------------------------------------------------------
def run_checks(src_root: Path, verbose: bool) -> tuple[int, int, int]:
    """戻り値: (ERROR件数, SKIP件数, 救済件数)

    ERROR件数には「docs/ にしか名前が無いもの」(ERROR(docs-only))も含む。
    """
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

    # 外部テキスト(救済する build/tools と、救済しない docs を分けて持つ)
    ext_text = load_external_text(REPO_ROOT, EXTERNAL_DIRS)
    doc_text = load_external_text(REPO_ROOT, DOC_DIRS)

    errors: list[Decl] = []
    doc_only: list[Decl] = []
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
                if mod_name == module:
                    if lineno == d.lineno:
                        continue  # 定義行自身
                    if ret_pat.match(stmt.strip()):
                        continue  # 戻り値代入行
                    # 自分の名札(`SRC & ".名前"`)は呼び出しではない(R2-16)
                    stmt = strip_self_labels(stmt, name)
                if pat.search(stmt):
                    used = True
                    break
            if used:
                break

        if used:
            continue

        # (b) build/ tools/ のテキスト(docs/ は救済しない = R1-01)
        if pat.search(ext_text):
            continue

        # (c) 動的連結の救済
        rescue_hit = next((p for p in dyn_prefixes if name.startswith(p)), None)
        if rescue_hit is not None:
            rescued.append((d, rescue_hit))
            continue

        # docs/ にしか名前が無いもの。救済せず、理由を分けて赤にする(R1-01)。
        if pat.search(doc_text):
            doc_only.append(d)
            continue

        errors.append(d)

    doc_only.sort(key=lambda d: (d.module, d.lineno))
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

    for d in doc_only:
        kind = "/".join(sorted(KIND_LABELS[k] for k in d.kinds))
        rel = _module_relpath(files, d.module, src_root)
        print("ERROR(docs-only) %s:%d %s %s (docs/ にしか名前がありません。"
              "仕様書への記載は配線ではないので救済しません。配線するか削除するか "
              "`' @unused:理由` を付けてください)" % (rel, d.lineno, kind, d.name))

    for d in errors:
        kind = "/".join(sorted(KIND_LABELS[k] for k in d.kinds))
        rel = _module_relpath(files, d.module, src_root)
        print("ERROR %s:%d %s %s (呼び出し元・文字列リテラル・build/tools "
              "のいずれにも出現しません)" % (rel, d.lineno, kind, d.name))

    n_error = len(errors) + len(doc_only)
    print("孤児Public候補(機械検出): %d件 (うち docs/ のみ %d件) / "
          "動的連結で救済: %d件 / @unused でSKIP: %d件"
          % (n_error + len(rescued), len(doc_only), len(rescued), len(skips)))
    return (n_error, len(skips), len(rescued))


def _module_relpath(files: list[Path], vb_name: str, src_root: Path) -> str:
    for f in files:
        if f.stem == vb_name:
            try:
                return str(f.relative_to(REPO_ROOT))
            except ValueError:
                return str(f)
    return vb_name


# ---------------------------------------------------------------------------
# 回帰テスト(W15 Round2 T-m3): 合成リポジトリで「赤くなる/ならない」の両方向を固定する。
# ---------------------------------------------------------------------------
# 以前はこの網が scratchpad の絶対パス直書きのスクリプトにしかなく、worktree を
# 消した時点で実行不能だった。**リポジトリの中**(この自己テスト)へ移し、毎回の
# ゲートで回るようにする。合成リポジトリの識別子は `ZzT` 接頭辞で始める
# (src/ に実在しない名前。実在名をこのファイルに書くと、このファイル自身が
#  tools/ の外部テキストとしてその名前を救済してしまう=冒頭の自己言及の罠)。
BAS_DOC_ONLY = (
    'Attribute VB_Name = "modZzT"\n'
    "Option Explicit\n"
    "\n"
    "Public Sub ZzTDocOnly()\n"
    "    Dim i As Long\n"
    "    i = 1\n"
    "End Sub\n"
)

BAS_SELF_LABEL = (
    'Attribute VB_Name = "modZzT"\n'
    "Option Explicit\n"
    'Private Const ZZT_SRC As String = "modZzT"\n'
    "\n"
    "Public Function ZzTNeverCalled() As Boolean\n"
    '    modLog.LogError "E0101", ZZT_SRC & ".ZzTNeverCalled", "bad"\n'
    "    ZzTNeverCalled = True\n"
    "End Function\n"
)

BAS_ONACTION = (
    'Attribute VB_Name = "modZzT"\n'
    "Option Explicit\n"
    "\n"
    "Public Sub ZzTWiredByOnAction()\n"
    "    Dim i As Long\n"
    "    i = 1\n"
    "End Sub\n"
    "\n"
    "Public Sub ZzTRegister()\n"
    '    sh.Buttons(1).OnAction = "modZzT.ZzTWiredByOnAction"\n'
    "    Call ZzTRegister2\n"
    "End Sub\n"
    "\n"
    "Public Sub ZzTRegister2()\n"
    "    Call ZzTRegister\n"
    "End Sub\n"
)

# `tools/vba_lint.py` の写し(登記だけ / 登記の外にドライバ文字列がある の2形)。
VBA_LINT_REGISTRY_ONLY = (
    "# 公開契約表の見出し。ZzTInHeadComment もここに名前だけがある。\n"
    "CONTRACT = {\n"
    '    "modZzT": {"closed": False, "required": ["ZzTDocOnly"]},\n'
    "}\n"
    "MODULE_REGISTRY = {\n"
    '    "modZzT",\n'
    "}\n"
)

VBA_LINT_WITH_DRIVER = VBA_LINT_REGISTRY_ONLY + (
    "\n"
    "# 登記の外。ここは実行時に VBA を呼ぶ文字列なので救済してよい。\n"
    'DRIVER = \'    Call modZzT.ZzTDocOnly\\n\'\n'
)


def _synth_repo(root: Path, bas_body: str, docs_text: str = "",
                tools_files: dict | None = None,
                build_files: dict | None = None) -> None:
    (root / "src" / "app").mkdir(parents=True, exist_ok=True)
    (root / "docs").mkdir(parents=True, exist_ok=True)
    (root / "tools").mkdir(parents=True, exist_ok=True)
    (root / "build").mkdir(parents=True, exist_ok=True)
    (root / "src" / "app" / "modZzT.bas").write_text(bas_body, encoding="utf-8")
    (root / "docs" / "zz_spec.md").write_text(docs_text, encoding="utf-8")
    for name, body in (tools_files or {}).items():
        (root / "tools" / name).write_text(body, encoding="utf-8")
    for name, body in (build_files or {}).items():
        (root / "build" / name).write_text(body, encoding="utf-8")


def _run_synth(bas_body: str, docs_text: str = "",
               tools_files: dict | None = None,
               build_files: dict | None = None) -> int:
    """合成リポジトリを検査して ERROR 件数だけ返す(出力は飲み込む)。"""
    global REPO_ROOT
    saved = REPO_ROOT
    tmp = Path(tempfile.mkdtemp(prefix="orphan_selftest_"))
    try:
        _synth_repo(tmp, bas_body, docs_text, tools_files, build_files)
        REPO_ROOT = tmp
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            n_error, _n_skip, _n_rescued = run_checks(tmp / "src", False)
        return n_error
    finally:
        REPO_ROOT = saved
        shutil.rmtree(tmp, ignore_errors=True)


def regression_cases() -> list[tuple[str, bool]]:
    """「条件が成立するときだけ赤くなる」の両方向を固定する。"""
    cases: list[tuple[str, bool]] = []

    # R1-01: docs/ にしか名前が無い Public は赤(仕様書の言及は配線ではない)。
    cases.append(("R1-01 docs/のみの言及では救済しない",
                  _run_synth(BAS_DOC_ONLY,
                             docs_text="14章: `ZzTDocOnly` は S5 を実行する。") == 1))
    # 逆方向: tools/ の普通のファイルの言及は従来どおり救済する(締めすぎ防止)。
    cases.append(("tools/ の普通のファイルの言及は救済",
                  _run_synth(BAS_DOC_ONLY,
                             tools_files={"zz_tool.py": "x = 'ZzTDocOnly'\n"}) == 0))
    # 逆方向: build/ の普通のファイルの言及も救済する。
    cases.append(("build/ の普通のファイルの言及は救済",
                  _run_synth(BAS_DOC_ONLY,
                             build_files={"zz_build.json":
                                          '{"onAction": "ZzTDocOnly"}'}) == 0))

    # T-M1 本丸: vba_lint の CONTRACT / MODULE_REGISTRY への登記だけでは救済しない。
    cases.append(("T-M1 vba_lint の登記だけでは救済しない",
                  _run_synth(BAS_DOC_ONLY,
                             tools_files={"vba_lint.py": VBA_LINT_REGISTRY_ONLY}) == 1))
    # 逆方向: 同じ vba_lint.py でも**登記の外**(VBAを呼ぶ文字列)なら救済する。
    cases.append(("T-M1 登記の外の呼び出し文字列は救済する",
                  _run_synth(BAS_DOC_ONLY,
                             tools_files={"vba_lint.py": VBA_LINT_WITH_DRIVER}) == 0))
    # T-M1: build/modules.json(モジュール台帳)への登記だけでも救済しない。
    cases.append(("T-M1 build/modules.json の登記だけでは救済しない",
                  _run_synth(BAS_DOC_ONLY,
                             build_files={"modules.json":
                                          '{"_note": "ZzTDocOnly を呼ぶ"}'}) == 1))
    # 逆方向: 同じ内容でもファイル名が違えば(台帳でなければ)救済する。
    cases.append(("T-M1 台帳でない build/ の同内容は救済する",
                  _run_synth(BAS_DOC_ONLY,
                             build_files={"zz_other.json":
                                          '{"_note": "ZzTDocOnly を呼ぶ"}'}) == 0))

    # R2-16: 自モジュールのログ用名札 `SRC & ".名前"` だけでは使用に数えない。
    cases.append(("R2-16 自分の名札だけの Public は赤", _run_synth(BAS_SELF_LABEL) == 1))
    # 逆方向: 同一モジュール内の**完全な** OnAction 宛先文字列は従来どおり救済。
    cases.append(("R2-16 同一モジュールのOnAction宛先は救済",
                  _run_synth(BAS_ONACTION) == 0))
    return cases


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

    # 自分の名札(W15 Round2 R2-16)。名札だけの出現は使用に数えない。
    label_stmt = 'modLog.LogError "E0101", MOD_SRC & ".Foo", "bad"'
    cases.append(("自分の名札 `SRC & \".名前\"` は落ちる",
                  strip_self_labels(label_stmt, "Foo") ==
                  'modLog.LogError "E0101", MOD_SRC & "", "bad"'))
    cases.append(("名札を落とすと \\bNAME\\b が残らない",
                  re.search(r"\bFoo\b", strip_self_labels(label_stmt, "Foo"),
                            re.IGNORECASE) is None))
    onaction_stmt = 'sh.Buttons(1).OnAction = "modZzExample.Foo"'
    cases.append(("同一モジュールのOnAction宛先は落ちない",
                  strip_self_labels(onaction_stmt, "Foo") == onaction_stmt))
    cases.append(("本物の呼び出しは落ちない",
                  strip_self_labels("Call Foo(1)", "Foo") == "Call Foo(1)"))
    cases.append(("別名の名札は落とさない",
                  strip_self_labels('x = SRC & ".FooBar"', "Foo") ==
                  'x = SRC & ".FooBar"'))

    # docs/ を救済集合から外したこと(W15 Round2 R1-01)の固定。
    cases.append(("救済する外部ディレクトリは build/ tools/ のみ",
                  tuple(EXTERNAL_DIRS) == ("build", "tools")))
    cases.append(("docs/ は救済集合ではない",
                  "docs" not in EXTERNAL_DIRS and tuple(DOC_DIRS) == ("docs",)))

    # 登記を救済集合から外したこと(W15 Round2 T-M1)の固定。
    src_py = (
        "HEAD = 1\n"
        "# 見出しコメント: ZzTInHeadComment\n"
        "# 2行目\n"
        "TABLE = {\n"
        '    "a": {"b": "}{"},   # ZzTInBlockComment と括弧 } を含む\n'
        '    "c": "# これはコメントではない",\n'
        "}\n"
        "TAIL = 'ZzTAfterTable'\n"
    )
    blanked = blank_assignment_block(src_py, "TABLE")
    cases.append(("登記ブロックは表本体ごと落ちる",
                  "ZzTInBlockComment" not in blanked))
    cases.append(("登記ブロックの見出しコメントも落ちる",
                  "ZzTInHeadComment" not in blanked))
    cases.append(("登記ブロックの外は落とさない",
                  "ZzTAfterTable" in blanked and "HEAD = 1" in blanked))
    cases.append(("登記ブロックを落としても行数は変わらない",
                  len(blanked.split("\n")) == len(src_py.split("\n"))))
    cases.append(("無い名前を指定しても何も落とさない",
                  blank_assignment_block(src_py, "NOT_THERE") == src_py))
    cases.append(("読めない Python は fail-closed(全部落とす)",
                  blank_assignment_block("def (:\n", "TABLE") == ""))
    cases.append(("登記だけのファイルはファイルごと落ちる",
                  strip_registry_text("build/modules.json", "ZzTLedger") is None))
    cases.append(("担当外のパスは素通し",
                  strip_registry_text("tools/zz_tool.py", "ZzTPlain") == "ZzTPlain"))
    cases.append(("登記リストは vba_lint の2表と台帳のみ",
                  sorted(REGISTRY_ONLY_TEXT) ==
                  ["build/modules.json", "tools/vba_lint.py"] and
                  REGISTRY_ONLY_TEXT["tools/vba_lint.py"] ==
                  ("MODULE_REGISTRY", "CONTRACT")))
    # 「実行時に VBA を呼ぶ」側は**外さない**(外すと本物の配線が赤くなる)。
    cases.append(("VBAを呼ぶ側は登記リストに入れない",
                  not any(p in REGISTRY_ONLY_TEXT for p in
                          ("build/build_rpn.py", "tools/run_lo_tests.py",
                           "tools/render_report.py", "tools/render_proposal.py"))))

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

    # 合成リポジトリの回帰網(W15 Round2 T-m3。以前は scratchpad にしか無かった)。
    cases.extend(regression_cases())

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
    ap.add_argument("--selftest", action="store_true",
                    help="自己テストと回帰網だけを回す(src は検査しない)")
    args = ap.parse_args()

    print("orphan_check: 孤児Public検査(裁定書38 班D §1(1))")
    if args.selftest:
        if not self_test():
            print("結果: 自己テスト失敗(検出器が壊れています)")
            return 2
        print("結果: 自己テストOK")
        return 0
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
