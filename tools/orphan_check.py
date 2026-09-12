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
    (b) build/ tools/ 配下のテキストファイルの**実行される部分**
        (手順書・ビルド入力の文字列から呼ばれる入口を救済する)
        **コメントと docstring は救済しない**(W15 §3 X3-2): docs/ を外した
        のと同じ理屈で、`tools/` `build/` の**散文**に名前が出ることも
        「使っている」ことではない。実測: 呼出0件の Public が
        `tools/*.py` の説明コメントや docstring・`tools/README.md` の地の文
        だけで緑になりうる状態だった(docs/ を外したのに同じ性質が残っていた)。
        救済に数えるのは次だけ(`strip_nonexecutable_text`):
          - .py … コメント(`#`)と docstring を落とした残り
                  (**文字列リテラルと実コードだけ**)
          - .md … フェンス付きコードブロックとインラインコード(`` ` ``)だけ
                  (地の文は散文=コメントと同じ)
          - .ps1/.bat/.yml/.ini/.cfg … 行コメント(`#` `REM` `::` `;`)を落とす
          - .json/.csv/.txt … データなのでそのまま(散文の器ではない)
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
    (c) 動的連結の救済(C_evidence A-1で確認した2パターン)。
        **救済は「宛先モジュール」と「接尾辞の形」の両方で閉じる**
        (W15 §3 X3-1)。以前は接頭辞の**前方一致だけ**で、モジュールも
        接尾辞も問わなかったため、src のどこかに `"ShowArea" & 式` が1つ
        あれば**無関係なモジュール**の `ShowAreaZzzNotWired` まで永久に
        救済された(統合レビューが実測で再現)。いまの規則:
        - `"モジュール名.接頭辞" & 式` 型 … 宛先 = その**モジュール名**、
          接頭辞 = ドットより後ろ(modUIResearch.CopyPrompt /
          modUIGuide.ShowAdvanced)
        - `"接頭辞" & 式` 型(モジュール修飾なし) … 宛先 = **その連結を
          書いているモジュール自身**
        - `モジュール名.HandlerName("接頭辞", …)` 型 … 宛先 = その
          モジュール(修飾が無ければ書いているモジュール。modUICase6)
        接尾辞(名前から接頭辞を除いた残り)は**必ず1文字以上**で、かつ
        - 連番型(`& CStr(n)` `& Format$(n, "00")` など数へ変換する式) …
          **数字だけ**
        - 名前組立型(それ以外の式・HandlerName) … 宛先モジュールか連結を
          書いたモジュールの**文字列リテラルに出てくる snake_case 語**を
          パスカル化した集合に**完全一致**(field_notes -> FieldNotes)
        でなければならない。

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

import gate_count  # noqa: E402  (要点行の契約。W15 §3 X3-3)
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
# 動的連結の救済(W15 §3 X3-1 で「宛先モジュール × 接尾辞の形」へ閉じた)
# ---------------------------------------------------------------------------
CONCAT_PREFIX = re.compile(r'"([^"]*)"\s*&')
# 名前組立関数(第1引数の文字列が接頭辞になる)。モジュール修飾は任意。
NAME_BUILDERS = ("HandlerName",)
HANDLER_NAME_CALL = re.compile(
    r"(?:(\w+)\s*\.\s*)?\b(?:%s)\s*\(\s*\"([^\"]*)\""
    % "|".join(NAME_BUILDERS), re.IGNORECASE)
# `& CStr(n)` `& Format$(n, "00")` のように**数へ変換する式**で終わる連結。
# このときハンドラ名の接尾辞は連番(数字だけ)にしかならない。
NUMERIC_TAIL = re.compile(
    r"^\s*(?:CStr|CLng|CInt|CDbl|CByte|Str\$?|Format\$?|Trim\$?)\s*\(|^\s*\d",
    re.IGNORECASE)
# 文字列リテラルの中から拾う snake_case 語(欄キー・data_key の類)。
SNAKE_TOKEN = re.compile(r"[a-z][a-z0-9]*(?:_[a-z0-9]+)*")

MIN_PREFIX_LEN = 3


def pascalize(token: str) -> str:
    """snake_case をハンドラ名の接尾辞と同じ形へ(field_notes -> FieldNotes)。

    製品側の組立(modUICase6.HandlerName)と同じ変換であることが救済の前提。
    自己テストが両者の一致を固定している。
    """
    return "".join(p[:1].upper() + p[1:] for p in token.split("_") if p)


def module_key_suffixes(stmts: list[tuple[int, str]]) -> set[str]:
    """そのモジュールの文字列リテラルに出てくる snake_case 語のパスカル形。"""
    out: set[str] = set()
    for _lineno, stmt in stmts:
        for lit in STRING_LITERAL.findall(stmt):
            for tok in SNAKE_TOKEN.findall(lit[1:-1]):
                out.add(pascalize(tok))
    return out


class DynRule:
    """動的連結1件ぶんの救済規則。

    module = 救済してよい**宛先モジュール**(ここ以外の同名接頭辞は救済しない)
    prefix = 接頭辞 / kind = "num"(連番) or "key"(名前組立)
    keys   = kind="key" のときに許す接尾辞の集合(完全一致)
    origin = この連結が書かれているモジュール(表示用)
    """

    __slots__ = ("module", "prefix", "kind", "keys", "origin")

    def __init__(self, module, prefix, kind, keys, origin):
        self.module = module
        self.prefix = prefix
        self.kind = kind
        self.keys = keys
        self.origin = origin

    def label(self) -> str:
        where = self.module if self.module == self.origin else (
            "%s <- %s" % (self.module, self.origin))
        return "%s + %s (%s)" % (
            self.prefix, "連番" if self.kind == "num" else "欄キー", where)


def collect_dynamic_rules(module_stmts: dict[str, list[tuple[int, str]]]
                          ) -> list[DynRule]:
    """全モジュールの statement から救済規則を組む(宛先モジュール付き)。"""
    known = set(module_stmts)
    keys_by_module = {m: module_key_suffixes(st) for m, st in module_stmts.items()}
    rules: list[DynRule] = []

    def allowed(target: str, origin: str) -> frozenset:
        return frozenset(keys_by_module.get(target, set())
                         | keys_by_module.get(origin, set()))

    # テストモジュールは最後に見る(救済の根拠としては本番の配線を先に出す)。
    ordered = sorted(module_stmts.items(),
                     key=lambda kv: (kv[0].lower().startswith("modtests"), kv[0]))
    for mod, stmts in ordered:
        for _lineno, stmt in stmts:
            for m in CONCAT_PREFIX.finditer(stmt):
                lit = m.group(1)
                kind = "num" if NUMERIC_TAIL.match(stmt[m.end():]) else "key"
                if "." in lit:
                    qual, local = lit.rsplit(".", 1)
                    qual = qual.rsplit(".", 1)[-1]
                    # 修飾が実在のモジュール名でなければ救済しない
                    # (`"copied_" & n` のようなキー組立を宛先と誤解しない)。
                    if qual not in known:
                        continue
                    target, prefix = qual, local
                else:
                    target, prefix = mod, lit
                if len(prefix) < MIN_PREFIX_LEN:
                    continue
                rules.append(DynRule(target, prefix, kind, allowed(target, mod), mod))
            for m in HANDLER_NAME_CALL.finditer(stmt):
                qual, lit = m.group(1), m.group(2)
                target = qual if (qual and qual in known) else mod
                if len(lit) < MIN_PREFIX_LEN:
                    continue
                rules.append(DynRule(target, lit, "key", allowed(target, mod), mod))
    return rules


def rescue_rule_for(module: str, name: str,
                    rules: list[DynRule]) -> DynRule | None:
    """その Public を救済してよい規則(無ければ None)。VBA は大小同一視。"""
    low = name.lower()
    for r in rules:
        if r.module != module:
            continue
        if not low.startswith(r.prefix.lower()):
            continue
        suffix = name[len(r.prefix):]
        if not suffix:
            continue  # 接頭辞そのものは動的連結では作られない
        if r.kind == "num":
            if suffix.isdigit():
                return r
            continue
        if suffix.lower() in {k.lower() for k in r.keys}:
            return r
    return None


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


# ---------------------------------------------------------------------------
# 実行されない散文の除去(W15 §3 X3-2)
# ---------------------------------------------------------------------------
# docs/ を救済集合から外した理由(「名前を書くことは使うことではない」)は、
# build/ tools/ の**コメント・docstring・README の地の文**にもそのまま当たる。
# ここで落とすのは「人へ向けた散文」だけで、文字列リテラル・コード・データは
# 残す(残さないと本物の配線まで赤くなる)。
LINE_COMMENT_EXTS = {".ps1": ("#",), ".bat": ("REM ", "::", "rem "),
                     ".yml": ("#",), ".yaml": ("#",),
                     ".ini": ("#", ";"), ".cfg": ("#", ";")}
MD_FENCE = re.compile(r"^\s*(?:```|~~~)")
MD_INLINE_CODE = re.compile(r"`([^`\n]+)`")


def strip_py_prose(text: str) -> str:
    """Python から コメントと docstring を落とす(行数は変えない)。

    読めない Python は **fail-closed**(全部落とす)。読めないものを
    「使用あり」の方向へ倒さない(blank_assignment_block と同じ方針)。
    """
    import tokenize  # 局所import(このツールの他の経路では使わない)

    try:
        toks = list(tokenize.generate_tokens(io.StringIO(text).readline))
    except (tokenize.TokenError, IndentationError, SyntaxError):
        return ""
    lines = text.split("\n")

    def blank(srow, scol, erow, ecol):
        for r in range(srow, erow + 1):
            i = r - 1
            if i < 0 or i >= len(lines):
                continue
            a = scol if r == srow else 0
            b = ecol if r == erow else len(lines[i])
            lines[i] = lines[i][:a] + " " * max(0, b - a) + lines[i][b:]

    for tok in toks:
        if tok.type == tokenize.COMMENT:
            blank(tok.start[0], tok.start[1], tok.end[0], tok.end[1])
    # docstring = 式文になっている文字列定数(モジュール/クラス/関数のどこでも)。
    try:
        tree = ast.parse(text)
    except SyntaxError:
        return ""
    for node in ast.walk(tree):
        if isinstance(node, ast.Expr) and isinstance(node.value, ast.Constant) \
                and isinstance(node.value.value, str):
            c = node.value
            blank(c.lineno, c.col_offset,
                  c.end_lineno or c.lineno, c.end_col_offset or 0)
    return "\n".join(lines)


def strip_md_prose(text: str) -> str:
    """Markdown はコードブロックとインラインコードだけを残す(地の文=散文)。"""
    kept: list[str] = []
    in_fence = False
    for line in text.split("\n"):
        if MD_FENCE.match(line):
            in_fence = not in_fence
            continue
        if in_fence:
            kept.append(line)
        else:
            kept.extend(MD_INLINE_CODE.findall(line))
    return "\n".join(kept)


def strip_line_comments(text: str, markers: tuple) -> str:
    out = []
    for line in text.split("\n"):
        st = line.lstrip()
        if any(st.startswith(mk) for mk in markers):
            continue
        out.append(line)
    return "\n".join(out)


def strip_nonexecutable_text(rel_path: str, text: str) -> str:
    """救済集合から「実行されない散文」を落とす(拡張子ごと)。"""
    ext = ("." + rel_path.rsplit(".", 1)[-1]).lower() if "." in rel_path else ""
    if ext == ".py":
        return strip_py_prose(text)
    if ext == ".md":
        return strip_md_prose(text)
    if ext in LINE_COMMENT_EXTS:
        return strip_line_comments(text, LINE_COMMENT_EXTS[ext])
    return text  # .json / .csv / .txt はデータ。散文の器ではないのでそのまま。


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
def load_external_text(repo_root: Path, dirs: tuple = EXTERNAL_DIRS,
                       strip_prose: bool = True) -> str:
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
            if strip_prose:
                text = strip_nonexecutable_text(rel, text)
            chunks.append(text)
    return "\n".join(chunks)


# ---------------------------------------------------------------------------
# 本検査
# ---------------------------------------------------------------------------
def run_checks(src_root: Path, verbose: bool,
               checked: "gate_count.Checked | None" = None
               ) -> tuple[int, int, int]:
    """戻り値: (ERROR件数, SKIP件数, 救済件数)

    ERROR件数には「docs/ にしか名前が無いもの」(ERROR(docs-only))も含む。
    checked を渡すと**実際に検査した項目数**を積む(W15 §3 X3-3)。
    """
    if checked is None:
        checked = gate_count.Checked()
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

    for info in modules:
        stmts = vba_lint.iter_statements(info.raw_text.split("\n"))
        module_stmts[info.vb_name] = stmts
        for lineno, stmt in stmts:
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

    # 動的連結の救済規則(宛先モジュール × 接尾辞の形。X3-1)
    dyn_rules = collect_dynamic_rules(module_stmts)

    # 外部テキスト(救済する build/tools と、救済しない docs を分けて持つ)
    ext_text = load_external_text(REPO_ROOT, EXTERNAL_DIRS)
    # docs/ は「名前がそこにしか無い」を言い当てるためだけに読むので、
    # 散文を落とさない(落とすと docs-only の診断そのものが効かなくなる)。
    doc_text = load_external_text(REPO_ROOT, DOC_DIRS, strip_prose=False)

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

        # (c) 動的連結の救済(宛先モジュールと接尾辞の形が合うものだけ)
        rule = rescue_rule_for(module, name, dyn_rules)
        if rule is not None:
            rescued.append((d, rule.label()))
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
            print("RESCUED %s:%d %s %s (動的連結: %s)" %
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

    checked.record("モジュール", len(modules))
    checked.record("Public宣言", len(decls))
    checked.record("動的連結規則", len(dyn_rules))

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

    # 動的連結の救済(W15 §3 X3-1: 宛先モジュール × 接尾辞の形で閉じる)。
    # **識別子は `ZzT` の合成名だけを書く**(冒頭の自己言及の罠。実在の
    # Public 名をこのファイルに書くと、このファイル自身が tools/ の外部
    # テキストとしてその名前を救済してしまう。実測で1度踏んだ)。
    synth_stmts = {
        "modZzTa": [
            (10, '"modZzTa.ZzTCopyIt" & CStr(n2)'),
            (11, 'Public Sub ZzTCopyIt1()'),
        ],
        "modZzTb": [(20, '"modZzTb.ZzTShowIt" & CStr(i + 1)')],
        "modZzTdraw": [
            (30, 'x = modZzTc.HandlerName("ZzTPasteTo", areaKey) & ";96"')],
        "modZzTc": [
            (40, 's = s & "field_notes|input_memo|現場メモ" & vbLf'),
            (41, 'Public Function HandlerName(ByVal stem As String)')],
        "modZzTother": [(50, 'Dim i As Long')],
    }
    rules = collect_dynamic_rules(synth_stmts)

    def rescued_name(module, name):
        return rescue_rule_for(module, name, rules) is not None

    # 正例(いま実際に救済されている3型が、宛先モジュールで救済されること)
    cases.append(("動的連結 連番(同一モジュール)",
                  rescued_name("modZzTa", "ZzTCopyIt1")))
    cases.append(("動的連結 連番(別の接頭辞)",
                  rescued_name("modZzTb", "ZzTShowIt5")))
    cases.append(("動的連結 名前組立(HandlerName+欄キー)",
                  rescued_name("modZzTc", "ZzTPasteToFieldNotes")))
    # X3-1(1) 宛先モジュールで閉じる: 同じ接頭辞でも**別モジュール**は救済しない。
    cases.append(("X3-1 無関係モジュールの同名接頭辞は救済しない",
                  not rescued_name("modZzTother", "ZzTCopyIt1")))
    cases.append(("X3-1 無関係モジュールの名前組立も救済しない",
                  not rescued_name("modZzTother", "ZzTPasteToFieldNotes")))
    # X3-1(2) 接尾辞の形で閉じる: 連番の宛先に非数字の接尾辞は通さない。
    cases.append(("X3-1 連番の宛先に非数字の接尾辞は救済しない",
                  not rescued_name("modZzTa", "ZzTCopyItNotWired")))
    # X3-1(3) 名前組立の宛先でも、欄キー集合に無い接尾辞は通さない。
    cases.append(("X3-1 未知の欄キーの接尾辞は救済しない",
                  not rescued_name("modZzTc", "ZzTPasteToNotWired")))
    cases.append(("X3-1 接頭辞そのもの(接尾辞なし)は救済しない",
                  not rescued_name("modZzTa", "ZzTCopyIt")))
    # 短すぎる接頭辞は規則にしない(誤爆防止。従来どおり)。
    cases.append(("動的連結 短すぎる接頭辞は規則にしない",
                  collect_dynamic_rules({"modZzTa": [(1, '"ab" & x')]}) == []))
    # 修飾が実在モジュールでない連結を宛先と誤解しない。
    cases.append(("動的連結 実在しない修飾は宛先にしない",
                  all(r.module == "modZzTa" for r in
                      collect_dynamic_rules({"modZzTa": [(1, '"copied_x.y" & n')]}))))
    # 接尾辞の作り方が製品側(欄キーのパスカル化)と同じであること。
    cases.append(("パスカル化が製品の組立と同じ",
                  pascalize("hearing_answers") == "HearingAnswers" and
                  pascalize("dossier") == "Dossier"))
    cases.append(("欄キー集合は文字列リテラルの snake_case から作る",
                  "FieldNotes" in module_key_suffixes(
                      [(1, 's = "dossier|input_dossier|x" & vbLf & "field_notes|y"')])))

    # X3-2: build/ tools/ の**散文**では救済しない(コメント・docstring・地の文)。
    cases.append(("X3-2 Python のコメントは落ちる",
                  "ZzTInPyComment" not in
                  strip_nonexecutable_text("tools/z.py", "x = 1  # ZzTInPyComment\n")))
    cases.append(("X3-2 Python の docstring は落ちる",
                  "ZzTInDocstring" not in
                  strip_nonexecutable_text("tools/z.py",
                                           '"""ZzTInDocstring を呼ぶ。"""\nx = 1\n')))
    cases.append(("X3-2 関数の docstring も落ちる",
                  "ZzTInFuncDoc" not in
                  strip_nonexecutable_text("tools/z.py",
                                           'def f():\n    "ZzTInFuncDoc"\n    return 1\n')))
    cases.append(("X3-2 Python の文字列リテラルは残る",
                  "ZzTInPyString" in
                  strip_nonexecutable_text("tools/z.py", "x = 'ZzTInPyString'\n")))
    cases.append(("X3-2 Python のコードは残る",
                  "ZzTPyIdent" in
                  strip_nonexecutable_text("tools/z.py", "ZzTPyIdent = 1\n")))
    cases.append(("X3-2 読めない Python は fail-closed",
                  strip_nonexecutable_text("tools/z.py", "def (:\n") == ""))
    cases.append(("X3-2 Markdown の地の文は落ちる",
                  "ZzTInProse" not in
                  strip_nonexecutable_text("tools/README.md", "ZzTInProse を呼びます。\n")))
    cases.append(("X3-2 Markdown のコードブロックは残る",
                  "ZzTInFence" in
                  strip_nonexecutable_text("tools/README.md",
                                           "説明\n```\nZzTInFence\n```\n")))
    cases.append(("X3-2 Markdown のインラインコードは残る",
                  "ZzTInBacktick" in
                  strip_nonexecutable_text("tools/README.md",
                                           "手順: `ZzTInBacktick` を押す。\n")))
    cases.append(("X3-2 ps1 の行コメントは落ちる",
                  "ZzTInPs1Comment" not in
                  strip_nonexecutable_text("build/win/z.ps1", "# ZzTInPs1Comment\n$a=1\n")))
    cases.append(("X3-2 json はデータなのでそのまま",
                  "ZzTInJson" in
                  strip_nonexecutable_text("build/z.json", '{"a": "ZzTInJson"}')))

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
    _SELFTEST_N[0] = 0 if bad else len(cases)
    return not bad


# 直近の自己テストで**実際に通った本数**(0 = 失敗 or 未実行)。要点行に出す。
_SELFTEST_N = [0]


def self_test_count() -> int:
    _SELFTEST_N[0] = 0
    self_test()
    return _SELFTEST_N[0]


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

    checked = gate_count.Checked()
    n_error, n_skip, n_rescued = run_checks(src_root, args.verbose, checked)

    n_self = self_test_count()
    if n_self <= 0:
        print("結果: 自己テスト失敗(検出器が壊れています)")
        return 2
    checked.record("自己テスト", n_self)

    # 「検査していないのに緑」を止める要点行(W15 §3 X3-3)。**実際に見た数**
    # だけを名乗り、0件なら赤で止まる。
    if gate_count.report(checked, required=("Public宣言", "自己テスト")):
        return 1

    if n_error:
        print("結果: NG (孤児Public %d件)" % n_error)
        return 1
    print("結果: OK (孤児Public 0件 / SKIP %d件 / 動的連結救済 %d件)" %
          (n_skip, n_rescued))
    return 0


if __name__ == "__main__":
    sys.exit(main())
