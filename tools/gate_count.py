#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""gate_count.py - 「検査していないのに緑」を止める共通の要点行(W15 §3 X3-3)

================================================================================
なぜ要るか(統合レビュー「ゲートの死角」):
    `tools/gate.py` は **returncode だけ**で赤緑を決め、要点行(`pick_summary`)は
    表示専用である。そのため各ツールが出す成功文言が **固定の作文**だと、
    実際には飛ばした検査まで「確認しました」と名乗れてしまう。実測された例:

      - `render_proposal.py` / `render_report.py` が node 不在で DOM 検査を
        まるごと飛ばしたのに「描画後DOM…を確認しました」と出して exit 0
      - `render_proposal.py` の対訳表検査が VBA の実装を一度も見ていないのに
        「対訳表との一致を確認しました」と出す
      - `notice_check.py --only dom` で検査②が「(未実行)」のまま OK 行が出る

    どれも「人が読む文言」と「機械が見る事実」がつながっていないことが原因。

この仕組み(3つだけ):
    (1) **数える**: 各ツールは検査した項目を `Checked.record(名前, 件数)` で
        積む。**実際に回した検査だけ**を積む(飛ばした検査は積まない)。
    (2) **名乗る**: `report()` が要点行を1行だけ出す。
            検査実施: 合計件 (名前=件, 名前=件, …)
        ここに出る名前は「実際に回した検査」だけなので、飛ばせば名前が消え、
        件数も落ちる。**固定の作文ができない**。
    (3) **落ちる**: 合計0件、または必須の検査が0件なら `report()` は非0を
        返し、ツールはそのまま exit する(returncode で赤になる=fail-closed)。
        gate.py 側は要点行に `検査実施: [1-9]\\d*件` を登録すれば、
        「要点行なし」=検査の証拠なし を目で見つけられる(登録行は司令塔)。

もう1つの役目(件数がリテラル定数でないこと。裁定書43 §2 Y-7):
    (1) の「数える」は、**その場で数えた実測値**でなければ意味がない。件数が
    `record("描画後DOM", n_dom + 14)` のようにリテラル定数を含んでいると、
    検査本体から判定を14本抜いても要点行が1文字も変わらず exit 0 になる
    (検証者が render_report.py と notice_check.py の2本で実証した)。これは
    本件の出発点である「描画後DOM18本を確認しました」という嘘の再生産なので、
    `record(名前, 件数)` の**件数の式に数値リテラルが現れないこと**を構文木で
    見る。直せない分は `LITERAL_COUNT_PENDING` へ**理由付きで登記**し、登記が
    陳腐化(定数が消えた)したら赤にする。

もう1つの役目(登記の網羅):
    `python3 tools/gate_count.py` は `tools/gate.py` の GATES を読み、
    **すべてのゲートの実行ファイル**が
      - 本契約に適合している(CONTRACT_TOOLS)か
      - 未適用として**理由付きで登記**されている(PENDING_TOOLS)か
    のどちらかであることを確かめる。どちらにも無いスクリプトが GATES に
    現れたら赤(新しい検問が黙って契約の外へ出るのを止める)。登記だけ残って
    GATES から消えたものも赤(掃除漏れを残さない)。

使い方:
    python3 tools/gate_count.py              # 登記の網羅 + 自己テスト
    python3 tools/gate_count.py --selftest   # 自己テストだけ
    exit code: 0 = OK / 1 = 契約違反 / 2 = 自己テスト失敗
================================================================================
"""
from __future__ import annotations

import argparse
import ast
import re
import sys
from pathlib import Path

TOOLS_DIR = Path(__file__).resolve().parent
REPO_ROOT = TOOLS_DIR.parent

# 要点行の書式(gate.py の登録行はこの形を見る)。
CHECKED_PREFIX = "検査実施: "
CHECKED_LINE = re.compile(r"^検査実施: (\d+)件")
# gate.py へ推奨する登録パターン(0件なら要点行として拾えない)。
GATE_PATTERN = r"検査実施: [1-9]\d*件"

# 本契約を適用済みのツール(W15 §3 X3-3 の第1弾)。
CONTRACT_TOOLS = {
    "orphan_check.py",
    "doc_gate.py",
    "render_report.py",
    "notice_check.py",
}

# 未適用の登記。**理由を書かずにここへ足さない**。
# 裁定書42 §3.4 は「班X1/X2 が触るファイルには手を出さない」と決めているので、
# render_proposal.py はこの波では班X1/X2 への handoff として登記する。
PENDING_TOOLS: dict[str, str] = {
    "render_proposal.py": "W15 §3.4: 班X1/X2 の担当ファイル。要点行の件数化は handoff",
    "vba_lint.py": "件数(ERROR 件数と対象モジュール数)を既に要点行に出している",
    "run_lo_tests.py": "PASS/FAIL/SKIP の実数を要点行に出している",
    "build_rpn.py": "ビルド。検査ではなく生成物を作る",
    "sheet_check.py": "「全N項目一致」で実数を出している",
    "ship_check.py": "PASS/FAIL の実数を出している",
    "bin_roundtrip.py": "条件数を要点行に出している",
    "lo_xlsm.py": "条件数を要点行に出している",
    "prompt_diff.py": "「一致: N件」で実数を出している",
    "dossier_check.py": "「全N本」で実数を出している",
    "config_check.py": "5点一致。件数化は次波",
    "action_check.py": "条件数を要点行に出している",
    "validate_check.py": "「計N件」で実数を出している",
    "enum_check.py": "件数化は次波",
    "caption_check.py": "「全N本」で実数を出している",
    "t48_check.py": "条件数を要点行に出している",
    "ribbon_wire_check.py": "件数化は次波",
    "ui_check.py": "条件数を要点行に出している",
    "gate_count.py": "本ファイル自身(契約の登記を見る側)",
    "bench_s1.py": "採点器のベンチ。selfcheck は変異注入を1件ずつ、"
                   "gold-check は「5社の gold を検算」と実数を出している",
}

# 件数にリテラル定数が残っているものの登記(裁定書43 §2 Y-7)。
# **理由を書かずにここへ足さない**。裁定書43 §2 の担当表は班Y2 の担当を
# notice_check.py / doc_gate.py / orphan_check.py / gate_count.py に限っている
# ので、render_report.py は handoff として登記する。
LITERAL_COUNT_PENDING: dict[str, str] = {
    "render_report.py": "裁定書43 §2 handoff: 班Y2 の担当ファイル外。"
                        "免責の態=4 / 描画後DOM=n_dom+14 / DATAリテラル=1 の"
                        "3箇所を実測へ直すのは別担当",
}


# ---------------------------------------------------------------------------
# (1)(2)(3) 数える・名乗る・落ちる
# ---------------------------------------------------------------------------
class Checked:
    """実際に回した検査だけを積む数え上げ。

    `record` を呼ばなかった検査は要点行に**名前ごと出ない**。これが
    「飛ばしたのに確認したと名乗る」を機械的に不可能にしている。
    """

    __slots__ = ("items",)

    def __init__(self) -> None:
        self.items: list[tuple[str, int]] = []

    def record(self, name: str, count: int) -> int:
        """検査1種を積む(count = 実際に見た項目数)。戻り値は count。"""
        if count < 0:
            raise ValueError("検査件数が負です: %s=%d" % (name, count))
        self.items.append((name, int(count)))
        return count

    def total(self) -> int:
        return sum(n for _name, n in self.items)

    def count_of(self, name: str) -> int:
        return sum(n for nm, n in self.items if nm == name)

    def names(self) -> list[str]:
        return [nm for nm, _n in self.items]

    def line(self) -> str:
        body = ", ".join("%s=%d" % (nm, n) for nm, n in self.items)
        return "%s%d件 (%s)" % (CHECKED_PREFIX, self.total(), body or "なし")


def report(checked: Checked, required: tuple = (), prefix: str = "") -> int:
    """要点行を出し、検査が空なら非0を返す(fail-closed)。

    required に挙げた名前が1件も回っていなければ、合計が0でなくても赤にする
    (「片方の検査だけ回して全部やったように見せる」を止める)。
    """
    print(prefix + checked.line())
    missing = [nm for nm in required if checked.count_of(nm) <= 0]
    if checked.total() <= 0:
        print(prefix + "結果: NG (検査を1件も実行していません。"
                       "実行できない検査を緑にはしません)")
        return 1
    if missing:
        print(prefix + "結果: NG (必須の検査を実行していません: %s)"
              % ", ".join(missing))
        return 1
    return 0


# ---------------------------------------------------------------------------
# 登記の網羅(gate.py の GATES を読む)
# ---------------------------------------------------------------------------
def gate_scripts(gate_src: str) -> list[tuple[str, str]]:
    """gate.py のソースから (ゲート名, スクリプトのファイル名) を抜く。

    GATES は `sys.executable` を含むので実行はせず、**構文木から**読む。
    """
    tree = ast.parse(gate_src)
    out: list[tuple[str, str]] = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Assign):
            continue
        if not any(isinstance(t, ast.Name) and t.id == "GATES"
                   for t in node.targets):
            continue
        if not isinstance(node.value, (ast.List, ast.Tuple)):
            continue
        for elt in node.value.elts:
            if not isinstance(elt, (ast.Tuple, ast.List)) or len(elt.elts) < 2:
                continue
            name_node, cmd_node = elt.elts[0], elt.elts[1]
            if not isinstance(name_node, ast.Constant):
                continue
            if not isinstance(cmd_node, (ast.List, ast.Tuple)):
                continue
            script = ""
            for a in cmd_node.elts:
                if isinstance(a, ast.Constant) and isinstance(a.value, str) \
                        and a.value.endswith(".py"):
                    script = a.value.replace("\\", "/").rsplit("/", 1)[-1]
                    break
            if script:
                out.append((str(name_node.value), script))
    return out


# 自己テストの器(合成の Checked に作り物の件数を積む)は「検査の件数」では
# ないので見ない。ここだけが例外で、本物の検査は必ず見る。
FIXTURE_FUNCS = ("self_test", "self_test_count", "regression_cases")


def _literal_in_count(node: "ast.AST",
                      const_names: "frozenset[str] | None" = None) -> bool:
    """件数の式に**数値リテラルの値**が混ざっているか。

    添字(`_WIRE_N[0]` の 0)は件数ではなく置き場所なので数えない。
    `len(rows)` のような実測は当然数えない。`4` や `n + 14` は数える。

    `const_names` は「その場で数えていない変数」の名前(`_const_names`)。
    `n = 14` と書いてから `record("x", n)` と呼ぶ**変数経由の抜け道**を塞ぐ
    (裁定書43 §2 Y-7 の司令塔手直し。以前は呼び出し行だけを見ていたので
    素通りした)。
    """
    if isinstance(node, ast.Constant):
        return isinstance(node.value, (int, float)) \
            and not isinstance(node.value, bool)
    if isinstance(node, ast.Name):
        return bool(const_names) and node.id in const_names
    if isinstance(node, ast.Subscript):
        return _literal_in_count(node.value, const_names)
    return any(_literal_in_count(ch, const_names)
               for ch in ast.iter_child_nodes(node))


def _is_const_expr(node: "ast.AST", known: "set[str]") -> bool:
    """式が「その場で数えていない定数」か(数値リテラルと定数変数だけで出来て
    いるか)。`len(...)`・添字・属性・内包表記が1つでも混ざれば False
    (=実測が混ざっているので件数として認める)。"""
    if isinstance(node, ast.Constant):
        return isinstance(node.value, (int, float)) \
            and not isinstance(node.value, bool)
    if isinstance(node, ast.Name):
        return node.id in known
    if isinstance(node, (ast.BinOp, ast.UnaryOp)):
        return all(_is_const_expr(ch, known)
                   for ch in ast.iter_child_nodes(node)
                   if isinstance(ch, ast.expr))
    return False


def _const_names(tree: "ast.AST") -> frozenset:
    """「数値定数しか代入されない変数」の名前を集める。

    同じ名前に1つでも実測(`len(rows)` 等)が代入されていれば定数ではない
    (=拾わない)。増分 `n += 1` も実測なので定数から外す。名前はモジュール
    全体で1つの集合にする(関数をまたいで同名を定数と実測に使い分ける書き方は
    そもそも読めないので、fail-closed 側に倒さず**実測が勝つ**)。
    """
    assigned: dict = {}
    for node in ast.walk(tree):
        targets: list = []
        value = None
        if isinstance(node, ast.Assign):
            targets, value = node.targets, node.value
        elif isinstance(node, ast.AnnAssign) and node.value is not None:
            targets, value = [node.target], node.value
        elif isinstance(node, ast.AugAssign):
            targets, value = [node.target], None
        elif isinstance(node, (ast.For, ast.AsyncFor)):
            targets, value = [node.target], None
        else:
            continue
        for tgt in targets:
            if not isinstance(tgt, ast.Name):
                continue
            assigned.setdefault(tgt.id, []).append(value)
    # 定数変数どうしの参照(`a = 2` / `b = a + 1`)を閉包で解く。
    known: set = set()
    for _round in range(8):
        grew = False
        for name, values in assigned.items():
            if name in known:
                continue
            if values and all(v is not None and _is_const_expr(v, known)
                              for v in values):
                known.add(name)
                grew = True
        if not grew:
            break
    return frozenset(known)


def literal_count_sites(src: str) -> list[tuple[int, str]]:
    """`record(名前, 件数)` の**件数の式に数値リテラルが現れる**箇所を返す。

    戻り値は (行番号, 件数の式の姿) のリスト。`4` も `n_dom + 14` も
    「その場で数えていない」ので同じように拾う。名前(第1引数)は見ない。
    """
    try:
        tree = ast.parse(src)
    except SyntaxError:
        return [(0, "<構文エラーで読めません>")]
    const = _const_names(tree)
    fixture_lines: set = set()
    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) \
                and node.name in FIXTURE_FUNCS:
            fixture_lines.update(range(node.lineno,
                                       (node.end_lineno or node.lineno) + 1))
    out: list[tuple[int, str]] = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        fn = node.func
        name = fn.attr if isinstance(fn, ast.Attribute) else (
            fn.id if isinstance(fn, ast.Name) else "")
        if name != "record" or len(node.args) < 2:
            continue
        if node.lineno in fixture_lines:
            continue
        count = node.args[1]
        if _literal_in_count(count, const):
            try:
                shown = ast.unparse(count)
            except Exception:       # pragma: no cover (古い Python)
                shown = "<式>"
            out.append((getattr(count, "lineno", 0), shown))
    return out


def audit_literal_counts(tools: set, verbose: bool) -> tuple[int, int]:
    """件数がリテラル定数でないことを見る。戻り値: (違反件数, 見た本数)。"""
    violations = 0
    seen = 0
    for script in sorted(tools):
        path = TOOLS_DIR / script
        if not path.exists():
            continue
        seen += 1
        sites = literal_count_sites(path.read_text(encoding="utf-8",
                                                   errors="ignore"))
        if script in LITERAL_COUNT_PENDING:
            if not sites:
                print("ERROR %s は LITERAL_COUNT_PENDING に登記されていますが、"
                      "件数にリテラル定数はもうありません(登記を消してください)"
                      % script)
                violations += 1
            elif verbose:
                print("  PENDING %s (%s)" % (script, LITERAL_COUNT_PENDING[script]))
            continue
        for lineno, shown in sites:
            print("ERROR %s:%d 要点行の件数がリテラル定数です: record(…, %s)"
                  "(実際に回した回数をカウンタで数えてください。定数だと検査を"
                  "抜いても件数が変わらず『確認しました』と名乗れます)"
                  % (script, lineno, shown))
            violations += 1
        if verbose and not sites:
            print("  OK      %s (件数は実測)" % script)
    return (violations, seen)


def tool_follows_contract(path: Path) -> bool:
    """そのツールが本契約(数える・名乗る・落ちる)を実際に使っているか。"""
    if not path.exists():
        return False
    src = path.read_text(encoding="utf-8", errors="ignore")
    return ("gate_count" in src and "Checked(" in src
            and "gate_count.report(" in src)


def audit(verbose: bool) -> tuple[int, Checked]:
    """戻り値: (違反件数, 数え上げ)"""
    checked = Checked()
    gate_py = TOOLS_DIR / "gate.py"
    if not gate_py.exists():
        print("ERROR tools/gate.py がありません")
        return (1, checked)
    entries = gate_scripts(gate_py.read_text(encoding="utf-8"))
    scripts = sorted({s for _n, s in entries})
    violations = 0

    adopted = 0
    for script in scripts:
        if script in CONTRACT_TOOLS:
            if tool_follows_contract(TOOLS_DIR / script):
                adopted += 1
                if verbose:
                    print("  OK      %s (契約適用)" % script)
            else:
                print("ERROR %s は契約適用のはずですが、要点行の仕組み"
                      "(gate_count.Checked / report)を使っていません" % script)
                violations += 1
        elif script in PENDING_TOOLS:
            if verbose:
                print("  PENDING %s (%s)" % (script, PENDING_TOOLS[script]))
        else:
            print("ERROR %s が GATES にありますが、契約にも未適用登記にも"
                  "ありません(検査件数を名乗らない検問を黙って増やさない)"
                  % script)
            violations += 1
    checked.record("ゲート登録", len(entries))
    checked.record("実行ファイル", len(scripts))
    checked.record("契約適用", adopted)

    # 掃除漏れ(GATES から消えたのに登記だけ残っている)。
    stale = sorted((set(PENDING_TOOLS) | CONTRACT_TOOLS) - set(scripts)
                   - {"gate_count.py"})
    for script in stale:
        print("ERROR %s は GATES にありませんが登記だけ残っています"
              "(登記を消してください)" % script)
        violations += 1
    checked.record("登記の掃除", len(PENDING_TOOLS) + len(CONTRACT_TOOLS))

    # gate.py の登録行そのものの衛生: 同じゲート名の二重登録。
    dup = sorted({n for n, _ in entries if [x for x, _ in entries].count(n) > 1})
    for n in dup:
        print("ERROR gate.py の GATES にゲート名 %r が二重に登録されています"
              "(--only で二度走り、赤の数も二重に数えられます)" % n)
        violations += 1
    checked.record("ゲート名の重複", len(entries))

    # 件数がリテラル定数でないこと(裁定書43 §2 Y-7)。契約適用ツールと、
    # 契約の側に立つ本ファイル自身を見る(自分だけ例外にしない)。
    lit_targets = set(CONTRACT_TOOLS) | set(LITERAL_COUNT_PENDING) | {"gate_count.py"}
    lit_violations, lit_seen = audit_literal_counts(lit_targets, verbose)
    violations += lit_violations
    checked.record("件数の実測", lit_seen)

    # 登記の掃除(GATES にも契約にも無いのに登記だけ残っている)。
    for script in sorted(set(LITERAL_COUNT_PENDING) - lit_targets):
        print("ERROR %s は LITERAL_COUNT_PENDING の登記だけが残っています"
              % script)
        violations += 1
    return (violations, checked)


# ---------------------------------------------------------------------------
# 自己テスト(骨抜き防止)
# ---------------------------------------------------------------------------
def self_test() -> bool:
    cases = []

    c = Checked()
    c.record("あ", 3)
    c.record("い", 2)
    cases.append(("合計を数える", c.total() == 5))
    cases.append(("要点行の書式", c.line() == "検査実施: 5件 (あ=3, い=2)"))
    cases.append(("要点行は CHECKED_LINE で読める",
                  CHECKED_LINE.match(c.line()) is not None
                  and CHECKED_LINE.match(c.line()).group(1) == "5"))
    cases.append(("回していない検査は名前ごと出ない", "う" not in c.line()))
    cases.append(("gate.py 推奨パターンは非0件だけに当たる",
                  re.search(GATE_PATTERN, c.line()) is not None))

    empty = Checked()
    cases.append(("0件の要点行は推奨パターンに当たらない",
                  re.search(GATE_PATTERN, empty.line()) is None))

    # 「落ちる」方向。report は標準出力を出すので握りつぶして戻り値だけ見る。
    import contextlib
    import io as _io

    def rc(counter, required=()):
        buf = _io.StringIO()
        with contextlib.redirect_stdout(buf):
            return report(counter, required)

    cases.append(("検査0件は赤", rc(Checked()) == 1))
    zero = Checked()
    zero.record("あ", 0)
    cases.append(("0件だけを積んでも赤", rc(zero) == 1))
    cases.append(("必須の検査を回していなければ赤",
                  rc(c, required=("う",)) == 1))
    cases.append(("回していれば緑", rc(c, required=("あ", "い")) == 0))
    cases.append(("必須が0件なら赤", rc(zero, required=("あ",)) == 1))

    # 件数がリテラル定数か(裁定書43 §2 Y-7)。両方向を固定する。
    cases.append(("Y-7 定数の件数を拾う",
                  [s for _l, s in literal_count_sites(
                      'c.record("a", 4)\n')] == ["4"]))
    cases.append(("Y-7 定数を足した式も拾う",
                  [s for _l, s in literal_count_sites(
                      'c.record("a", n + 14)\n')] == ["n + 14"]))
    cases.append(("Y-7 実測の件数は拾わない",
                  literal_count_sites('c.record("a", len(rows))\n') == []))
    cases.append(("Y-7 カウンタの件数は拾わない",
                  literal_count_sites('c.record("a", n_hit)\n') == []))
    cases.append(("Y-7 名前(第1引数)の中は見ない",
                  literal_count_sites('c.record("SEC-09", n)\n') == []))
    cases.append(("Y-7 record 以外は見ない",
                  literal_count_sites('c.append("a", 4)\n') == []))
    cases.append(("Y-7 読めない Python は fail-closed",
                  literal_count_sites("def (:\n") != []))
    cases.append(("Y-7 添字の数字は件数ではない",
                  literal_count_sites('c.record("a", _N[0])\n') == []))
    cases.append(("Y-7 添字でも足した定数は拾う",
                  [s for _l, s in literal_count_sites(
                      'c.record("a", _N[0] + 1)\n')] == ["_N[0] + 1"]))
    # Y-7(司令塔手直し): 変数を1つ挟んだ抜け道も塞ぐ。
    cases.append(("Y-7 変数経由の定数も拾う",
                  literal_count_sites(
                      'def a():\n    n = 14\n    c.record("x", n)\n') != []))
    cases.append(("Y-7 定数どうしの計算も拾う",
                  literal_count_sites(
                      'def a():\n    n = 2\n    m = n + 1\n'
                      '    c.record("x", m)\n') != []))
    cases.append(("Y-7 実測を代入した変数は拾わない",
                  literal_count_sites(
                      'def a():\n    n = len(rows)\n'
                      '    c.record("x", n)\n') == []))
    cases.append(("Y-7 数え上げた変数は拾わない",
                  literal_count_sites(
                      'def a():\n    n = 0\n    n += 1\n'
                      '    c.record("x", n)\n') == []))
    cases.append(("Y-7 実測が1つでもあれば定数ではない",
                  literal_count_sites(
                      'def a():\n    n = 3\n    n = len(rows)\n'
                      '    c.record("x", n)\n') == []))
    cases.append(("Y-7 自己テストの器は見ない",
                  literal_count_sites(
                      'def self_test():\n    c.record("a", 3)\n') == []))
    cases.append(("Y-7 本物の検査は自己テストの外なので見る",
                  literal_count_sites(
                      'def audit():\n    c.record("a", 3)\n') != []))

    # GATES の読み取り(構文木)。
    src = (
        "import sys\n"
        "GATES = [\n"
        '    ("zza", [sys.executable, "tools/zz_a.py"], r"x"),\n'
        '    ("zzb", [sys.executable, "tools/zz_b.py", "--mode", "x"], r"y"),\n'
        "]\n"
    )
    got = gate_scripts(src)
    cases.append(("GATES を構文木から読む",
                  got == [("zza", "zz_a.py"), ("zzb", "zz_b.py")]))
    cases.append(("GATES が無ければ空", gate_scripts("X = 1\n") == []))

    bad = [name for name, ok in cases if not ok]
    for name in bad:
        print("  自己テスト NG: %s" % name)
    print("  自己テスト: %d/%d" % (len(cases) - len(bad), len(cases)))
    _SELFTEST_N[0] = len(cases)
    return not bad


# 自己テストの本数(要点行へ出す実測値。定数を書かない = Y-7 を自分にも適用)。
_SELFTEST_N = [0]


def self_test_count() -> int:
    _SELFTEST_N[0] = 0
    self_test()
    return _SELFTEST_N[0]


def main() -> int:
    ap = argparse.ArgumentParser(
        description="検問の要点行の契約(検査件数)と、その登記の網羅を見る")
    ap.add_argument("--selftest", action="store_true", help="自己テストだけ")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args()

    print("gate_count: 要点行の契約(W15 §3 X3-3)")
    if args.selftest:
        if not self_test():
            print("結果: 自己テスト失敗(検出器が壊れています)")
            return 2
        print("結果: 自己テストOK")
        return 0

    violations, checked = audit(args.verbose)
    if not self_test():
        print("結果: 自己テスト失敗(検出器が壊れています)")
        return 2
    checked.record("自己テスト", _SELFTEST_N[0])

    rc = report(checked, required=("ゲート登録", "契約適用"))
    if violations:
        print("結果: NG (契約違反 %d件)" % violations)
        return 1
    if rc:
        return 1
    print("結果: OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
