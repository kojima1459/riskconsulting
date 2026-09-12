#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
prompt_diff.py - 15章のプロンプト/スキーマ本文と .bas 実装の一致検査(17章 T-01・T-23)

================================================================================
役割:
    docs/spec/15_プロンプトとJSONスキーマ.md から「プロンプト本文」を抽出し、
    src/app/modPrompts*.bas / modSchemas*.bas の対応関数が返す文字列と突き合わせる。
    突合キーは 15章§10.2 の「節⇔関数名対応表」。抽出規約は 15章§10.1。

    T-23 の受入条件は **0 diff**。W0(T-01)時点ではプロンプト系モジュールが
    1本も実装されていないため、「対象0件・スキップ」を明示して正常終了する。

15章§10.1 抽出規約(実装した内容):
    (a) 本文はコードフェンス内のみ。フェンスの開始行・終了行そのものは含めない。
        1つの節に複数のフェンスがある場合、**本文は最初のフェンスのみ**。
        2本目以降は冒頭に「例:」を冠した例示であり突合対象外。
        1つの節に複数の関数が対応する場合(§1.2の3ブロック・§5の構成指示2本)は
        §10.2の対応表がフェンスと関数を1対1に割り当てる(下の SECTION_MAP)。
    (b) フェンス内の行頭「※」の行は本文である。除去しない。
    (c) プレースホルダは `{{識別子}}` の形だけを残す。`{{識別子 ※...}}` の
        ※以降は実装向け注記なので除去する。第3形(日本語ラベル形)
        `{{識別子の日本語: ラベル列挙}}` はコロン以降を除去して
        `{{識別子の日本語}}` へ正規化する。
    (d) 行がまるごと条件付きで消えるプレースホルダの行
        (`{{BLOCK_RENEWAL_S1 ※renewalのみ}}` など3行)は、15章側・コード側の
        双方から除外して比較する。

一致検査の正規化規則(§10.1末尾):
    改行は LF に統一し末尾改行は付けない。VBAソース上の `""` は `"` に戻して
    比較する。各行の行末の半角空白は両側で除去する。全角文字はそのまま比較する
    (CP932外文字は tools/vba_lint.py の check_cp932_safe が別途FAILさせる)。

使い方:
    python3 tools/prompt_diff.py
    python3 tools/prompt_diff.py --strict   # 未実装関数も差分として数える
    exit code: 差分件数(0=一致)。125件を超えたら125で頭打ちにする

スキーマのJSON検査(17章 T-23 DoD「schemaはJSONとしてパース可能」の機械化):
    Schema* 系の抽出本文を json.loads でパースし、さらに 14章§3 が direct 経路の
    必須要件と定める **strict要件**(すべてのオブジェクトで properties と required が
    一致し、additionalProperties: false が付いている)を再帰的に検査する。
    違反は差分と同じく exit code に算入する(--strict の合格条件に含まれる)。
    これが無いと、スキーマ本体が壊れても「15章と一字一句一致」だけは通ってしまい、
    実機のdirect経路で初めて落ちる。

    --strict は「未実装=差分」とみなす最終確認用スイッチ。突合先ディレクトリが
    無い/対象 .bas が0本(=全関数未実装)のときも、--strict では不合格(exit 1)に
    倒す。**--strict を付けない既定はこの状態を「T-23で実装予定のスキップ」として
    exit 0 のまま通す**(W0～W1では実装前なのが正常なため)。この非対称が無いと
    「対象0件」を無条件に合格扱いにしてしまい --strict が骨抜きになる(裁定書4 項目13)。
================================================================================
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

TOOLS_DIR = Path(__file__).resolve().parent
REPO_ROOT = TOOLS_DIR.parent
DEFAULT_SPEC = REPO_ROOT / "docs" / "spec" / "15_プロンプトとJSONスキーマ.md"
DEFAULT_SRC = REPO_ROOT / "src" / "app"
# 突合対象のファイル名パターン(15章§10.2 の実装モジュール列。分割時は末尾に数字)。
SRC_GLOBS = ("modPrompts*.bas", "modSchemas*.bas")

# ==============================================================================
# 15章§10.2 節⇔関数名対応表(突合キー)
# ------------------------------------------------------------------------------
# (見出し行に含まれる目印, [その節のフェンスに順番に対応する関数名...], 実装モジュール)
# 見出し行そのものに関数名が書かれている節はそれを目印にし、ブロック定数名で
# 書かれている節(§1.1/§1.2/§1.3/§5の構成指示)は定数名を目印にする。
# §9(WT/FG)は「まだコードフェンスを持たず要旨のみ」のため Phase 1.5 まで対象外。
# ==============================================================================
SECTION_MAP: list[tuple[str, list[str], str]] = [
    ("BLOCK_CTX", ["BlockCtx"], "modPromptsBlocks"),
    ("BLOCK_RENEWAL_S1 / S2 / S3",
     ["BlockRenewalS1", "BlockRenewalS2", "BlockRenewalS3"], "modPromptsBlocks"),
    ("BLOCK_NEW_S2", ["BlockNewS2"], "modPromptsBlocks"),
    ("BLOCK_ROUND2_FOCUS", ["BlockRound2Focus"], "modPromptsBlocks"),
    ("BLOCK_GUARD", ["BlockGuard"], "modPromptsBlocks"),
    ("BuildS1System", ["BuildS1System"], "modPromptsCore"),
    ("BuildS1User", ["BuildS1User"], "modPromptsCore"),
    ("SchemaS1()", ["SchemaS1"], "modSchemas"),
    ("BuildS2System", ["BuildS2System"], "modPromptsCore"),
    ("BuildS2User", ["BuildS2User"], "modPromptsCore"),
    ("SchemaS2()", ["SchemaS2"], "modSchemas"),
    ("BuildS3System", ["BuildS3System"], "modPromptsCore"),
    ("BuildS3User", ["BuildS3User"], "modPromptsCore"),
    ("SchemaS3()", ["SchemaS3"], "modSchemas"),
    ("BuildS2CriticSystem", ["BuildS2CriticSystem"], "modPromptsOps"),
    ("BuildS2CriticUser", ["BuildS2CriticUser"], "modPromptsOps"),
    ("SchemaS2C()", ["SchemaS2C"], "modSchemas"),
    ("BuildS3CriticSystem", ["BuildS3CriticSystem"], "modPromptsOps"),
    ("BuildS3CriticUser", ["BuildS3CriticUser"], "modPromptsOps"),
    ("SchemaS3C()", ["SchemaS3C"], "modSchemas"),
    ("ReviseSuffix", ["ReviseSuffix"], "modPromptsOps"),
    ("BLOCK_S4_PROPOSAL", ["BlockS4Proposal"], "modPromptsBlocks"),
    ("BLOCK_S4_ALLIANCE", ["BlockS4Alliance"], "modPromptsBlocks"),
    ("BuildS4System", ["BuildS4System"], "modPromptsCore"),
    ("BuildS4User", ["BuildS4User"], "modPromptsCore"),
    ("SchemaS4()", ["SchemaS4"], "modSchemas"),
    ("BuildPFSystem", ["BuildPFSystem"], "modPromptsOps"),
    ("BuildPFUser", ["BuildPFUser"], "modPromptsOps"),
    ("SchemaPF()", ["SchemaPF"], "modSchemas"),
    ("BuildSparringSystem", ["BuildSparringSystem"], "modPromptsOps"),
    ("RepairSuffix", ["RepairSuffix"], "modPromptsOps"),
    # W15・裁定書38 班C: 顧客向け提案書(S5)。15章§5.6・20章。
    ("BuildS5System", ["BuildS5System"], "modPromptsS5"),
    ("BuildS5User", ["BuildS5User"], "modPromptsS5"),
    ("SchemaS5()", ["SchemaS5"], "modSchemas2"),
]

# ==============================================================================
# 配線ゲート(裁定書37 B-01/B-02。B_dr_quality.md §3 B-02/§4)
# ------------------------------------------------------------------------------
# 15章が「systemの末尾へ差し込む」と定める Block* 9関数が、src/ 全体で
# 1件以上「呼ばれている」ことを検査する。SECTION_MAP の文言一致ゲートとは別物
# (文言が15章と一字一句一致していても、どこからも呼ばれていなければ0件のまま
# 本番に出る=伝書鳩3-1で実際に起きたBlockGuard未配線)。
# fail-closed: src/ に .bas が読めない・1本も無いときも赤にする。
# ==============================================================================
WIRED_BLOCKS: dict[str, str] = {
    "BlockCtx": "§1.1", "BlockRenewalS1": "§1.2", "BlockRenewalS2": "§1.2",
    "BlockRenewalS3": "§1.2", "BlockNewS2": "§1.2b", "BlockRound2Focus": "§1.2c",
    "BlockGuard": "§1.3", "BlockS4Proposal": "§5", "BlockS4Alliance": "§5",
}
WIRING_ROOT = REPO_ROOT / "src"
_WIRE_DEF_RE = re.compile(r"^\s*(Public|Private)\s+Function\s+(\w+)", re.IGNORECASE)


def block_wiring_counts() -> tuple[dict[str, int], bool]:
    """{関数名: 呼び出し件数} と、src/ が読めたか(.basが1本以上あったか)を返す。
    定義行(Public|Private Function Xxx)自身と、`'` で始まるコメント行は
    呼び出しに数えない。"""
    counts = {fn: 0 for fn in WIRED_BLOCKS}
    try:
        # src/test はテストからの呼び出しであり「製品の配線」ではないので数えない
        # (テストだけが呼んでいる状態を緑にしないため。司令塔検収 W14)。
        bas_files = sorted(p for p in WIRING_ROOT.rglob("*.bas")
                           if "test" not in p.relative_to(WIRING_ROOT).parts)
    except OSError:
        return counts, False
    if not bas_files:
        return counts, False

    call_patterns = {fn: re.compile(r"\b" + re.escape(fn) + r"\s*\(")
                      for fn in WIRED_BLOCKS}
    read_ok = False
    for path in bas_files:
        try:
            lines = path.read_text(encoding="utf-8").splitlines()
        except OSError:
            continue
        read_ok = True
        for line in lines:
            if line.strip().startswith("'"):
                continue
            m = _WIRE_DEF_RE.match(line)
            def_name = m.group(2) if m else None
            for fn, pat in call_patterns.items():
                if def_name == fn:
                    continue
                if pat.search(line):
                    counts[fn] += 1
    return counts, read_ok


_CALLSITE_RE = re.compile(r"\b(Build\w+System)\s*\(")
# 15章§1.3 の例外=壁打ち。案件チャット(BuildChatSystem)は15章§1.3 の7本に
# 含まれない(ガード要否は Z-52 で裁定)ため、ここでは対象外として明示する。
_CALLSITE_EXEMPT = {"BuildSparringSystem"}
# BuildChatSystem は裁定書38 Z-52 で AsmGuarded 対象に加わった(案件チャットも
# 貼付資料を注入するため)。呼出点(modNaviChat.Ask)は AsmGuarded( で包んで
# いるので EXEMPT から外した(除外していたのは裁定書37時点で対象外だったため)。


def check_guard_callsites() -> list[str]:
    """司令塔検収(W14 抜き打ち): 配線ゲートは「1件以上呼ばれている」しか見ない
    ので、9代入点のうち1つから AsmGuarded を外しても緑のままだった(実演済み)。
    そこで system 組立関数 Build*System を呼ぶ行は、modPromptsOps 自身
    (AsmS4System の内側で包む)と _CALLSITE_EXEMPT を除き、同じ行で AsmGuarded(
    に包まれていることを要求する。定義行・コメント行・src/test は対象外。
    S4 は modPromptsOps.AsmS4System の戻り値代入行が AsmGuarded( を含むことを見る。"""
    problems: list[str] = []
    s4_guarded = False
    for path in sorted(WIRING_ROOT.rglob("*.bas")):
        rel = path.relative_to(WIRING_ROOT)
        if "test" in rel.parts:
            continue
        lines = path.read_text(encoding="utf-8").splitlines()
        if path.name == "modPromptsOps.bas":
            s4_guarded = any(
                l.lstrip().startswith("AsmS4System =") and "AsmGuarded(" in l
                for l in lines)
            continue
        for no, line in enumerate(lines, 1):
            if line.strip().startswith("'") or _WIRE_DEF_RE.match(line):
                continue
            m = _CALLSITE_RE.search(line)
            if not m or m.group(1) in _CALLSITE_EXEMPT:
                continue
            if "AsmGuarded(" not in line:
                problems.append(
                    f"[代入点] {rel}:{no} の {m.group(1)}() が AsmGuarded( に包まれていません"
                    f"(裁定書37 B-01: 9代入点は全て AsmGuarded を通す)")
    if not s4_guarded:
        problems.append("[代入点] modPromptsOps.AsmS4System の戻り値が AsmGuarded( で包まれていません")
    return problems


def check_block_wiring() -> list[str]:
    """WIRED_BLOCKS の各関数が1件以上呼ばれているかを検査する。
    戻り値は問題メッセージの一覧(空=全件配線済み)。fail-closed。"""
    counts, ok = block_wiring_counts()
    if not ok:
        return ["[配線] src/ に .bas が1本も読めません(fail-closed)"]
    problems: list[str] = []
    for fn in sorted(WIRED_BLOCKS):
        if counts[fn] == 0:
            problems.append(
                f"[配線] {fn}() が src/ で1件も呼ばれていません"
                f"(15章{WIRED_BLOCKS[fn]}が「差し込む」と定める)")
    return problems


HEADING_RE = re.compile(r"^#{2,4}\s+(.*)$")
FENCE_RE = re.compile(r"^```")
EXAMPLE_LEAD_RE = re.compile(r"^\s*例[:：]")

# (c) プレースホルダの注記除去
PLACEHOLDER_NOTE_RE = re.compile(r"\{\{([^{}※]*?)\s*※[^{}]*?\}\}")
PLACEHOLDER_JP_LABEL_RE = re.compile(r"\{\{([^{}:：]*?の日本語)\s*[:：][^{}]*?\}\}")
# (d) 条件付きで丸ごと消える行(正規化後の形で判定する)
CONDITIONAL_LINE_RE = re.compile(r"^\s*\{\{BLOCK_RENEWAL_S[123]\}\}\s*$")


def normalize_body(text: str) -> str:
    """§10.1 (b)(c)(d) と正規化規則を適用して比較用の本文にする。"""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    out = []
    for line in text.split("\n"):
        line = PLACEHOLDER_NOTE_RE.sub(r"{{\1}}", line)
        line = PLACEHOLDER_JP_LABEL_RE.sub(r"{{\1}}", line)
        if CONDITIONAL_LINE_RE.match(line):
            continue          # (d) 両側から除外する
        out.append(line.rstrip(" \t"))
    while out and out[-1] == "":
        out.pop()
    return "\n".join(out)


# ==============================================================================
# 15章側の抽出
# ==============================================================================
def extract_spec_bodies(spec_path: Path) -> tuple[dict[str, str], list[str]]:
    """{関数名: 本文} と、対応表に載っているのに節が見つからなかった関数名の一覧。"""
    lines = spec_path.read_text(encoding="utf-8").split("\n")

    # 見出しごとに (見出し行, [フェンス本文...]) を集める。
    sections: list[tuple[str, list[str], list[str]]] = []
    cur_head = ""
    cur_fences: list[str] = []
    cur_leads: list[str] = []
    buf: list[str] = []
    in_fence = False
    lead = ""

    def flush_section():
        if cur_head:
            sections.append((cur_head, list(cur_fences), list(cur_leads)))

    for i, line in enumerate(lines):
        if not in_fence:
            m = HEADING_RE.match(line)
            if m:
                flush_section()
                cur_head = line.strip()
                cur_fences = []
                cur_leads = []
                continue
        if FENCE_RE.match(line):
            if not in_fence:
                in_fence = True
                buf = []
                lead = ""
                for j in range(i - 1, -1, -1):
                    if lines[j].strip():
                        lead = lines[j].strip()
                        break
            else:
                in_fence = False
                cur_fences.append("\n".join(buf))
                cur_leads.append(lead)
            continue
        if in_fence:
            buf.append(line)
    flush_section()

    bodies: dict[str, str] = {}
    not_found: list[str] = []
    for marker, funcs, _module in SECTION_MAP:
        hit = None
        for head, fences, leads in sections:
            if marker in head:
                hit = (head, fences, leads)
                break
        if hit is None:
            not_found.extend(funcs)
            continue
        _head, fences, leads = hit
        # (a) 「例:」を冠したフェンスは例示なので除外する。
        real = [f for f, ld in zip(fences, leads) if not EXAMPLE_LEAD_RE.match(ld)]
        if len(real) < len(funcs):
            not_found.extend(funcs[len(real):])
        for idx, fn in enumerate(funcs):
            if idx < len(real):
                bodies[fn] = normalize_body(real[idx])
    return bodies, not_found


# ==============================================================================
# .bas 側の評価
# ==============================================================================
CONT_RE = re.compile(r"\s_$")
FUNC_HEAD_RE = re.compile(
    r"^\s*Public\s+Function\s+([A-Za-z_]\w*)\s*\(", re.IGNORECASE)
FUNC_END_RE = re.compile(r"^\s*End\s+Function\b", re.IGNORECASE)

VBA_CONSTS = {
    "vblf": "\n",
    "vbcrlf": "\r\n",
    "vbcr": "\r",
    "vbtab": "\t",
    "vbnullstring": "",
}


class EvalError(Exception):
    pass


def _merge_continuations(raw_lines: list[str]) -> list[str]:
    out: list[str] = []
    acc = ""
    for raw in raw_lines:
        line = raw.rstrip("\n\r")
        if line.lstrip().startswith("'"):
            if acc:
                out.append(acc)
                acc = ""
            out.append(line)
            continue
        rstripped = line.rstrip()
        if rstripped.endswith(" _") or rstripped == "_":
            acc = (acc + " " if acc else "") + rstripped[:-1].rstrip()
            continue
        acc = (acc + " " if acc else "") + line
        out.append(acc)
        acc = ""
    if acc:
        out.append(acc)
    return out


def _strip_comment(line: str) -> str:
    in_str = False
    for i, c in enumerate(line):
        if c == '"':
            in_str = not in_str
        elif c == "'" and not in_str:
            return line[:i]
    return line


def _split_top_level_amp(expr: str) -> list[str]:
    parts, cur, depth, in_str = [], [], 0, False
    for c in expr:
        if c == '"':
            in_str = not in_str
            cur.append(c)
        elif not in_str and c == "(":
            depth += 1
            cur.append(c)
        elif not in_str and c == ")":
            depth = max(0, depth - 1)
            cur.append(c)
        elif not in_str and depth == 0 and c == "&":
            parts.append("".join(cur))
            cur = []
        else:
            cur.append(c)
    parts.append("".join(cur))
    return parts


def _assign_index(stmt: str) -> int:
    depth, in_str = 0, False
    for i, c in enumerate(stmt):
        if c == '"':
            in_str = not in_str
        elif not in_str:
            if c == "(":
                depth += 1
            elif c == ")":
                depth = max(0, depth - 1)
            elif c == "=" and depth == 0:
                return i
    return -1


def _eval_term(term: str, env: dict[str, str]) -> str:
    t = term.strip()
    if not t:
        raise EvalError("空の項があります")
    if t.startswith('"') and t.endswith('"') and len(t) >= 2:
        inner = t[1:-1]
        if '"' in inner.replace('""', ""):
            raise EvalError(f"文字列リテラルを解釈できません: {t[:40]}")
        return inner.replace('""', '"')      # VBAの "" は " に戻す
    low = t.lower()
    if low in VBA_CONSTS:
        return VBA_CONSTS[low]
    if low in env:
        return env[low]
    raise EvalError(
        f"未対応の項 {t[:40]!r}(本関数は文字列リテラル・vbLf等の組込定数・"
        f"同一関数内で組み立てた変数の連結だけを評価します)")


def eval_function_body(lines: list[str], func_name: str) -> str:
    """`s = s & "..." & vbLf` 方式で組み立てた戻り値を評価して返す。"""
    env: dict[str, str] = {}
    for raw in lines:
        stmt = _strip_comment(raw).strip()
        if not stmt:
            continue
        low = stmt.lower()
        if low.startswith(("dim ", "static ", "const ", "option ", "exit ",
                           "on error", "if ", "elseif ", "else", "end if",
                           "for ", "next", "do ", "loop", "select ", "case ",
                           "with ", "end with", "while ", "wend")):
            if low.startswith(("if ", "elseif ", "for ", "do ", "select ",
                               "while ", "with ")):
                raise EvalError(f"制御構文は評価できません: {stmt[:60]}")
            continue
        eq = _assign_index(stmt)
        if eq < 0:
            raise EvalError(f"代入以外の文は評価できません: {stmt[:60]}")
        lhs = stmt[:eq].strip()
        if not re.fullmatch(r"[A-Za-z_]\w*", lhs):
            raise EvalError(f"単純変数への代入以外は評価できません: {stmt[:60]}")
        value = "".join(_eval_term(p, env) for p in _split_top_level_amp(stmt[eq + 1:]))
        env[lhs.lower()] = value
    if func_name.lower() not in env:
        raise EvalError(f"戻り値 {func_name} への代入が見つかりません")
    return env[func_name.lower()]


def extract_code_bodies(src_dir: Path):
    """{関数名: (ファイル, 本文 or None, エラー)} を返す。"""
    files: list[Path] = []
    for pat in SRC_GLOBS:
        files.extend(sorted(src_dir.glob(pat)))

    result: dict[str, tuple[Path, str | None, str]] = {}
    for path in files:
        raw_lines = path.read_text(encoding="utf-8").split("\n")
        merged = _merge_continuations(raw_lines)
        cur_name = None
        body: list[str] = []
        for line in merged:
            if cur_name is None:
                m = FUNC_HEAD_RE.match(_strip_comment(line))
                if m:
                    cur_name = m.group(1)
                    body = []
                continue
            if FUNC_END_RE.match(_strip_comment(line)):
                try:
                    result[cur_name] = (path, normalize_body(
                        eval_function_body(body, cur_name)), "")
                except EvalError as e:
                    result[cur_name] = (path, None, str(e))
                cur_name = None
                continue
            body.append(line)
    return files, result


# ==============================================================================
# Schema* のJSON検査(17章 T-23 DoD の機械化。14章§3 の strict 要件)
# ==============================================================================
SCHEMA_FUNCS = ("SchemaS1", "SchemaS2", "SchemaS3", "SchemaS4", "SchemaS5",
                "SchemaS2C", "SchemaS3C", "SchemaPF", "SchemaWT", "SchemaFG")


def _walk_schema(node, path: str, problems: list[str]) -> None:
    if isinstance(node, dict):
        if node.get("type") == "object" or "properties" in node:
            props = node.get("properties")
            if not isinstance(props, dict):
                problems.append(f"{path}: type=object なのに properties がありません")
            else:
                req = node.get("required")
                if not isinstance(req, list):
                    problems.append(
                        f"{path}: required がありません"
                        "(strict:true は全プロパティを required に要求します)")
                else:
                    missing = [k for k in props if k not in req]
                    extra = [k for k in req if k not in props]
                    if missing:
                        problems.append(f"{path}: required に無い properties: {missing}")
                    if extra:
                        problems.append(f"{path}: properties に無い required: {extra}")
            if node.get("additionalProperties") is not False:
                problems.append(f"{path}: additionalProperties: false がありません")
        for k, v in node.items():
            _walk_schema(v, f"{path}.{k}", problems)
    elif isinstance(node, list):
        for i, v in enumerate(node):
            _walk_schema(v, f"{path}[{i}]", problems)


def check_schema_body(name: str, body: str) -> list[str]:
    """1本のスキーマ本文を検査して違反メッセージの一覧を返す(空=合格)。"""
    try:
        obj = json.loads(body)
    except ValueError as e:
        return [f"JSONとしてパースできません: {e}"]
    problems: list[str] = []
    _walk_schema(obj, name, problems)
    return problems


def run_schema_checks(code: dict) -> int:
    """実装済みの Schema* を全部検査し、違反件数を返す。"""
    print("-" * 78)
    print("スキーマJSON検査 (json.loads + strict要件: properties=required / "
          "additionalProperties:false)")
    bad = 0
    checked = 0
    for fn in SCHEMA_FUNCS:
        entry = code.get(fn)
        if entry is None:
            continue
        path, body, _err = entry
        if body is None:
            continue
        checked += 1
        problems = check_schema_body(fn, body)
        if problems:
            bad += len(problems)
            print(f"  NG    {fn} [{path.name}]: {len(problems)}件")
            for msg in problems[:6]:
                print(f"        - {msg}")
        else:
            print(f"  OK    {fn} [{path.name}] ({len(body)}字)")
    if checked == 0:
        print("  (実装済みの Schema* がありません)")
    return bad


def show_diff(name: str, want: str, got: str) -> None:
    wl, gl = want.split("\n"), got.split("\n")
    for i in range(max(len(wl), len(gl))):
        w = wl[i] if i < len(wl) else "(行なし)"
        g = gl[i] if i < len(gl) else "(行なし)"
        if w != g:
            print(f"      15章 L{i + 1}: {w!r}")
            print(f"      コード L{i + 1}: {g!r}")
            return
    print(f"      (行の内容は一致。長さ: 15章={len(want)}字 コード={len(got)}字)")


def main() -> int:
    ap = argparse.ArgumentParser(
        description="15章のプロンプト/スキーマ本文と .bas 実装の一致検査")
    ap.add_argument("--spec", default=str(DEFAULT_SPEC), help="15章のパス")
    ap.add_argument("--src", default=str(DEFAULT_SRC),
                    help="突合対象の .bas を置いたディレクトリ(既定: src/app)")
    ap.add_argument("--strict", action="store_true",
                    help="未実装関数も差分として数える(T-23の最終確認用)。"
                         "突合先ディレクトリ不在・対象0件のときも不合格(exit 1)にする"
                         "(--strict なしはこれらを exit 0 でスキップ)")
    ap.add_argument("--list", action="store_true",
                    help="配線ゲート(WIRED_BLOCKS)の呼び出し件数を9関数すべて"
                         "一覧表示する(既定は問題がある行だけを表示)")
    args = ap.parse_args()

    spec_path = Path(args.spec).resolve()
    src_dir = Path(args.src).resolve()

    print("=" * 78)
    print("prompt_diff レポート - 15章 <-> modPrompts*/modSchemas* の一致検査")
    print("=" * 78)

    if not spec_path.exists():
        print(f"15章が見つかりません: {spec_path}")
        return 1

    bodies, section_missing = extract_spec_bodies(spec_path)
    print(f"15章から抽出した本文: {len(bodies)}件 (対応表 15章§10.2 の"
          f"{sum(len(f) for _m, f, _mod in SECTION_MAP)}関数中)")
    if section_missing:
        print(f"  注意: 節またはフェンスが見つからなかった関数: {section_missing}")

    if not src_dir.exists():
        print(f"\n対象0件: 突合先ディレクトリがありません({src_dir})。")
        print("modPrompts* / modSchemas* は T-23 で実装します。")
        if args.strict:
            print("  --strict 指定: 突合先が存在しない=未実装を差分として扱い、不合格"
                  "(exit 1)にします(T-23の最終確認用。実装が揃えば 0 diff を要求)。")
            return 1
        print("  スキップ(--strict なしのため exit 0)。")
        return 0

    files, code = extract_code_bodies(src_dir)
    if not files:
        print(f"\n対象0件: {src_dir} に modPrompts*.bas / modSchemas*.bas が"
              "1本もありません。")
        print("これらは T-23(W2)で実装します。実装後は本スクリプトが0 diffを要求します。")
        if args.strict:
            print("  --strict 指定: 突合対象が0本=未実装を差分として扱い、不合格"
                  "(exit 1)にします(T-23の最終確認用)。")
            return 1
        print("  スキップ(--strict なしのため exit 0)。")
        return 0

    print(f"突合対象ファイル: {[f.name for f in files]}")
    print("-" * 78)

    diffs = 0
    skipped: list[str] = []
    matched = 0
    for _marker, funcs, module in SECTION_MAP:
        for fn in funcs:
            want = bodies.get(fn)
            if want is None:
                diffs += 1
                print(f"  DIFF  {fn}: 15章側の本文を抽出できませんでした"
                      f"(§10.2の対応表と節見出しを確認してください)")
                continue
            if fn not in code:
                skipped.append(f"{fn}({module})")
                continue
            path, got, err = code[fn]
            if got is None:
                diffs += 1
                print(f"  DIFF  {fn} [{path.name}]: 戻り値を評価できません: {err}")
                continue
            if got != want:
                diffs += 1
                print(f"  DIFF  {fn} [{path.name}]: 15章と不一致")
                show_diff(fn, want, got)
                continue
            matched += 1

    diffs += run_schema_checks(code)

    print("-" * 78)
    print("配線ゲート (裁定書37 B-02。15章が定める Block* 9関数の呼び出し件数)")
    wiring_counts, wiring_ok = block_wiring_counts()
    wiring_problems = check_block_wiring() + check_guard_callsites()
    if not wiring_ok:
        print("  [配線] src/ に .bas が1本も読めません(fail-closed)")
    elif args.list or wiring_problems:
        for fn in sorted(WIRED_BLOCKS):
            status = "OK" if wiring_counts[fn] > 0 else "NG"
            print(f"  {status:<4}  {fn} [15章{WIRED_BLOCKS[fn]}]: "
                  f"{wiring_counts[fn]}件")
    else:
        print(f"  OK    {len(WIRED_BLOCKS)}関数すべて1件以上配線済み")
    for msg in wiring_problems:
        if msg.startswith("[代入点]"):
            print(f"  NG    {msg}")
    diffs += len(wiring_problems)

    print("-" * 78)
    print(f"一致: {matched}件 / 差分(スキーマJSON違反・配線ゲートを含む): {diffs}件 / "
          f"未実装(スキップ): {len(skipped)}件")
    if skipped:
        print(f"  未実装: {', '.join(skipped)}")
        if args.strict:
            diffs += len(skipped)
            print("  --strict 指定のため未実装分も差分に数えました。")
        else:
            print("  ※ T-23の受入条件は「未実装0件かつ差分0件」です"
                  "(最終確認は --strict を付けて実行してください)。")
    print(f"結果: 差分 {diffs} 件(exit code)")
    return min(diffs, 125)


if __name__ == "__main__":
    sys.exit(main())
