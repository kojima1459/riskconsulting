#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
render_proposal.py - 実物のサンプル提案書(Wide 22枚)を出す(20章§10・裁定書38 班C)

================================================================================
役割:
    tools/render_report.py と同型。LibreOffice(soffice --headless)へ純文字列
    モジュール一式を読み込ませ、mock素材(modMockLlm.ResponseById の MK-S2-RNW /
    MK-S5)を入力として modExportProposal の**純組立関数**
        BuildProposalMetaJson(...) -> 20章§3 の meta
        BuildProposalHtml(...)     -> HTML全文
    を実行し、dist/サンプル提案書.html へ書き出す。実行機構は
    tools/run_lo_tests.py を import して流用する(二重管理しない)。

2つのモード(裁定書38 班C「standard と --faithful の2モード」):
    standard(既定) : mock素材そのまま。22枚すべてに本文が載る「見せる側」。
    --faithful     : **当方が足さない**素の状態を見る。提案書JSONの任意ブロック
                     (hard_risks / ideas / four / theme_table / decisions /
                      share_items / notes)を空にして流し、
                     **欠けても22枚が描かれ、空の枚には20章§4.1 の1行が出る**
                     ことを実測する。この規定の回帰網はここだけである。

使い方:
    python3 tools/render_proposal.py
    python3 tools/render_proposal.py --faithful
    python3 tools/render_proposal.py --out <path>
    # exit code: 0 = 生成+検査OK / 1 = 生成できたが検査NG / 2 = 生成できず

検査(20章§10 の11項目)は check_* 関数がそのまま持つ。期待値は**20章と15章と
対訳表のMarkdownからパースする**(ツール側に写経しない=二重管理にしない)。
================================================================================
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

TOOLS_DIR = Path(__file__).resolve().parent
REPO_ROOT = TOOLS_DIR.parent
sys.path.insert(0, str(TOOLS_DIR))

import run_lo_tests as lo  # noqa: E402  (実行機構の流用。二重管理しない)

DEFAULT_OUT = REPO_ROOT / "dist" / "サンプル提案書.html"
FAITHFUL_OUT = REPO_ROOT / "dist" / "サンプル提案書_faithful.html"
SPEC20 = REPO_ROOT / "docs" / "spec" / "20_提案書Wideテンプレート仕様.md"
SPEC15 = REPO_ROOT / "docs" / "spec" / "15_プロンプトとJSONスキーマ.md"
GLOSSARY = REPO_ROOT / "docs" / "design" / "提案書_wide" / "対訳表_社内語から顧客語.md"

RENDER_MODULES = [
    "modUtil", "modUtilText", "modUtilPath", "modJsonLite", "modLog",
    "modProposalHtml1", "modProposalHtml2", "modProposalHtml3", "modProposalHtml4",
    "modExportProposal",
    # BuildProposalData が S2 由来の文字列へ対訳表の機械置換を掛けるので必須。
    "modValidate4",
    "modMockLlm", "modMockLlm2", "modMockLlm3", "modMockLlm4",
]

BANNED_JS = ["innerHTML", "insertAdjacentHTML", "document.write", "outerHTML"]
EXTERNAL_URL = re.compile(r"https?://|(?<![:\w])//cdn[\w.-]*/", re.IGNORECASE)

# 20章§5.1「テンプレが定義してよいCSS変数の閉じた一覧(19個)」。
THEME_VARS = [
    "--slide-w", "--slide-h", "--font-sans", "--font-base", "--line-height",
    "--ink", "--sub", "--paper", "--line", "--brand", "--brand-deep",
    "--accent", "--soft", "--heat-1", "--heat-2", "--heat-3", "--heat-4",
    "--note-bg", "--pad-x",
]

# 20章§3「DATAに入れないもの(構造的に持たない)」。
FORBIDDEN_IN_DATA = [
    "dossier_tier", "quality_mode", "case_type", "round_no", "s4_variant",
    "talk_script", "field_insights", "taboo", "report_path",
]

SAMPLE_COMPANY = "株式会社浜松スイーツファクトリー"
SAMPLE_GENERATED_AT = "2026/09/12 10:00:00"
SAMPLE_APP_VERSION = "2.4.0"
SAMPLE_REVIEWER = "浜松支店 山田"
SAMPLE_REVIEWED_AT = "2026/09/12 10:05:00"

# --faithful で空にするブロック(20章§4「データが無い枚も必ず描く」の実測)。
FAITHFUL_EMPTY_KEYS = ["hard_risks", "ideas", "four", "theme_table",
                       "decisions", "share_items", "notes"]
# 見出し(headline)も空にする。配列だけを空にしても各枚が見出しの一文を描いて
# しまい「本文が1つも無い枚」が作れないため(20章§4.1 の1行が出る条件)。
FAITHFUL_EMPTY_OBJS = ["headline"]


def basic_driver(out_url: str, faithful: bool) -> str:
    """LO Basic のドライバ(RenderMainモジュール)。素材の間引きはここだけで行う。"""
    shaping = ""
    if faithful:
        for key in FAITHFUL_EMPTY_KEYS:
            shaping += f'    s5 = EmptyArray(s5, "{key}")\n'
        for key in FAITHFUL_EMPTY_OBJS:
            shaping += f'    s5 = EmptyObject(s5, "{key}")\n'
    return (
        "Option Explicit\n"
        "\n"
        "' --- 素材の間引き(このモジュールだけの都合。製品コードには無い) ---\n"
        "Function EmptyArray(ByVal src As String, ByVal keyName As String) As String\n"
        "    Dim anchorText As String\n"
        '    anchorText = """" & keyName & """:"\n'
        "    Dim p As Long, q As Long, depth As Long, i As Long\n"
        "    EmptyArray = src\n"
        "    p = InStr(1, src, anchorText)\n"
        "    If p = 0 Then Exit Function\n"
        "    q = InStr(p, src, \"[\")\n"
        "    If q = 0 Then Exit Function\n"
        "    depth = 0\n"
        "    i = q\n"
        "    Do While i <= Len(src)\n"
        '        If Mid(src, i, 1) = "[" Then depth = depth + 1\n'
        '        If Mid(src, i, 1) = "]" Then\n'
        "            depth = depth - 1\n"
        "            If depth = 0 Then Exit Do\n"
        "        End If\n"
        "        i = i + 1\n"
        "    Loop\n"
        "    If depth <> 0 Then Exit Function\n"
        '    EmptyArray = Left(src, q) & Mid(src, i)\n'
        "End Function\n"
        "\n"
        "Function EmptyObject(ByVal src As String, ByVal keyName As String) As String\n"
        "    Dim anchorText As String\n"
        '    anchorText = """" & keyName & """:"\n'
        "    Dim p As Long, q As Long, depth As Long, i As Long\n"
        "    EmptyObject = src\n"
        "    p = InStr(1, src, anchorText)\n"
        "    If p = 0 Then Exit Function\n"
        '    q = InStr(p, src, "{")\n'
        "    If q = 0 Then Exit Function\n"
        "    depth = 0\n"
        "    i = q\n"
        "    Do While i <= Len(src)\n"
        '        If Mid(src, i, 1) = "{" Then depth = depth + 1\n'
        '        If Mid(src, i, 1) = "}" Then\n'
        "            depth = depth - 1\n"
        "            If depth = 0 Then Exit Do\n"
        "        End If\n"
        "        i = i + 1\n"
        "    Loop\n"
        "    If depth <> 0 Then Exit Function\n"
        "    EmptyObject = Left(src, q) & Mid(src, i)\n"
        "End Function\n"
        "\n"
        "Sub WriteUtf8(ByVal fileUrl As String, ByVal bodyText As String)\n"
        "    Dim oSFA As Object, oOut As Object, oText As Object\n"
        '    Set oSFA = createUnoService("com.sun.star.ucb.SimpleFileAccess")\n'
        "    If oSFA.exists(fileUrl) Then oSFA.kill(fileUrl)\n"
        "    Set oOut = oSFA.openFileWrite(fileUrl)\n"
        '    Set oText = createUnoService("com.sun.star.io.TextOutputStream")\n'
        "    oText.setOutputStream(oOut)\n"
        '    oText.setEncoding("UTF-8")\n'
        "    oText.writeString(bodyText)\n"
        "    oText.closeOutput()\n"
        "End Sub\n"
        "\n"
        "Sub Main\n"
        "    Dim s2 As String, s5 As String\n"
        '    s2 = modMockLlm.ResponseById("MK-S2-RNW")\n'
        '    s5 = modMockLlm.ResponseById("MK-S5")\n'
        f"{shaping}"
        "    Dim metaJson As String\n"
        "    metaJson = modExportProposal.BuildProposalMetaJson( _\n"
        f'        "{SAMPLE_COMPANY}", s5, "{SAMPLE_GENERATED_AT}", '
        f'"{SAMPLE_APP_VERSION}", _\n'
        f'        "{SAMPLE_REVIEWER}", "{SAMPLE_REVIEWED_AT}")\n'
        "    Dim docText As String\n"
        "    docText = modExportProposal.BuildProposalHtml(metaJson, s2, s5)\n"
        f'    WriteUtf8 "{out_url}", docText\n'
        "End Sub\n"
    )


DOM_STUB_JS = r"""
'use strict';
// 20章§1 が許す操作(createElement / textContent / setAttribute / appendChild /
// getElementById / addEventListener / style.setProperty)だけを持つ最小DOMスタブ。
// これで足りること自体が「innerHTML系を使っていない」ことの裏取りになる。
const fs = require('fs');
const html = fs.readFileSync(process.argv[2], 'utf8');

let byId = Object.create(null);
function Style() { this._p = Object.create(null); }
Style.prototype.setProperty = function (k, v) { this._p[k] = String(v); };
function El(tag) {
  this.tagName = tag; this.nodeType = 1; this.childNodes = [];
  this.attrs = Object.create(null); this.className = ''; this._text = '';
  this.parentNode = null; this.style = new Style(); this.clientWidth = 1280;
}
Object.defineProperty(El.prototype, 'textContent', {
  get() { return this._text; },
  set(v) { this._text = String(v); this.childNodes = []; }
});
El.prototype.appendChild = function (c) {
  c.parentNode = this; this.childNodes.push(c); return c;
};
El.prototype.setAttribute = function (k, v) {
  this.attrs[k] = String(v);
  if (k === 'id') { byId[String(v)] = this; }
};
El.prototype.addEventListener = function () {};

const scripts = [];
const re = /<script>([\s\S]*?)<\/script>/g;
let m;
while ((m = re.exec(html)) !== null) { scripts.push(m[1]); }
if (scripts.length !== 2) { console.error('SCRIPT_BLOCKS=' + scripts.length); process.exit(3); }

byId = Object.create(null);
const root = new El('body');
const deck = new El('main');
deck.setAttribute('id', 'deck');
root.appendChild(deck);
const all = [];
global.document = {
  body: root,
  createElement: (t) => { const n = new El(t); all.push(n); return n; },
  getElementById: (id) => (byId[id] || null),
  getElementsByClassName: (cls) => all.filter((n) => n.className === cls)
};
global.window = { addEventListener: function () {} };
global.location = { search: '' };

(0, eval)(scripts[0]);            // var DATA=JSON.parse("...")
(0, eval)(scripts[1]);            // ランタイム(登録配列の走査・描画)

const ids = [];
const nos = [];
const texts = [];
(function walk(n) {
  if (n.attrs && n.attrs.id) { ids.push(n.attrs.id); }
  if (n.attrs && n.attrs['data-no']) { nos.push(Number(n.attrs['data-no'])); }
  if (n._text) { texts.push(n._text); }
  for (const c of n.childNodes) { walk(c); }
})(root);
console.log(JSON.stringify({ ids: ids, nos: nos, texts: texts }));
"""


def run_render(soffice: str, work_dir: Path, faithful: bool, verbose: bool) -> str | None:
    all_modules = dict(lo.discover_modules(REPO_ROOT / "src"))
    type_blocks = lo.collect_public_type_blocks(all_modules)

    modules: dict[str, str] = {}
    missing = []
    for name in RENDER_MODULES:
        path = all_modules.get(name)
        if path is None:
            missing.append(name)
            continue
        modules[name] = path.read_text(encoding="utf-8", errors="replace")
    if missing:
        print(f"[render_proposal] FAIL: 未実装のモジュールがあります: {', '.join(missing)}")
        return None

    out_file = work_dir / "proposal.html"
    out_url = "file://" + out_file.as_posix()
    modules["RenderMain"] = basic_driver(out_url, faithful)

    profile_dir = work_dir / "profile_proposal"
    template = lo.ensure_template_profile(soffice, verbose)
    lo.fresh_profile_copy(template, profile_dir)
    lo.write_library(profile_dir, "RpnProposal", modules, type_blocks)
    lo.register_libraries(profile_dir, ["RpnProposal"])

    uri = ("vnd.sun.star.script:RpnProposal.RenderMain.Main"
           "?language=Basic&location=application")
    rc, out, err = lo.run_uri(soffice, profile_dir, uri, 180)
    if not out_file.exists():
        print(f"[render_proposal] FAIL: HTMLが生成されませんでした(soffice exit={rc})")
        if verbose:
            print(f"  stdout: {out.strip()}\n  stderr: {err.strip()}")
        return None
    return out_file.read_text(encoding="utf-8", errors="replace")


# ==============================================================================
# 20章 §4.2 の登録表・§4.1 の1行・§8 の免責を Markdown から読む(写経しない)
# ==============================================================================
SLIDE_ROW = re.compile(r"^\|\s*(\d{1,2})\s*\|\s*([a-z][a-z0-9-]*)\s*\|\s*"
                       r"(cover|section|content|back)\s*\|\s*([^|]*?)\s*\|")


def parse_slides(spec: str) -> list[tuple[int, str, str, str]]:
    out = []
    for line in spec.splitlines():
        mm = SLIDE_ROW.match(line)
        if not mm:
            continue
        title = mm.group(4)
        if title.startswith("（表紙") or title.startswith("（裏表紙"):
            title = ""
        out.append((int(mm.group(1)), mm.group(2), mm.group(3), title))
    return out


def parse_disclaimer(spec: str) -> str:
    body = spec.split("## 8. ")[1]
    mm = re.search(r"`(本資料は、[^`]+)`", body)
    return mm.group(1) if mm else ""


def parse_todo_line(spec: str) -> str:
    body = spec.split("### 4.1 ")[1].split("### 4.2")[0]
    mm = re.search(r"`(この項目は[^`]+)`", body)
    return mm.group(1) if mm else ""


def parse_glossary_terms(text: str) -> list[str]:
    """対訳表の「| n | 社内語 | 顧客語 |」行から社内語と顧客語の対を拾う。"""
    out = []
    for line in text.splitlines():
        mm = re.match(r"^\|\s*(\d{1,3})\s*\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|\s*$", line)
        if mm:
            out.append((mm.group(2), mm.group(3)))
    return out


def spec15_s5_system() -> str:
    """15章§5.6 の system フェンス本文。"""
    text = SPEC15.read_text(encoding="utf-8")
    body = text.split("### system（BuildS5System）")[1]
    parts = body.split("```")
    return parts[1] if len(parts) > 2 else ""


# ==============================================================================
# 検査
# ==============================================================================
def check_source(html: str, slides: list) -> list[str]:
    problems = []
    if '<meta charset="utf-8"' not in html[:400]:
        problems.append('<head> の先頭付近に <meta charset="utf-8"> がありません(20章§9.1)')
    flat = html.replace("\n", "")
    for no, slug, master, title in slides:
        row = f"{{no:{no},slug:'{slug}',master:'{master}',title:'{title}'"
        if row not in flat:
            problems.append(
                f"登録表の{no}枚目が20章§4.2と一致しません"
                f"(期待 no:{no} slug:{slug} master:{master} title:{title})")
    for word in BANNED_JS:
        if word in html:
            problems.append(f"禁止された書込口 {word} が出現します(20章§1)")
    if "JSON.parse(" not in html:
        problems.append("DATAが JSON.parse 形式で埋まっていません(20章§9.1)")
    if "@page{size:A4 landscape" not in flat:
        problems.append("@page{size:A4 landscape} がありません(20章§6②)")
    for mm in EXTERNAL_URL.finditer(html):
        problems.append(f"外部URLが出現します: {html[mm.start():mm.start() + 40]!r}(20章§1)")
    return problems


def check_theme_css(html: str) -> list[str]:
    problems = []
    m = re.search(r":root\{(.*?)\n\}", html, re.S)
    if m is None:
        problems.append("CSS変数の :root{ ... } ブロックが見つかりません(20章§5.1)")
        return problems
    body = m.group(1)
    declared = re.findall(r"(--[a-z0-9-]+)\s*:", body)
    if sorted(declared) != sorted(THEME_VARS):
        missing = sorted(set(THEME_VARS) - set(declared))
        extra = sorted(set(declared) - set(THEME_VARS))
        problems.append(f"CSS変数の19個が過不足です(不足={missing} 余分={extra})")
    style = re.search(r"<style>(.*?)</style>", html, re.S)
    if style is not None:
        common = style.group(1).replace(m.group(0), "")
        bad = [h for h in re.findall(r"#[0-9A-Fa-f]{3,6}\b", common) if h.lower() != "#fff"]
        if bad:
            problems.append(f":root の外に生の色指定があります: {sorted(set(bad))}(20章§5.1)")
    return problems


def check_data_literal(html: str) -> list[str]:
    problems = []
    pre = 'var DATA=JSON.parse("'
    a = html.find(pre)
    if a < 0:
        return ['DATAが `var DATA=JSON.parse("...")` の形で埋まっていません(20章§9.1)']
    a += len(pre)
    b = html.find('");', a)
    if b < 0:
        return ["DATAの文字列リテラルが閉じていません(20章§9.1)"]
    lit = html[a:b]
    bad = lit.find("<")
    if bad >= 0:
        problems.append(
            f"DATAの文字列リテラル内に生の `<` があります(位置{bad})。"
            f"20章§9.1 は「すべての `<` を \\u003C へ」を要求します")
    for key in FORBIDDEN_IN_DATA:
        if '\\"' + key + '\\"' in lit or f'"{key}"' in lit:
            problems.append(
                f"DATAに内部の値 `{key}` が入っています"
                f"(20章§3「入れてから隠すのではなく置く枝を持たない」)")
    n_open = len(re.findall(r"<script>", html))
    n_close = len(re.findall(r"</script>", html))
    if (n_open, n_close) != (2, 2):
        problems.append(f"<script>ブロックが2組ではありません(開き{n_open}/閉じ{n_close})")
    return problems


def check_literals(html: str, disclaimer: str, todo: str) -> list[str]:
    problems = []
    if not disclaimer:
        problems.append("20章§8の免責固定文を読み取れませんでした")
    elif ("'" + disclaimer + "'") not in html:
        problems.append(
            f"20章§8の免責固定文が1本のJSリテラルとして入っていません: [{disclaimer}]")
    if not todo:
        problems.append("20章§4.1の「データが無い枚の1行」を読み取れませんでした")
    elif ("'" + todo + "'") not in html:
        problems.append(f"20章§4.1の1行が逐語で入っていません: [{todo}]")
    return problems


def check_glossary(problems_sink: list[str]) -> None:
    """20章§10-11: 対訳表の全語が15章§5.6 の system 本文に現れること。"""
    try:
        pairs = parse_glossary_terms(GLOSSARY.read_text(encoding="utf-8"))
    except OSError:
        problems_sink.append(f"対訳表 {GLOSSARY} を読めません")
        return
    if len(pairs) < 30:
        problems_sink.append(
            f"対訳表の語が{len(pairs)}件です(裁定書38 班C は30語以上を求めています)")
    system = spec15_s5_system()
    if not system:
        problems_sink.append("15章§5.6 の system フェンスを読み取れませんでした")
        return
    for src, dst in pairs:
        if f"{src} -> {dst}" not in system:
            problems_sink.append(
                f"対訳表の「{src} -> {dst}」が15章§5.6 の system 本文にありません"
                f"(対訳表とプロンプトが二重管理になっています)")


def check_dom(html_path: Path, slides: list, todo: str, faithful: bool,
              verbose: bool) -> list[str]:
    node = shutil.which("node") or shutil.which("nodejs")
    if node is None:
        print("[render_proposal] (node が無いためDOM検査はスキップしました)")
        return []
    stub = html_path.parent / "dom_stub_proposal.js"
    stub.write_text(DOM_STUB_JS, encoding="utf-8")
    proc = subprocess.run([node, str(stub), str(html_path)],
                          capture_output=True, text=True, timeout=120)
    if proc.returncode != 0:
        return [f"DOM検査でページのJSが落ちました: {proc.stderr.strip()[:400]}"]
    try:
        got = json.loads(proc.stdout.strip().splitlines()[-1])
    except (ValueError, IndexError):
        return [f"DOM検査の出力を読めませんでした: {proc.stdout.strip()[:200]}"]

    problems = []
    ids = set(got["ids"])
    for no, slug, _master, _title in slides:
        if ("sl-" + slug) not in ids:
            problems.append(f"描画後のDOMに sl-{slug} がありません({no}枚目)")
    nos = sorted(got["nos"])
    if nos != list(range(1, len(slides) + 1)):
        problems.append(
            f"data-no が 1..{len(slides)} の連番ではありません(実際={nos})"
            f"。20章§4「枚数は22枚で固定」")
    if verbose:
        print(f"[render_proposal] 生成されたid: {sorted(ids)}")

    texts = got["texts"]
    hit = sum(1 for t in texts if t == todo)
    if faithful:
        # --faithful は任意ブロックを空にした素材なので、20章§4.1 の1行が
        # 出ていなければ「欠けても22枚描く」規定が働いていない。
        if hit == 0:
            problems.append(
                "--faithful(任意ブロックを空にした素材)なのに20章§4.1の1行が"
                "1枚も出ていません(空の枚を描いていない疑い)")
    else:
        if hit > 0:
            problems.append(
                f"mock素材そのままなのに20章§4.1の1行が{hit}枚に出ています"
                f"(mock が22枚ぶんのデータを満たしていない)")
    return problems


def main() -> int:
    ap = argparse.ArgumentParser(description="サンプル提案書(Wide 22枚)を生成する")
    ap.add_argument("--out", default=None, help="出力先(既定 dist/サンプル提案書.html)")
    ap.add_argument("--faithful", action="store_true",
                    help="任意ブロックを空にした素材で流す(欠けても22枚出ることの実測)")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args()

    out_path = Path(args.out) if args.out else (FAITHFUL_OUT if args.faithful else DEFAULT_OUT)

    spec = SPEC20.read_text(encoding="utf-8")
    slides = parse_slides(spec)
    if len(slides) != 22:
        print(f"[render_proposal] FAIL: 20章§4.2の表から{len(slides)}枚しか読めません(期待22)")
        return 2
    disclaimer = parse_disclaimer(spec)
    todo = parse_todo_line(spec)

    soffice = lo.find_soffice()
    work_dir = Path(tempfile.mkdtemp(prefix="rpn_proposal_"))
    try:
        html = run_render(soffice, work_dir, args.faithful, args.verbose)
        if html is None:
            return 2
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(html, encoding="utf-8-sig", newline="")
        print(f"[render_proposal] 生成: {out_path} ({len(html)}字)")

        problems = check_source(html, slides)
        problems += check_theme_css(html)
        problems += check_data_literal(html)
        problems += check_literals(html, disclaimer, todo)
        check_glossary(problems)
        problems += check_dom(out_path, slides, todo, args.faithful, args.verbose)
        if problems:
            print("[render_proposal] NG:")
            for p in problems:
                print(f"  - {p}")
            return 1
        print(f"[render_proposal] OK: 全{len(slides)}枚の登録と描画後DOM、"
              f"20章§5.1の19変数・§8の免責・対訳表との一致を確認しました。")
        return 0
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
