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


def basic_driver(out_url: str, faithful: bool,
                 gl_in_url: str = "", gl_out_url: str = "") -> str:
    """LO Basic のドライバ(RenderMainモジュール)。素材の間引きはここだけで行う。

    gl_in_url / gl_out_url を渡すと、同じ LO の実行の中で
    **実物の modValidate4.SoftenTaboo** に1行ずつ入力を通し、
    「入力<TAB>1回目の出力<TAB>2回目の出力」を書き出す(対訳表の実効出力の検問。
    裁定書42 §1)。LO の起動は1回のままにしたいのでドライバへ相乗りさせる。
    """
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
        + (f'    GlossarySweep "{gl_in_url}", "{gl_out_url}"\n'
           if gl_in_url and gl_out_url else "")
        + "End Sub\n"
        "\n"
        "' 対訳表の実効出力(裁定書42 §1)。入力を1行ずつ実物の SoftenTaboo へ通し、\n"
        "'   2回通した結果も並べて書き出す(冪等の実測も同じ1回の実行で済ませる)。\n"
        "Sub GlossarySweep(ByVal inUrl As String, ByVal outUrl As String)\n"
        "    Dim oSFA As Object, oIn As Object, oTin As Object\n"
        "    Dim lineText As String, acc As String, r1 As String, r2 As String\n"
        "    Dim n1 As Long, n2 As Long\n"
        '    Set oSFA = createUnoService("com.sun.star.ucb.SimpleFileAccess")\n'
        "    Set oIn = oSFA.openFileRead(inUrl)\n"
        '    Set oTin = createUnoService("com.sun.star.io.TextInputStream")\n'
        "    oTin.setInputStream(oIn)\n"
        '    oTin.setEncoding("UTF-8")\n'
        "    Do While Not oTin.isEOF()\n"
        "        lineText = oTin.readLine()\n"
        "        ' 最後の1行だけ改行が付いたまま返ることがある(実測)。付いたまま\n"
        "        '   だと出力のTSVが1件分ずれて、その1件が黙って検査から落ちる。\n"
        "        Do While Len(lineText) > 0\n"
        "            If Right$(lineText, 1) <> Chr(10) And Right$(lineText, 1) <> Chr(13) Then\n"
        "                Exit Do\n"
        "            End If\n"
        "            lineText = Left$(lineText, Len(lineText) - 1)\n"
        "        Loop\n"
        "        If Len(lineText) > 0 Then\n"
        "            r1 = modValidate4.SoftenTaboo(lineText, n1)\n"
        "            r2 = modValidate4.SoftenTaboo(r1, n2)\n"
        "            acc = acc & lineText & Chr(9) & r1 & Chr(9) & r2 & Chr(10)\n"
        "        End If\n"
        "    Loop\n"
        "    oTin.closeInput()\n"
        "    WriteUtf8 outUrl, acc\n"
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

// ページを1回描いて結果を集める。mutate は DATA を壊す関数(省略可)。
// 壊した素材でも「落ちずに22枚描き、壊れた枚に印が立つ」ことを見るため、
// 同じ手順を2回回す(裁定書40 S-M3。発表者ノートの例外で run() ごと死ぬと
// 提案書が白紙になるが、文字列 grep ではその違いが見えない)。
function render(mutate) {
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
  if (mutate) { mutate(global.DATA); }
  (0, eval)(scripts[1]);            // ランタイム(登録配列の走査・描画)

  const ids = [];
  const nos = [];
  const texts = [];
  let renderErrors = 0;
  (function walk(n) {
    if (n.attrs && n.attrs.id) { ids.push(n.attrs.id); }
    if (n.attrs && n.attrs['data-no']) { nos.push(Number(n.attrs['data-no'])); }
    if (n.attrs && n.attrs['data-render-error']) { renderErrors += 1; }
    if (n._text) { texts.push(n._text); }
    for (const c of n.childNodes) { walk(c); }
  })(root);
  return { ids: ids, nos: nos, texts: texts, renderErrors: renderErrors };
}

const normal = render(null);
// 発表者ノートの素材だけを壊す(20章§7。値源は S5 の notes[] で、
// s5_edited を人が直す経路もあるため現実に起こりうる形)。
const poisoned = render(function (D) { if (D && D.p) { D.p.notes = [null]; } });
console.log(JSON.stringify({
  ids: normal.ids, nos: normal.nos, texts: normal.texts,
  renderErrors: normal.renderErrors,
  poisonedIds: poisoned.ids, poisonedNos: poisoned.nos,
  poisonedErrors: poisoned.renderErrors
}));
"""


def run_render(soffice: str, work_dir: Path, faithful: bool, verbose: bool,
               gl_inputs: list[str] | None = None) -> str | None:
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
    gl_in_url = gl_out_url = ""
    if gl_inputs:
        gl_in = work_dir / "glossary_in.txt"
        gl_in.write_text("\n".join(gl_inputs) + "\n", encoding="utf-8")
        gl_in_url = "file://" + gl_in.as_posix()
        gl_out_url = "file://" + (work_dir / GLOSSARY_OUT_NAME).as_posix()
    modules["RenderMain"] = basic_driver(out_url, faithful, gl_in_url, gl_out_url)

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
        mm = re.match(r"^\|\s*(\d{1,3})\s*\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|"
                      r"\s*(?:replace|warn)\s*\|\s*$", line)
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


# ==============================================================================
# 対訳表の検問(裁定書43 §1-5。**3本立て**)
# ------------------------------------------------------------------------------
# 1. check_glossary_impl      宣言(対訳表 Markdown)⇔実装(modValidate4)の
#                             **全列突合**。対の社内語・顧客語・mode の3列と、
#                             §6 の**終端集合を1文字ずつ**、mode の語彙も見る。
# 2. check_glossary_effective LibreOffice で**実物の SoftenTaboo** を回し、
#                             全対 × 全文脈(終端集合の各文字・非終端の代表文字・
#                             活用語尾)の出力に**壊れの徴候**が1件も無いことを
#                             見る((a)同一文字3連続 (b)顧客語末尾と直後の重複
#                             (c)置換が起きたのに直後が終端集合でない)。加えて
#                             対訳表の宣言だけから組み立てた参照実装と1件ずつ
#                             比べる(値源と実装を別経路で辿る)。
# 3. 冪等                     2回通して変わらないこと(2 と同じ採取で見る)。
# 前波はここが「対の数と一部」しか見ておらず、実装の語尾リストを半分に削っても
# 全ゲートが緑だった。列を1つ足したら宣言と実装の両方に要る。
# ==============================================================================
GLOSSARY_OUT_NAME = "glossary_out.tsv"

# 非終端の代表文字と活用語尾(**終端集合は対訳表§6 の宣言から読む**ので、ここに
# は1文字も書かない=列挙を2箇所に置かない)。漢字/ひらがな/カタカナ/英数。
GL_NONTERM = [
    "額", "保険", "計画", "認証", "指標", "力", "策", "品", "型", "直す",
    "する", "した", "して", "している", "され", "させる", "できる",
    "て", "た", "ている", "ながら", "やすい", "ました", "ます", "ない",
    "など", "まで", "リスク", "シート", "ABC", "1",
]
GL_FRAME = "本件の{0}をご説明します。"
GL_MODES = ("replace", "warn")
# 行単位で LO へ渡す経路なので、タブ・CR・LF は入力に載せられない(TSV が壊れる)。
# この3文字は check_glossary_impl の**終端集合の1文字ずつの突合**で見る。
GL_UNSENDABLE = "\t\r\n"


def _md_section(text: str, head: str, next_heads: list[str]) -> str:
    """`head` で始まる節の本文(次の見出しの手前まで)。無ければ ""。"""
    if head not in text:
        return ""
    body = text.split(head, 1)[1]
    cut = len(body)
    for nh in next_heads:
        pos = body.find(nh)
        if pos >= 0:
            cut = min(cut, pos)
    return body[:cut]


def _md_rows(body: str) -> list[list[str]]:
    """Markdown の表の行(区切り行と見出し行を除く)をセルの配列で返す。"""
    out = []
    for line in body.splitlines():
        line = line.strip()
        if not line.startswith("|") or not line.endswith("|"):
            continue
        cells = [c.strip() for c in line[1:-1].split("|")]
        if not cells or set("".join(cells)) <= set("-: "):
            continue
        out.append(cells)
    return out


def term_from_md(text: str) -> tuple[str, list[str]]:
    """対訳表§6 の終端集合ブロック(フェンス)を読む。`U+XXXX` 表記も解く。"""
    problems: list[str] = []
    body = _md_section(text, "### 終端集合", ["## 6.1"])
    parts = body.split("```")
    if len(parts) < 3:
        return "", ["対訳表§6 の終端集合のブロックを読み取れませんでした"]
    chars: list[str] = []
    for line in parts[1].splitlines():
        line = line.strip()
        if not line:
            continue
        label, _sp, rest = line.partition(" ")
        if label == "コード":
            for tok in rest.split():
                mm = re.fullmatch(r"U\+([0-9A-Fa-f]{4,6})", tok)
                if mm:
                    chars.append(chr(int(mm.group(1), 16)))
                else:
                    problems.append(f"対訳表§6 の終端集合に読めない字があります: {tok}")
        else:
            chars.extend(rest.strip())
    dup = [c for c in set(chars) if chars.count(c) > 1]
    if dup:
        problems.append(f"対訳表§6 の終端集合に重複があります: {''.join(sorted(dup))}")
    return "".join(chars), problems


def glossary_from_md(text: str) -> tuple[list[tuple[str, str, str]], dict]:
    """対訳表 Markdown から (社内語, 顧客語, mode) の一覧と置換規則を読む。

    対を書く表は3つあり、**どれも mode 列を持つ**(裁定書43 §1-2):
    - §1〜§3 の番号付き46語 / §4.1 の代替語3語 / §6.4 の表記ゆれ2語
    """
    problems: list[str] = []
    rows: list[tuple[str, str, str]] = []

    for line in text.splitlines():
        mm = re.match(r"^\|\s*(\d{1,3})\s*\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|"
                      r"\s*([^|]*?)\s*\|\s*$", line)
        if mm:
            rows.append((mm.group(2), mm.group(3), mm.group(4)))

    sec41 = _md_section(text, "### 4.1 ", ["## 5."])
    for cells in _md_rows(sec41):
        if len(cells) >= 3 and cells[0] != "社内語":
            rows.append((cells[0], cells[1], cells[2]))

    sec64 = _md_section(text, "## 6.4 ", ["## 6.5"])
    for cells in _md_rows(sec64):
        if len(cells) >= 4 and cells[0] != "表記ゆれ":
            rows.append((cells[0], cells[2], cells[3]))

    for src, dst, mode in rows:
        if mode not in GL_MODES:
            problems.append(
                f"対訳表の「{src}」の mode が [{mode}] です"
                f"({' / '.join(GL_MODES)} のどちらかを必ず書きます)")

    term, term_problems = term_from_md(text)
    problems.extend(term_problems)
    rules = {"problems": problems, "term": term}
    body = _md_section(text, "## 6.5 ", ["## 6.6"])
    mm = re.search(r"V4_RUN_MAX = (\d+)", body)
    rules["run_max"] = int(mm.group(1)) if mm else 0
    mm = re.search(r"V4_UNDO_TAIL_MAX = (\d+)", body)
    rules["undo_tail"] = int(mm.group(1)) if mm else 0
    if not rules["run_max"] or not rules["undo_tail"]:
        problems.append("対訳表§6.5 から取り消し規則の字数を読み取れませんでした")
    return rows, rules


VBA_TOKEN = re.compile(r'"((?:[^"]|"")*)"|([A-Za-z_][A-Za-z0-9_]*)')
VBA_MODES = {"V4_REPLACE": "replace", "V4_WARN": "warn"}
VBA_TERM_TOKEN = re.compile(r'ChrW\(&H([0-9A-Fa-f]+)\)|([A-Za-z_]\w*)')
VBA_TERM_CONST = re.compile(r'^Private Const (V4_TERM_\w+) As String = "((?:[^"]|"")*)"',
                            re.MULTILINE)
VBA_CONSTANTS = {"vbTab": "\t", "vbCr": "\r", "vbLf": "\n"}


def glossary_from_vba(src: str) -> list[tuple[str, str, str]]:
    """modValidate4.bas の TabooPairs() が組み立てる行(社内語・顧客語・mode)。"""
    if "Public Function TabooPairs()" not in src:
        return []
    body = src.split("Public Function TabooPairs()", 1)[1].split("End Function", 1)[0]
    rows = []
    for line in body.splitlines():
        if "AdPair " not in line:
            continue
        texts, mode = [], ""
        for lit, ident in VBA_TOKEN.findall(line.split("AdPair ", 1)[1]):
            if ident:
                mode = VBA_MODES.get(ident, mode)
            else:
                texts.append(lit.replace('""', '"'))
        if len(texts) >= 2:
            rows.append((texts[0], texts[1], mode))
    return rows


def term_from_vba(src: str) -> tuple[str, list[str]]:
    """modValidate4.TermChars() が返す終端集合を**実装の字面から**組み立てる。"""
    consts = {m.group(1): m.group(2).replace('""', '"')
              for m in VBA_TERM_CONST.finditer(src)}
    if "Private Function TermChars() As String" not in src:
        return "", ["modValidate4.TermChars() を読み取れませんでした"]
    body = src.split("Private Function TermChars() As String", 1)[1]
    body = body.split("End Function", 1)[0]
    expr = body.split("TermChars =", 1)[1]
    problems, out = [], []
    for mm in VBA_TERM_TOKEN.finditer(expr):
        if mm.group(1):
            out.append(chr(int(mm.group(1), 16)))
        elif mm.group(2) in consts:
            out.append(consts[mm.group(2)])
        elif mm.group(2) in VBA_CONSTANTS:
            out.append(VBA_CONSTANTS[mm.group(2)])
        elif mm.group(2) == "_":
            continue                     # 行継続
        else:
            problems.append(
                f"modValidate4.TermChars() に読めない項があります: {mm.group(2)}")
    rest = VBA_TERM_TOKEN.sub("", expr)
    if set(rest) - set(" &_\r\n"):
        problems.append(
            f"modValidate4.TermChars() に想定外の式があります: {rest.strip()}")
    return "".join(out), problems


def check_glossary_impl(problems_sink: list[str]) -> int:
    """1本目: 宣言(対訳表)と実装(modValidate4)の**全列**を突き合わせる。"""
    md_rows, rules = glossary_from_md(GLOSSARY.read_text(encoding="utf-8"))
    problems_sink.extend(rules["problems"])
    vba_src = (REPO_ROOT / "src" / "app" / "modValidate4.bas").read_text(encoding="utf-8")
    vba_rows = glossary_from_vba(vba_src)
    if not vba_rows:
        problems_sink.append("modValidate4.bas の TabooPairs() を読み取れませんでした")
        return 0
    md_map = {s: (d, m) for s, d, m in md_rows}
    vba_map = {s: (d, m) for s, d, m in vba_rows}
    for word in sorted(set(md_map) | set(vba_map)):
        if word not in vba_map:
            problems_sink.append(
                f"対訳表の「{word}」が modValidate4.TabooPairs() にありません"
                "(宣言だけあって機械置換されない=顧客資料に社内語が残る)")
        elif word not in md_map:
            problems_sink.append(
                f"modValidate4.TabooPairs() の「{word}」が対訳表にありません"
                "(実装だけが知っている禁止語=誰も検算していない)")
        elif md_map[word] != vba_map[word]:
            problems_sink.append(
                f"「{word}」の顧客語/mode が対訳表と実装で違います"
                f"(対訳表={md_map[word]} / TabooPairs={vba_map[word]})")

    # 終端集合を**1文字ずつ**。ここが前波の死角(実装の語尾リストを半分に削っても
    # 全ゲートが緑だった)を閉じる本体である。
    vba_term, term_problems = term_from_vba(vba_src)
    problems_sink.extend(term_problems)
    md_term = rules["term"]
    for ch in sorted(set(md_term) - set(vba_term)):
        problems_sink.append(
            f"対訳表§6 の終端集合の U+{ord(ch):04X} が modValidate4 にありません")
    for ch in sorted(set(vba_term) - set(md_term)):
        problems_sink.append(
            f"modValidate4 の終端集合の U+{ord(ch):04X} が対訳表§6 にありません")
    if not md_term:
        problems_sink.append("対訳表§6 の終端集合が空です(検査を飛ばして緑にはしません)")
    # 取り消し規則(対訳表§6.5)。終端集合がある限り発火しない二重の安全網なので
    # 出力からは観測できない。**宣言の数字と実装の定数**、および「ReplaceOk が
    # UndoNeeded を呼ぶこと」を見て、規則を外す変異が必ず赤くなるようにする。
    for name, want in (("V4_RUN_MAX", rules["run_max"]),
                       ("V4_UNDO_TAIL_MAX", rules["undo_tail"])):
        mm = re.search(rf"Private Const {name} As Long = (\d+)", vba_src)
        if not mm:
            problems_sink.append(f"modValidate4 に {name} がありません(対訳表§6.5)")
        elif int(mm.group(1)) != want:
            problems_sink.append(
                f"取り消し規則の {name} が対訳表§6.5 の宣言({want})と違います"
                f"(実装={mm.group(1)})")
    body = vba_src.split("Private Function ReplaceOk(", 1)
    if len(body) < 2 or "UndoNeeded(" not in body[1].split("End Function", 1)[0]:
        problems_sink.append(
            "modValidate4.ReplaceOk が UndoNeeded を呼んでいません"
            "(対訳表§6.5 の取り消し規則が外れています)")
    # mode の語彙も実装と突き合わせる(印を増やしていないことの機械検査)。
    vba_modes = sorted(
        m.group(2) for m in re.finditer(
            r'Private Const (V4_REPLACE|V4_WARN) As String = "(\w+)"', vba_src))
    if vba_modes != sorted(GL_MODES):
        problems_sink.append(
            f"modValidate4 の mode の語彙が {vba_modes} です"
            f"(対訳表§6.1 の2区分 {sorted(GL_MODES)} だけにします)")
    return len(md_map)


# ---- 対訳表の宣言だけから組み立てる参照実装(VBA とは別経路で同じ答えを出す) ----
def _is_ascii_word(word: str) -> bool:
    return all(32 <= ord(c) <= 126 for c in word)


def _boundary_ok(hay: str, pos: int, word: str) -> bool:
    if not _is_ascii_word(word):
        return True
    if pos > 0 and hay[pos - 1].isascii() and hay[pos - 1].isalpha():
        return False
    end = pos + len(word)
    return not (end < len(hay) and hay[end].isascii() and hay[end].isalpha())


def _has_run(text: str, run_max: int) -> bool:
    n = 1
    for i in range(1, len(text)):
        if text[i] == text[i - 1]:
            n += 1
            if n >= run_max:
                return True
        else:
            n = 1
    return False


def _undo_needed(dst: str, hay: str, pos: int, rules: dict) -> bool:
    """取り消し規則(対訳表§6.5)。modValidate4.UndoNeeded と同じ判断。"""
    for k in range(1, rules["undo_tail"] + 1):
        if k <= len(dst) and hay[pos:pos + k] == dst[-k:]:
            return True
    seam = dst[-(rules["run_max"] - 1):] + hay[pos:pos + rules["run_max"] - 1]
    return _has_run(seam, rules["run_max"])


def _replace_ok(hay: str, pos: int, dst: str, rules: dict) -> bool:
    """置換してよい位置か(対訳表§6 の終端集合 + §6.5 の取り消し規則)。"""
    if pos < len(hay) and hay[pos] not in rules["term"]:
        return False
    return not _undo_needed(dst, hay, pos, rules)


def soften_reference(text: str, rows: list[tuple[str, str, str]], rules: dict,
                     passes: int = 4) -> str:
    """対訳表§6〜§6.6 の規約どおりに置換する参照実装(modValidate4 とは別実装)。"""
    use = sorted([r for r in rows if r[2] == "replace"], key=lambda r: -len(r[0]))
    heads = {r[0][0] for r in use}
    for _ in range(passes):
        out, i, hits, n = [], 0, 0, len(text)
        while i < n:
            hit = None
            if text[i] in heads:
                for src, dst, _mode in use:
                    if text[i:i + len(src)] != src:
                        continue
                    if not _boundary_ok(text, i, src):
                        continue
                    if not _replace_ok(text, i + len(src), dst, rules):
                        break            # 終端でない位置は打ち切る(規約2)
                    hit = (src, dst)
                    break
            if hit is None:
                out.append(text[i])
                i += 1
                continue
            src, dst = hit
            out.append(dst)
            i += len(src)
            hits += 1
        text = "".join(out)
        if hits == 0:
            break
    return text


def glossary_inputs(rows: list[tuple[str, str, str]], rules: dict) -> list[str]:
    """全対 × 全文脈の入力文(LO へ渡す1行1件)。

    文脈は「素」「終端集合の各文字(**対訳表の宣言から**)」「非終端の代表文字」
    「活用語尾」。対や終端集合を1つ足すと、検査も自動で増える。
    """
    suffixes = [""] + [c for c in rules["term"] if c not in GL_UNSENDABLE]
    suffixes += GL_NONTERM
    seen, out = set(), []
    for src, _dst, _mode in rows:
        for suf in suffixes:
            for body in (src + suf, GL_FRAME.format(src + suf)):
                if body not in seen:
                    seen.add(body)
                    out.append(body)
    return out


def check_glossary_effective(problems_sink: list[str], sweep_path: Path) -> int:
    """2本目と3本目: 実物の SoftenTaboo の出力に壊れの徴候が無いこと・冪等。"""
    if not sweep_path.exists():
        problems_sink.append(
            "対訳表の実効出力(SoftenTaboo)を LibreOffice で採取できませんでした"
            "(検査を飛ばして緑にはしません)")
        return 0
    rows, rules = glossary_from_md(GLOSSARY.read_text(encoding="utf-8"))
    term = set(rules["term"])
    checked, shown = 0, 0
    seen_inputs: set[str] = set()

    def report(msg: str) -> None:
        nonlocal shown
        if shown < 12:
            shown += 1
            problems_sink.append(msg)

    for line in sweep_path.read_text(encoding="utf-8").splitlines():
        cells = line.split("\t")
        if len(cells) < 3:
            continue
        body, got1, got2 = cells[0], cells[1], cells[2]
        checked += 1
        seen_inputs.add(body)
        # (a) 同一文字が3つ以上連続(入力に無かったものだけ)
        if _has_run(got1, rules["run_max"]) and not _has_run(body, rules["run_max"]):
            report(f"対訳表の実効出力に同一文字の{rules['run_max']}連続が出ました: "
                   f"[{body}] → [{got1}]")
        for src, dst, mode in rows:
            if mode != "replace" or dst in body or dst not in got1:
                continue
            p = got1.find(dst)
            while p >= 0:
                nxt = p + len(dst)
                # (b) 顧客語の末尾1〜3文字が直後の文字列と重複する
                for k in range(1, rules["undo_tail"] + 1):
                    if k <= len(dst) and got1[nxt:nxt + k] == dst[-k:]:
                        report(f"対訳表の実効出力で顧客語の末尾が直後と重複しました: "
                               f"[{body}] → [{got1}]")
                        break
                # (c) 置換が起きたのに、置換した位置の直後が終端集合でない
                if nxt < len(got1) and got1[nxt] not in term:
                    report(f"終端集合でない文脈で置換が起きました: "
                           f"[{body}] → [{got1}](「{src}」の直後は "
                           f"U+{ord(got1[nxt]):04X})")
                p = got1.find(dst, p + 1)
        # (d) 対訳表の宣言だけから組み立てた参照実装と一致すること
        want = soften_reference(body, rows, rules)
        if got1 != want:
            report(f"対訳表の実効出力が宣言と違います: [{body}] → 実装[{got1}] / "
                   f"対訳表どおりなら[{want}]")
        # 3本目: 冪等
        if got2 != got1:
            report(f"SoftenTaboo が冪等ではありません: [{body}] → [{got1}] → [{got2}]")
    if checked == 0:
        problems_sink.append("対訳表の実効出力を1件も検査できませんでした")
    # 入力した文脈が**1件残らず**採れていること(採取が欠けた分だけ検査が静かに
    # 減るのを防ぐ。実測で最後の1行が落ちていた)。
    lost = [s for s in glossary_inputs(rows, rules) if s not in seen_inputs]
    if lost:
        problems_sink.append(
            f"対訳表の実効出力の採取が{len(lost)}件欠けています"
            f"(例: [{lost[0]}])。検査を飛ばして緑にはしません")
    return checked


def spec15_s5_required() -> list[str]:
    """15章§5.6 Schema-S5 の**最外 required**(16キー)を仕様から読む。"""
    text = SPEC15.read_text(encoding="utf-8")
    body = text.split("### Schema-S5")[1]
    parts = body.split("```")
    if len(parts) < 3:
        return []
    schema = parts[1]
    if schema.lstrip().startswith("json"):
        schema = schema.lstrip()[4:]
    try:
        obj = json.loads(schema)
    except ValueError:
        return []
    req = obj.get("required")
    return list(req) if isinstance(req, list) else []


def impl_s5_required() -> list[str]:
    """modExportProposal.bas の EP_S5_REQUIRED("|"区切り)を実装から読む。"""
    src = (REPO_ROOT / "src" / "app" / "modExportProposal.bas").read_text(
        encoding="utf-8")
    lines = src.splitlines()
    body = ""
    for i, line in enumerate(lines):
        if "Const EP_S5_REQUIRED" not in line:
            continue
        body = line
        # VBAの行継続(末尾の ` _`)をたどって1本の宣言に戻す。
        while body.rstrip().endswith("_") and i + 1 < len(lines):
            i += 1
            body = body.rstrip()[:-1] + lines[i]
        break
    if not body:
        return []
    joined = "".join(re.findall(r'"([^"]*)"', body))
    return [k for k in joined.split("|") if k]


def check_s5_required(problems_sink: list[str]) -> None:
    """裁定書40 S-m: Schema-S5 の required の**3つ目の写し**を突き合わせる。

    値源は (1) 15章§5.6 の required 行 (2) modSchemas2.SchemaS5()
    (3) modExportProposal.EP_S5_REQUIRED の3箇所ある。(1)-(2) は
    tools/prompt_diff.py --strict が見ているが、(3) を見るゲートが無かったため、
    required が増減しても E0502 の必須キー検査だけが黙って古いままになる
    (=裁定書39 R2-12 で立てた関門が静かに穴になる)。ここで閉じる。
    """
    spec = spec15_s5_required()
    impl = impl_s5_required()
    if not spec:
        problems_sink.append(
            "15章§5.6 Schema-S5 の required を読み取れませんでした")
        return
    if not impl:
        problems_sink.append(
            "modExportProposal.bas の EP_S5_REQUIRED を読み取れませんでした")
        return
    if spec != impl:
        problems_sink.append(
            "EP_S5_REQUIRED が15章§5.6 Schema-S5 の required と一致しません"
            f"(15章={spec} / 実装={impl})。必須キー検査(E0502)が古い写しの"
            "ままになっています")


# ==============================================================================
# 確認導線(裁定書42 §2-3)。班X1 の対訳表の節とは独立した節である。
# ------------------------------------------------------------------------------
# なぜ要るか(ゲートの死角#7): W15 Round3 まで、確認導線(=確認者名が無ければ
# 顧客提示物を出さない)を守っていたのは lo-pure の純テスト4本だけだった。
# 統合レビューは modExportProposal.NeedsReviewMessage を「常に空を返す」へ潰す
# 変異で、render-p / render-pf / lint / doc-gate / action / ui / validate が
# **全部緑**になることを実測している。提案書に関わる機械検問は「22枚出るか」
# しか見ておらず、「**出てはいけない条件で出ないか**」を1件も見ていなかった。
#
# ここで何を測るか(すべて LibreOffice で**製品コードを実際に呼ぶ**):
#   (1) 出してはいけない: 空白類だけの確認者名12通り × 2経路
#       ・提案書  modExportProposal.GenerateProposalHtml が
#         reason=EP_NEED_REVIEW / outPath 空 / **出力先にファイルが増えない**
#       ・レポート modExportHtml.ReviewerOf が空文字(=18章§3.5 の免責が
#         「担当者が確認・編集したもの」へ切り替わらない)
#       この2経路を**同じ表**で測ることが肝である。裁定書40 S-m では提案書側
#       だけを直したため NBSP でレポートだけが「確認済み」になっていた
#       (=「片方だけ直す」型の3回目)。表を1つにすれば非対称は原理的に作れない。
#   (2) 出さなければいけない: 可視文字のある確認者名では提案書が**実際に
#       書き出される**(ファイルが存在し、出力の直前で止まっていない)。
#       これが無いと「全部断る」実装でも(1)が緑になる=出来レースになる。
#
# 空白類の一覧(BLANK_INPUTS)は**この検問が持つ独立した期待値**であり、
# 実装(modUtilText.HasVisibleText)から読み出さない。実装から読むと、実装を
# 壊した変異に合わせて期待値も一緒に動き、何も検査しないのと同じになる。
# 文言 EP_NEED_REVIEW だけは modExportProposal.bas の Const から読む(値源は
# 1つ=二重管理にしない。文言が変わったら検問も一緒に動いてよい)。
# ==============================================================================

# 出力してはいけない確認者名。(ラベル, LO Basic の式)。
# 足すときは modUtilText.HasVisibleText の Select Case にも同じ文字を足すこと。
BLANK_INPUTS: list[tuple[str, str]] = [
    ("空", '""'),
    ("半角空白", '" "'),
    ("全角空白", "ChrW(12288) & ChrW(12288)"),
    ("TAB", "Chr(9)"),
    ("LF", "Chr(10)"),
    ("CR", "Chr(13)"),
    ("VT", "Chr(11)"),
    ("FF", "Chr(12)"),
    ("NBSP", "ChrW(160)"),
    ("ZWSP", "ChrW(8203)"),
    ("BOM", "ChrW(65279)"),
    ("混在", 'Chr(9) & ChrW(160) & ChrW(12288) & ChrW(8203) & " " & ChrW(65279)'),
]
# 出力しなければいけない確認者名(出来レース防止の対照)。
VISIBLE_INPUTS: list[tuple[str, str]] = [
    ("氏名", '"浜松支店 山田"'),
    ("空白で囲んだ1字", 'ChrW(160) & "田" & ChrW(8203)'),
]

# 確認導線の実測に要るモジュール。製品コードは**本物**を読み込み、
# 案件データ・設定・ログの3つだけをスタブに差し替える(統合レビューと同じ手)。
REVIEW_REAL_MODULES = [
    "modUtil", "modUtilText", "modUtilPath", "modJsonLite", "modPii",
    "modValidate4", "modProposalHtml1", "modProposalHtml2",
    "modProposalHtml3", "modProposalHtml4", "modExportProposal",
    "modExportHtml",
    "modMockLlm", "modMockLlm2", "modMockLlm3", "modMockLlm4",
    "modAppTypes", "modTypes",
]

REVIEW_STUBS: dict[str, str] = {
    "modCaseStore": (
        "Option Explicit\n"
        "Public Function IsValidCaseId(ByVal caseId As String) As Boolean\n"
        '    IsValidCaseId = (Left(caseId, 2) = "C-")\n'
        "End Function\n"
        "Public Function LoadData(ByVal caseId As String, ByVal k As String) As String\n"
        '    If k = "s5_json" Then LoadData = StubBag.S5()\n'
        "End Function\n"
        "Public Function ResolveStepJson(ByVal caseId As String, ByVal n As Long) As String\n"
        "    If n = 2 Then ResolveStepJson = StubBag.S2()\n"
        "End Function\n"
    ),
    "modCaseRead": (
        "Option Explicit\n"
        "Public Function ReadCaseCtx(ByVal caseId As String, ByRef ctx As TCaseCtx, _\n"
        "                            ByRef roundNo As Long, ByRef qualityMode As String, _\n"
        "                            ByRef s4Variant As String, ByRef dossierTier As String) As Boolean\n"
        '    ctx.company = "' + SAMPLE_COMPANY + '"\n'
        '    ctx.case_type = "new"\n'
        '    ctx.industry_name = "菓子製造"\n'
        "    roundNo = 1\n"
        '    qualityMode = "standard"\n'
        '    dossierTier = "B"\n'
        "    ReadCaseCtx = True\n"
        "End Function\n"
    ),
    "modConfig": (
        "Option Explicit\n"
        "Public Function GetStr(ByVal k As String, Optional ByVal dflt As String) As String\n"
        '    If k = "data_dir" Then\n'
        "        GetStr = StubBag.DataDir()\n"
        '    ElseIf k = "app_version" Then\n'
        '        GetStr = "' + SAMPLE_APP_VERSION + '"\n'
        "    Else\n"
        "        GetStr = dflt\n"
        "    End If\n"
        "End Function\n"
        "Public Function GetLong(ByVal k As String, Optional ByVal dflt As Long) As Long\n"
        "    GetLong = dflt\n"
        "End Function\n"
        "Public Function GetBool(ByVal k As String, Optional ByVal dflt As Boolean) As Boolean\n"
        "    GetBool = dflt\n"
        "End Function\n"
    ),
    "modLog": (
        "Option Explicit\n"
        "Public Sub LogUsage(ByVal ev As String, Optional ByVal caseId As String, "
        "Optional ByVal detail As String)\n"
        '    StubBag.AddLog "USAGE|" & ev & "|" & detail\n'
        "End Sub\n"
        "Public Sub LogError(ByVal code As String, Optional ByVal src As String, "
        "Optional ByVal detail As String, Optional ByVal n As Long)\n"
        '    StubBag.AddLog "ERROR|" & code & "|" & detail\n'
        "End Sub\n"
        "Public Sub LogRun(ByVal a As String, Optional ByVal b As String, "
        "Optional ByVal c As String)\n"
        "End Sub\n"
    ),
    "StubBag": (
        "Option Explicit\n"
        "Private mS5 As String\n"
        "Private mS2 As String\n"
        "Private mDir As String\n"
        "Private mLog As String\n"
        "Public Sub SetUp(ByVal s5 As String, ByVal s2 As String, ByVal d As String)\n"
        "    mS5 = s5\n"
        "    mS2 = s2\n"
        "    mDir = d\n"
        '    mLog = ""\n'
        "End Sub\n"
        "Public Function S5() As String\n"
        "    S5 = mS5\n"
        "End Function\n"
        "Public Function S2() As String\n"
        "    S2 = mS2\n"
        "End Function\n"
        "Public Function DataDir() As String\n"
        "    DataDir = mDir\n"
        "End Function\n"
        "Public Sub AddLog(ByVal t As String)\n"
        "    mLog = mLog & t & Chr(10)\n"
        "End Sub\n"
        "Public Function LogText() As String\n"
        "    LogText = mLog\n"
        "End Function\n"
    ),
}


def review_driver(out_dir: str, result_url: str) -> str:
    """確認者名の表を1件ずつ**製品コードへ実際に通す**ドライバ。"""
    rows = ""
    for label, expr in BLANK_INPUTS:
        rows += f'    outT = outT & Probe("BLANK", "{label}", {expr}, s5, s2)\n'
    for label, expr in VISIBLE_INPUTS:
        rows += f'    outT = outT & Probe("VISIBLE", "{label}", {expr}, s5, s2)\n'
    return (
        "Option Explicit\n"
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
        "' 1件 = 提案書(実生成)とレポート(確認者名の正規化)の**両方**を同じ入力で測る。\n"
        "Function Probe(ByVal kindText As String, ByVal label As String, _\n"
        "               ByVal reviewer As String, ByVal s5 As String, _\n"
        "               ByVal s2 As String) As String\n"
        '    StubBag.SetUp s5, s2, "' + out_dir + '"\n'
        "    Dim p As String, r As String, rv As String\n"
        '    r = modExportProposal.GenerateProposalHtml("C-0001", p, reviewer)\n'
        "    rv = modExportHtml.ReviewerOf(reviewer)\n"
        '    Probe = kindText & Chr(9) & label & Chr(9) & "reason=" & r & _\n'
        '            Chr(9) & "path=" & p & Chr(9) & "reviewer=[" & rv & "]" & _\n'
        '            Chr(9) & "log=" & Replace(StubBag.LogText(), Chr(10), "~") & Chr(10)\n'
        "End Function\n"
        "\n"
        "Sub Main\n"
        "    Dim s5 As String, s2 As String, outT As String\n"
        '    s5 = modMockLlm.ResponseById("MK-S5")\n'
        '    s2 = modMockLlm.ResponseById("MK-S2-RNW")\n'
        '    outT = ""\n'
        + rows +
        '    WriteUtf8 "' + result_url + '", outT\n'
        "End Sub\n"
    )


def need_review_literal() -> str:
    """modExportProposal.bas の EP_NEED_REVIEW(文言の唯一の値源)を読む。"""
    src = (REPO_ROOT / "src" / "app" / "modExportProposal.bas").read_text(
        encoding="utf-8")
    m = re.search(r'Const\s+EP_NEED_REVIEW\s+As\s+String\s*=\s*"([^"]*)"', src)
    return m.group(1) if m else ""


def run_review_probe(soffice: str, work_dir: Path,
                     verbose: bool) -> tuple[list[list[str]], int] | None:
    all_modules = dict(lo.discover_modules(REPO_ROOT / "src"))
    type_blocks = lo.collect_public_type_blocks(all_modules)
    modules: dict[str, str] = {}
    missing = []
    for name in REVIEW_REAL_MODULES:
        path = all_modules.get(name)
        if path is None:
            missing.append(name)
            continue
        modules[name] = path.read_text(encoding="utf-8", errors="replace")
    if missing:
        print(f"[render_proposal] FAIL: 未実装のモジュールがあります: {', '.join(missing)}")
        return None
    modules.update(REVIEW_STUBS)

    out_dir = work_dir / "review_out"
    out_dir.mkdir(parents=True, exist_ok=True)
    result = work_dir / "review_result.txt"
    modules["ReviewMain"] = review_driver(out_dir.as_posix(),
                                          "file://" + result.as_posix())

    profile_dir = work_dir / "profile_review"
    template = lo.ensure_template_profile(soffice, verbose)
    lo.fresh_profile_copy(template, profile_dir)
    lo.write_library(profile_dir, "RpnReview", modules, type_blocks)
    lo.register_libraries(profile_dir, ["RpnReview"])
    uri = ("vnd.sun.star.script:RpnReview.ReviewMain.Main"
           "?language=Basic&location=application")
    rc, out, err = lo.run_uri(soffice, profile_dir, uri, 300)
    if not result.exists():
        print(f"[render_proposal] FAIL: 確認導線の実測を実行できません(soffice exit={rc})")
        if verbose:
            print(f"  stdout: {out.strip()}\n  stderr: {err.strip()}")
        return None
    rows = [line.split("\t") for line in
            result.read_text(encoding="utf-8").splitlines() if line.strip()]
    produced = len([p for p in out_dir.iterdir() if p.is_file()])
    return rows, produced


def check_review_path(problems_sink: list[str], soffice: str, work_dir: Path,
                      verbose: bool) -> int:
    """確認導線を実測する。戻り値=**実際に検査した項目数**(0=検問として無効)。"""
    need = need_review_literal()
    if not need:
        problems_sink.append(
            "modExportProposal.bas の EP_NEED_REVIEW を読み取れませんでした")
        return 0
    got = run_review_probe(soffice, work_dir, verbose)
    if got is None:
        problems_sink.append("確認導線の実測(LibreOffice)を実行できませんでした")
        return 0
    rows, produced = got

    checked = 0
    seen: set[tuple[str, str]] = set()
    for row in rows:
        if len(row) < 3:
            continue
        kind, label = row[0], row[1]
        fields: dict[str, str] = {}
        for cell in row[2:]:
            k, _, v = cell.partition("=")
            fields[k] = v
        reason = fields.get("reason", "")
        path = fields.get("path", "")
        reviewer = fields.get("reviewer", "")
        logText = fields.get("log", "")
        seen.add((kind, label))
        if kind == "BLANK":
            checked += 3
            if reason != need or path:
                problems_sink.append(
                    f"確認者名が[{label}]だけでも提案書が止まりません"
                    f"(reason=[{reason}] path=[{path}]。期待=[{need}]・path空)")
            if "proposal_skipped|not_reviewed" not in logText:
                problems_sink.append(
                    f"確認者名が[{label}]のとき usage_log に "
                    f"proposal_skipped|not_reviewed が残りません(log=[{logText}])")
            if reviewer != "[]":
                problems_sink.append(
                    f"確認者名が[{label}]だけでもレポートは確認済みとして扱います"
                    f"(modExportHtml.ReviewerOf=[{reviewer}]。期待=空)。"
                    "レポートと提案書で空白判定が非対称です")
        else:
            checked += 2
            if reason or not path:
                problems_sink.append(
                    f"可視文字のある確認者名[{label}]で提案書が出ません"
                    f"(reason=[{reason}] path=[{path}])。「全部断る」実装でも"
                    "空白側の検査は緑になるため、この対照が要ります")
            if reviewer == "[]":
                problems_sink.append(
                    f"可視文字のある確認者名[{label}]をレポートが未確認として扱います"
                    "(modExportHtml.ReviewerOf が空)")
    want = {("BLANK", lbl) for lbl, _ in BLANK_INPUTS}
    want |= {("VISIBLE", lbl) for lbl, _ in VISIBLE_INPUTS}
    lost = want - seen
    if lost:
        problems_sink.append(f"確認導線の実測が欠けています: {sorted(lost)}")
    checked += 1
    if produced != len(VISIBLE_INPUTS):
        problems_sink.append(
            f"出力先に残った提案書が{produced}件です(期待{len(VISIBLE_INPUTS)}件="
            "可視文字の分だけ)。空白類の確認者名で1件でも書き出されていたら、"
            "未確認の顧客提示物が出ています")
    return checked


# ==============================================================================
# 人の入力の空判定を1本に寄せる(裁定書42 §2-1 の横展開を**規則として固定**)。
# ------------------------------------------------------------------------------
# 確認者名を1本に寄せても、会社名・テーマ・本文で `LenB(Trim$(x)) = 0` が新しく
# 書かれれば同じ欠陥がまた生える(Trim$ は Chr(32) しか落とさないので、全角空白・
# NBSP・ZWSP だけの入力が「入力あり」として通る)。そこで**人の入力を指す名前**
# で空判定を書いている行を機械で数え、既知の一覧(BLANK_STYLE_BASELINE)より
# 増えたら赤にする。新しい未変換を作れなくし、残りの債務を目に見える形で置く。
#
# 名前で当たりを付ける以上、機械JSONの局所変数が同じ名前のこともある。除外は
# baseline に**理由付きで**書き、黙って消さない(docs/29 §5.3「黙って直さない」)。
# ==============================================================================

# 人の入力を指す識別子(この名前で受けたものは画面から来た文字列とみなす)。
HUMAN_INPUT_NAMES = ("company", "companyName", "reviewer", "reviewedBy",
                     "reviewerName", "theme", "themeText", "body", "bodyText",
                     "utterance", "displayName", "display_name", "caseName",
                     "answer")
BLANK_STYLE_RE = re.compile(
    r"LenB\(Trim\$\((?P<a>" + "|".join(HUMAN_INPUT_NAMES) + r")\)\)"
    r"|Trim\$\((?P<b>" + "|".join(HUMAN_INPUT_NAMES) + r")\)\s*(=|<>)\s*(\"\"|vbNullString)")

# 既知の未変換(裁定書42 §2 の担当ファイル外=司令塔へ handoff)と、名前が
# たまたま一致しただけの機械JSON。**この表を増やすには裁定が要る**。
BLANK_STYLE_BASELINE = {
    # (ファイル, 行に現れる識別子): 理由
    ("app/modCaseStore.bas", "company"):
        "handoff: 案件作成時の会社名。担当ファイル外(裁定書42 §2-4)",
    ("app/modExportHtml.bas", "body"):
        "対象外: OrNull の body は ExtractJsonBlock の戻り値(機械JSON)",
    ("app/modInboxStore.bas", "themeText"):
        "handoff: 受信箱のテーマ。担当ファイル外",
    ("app/modInboxStore.bas", "theme"):
        "handoff: 受信箱のテーマ。担当ファイル外",
    ("app/modPlayOps.bas", "themeText"):
        "handoff: 受信箱の投函。担当ファイル外",
    ("app/modPlayOps.bas", "bodyText"):
        "handoff: 受信箱の投函。担当ファイル外",
    ("app/modSparring.bas", "bodyText"):
        "handoff: スパーリングの発話。担当ファイル外",
    ("ui/modUISparring.bas", "utterance"):
        "handoff: スパーリングの発話。担当ファイル外",
    ("ui/navi/modNaviActions2.bas", "theme"):
        "handoff: 受信箱の投函(ActInboxPost)。担当ファイル外",
    ("ui/navi/modNaviActions2.bas", "body"):
        "handoff: 受信箱の投函(ActInboxPost)。担当ファイル外",
}


def check_blank_style(problems_sink: list[str]) -> int:
    """人の入力の空判定が Trim$ で書かれていないか。戻り値=走査したファイル数。"""
    src_root = REPO_ROOT / "src"
    files = sorted(src_root.rglob("*.bas"))
    found: set[tuple[str, str]] = set()
    for path in files:
        rel = path.relative_to(src_root).as_posix()
        if rel.startswith("test/"):
            continue          # テストは意図的に古い書き方を再現することがある
        for lineno, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
            # 1行に2つ以上書かれることがある(`A = 0 Or B = 0`)ので全件を見る。
            for m in BLANK_STYLE_RE.finditer(line):
                name = m.group("a") or m.group("b")
                key = (rel, name)
                found.add(key)
                if key not in BLANK_STYLE_BASELINE:
                    problems_sink.append(
                        f"{rel}:{lineno} 人の入力({name})の空判定が Trim$ です。"
                        "modUtilText.HasVisibleText へ寄せてください"
                        "(裁定書42 §2-1。全角空白・NBSP・ZWSP が素通りします)")
    stale = sorted(set(BLANK_STYLE_BASELINE) - found)
    if stale:
        problems_sink.append(
            f"BLANK_STYLE_BASELINE に、もう存在しない行が残っています: {stale}"
            "(直したら表からも消すこと)")
    return len(files)


def find_node() -> str | None:
    """DOM検査に使う node の場所(無ければ None)。"""
    return shutil.which("node") or shutil.which("nodejs")


def check_dom(html_path: Path, slides: list, todo: str, faithful: bool,
              verbose: bool) -> list[str]:
    node = find_node()
    if node is None:
        # 裁定書41 §2: **fail-open 禁止**。旧実装はここで空リストを返していたので、
        # DOMを1枚も組まないまま「全22枚の登録と描画後DOM…を確認しました」と出して
        # exit 0 になっていた(裁定書40 T-M1 の orphan_check と同型)。
        return ["node が見つからないため描画後DOMの検査を実行できません"
                "(検査を飛ばして緑にはしません。裁定書41 §2)"]
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

    # 裁定書40 S-M3: 発表者ノートの描画で例外が飛んでも、run() ごと落ちて
    # 提案書が白紙になってはいけない。壊した素材でも22枚が描かれ、壊れた枚には
    # 印(data-render-error)が立つこと。素の素材では印が1つも立たないこと
    # (条件が成立するときだけ出る / しないときは出ない、の両方向)。
    if got.get("renderErrors", 0) != 0:
        problems.append(
            f"素の mock なのに data-render-error が{got['renderErrors']}枚に"
            "立っています(描画が壊れています)")
    poisoned_ids = set(got.get("poisonedIds", []))
    for no, slug, _master, _title in slides:
        if ("sl-" + slug) not in poisoned_ids:
            problems.append(
                f"発表者ノートを壊すと sl-{slug}({no}枚目)が描かれません"
                "(notes() の例外が走査ごと止めています。20章§7・裁定書40 S-M3)")
            break
    if got.get("poisonedErrors", 0) == 0:
        problems.append(
            "発表者ノートを壊しても data-render-error が1枚も立ちません"
            "(壊れたページが「わざと保留した項目」に見えます。裁定書39 R2-12)")

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

    # 裁定書41 §2: DOM検査には node が要る。無いまま進むと「全22枚の登録と
    # 描画後DOM…を確認しました」という**偽の成功文言**で緑になるので、ここで
    # 赤にして止める(tools/notice_check.py と同じ exit 2 = 検査を実行できない)。
    if find_node() is None:
        print("[render_proposal] 結果: node が見つからないため描画後DOMの検査を"
              "実行できません(検査を飛ばして緑にはしません)")
        return 2

    # 裁定書43 §1-5: 対訳表の実効出力を同じ LO の実行で採る(全対 × 全文脈)。
    gl_rows, gl_rules = glossary_from_md(GLOSSARY.read_text(encoding="utf-8"))
    gl_inputs = glossary_inputs(gl_rows, gl_rules)

    soffice = lo.find_soffice()
    work_dir = Path(tempfile.mkdtemp(prefix="rpn_proposal_"))
    try:
        html = run_render(soffice, work_dir, args.faithful, args.verbose, gl_inputs)
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
        n_pairs = check_glossary_impl(problems)
        n_term = len(set(gl_rules["term"]))
        n_eff = check_glossary_effective(problems, work_dir / GLOSSARY_OUT_NAME)
        check_s5_required(problems)
        problems += check_dom(out_path, slides, todo, args.faithful, args.verbose)
        # 確認導線(裁定書42 §2-3)。--faithful は素材の間引きを見るモードで
        # 確認導線とは無関係なので、標準モードでだけ実測する(LOの往復が増える)。
        # 飛ばした回は要点行にも「確認しました」と書かない。
        reviewed = 0
        scanned = check_blank_style(problems)
        if not args.faithful:
            reviewed = check_review_path(problems, soffice, work_dir, args.verbose)
            if reviewed == 0 and not problems:
                problems.append(
                    "確認導線の検査を1件も実行できませんでした"
                    "(0件の検査を緑にはしません)")
        if problems:
            print("[render_proposal] NG:")
            for p in problems:
                print(f"  - {p}")
            return 1
        # 要点行は**実際に回した検査だけ**を名乗る(件数を出す。裁定書42)。
        review_text = f"人の入力の空判定を{scanned}モジュール分走査・"
        if reviewed:
            review_text += (f"確認導線{reviewed}件"
                            f"(空白類{len(BLANK_INPUTS)}通りの確認者名で提案書が"
                            "生成されず、レポート側も未確認として扱うこと・"
                            "可視文字では実際に書き出されること)・")
        print(f"[render_proposal] OK: 全{len(slides)}枚の登録と描画後DOM"
              "(発表者ノートを壊しても22枚が描かれ印が立つことを含む)、"
              "20章§5.1の19変数・§8の免責・"
              f"対訳表{n_pairs}対と15章§5.6/TabooPairs の全列一致"
              f"(終端集合{n_term}文字を1文字ずつ含む)、"
              f"SoftenTaboo の実効出力{n_eff}件に壊れの徴候が無いこと"
              "(同一文字3連続・顧客語末尾の重複・終端でない置換)と冪等、"
              f"{review_text}"
              "EP_S5_REQUIRED と15章§5.6 required の一致を確認しました。")
        return 0
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
