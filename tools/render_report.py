#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
render_report.py - 実物のサンプルHTMLレポートを出す(17章 T-33 / T-35 の受入確認)

================================================================================
役割:
    18章のHTMLレポートを「人がブラウザで開ける実物」として出す。lint と
    run_lo_tests が「壊れていないこと」を見るのに対し、本スクリプトは
    **見た目そのもの**を人がレビューできる状態に落とす。

    LibreOffice(soffice --headless)へ純文字列モジュール一式を読み込ませ、
    mock素材(modMockLlm.ResponseById。15章§8.1の11 ID)を入力として
    modExportHtml の**純組立関数**
        BuildMetaJson(...)  -> 18章§2の meta
        BuildReportHtml(...) -> HTML全文
    を実行し、結果を dist/サンプルレポート.html へ書き出す。
    実行機構(雛形プロファイル・.xba変換・型ブロック注入・timeout)は
    tools/run_lo_tests.py をそのまま import して流用する(二重管理しない)。

なぜファイル書出だけPython側なのか:
    製品の書出は ADODB.Stream(Charset="utf-8"・BOMあり。18章§5.3(3))で行うが、
    ADODB は Windows 専用で LibreOffice/Linux には無い。そこで本ツールは
    「組立=製品コード(VBA)/書出=ツール(Python)」に分け、Python 側で
    **同じバイト配置(UTF-8・BOMあり)**にして書く。組立ロジックは1行も
    ツール側に持たない(持った瞬間に「サンプルだけ綺麗」になるため)。

サンプル素材の合成について(**製品コードには一切入らない**):
    mock は MK-S2-RNW が gaps つき・emerging_risks 空、MK-S2-NEW が
    emerging_risks つき・gaps 空、status は全件 proposed である。
    このままだと SEC-08(カバレッジ)か SEC-09(ニューリスク)のどちらかが
    薄くなり、SEC-16(訪問で分かったこと)は 18章§3 の規定どおり非表示になる。
    見本として全セクションを1枚で見せるため、**このツールの中だけで**
      (a) MK-S2-NEW の emerging_risks を MK-S2-RNW へ移植
      (b) risk_no 3/6/8 の status を confirmed / rejected / new に変える
      (c) meta.round_no を 2 にする
    という素材合成を行う。合成は生成されたBasicドライバ側の文字列操作であり、
    src/ の製品コードにも mock 本体にもこの分岐は存在しない。
    `--faithful` を付けると合成を行わない(初回ラウンドの素のmockそのまま。
    このとき SEC-16 は仕様どおり非表示になる)。

使い方:
    python3 tools/render_report.py                 # dist/サンプルレポート.html
    python3 tools/render_report.py --theme mono    # dist/サンプルレポート_mono.html
    python3 tools/render_report.py --faithful      # 素材合成なし(round 1)
    python3 tools/render_report.py --out <path>    # 出力先を明示する
    # exit code: 0 = 生成+検査OK / 1 = 生成できたが検査NG / 2 = 生成できず

検査(DoD):
    (1) 先頭付近に `<meta charset="utf-8"` がある(18章§5.3(3))
    (2) 18章§3の全16セクションが登録表にある(id と slug と描画関数名の3点)
    (3) node があれば、最小DOMスタブでページ内スクリプトを実際に走らせ、
        `sec-<slug>` の要素が16本すべて生成されることまで確認する。
        セクションの実体はブラウザ側のJSが作る(18章§4.1)ので、HTMLソースを
        grep しても cover 以外のアンカーは出てこない。ソース検査(2)だけでは
        「登録したが描けない」を見逃すため、可能なら(3)まで行う。
    (4) innerHTML / insertAdjacentHTML / document.write / outerHTML= が
        1つも出現しない(18章§4.1・17章 T-46 の出荷前検問と同じ観点)
================================================================================
"""

from __future__ import annotations

import argparse
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

DEFAULT_OUT = REPO_ROOT / "dist" / "サンプルレポート.html"
MONO_OUT = REPO_ROOT / "dist" / "サンプルレポート_mono.html"

# LOへ注入するモジュール(存在するものだけ)。純組立関数が実際に辿るものと、
# その入力になる mock だけに絞る。ここに無いモジュールへの参照はLO Basicでは
# 実行時解決なので、呼ばれない限りコンパイルも実行も通る(run_lo_tests 技術メモ4)。
RENDER_MODULES = [
    "modUtil", "modUtilText", "modJsonLite", "modLog",
    "modHtmlTheme",
    "modHtmlTemplate1", "modHtmlTemplate2", "modHtmlTemplate3",
    "modHtmlTemplate4", "modHtmlTemplate5",
    "modExportHtml",
    "modMockLlm", "modMockLlm2",
]

# 18章§3のセクション一覧(ID / slug / 描画関数名)。この表が本ツールの検査基準。
SECTIONS = [
    ("SEC-01", "cover", "renderCover"),
    ("SEC-02", "exec", "renderExec"),
    ("SEC-03", "profile", "renderProfile"),
    ("SEC-04", "sufficiency", "renderSufficiency"),
    ("SEC-05", "riskuniv", "renderRiskUniv"),
    ("SEC-06", "riskmap", "renderRiskMap"),
    ("SEC-07", "risks", "renderRisks"),
    ("SEC-08", "coverage", "renderCoverage"),
    ("SEC-09", "newrisk", "renderNewRisk"),
    ("SEC-16", "round-update", "renderRoundUpdate"),
    ("SEC-10", "story", "renderStory"),
    ("SEC-11", "prevent", "renderPrevent"),
    ("SEC-12", "limit", "renderLimit"),
    ("SEC-13", "hearing", "renderHearing"),
    ("SEC-14", "source", "renderSource"),
    ("SEC-15", "disclaimer", "renderDisclaimer"),
]

# 18章§4.1 が構造的に排除している書き込み口。1つでも出たら出荷しない。
BANNED_JS = ["innerHTML", "insertAdjacentHTML", "document.write", "outerHTML"]

# 18章§5.1「テーマが定義してよいCSS変数の閉じた一覧」(28個)。テーマは**過不足
# なく**これだけを定義し、共通CSSはこれ以外の色を持たない(唯一の例外は #fff)。
THEME_VARS = [
    "--page-width", "--page-pad", "--font-sans", "--font-serif", "--font-size",
    "--line-height",
    "--paper", "--ink", "--sub", "--mist", "--line",
    "--ai", "--kaki", "--matsu", "--deep",
    "--warn", "--warn-line",
    "--heat-1", "--heat-2", "--heat-3", "--heat-4", "--heat-5",
    "--tr-cover", "--tr-partial", "--tr-hard",
    "--iq-ok", "--iq-partial", "--iq-missing",
]

# サンプルの meta(18章§2)。日時を固定しておくと、見た目の差分レビューで
# 「毎回変わる1行」がノイズにならない。
SAMPLE_CASE_ID = "C-20260901-001"
SAMPLE_COMPANY = "株式会社浜松スイーツファクトリー"
SAMPLE_INDUSTRY_CODE = "09"
SAMPLE_INDUSTRY_NAME = "食料品製造業"
SAMPLE_GENERATED_AT = "2026/09/01 14:07:22"
SAMPLE_APP_VERSION = "2.4.0"


def basic_driver(out_url: str, theme: str, faithful: bool) -> str:
    """LO Basic のドライバ(RenderMainモジュール)。素材合成はここだけで行う。"""
    round_no = "1" if faithful else "2"
    shaping = ""
    if not faithful:
        shaping = (
            "    Dim s2new As String\n"
            '    s2new = modMockLlm.ResponseById("MK-S2-NEW")\n'
            "    Dim emerging As String\n"
            '    emerging = SliceBetween(s2new, """emerging_risks"":", ",""open_questions""")\n'
            "    If Len(emerging) > 0 Then\n"
            '        s2 = Replace(s2, """emerging_risks"":[]", """emerging_risks"":" & emerging)\n'
            "    End If\n"
            '    s2 = SetStatusOf(s2, 3, "confirmed")\n'
            '    s2 = SetStatusOf(s2, 6, "rejected")\n'
            '    s2 = SetStatusOf(s2, 8, "new")\n'
        )
    return (
        "Option Explicit\n"
        "\n"
        "' --- サンプル素材の合成(このモジュールだけの都合。製品コードには無い) ---\n"
        "Function SliceBetween(ByVal src As String, ByVal startTag As String, _\n"
        "                      ByVal endTag As String) As String\n"
        "    Dim p1 As Long, p2 As Long\n"
        "    p1 = InStr(1, src, startTag)\n"
        "    If p1 = 0 Then Exit Function\n"
        "    p1 = p1 + Len(startTag)\n"
        "    p2 = InStr(p1, src, endTag)\n"
        "    If p2 = 0 Then Exit Function\n"
        "    SliceBetween = Mid(src, p1, p2 - p1)\n"
        "End Function\n"
        "\n"
        "Function SetStatusOf(ByVal src As String, ByVal riskNo As Long, _\n"
        "                     ByVal newStatus As String) As String\n"
        "    SetStatusOf = src\n"
        "    Dim anchorText As String\n"
        '    anchorText = """risk_no"":" & CStr(riskNo) & ","\n'
        "    Dim p As Long, q As Long\n"
        "    p = InStr(1, src, anchorText)\n"
        "    If p = 0 Then Exit Function\n"
        '    q = InStr(p, src, """status"":""proposed""")\n'
        "    If q = 0 Then Exit Function\n"
        '    SetStatusOf = Left(src, q - 1) & """status"":""" & newStatus & """" & _\n'
        '                  Mid(src, q + Len("""status"":""proposed"""))\n'
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
        "    Dim s1 As String, s2 As String, s3 As String\n"
        '    s1 = modMockLlm.ResponseById("MK-S1-RNW")\n'
        '    s2 = modMockLlm.ResponseById("MK-S2-RNW")\n'
        '    s3 = modMockLlm.ResponseById("MK-S3")\n'
        f"{shaping}"
        "    Dim metaJson As String\n"
        "    metaJson = modExportHtml.BuildMetaJson( _\n"
        f'        "{SAMPLE_CASE_ID}", "{SAMPLE_COMPANY}", "{SAMPLE_INDUSTRY_CODE}", _\n'
        f'        "{SAMPLE_INDUSTRY_NAME}", "renewal", "t2_full", "deep", {round_no}, _\n'
        f'        "proposal", "{SAMPLE_GENERATED_AT}", "{SAMPLE_APP_VERSION}", "{theme}", "")\n'
        "    Dim docText As String\n"
        f'    docText = modExportHtml.BuildReportHtml(metaJson, s1, s2, s3, "{theme}")\n'
        f'    WriteUtf8 "{out_url}", docText\n'
        "End Sub\n"
    )


DOM_STUB_JS = r"""
'use strict';
// 18章§4.1 が許す操作(createElement / textContent / setAttribute / appendChild /
// removeChild / getElementById / addEventListener)だけを持つ最小DOMスタブ。
// これで足りるということ自体が「innerHTML系を使っていない」ことの裏取りになる。
const fs = require('fs');
const html = fs.readFileSync(process.argv[2], 'utf8');

const byId = Object.create(null);
function El(tag) {
  this.tagName = tag; this.nodeType = 1; this.childNodes = [];
  this.attrs = Object.create(null); this.className = ''; this._text = '';
  this.parentNode = null;
}
Object.defineProperty(El.prototype, 'textContent', {
  get() { return this._text; },
  set(v) { this._text = String(v); this.childNodes = []; }
});
Object.defineProperty(El.prototype, 'firstChild', {
  get() { return this.childNodes.length ? this.childNodes[0] : null; }
});
El.prototype.appendChild = function (c) {
  if (c.parentNode) { c.parentNode.removeChild(c); }
  c.parentNode = this; this.childNodes.push(c); return c;
};
El.prototype.removeChild = function (c) {
  const i = this.childNodes.indexOf(c);
  if (i >= 0) { this.childNodes.splice(i, 1); }
  c.parentNode = null; return c;
};
El.prototype.setAttribute = function (k, v) {
  this.attrs[k] = String(v);
  if (k === 'id') { byId[String(v)] = this; }
};
El.prototype.addEventListener = function () {};

function mk(tag, id) { const n = new El(tag); if (id) { n.setAttribute('id', id); } return n; }
const root = mk('body', null);
const doc = mk('main', 'doc');
root.appendChild(doc);
// 静的HTML側(BodyShellHtml)にある id を先に用意する。
doc.appendChild(mk('section', 'sec-cover'));
doc.appendChild(mk('nav', 'toc'));
root.appendChild(mk('div', 'warnbox'));
root.appendChild(mk('button', 'btnPrint'));

global.document = {
  createElement: (t) => new El(t),
  getElementById: (id) => (byId[id] || null)
};
global.window = { print: function () {} };

const scripts = [];
const re = /<script>([\s\S]*?)<\/script>/g;
let m;
while ((m = re.exec(html)) !== null) { scripts.push(m[1]); }
if (scripts.length < 2) { console.error('SCRIPT_BLOCKS=' + scripts.length); process.exit(3); }
for (const s of scripts) { (0, eval)(s); }

const seen = [];
(function walk(n) {
  if (n.attrs && n.attrs.id) { seen.push(n.attrs.id); }
  for (const c of n.childNodes) { walk(c); }
})(doc);
console.log(seen.join('\n'));
"""


def run_render(soffice: str, work_dir: Path, theme: str, faithful: bool,
               verbose: bool) -> str | None:
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
        print(f"[render_report] FAIL: 未実装のモジュールがあります: {', '.join(missing)}")
        return None

    out_file = work_dir / "report.html"
    out_url = "file://" + out_file.as_posix()
    modules["RenderMain"] = basic_driver(out_url, theme, faithful)

    profile_dir = work_dir / "profile_render"
    template = lo.ensure_template_profile(soffice, verbose)
    lo.fresh_profile_copy(template, profile_dir)
    lo.write_library(profile_dir, "RpnRender", modules, type_blocks)
    lo.register_libraries(profile_dir, ["RpnRender"])

    uri = ("vnd.sun.star.script:RpnRender.RenderMain.Main"
           "?language=Basic&location=application")
    rc, out, err = lo.run_uri(soffice, profile_dir, uri, 180)
    if not out_file.exists():
        print(f"[render_report] FAIL: HTMLが生成されませんでした(soffice exit={rc})")
        if verbose:
            print(f"  stdout: {out.strip()}\n  stderr: {err.strip()}")
        return None
    return out_file.read_text(encoding="utf-8", errors="replace")


def check_source(html: str) -> list[str]:
    problems = []
    head = html[:400]
    if '<meta charset="utf-8"' not in head:
        problems.append('<head> の先頭付近に <meta charset="utf-8"> がありません(18章§5.3(3))')
    for sec_id, slug, fn in SECTIONS:
        if f"id:'{sec_id}',slug:'{slug}'" not in html:
            problems.append(f"登録表に {sec_id}(slug={slug})の行がありません(18章§4.2)")
        if f"function {fn}(" not in html:
            problems.append(f"描画関数 {fn} が見つかりません({sec_id})")
    for word in BANNED_JS:
        if word in html:
            problems.append(f"禁止された書込口 {word} が出現します(18章§4.1)")
    if "JSON.parse(" not in html:
        problems.append("DATAが JSON.parse 形式で埋まっていません(18章§5.3(1))")
    problems += check_theme_css(html)
    return problems


def check_theme_css(html: str) -> list[str]:
    """18章§5.1の機械検査(1)(2)(3)をレポート本文に対して行う。"""
    problems = []
    m = re.search(r":root\{(.*?)\n\}", html, re.S)
    if m is None:
        problems.append("テーマCSSの :root{ ... } ブロックが見つかりません(18章§5.1)")
        return problems
    body = m.group(1)
    declared = re.findall(r"(--[a-z0-9-]+)\s*:", body)
    for line in body.strip().splitlines():
        if line.strip() and not line.strip().startswith("--"):
            problems.append(f"テーマCSSに変数以外の宣言があります: {line.strip()[:60]}")
    if sorted(declared) != sorted(THEME_VARS):
        missing = sorted(set(THEME_VARS) - set(declared))
        extra = sorted(set(declared) - set(THEME_VARS))
        problems.append(f"テーマCSSの28変数が過不足です(不足={missing} 余分={extra})")
    # 共通CSS(= :root ブロック以外の <style> 本文)に生の色指定が無いこと。
    style = re.search(r"<style>(.*?)</style>", html, re.S)
    if style is not None:
        common = style.group(1).replace(m.group(0), "")
        bad = [h for h in re.findall(r"#[0-9A-Fa-f]{3,6}\b", common) if h.lower() != "#fff"]
        if bad:
            problems.append(f"共通CSSに var(--) を伴わない色指定があります: {sorted(set(bad))}")
    return problems


def check_dom(html_path: Path, verbose: bool, faithful: bool) -> list[str]:
    """node があれば最小DOMスタブでページのJSを実際に走らせて確認する。

    --faithful(素材合成なし=初回ラウンド)のときは SEC-16 が18章§3の規定
    どおり非表示になるのが**正しい**ので、DOMに無いことを期待値にする。
    """
    node = shutil.which("node") or shutil.which("nodejs")
    if node is None:
        print("[render_report] (node が無いためDOM検査はスキップしました)")
        return []
    stub = html_path.parent / "dom_stub.js"
    stub.write_text(DOM_STUB_JS, encoding="utf-8")
    proc = subprocess.run([node, str(stub), str(html_path)],
                          capture_output=True, text=True, timeout=120)
    if proc.returncode != 0:
        return [f"DOM検査でページのJSが落ちました: {proc.stderr.strip()[:400]}"]
    ids = set(l.strip() for l in proc.stdout.splitlines() if l.strip())
    if verbose:
        print(f"[render_report] DOMに現れたid: {sorted(ids)}")
    problems = []
    for sec_id, slug, _fn in SECTIONS:
        hidden_expected = faithful and sec_id == "SEC-16"
        present = ("sec-" + slug) in ids
        if hidden_expected and present:
            problems.append(f"初回ラウンドなのに sec-{slug} が描かれています({sec_id}・18章§3)")
        elif not hidden_expected and not present:
            problems.append(f"描画後のDOMに sec-{slug} がありません({sec_id})")
    return problems


def main() -> int:
    ap = argparse.ArgumentParser(description="サンプルHTMLレポートを生成する(T-33/T-35)")
    ap.add_argument("--theme", default="standard", help="18章§5.2のテーマ名(standard / mono)")
    ap.add_argument("--out", default=None, help="出力先(既定 dist/サンプルレポート.html)")
    ap.add_argument("--faithful", action="store_true",
                    help="サンプル素材の合成を行わない(素のmock・round 1)")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args()

    if args.out:
        out_path = Path(args.out)
    elif args.theme != "standard":
        out_path = MONO_OUT
    else:
        out_path = DEFAULT_OUT

    soffice = lo.find_soffice()
    work_dir = Path(tempfile.mkdtemp(prefix="rpn_render_"))
    try:
        html = run_render(soffice, work_dir, args.theme, args.faithful, args.verbose)
        if html is None:
            return 2
        out_path.parent.mkdir(parents=True, exist_ok=True)
        # 製品は ADODB.Stream(utf-8・BOMあり)で書く(18章§5.3(3))。同じバイト
        # 配置になるよう utf-8-sig で書き出す。
        out_path.write_text(html, encoding="utf-8-sig", newline="")
        print(f"[render_report] 生成: {out_path} ({len(html)}字)")

        problems = check_source(html)
        problems += check_dom(out_path, args.verbose, args.faithful)
        if problems:
            print("[render_report] NG:")
            for p in problems:
                print(f"  - {p}")
            return 1
        shown = len(SECTIONS) - (1 if args.faithful else 0)
        print(f"[render_report] OK: <meta charset=\"utf-8\"> と全{len(SECTIONS)}"
              f"セクションの登録・描画後DOM{shown}本を確認しました。")
        return 0
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
