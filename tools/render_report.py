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
    (2) 18章§3の全17セクション(v1.2で SEC-17 growth を追加)が登録表にある
        (id と slug と描画関数名の3点)
    (3) node があれば、最小DOMスタブでページ内スクリプトを実際に走らせ、
        `sec-<slug>` の要素がすべて生成されることと、上部ナビ・目次の
        アンカーが表示対象と1対1であること(11章§3.8.3)まで確認する。
        セクションの実体はブラウザ側のJSが作る(18章§4.1)ので、HTMLソースを
        grep しても cover 以外のアンカーは出てこない。ソース検査(2)だけでは
        「登録したが描けない」を見逃すため、可能なら(3)まで行う。
    (4) innerHTML / insertAdjacentHTML / document.write / outerHTML= が
        1つも出現しない(18章§4.1・17章 T-46 の出荷前検問と同じ観点)
    (5) DATAの文字列リテラル内に**生の `<` が1文字も無い**(18章§5.3(1) v1.1)。
        `</` だけを逃がす旧規約では `<!--<script>` を含むLLM出力でHTMLトークナイザが
        script data double escaped 状態へ入り、ページが白紙化する。
    (6) 18章の**固定文と見出しの逐語照合**(§3の見出し16件・§3.5の免責4行・
        §3.4のヒアリング2文・SEC-08注記・SEC-09/SEC-12のnote)。期待値は18章
        Markdownからパースする(ツール側に写経しない)。
    (7) DOMスタブのパスB: `meta.round_no` を1に落とすと SEC-16 が本文からも
        目次からも消える(18章§3。この規定の回帰網はここだけ)。
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

DEFAULT_OUT = REPO_ROOT / "dist" / "サンプルレポート.html"
MONO_OUT = REPO_ROOT / "dist" / "サンプルレポート_mono.html"

# LOへ注入するモジュール(存在するものだけ)。純組立関数が実際に辿るものと、
# その入力になる mock だけに絞る。ここに無いモジュールへの参照はLO Basicでは
# 実行時解決なので、呼ばれない限りコンパイルも実行も通る(run_lo_tests 技術メモ4)。
RENDER_MODULES = [
    "modUtil", "modUtilText", "modJsonLite", "modLog",
    "modHtmlTheme",
    "modHtmlTemplate1", "modHtmlTemplate2", "modHtmlTemplate3",
    "modHtmlTemplate4", "modHtmlTemplate5", "modHtmlTemplate6",
    "modHtmlTemplate7",
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
    ("SEC-16", "round-update", "renderRoundUpdate"),
    ("SEC-07", "risks", "renderRisks"),
    ("SEC-08", "coverage", "renderCoverage"),
    ("SEC-11", "prevent", "renderPrevent"),
    ("SEC-12", "limit", "renderLimit"),
    ("SEC-09", "newrisk", "renderNewRisk"),
    ("SEC-17", "growth", "renderGrowth"),
    ("SEC-10", "story", "renderStory"),
    ("SEC-13", "hearing", "renderHearing"),
    ("SEC-14", "source", "renderSource"),
    ("SEC-15", "disclaimer", "renderDisclaimer"),
]


# 18章§4.1 が構造的に排除している書き込み口。1つでも出たら出荷しない。
BANNED_JS = ["innerHTML", "insertAdjacentHTML", "document.write", "outerHTML"]

# 18章§5.1「テーマが定義してよいCSS変数の閉じた一覧」(v1.2で39個)。テーマは
# **過不足なく**これだけを定義し、共通CSSはこれ以外の16進色を持たない
# (唯一の例外は #fff。rgba() の半透明は§5.1の規約で例外)。
THEME_VARS = [
    "--page-width", "--page-pad", "--font-sans", "--font-serif", "--font-size",
    "--line-height",
    "--bg", "--paper", "--ink", "--sub", "--mist", "--line",
    "--brand", "--brand2", "--accent", "--navy", "--kaki", "--matsu", "--deep",
    "--soft-brand", "--soft-red", "--soft-amber", "--soft-green", "--soft-blue",
    "--soft-purple", "--shadow",
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
//
// 2パスで走らせる:
//   A = DATA をそのまま流す(通常の描画)
//   B = DATA.meta.round_no を 1 に落として流す(18章§3 SEC-16 の
//       「round_no が2未満ならセクションごと非表示。目次からも落とす」の実測)
// Bを別に持つのは、A の素材が round_no=2 のときだけ SEC-16 の分岐に意味が
// 生じるため。--faithful の素材(round_no=1・status全件proposed)では後段の
// 「3つのstatusが0件なら return」ガードが先に効いてしまい、round_no 側の
// ガードを消しても誰も気づけない(W3検証 MAJOR1 の未検出変異がこれ)。
const fs = require('fs');
const html = fs.readFileSync(process.argv[2], 'utf8');

let byId = Object.create(null);
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

const scripts = [];
const re = /<script>([\s\S]*?)<\/script>/g;
let m;
while ((m = re.exec(html)) !== null) { scripts.push(m[1]); }
if (scripts.length !== 2) { console.error('SCRIPT_BLOCKS=' + scripts.length); process.exit(3); }

// 1ページぶんのDOMを作って scripts を流し、現れたidと目次のhrefを集める。
function runPass(mutate) {
  byId = Object.create(null);
  function mk(tag, id) { const n = new El(tag); if (id) { n.setAttribute('id', id); } return n; }
  const root = mk('body', null);
  // 静的HTML側(BodyShellHtml。18章§3.6 v1.2)にある id を先に用意する。
  // 上部ナビ(toc)はヒーローより前、本文先頭の目次(tocprint)と表紙(sec-cover)は
  // <main id="doc"> の中にある。
  root.appendChild(mk('nav', 'toc'));
  root.appendChild(mk('div', 'warnbox'));
  root.appendChild(mk('button', 'btnPrint'));
  const doc = mk('main', 'doc');
  root.appendChild(doc);
  doc.appendChild(mk('nav', 'tocprint'));
  doc.appendChild(mk('section', 'sec-cover'));
  global.document = {
    createElement: (t) => new El(t),
    getElementById: (id) => (byId[id] || null)
  };
  global.window = { print: function () {} };

  (0, eval)(scripts[0]);            // var DATA=JSON.parse("...")
  if (mutate) { mutate(global.DATA); }
  (0, eval)(scripts[1]);            // ランタイム(登録配列の走査・描画・目次)

  const ids = [];
  const hrefs = [];
  (function walk(n) {
    if (n.attrs && n.attrs.id) { ids.push(n.attrs.id); }
    if (n.attrs && n.attrs.href) { hrefs.push(n.attrs.href); }
    for (const c of n.childNodes) { walk(c); }
  })(root);
  return { ids: ids, hrefs: hrefs };
}

const passA = runPass(null);
const passB = runPass(function (D) { if (D && D.meta) { D.meta.round_no = 1; } });
console.log(JSON.stringify({ A: passA, B: passB }));
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
        problems.append(f"テーマCSSの39変数が過不足です(不足={missing} 余分={extra})")
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

    パスA(素材そのまま): 表示されるべきセクションが全部描かれているか。
      --faithful(素材合成なし=初回ラウンド)のときは SEC-16 が18章§3の規定
      どおり非表示になるのが**正しい**ので、DOMに無いことを期待値にする。
    パスB(meta.round_no を 1 に落とす): 18章§3 SEC-16 の
      「`meta.round_no` が2未満（初回ラウンド）なら**セクションごと非表示**
      （目次からも落とす）」を、status が new/confirmed/rejected で埋まった
      素材のまま round_no だけ下げて実測する。**この規定を守る回帰網はここ
      だけ**なので、本文(sec-round-update)と目次(#sec-round-update)の両方が
      消えることを見る。
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
    try:
        got = json.loads(proc.stdout.strip().splitlines()[-1])
    except (ValueError, IndexError):
        return [f"DOM検査の出力を読めませんでした: {proc.stdout.strip()[:200]}"]

    ids_a = set(got["A"]["ids"])
    ids_b = set(got["B"]["ids"])
    hrefs_b = set(got["B"]["hrefs"])
    if verbose:
        print(f"[render_report] パスA のid: {sorted(ids_a)}")
        print(f"[render_report] パスB(round_no=1)のid: {sorted(ids_b)}")
    problems = []
    for sec_id, slug, _fn in SECTIONS:
        hidden_expected = faithful and sec_id == "SEC-16"
        present = ("sec-" + slug) in ids_a
        if hidden_expected and present:
            problems.append(f"初回ラウンドなのに sec-{slug} が描かれています({sec_id}・18章§3)")
        elif not hidden_expected and not present:
            problems.append(f"描画後のDOMに sec-{slug} がありません({sec_id})")

    # --- 18章§3 SEC-16 の round_no ガード(パスB) ---
    if "sec-round-update" in ids_b:
        problems.append(
            "meta.round_no=1 なのに sec-round-update が描かれています"
            "(18章§3 SEC-16「round_no が2未満ならセクションごと非表示」)")
    if "#sec-round-update" in hrefs_b:
        problems.append(
            "meta.round_no=1 なのに目次に #sec-round-update が残っています"
            "(18章§3「非表示のセクションは目次からも同時に落とす」)")
    # パスBが「SEC-16以外まで消えた」状態でないこと(ガードの効きすぎ・空振り防止)。
    if "sec-exec" not in ids_b:
        problems.append("パスB(round_no=1)で sec-exec まで消えています(DOM検査が空振り)")

    # --- 11章§3.8.3(1)(2): 上部ナビのアンカー本数 = 表示対象数、id は slug と1対1 ---
    # 見出しを持たない SEC-01 cover はナビにも目次にも並べない(18章§3.6)ので、
    # 期待本数は「描かれたセクション - cover」。上部ナビと印刷用目次は同じ一覧から
    # 作るので、アンカーは1セクションにつき2本(帯 + 目次)出る。
    shown_slugs = [slug for _sid, slug, _fn in SECTIONS
                   if ("sec-" + slug) in ids_a and slug != "cover"]
    want = sorted("#sec-" + s for s in shown_slugs)
    got_nav = sorted(h for h in got["A"]["hrefs"] if h.startswith("#sec-"))
    if sorted(set(got_nav)) != want:
        problems.append(
            f"上部ナビ・目次のアンカーが表示対象と一致しません"
            f"(不足={sorted(set(want) - set(got_nav))} 余分={sorted(set(got_nav) - set(want))}"
            f"。11章§3.8.3(1)(2)・18章§3.6)")
    if len(got_nav) != len(want) * 2:
        problems.append(
            f"アンカーの本数が{len(got_nav)}本です(上部ナビ{len(want)}本+印刷用目次"
            f"{len(want)}本の計{len(want) * 2}本を期待。18章§3.6)")
    return problems


# ==============================================================================
# 18章の固定文・見出しの逐語照合(漂流の恒久検問)
# ------------------------------------------------------------------------------
# validate_check.py が15章§11のエラー文テンプレに対して行っている「一字一句の
# 近傍照合」を、18章の固定文にも掛ける。W3では SEC-08 注記への「です。」付加と
# §3.4 超過1行の半角括弧という2件の漂流が、目視でしか見つからなかった。
# **期待値は18章Markdownからパースする**(ツール側に写経しない=二重管理にしない)。
# ==============================================================================
SPEC18 = REPO_ROOT / "docs" / "spec" / "18_HTMLレポートテンプレート仕様.md"


def spec18_text() -> str:
    return SPEC18.read_text(encoding="utf-8")


def parse_sec_headings(spec: str) -> list[tuple[str, str, str]]:
    """§3の表から (SEC-nn, slug, 見出し（既定）) を拾う。見出しが正の値源。"""
    out = []
    for line in spec.splitlines():
        if not line.startswith("| SEC-"):
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) < 3:
            continue
        sec_id, slug, title = cells[0], cells[1], cells[2]
        if title.startswith("（表紙"):      # SEC-01 は「見出しなし」
            title = ""
        out.append((sec_id, slug, title))
    return out


def parse_disclaimer(spec: str) -> list[str]:
    """§3.5「この4行を必ず含める」の固定文をバッククォートから取る。"""
    body = spec.split("### 3.5 ")[1].split("### 3.6")[0]
    out = []
    for line in body.splitlines():
        mm = re.match(r"^\d+\. `(.+?)`", line.strip())
        if mm:
            out.append(mm.group(1))
    return out


def parse_growth_note(spec: str) -> list[str]:
    """§3.7「節の先頭に次の1文を逐語で置く」の固定文をバッククォートから取る。"""
    body = spec.split("### 3.7 ")[1].split("## 4. ")[0]
    out = []
    for line in body.splitlines():
        mm = re.match(r"^\d+\. `(.+?)`", line.strip())
        if mm:
            out.append(mm.group(1))
    return out


def parse_fixed_phrases(spec: str) -> list[tuple[str, str]]:
    """§3.4・§3のSEC-08注記・SEC-09/SEC-12のnoteを (出典, 逐語) で返す。"""
    out: list[tuple[str, str]] = []
    s34 = spec.split("### 3.4 ")[1].split("### 3.5")[0]
    for label, pat in (("§3.4 超過1行", r"「(ほか\{n\}問（[^」]+）)」"),
                       ("§3.4 不足情報の設問文", r"`(\{item\}について教えてください)`")):
        mm = re.search(pat, s34)
        if mm:
            out.append((label, mm.group(1)))
    s3 = spec.split("## 3. ")[1].split("### 3.1")[0]
    for label, pat in (
        ("§3 SEC-08 注記", r"「(新規案件のため現契約なし。[^」]+)」の注記"),
        ("§3 SEC-09 0件時の1行", r"「(現時点で特筆すべき[^」]+)」の1行"),
        ("§3 SEC-12 0件時の1行", r"3ブロックとも0件なら「(該当なし)」の1行"),
    ):
        mm = re.search(pat, s3)
        if mm:
            out.append((label, mm.group(1)))
    return out


def _literal_missing(html: str, frag: str) -> bool:
    """固定文の断片が「JSの単引用符リテラルまるごと」として現れるか。

    ただの部分文字列一致にしないのは、**末尾に語を足す漂流**（W3で実際に起きた
    SEC-08 注記への「です。」付加）が部分文字列一致では素通りするため。テンプレは
    固定文を `T(el,'p','note','……')` の形で1本のリテラルとして書くので、前後を
    引用符ごと照合すれば付加も削除も同時に捕まる。プレースホルダ（`{item}` 等）で
    割れた断片も `'……'+S(x)+'……'` の形になるため、断片ごとに同じ規則で当たる。
    """
    return ("'" + frag + "'") not in html


def check_spec18_literals(html: str) -> list[str]:
    """18章の固定文・見出しが生成HTMLに逐語で入っているか。"""
    problems = []
    spec = spec18_text()

    headings = parse_sec_headings(spec)
    if len(headings) != len(SECTIONS):
        problems.append(f"18章§3の表から{len(headings)}行しか読めません(期待{len(SECTIONS)})")
    for sec_id, slug, title in headings:
        if not title:
            continue
        row = f"{{id:'{sec_id}',slug:'{slug}',title:'{title}'"
        if row not in html.replace("\n", ""):
            problems.append(
                f"{sec_id} の見出しが18章§3の「{title}」と逐語一致しません"
                f"(登録表の title。18章§4.2)")

    disc = parse_disclaimer(spec)
    if len(disc) != 4:
        problems.append(f"18章§3.5の免責が4行読めません({len(disc)}行)")
    for line in disc:
        # 4行目は {meta.xxx} を含むテンプレなので、プレースホルダで割った
        # リテラル片をすべて照合する。
        for frag in [f for f in re.split(r"\{[^}]+\}", line) if f.strip()]:
            if _literal_missing(html, frag):
                problems.append(
                    f"18章§3.5の免責固定文が逐語で入っていません: [{frag}]"
                    f"(前後に語を足していないか。1本のJSリテラルとして書くこと)")

    for line in parse_growth_note(spec):
        if _literal_missing(html, line):
            problems.append(
                f"18章§3.7 SEC-17 の免責固定文が逐語で入っていません: [{line}]"
                f"(前後に語を足していないか。1本のJSリテラルとして書くこと)")

    for label, phrase in parse_fixed_phrases(spec):
        for frag in [f for f in re.split(r"\{[^}]+\}", phrase) if f.strip()]:
            if _literal_missing(html, frag):
                problems.append(
                    f"{label}の固定文が逐語で入っていません: [{frag}]"
                    f"(前後に語を足していないか。1本のJSリテラルとして書くこと)")
    return problems


def check_data_literal(html: str) -> list[str]:
    """18章§5.3(1) v1.1 の受入条件: DATAの文字列リテラル内に生の `<` が無いこと。

    `</` だけを逃がす旧規約では、LLM出力の1フィールドに `<!--<script>`
    (`-->` を伴わない形)が入るとHTMLトークナイザが script data double escaped
    状態へ入り、DATAブロックの正規の `</script>` が終端として働かず、ページが
    1セクションも描かれない真っ白な状態になる。`<` が1文字も無ければ
    `</script>` も `<!--` も `<script` も構造上現れない。
    """
    problems = []
    pre = 'var DATA=JSON.parse("'
    a = html.find(pre)
    if a < 0:
        return ["DATAが `var DATA=JSON.parse(\"...\")` の形で埋まっていません(18章§5.3(1))"]
    a += len(pre)
    b = html.find('");', a)
    if b < 0:
        return ["DATAの文字列リテラルが閉じていません(18章§5.3(1))"]
    lit = html[a:b]
    bad = lit.find("<")
    if bad >= 0:
        problems.append(
            f"DATAの文字列リテラル内に生の `<` があります(位置{bad}・周辺=[{lit[max(0, bad - 20):bad + 30]}])。"
            f"18章§5.3(1) v1.1 は「すべての `<` を \\u003C へ」を要求します")
    # 文書全体としても <script> の対が2組ちょうどであること(トークナイザが
    # 迷わない=白紙化しないことの構造的な裏取り)。
    n_open = len(re.findall(r"<script>", html))
    n_close = len(re.findall(r"</script>", html))
    if (n_open, n_close) != (2, 2):
        problems.append(
            f"<script>ブロックが2組ではありません(開き{n_open}/閉じ{n_close}。"
            f"DATAブロックとランタイムブロックの2本が18章§5.3(1)・§4.1の想定)")
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
        problems += check_data_literal(html)
        problems += check_spec18_literals(html)
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
