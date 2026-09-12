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
    python3 tools/render_report.py --reviewed 山田太郎  # 確認済みの態(R2-17)
    python3 tools/render_report.py --out <path>    # 出力先を明示する
    python3 tools/render_report.py --selftest      # 検査器の回帰網だけ(LO不要)
    # exit code: 0 = 生成+検査OK / 1 = 生成できたが検査NG / 2 = 生成できず・自己テスト失敗

検査(DoD):
    (1) 先頭付近に `<meta charset="utf-8"` がある(18章§5.3(3))
    (2) 18章§3の全18セクション(v1.3で SEC-18 talk を追加)が登録表にある
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
    (8) 免責の3項分岐(18章§3.5・裁定書37 B-06)の**両態**。既定(確認者名なし)
        では表紙チップ「確認前」と `<noscript>`「（AI生成・担当者確認前）」、
        `--reviewed <名前>` では表紙チップ「確認済 <名前>」・SEC-15・
        `<noscript>` の確認済み文・`meta.reviewed_by` / `meta.reviewed_at`。
        W15 Round2 R2-17 まで、サンプル生成は `reviewedBy=""` 固定で
        **確認済みの2態が回帰網に1度も載っていなかった**。
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
REVIEWED_OUT = REPO_ROOT / "dist" / "サンプルレポート_確認済.html"

# LOへ注入するモジュール(存在するものだけ)。純組立関数が実際に辿るものと、
# その入力になる mock だけに絞る。ここに無いモジュールへの参照はLO Basicでは
# 実行時解決なので、呼ばれない限りコンパイルも実行も通る(run_lo_tests 技術メモ4)。
RENDER_MODULES = [
    "modUtil", "modUtilText", "modJsonLite", "modLog",
    "modHtmlTheme",
    "modHtmlTemplate1", "modHtmlTemplate2", "modHtmlTemplate3",
    "modHtmlTemplate4", "modHtmlTemplate5", "modHtmlTemplate6",
    "modHtmlTemplate7", "modHtmlTemplate8",
    "modExportHtml",
    # W14(裁定書37 B-03): BuildMetaJson が meta.ground_unmatched の組立で
    #   modGround.NoteJsonArray を**実際に呼ぶ**ので必須(呼ばれるものは載せる)。
    "modGround",
    # modMockLlm3 は W7(T-55)で新設した MK-S2-NEW / MK-S2-RNW の本体。
    #   modMockLlm.ResponseById が両IDでここへ委譲するので、載せ忘れると
    #   Basicライブラリに関数が無く soffice がダイアログで止まる
    #   (=exit 124 のタイムアウトになり「HTMLが生成されませんでした」に化ける)。
    "modMockLlm", "modMockLlm2", "modMockLlm3",
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
    ("SEC-18", "talk", "renderTalk"),
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
# --reviewed のときに meta.reviewed_at へ入れる固定値(出力を決定的にするため)。
SAMPLE_REVIEWED_AT = "2026/09/01 15:30:10"


def vba_str(value: str) -> str:
    """VBA の文字列リテラルへ埋め込む(二重引用符を重ねる)。"""
    return value.replace('"', '""')


def html_safe(value: str) -> str:
    """`modUtilText.HtmlSafe`(18章§5.3 NFR-S7③)と同じ置換・同じ順序。

    HTML**ソース**と突き合わせる期待値はこれを通す(W15 Round2 T-m1)。
    通さないと、確認者名に `"` `<` `&` `'` `>` が1文字でも入った時点で
    「生成物は正しいのに検査だけが赤くなる」偽陽性になる。DOM の textContent
    と突き合わせる側(check_dom)は**素のまま**が正しいので通さない。
    順序は製品コードと同じ: & を最初に置かないと二重エスケープになる。
    """
    t = value.replace("&", "&amp;")
    t = t.replace("<", "&lt;")
    t = t.replace(">", "&gt;")
    t = t.replace('"', "&quot;")
    t = t.replace("'", "&#39;")
    return t


# 18章§3.5 の3項分岐のうち「確認済み」の1文。**確認者名は素のまま**受け取る。
# 使い分け(T-m1):
#   DOM(textContent)と照合 -> disclaimer_reviewed_line(name)
#   HTMLソース(<noscript>)と照合 -> disclaimer_reviewed_line(html_safe(name))
DISCLAIMER_REVIEWED_HEAD = "本資料はAI支援により作成した骨子を担当者が確認・編集したものです"
DISCLAIMER_UNREVIEWED = "（AI生成・担当者確認前）"


def disclaimer_reviewed_line(name_as_rendered: str) -> str:
    return (DISCLAIMER_REVIEWED_HEAD + "（確認: " + name_as_rendered + " / "
            + SAMPLE_REVIEWED_AT + "）。")


def basic_driver(out_url: str, theme: str, faithful: bool,
                 reviewed_by: str = "") -> str:
    """LO Basic のドライバ(RenderMainモジュール)。素材合成はここだけで行う。"""
    round_no = "1" if faithful else "2"
    reviewed_at = SAMPLE_REVIEWED_AT if reviewed_by else ""
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
        f'        "proposal", "{SAMPLE_GENERATED_AT}", "{SAMPLE_APP_VERSION}", "{theme}", _\n'
        # 末尾は warnText / reviewedBy / reviewedAt / groundNote / s1WarnNote。
        # s1_warn はサンプルでは空にする(実在しない警告をサンプルに焼かない。
        # 出典表そのものは MK-S1-RNW の sources 3件で描かれる。裁定書38 班A)。
        # reviewedBy/reviewedAt は --reviewed のときだけ入る(R2-17)。空のときが
        # 「AI生成・担当者確認前」の態、入れたときが「確認済み」の態で、18章§3.5
        # の3項分岐の**両方**を回帰網に載せるためのスイッチである。
        f'        "", "{vba_str(reviewed_by)}", "{reviewed_at}", "", "")\n'
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
  const texts = [];
  (function walk(n) {
    if (n.attrs && n.attrs.id) { ids.push(n.attrs.id); }
    if (n.attrs && n.attrs.href) { hrefs.push(n.attrs.href); }
    if (n._text) { texts.push(n._text); }
    for (const c of n.childNodes) { walk(c); }
  })(root);
  return { ids: ids, hrefs: hrefs, texts: texts };
}

const passA = runPass(null);
const passB = runPass(function (D) { if (D && D.meta) { D.meta.round_no = 1; } });
// C = 18章§3.8/§3.9 の回帰(裁定書25 S1/S2。17章 T-56 のDoD④)。
//   ・新規案件の形(current_coverage=0件・gaps=2件)にしても SEC-08 が描かれる
//     (旧規定「両方0件なら非表示」の撤回。素材そのままだと現契約が入っている
//      ので、この分岐は落として初めて実測できる)
//   ・talk_script を与えると SEC-18 が本文と目次の両方に出る
const passC = runPass(function (D) {
  if (!D) { return; }
  if (D.s1) { D.s1.current_coverage = []; }
  if (D.s2) {
    D.s2.gaps = [
      { gap_no: 1, gap_type: 'uninsured', target: '生産物賠償', description: 'テスト',
        risk_evidence: 'リスク側', coverage_evidence: '該当契約なし' },
      { gap_no: 2, gap_type: 'uninsured', target: '休業損失', description: 'テスト',
        risk_evidence: 'リスク側', coverage_evidence: '該当契約なし' }
    ];
  }
  if (D.s3) {
    D.s3.talk_script = {
      opening: '御社の利益を止めないための話をさせてください。',
      flow: ['いまの事業の姿を確かめます。', '止まると困る所を並べます。',
             'いまの保険で足りるかを見ます。', '足りない所の埋め方を出します。'],
      closing: '次回までに現契約の写しをご用意ください。',
      taboo: ['前任者の担当時期の話には触れない。']
    };
  }
});
// D = 与えないときに SEC-18 が本文からも目次からも消えること(§3の空のときの挙動)。
const passD = runPass(function (D) {
  if (D && D.s3) { delete D.s3.talk_script; }
});
console.log(JSON.stringify({ A: passA, B: passB, C: passC, D: passD }));
"""


def run_render(soffice: str, work_dir: Path, theme: str, faithful: bool,
               verbose: bool, reviewed_by: str = "") -> str | None:
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
    modules["RenderMain"] = basic_driver(out_url, theme, faithful, reviewed_by)

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


def parse_data_json(html: str) -> dict | None:
    """`var DATA=JSON.parse("……")` の中身を辞書にして返す(読めなければ None)。

    二重に包まれている(JSの文字列リテラル → その中身がJSON)ので、先に
    リテラルとして解いてから JSON として読む。`\\u003C` エスケープ(18章§5.3(1))
    もこの経路でそのまま解ける。
    """
    pre = 'var DATA=JSON.parse("'
    a = html.find(pre)
    if a < 0:
        return None
    a += len(pre)
    b = html.find('");', a)
    if b < 0:
        return None
    try:
        return json.loads(json.loads('"' + html[a:b] + '"'))
    except ValueError:
        return None


def check_reviewed(html: str, reviewed_by: str) -> list[str]:
    """18章§3.5・裁定書37 B-06 の**確認済みの態**をソース側で確かめる(R2-17)。

    `--reviewed` を付けない既定の実行は「AI生成・担当者確認前」の態しか通らず、
    3項分岐のうち2態(SEC-15 と `<noscript>` の確認済み文)が**回帰網に1度も
    載っていなかった**(W15 Round2 R2-17)。ここでは
      (1) DATA に `reviewed_by` / `reviewed_at` が入る
      (2) `<noscript>`(JSが動かない環境の免責)が確認済みの文へ切り替わり、
          確認前の文が**残っていない**
    を見る。表紙チップ(確認済/確認前)と SEC-15 はJSが描くので check_dom が見る。
    """
    problems = []
    noscript = ""
    m = re.search(r"<noscript>(.*?)</noscript>", html, re.S)
    if m is not None:
        noscript = m.group(1)
    else:
        problems.append("<noscript> ブロックがありません(18章§3.5)")

    meta = (parse_data_json(html) or {}).get("meta", {})
    got_by = meta.get("reviewed_by", None)
    got_at = meta.get("reviewed_at", None)

    if reviewed_by:
        if got_by != reviewed_by:
            problems.append(
                f"meta.reviewed_by が [{got_by!r}] です(--reviewed で渡した"
                f"[{reviewed_by}] が meta に入っていない)")
        if got_at != SAMPLE_REVIEWED_AT:
            problems.append(
                f"meta.reviewed_at が [{got_at!r}] です"
                f"(期待 [{SAMPLE_REVIEWED_AT}])")
        # <noscript> は HTML **ソース**なので、確認者名は HtmlSafe 後の姿で入る
        # (T-m1)。素の名前で照合すると `"` `<` `&` などで偽陽性になる。
        want = disclaimer_reviewed_line(html_safe(reviewed_by))
        if want not in noscript:
            problems.append(
                f"<noscript> の免責が確認済みの文になっていません(期待: [{want}])")
        if DISCLAIMER_UNREVIEWED in noscript:
            problems.append(
                "確認者名を渡したのに <noscript> に「（AI生成・担当者確認前）」が"
                "残っています(18章§3.5 の3項分岐が切り替わっていない)")
    else:
        if got_by != "":
            problems.append('確認者名を渡していないのに meta.reviewed_by が'
                            f"[{got_by!r}] です")
        if DISCLAIMER_UNREVIEWED not in noscript:
            problems.append(
                "確認者名を渡していないのに <noscript> が「（AI生成・担当者確認前）」"
                "になっていません(18章§3.5)")
        if DISCLAIMER_REVIEWED_HEAD in noscript:
            problems.append(
                "確認者名が空なのに <noscript> が「担当者が確認・編集した」と"
                "名乗っています(裁定書37 B-06)")
    return problems


def check_dom(html_path: Path, verbose: bool, faithful: bool,
              reviewed_by: str = "") -> list[str]:
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
        # SEC-18 は素材(mock MK-S3)が talk_script を持つかどうかで出方が変わる。
        # パスAで固定の期待値を置くと、mock 側の改訂でこの検問が意味を失うので、
        # SEC-18 の回帰は**パスC/D**(与えたら出る / 与えなければ消える)が持つ。
        if sec_id == "SEC-18":
            continue
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

    # --- 18章§3.8 SEC-08(パスC): 新規案件の形でもリスク単位の主表を描く ---
    ids_c = set(got["C"]["ids"])
    hrefs_c = set(got["C"]["hrefs"])
    if "sec-coverage" not in ids_c:
        problems.append(
            "新規案件の形(current_coverage=0件・gaps=2件)で sec-coverage が"
            "描かれていません(18章§3.8 v1.3。旧「両方0件なら非表示」は撤回)")
    # --- 18章§3.9 SEC-18(パスC/D): 与えたら本文と目次の両方に出る / 消える ---
    if "sec-talk" not in ids_c:
        problems.append("talk_script を与えたのに sec-talk が描かれていません(18章§3.9)")
    if "#sec-talk" not in hrefs_c:
        problems.append("talk_script を与えたのに目次に #sec-talk がありません(18章§3.6)")
    ids_d = set(got["D"]["ids"])
    hrefs_d = set(got["D"]["hrefs"])
    if "sec-talk" in ids_d:
        problems.append(
            "talk_script が無いのに sec-talk が描かれています"
            "(18章§3 SEC-18「talk_script が無いならセクションごと非表示」)")
    if "#sec-talk" in hrefs_d:
        problems.append(
            "talk_script が無いのに目次に #sec-talk が残っています"
            "(18章§3「非表示のセクションは目次からも同時に落とす」)")
    if "sec-story" not in ids_d:
        problems.append("パスD で sec-story まで消えています(DOM検査が空振り)")

    # --- 裁定書37 B-06(パスA): 表紙チップと SEC-15 の確認済み/確認前(R2-17) ---
    texts_a = got["A"].get("texts", [])
    if reviewed_by:
        want_chip = "確認済 " + reviewed_by
        if want_chip not in texts_a:
            problems.append(
                f"表紙のチップが「{want_chip}」になっていません"
                f"(裁定書37 B-06。--reviewed の確認済みの態)")
        if "確認前" in texts_a:
            problems.append("確認者名を渡したのに表紙に「確認前」のチップが出ています")
        # DOM の textContent はエスケープが解けた**素の**姿なので html_safe は
        # 通さない(<noscript> はソースなので通す。T-m1 の非対称)。
        want_disc = disclaimer_reviewed_line(reviewed_by)
        if want_disc not in texts_a:
            problems.append(
                f"SEC-15 の1行目が確認済みの文になっていません(期待: [{want_disc}])")
    else:
        if "確認前" not in texts_a:
            problems.append(
                "表紙のチップに「確認前」がありません(裁定書37 B-06。確認前の態)")
        if any(t.startswith("確認済 ") for t in texts_a):
            problems.append("確認者名が空なのに表紙に「確認済」のチップが出ています")

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


def _numbered_literals(body: str) -> list[str]:
    """節の本文から `1. \`……\`` の形の固定文を拾う。"""
    out = []
    for line in body.splitlines():
        mm = re.match(r"^\d+\. `(.+?)`", line.strip())
        if mm:
            out.append(mm.group(1))
    return out


def parse_growth_note(spec: str) -> list[str]:
    """§3.7「節の先頭に次の1文を逐語で置く」の固定文をバッククォートから取る。

    切り出しの終端は **`"### 3.8"`**(v1.3)。`"## 4. "` までにすると §3.9 の
    固定文(SEC-18 の1文)まで拾ってしまい、SEC-18 の漂流を「§3.7 SEC-17 の
    免責固定文が入っていない」と**誤ってラベルする**(実際にW7で発生した)。
    """
    return _numbered_literals(spec.split("### 3.7 ")[1].split("### 3.8")[0])


def parse_talk_note(spec: str) -> list[str]:
    """§3.9「節の先頭に次の1文を逐語で置く」の固定文(SEC-18)を取る。"""
    return _numbered_literals(spec.split("### 3.9 ")[1].split("## 4. ")[0])


def parse_sec08_notes(spec: str) -> list[str]:
    """§3.8 の注記(新規案件の1行・確度の見立ての1行)を逐語で取る。"""
    body = spec.split("### 3.8 ")[1].split("### 3.9")[0]
    out = []
    for pat in (r"`(新規案件のため現契約なし。[^`]+)`",
                r"「(確認前の見立てを含みます)」"):
        mm = re.search(pat, body)
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

    talk = parse_talk_note(spec)
    if len(talk) != 1:
        problems.append(f"18章§3.9 の固定文が1本読めません({len(talk)}本)")
    for line in talk:
        if _literal_missing(html, line):
            problems.append(
                f"18章§3.9 SEC-18 の固定文が逐語で入っていません: [{line}]"
                f"(前後に語を足していないか。1本のJSリテラルとして書くこと)")

    notes08 = parse_sec08_notes(spec)
    if len(notes08) != 2:
        problems.append(f"18章§3.8 の注記が2本読めません({len(notes08)}本)")
    for line in notes08:
        if _literal_missing(html, line):
            problems.append(
                f"18章§3.8 SEC-08 の注記が逐語で入っていません: [{line}]"
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


# ==============================================================================
# 自己テスト(W15 Round2 T-m3): 確認前/確認済みの**両方向**を LibreOffice 無しで固定する。
# ------------------------------------------------------------------------------
# R2-17 の回帰網は scratchpad の絶対パス直書きスクリプトにしか無く、worktree を
# 消せば実行不能だった。ここへ移し、`--selftest` と毎回の生成の前段で回す。
# 合成HTMLは check_reviewed が読む2箇所(<noscript> と DATA リテラル)だけを
# 製品と同じ作りで持つ(JSの文字列リテラル → その中身がJSON の二重包み)。
# ==============================================================================
def _sample_html(reviewed_by: str, escape_noscript: bool = True) -> str:
    """確認前/確認済みの最小HTMLを合成する。

    escape_noscript=False は「製品が HtmlSafe を掛け忘れた」形(負例)。
    """
    if reviewed_by:
        shown = html_safe(reviewed_by) if escape_noscript else reviewed_by
        disc = disclaimer_reviewed_line(shown)
        at = SAMPLE_REVIEWED_AT
    else:
        disc = "本資料はAIが作成した骨子です" + DISCLAIMER_UNREVIEWED + "。"
        at = ""
    meta = {"meta": {"reviewed_by": reviewed_by, "reviewed_at": at}}
    inner = json.dumps(json.dumps(meta, ensure_ascii=False))  # 二重包み
    return ("<noscript><p>" + disc + "</p></noscript>\n"
            "<script>var DATA=JSON.parse(" + inner + ");</script>\n")


def self_test() -> bool:
    cases: list[tuple[str, bool]] = []
    name = "山田 太郎"
    plain = _sample_html("")
    reviewed = _sample_html(name)

    # HtmlSafe の写し(製品 modUtilText.HtmlSafe と同じ置換・同じ順序)。
    cases.append(("HtmlSafe 5文字",
                  html_safe("&<>\"'") == "&amp;&lt;&gt;&quot;&#39;"))
    cases.append(("HtmlSafe は & を最初に置く(二重エスケープしない順)",
                  html_safe("&amp;") == "&amp;amp;"))
    cases.append(("HtmlSafe 無害な名前は素通し", html_safe(name) == name))

    # 態の照合(正例・負例の両方向)。
    cases.append(("確認前を確認前として検査 → 緑", check_reviewed(plain, "") == []))
    cases.append(("確認済みを確認済みとして検査 → 緑",
                  check_reviewed(reviewed, name) == []))
    cases.append(("確認前を確認済みとして検査 → 赤",
                  len(check_reviewed(plain, name)) > 0))
    cases.append(("確認済みを確認前として検査 → 赤",
                  len(check_reviewed(reviewed, "")) > 0))
    cases.append(("別人の名前で検査 → 赤",
                  len(check_reviewed(reviewed, "鈴木 次郎")) > 0))
    cases.append(("<noscript> が無ければ赤",
                  len(check_reviewed('<script>var DATA=JSON.parse("{}");</script>',
                                     "")) > 0))

    # T-m1: HtmlSafe が掛かる文字を含む確認者名。
    tricky = 'A"B&C<D'
    cases.append(("HtmlSafe 対象文字を含む名前でも緑(偽陽性を出さない)",
                  check_reviewed(_sample_html(tricky), tricky) == []))
    cases.append(("製品がエスケープを掛け忘れたら赤(見逃さない)",
                  len(check_reviewed(_sample_html(tricky, escape_noscript=False),
                                     tricky)) > 0))

    # DOM 側(textContent)は**素のまま**が正しい = ソース側との非対称を固定する。
    cases.append(("ソース側の期待値は HtmlSafe 後",
                  disclaimer_reviewed_line(html_safe(tricky)) !=
                  disclaimer_reviewed_line(tricky)))
    cases.append(("DOM側の期待値は素の名前を含む",
                  tricky in disclaimer_reviewed_line(tricky)))

    # DATA リテラルの読み出し(二重包みを解けているか)。
    got = parse_data_json(reviewed) or {}
    cases.append(("meta.reviewed_by を読める",
                  got.get("meta", {}).get("reviewed_by") == name))
    cases.append(("meta.reviewed_at を読める",
                  got.get("meta", {}).get("reviewed_at") == SAMPLE_REVIEWED_AT))
    cases.append(("DATA が無ければ None", parse_data_json("<p>x</p>") is None))

    bad = [n for n, ok in cases if not ok]
    for n in bad:
        print("  自己テスト NG: %s" % n)
    print("  自己テスト: %d/%d" % (len(cases) - len(bad), len(cases)))
    return not bad


def main() -> int:
    ap = argparse.ArgumentParser(description="サンプルHTMLレポートを生成する(T-33/T-35)")
    ap.add_argument("--theme", default="standard", help="18章§5.2のテーマ名(standard / mono)")
    ap.add_argument("--out", default=None, help="出力先(既定 dist/サンプルレポート.html)")
    ap.add_argument("--faithful", action="store_true",
                    help="サンプル素材の合成を行わない(素のmock・round 1)")
    ap.add_argument("--reviewed", default="", metavar="確認者名",
                    help="確認者名を入れて「確認済み」の態で出す(18章§3.5の"
                         "3項分岐のうち残り2態。既定 dist/サンプルレポート_確認済.html)")
    ap.add_argument("--selftest", action="store_true",
                    help="検査器の自己テスト(回帰網)だけを回す(LibreOffice 不要)")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args()

    if args.selftest:
        print("[render_report] 検査器の自己テスト(R2-17 / T-m1)")
        if not self_test():
            print("[render_report] 結果: 自己テスト失敗(検査器が壊れています)")
            return 2
        print("[render_report] 結果: 自己テストOK")
        return 0

    if args.out:
        out_path = Path(args.out)
    elif args.theme != "standard":
        out_path = MONO_OUT
    elif args.reviewed:
        out_path = REVIEWED_OUT
    else:
        out_path = DEFAULT_OUT

    # 検査器そのものの回帰網。毎回のゲートで回す(骨抜き防止。T-m3)。
    if not self_test():
        print("[render_report] 結果: 自己テスト失敗(検査器が壊れています)")
        return 2

    soffice = lo.find_soffice()
    work_dir = Path(tempfile.mkdtemp(prefix="rpn_render_"))
    try:
        html = run_render(soffice, work_dir, args.theme, args.faithful,
                          args.verbose, args.reviewed)
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
        problems += check_reviewed(html, args.reviewed)
        problems += check_dom(out_path, args.verbose, args.faithful, args.reviewed)
        if problems:
            print("[render_report] NG:")
            for p in problems:
                print(f"  - {p}")
            return 1
        shown = len(SECTIONS) - (1 if args.faithful else 0)
        state = f"確認済み({args.reviewed})" if args.reviewed else "担当者確認前"
        print(f"[render_report] OK: <meta charset=\"utf-8\"> と全{len(SECTIONS)}"
              f"セクションの登録・描画後DOM{shown}本・免責の態={state} を"
              f"確認しました。")
        return 0
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
