#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""notice_check.py - 「注記の出し方」の回帰検問(裁定書40 Q-M2 / Q-m2)

================================================================================
なぜ要るか(裁定書40 §0・検証者レポートB):
    W15 Round2 Fix の独立検証で、班Q が足した2つの注記のうち**2箇所の回帰網が
    全ゲートを素通り**した。

      (1) SEC-09 の「原文未照合」(裁定書39 R1-08)
          回帰テストが生成済みJSの**文字列 grep 3本**だけだったので、
          `modHtmlTemplate4` の `if(gm['E'+(i+1)])` を `if(!gm['E'+(i+1)])` へ
          反転させても純層926本と render 系が緑のままだった。反転すると
          「原文で見つかったカード全部に『原文未照合』と出し、見つからなかった
          カードには出さない」という**真逆の表示**になり、正しい根拠を捏造
          扱いする最悪の誤表示になる。ページ内JSは VBA から走らせられないので、
          純層テストはこの分岐を原理的に踏めない。

      (2) 16章 E-02 の警告帯の配線(裁定書39 R1-07(a))
          追加テストは純関数 `modUICase.IqBannerTextOf` しか見ておらず、
          `modUICase2.DrawStep` からの呼び出しを削除しても誰も気付かなかった。
          R1-07 の欠陥そのものが「判定はあるのに画面に出ない」という**配線の
          欠落**だったので、同じ欠陥が無検出で再発しうる。

    そこで本ツールは、純層テストが原理的に届かない2点だけを受け持つ。

検査①(実DOM・両方向): SEC-09 の「原文未照合」
    LibreOffice で実物のHTMLレポートを組み立て(組立は製品コード=VBA が行う。
    tools/render_report.py の run_render をそのまま流用する)、node の最小DOM
    スタブでページ内スクリプトを**実際に走らせて**、次の4つを実測する。
      P0 meta.ground_unmatched = []        -> 印は0件(**無ければ出ない**)
      P1 meta.ground_unmatched = ["E1"]    -> 印は1件で、**1枚目のカードだけ**
      P2 ニューリスクを1件足して ["E2"]    -> 印は1件で、**2枚目のカードだけ**
      P3 meta.ground_unmatched = ["E9"]    -> 印は0件(範囲外の番号で全面に出ない)
    判定を反転させる変異では P0(0件のはずが全カードに出る)と P1(1枚目に出ない)
    の両方が赤くなる。採番を1つずらす変異(`'E'+i`)では P1/P2 が赤くなる。

    **同じ型の横展開**: SEC-14(risks[] の「原文照合」列。modHtmlTemplate5)も
    `gm[risk_no] ? '原文未照合' : ''` という同型の分岐で、実DOMの回帰が同じく
    無かった(純層テスト22も文字列 grep だけ)。素の描画から実在する risk_no を
    拾い、[]のときは1行も印が付かず、[その No]のときはその行だけに付くことを
    P0/P4 で測る。

検査②(配線・静的): 16章 E-02 の警告帯
    「判定はあるのに画面に出ない」を止めるのが目的なので、**呼び出しが在るか**
    を手続き単位で読む(ribbon_wire_check.py と同じ考え方)。
      (a) modUICase2.DrawStep が afterRun を受け取り、StepNoticeOf の戻り値を
          ShowStepNotice へ渡していること
      (b) modUICase2.ShowStepNotice が hm_warning とトーストの両方へ書くこと
      (c) modUIHome.DrawAllSteps が afterRun をそのまま DrawStep へ渡すこと
      (d) modUIHome2.RunStepUi / HomeRunAll が**実行直後だけ** True を渡すこと
      (e) modUIHome2.ShowDeepWarning が LastStepNotice() を取り直して併記すること
      (f) 実行ではない呼び口(modNaviActions の open_step_sheet・RefreshHome の
          描き直し・企業ファイル取込)が True を渡して**いない**こと
    (a)(e) は検証者が素通りさせた2つの変異(呼び出し行の削除・上書き)に対応する。

使い方:
    python3 tools/notice_check.py
    python3 tools/notice_check.py --verbose
    exit code: 0 = 両方OK / 1 = 検査NG / 2 = 検査を実行できない(node無し等)

    **node が無いときは 0 を返さない**(2で落とす)。「実行できなかった」を
    「合格」に化けさせない(裁定書40 §0 の fail-open 禁止)。
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

import render_report as rr  # noqa: E402  (HTMLの組立機構を流用。二重管理しない)

SRC_ROOT = REPO_ROOT / "src"

# 18章§3 SEC-09 の逐語の先頭(modHtmlTemplate4.SecNewRiskJs が出す1行)。
NOTE_HEAD = "原文未照合:"
NEWRISK_ID = "sec-newrisk"
CARD_CLASS = "radar-item"

# ---------------------------------------------------------------------------
# 検査①: 実DOM
# ---------------------------------------------------------------------------
DOM_STUB_JS = r"""
'use strict';
// 18章§4.1 が許す操作だけを持つ最小DOMスタブ(tools/render_report.py と同型)。
// 4パスを走らせ、SEC-09 のカードごとに本文を集めて返す。
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

function findById(root, id) {
  let hit = null;
  (function w(n) {
    if (hit) { return; }
    if (n.attrs && n.attrs.id === id) { hit = n; return; }
    for (const c of n.childNodes) { w(c); }
  })(root);
  return hit;
}
function byClass(node, cls) {
  const out = [];
  (function w(n) {
    if (n.className && (' ' + n.className + ' ').indexOf(' ' + cls + ' ') >= 0) { out.push(n); }
    for (const c of n.childNodes) { w(c); }
  })(node);
  return out;
}
function textsOf(node) {
  const out = [];
  (function w(n) {
    if (n._text) { out.push(n._text); }
    for (const c of n.childNodes) { w(c); }
  })(node);
  return out;
}

function runPass(mutate) {
  byId = Object.create(null);
  function mk(tag, id) { const n = new El(tag); if (id) { n.setAttribute('id', id); } return n; }
  const root = mk('body', null);
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

  const sec = findById(root, 'sec-newrisk');
  const cards = sec ? byClass(sec, 'radar-item') : [];
  // SEC-14(sec-source)の主表。行ごとに td の本文を並べて返す
  // (列は ['No','リスク名','引用した記述','出所','原文照合'])。
  const src = findById(root, 'sec-source');
  const rows = [];
  if (src) {
    (function w(n) {
      if (n.tagName === 'tr') {
        const cells = [];
        for (const c of n.childNodes) {
          if (c.tagName === 'td') { cells.push(c._text || ''); }
        }
        if (cells.length) { rows.push(cells); }
      }
      for (const c of n.childNodes) { w(c); }
    })(src);
  }
  return {
    hasSection: !!sec,
    cards: cards.map(function (c) { return textsOf(c); }),
    hasSource: !!src,
    rows: rows
  };
}

function setGu(list) {
  return function (D) {
    if (!D) { return; }
    if (!D.meta) { D.meta = {}; }
    D.meta.ground_unmatched = list;
  };
}
const out = {
  P0: runPass(setGu([])),
  P1: runPass(setGu(['E1'])),
  P2: runPass(function (D) {
    if (!D) { return; }
    if (!D.meta) { D.meta = {}; }
    const er = (D.s2 && D.s2.emerging_risks) ? D.s2.emerging_risks : [];
    if (er.length) {
      const extra = JSON.parse(JSON.stringify(er[0]));
      extra.risk_name = '検問用に足した2件目';
      er.push(extra);
      D.s2.emerging_risks = er;
    }
    D.meta.ground_unmatched = ['E2'];
  }),
  P3: runPass(setGu(['E9']))
};
// SEC-14(risks[] の「原文照合」列)も同じ型の分岐なので同じ両方向で測る。
// 突き合わせのキーは risk_no なので、素の描画から実在する No を拾って使う。
const firstNo = out.P0.rows.length ? out.P0.rows[0][0] : '';
out.firstNo = firstNo;
out.P4 = runPass(setGu([firstNo]));
console.log(JSON.stringify(out));
"""


def note_positions(cards: list[list[str]]) -> list[int]:
    """「原文未照合:」の1行を持つカードの添字(0始まり)。"""
    hit = []
    for i, texts in enumerate(cards):
        if any(t.startswith(NOTE_HEAD) for t in texts):
            hit.append(i)
    return hit


def check_dom_ground(verbose: bool) -> tuple[list[str], str]:
    """実DOMで SEC-09 の「原文未照合」を両方向に測る。"""
    node = shutil.which("node") or shutil.which("nodejs")
    if node is None:
        return (["node が見つからないため検査①(実DOM)を実行できません"
                 "(「実行できなかった」を合格にしない)"], "")

    soffice = rr.lo.find_soffice()
    work_dir = Path(tempfile.mkdtemp(prefix="rpn_notice_"))
    try:
        html = rr.run_render(soffice, work_dir, "standard", False, verbose, "")
        if html is None:
            return (["HTMLレポートを組み立てられませんでした"
                     "(tools/render_report.py の run_render が失敗)"], "")
        html_path = work_dir / "report.html"
        html_path.write_text(html, encoding="utf-8-sig", newline="")
        stub = work_dir / "notice_dom_stub.js"
        stub.write_text(DOM_STUB_JS, encoding="utf-8")
        proc = subprocess.run([node, str(stub), str(html_path)],
                              capture_output=True, text=True, timeout=120)
        if proc.returncode != 0:
            return ([f"DOM検査でページのJSが落ちました: {proc.stderr.strip()[:400]}"], "")
        try:
            got = json.loads(proc.stdout.strip().splitlines()[-1])
        except (ValueError, IndexError):
            return ([f"DOM検査の出力を読めませんでした: {proc.stdout.strip()[:200]}"], "")
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)

    problems: list[str] = []
    p0, p1, p2, p3 = got["P0"], got["P1"], got["P2"], got["P3"]

    if not p0["hasSection"] or len(p0["cards"]) < 1:
        return (["SEC-09(sec-newrisk)のカードが1枚も描かれていません"
                 "(素材に emerging_risks が無い=検査が空振り)"], "")

    n_base = len(p0["cards"])
    hit0 = note_positions(p0["cards"])
    hit1 = note_positions(p1["cards"])
    hit2 = note_positions(p2["cards"])
    hit3 = note_positions(p3["cards"])
    summary = (f"SEC-09 カード{n_base}枚 / 印の位置: "
               f"[]={hit0} [E1]={hit1} 2件目+[E2]={hit2} [E9]={hit3}")
    if verbose:
        print(f"[notice_check] {summary}")

    if hit0:
        problems.append(
            f"meta.ground_unmatched が空なのに「{NOTE_HEAD}」がカード{hit0}に"
            "出ています(18章§3 SEC-09。判定が反転している疑い)")
    if hit1 != [0]:
        problems.append(
            f"meta.ground_unmatched=[\"E1\"] のとき印が1枚目のカードだけに"
            f"出ていません(実際の位置={hit1})")
    if len(p2["cards"]) < 2:
        problems.append("パスP2 でニューリスクを2枚にできていません(検査が空振り)")
    elif hit2 != [1]:
        problems.append(
            f"meta.ground_unmatched=[\"E2\"] のとき印が2枚目のカードだけに"
            f"出ていません(実際の位置={hit2}。E採番は出現順の1始まり="
            "modGround.GroundNotes と同じ数え方)")
    if hit3:
        problems.append(
            f"範囲外の番号 [\"E9\"] で「{NOTE_HEAD}」がカード{hit3}に出ています"
            "(番号を突き合わせずに配列の有無だけを見ている疑い)")

    # --- 同じ型の横展開: SEC-14(risks[])の「原文照合」列も両方向で測る ---
    # 18章§3 SEC-14 は `gm[risk_no] ? '原文未照合' : ''` という同型の分岐で、
    # ここにも実DOMの回帰が無かった(裁定書40 Q-M2 の横展開)。
    rows0 = p0["rows"]
    if not p0.get("hasSource") or not rows0:
        problems.append("SEC-14(sec-source)の主表が1行も描かれていません(検査が空振り)")
        return (problems, summary)
    first_no = got.get("firstNo", "")
    marked0 = [r[0] for r in rows0 if len(r) > 4 and r[4].strip()]
    rows4 = got["P4"]["rows"]
    marked4 = [r[0] for r in rows4 if len(r) > 4 and r[4].strip()]
    summary += f" / SEC-14 {len(rows0)}行 印: []={marked0} [{first_no}]={marked4}"
    if verbose:
        print(f"[notice_check] SEC-14 rows={len(rows0)} []={marked0} "
              f"[{first_no}]={marked4}")
    if marked0:
        problems.append(
            f"meta.ground_unmatched が空なのに SEC-14 の「原文照合」列が"
            f"No.{marked0} で埋まっています(18章§3 SEC-14。判定が反転している疑い)")
    if marked4 != [first_no]:
        problems.append(
            f"meta.ground_unmatched=[\"{first_no}\"] のとき SEC-14 で印が付く行が"
            f"No.{first_no} だけになっていません(実際={marked4})")
    return (problems, summary)


# ---------------------------------------------------------------------------
# 検査②: 配線(静的)
# ---------------------------------------------------------------------------
PROC_HEAD = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function|Property\s+\w+)\s+(\w+)",
    re.IGNORECASE)
PROC_END = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.IGNORECASE)


def read_module(name: str) -> str:
    hits = sorted(SRC_ROOT.rglob(name))
    if not hits:
        raise FileNotFoundError(f"{name} が src/ 配下に見つかりません")
    return hits[0].read_text(encoding="utf-8", errors="replace")


def proc_body(text: str, proc_name: str) -> str | None:
    """手続き1本の本文(宣言行を含む)。見つからなければ None。"""
    lines = text.splitlines()
    out: list[str] = []
    inside = False
    for ln in lines:
        if not inside:
            m = PROC_HEAD.match(ln)
            if m and m.group(2).lower() == proc_name.lower():
                inside = True
                out.append(ln)
            continue
        out.append(ln)
        if PROC_END.match(ln):
            break
    return "\n".join(out) if out else None


def joined_body(text: str, proc_name: str) -> str | None:
    """行継続(` _`)を畳んだ手続き本文。1行にまたがる呼び出しを見るため。"""
    body = proc_body(text, proc_name)
    if body is None:
        return None
    return re.sub(r"\s+_\s*\n\s*", " ", body)


def want(problems: list[str], body: str | None, proc: str,
         needle: str, why: str) -> None:
    if body is None:
        problems.append(f"{proc} が見つかりません(改名・削除の疑い)")
        return
    if needle not in body:
        problems.append(f"{proc} に `{needle}` がありません({why})")


def want_not(problems: list[str], body: str | None, proc: str,
             needle: str, why: str) -> None:
    if body is None:
        problems.append(f"{proc} が見つかりません(改名・削除の疑い)")
        return
    if needle in body:
        problems.append(f"{proc} に `{needle}` があります({why})")


def check_banner_wire(verbose: bool) -> tuple[list[str], str]:
    problems: list[str] = []
    case2 = read_module("modUICase2.bas")
    home = read_module("modUIHome.bas")
    home2 = read_module("modUIHome2.bas")
    navi = read_module("modNaviActions.bas")

    # (a) 描画側: afterRun を受け取り、判定の戻り値を書き手へ渡している
    draw = joined_body(case2, "DrawStep")
    want(problems, draw, "modUICase2.DrawStep",
         "Optional ByVal afterRun As Boolean",
         "16章 E-02 は「実行後」の規定。実行直後かどうかを呼び口から受け取る")
    want(problems, draw, "modUICase2.DrawStep",
         "ShowStepNotice StepNoticeOf(stepNo, jsonText, afterRun)",
         "裁定書39 R1-07(a) の配線そのもの。この1行を消すと"
         "「判定はあるのに画面に出ない」が無検出で再発する")

    # (b) 書き手: 画面の2箇所(警告欄とトースト)へ出す
    show = joined_body(case2, "ShowStepNotice")
    want(problems, show, "modUICase2.ShowStepNotice", "WriteWarnCell",
         "HOME の警告欄(hm_warning)は WriteWarnCell の1本を通す"
         "(直に WriteNamed すると他の警告を上書きで消す)")
    want(problems, show, "modUICase2.ShowStepNotice", "modUIToast.ShowToast",
         "裁定書17 H2/H4: セルは見られていないのでトーストでも出す")
    want(problems, show, "modUICase2.ShowStepNotice", "gStepNotice = noticeText",
         "裁定書40 Q-m3: 上書きする側が取り直せるよう、出した本文を覚える")
    cell = joined_body(case2, "WriteWarnCell")
    want(problems, cell, "modUICase2.WriteWarnCell",
         "NoticeJoin(gStepNotice, TruncWarnText())",
         "裁定書40 Q-m3 の横展開: 16章 E-02 の帯と部屋あふれの警告が"
         "同じ1枠を奪い合うので、両方を併記する")
    want(problems, cell, "modUICase2.WriteWarnCell", "modUISheet.WriteNamed U2_WARN",
         "hm_warning へ実際に書くのはここ1本")
    trunc = joined_body(case2, "NoteTruncation")
    want(problems, trunc, "modUICase2.NoteTruncation", "WriteWarnCell",
         "部屋あふれの警告も同じ1本を通す(直に書くと E-02 の帯が消える)")
    want_not(problems, trunc, "modUICase2.NoteTruncation",
             "modUISheet.WriteNamed U2_WARN",
             "裁定書40 Q-m3 の横展開: ここで直に書くと直前の警告を上書きする")

    # (c) 一括描画: afterRun をそのまま通す(勝手に True にしない)
    drawall = joined_body(home, "DrawAllSteps")
    want(problems, drawall, "modUIHome.DrawAllSteps",
         "modUICase2.DrawStep caseId, n, afterRun",
         "呼び口が渡した値をそのまま通す(ここで True に固定すると"
         "案件切替の描き直しでも E-02 の帯が出る)")

    # (d) 実行直後の呼び口だけが True を渡す
    runstep = joined_body(home2, "RunStepUi")
    want(problems, runstep, "modUIHome2.RunStepUi",
         "modUICase2.DrawStep caseId, stepNo, True",
         "1段実行の直後は「実行後」なので帯を出してよい唯一の呼び口")
    want(problems, runstep, "modUIHome2.RunStepUi", "modUICase2.ResetStepNotice",
         "実行の開始時に前回の帯を1回だけ消す")
    runall = joined_body(home2, "HomeRunAll")
    want(problems, runall, "modUIHome2.HomeRunAll",
         "modUIHome.DrawAllSteps caseId, True",
         "一括実行の直後も「実行後」")
    want(problems, runall, "modUIHome2.HomeRunAll", "modUICase2.ResetStepNotice",
         "実行の開始時に前回の帯を1回だけ消す")

    # (e) deep の警告が E-02 の帯を上書きで消さない
    deep = joined_body(home2, "ShowDeepWarning")
    want(problems, deep, "modUIHome2.ShowDeepWarning",
         "modUICase2.NoticeJoin(modUICase2.LastStepNotice(), warnText)",
         "裁定書40 Q-m3: hm_warning は1枠なので、そのまま書くと E-02 の帯が消える")

    # (f) 実行ではない呼び口は True を渡さない
    for mod_name, text, proc in (("modNaviActions", navi, "Dispatch"),
                                 ("modUIHome", home, "RefreshHome"),
                                 ("modUIHome2", home2, "HomeCompanyOpen")):
        body = joined_body(text, proc)
        if body is None:
            continue
        for bad in ("DrawStep(caseId, n, True)", "DrawStep caseId, n, True",
                    "DrawAllSteps caseId, True"):
            want_not(problems, body, f"{mod_name}.{proc}", bad,
                     "実行していない呼び口なので 16章 E-02 の帯を出してはいけない")

    # 念のため全文でも「実行以外の場所での True 渡し」を数える。
    true_calls = 0
    for text in (case2, home, home2, navi):
        true_calls += len(re.findall(r"DrawStep\s*\(?\s*caseId,\s*\w+,\s*True",
                                     re.sub(r"\s+_\s*\n\s*", " ", text)))
        true_calls += len(re.findall(r"DrawAllSteps\s+caseId,\s*True",
                                     re.sub(r"\s+_\s*\n\s*", " ", text)))
    if true_calls != 2:
        problems.append(
            f"afterRun=True を渡している呼び口が{true_calls}箇所あります"
            "(16章 E-02 の「実行後」= RunStepUi と HomeRunAll の2箇所だけ)")
    summary = f"afterRun=True の呼び口 {true_calls}箇所(期待2)"
    if verbose:
        print(f"[notice_check] {summary}")
    return (problems, summary)


def main() -> int:
    ap = argparse.ArgumentParser(
        description="SEC-09の原文未照合(実DOM)と16章E-02の警告帯(配線)の回帰検問")
    ap.add_argument("--only", choices=["dom", "wire"], default=None,
                    help="片方だけ走らせる(既定は両方)")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args()

    problems: list[str] = []
    blocked = False
    dom_sum = wire_sum = "(未実行)"

    if args.only in (None, "wire"):
        wire_problems, wire_sum = check_banner_wire(args.verbose)
        problems += [f"検査②(配線): {p}" for p in wire_problems]

    if args.only in (None, "dom"):
        dom_problems, dom_sum = check_dom_ground(args.verbose)
        if dom_problems and not dom_sum:
            blocked = True
        problems += [f"検査①(実DOM): {p}" for p in dom_problems]

    if problems:
        print("[notice_check] NG:")
        for p in problems:
            print(f"  - {p}")
        return 2 if blocked else 1

    print(f"[notice_check] OK: 検査①(実DOM SEC-09) {dom_sum} / "
          f"検査②(E-02 配線) {wire_sum}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
