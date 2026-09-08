#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ui_check.py - HTML画面(ui/ と src/ui/navi/)の検問(裁定書34 §1.4・21本目のゲート)

================================================================================
なぜ要るか:
    W12-A で「利用者が触る画面」は HTML(モードレスの1枚窓)になった。画面は
    ui/ の5本(index.html / style.css / markdown.js / views.js / app.js)で、
    VBA 側の入口は modNaviHost の action 許可リストだけである。この2つは
    別のファイルなので、片方だけ直す事故が起こる。lint も LO もそこは見ない
    (HTML/JS はVBAではないし、LO はフォームを知らない)。本ツールがその隙間を
    埋める。**HTML画面が実際に描かれるかどうかは Windows 実機でしか確認できない**
    (17章§7 Z-43)。ここで見るのは「配線と閉じ込め」だけである。

検査(6項目。裁定書34 §1.4 の(1)〜(6)):
    (1) ui/ に外部URL(http:// / https:// / //cdn)が無い
        -> 会社PCは外部へ出られないし、出られる端末で出てしまうのはもっと悪い。
           画面は本体と同じフォルダの ui/ だけで閉じる。
    (2) ui/ に eval( / new Function( / document.write( が無い
        -> HTML への差し込みは VBA(modNaviHost.HostReadPage)が
           `<!--INLINE_STYLE-->` / `<!--INLINE_SCRIPT-->` の2箇所で行う。
           JS 側が自分で文字列をコードにする口は作らない(16章 NFR-S7)。
    (3) index.html に vbaPayload / vbaResponse の textarea がある
        -> HTML と VBA の受け渡しはこの2枚の textarea だけを通る(14章§7)。
    (4) JS が送る action 名の集合 == VBA(modNaviHost.IsAllowed)が許す集合
        -> 片方にしか無い名前を**両方向とも**列挙して ERROR。例外は
           VBA_ONLY_ALLOWED に理由つきで書いたものだけ。
    (5) 実フォーム src/ui/navi/frmNaviHtml.frm の Public と
        LO専用スタブ wintest/lo_stubs/frmNaviHtml_stub.bas の Public が一致
        -> ずれると「LOでは通るのに実Excelで落ちる(逆も)」が起きる。
    (6) ui/ の合計サイズが 300KB 以下
        -> 1枚のHTMLへ全部差し込んでから WebBrowser に食わせるので、
           大きくすると起動が目に見えて遅くなる。

使い方:
    python3 tools/ui_check.py
    python3 tools/ui_check.py --verbose     # 照合した action 名を全部出す
    exit code: 0 = ERROR 0件 / 1 = ERROR 1件以上 / 2 = 自己テスト失敗

自己テスト(骨抜き防止):
    毎回、末尾で負例(外部URL・eval・action名の食い違い・スタブの過不足)を
    合成データで走らせ、**負例で検出できなければ exit 2** で止める。
================================================================================
"""
from __future__ import annotations

import argparse
import os
import re
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
UI_DIR = os.path.join(REPO, "ui")
NAVI_DIR = os.path.join(REPO, "src", "ui", "navi")
FORM_PATH = os.path.join(NAVI_DIR, "frmNaviHtml.frm")
STUB_PATH = os.path.join(REPO, "wintest", "lo_stubs", "frmNaviHtml_stub.bas")
HOST_PATH = os.path.join(NAVI_DIR, "modNaviHost.bas")

UI_FILES = ("index.html", "style.css", "markdown.js", "views.js", "app.js")
UI_MAX_BYTES = 300 * 1024

# (1) 外部の持ち込み。`//cdn` は protocol-relative な CDN 参照。
EXTERNAL_URL = re.compile(r"https?://|(?<![:\w])//cdn[\w.-]*/", re.IGNORECASE)
# 例外: 仕様上どうしても文字列として現れる URL 風の綴り。**空にしておく**
# (例外を足すときは、なぜ外部へ出ないのかを1行書くこと)。
EXTERNAL_URL_EXEMPT: tuple[str, ...] = (
    # HTML の名前空間 URI。ブラウザは取りに行かない(識別子として使うだけ)。
    "http://www.w3.org/",
)

# (2) 文字列をコードにする口。
DYNAMIC_CODE = (
    (re.compile(r"\beval\s*\("), "eval("),
    (re.compile(r"\bnew\s+Function\s*\("), "new Function("),
    (re.compile(r"\bdocument\s*\.\s*write\s*\("), "document.write("),
)

# (3) HTML と VBA の受け渡し口(14章§7)。
REQUIRED_IDS = ("vbaPayload", "vbaResponse")

# (4) JS 側の action 名の出どころ。
#   ・send('x', ...) / send("x", ...) の第1引数(三項演算子の両側も拾う)
#   ・data-action="x" / data-confirm-action="x"(クリックで actionClick が
#     そのまま send する属性。値は静的なリテラルだけ)
SEND_CALL = re.compile(r"\bsend\s*\(([^();]{0,160})")
JS_STRING = re.compile(r"'([a-z][a-z0-9_]*)'|\"([a-z][a-z0-9_]*)\"")
DATA_ACTION_ATTR = re.compile(r"data-(?:confirm-)?action=\"([a-z][a-z0-9_]*)\"")

# VBA 側の許可リスト(modNaviHost.IsAllowed の Case 行)。
IS_ALLOWED_BLOCK = re.compile(
    r"Public\s+Function\s+IsAllowed\b.*?End\s+Function", re.DOTALL | re.IGNORECASE)
CASE_LITERAL = re.compile(r'"([a-z][a-z0-9_]*)"')

# VBA にはあるが画面が送らない action(理由つきでここに書いたものだけ許す)。
VBA_ONLY_ALLOWED = {
    # 第2段(F-16)の口。modNaviActions.PhaseTwoStepEdit は必ず断りを返す
    # (「直接編集は第2段の機能です。「シートで編集」を利用してください。」)。
    # 許可リストに載っているのは、載せずに来た要求を「未知の action」として
    # 一括で弾くのではなく、**用意はあるが今は断る**と画面へ伝えるため。
    "save_step_edit",
}

# (5) Public の抜き出し。フォームは Property Get / Public 変数も口である。
PUBLIC_DECL = re.compile(
    r"^\s*Public\s+(?:(?:Sub|Function)\s+(\w+)|Property\s+(?:Get|Let|Set)\s+(\w+)|(\w+)\s+As\s+)",
    re.IGNORECASE | re.MULTILINE)


def read(path: str) -> str:
    with open(path, encoding="utf-8") as fp:
        return fp.read()


# ---------------------------------------------------------------------------
# 抽出(自己テストから直接叩けるよう、ファイルI/Oと分けておく)
# ---------------------------------------------------------------------------
def external_urls(text: str) -> list[str]:
    hits = []
    for m in EXTERNAL_URL.finditer(text):
        tail = text[m.start():m.start() + 40]
        if any(tail.startswith(ex) for ex in EXTERNAL_URL_EXEMPT):
            continue
        hits.append(text[m.start():m.start() + 60].split('"')[0])
    return hits


def dynamic_code(text: str) -> list[str]:
    return [label for pat, label in DYNAMIC_CODE if pat.search(text)]


def js_actions(text: str) -> set[str]:
    """JS/HTML が送る action 名。send() の第1引数と data-action 属性から。"""
    out: set[str] = set()
    for m in SEND_CALL.finditer(text):
        args = m.group(1)
        head = args.split(",", 1)[0]          # 第1引数だけ(三項演算子は残る)
        for s in JS_STRING.finditer(head):
            out.add(s.group(1) or s.group(2))
    for m in DATA_ACTION_ATTR.finditer(text):
        out.add(m.group(1))
    return out


def vba_actions(host_text: str) -> set[str]:
    """modNaviHost.IsAllowed が True を返す action 名。"""
    block = IS_ALLOWED_BLOCK.search(host_text)
    if not block:
        return set()
    out: set[str] = set()
    for line in block.group(0).split("\n"):
        if re.match(r"\s*Case\b", line, re.IGNORECASE):
            out |= set(CASE_LITERAL.findall(line))
    return out


def public_names(text: str) -> set[str]:
    out: set[str] = set()
    for m in PUBLIC_DECL.finditer(text):
        name = m.group(1) or m.group(2) or m.group(3)
        if name and name.lower() != "const":
            out.add(name)
    return out


# ---------------------------------------------------------------------------
# 本検査
# ---------------------------------------------------------------------------
def run_checks(verbose: bool) -> int:
    errors = 0

    def err(msg: str) -> None:
        nonlocal errors
        errors += 1
        print("ERROR " + msg)

    # ---- 前提: ui/ の5本がそろっているか ----------------------------------
    if not os.path.isdir(UI_DIR):
        print("ERROR ui/ がありません(検査不能)")
        return 1
    texts: dict[str, str] = {}
    total = 0
    for name in UI_FILES:
        path = os.path.join(UI_DIR, name)
        if not os.path.exists(path):
            err("ui/%s がありません(HTML画面は5本そろって動きます)" % name)
            continue
        texts[name] = read(path)
        total += os.path.getsize(path)

    # ---- (1) 外部URL -------------------------------------------------------
    n1 = 0
    for name, text in texts.items():
        for hit in external_urls(text):
            err("(1) ui/%s に外部URL: %s" % (name, hit))
            n1 += 1
    print("  (1) 外部URL                    %d件" % n1)

    # ---- (2) 文字列をコードにする口 ----------------------------------------
    n2 = 0
    for name, text in texts.items():
        for label in dynamic_code(text):
            err("(2) ui/%s に %s があります" % (name, label))
            n2 += 1
    print("  (2) eval/new Function/write    %d件" % n2)

    # ---- (3) 受け渡しの textarea -------------------------------------------
    html = texts.get("index.html", "")
    missing_ids = [i for i in REQUIRED_IDS
                   if not re.search(r'<textarea[^>]*\bid="%s"' % i, html)]
    for i in missing_ids:
        err('(3) index.html に <textarea id="%s"> がありません(14章§7)' % i)
    print("  (3) 受け渡しの textarea        %d件不足" % len(missing_ids))

    # ---- (4) action 名の突合 -----------------------------------------------
    if not os.path.exists(HOST_PATH):
        err("(4) %s がありません" % os.path.relpath(HOST_PATH, REPO))
        js_set: set[str] = set()
        vba_set: set[str] = set()
    else:
        js_set = set()
        for name in ("app.js", "views.js", "index.html"):
            js_set |= js_actions(texts.get(name, ""))
        vba_set = vba_actions(read(HOST_PATH))
        if not vba_set:
            err("(4) modNaviHost.IsAllowed の Case を読み取れませんでした")
    only_js = sorted(js_set - vba_set)
    only_vba = sorted(vba_set - js_set - VBA_ONLY_ALLOWED)
    for a in only_js:
        err("(4) 画面が送るのに VBA が許可していない action: %s "
            "(modNaviHost.IsAllowed へ足すか、画面側の綴りを直す)" % a)
    for a in only_vba:
        err("(4) VBA が許可しているのに画面が送らない action: %s "
            "(画面へ配線するか、tools/ui_check.py の VBA_ONLY_ALLOWED へ理由つきで書く)" % a)
    print("  (4) action 名の突合            JS %d / VBA %d / 例外 %d / 不一致 %d"
          % (len(js_set), len(vba_set), len(VBA_ONLY_ALLOWED),
             len(only_js) + len(only_vba)))
    if verbose:
        print("      JS : " + ", ".join(sorted(js_set)))
        print("      VBA: " + ", ".join(sorted(vba_set)))

    # ---- (5) 実フォームと LO スタブの Public -------------------------------
    if not os.path.exists(FORM_PATH):
        err("(5) %s がありません" % os.path.relpath(FORM_PATH, REPO))
        diff_pub: list[str] = []
    elif not os.path.exists(STUB_PATH):
        err("(5) %s がありません(LO経路のコンパイルが通らなくなります)"
            % os.path.relpath(STUB_PATH, REPO))
        diff_pub = []
    else:
        form_pub = public_names(read(FORM_PATH))
        stub_pub = public_names(read(STUB_PATH))
        diff_pub = sorted(form_pub ^ stub_pub)
        for a in sorted(form_pub - stub_pub):
            err("(5) 実フォームにあって LOスタブに無い Public: %s" % a)
        for a in sorted(stub_pub - form_pub):
            err("(5) LOスタブにあって実フォームに無い Public: %s" % a)
        if verbose:
            print("      Form: " + ", ".join(sorted(form_pub)))
    print("  (5) フォームとLOスタブのPublic  不一致 %d件" % len(diff_pub))

    # ---- (6) 合計サイズ ----------------------------------------------------
    if total > UI_MAX_BYTES:
        err("(6) ui/ の合計が %s bytes で上限 %s bytes を超えています"
            % (format(total, ","), format(UI_MAX_BYTES, ",")))
    print("  (6) 合計サイズ                 %s / %s bytes"
          % (format(total, ","), format(UI_MAX_BYTES, ",")))

    return errors


# ---------------------------------------------------------------------------
# 自己テスト(負例で検出できなければ検問が壊れている)
# ---------------------------------------------------------------------------
def self_test() -> bool:
    cases = []

    # (1) 外部URL: 負例=CDN参照 / 正例=名前空間URIだけ
    cases.append(("外部URL 負例",
                  len(external_urls('<script src="https://cdn.x/a.js">')) > 0))
    cases.append(("外部URL 負例(protocol-relative)",
                  len(external_urls('<link href="//cdnjs.com/a.css">')) > 0))
    cases.append(("外部URL 正例",
                  len(external_urls('<svg xmlns="http://www.w3.org/2000/svg">')) == 0))

    # (2) 動的コード
    cases.append(("eval 負例", "eval(" in dynamic_code("var x=eval('1+1');")))
    cases.append(("new Function 負例",
                  "new Function(" in dynamic_code("var f=new Function('return 1');")))
    cases.append(("動的コード 正例", dynamic_code("var evaluated=1;") == []))

    # (4) action 名の抽出
    js = js_actions("send('open_case',d);send(spar?'sparring_send':'chat',d);"
                    '<button data-action="company_save" data-confirm-action="run_tests">')
    cases.append(("action抽出 send", {"open_case", "sparring_send", "chat"} <= js))
    cases.append(("action抽出 属性", {"company_save", "run_tests"} <= js))
    cases.append(("action抽出 第2引数は拾わない", "d" not in js))

    vba = vba_actions('Public Function IsAllowed(ByVal action As String) As Boolean\n'
                      '    Select Case action\n'
                      '    Case "initialize", "chat"\n'
                      '        IsAllowed = True\n'
                      '    End Select\n'
                      'End Function\n')
    cases.append(("VBA許可抽出", vba == {"initialize", "chat"}))
    cases.append(("突合 負例(画面だけ)", len({"open_case"} - vba) > 0))
    cases.append(("突合 負例(VBAだけ)", len(vba - {"initialize"}) > 0))

    # (5) Public の抽出
    pub = public_names("Public AllowClose As Boolean\n"
                       "Public Sub CycleSize()\n"
                       "Public Function TakePendingText() As String\n"
                       "Public Property Get IsReady() As Boolean\n"
                       "Private Sub Hidden()\n")
    cases.append(("Public抽出",
                  pub == {"AllowClose", "CycleSize", "TakePendingText", "IsReady"}))
    cases.append(("Public抽出 負例(過不足を見つける)",
                  len(pub ^ {"AllowClose", "CycleSize"}) > 0))

    bad = [name for name, ok in cases if not ok]
    for name in bad:
        print("  自己テスト NG: %s" % name)
    print("  自己テスト: %d/%d" % (len(cases) - len(bad), len(cases)))
    return not bad


def main() -> int:
    ap = argparse.ArgumentParser(description="HTML画面(ui/ と src/ui/navi/)の検問")
    ap.add_argument("--verbose", action="store_true",
                    help="照合した action 名・Public 名を全部出す")
    args = ap.parse_args()

    print("ui_check: HTML画面の配線と閉じ込め(裁定書34 §1.4)")
    errors = run_checks(args.verbose)

    if not self_test():
        print("結果: 自己テスト失敗(検出器が壊れています)")
        return 2

    if errors:
        print("ERROR: %d 件" % errors)
        print("結果: NG")
        return 1
    print("結果: OK 6条件(外部URL無し / 動的コード無し / 受け渡しtextarea / "
          "action名の一致 / フォームとLOスタブのPublic一致 / 合計サイズ)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
