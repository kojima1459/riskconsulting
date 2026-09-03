#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""caption_check.py - ナビのボタン名(キャプション)の4系統逐語照合(W5.2.1 裁定)。

なぜ要るのか(W5.2 検証班の指摘):
  ナビのボタン名は**同じ文字列が4か所に別々に書かれている**。実装(配置表の
  定数)・使い方タブのボタン早見・13章§2.10(f)の配置表・初回ツアーの文言で、
  どれか1つだけ直すと利用者は「押せと書いてあるボタンが画面に無い」状態に
  なる。W5/W5.2 で実際に旧名([① 案件を作る] 等)が文書に残った。
  人手の目視でしか守れていなかったこの一致を、静的検問で機械化する。

照合する4系統(v3.2: HOME 廃止によりナビへ移した):
  (A) src/ui/modUINav.bas の UN_ROW_COACH / UN_ROW_SEC1..4
      = 「図形名;キャプション;OnAction;幅pt」の配置表。**これが唯一の値源**。
  (B) build/build_rpn.py の GUIDE_NAV_BUTTONS(使い方タブのボタン早見)
  (C) docs/spec/13_データ設計.md §2.10(f) の配置表([～]表記)
  (D) src/ui/modUIGuide.bas の初回ツアー文言(TitleOf / BodyOf)の [～] 表記
      ※ツアーの本文は modUINav.StepText を引くので、そちらも走査する。

  (A)(B)(C) は**集合として完全一致**を要求する。(D) はツアーが全ボタンに
  触れるわけではないので **(A)の部分集合**であることだけを要求する。

(E) 使い方タブ⑦「上級」の動作ボタン3本(v3.2.1・裁定書22):
  [第2ラウンドを始める][企業ファイルへ保存][企業ファイルを開く]は**ナビの
  配置表には載らない**(ナビに置かないボタンなので (A) の集合には入れない)が、
  同じ「同じ文字列が3か所に別々に書かれている」問題を持つ。
      (E1) src/ui/modUIGuide.bas の AdvActionRow()(図形名;キャプション;
           OnAction;幅pt) = **唯一の値源**
      (E2) build/build_rpn.py の GUIDE_ADVANCED_ACTIONS(キャプション・
           いつ使うか・OnAction)
      (E3) docs/spec/13_データ設計.md §2.18 の動作ボタン表([～]表記と OnAction)
  キャプションと OnAction の**両方**を集合として完全一致で照合する
  (OnAction を取り違えると押しても別の処理が走る=キャプション照合だけでは
   捕まらない)。

(F) 区画②の7欄の説明1行・例文2行(v2.6・裁定書25 S3。17章 T-56 のDoD③):
  11章§3.3.3(a) の表が**唯一の逐語の値源**であり、実体は
  `build/sheets_main.json` のナビ区画②(`nv_sec2`)のテキスト行にある。
  W7 で7本目「決算・財務」を足したときに、この2系統がずれていても誰も
  気づけなかった(ボタン名と違い、説明文はどの検問も見ていなかった)。
  照合は欄ごとに「説明1行 + 例文2行」を順序つきで突き合わせる。

  **表のセルの引用符の読み方(11章§3.3.3(a) の書式)**: セルは全体を「」で
  括り、その内側の引用符は『』で表す(§3.3.2 のワイヤーが示す実際の表示は
  内側も「」である)。したがって照合の前に、セル全体の外側の「」を1組だけ
  外し、内側の『』を「」へ戻す。**この2つ以外の正規化は行わない**
  (1字でも違えば赤にする)。

繰り返しキャプションを載せない規約(13章§2.10(f)):
  区画①の[コピー]8本と区画②の6欄×3種18本は、同じ語が複数出ると逐語照合の
  突合が壊れるため**配置表にも早見にも載せない**(生成規則の側が正を持つ)。
  実装の配置表にキャプションの重複があれば、その時点で赤にする。

骨抜き防止(裁定):
  対象ファイルが無い・アンカー(定数名/表/関数)が見つからない・抽出結果が
  0件のときは、照合が「たまたま通る」ことが無いように**必ず exit 1** する。

使い方: python3 tools/caption_check.py   (gate.py の "caption" ゲート)
"""
import io
import json
import os
import re
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

SRC_NAV = os.path.join("src", "ui", "modUINav.bas")
SRC_GUIDE = os.path.join("src", "ui", "modUIGuide.bas")
SRC_BUILD = os.path.join("build", "build_rpn.py")
SRC_SPEC = os.path.join("docs", "spec", "13_データ設計.md")

# (A) の配置表を持つ定数名(この並びが画面上の並び順でもある)。
ROW_CONSTS_COACH = ["UN_ROW_COACH"]
ROW_CONSTS_SEC = ["UN_ROW_SEC1B", "UN_ROW_SEC1", "UN_ROW_SEC2",
                  "UN_ROW_SEC3", "UN_ROW_SEC4"]

# 早見(B)だけに許す集約表記 -> 展開後のキャプション。いまは1件も無い。
AGGREGATES: dict = {}

# 13章§2.10(f) が「配置表に載せない」と定めた**繰り返しキャプション**。
# 同じ語が画面に複数出るため配置表に置けない(逐語照合の突合が壊れる)。生成規則の
# 側が正を持つ: [コピー]=modUIResearch の8本 / [ここに貼る][中身を見る][消す]=
# modUICase6 の6欄×3種 / [表示する]=modUIGuide の上級5枚 / [ナビへ戻る]=
# modUISheet.EnsureBackButton が可視化したシートへ置く1本。
# ツアーやコーチ帯の文がこれらの語に触れても、(D)の部分集合検査では見逃す。
REPEATING_CAPTIONS = {
    "コピー", "ここに貼る", "中身を見る", "消す", "表示する", "ナビへ戻る",
    "記録を見る", "テストを実行", "ツアーをもう一度見る",
}

SRC_SPEC11 = os.path.join("docs", "spec", "11_画面設計.md")
SRC_LEDGER = os.path.join("build", "sheets_main.json")

BRACKET = re.compile(r"\[([^\[\]]+)\]")
# (E3) 13章§2.18 の動作ボタン表の行(1列目が [～] で始まる行だけを表とみなす)。
SPEC_ADV_ROW = re.compile(r"^\|\s*\[([^\]]+)\]\s*\|")
errors = []


def fail(msg):
    errors.append(msg)


def read(rel):
    path = os.path.join(REPO, rel)
    if not os.path.isfile(path):
        fail("対象ファイルが見つからない: %s" % rel)
        return None
    return io.open(path, encoding="utf-8").read()


def vba_const_value(text, name, rel):
    """VBAの `Private Const <name> As String = ...` の値(文字列リテラル群)を返す。

    行末の ` _` による継続行を1つの論理行へつないでから、二重引用符の
    リテラルだけを取り出す(& vbLf & の連結は区切りとして扱えばよい)。
    """
    pat = re.compile(r"^\s*(?:Private|Public)\s+Const\s+%s\s+As\s+String\s*=" % re.escape(name),
                     re.M)
    m = pat.search(text)
    if not m:
        fail("%s に定数 %s が見つからない(配置表の書き方が変わった可能性)" % (rel, name))
        return []
    lines = text[m.start():].split("\n")
    logical = []
    for ln in lines:
        stripped = ln.rstrip()
        logical.append(stripped)
        if not stripped.endswith(" _"):
            break
    stmt = "\n".join(logical)
    return re.findall(r'"([^"]*)"', stmt)


def captions_from_nav():
    """(A) 実装の配置表。戻り値 (コーチ帯, 4区画) の順序つきリスト。"""
    text = read(SRC_NAV)
    if text is None:
        return [], []

    def collect(names):
        out = []
        for nm in names:
            for lit in vba_const_value(text, nm, SRC_NAV):
                flds = lit.split(";")
                if len(flds) < 4:
                    fail("%s の %s に「図形名;キャプション;OnAction;幅pt」でない要素: %r"
                         % (SRC_NAV, nm, lit))
                    continue
                cap = flds[1].strip()
                if not cap:
                    fail("%s の %s にキャプションが空の要素がある" % (SRC_NAV, nm))
                    continue
                out.append(cap)
        return out

    return collect(ROW_CONSTS_COACH), collect(ROW_CONSTS_SEC)


def captions_from_guide_table():
    """(B) build_rpn.py の GUIDE_NAV_BUTTONS。集約表記は展開して返す。
    要素は (キャプション, 説明, 置き場所の章) の3つ組。"""
    text = read(SRC_BUILD)
    if text is None:
        return []
    m = re.search(r"^GUIDE_NAV_BUTTONS\s*=\s*\[(.*?)^\]", text, re.M | re.S)
    if not m:
        fail("%s に GUIDE_NAV_BUTTONS の定義が見つからない" % SRC_BUILD)
        return []
    out = []
    for cap, _desc, _ch in re.findall(
            r'\(\s*"([^"]*)"\s*,\s*"([^"]*)"\s*,\s*"([^"]*)"\s*\)', m.group(1)):
        out.extend(AGGREGATES.get(cap, [cap]))
    return out


# 13章§2.10(f) の配置表の1列目に現れる区画名(この行だけを配置表とみなす)。
SPEC_SECTION_CELLS = ("帯", "①", "②", "③", "④")


def captions_from_spec():
    """(C) 13章§2.10(f) の配置表。ボタン列の [～] を拾う。"""
    text = read(SRC_SPEC)
    if text is None:
        return []
    m = re.search(r"^### 2\.10 .*?$(.*?)^### 2\.11 ", text, re.M | re.S)
    if not m:
        fail("%s に §2.10 の節が見つからない" % SRC_SPEC)
        return []
    out = []
    rows = 0
    for line in m.group(1).split("\n"):
        cells0 = [c.strip() for c in line.strip().strip("|").split("|")]
        if not line.startswith("|") or cells0[0] not in SPEC_SECTION_CELLS:
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) < 3:
            fail("%s §2.10(f) の配置表の行が3列でない: %r" % (SRC_SPEC, line[:40]))
            continue
        rows += 1
        out.extend(c.strip() for c in BRACKET.findall(cells[-1]))
    if rows == 0:
        fail("%s §2.10(f) に配置表の行(区画名「帯」「①」～「④」)が1行も無い" % SRC_SPEC)
    return out


def captions_from_tour():
    """(D) 初回ツアーの文言に現れる [～] 表記。

    v3.2 でツアーの本文は modUINav.StepText を引く形になった(11章§3.6 末尾:
    「文面はコーチ帯の6文から流用し、1字も別の文を作らない」)。したがって
    modUIGuide の TitleOf / BodyOf と、modUINav.StepText の両方を走査する。
    """
    out = []
    found = 0
    text = read(SRC_GUIDE)
    if text is not None:
        for fn in ("TitleOf", "BodyOf"):
            m = re.search(r"^Private Function %s\b.*?^End Function" % fn, text, re.M | re.S)
            if not m:
                fail("%s に %s が見つからない(ツアー文言の置き場所が変わった可能性)"
                     % (SRC_GUIDE, fn))
                continue
            found += 1
            for lit in re.findall(r'"([^"]*)"', m.group(0)):
                out.extend(c.strip() for c in BRACKET.findall(lit))
    nav = read(SRC_NAV)
    if nav is not None:
        m = re.search(r"^Public Function StepText\b.*?^End Function", nav, re.M | re.S)
        if not m:
            fail("%s に StepText が見つからない(コーチ帯の逐語表の置き場所が変わった)"
                 % SRC_NAV)
        else:
            found += 1
            for lit in re.findall(r'"([^"]*)"', m.group(0)):
                out.extend(c.strip() for c in BRACKET.findall(lit))
    if found == 0:
        fail("ツアー文言を1本も読めなかった")
    return out


def adv_actions_from_guide():
    """(E1) modUIGuide.AdvActionRow() の配置表。(キャプション, OnAction) の順序つき。"""
    text = read(SRC_GUIDE)
    if text is None:
        return []
    out = []
    for lit in vba_const_value(text, "UG_ROW_ADV_ACT", SRC_GUIDE):
        flds = lit.split(";")
        if len(flds) < 4:
            fail("%s の UG_ROW_ADV_ACT に「図形名;キャプション;OnAction;幅pt」でない要素: %r"
                 % (SRC_GUIDE, lit))
            continue
        cap = flds[1].strip()
        act = flds[2].strip()
        if not cap or not act:
            fail("%s の UG_ROW_ADV_ACT にキャプションかOnActionが空の要素がある" % SRC_GUIDE)
            continue
        out.append((cap, act))
    return out


def adv_actions_from_build():
    """(E2) build_rpn.py の GUIDE_ADVANCED_ACTIONS。(キャプション, OnAction)。"""
    text = read(SRC_BUILD)
    if text is None:
        return []
    m = re.search(r"^GUIDE_ADVANCED_ACTIONS\s*=\s*\[(.*?)^\]", text, re.M | re.S)
    if not m:
        fail("%s に GUIDE_ADVANCED_ACTIONS の定義が見つからない" % SRC_BUILD)
        return []
    out = []
    for cap, _when, act in re.findall(
            r'\(\s*"([^"]*)"\s*,\s*\n?\s*"([^"]*)"\s*,\s*\n?\s*"([^"]*)"\s*,?\s*\)',
            m.group(1)):
        out.append((cap.strip(), act.strip()))
    return out


def adv_actions_from_spec():
    """(E3) 13章§2.18 の動作ボタン表。(キャプション, OnAction)。"""
    text = read(SRC_SPEC)
    if text is None:
        return []
    m = re.search(r"^### 2\.18 .*?$(.*?)^### 2\.19 ", text, re.M | re.S)
    if not m:
        fail("%s に §2.18 の節が見つからない" % SRC_SPEC)
        return []
    out = []
    for line in m.group(1).split("\n"):
        hit = SPEC_ADV_ROW.match(line.strip())
        if not hit:
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) < 3:
            fail("%s §2.18 の動作ボタン表の行が3列以上でない: %r" % (SRC_SPEC, line[:40]))
            continue
        acts = re.findall(r"`([\w.]+)`", cells[2])
        if not acts:
            fail("%s §2.18 の動作ボタン表の行に OnAction が無い: %r" % (SRC_SPEC, line[:40]))
            continue
        out.append((hit.group(1).strip(), acts[0]))
    return out


def _unquote_cell(cell):
    """11章§3.3.3(a) の表のセルを、画面へ出る実際の文字列へ戻す。

    (1) 全体を括る「」を1組だけ外す (2) 内側の『』を「」へ戻す。
    それ以外の正規化はしない(部分一致・空白の畳み込みもしない)。
    """
    s = cell.strip()
    if s.startswith("「") and s.endswith("」"):
        s = s[1:-1]
    return s.replace("『", "「").replace("』", "」")


def field_texts_from_spec11():
    """(F) 11章§3.3.3(a) の表 -> {欄名: [説明, 例文1, 例文2, ...]}。"""
    text = read(SRC_SPEC11)
    if text is None:
        return {}
    m = re.search(r"^##### \(a\) .*?$(.*?)^##### \(b\) ", text, re.M | re.S)
    if not m:
        fail("%s に §3.3.3(a) の節が見つからない" % SRC_SPEC11)
        return {}
    out = {}
    for line in m.group(1).split("\n"):
        if not line.startswith("| `"):
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) < 3:
            fail("%s §3.3.3(a) の表の行が3列でない: %r" % (SRC_SPEC11, line[:40]))
            continue
        name = cells[0].strip("`").strip()
        rows = [_unquote_cell(cells[1])]
        rows += [_unquote_cell(x) for x in cells[2].split("<br>") if x.strip()]
        out[name] = rows
    if not out:
        fail("%s §3.3.3(a) の表から1欄も読み取れなかった" % SRC_SPEC11)
    return out


def field_texts_from_ledger():
    """(F) build/sheets_main.json のナビ区画② -> {欄名: [説明, 例文...]}。

    見出し行 `── ⑥ 決算・財務 ──　[...]` から次の見出し行までの text 要素の
    うち、状態行などのレンジ行より前に並ぶ連続したテキストを拾う。
    """
    path = os.path.join(REPO, SRC_LEDGER)
    if not os.path.isfile(path):
        fail("対象ファイルが見つからない: %s" % SRC_LEDGER)
        return {}
    data = json.load(io.open(path, encoding="utf-8"))
    nav = None
    for sh in data.get("sheets", []):
        if sh.get("name") == "ナビ":
            nav = sh
    if nav is None:
        fail("%s に「ナビ」シートが無い" % SRC_LEDGER)
        return {}
    sec2 = None
    for sec in nav.get("sections", []):
        if sec.get("anchor") == "nv_sec2":
            sec2 = sec
    if sec2 is None:
        fail("%s のナビに区画②(nv_sec2)が無い" % SRC_LEDGER)
        return {}

    head = re.compile(r"^──\s*[①-⑳]\s*(.+?)\s*──")
    out = {}
    cur = None
    for fld in sec2.get("fields", []):
        txt = fld.get("text")
        if txt is None:
            # レンジ行。欄の説明・例文はこの行より前にしか置かない。
            cur = None
            continue
        hit = head.match(txt)
        if hit:
            cur = hit.group(1).strip()
            out[cur] = []
        elif cur is not None:
            out[cur].append(txt)
    if not out:
        fail("%s のナビ区画②から欄の見出しを1本も読み取れなかった" % SRC_LEDGER)
    return out


def check_field_texts():
    """(F) 11章§3.3.3(a) <-> build/sheets_main.json の逐語照合。"""
    spec = field_texts_from_spec11()
    led = field_texts_from_ledger()
    if not spec or not led:
        return 0
    # 現場メモは(b)の枠内先置きであり(a)の表に無い。台帳側だけにあってよい。
    for name, want in spec.items():
        got = led.get(name)
        if got is None:
            fail("11章§3.3.3(a) の欄「%s」がナビ区画②(%s)に無い" % (name, SRC_LEDGER))
            continue
        if got != want:
            fail("欄「%s」の説明・例文が11章§3.3.3(a)と逐語一致しない:\n"
                 "      11章 = %r\n      台帳 = %r" % (name, want, got))
    return len(spec)


# (G) フッター(裁定書26 D)。キャプションの先頭の丸C(U+00A9)は CP932 に無いので
# 実装は ChrW() で組み立てる(定数には残りの語だけが入る)。したがって (A) の
# 配置表からは読めない。実装の UN_FOOTER_TEXT と 13章§2.10 のフッター行を
# 直接突き合わせる。**ナビと使い方に同じ1本を置く**ので (A) の集合には入れない
# (同じ語が2枚に出るため、配置表に載せると重複検査で赤になる)。
FOOTER_MARK = "©"
FOOTER_SHAPE = "btn_nv_footer"


def footer_caption_from_nav():
    """(G1) 実装のフッターのキャプション。"""
    text = read(SRC_NAV)
    if text is None:
        return None
    m = re.search(r'^\s*Private\s+Const\s+UN_FOOTER_TEXT\s+As\s+String\s*=\s*"([^"]*)"',
                  text, re.M)
    if not m:
        fail("%s に定数 UN_FOOTER_TEXT が見つからない(フッターの値源が変わった可能性)"
             % SRC_NAV)
        return None
    return FOOTER_MARK + m.group(1)


def footer_caption_from_spec():
    """(G2) 13章§2.10 のフッター行に書かれたキャプション。"""
    text = read(SRC_SPEC)
    if text is None:
        return None
    m = re.search(r"^### 2\.10 .*?$(.*?)^### 2\.11 ", text, re.M | re.S)
    if not m:
        fail("%s に §2.10 の節が見つからない" % SRC_SPEC)
        return None
    for line in m.group(1).split("\n"):
        if FOOTER_SHAPE not in line:
            continue
        caps = BRACKET.findall(line)
        if not caps:
            fail("%s §2.10 のフッター行に [～] のキャプションが無い" % SRC_SPEC)
            return None
        return caps[0].strip()
    fail("%s §2.10 に図形名 %s を書いたフッターの行が無い" % (SRC_SPEC, FOOTER_SHAPE))
    return None


def check_footer():
    """(G) フッターのキャプションの逐語照合。戻り値=照合したキャプション。"""
    impl = footer_caption_from_nav()
    spec = footer_caption_from_spec()
    if impl is None or spec is None:
        return None
    if impl != spec:
        fail("フッターのキャプションが13章§2.10と逐語一致しない:\n"
             "      実装 = %r\n      13章 = %r" % (impl, spec))
    return impl


def compare_sets(label, impl, other):
    missing = sorted(set(impl) - set(other))
    extra = sorted(set(other) - set(impl))
    for cap in missing:
        fail("%s に実装のキャプションが無い: [%s]" % (label, cap))
    for cap in extra:
        fail("%s にだけあるキャプション(実装に無い): [%s]" % (label, cap))


def main():
    main_caps, sub_caps = captions_from_nav()
    impl = main_caps + sub_caps

    # 骨抜き防止: 抽出0件は「一致した」ではなく検問の失敗として扱う。
    if not main_caps:
        fail("%s からコーチ帯のキャプションを1件も抽出できなかった" % SRC_NAV)
    if not sub_caps:
        fail("%s から4区画のキャプションを1件も抽出できなかった" % SRC_NAV)
    if len(impl) != len(set(impl)):
        dup = sorted({c for c in impl if impl.count(c) > 1})
        fail("実装の配置表にキャプションの重複がある: %s" % ", ".join(dup))

    guide = captions_from_guide_table()
    spec = captions_from_spec()
    tour = captions_from_tour()

    if impl:
        if not guide:
            fail("%s の GUIDE_NAV_BUTTONS からキャプションを1件も抽出できなかった" % SRC_BUILD)
        else:
            compare_sets("使い方タブのボタン早見(%s の GUIDE_NAV_BUTTONS)" % SRC_BUILD,
                         impl, guide)
        if not spec:
            fail("%s §2.10(f) の配置表からキャプションを1件も抽出できなかった" % SRC_SPEC)
        else:
            compare_sets("13章§2.10(f) の配置表", impl, spec)
        if not tour:
            fail("%s のツアー文言から [～] 表記を1件も抽出できなかった" % SRC_GUIDE)
        else:
            # ツアーは全ボタンに触れないので**部分集合**であることだけを求める。
            # 繰り返しキャプション(13章§2.10(f))は配置表に載らないので除く。
            for cap in sorted(set(tour) - set(impl) - REPEATING_CAPTIONS):
                fail("初回ツアーの文言に、実装に無いボタン名がある: [%s] (%s)" % (cap, SRC_GUIDE))

    # (E) 使い方タブ⑦の動作ボタン3本(v3.2.1・裁定書22)。
    adv_impl = adv_actions_from_guide()
    adv_build = adv_actions_from_build()
    adv_spec = adv_actions_from_spec()
    if not adv_impl:
        fail("%s の AdvActionRow()(UG_ROW_ADV_ACT)から⑦上級の動作ボタンを"
             "1件も抽出できなかった" % SRC_GUIDE)
    else:
        if not adv_build:
            fail("%s の GUIDE_ADVANCED_ACTIONS から1件も抽出できなかった" % SRC_BUILD)
        else:
            compare_sets("使い方タブ⑦の動作ボタン(%s の GUIDE_ADVANCED_ACTIONS)"
                         % SRC_BUILD, adv_impl, adv_build)
        if not adv_spec:
            fail("%s §2.18 の動作ボタン表から1件も抽出できなかった" % SRC_SPEC)
        else:
            compare_sets("13章§2.18 の動作ボタン表", adv_impl, adv_spec)

    # (F) 区画②の7欄の説明1行・例文2行(11章§3.3.3(a)。v2.6・T-56)。
    n_fields = check_field_texts()

    # (G) フッター(裁定書26 D)。
    footer = check_footer()

    if errors:
        print("[caption_check] ナビのボタン名の逐語照合")
        for e in errors:
            print("  NG: " + e)
        print("結果: NG %d件。値源は %s の配置表(UN_ROW_*)であり、"
              "早見・13章§2.10(f)・ツアー文言をそれに合わせる。" % (len(errors), SRC_NAV))
        return 1

    print("照合先: 使い方タブ早見 %d件 / 13章§2.10(f) %d件 / ツアー文言 %d件(部分集合)"
          % (len(guide), len(spec), len(tour)))
    print("使い方タブ⑦の動作ボタン: %d本(キャプション＋OnActionを build と 13章§2.18 へ照合)"
          % len(adv_impl))
    print("区画②の欄の説明・例文: %d欄を11章§3.3.3(a)と逐語照合" % n_fields)
    print("フッター: [%s] を13章§2.10と逐語照合" % (footer or ""))
    print("OK: 全%d本のキャプションが3系統と一致しました（コーチ帯%d / 4区画%d）"
          % (len(impl), len(main_caps), len(sub_caps)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
