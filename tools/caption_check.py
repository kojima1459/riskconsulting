#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""caption_check.py - HOMEのボタン名(キャプション)の4系統逐語照合(W5.2.1 裁定)。

なぜ要るのか(W5.2 検証班の指摘):
  HOMEのボタン名は**同じ文字列が4か所に別々に書かれている**。実装(配置表の
  定数)・操作ガイド③のボタン早見表・13章§2.10の配置表・初回ツアーの文言で、
  どれか1つだけ直すと利用者は「押せと書いてあるボタンが画面に無い」状態に
  なる。W5/W5.2 で実際に旧名([① 案件を作る] 等)が文書に残った。
  人手の目視でしか守れていなかったこの一致を、静的検問で機械化する。

照合する4系統:
  (A) src/ui/modUIHome.bas の UH_ROW_MAIN1..3 / UH_ROW_SUB1..4
      = 「図形名;キャプション;OnAction;幅pt」の配置表。**これが唯一の値源**。
  (B) build/build_rpn.py の GUIDE_HOME_BUTTONS(操作ガイド③のボタン早見表)
  (C) docs/spec/13_データ設計.md §2.10 の配置表([～]表記)
  (D) src/ui/modUIGuide.bas の初回ツアー文言(TitleOf / BodyOf)の [～] 表記

  (A)(B)(C) は**集合として完全一致**を要求する。(D) はツアーが全ボタンに
  触れるわけではないので **(A)の部分集合**であることだけを要求する。

集約表記の許容ルール(13章§2.10 の但し書き):
  早見表(B)は [S1] [S2] [S3] [S4] の4本を「S1 / S2 / S3 / S4」の1行へ集約
  してよい。この1件だけを AGGREGATES で展開してから照合する。**新しい集約を
  足すときはここへ書く**(照合の抜け道を実装側に作らない)。

骨抜き防止(裁定):
  対象ファイルが無い・アンカー(定数名/表/関数)が見つからない・抽出結果が
  0件のときは、照合が「たまたま通る」ことが無いように**必ず exit 1** する。

使い方: python3 tools/caption_check.py   (gate.py の "caption" ゲート)
"""
import io
import os
import re
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

SRC_HOME = os.path.join("src", "ui", "modUIHome.bas")
SRC_GUIDE = os.path.join("src", "ui", "modUIGuide.bas")
SRC_BUILD = os.path.join("build", "build_rpn.py")
SRC_SPEC = os.path.join("docs", "spec", "13_データ設計.md")

# (A) の配置表を持つ定数名(この並びが画面上の並び順でもある)。
ROW_CONSTS_MAIN = ["UH_ROW_MAIN1", "UH_ROW_MAIN2", "UH_ROW_MAIN3"]
ROW_CONSTS_SUB = ["UH_ROW_SUB1", "UH_ROW_SUB2", "UH_ROW_SUB3", "UH_ROW_SUB4"]

# 早見表(B)だけに許す集約表記 -> 展開後のキャプション。
AGGREGATES = {"S1 / S2 / S3 / S4": ["S1", "S2", "S3", "S4"]}

BRACKET = re.compile(r"\[([^\[\]]+)\]")
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


def captions_from_home():
    """(A) 実装の配置表。戻り値 (主要動線, くわしい操作) の順序つきリスト。"""
    text = read(SRC_HOME)
    if text is None:
        return [], []

    def collect(names):
        out = []
        for nm in names:
            for lit in vba_const_value(text, nm, SRC_HOME):
                flds = lit.split(";")
                if len(flds) < 4:
                    fail("%s の %s に「図形名;キャプション;OnAction;幅pt」でない要素: %r"
                         % (SRC_HOME, nm, lit))
                    continue
                cap = flds[1].strip()
                if not cap:
                    fail("%s の %s にキャプションが空の要素がある" % (SRC_HOME, nm))
                    continue
                out.append(cap)
        return out

    return collect(ROW_CONSTS_MAIN), collect(ROW_CONSTS_SUB)


def captions_from_guide_table():
    """(B) build_rpn.py の GUIDE_HOME_BUTTONS。集約表記は展開して返す。"""
    text = read(SRC_BUILD)
    if text is None:
        return []
    m = re.search(r"^GUIDE_HOME_BUTTONS\s*=\s*\[(.*?)^\]", text, re.M | re.S)
    if not m:
        fail("%s に GUIDE_HOME_BUTTONS の定義が見つからない" % SRC_BUILD)
        return []
    out = []
    for cap, _desc in re.findall(r'\(\s*"([^"]*)"\s*,\s*"([^"]*)"\s*\)', m.group(1)):
        out.extend(AGGREGATES.get(cap, [cap]))
    return out


def captions_from_spec():
    """(C) 13章§2.10 の配置表。ボタン列の [～] を拾う。"""
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
        if not line.startswith("| 主要動線") and not line.startswith("| くわしい操作"):
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) < 3:
            fail("%s §2.10 の配置表の行が3列でない: %r" % (SRC_SPEC, line[:40]))
            continue
        rows += 1
        out.extend(c.strip() for c in BRACKET.findall(cells[-1]))
    if rows == 0:
        fail("%s §2.10 に配置表の行(「主要動線」「くわしい操作」)が1行も無い" % SRC_SPEC)
    return out


def captions_from_tour():
    """(D) 初回ツアーの文言(TitleOf / BodyOf)に現れる [～] 表記。"""
    text = read(SRC_GUIDE)
    if text is None:
        return []
    out = []
    found = 0
    for fn in ("TitleOf", "BodyOf"):
        m = re.search(r"^Private Function %s\b.*?^End Function" % fn, text, re.M | re.S)
        if not m:
            fail("%s に %s が見つからない(ツアー文言の置き場所が変わった可能性)"
                 % (SRC_GUIDE, fn))
            continue
        found += 1
        for lit in re.findall(r'"([^"]*)"', m.group(0)):
            out.extend(c.strip() for c in BRACKET.findall(lit))
    if found == 0:
        fail("%s からツアー文言を1本も読めなかった" % SRC_GUIDE)
    return out


def compare_sets(label, impl, other):
    missing = sorted(set(impl) - set(other))
    extra = sorted(set(other) - set(impl))
    for cap in missing:
        fail("%s に実装のキャプションが無い: [%s]" % (label, cap))
    for cap in extra:
        fail("%s にだけあるキャプション(実装に無い): [%s]" % (label, cap))


def main():
    main_caps, sub_caps = captions_from_home()
    impl = main_caps + sub_caps

    # 骨抜き防止: 抽出0件は「一致した」ではなく検問の失敗として扱う。
    if not main_caps:
        fail("%s から主要動線のキャプションを1件も抽出できなかった" % SRC_HOME)
    if not sub_caps:
        fail("%s から「くわしい操作」のキャプションを1件も抽出できなかった" % SRC_HOME)
    if len(impl) != len(set(impl)):
        dup = sorted({c for c in impl if impl.count(c) > 1})
        fail("実装の配置表にキャプションの重複がある: %s" % ", ".join(dup))

    guide = captions_from_guide_table()
    spec = captions_from_spec()
    tour = captions_from_tour()

    if impl:
        if not guide:
            fail("%s の GUIDE_HOME_BUTTONS からキャプションを1件も抽出できなかった" % SRC_BUILD)
        else:
            compare_sets("操作ガイド③のボタン早見表(%s の GUIDE_HOME_BUTTONS)" % SRC_BUILD,
                         impl, guide)
        if not spec:
            fail("%s §2.10 の配置表からキャプションを1件も抽出できなかった" % SRC_SPEC)
        else:
            compare_sets("13章§2.10 の配置表", impl, spec)
        if not tour:
            fail("%s のツアー文言から [～] 表記を1件も抽出できなかった" % SRC_GUIDE)
        else:
            # ツアーは全ボタンに触れないので**部分集合**であることだけを求める。
            for cap in sorted(set(tour) - set(impl)):
                fail("初回ツアーの文言に、実装に無いボタン名がある: [%s] (%s)" % (cap, SRC_GUIDE))

    if errors:
        print("[caption_check] HOMEのボタン名の逐語照合")
        for e in errors:
            print("  NG: " + e)
        print("結果: NG %d件。値源は %s の配置表(UH_ROW_*)であり、"
              "早見表・13章§2.10・ツアー文言をそれに合わせる。" % (len(errors), SRC_HOME))
        return 1

    print("照合先: 操作ガイド早見表 %d件 / 13章§2.10 %d件 / ツアー文言 %d件(部分集合)"
          % (len(guide), len(spec), len(tour)))
    print("OK: 全%d本のキャプションが3系統と一致しました（主要動線%d / くわしい操作%d）"
          % (len(impl), len(main_caps), len(sub_caps)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
