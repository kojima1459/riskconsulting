#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_lo_tests.py - LibreOffice headless によるVBA実行テスト(17章§1 テスト3層の(c))

================================================================================
役割:
    vba_lint.py が「読むだけ」で見つけられる違反を検出するのに対し、本スクリプトは
    実際に LibreOffice(soffice --headless)にVBAソースを読み込ませ、
      モード1: 純ロジックモジュール一式を実行して modTestRunner.RunAllPureTests
               -> ReportText の結果(PASS/FAIL/SKIP)を回収する
      モード2: 全モジュール(Excel依存を含む)を1本ずつ隔離したライブラリに
               読み込み、実行はせず「コンパイルが通るか」だけを確認する
    の2段構えでテストする。**(c)緑はコミット条件であって出荷条件ではない**
    (出荷条件は wintest 実機の層(b))。

移植元: PoC「マイ本棚AI」 tools/run_lo_tests.py(product/nexus-agent)。
    LibreOffice 側の癖(下の技術メモ)と隔離ライブラリ方式はそのまま維持し、
    RPNのsrcレイアウト(src/{core,app,ui,test})と17章§4-1のランナー要件に
    合わせて次を変更した。
      ・PURE_ALLOWLIST を RPN のモジュール名へ入れ替え(未実装は注入をスキップ)
      ・PoC の EXPECTED_PASS_MIN(PASS件数の下限ラチェット)を廃止し、
        **wintest/tests_expected.txt による実行本数の完全一致検査**へ置き換えた。
        下限ラチェットより厳しい条件なので緩和ではない(PASSが増えても減っても
        tests_expected を更新しない限り落ちる)。
      ・EXPECTED_SKIP_MAX は 0 から始める(17章§4-1「SKIP 0件」)。

技術メモ(PoCが実験して確定させた挙動。ここを外すとテストが原因不明で落ちる):
    1. soffice --headless で vnd.sun.star.script: 形式のURIを叩いてマクロを
       実行させるには、-env:UserInstallation で指すプロファイルが「一度でも
       正常に起動を終えたことがある」状態でなければならない。手組みしただけの
       真っさらなプロファイルではスクリプトが黙って実行されない(exit codeは0の
       まま何も起きない)。このためまず --terminate_after_init で一次起動して
       プロファイルを初期化してから、そこへ Basic モジュールを注入する。
    2. モジュール間で Public Type を跨いで参照すると、既定(VBA非互換)の
       StarBasicコンパイラではコンパイルがスタックし呼び出しがハングする。
       各モジュール先頭に "Option VBASupport 1" を追加するとこの問題が解消する。
       元の .bas ファイルは変更せず、LO実行用の一時コピーにだけ付与する。
    3. 1つのライブラリ内に構文エラーを含むモジュールが1つでもあると、その
       ライブラリ内の「どのマクロを呼んでも」呼び出しがハングする。したがって
       モード2ではモジュール1本ずつを専用ライブラリに隔離する。
    4. Excel固有オブジェクトは、実際に「実行が到達」しない限りコンパイルは通る
       (未定義のグローバル識別子の解決は実行時に遅延される)。モード2は対象
       モジュールの中身を実行せず、同居させたダミーの Chk_Driver.Probe() だけを
       呼ぶ(=ライブラリ全体のコンパイルは強制するが本体は実行しない)。
    5. タイムアウトの検知とプロセス後始末は coreutils の `timeout --kill-after`
       に委譲する(自前のkill処理より信頼できた)。
    6. `Public Function Foo(...) As String()`(配列を返す関数)の宣言は、
       Option VBASupport 1 を付けても LibreOffice Basic ではコンパイルが通らない
       (ハングする)。.xba へ変換するときにだけ `As T()` を `As Variant` へ
       機械的に書き換える(_fix_array_return_types)。元の .bas は変更しない。
       実Excel側は元の宣言のままビルドされるため、公開契約は変更していない。
    7. **Public Type はモジュールを跨いで「実行時に」解決されない**(技術メモ2は
       コンパイル停止が消えるだけで、実行時の型解決までは直らない)。定義側とは
       別のモジュールで `Dim c As TCaseCtx` と書くと、LO Basic は型を解決できず
       空のオブジェクトを作り(TypeName=Object / IsObject=True)、最初の
       フィールド代入 `c.case_type = ...` で実行時エラー91
       (Object variable not set)になる。定義モジュール側にダミーのFunctionを
       置いて先に呼び出しても解消しない(遅延ロードの問題ではない)。
       対処: .xba へ変換するときにだけ、**その型を参照していて自分では定義して
       いないモジュールへ Public Type ブロックの写しを前置する**
       (_collect_public_type_blocks / _inject_type_blocks)。同一ライブラリ内で
       同名 Type が複数モジュールに重複定義されても LO は衝突させず、
       写しを持つモジュール同士なら UDT を ByRef/ByVal で受け渡しできることを
       実測で確認した。元の .bas は変更しないので、実Excel側は本物の
       modTypes/modAppTypes を単一の定義元として参照したままである
       (=公開契約も12章§4の依存規則も変えていない)。

使い方:
    python3 tools/run_lo_tests.py                  # モード1+モード2 両方
    python3 tools/run_lo_tests.py --mode pure      # モード1のみ
    python3 tools/run_lo_tests.py --mode compile   # モード2のみ
    python3 tools/run_lo_tests.py --keep-profile   # 一時プロファイルを残す
    exit code: 0 = 全テストPASS+全モジュールコンパイル成功 / 1 = いずれか失敗
================================================================================
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from xml.sax.saxutils import escape as xml_escape

TOOLS_DIR = Path(__file__).resolve().parent
REPO_ROOT = TOOLS_DIR.parent
DEFAULT_SRC_ROOT = REPO_ROOT / "src"
TESTS_EXPECTED_FILE = REPO_ROOT / "wintest" / "tests_expected.txt"

SOFFICE_CANDIDATES = ["/usr/bin/soffice", "soffice"]

# ==============================================================================
# 未実行(SKIP)の天井
# ------------------------------------------------------------------------------
# [SKIP] を貼ったテストは PASS にも FAIL にも現れない。つまり件数を見張らないと
# 「落ちるテストの頭に [SKIP] を貼れば緑に戻せる」。17章§4-1 は出荷条件として
# SKIP 0件を求めているので、ここも0から始める。
# 増やしてよいのは「LOでは原理的に実行できない」と確認できたときだけ。理由を
# テスト側のコメントに書いたうえでこの数字を上げること(=人が1回考えた証跡)。
# 更新履歴:
#   2026-08-29 W0(T-01): 初期値 0。
# ==============================================================================
EXPECTED_SKIP_MAX = 0

# ==============================================================================
# モード1(純ロジック実行)に注入するモジュール(存在するものだけを注入する)
# ------------------------------------------------------------------------------
# ここに無いモジュールをテストから呼ぶと、LOでは実行時エラー12
# (Variable not defined)になり、そのテスト群が「実行されないまま全部PASSに
# 見える」状態になる。テストが新しいモジュールを叩くようになったら必ず足すこと。
# なお「モジュール全体がR4準拠である」必要はない。テストが実際に呼ぶ関数が
# Excel/COMに触れなければよい(実行に到達しない識別子は未解決のままでよい=
# 技術メモ4)。
# ==============================================================================
PURE_ALLOWLIST = [
    # core(W1で移植済)。純ロジックなのでそのまま実行テストにかけられる。
    "modTypes", "modUtil", "modUtilText", "modJsonLite",
    # core のうちExcel/COMに触れる関数を持つが、テストが呼ぶのは純関数だけの
    # モジュール(技術メモ4。W1のG8/G9/G10/G11が叩く)。
    "modConfig", "modLog", "modGatewayRPN", "modGatewayDirect",
    # app の純文字列・純ロジック(W2)。
    "modAppTypes", "modPromptsBlocks", "modPromptsCore", "modPromptsOps",
    "modSchemas", "modValidate", "modValidate2", "modPii",
    # T-35(W3)のHTMLテンプレ系。18章§1の表が「純文字列モジュール」と定めた
    # とおりExcelトークンを1つも持たないので、そのまま実行テストにかけられる
    # (lint の R4 が Excelトークンの混入を機械的に禁止している)。
    "modHtmlTheme", "modHtmlTemplate1", "modHtmlTemplate2", "modHtmlTemplate3",
    "modHtmlTemplate4", "modHtmlTemplate5", "modHtmlTemplate6",
    "modHtmlTemplate7",
    # T-33(W3)。modExportHtml はファイルI/O(ADODB.Stream)とstore経由の読取を
    # 持つが、純組立関数(BuildReportHtml / BuildMetaJson)はどちらにも触れない
    # ため、テストが叩くのはその2本だけ(技術メモ4)。tools/render_report.py も
    # 同じ2本だけを呼んで dist/ のサンプルHTMLを出す。
    "modExportHtml",
    # app のうちExcel/COMに触れる関数を持つが、テストが呼ぶのは14章§6が公開を
    # 宣言した純関数だけのモジュール(技術メモ4。裁定書6 項目6/7・裁定書7 B-6)。
    # modKnowledgeFmt は全体が純文字列(整形と15章§0.7の切詰め)。
    # modPipeline は store/log/LLM経由でExcelに触れるが、テストが叩くのは
    # 14章§6が公開を宣言した判定核16本(純関数)だけ。
    "modKnowledgeFmt", "modCaseStore", "modPipeline",
    # T-25/T-28(裁定書8 B-7/B-10)で 14章§6 が公開を宣言した判定核を持つ3本。
    # いずれもモジュール全体としてはシート・LLMに触れるが、テストが叩くのは
    # 純関数だけ(技術メモ4)。
    #   modInboxStore : 採番 BuildInboxId / ID書式 IsValidInboxId / 状態遷移
    #     CanInboxTransition / E-41 の JudgementError / 関心度の InterestKeyOf・
    #     FmtInterestLine・InterestSummaryOf(FR-17の集計。裁定書9-1)
    #   modPlayOps    : PfSurvivalOf / PfPredTypesOf / PfRefIds / PfFailCodeOf /
    #     CaseIdOfPfLine
    #   modPipeline2  : CritiqueStepOf / ReviseStepOf / NeedsRevision /
    #     CritiqueDigest / DeepOutcomeOf / AdoptRevisionOf(E-36の採用。
    #     裁定書9-2)/ DeepWarningOf / DeepRouteOf
    "modInboxStore", "modPlayOps", "modPipeline2",
    # W7(T-55)。modPipeline の分割先。テストが叩くのは純関数3本
    #   (FinanceBlockText / IncidentsBlockText / FocusLineIdsAttr)だけで、
    #   S*UserText / *Of 系はシートを読むため実行に到達しない(技術メモ4)。
    "modPipeline3",
    # T-27(裁定書8 B-9)で 14章§6 が公開を宣言した壁打ちの純核3本を持つ。
    # モジュール全体は store/受信箱/LLM経由でExcelに触れるが、テストが叩くのは
    #   modSparring : HistoryJoinOf(保存形式 -> 新しい順の";;;"連結)/
    #     TrimHistoryOf(E-44の履歴上限)/CanContinueSparring(E-05(3)を含む
    #     送信可否の fail-closed 判定)
    # だけ(技術メモ4)。
    "modSparring",
    # T-26(裁定書8 B-8)。モジュール全体はシートに触れるが、テストが叩くのは
    # 純関数だけ: BuildJudgeId / IsValidJudgeId / IsValidDecision /
    # IsValidJudgeResult(採番・ID書式・19章§3のdecision enum検証・
    # 13章§2.7のresult enum検証)。
    "modJudgeStore",
    # W4.1(裁定書9)。ui層だがテストが叩くのは19章§3の変換表の純関数3本だけ:
    #   modUICase : EnumPairsCsv(変換表の唯一の値源)/ EnumJa / EnumEn
    # (技術メモ4。モジュールの他の関数はシートに触れるが実行に到達しない)。
    "modUICase",
    # W6第1弾(T-49)。どちらも core の純関数だけで構成され Excel を1つも触らない。
    #   modUIGeom  : 帯・ボタンの並びとカードの高さ・表示時間の算数
    #   modNavText : StripDrFooter / PreviewLines / SplitFieldNotes / JoinFieldNotes
    "modUIGeom", "modNavText",
    # W6.1(裁定書22)。ui層だがテストが叩くのは14章§6が公開を宣言した純関数だけ
    # (技術メモ4。モジュールの他の関数はシート・クリップボードに触れるが実行に
    #  到達しない):
    #   modUIResearch : FillTemplate / PlaceholderTable / PlaceholderKeys(M1)
    #   modUICase6    : AreaTable / AreaKeys / AreaField / HandlerName(M4)
    #   modUINav      : StepRuleOf / StepFor / StepActionOf / StepText /
    #                   StepAnchor(M4。優先順位10行の判定核)
    "modUIResearch", "modUICase6", "modUINav",
    # W6.2(裁定書23追補2)。ui層だがテストが叩くのは MaxWaitText(秒→分の切り上げ)
    # だけ。Excel・シート・モジュール変数のどれにも触れない純関数である
    # (技術メモ4。SetStage 等の他の関数はシートに触るが実行に到達しない)。
    "modUIProgress",
    # test 層。modTestRunner はモード1の入口そのもの。
    "modTestRunner",
    "modTestsPure", "modTestsPure2", "modTestsPure3", "modTestsPure4",
    "modTestsPure5", "modTestsPure6", "modTestsPure7", "modTestsPure8",
    # modTestsPure9: 30,000字契約による modTestsPure8 の分割先。14章§6が公開を
    # 宣言済みで回帰網の無かった純核13本(modInboxStore の採番/ID書式/遷移/
    # 関心度の下請け2本・modJudgeStore の採番/ID書式/enum検証・modPlayOps の
    # PfPredTypesOf/PfRefIds/PfFailCodeOf/CaseIdOfPfLine)を叩く。
    "modTestsPure9",
    # modTestsPure10: T-35(W3)HTMLレポートの純部の契約テスト45本。18章全文と
    # 19章§3だけを根拠に modHtmlTheme.ThemeNames / ThemeCss、modHtmlTemplate1 の
    # BuildDocument / HeadHtml / BodyShellHtml / SectionsJs / RuntimeJs、
    # modHtmlTemplate4.SecRoundUpdateJs を叩く(叩く製品モジュールはいずれも上で
    # 登録済み)。modTestsPure9.RunAll の末尾から呼ぶ。
    "modTestsPure10",
    # modTestsPure11: 30,000字契約(12章§2)による modTestsPure10 の分割先。
    # 16章E-47・18章§5.3のエスケープ契約20本(modUtilText.HtmlSafe / JsStringSafe)
    # と、両群が共有する攻撃素材 DataJsonAttack を持つ。
    # modTestsPure10.RunAll の末尾から呼ぶ。
    "modTestsPure11",
    # modTestsPure12: W4.1(裁定書9)の回帰42本。B5(;置換)/A-6(充足度ラベル)/
    # B2(受信箱判定judge_to)/B4・B7(ファイル名規則)/B9・N1(E-35/E-36警告)。
    # 叩くのは modUtil / modUtilText / modUICase / modInboxStore / modPipeline2 の
    # 純関数だけ。modTestsPure11.RunAll の末尾から呼ぶ。
    "modTestsPure12",
    # modTestsPure13: W6第1弾(1画面ナビ・T-49)の純層18本。11章v3.2 の
    # §3.3.2 / §3.3.5 / §3.3.6 だけを根拠に modNavText.StripDrFooter(7本)/
    # PreviewLines(4本)/ SplitFieldNotes・JoinFieldNotes(7本)を叩く。
    # modTestsPure12.RunAll の末尾から呼ぶ。
    "modTestsPure13",
    # modTestsPure14: W6.1(裁定書22)の純層28本。11章v3.2.1 §3.1.1 / §3.2 /
    # §3.3.7 と 13章§2.11・§2.19、docs/08 の実測だけを根拠に
    # modUIResearch.FillTemplate(11本)/ modUICase6.HandlerName・AreaTable(3本)/
    # modUINav.StepText・StepAnchor(2本)/ StepFor(10本)/
    # modNavText.FitsInRows(2本)を叩く。modTestsPure13.RunAll の末尾から呼ぶ。
    "modTestsPure14",
    # modTestsPure15: W7(裁定書25・17章 T-55)の純層12本。15章 v2.6 の新設・改訂
    #   検証ルール(V-S1-12/13・V-S2-12b/18・V-S3-19/20/21 と V-S1-04・V-S2-16 の
    #   改訂)を叩く。modTestsPure14.RunAll の末尾から呼ぶ。
    "modTestsPure15",
    "modMockLlm", "modMockLlm2", "modMockLlm3",
]

TEMPLATE_PROFILE_DIR = Path(tempfile.gettempdir()) / "rpn_lo_template_profile"

ATTR_LINE_PATTERN = re.compile(r"^\s*Attribute\s+")
# "Function Foo(...) As T()" のみ(引数側の "name() As T" は対象外)を検出する。
ARRAY_RETURN_TYPE_PATTERN = re.compile(
    r"(Function\s+\w+\s*\((?:[^()]|\([^()]*\))*\)\s*)As\s+([A-Za-z_]\w*)\s*\(\s*\)",
    re.IGNORECASE | re.DOTALL,
)

# 技術メモ7: モジュール跨ぎで共有される Public Type ブロック(`Private Type` は
# モジュール内に閉じた型なので写さない。宣言子を省いた `Type Foo` は標準モジュール
# では Public 扱いなので対象に含める)。
PUBLIC_TYPE_BLOCK_PATTERN = re.compile(
    r"^[ \t]*(?:Public[ \t]+)?Type[ \t]+(\w+)[ \t]*$.*?^[ \t]*End[ \t]+Type[ \t]*$",
    re.IGNORECASE | re.MULTILINE | re.DOTALL,
)
# 型ブロックを差し込んでよい位置(先頭の Option 行・空行・行コメントの直後)。
LEADING_NONCODE_PATTERN = re.compile(r"^[ \t]*(?:'|Option[ \t]|$)", re.IGNORECASE)

XBA_TEMPLATE = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    '<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
    '<script:module xmlns:script="http://openoffice.org/2000/script" '
    'script:name="{name}" script:language="StarBasic">{body}\n'
    "</script:module>\n"
)

XLB_TEMPLATE = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    '<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">\n'
    '<library:library xmlns:library="http://openoffice.org/2000/library" '
    'library:name="{libname}" library:readonly="false" library:passwordprotected="false">\n'
    "{elements}\n"
    "</library:library>\n"
)

XLC_TEMPLATE = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    '<!DOCTYPE library:libraries PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "libraries.dtd">\n'
    '<library:libraries xmlns:library="http://openoffice.org/2000/library" '
    'xmlns:xlink="http://www.w3.org/1999/xlink">\n'
    "{libs}\n"
    "</library:libraries>\n"
)

CHK_DRIVER_SRC = (
    "Option Explicit\n\n"
    "Public Function Probe() As Boolean\n"
    "    Probe = True\n"
    "End Function\n"
)

CLS_HEADER_PATTERN = re.compile(
    r"^\s*(VERSION\s+[\d.]+\s+CLASS|BEGIN|MultiUse\s*=.*|END)\s*$", re.IGNORECASE)


def find_soffice() -> str:
    for cand in SOFFICE_CANDIDATES:
        p = shutil.which(cand) or (cand if Path(cand).exists() else None)
        if p:
            return p
    print("[run_lo_tests] soffice が見つかりません。LibreOfficeをインストールしてください。")
    sys.exit(2)


def read_tests_expected() -> int | None:
    """wintest/tests_expected.txt の1行目(10進整数)を読む。無ければ None。"""
    if not TESTS_EXPECTED_FILE.exists():
        return None
    for line in TESTS_EXPECTED_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if re.fullmatch(r"\d+", line):
            return int(line)
        return None
    return None


def strip_attributes(text: str) -> str:
    """Attribute行に加え、.clsファイル先頭のクラスヘッダブロックも除去する。
    これらはVBEのエクスポート形式であってBasicソースではないため、残すと
    LO Basicが構文エラー(ハング)になる。"""
    out = []
    in_header = True
    for l in text.splitlines():
        if in_header:
            if CLS_HEADER_PATTERN.match(l) or ATTR_LINE_PATTERN.match(l) or l.strip() == "":
                continue
            in_header = False
        if not ATTR_LINE_PATTERN.match(l):
            out.append(l)
    return "\n".join(out)


def fix_array_return_types(text: str) -> str:
    """"Function Foo(...) As T()" を "As Variant" へ書き換える(技術メモ6)。"""
    return ARRAY_RETURN_TYPE_PATTERN.sub(lambda m: m.group(1) + "As Variant", text)


def collect_public_type_blocks(all_modules: dict[str, Path]) -> dict[str, str]:
    """src全体から {型名: "Public Type ... End Type"} を集める(技術メモ7)。

    型名が2箇所以上で定義されていたら、どちらを写すかを機械が決められないので
    素通しせず落とす(黙って片方を選ぶと、実Excelとの差が誰にも見えなくなる)。
    """
    blocks: dict[str, str] = {}
    owners: dict[str, list[str]] = {}
    for name in sorted(all_modules):
        text = all_modules[name].read_text(encoding="utf-8", errors="replace")
        for m in PUBLIC_TYPE_BLOCK_PATTERN.finditer(text):
            type_name = m.group(1)
            owners.setdefault(type_name, []).append(name)
            blocks[type_name] = m.group(0).rstrip()
    dup = {t: mods for t, mods in owners.items() if len(mods) > 1}
    if dup:
        for t, mods in sorted(dup.items()):
            print(f"[run_lo_tests] FAIL: Public Type {t} が複数モジュールで定義されています"
                  f"({', '.join(mods)})。写す先を機械が決められません。")
        sys.exit(2)
    return blocks


def module_defines_type(body: str, type_name: str) -> bool:
    return any(m.group(1).lower() == type_name.lower()
               for m in PUBLIC_TYPE_BLOCK_PATTERN.finditer(body))


def code_only(body: str) -> str:
    """行コメントを落とした本文(型名が本当に「使われて」いるかの判定用)。"""
    return "\n".join(l for l in body.splitlines() if not l.lstrip().startswith("'"))


def inject_type_blocks(body: str, type_blocks: dict[str, str]) -> str:
    """自分では定義していないが参照している Public Type の写しを前置する。"""
    code = code_only(body)
    wanted = []
    for type_name, block in type_blocks.items():
        if not re.search(r"\b" + re.escape(type_name) + r"\b", code):
            continue
        if module_defines_type(body, type_name):
            continue
        wanted.append(block)
    if not wanted:
        return body
    lines = body.splitlines()
    at = 0
    while at < len(lines) and LEADING_NONCODE_PATTERN.match(lines[at]):
        at += 1
    header = ["' --- run_lo_tests.py が注入した型定義の写し(技術メモ7)。",
              "'     元の .bas は変更していない。実Excelは定義元1本を参照する。"]
    return "\n".join(lines[:at] + header + wanted + [""] + lines[at:])


def to_module_body(source_text: str, type_blocks: dict[str, str] | None = None) -> str:
    body = strip_attributes(source_text)
    body = fix_array_return_types(body)
    if type_blocks:
        body = inject_type_blocks(body, type_blocks)
    return "Option VBASupport 1\n" + body


def write_module_xba(lib_dir: Path, name: str, source_text: str,
                     type_blocks: dict[str, str] | None = None) -> None:
    body = xml_escape(to_module_body(source_text, type_blocks))
    (lib_dir / f"{name}.xba").write_text(XBA_TEMPLATE.format(name=name, body=body),
                                         encoding="utf-8")


def write_library(profile_dir: Path, lib_name: str, modules: dict[str, str],
                  type_blocks: dict[str, str] | None = None) -> None:
    lib_dir = profile_dir / "user" / "basic" / lib_name
    lib_dir.mkdir(parents=True, exist_ok=True)
    elements = []
    for name, src in modules.items():
        write_module_xba(lib_dir, name, src, type_blocks)
        elements.append(f' <library:element library:name="{name}"/>')
    xlb = XLB_TEMPLATE.format(libname=lib_name, elements="\n".join(elements))
    (lib_dir / "script.xlb").write_text(xlb, encoding="utf-8")


def register_libraries(profile_dir: Path, lib_names: list[str]) -> None:
    libs = [' <library:library library:name="Standard" library:link="false"/>']
    for n in lib_names:
        libs.append(f' <library:library library:name="{n}" library:link="false"/>')
    (profile_dir / "user" / "basic" / "script.xlc").write_text(
        XLC_TEMPLATE.format(libs="\n".join(libs)), encoding="utf-8")


def ensure_template_profile(soffice: str, verbose: bool) -> Path:
    """一度だけ soffice を一次起動させて雛形プロファイルを作る(技術メモ1)。"""
    marker = TEMPLATE_PROFILE_DIR / "user" / "basic" / "Standard" / "script.xlb"
    if marker.exists():
        return TEMPLATE_PROFILE_DIR
    if verbose:
        print(f"[run_lo_tests] 雛形プロファイルを初期化中: {TEMPLATE_PROFILE_DIR}")
    if TEMPLATE_PROFILE_DIR.exists():
        shutil.rmtree(TEMPLATE_PROFILE_DIR)
    subprocess.run([
        "timeout", "--kill-after=5", "60",
        soffice, "--headless", "--invisible", "--nologo", "--norestore",
        f"-env:UserInstallation=file://{TEMPLATE_PROFILE_DIR}",
        "--terminate_after_init",
    ], capture_output=True, text=True)
    if not marker.exists():
        print("[run_lo_tests] 雛形プロファイルの初期化に失敗しました(soffice起動不可の可能性)。")
        sys.exit(2)
    return TEMPLATE_PROFILE_DIR


def fresh_profile_copy(template: Path, dest: Path) -> None:
    shutil.copytree(template, dest)


def run_uri(soffice: str, profile_dir: Path, uri: str, timeout_sec: int):
    proc = subprocess.run([
        "timeout", "--kill-after=5", str(timeout_sec),
        soffice, "--headless", "--invisible", "--nologo", "--norestore",
        f"-env:UserInstallation=file://{profile_dir}",
        uri,
    ], capture_output=True, text=True)
    return proc.returncode, proc.stdout, proc.stderr


def discover_modules(src_root: Path) -> list[tuple[str, Path]]:
    """(モジュール名, パス) のリスト。名前は Attribute VB_Name があればそれ。"""
    out = []
    for path in sorted(list(src_root.rglob("*.bas")) + list(src_root.rglob("*.cls"))):
        text = path.read_text(encoding="utf-8", errors="replace")
        name = path.stem
        m = re.search(r'^\s*Attribute\s+VB_Name\s*=\s*"([^"]*)"', text, re.MULTILINE)
        if m:
            name = m.group(1)
        out.append((name, path))
    return out


# ==============================================================================
# モード1: 純ロジック実行(RunAllPureTests -> ReportText)
# ==============================================================================
def run_pure_mode(soffice: str, template: Path, all_modules: dict[str, Path],
                  work_dir: Path, timeout_sec: int, verbose: bool,
                  type_blocks: dict[str, str] | None = None):
    print("=" * 78)
    print("モード1: LibreOffice上で純ロジックテスト(modTestRunner.RunAllPureTests)を実行")
    print("=" * 78)

    pure_srcs: dict[str, str] = {}
    missing = []
    for name in PURE_ALLOWLIST:
        path = all_modules.get(name)
        if path is None:
            missing.append(name)
            continue
        pure_srcs[name] = path.read_text(encoding="utf-8", errors="replace")

    if missing:
        print(f"  (未実装のため注入をスキップ: {', '.join(missing)})")

    if "modTestRunner" not in pure_srcs:
        print("  modTestRunner.bas が見つからないためモード1を実行できません。")
        return False, "modTestRunner.bas not found"

    expected = read_tests_expected()
    if expected is None:
        print(f"  FAIL: {TESTS_EXPECTED_FILE} が読めません"
              "(1行目に10進整数のみを書いてください。17章§4-1)")
        return False, "tests_expected unreadable"
    print(f"  tests_expected = {expected} ({TESTS_EXPECTED_FILE})")

    out_path = work_dir / "pure_result.txt"
    if out_path.exists():
        out_path.unlink()

    test_main_src = (
        "Option Explicit\n\n"
        "Sub Main\n"
        "    On Error Resume Next\n"
        "    Err.Clear\n"
        f"    modTestRunner.SetExpectedCount {expected}\n"
        "    modTestRunner.RunAllPureTests\n"
        "    Dim runErr As String\n"
        "    If Err.Number <> 0 Then\n"
        '        runErr = "RUNNER_ERROR " & Err.Number & ": " & Err.Description\n'
        "        Err.Clear\n"
        "    End If\n"
        "    On Error GoTo 0\n\n"
        "    Dim iFile As Integer\n"
        "    iFile = FreeFile\n"
        f'    Open "{out_path.as_posix()}" For Output As #iFile\n'
        "    Print #iFile, modTestRunner.ReportText()\n"
        "    If Len(runErr) > 0 Then Print #iFile, runErr\n"
        "    Close #iFile\n"
        "End Sub\n"
    )

    profile_dir = work_dir / "profile_pure"
    fresh_profile_copy(template, profile_dir)
    modules = dict(pure_srcs)
    modules["TestMain"] = test_main_src
    write_library(profile_dir, "RpnPureRun", modules, type_blocks)
    register_libraries(profile_dir, ["RpnPureRun"])

    uri = "vnd.sun.star.script:RpnPureRun.TestMain.Main?language=Basic&location=application"
    rc, out, err = run_uri(soffice, profile_dir, uri, timeout_sec)

    if not out_path.exists():
        msg = f"実行結果ファイルが生成されませんでした(soffice exit={rc}, timeout={timeout_sec}s)"
        print(f"  FAIL: {msg}")
        if verbose:
            print(f"    stdout: {out.strip()}")
            print(f"    stderr: {err.strip()}")
        return False, msg

    report = out_path.read_text(encoding="utf-8", errors="replace").strip()
    print(report if report else "(空の結果)")

    m = re.search(r"PASS\s+(\d+)\s*/\s*FAIL\s+(\d+)\s*/\s*SKIP\s+(\d+)", report)
    if m:
        pass_count, fail_count, skip_count = (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    else:
        # 見出し行が読めない = 集計そのものが壊れている。合格側へ倒さない。
        pass_count = fail_count = skip_count = None
    has_runner_error = "RUNNER_ERROR" in report

    ok = (fail_count == 0) and not has_runner_error

    if pass_count is None or skip_count is None:
        print("  FAIL: 集計行(PASS n / FAIL m / SKIP k)を読み取れませんでした。")
        ok = False
    else:
        executed = pass_count + fail_count
        if skip_count > EXPECTED_SKIP_MAX:
            print(f"  FAIL: 未実行(SKIP)が {skip_count} 件で、天井 {EXPECTED_SKIP_MAX} 件を超えました。")
            print("        [SKIP] を貼るとテストは PASS にも FAIL にも現れません。")
            print("        増やしてよいのは『LOでは原理的に実行できない』と確認できたときだけです。")
            print("        理由をテスト側のコメントに書いたうえで EXPECTED_SKIP_MAX を更新してください。")
            ok = False
        if executed != expected:
            print(f"  FAIL: 実行本数が {executed} 件で tests_expected({expected})と一致しません。")
            print("        テストを増減したら wintest/tests_expected.txt を同時に更新してください。")
            ok = False
        if ok:
            print(f"  ベースライン照合 OK: 実行本数 {executed} = tests_expected {expected} / "
                  f"SKIP {skip_count} <= {EXPECTED_SKIP_MAX}")

    return ok, report


# ==============================================================================
# モード2: 全モジュールのコンパイルチェック(実行はしない)
# ==============================================================================
def safe_lib_name(module_name: str) -> str:
    return "Chk_" + re.sub(r"[^A-Za-z0-9_]", "_", module_name)


def run_compile_mode(soffice: str, template: Path, all_modules: dict[str, Path],
                     work_dir: Path, timeout_sec: int, verbose: bool,
                     type_blocks: dict[str, str] | None = None):
    print("\n" + "=" * 78)
    print("モード2: 全モジュールの構文コンパイルチェック(実行はしない)")
    print("=" * 78)

    modtypes_path = all_modules.get("modTypes")
    modtypes_src = (modtypes_path.read_text(encoding="utf-8", errors="replace")
                    if modtypes_path else None)

    results: list[tuple[str, bool, str]] = []
    all_ok = True

    for name, path in sorted(all_modules.items()):
        target_src = path.read_text(encoding="utf-8", errors="replace")
        lib_name = safe_lib_name(name)

        modules_for_lib: dict[str, str] = {}
        if name != "modTypes" and modtypes_src is not None:
            modules_for_lib["modTypes"] = modtypes_src
        modules_for_lib[name] = target_src
        modules_for_lib["Chk_Driver"] = CHK_DRIVER_SRC

        profile_dir = work_dir / f"profile_{lib_name}"
        fresh_profile_copy(template, profile_dir)
        write_library(profile_dir, lib_name, modules_for_lib, type_blocks)
        register_libraries(profile_dir, [lib_name])

        uri = f"vnd.sun.star.script:{lib_name}.Chk_Driver.Probe?language=Basic&location=application"
        t0 = time.time()
        rc, _out, _err = run_uri(soffice, profile_dir, uri, timeout_sec)
        elapsed = time.time() - t0

        ok = (rc == 0)
        detail = f"exit={rc} ({elapsed:.1f}s)"
        if rc == 124:
            detail = f"タイムアウト({timeout_sec}s) - 構文エラーの疑い"
        results.append((name, ok, detail))
        all_ok = all_ok and ok

        print(f"  {'PASS' if ok else 'FAIL':<4} {name:<24} {detail}")
        shutil.rmtree(profile_dir, ignore_errors=True)

    if not results:
        print("  (対象モジュールが1本もありません)")
    return all_ok, results


def main() -> int:
    parser = argparse.ArgumentParser(description="リスク提案ナビ LibreOffice実行テスト")
    parser.add_argument("--path", type=str, default=str(DEFAULT_SRC_ROOT),
                        help="対象src(既定: <repo>/src)")
    parser.add_argument("--mode", choices=["pure", "compile", "all"], default="all")
    parser.add_argument("--pure-timeout", type=int, default=120,
                        help="モード1のタイムアウト秒(既定120)")
    parser.add_argument("--compile-timeout", type=int, default=20,
                        help="モード2の1モジュールあたりタイムアウト秒(既定20)")
    parser.add_argument("--keep-profile", action="store_true",
                        help="一時プロファイルを削除せず残す(デバッグ用)")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    src_root = Path(args.path).resolve()
    if not src_root.exists():
        print(f"[run_lo_tests] 対象ディレクトリが存在しません: {src_root}")
        return 2

    soffice = find_soffice()
    template = ensure_template_profile(soffice, args.verbose)

    all_modules: dict[str, Path] = dict(discover_modules(src_root))
    print(f"[run_lo_tests] soffice={soffice}")
    print(f"[run_lo_tests] 対象モジュール数: {len(all_modules)} (in {src_root})")

    type_blocks = collect_public_type_blocks(all_modules)
    if type_blocks:
        print(f"[run_lo_tests] 跨ぎ参照へ写す Public Type(技術メモ7): "
              f"{', '.join(sorted(type_blocks))}")

    work_dir = Path(tempfile.mkdtemp(prefix="rpn_lo_run_"))
    overall_ok = True
    try:
        if args.mode in ("pure", "all"):
            ok, _report = run_pure_mode(soffice, template, all_modules, work_dir,
                                        args.pure_timeout, args.verbose, type_blocks)
            overall_ok = overall_ok and ok
        if args.mode in ("compile", "all"):
            ok, _results = run_compile_mode(soffice, template, all_modules, work_dir,
                                            args.compile_timeout, args.verbose, type_blocks)
            overall_ok = overall_ok and ok
    finally:
        if args.keep_profile:
            print(f"\n[run_lo_tests] --keep-profile 指定のため一時ディレクトリを残します: {work_dir}")
        else:
            shutil.rmtree(work_dir, ignore_errors=True)

    print("\n" + "-" * 78)
    print("結果: OK(exit code 0)" if overall_ok else "結果: NG(exit code 1)")
    print("-" * 78)
    return 0 if overall_ok else 1


if __name__ == "__main__":
    sys.exit(main())
