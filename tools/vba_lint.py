#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
vba_lint.py - リスク提案ナビ(RPN) VBAソースの静的Lint

================================================================================
位置づけ(17章§1 テスト3層の (c) 実機前静的):
    src/ 配下の全 .bas / .cls を対象に、実行せずに読める範囲の契約違反・
    危険なコードパターンを機械的に検出する。**(c)緑はコミット条件であって
    出荷条件ではない**(出荷条件は wintest 実機の層(b))。

移植元: PoC「マイ本棚AI」 tools/vba_lint.py(product/nexus-agent)。
    実機事故の再発防止として積み上がった検査群はそのまま維持し、
    RPNに存在しない規則(opt層R2・modUiLock・modSkin.ShowToast 等の
    PoC固有の例外表)だけを削除・読み替えた。**既存規則の緩和はしていない**。
    RPN向けに強化した点は各検査の docstring に「RPNでの変更」として明記する。

    削除した検査の記録:
    ・PoC の check_safeleft_warning(SafeLeft 迂回のセル書込を WARN で促す助言)は
      非移植とした。ERROR級の check_cell_write_guard(16章NFR-S7①: 外部由来テキストの
      セル書込は SetCellSafe 以外を禁止)が上位互換(同種の書込経路を WARN より強く
      全て捕捉)であり、弱い WARN を重ねる意味がないため。**緩和ではなく吸収**。

RPNの規約(12章§2・§4 / 16章NFR-S7 / 17章T-46):
    R1 依存方向 ui -> app -> core の一方向(core は app/ui を参照しない)
    R3 Application.Run は modGatewayRPN のみ
    R4 Excelトークン(Worksheets / Range( / Application. / ThisWorkbook /
       MsgBox / ActiveSheet)は ui層 と「指定モジュール」のみ
    1モジュール30,000字契約(28,000字で警告)
    CP932安全(VBEはCP932でソースを保持するため、CP932外文字は "?" 化する)
    NFR-S7 外部由来テキストの書き込み口の一元化(セル=SetCellSafe /
       HTML=HtmlSafe・JsStringSafe)。17章T-46①を(c)層でも毎コミット走らせる

使い方:
    python3 tools/vba_lint.py                 # <repo>/src 配下を検査
    python3 tools/vba_lint.py --path <dir>    # 検査対象を変更(主にテスト用)
    exit code: 0 = ERROR 0件 / 1 = ERROR 1件以上
================================================================================
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

TOOLS_DIR = Path(__file__).resolve().parent
REPO_ROOT = TOOLS_DIR.parent
DEFAULT_SRC_ROOT = REPO_ROOT / "src"

MAX_MODULE_CHARS = 30000
# 残り2,000字を切ったら警告する(バグ修正1件ぶんの余裕がある状態を保つため)。
MODULE_WARN_CHARS = 28000

# ==============================================================================
# 12章§2 モジュール一覧(レイヤ構成図の全モジュール)
# ------------------------------------------------------------------------------
# ここに無いモジュール名が src/ に現れたら WARN で申告する(仕様に無いモジュール
# が黙って増えるのを防ぐ)。分割モジュール(modTestsPure1..n / modSchemas1..n /
# modHtmlTemplate1..n)は末尾数字を許すため正規表現側で受ける。
# ==============================================================================
MODULE_REGISTRY = {
    # ---- ui 層 ----
    "modBoot", "modUIHome", "modUICase", "modUIInbox", "modUISparring",
    "modUIProgress",
    # ---- app 層 ----
    "modPipeline", "modPlayOps", "modSparring", "modCaseStore", "modInboxStore",
    "modJudgeStore", "modKnowledge", "modValidate", "modCompanyFile", "modPii",
    "modExportHtml", "modExportPpt", "modExportHearing", "modAppTypes",
    "modPromptsCore", "modPromptsBlocks", "modPromptsOps", "modSchemas",
    "modHtmlTheme",
    # ---- core 層 ----
    "modGatewayRPN", "modGatewayDirect", "modJsonLite", "modConfig", "modLog",
    "modUtil", "modUtilText", "modTypes",
    # ---- test 層 ----
    "modTestRunner", "modTestsExcel", "modMockLlm",
}
# 分割される可能性のあるモジュール名(末尾に1以上の数字が付く)。
# modMockLlm1..n は 12章§2(v2.4.1)が 30,000字契約による分割を明記している。
MODULE_REGISTRY_PREFIXES = re.compile(
    r"^(modTestsPure|modSchemas|modHtmlTemplate|modPromptsCore|modPromptsBlocks"
    r"|modPromptsOps|modMockLlm)\d*$"
)

# ==============================================================================
# 公開契約表: モジュール名 -> {"closed": bool, "required": [Public名...]}
# ------------------------------------------------------------------------------
# 正は 14章§6(公開関数シグネチャ)＋12章§2(モジュール一覧)＋15章§10.2
# (節⇔関数名対応表)。**ここに載っていないモジュールは契約チェックの対象外**、
# **ファイルがまだ無いモジュールはSKIP(エラーにしない)**。W0時点では全モジュール
# 未実装なので、この表は「実装され次第そのモジュールに効く」検査になる。
#
# closed の判断:
#   closed=True  : 章側が「これで全部」と明言している表だけ(15章§10.2の
#                  Block* 7関数 / 本Lintが定義する modTestRunner)。
#   closed=False : それ以外。14章§6は主要な公開口の宣言であって「Publicの
#                  全数表」ではない(modUtilText の SafeLeft/ElapsedMsSince 等、
#                  12章の移植対応表側にしか出てこない関数がある)ため、
#                  追加Publicを違反にしない。required は全部無いとERROR。
#
# Phase 1.5 の関数(SchemaWT / SchemaFG / CheckWT / CheckFG / GeneratePpt)は
# required に入れない(Phase 1 の実装でERRORになるため)。
# ==============================================================================
CONTRACT: dict[str, dict] = {
    # ---- core 層 ----
    "modGatewayRPN": {
        "closed": False,
        "required": ["CallStep", "CallChat", "RibbonAvailable", "RunLimitCheck"],
    },
    "modJsonLite": {
        "closed": False,
        "required": [
            "ExtractJsonBlock", "GetStr", "GetLong", "GetBoolJ",
            "GetArrayItems", "EscapeJsonStr", "UnescapeJsonStr",
        ],
    },
    "modUtilText": {
        "closed": False,
        # SetCellSafe は 16章NFR-S7① / 17章T-10 の新設関数(外部由来テキストの
        # 唯一のセル書き込み口)。JsStringSafe/HtmlSafe は NFR-S7③。
        "required": [
            "SetCellSafe", "JsStringSafe", "HtmlSafe", "SanitizeInput",
            "SanitizeFileName", "NormalizeForHash", "Fnv1a64Hex",
        ],
    },
    # modTypes は core層の汎用型(Public Type)のみを持つ(12章§2・§4。TCaseCtx等の
    # ドメイン型は app層 modAppTypes へ移設)。14章§6は Public Function シグネチャを
    # 定義していない(型モジュール)ため required は空。CONTRACT に載せるのは完全性
    # 自己検査(MODULE_REGISTRY⇔CONTRACT)を満たすため。closed=False で追加 Public は許容。
    "modTypes": {"closed": False, "required": []},
    # ---- app 層 ----
    "modValidate": {
        "closed": False,
        "required": [
            "NormalizeLlmJson", "CheckS1", "CheckS2", "CheckS3", "CheckS4",
            "CheckS2C", "CheckS3C", "CheckPF",
        ],
    },
    "modKnowledge": {
        "closed": False,
        "required": [
            "LoadKnowledge", "RiskLibFor", "MenusSummaryFor", "MenusFor",
            "LinesText", "CasesFor", "SchemesFor", "PatternsText", "RulesText",
            "ResearchingText", "MechsText", "LastInjectedIds", "ResetInjectedIds",
            "MenuIdExists", "LineIdExists", "SchemeIdExists", "CaseLibIdExists",
            "PatternIdExists", "AppendServiceGap",
        ],
    },
    "modPromptsCore": {
        "closed": False,
        "required": [
            "BuildS1System", "BuildS1User", "BuildS2System", "BuildS2User",
            "BuildS3System", "BuildS3User", "BuildS4System", "BuildS4User",
        ],
    },
    "modPromptsBlocks": {
        # 15章§10.2 末尾が「Block* の7関数」と明言しているため closed=True。
        "closed": True,
        "required": [
            "BlockCtx", "BlockRenewalS1", "BlockRenewalS2", "BlockRenewalS3",
            "BlockGuard", "BlockS4Proposal", "BlockS4Alliance",
        ],
    },
    "modPromptsOps": {
        "closed": False,
        "required": [
            "BuildS2CriticSystem", "BuildS2CriticUser", "BuildS3CriticSystem",
            "BuildS3CriticUser", "ReviseSuffix", "BuildSparringSystem",
            "BuildPFSystem", "BuildPFUser", "RepairSuffix",
        ],
    },
    "modSchemas": {
        "closed": False,
        "required": [
            "SchemaS1", "SchemaS2", "SchemaS3", "SchemaS4",
            "SchemaS2C", "SchemaS3C", "SchemaPF",
        ],
    },
    "modPipeline": {"closed": False, "required": ["RunAll", "RunStep"]},
    "modPlayOps": {"closed": False, "required": ["RunPreflight"]},
    "modCaseStore": {
        "closed": False,
        "required": [
            "NewCase", "SaveData", "LoadData", "ResolveStepJson", "SetStatus",
            "InvalidateDownstream", "RepairStates", "FreezeRound",
        ],
    },
    "modInboxStore": {
        "closed": False,
        "required": ["NewInboxItem", "SetInboxJudgement"],
    },
    "modJudgeStore": {"closed": False, "required": ["NewJudgement"]},
    "modExportHtml": {"closed": False, "required": ["GenerateHtmlReport"]},
    "modExportHearing": {"closed": False, "required": ["BuildHearingSheet"]},
    # modExportPpt: 14章§6の GeneratePpt は **Phase 1.5**(§6の注記・§1の表)。
    # Phase 1 実装で required に入れると未実装ERRORになるため required は空にする
    # (CONTRACTの方針: Phase 1.5関数は required に入れない)。closed=False。
    "modExportPpt": {"closed": False, "required": []},
    # modAppTypes: TCaseCtx 等のドメイン型(Public Type)を持つ app層モジュール
    # (14章§6 の TCaseCtx 定義)。Public Function シグネチャは章に無いため required は空。
    "modAppTypes": {"closed": False, "required": []},
    # modPii: 保存・外部送信・レポート出力前のPII走査本体(16章 E-05・12章§2/§4)。
    # 14章§6は走査本体の Public Function シグネチャを固定していない(責務のみ規定)ため
    # required は空。呼び出し側(modUICase/modUIInbox/modSparring/modJudgeStore/
    # modExportHtml/modCompanyFile)が前段で必ず通す関係だけが規定される。
    "modPii": {"closed": False, "required": []},
    # modHtmlTemplate1: 18章§4.4 の分割表で「基底モジュール」が持つと明記された5関数。
    # (SecXxxJs等のセクション関数は modHtmlTemplate2..n 側にあり流動的なので、
    #  分割後モジュールには固定契約を課さない=CONTRACTに載せない。closed=False)。
    "modHtmlTemplate1": {
        "closed": False,
        "required": ["BuildDocument", "HeadHtml", "BodyShellHtml", "SectionsJs",
                     "RuntimeJs"],
    },
    # modHtmlTheme: 18章§5.2 の純文字列モジュール(テーマCSSの :root{...} だけを返す)。
    "modHtmlTheme": {"closed": False, "required": ["ThemeNames", "ThemeCss"]},
    # ---- ui 層 ----
    "modBoot": {"closed": False, "required": ["Boot"]},
    "modUIProgress": {
        "closed": False,
        "required": ["SetStage", "TryEnterUiLock", "ExitUiLock", "ParkFocus"],
    },
    # ---- test 層 ----
    "modTestRunner": {
        # 本Lintと17章§4-1のランナー要件で公開口を確定させているため closed=True。
        "closed": True,
        "required": [
            "ResetTests", "Check", "Failures", "ReportText", "RunAllPureTests",
            "SetExpectedCount",
        ],
    },
    # modTestsExcel: 14章§6のtest層契約(層(b)=実Excel E2Eスモークの入口。17章T-47)。
    "modTestsExcel": {"closed": False, "required": ["RunAllExcelTests"]},
}

# ==============================================================================
# R4: Excelトークン規則(12章§2)
# ------------------------------------------------------------------------------
# 「ui層以外に Worksheets/Range/MsgBox 等のExcelトークン禁止(modConfig/
#   modCaseStore等のstore系モジュール内は自身の責務範囲で可)」
#
# RPNでの変更(PoCからの読み替え・**強化**):
#   PoC は「純ロジックと決めたモジュールの一覧(PURE_LOGIC_MODULES)」だけを
#   検査する許可リスト方式だった。RPNの R4 本文は逆に「ui層以外は原則禁止」
#   なので、**既定禁止＋明示許可**へ反転する(検査範囲が広がる=緩和ではない)。
#   ここへ足すときは必ず理由を1行書くこと。どのモジュールがExcelに触れるかが
#   1箇所で分かる状態を保つ。
# ==============================================================================
R4_EXCEL_ALLOWED_MODULES = {
    # modGatewayRPN: R3 の Application.Run("ChatGPT", ...) の唯一の置き場
    #   (14章§2。core層だがリボン呼び出しの責務上ここだけは Application. が要る)。
    "modGatewayRPN",
    # modConfig: config シートの読み書きが責務そのもの(13章§2.3)。
    "modConfig",
    # modLog: err_log / usage_log / run_log シートへの記録が責務(13章§2.4)。
    "modLog",
    # 以下 store 系。案件・受信箱・判断台帳・ナレッジブックのシートI/Oが責務
    # (12章§2 の R4 但し書き「store系モジュール内は自身の責務範囲で可」)。
    "modCaseStore", "modInboxStore", "modJudgeStore", "modKnowledge",
    # modCompanyFile: 企業ドシエファイル(1社1.xlsx)の書出・取込(13章§2.8)。
    "modCompanyFile",
    # modExportHearing: ヒアリングシート(本体ブック内のシート)を組み立てる。
    "modExportHearing",
    # modUtilText: SetCellSafe 内のセル書込に限る(12章§4 v2.4.1・16章NFR-S7(1))。
    #   「外部由来テキストのセル書込口はこの1関数」と定めた以上、その関数本体だけは
    #   セルに触れざるを得ない。純変換部は SanitizeForCell として分離してあり、
    #   そちらはExcel非依存=層(a)でテストする。
    "modUtilText",
    # ※ modExportHtml / modExportPpt は「案件JSONを受け取って外部ファイルを
    #   書く」役なので既定では許可しない。シートを直接読む必要が出たら
    #   modCaseStore 経由にするか、理由を添えてここへ足すこと。
    # ※ modPromptsCore/Blocks/Ops・modSchemas・modHtmlTemplate*・modHtmlTheme は
    #   12章§4により純文字列モジュール。絶対にここへ足さない。
}

# test層は R4 の適用外(全層を参照してよい)。ただし modTestRunner と
# modTestsPure* は LibreOffice 実行テスト(17章§1 (c))の本体であり、
# Excelトークンが入った瞬間に (c) が丸ごと動かなくなるため純ロジックを強制する。
R4_PURE_REQUIRED_TEST = re.compile(r"^(modTestRunner|modTestsPure\d*)$")

FORBIDDEN_TOKEN_PATTERNS = [
    (re.compile(r"\bWorksheets\b"), "Worksheets"),
    (re.compile(r"\bRange\s*\("), "Range("),
    (re.compile(r"\bApplication\."), "Application."),
    (re.compile(r"\bThisWorkbook\b"), "ThisWorkbook"),
    (re.compile(r"\bMsgBox\b"), "MsgBox"),
    (re.compile(r"\bActiveSheet\b"), "ActiveSheet"),
]

# ==============================================================================
# R3: Application.Run 第1引数リテラルのホワイトリスト(12章§2・14章§2)
# ------------------------------------------------------------------------------
# RPNでの変更: PoC の "GetEmbeddings"(ベクトル検索。RPNでは使わない)を削除し、
#   許可モジュールを modGateway -> modGatewayRPN、modFeatures(opt機能ディス
#   パッチ。RPNに存在しない)を削除した。**許可の幅は狭くなっている**。
# ==============================================================================
RUN_LITERAL_WHITELIST_EXACT = {"ChatGPT", "LimitCheck", "modBoot.Boot"}
# ChatGPT / LimitCheck の直接 Application.Run は modGatewayRPN 内のみ許可(R3本文)。
RUN_LITERAL_GATEWAY_ONLY = {"ChatGPT", "LimitCheck"}
RUN_GATEWAY_MODULE = "modGatewayRPN"
# 変数経由(非リテラル)の Application.Run はこのモジュールのみ許可。
RUN_VARIABLE_ALLOWED_MODULES = {"modGatewayRPN"}

# ==============================================================================
# 16章 NFR-S7 / 17章 T-46①: 外部由来テキストの書き込み口の一元化
# ------------------------------------------------------------------------------
# 検出パターンは16章NFR-S7・17章T-46①の逐語。除外は modUtilText(防御関数の
# 定義本体)のみ。定数・ヘッダの書込は行末に ' SAFE:const の注記を必須とし、
# **注記の無いヒットのみFAIL**とする。
#
# 精度の但し書き(見逃しではなく誤検知の抑制):
#   `If ws.Cells(r, 1).Value = "" Then` のような**比較**は書き込みではない。
#   そこで「文の先頭が制御構文キーワードでなく、かつトップレベルの代入である」
#   文だけを対象にし、その**左辺**にパターンを当てる。書き込み(代入)は必ず
#   この形になるので検出範囲は落ちない。
# ==============================================================================
CELL_WRITE_PATTERNS = [
    re.compile(r"\.Value2?\s*="),
    re.compile(r"\.Formula\w*\s*="),
    re.compile(r"(Cells|Range|Offset)\([^)]*\)\s*="),
]
CELL_WRITE_EXEMPT_MODULES = {"modUtilText"}
SAFE_CONST_MARKER = "' SAFE:const"
CELL_WRITE_SAFE_CALL = re.compile(r"\bSetCellSafe\s*\(", re.IGNORECASE)

# HTML/JS 埋め込み(NFR-S7③)。対象は HTML を組み立てるモジュールのみ。
HTML_BUILDER_MODULE = re.compile(r"^(modExportHtml|modHtmlTemplate\d*|modHtmlTheme)$")
HTML_SAFE_CALL = re.compile(r"\b(HtmlSafe|JsStringSafe)\s*\(", re.IGNORECASE)
SAFE_HTML_MARKER = "' SAFE:html"
# 連結してよい素のトークン(VBAの組み込み定数と数値リテラル)。
HTML_SAFE_BARE_TOKENS = {
    "vblf", "vbcrlf", "vbcr", "vbtab", "vbnullstring", "true", "false",
}

# ==============================================================================
# 12章§4: core層のコードに製品固有の語彙を書かない
# ------------------------------------------------------------------------------
# 受入条件は「core層の全ファイルを grep して製品名・シート名の出現がゼロ」
# (17章 T-10)。W0時点で機械化しておき、W1のcore移植で違反が生えないようにする。
# ==============================================================================
CORE_PRODUCT_VOCAB = [
    "リスク提案ナビ",
    "ナレッジブック",
    "案件一覧", "案件入力", "case_data", "受信箱", "判断台帳", "フィードバック",
    "ヒアリングシート", "壁打ち",
    "S1_企業プロファイル", "S2_リスク仮説", "S3_提案", "S4_骨子",
    "はじめにお読みください",
]

# ==============================================================================
# R2(opt層の直接参照禁止)は RPN に opt 層が存在しないため**規則ごと削除**した。
# PoC の modUiLock / modSkin.ShowToast / modProgressBar / modWorkExcel の
# R1例外表も、対応するモジュールが RPN に存在しないため削除した。
# ==============================================================================

# ==============================================================================
# R1例外表(製品コード -> test層。名指しペアのみ。裁定書5 項目7)
# ------------------------------------------------------------------------------
# 原則は「どの層からも test層を参照してはならない」だが、mockトランスポートの
# 参照だけは【仕様が明示的に命じたもの】であり例外として1件だけ許可する。
#
# 根拠(仕様書4箇所):
#   12章§2 移植対応表「modMockLlm(src/test/)。modGatewayRPN のmock分岐が呼ぶ
#           唯一の相手」/ 14章§4(a) 同旨 / 15章§8 冒頭 / 17章 T-14
# 禁止理由が当てはまらないこと: 本ルールの根拠は「製品コードがテストに依存すると
#   配布物からテストを外せなくなる」だが、modMockLlm は build/modules.json に
#   登録され vba_src へ焼き込まれて【配布物に同梱される本体内mockトランスポート】
#   であり、外す対象ではない。
# 例外の粒度: (参照元モジュール, 参照先モジュール, メンバ) の完全一致のみ。
#   他の core->test 参照、および modGatewayRPN から modMockLlm の別メンバへの
#   参照は引き続き ERROR。緩和を1行増やすには司令塔の裁定を要する。
# ==============================================================================
R1_TEST_LAYER_EXCEPTIONS = {
    ("modGatewayRPN", "modMockLlm", "MockResponse"),
}

# 型落ち検出: Dim/Static文
DIM_STMT_PATTERN = re.compile(r"^(Dim|Static)\s+(.*)$", re.IGNORECASE)
AS_KEYWORD_PATTERN = re.compile(r"\bAs\b", re.IGNORECASE)
AS_INTEGER_PATTERN = re.compile(r"\bAs\s+Integer\b", re.IGNORECASE)

# ReDim x(0 To -1) / (5 To 2) 等の「上限が負」の負範囲(実機VBAで実行時エラー9)。
NEGATIVE_REDIM_PATTERN = re.compile(
    r"\bReDim\b[^(]*\([^)]*\bTo\s+-\d+\s*\)", re.IGNORECASE
)

ATTRIBUTE_VBNAME_PATTERN = re.compile(
    r'^\s*Attribute\s+VB_Name\s*=\s*"([^"]*)"', re.IGNORECASE)
OPTION_EXPLICIT_PATTERN = re.compile(r"^\s*Option\s+Explicit\s*$", re.IGNORECASE)

PUB_SUB_PATTERN = re.compile(r"^Public\s+Sub\s+([A-Za-z_]\w*)", re.IGNORECASE)
PUB_FUNC_PATTERN = re.compile(r"^Public\s+Function\s+([A-Za-z_]\w*)", re.IGNORECASE)
PUB_CONST_PATTERN = re.compile(r"^Public\s+Const\s+([A-Za-z_]\w*)", re.IGNORECASE)
PUB_TYPE_PATTERN = re.compile(r"^Public\s+Type\s+([A-Za-z_]\w*)", re.IGNORECASE)

DOTTED_REF_PATTERN = re.compile(r"\b([A-Za-z_]\w*)\.([A-Za-z_]\w*)")
# 注意: ThisWorkbook はここに含めない。"ThisWorkbook.Worksheets" 等はExcel組込の
# オブジェクトモデルであって我々が定義したPublicメンバーではないため、
# モジュール形状の判定対象は mod* のみに絞る(RPNに opt* は無い)。
MODULE_SHAPED_NAME = re.compile(r"^mod[A-Z]\w*$")


# ==============================================================================
# レイヤー判定(12章§2 のディレクトリ配置基準)
# ==============================================================================
LAYER_CORE = 0
LAYER_APP = 1
LAYER_UI = 2
LAYER_TEST = "test"

LAYER_LABEL = {
    LAYER_CORE: "core層(src/core)",
    LAYER_APP: "app層(src/app)",
    LAYER_UI: "ui層(src/ui)",
    LAYER_TEST: "test層(src/test)",
}


def classify_layer(relpath: Path):
    parts = relpath.parts
    if not parts:
        return None
    top = parts[0]
    if top == "core":
        return LAYER_CORE
    if top == "app":
        return LAYER_APP
    if top == "ui":
        return LAYER_UI
    if top == "test":
        return LAYER_TEST
    return None


# ==============================================================================
# ソース前処理: 行連結・コメント除去・文分割
# ==============================================================================
def merge_continuations(raw_lines: list[str]) -> list[tuple[int, str]]:
    """" _" で終わる行を次行と結合し、(開始行番号, 論理行) のリストにする。"""
    out: list[tuple[int, str]] = []
    acc = ""
    acc_start = None
    for i, raw in enumerate(raw_lines, start=1):
        line = raw.rstrip("\n\r")
        if acc_start is None:
            acc_start = i
        rstripped = line.rstrip()
        # コメント行末の " _" は行継続ではない(VBAの構文上もコメントは継続しない)。
        is_comment_line = line.lstrip().startswith("'")
        cont = (not is_comment_line) and (rstripped.endswith(" _") or rstripped == "_")
        line_wo_cont = rstripped[:-1].rstrip() if cont else line
        acc = f"{acc} {line_wo_cont}" if acc else line_wo_cont
        if not cont:
            out.append((acc_start, acc))
            acc = ""
            acc_start = None
    if acc:
        out.append((acc_start, acc))
    return out


def strip_comment(line: str) -> str:
    """文字列リテラル内の ' は無視して、行コメント(')以降を切り落とす。"""
    in_str = False
    for i, c in enumerate(line):
        if c == '"':
            in_str = not in_str
        elif c == "'" and not in_str:
            return line[:i]
    return line


def split_top_level(s: str, sep: str) -> list[str]:
    """文字列リテラル・カッコの中は無視してトップレベルの sep で分割する。"""
    parts: list[str] = []
    depth = 0
    in_str = False
    cur: list[str] = []
    for c in s:
        if c == '"':
            in_str = not in_str
            cur.append(c)
        elif not in_str and c == "(":
            depth += 1
            cur.append(c)
        elif not in_str and c == ")":
            depth = max(0, depth - 1)
            cur.append(c)
        elif not in_str and depth == 0 and c == sep:
            parts.append("".join(cur))
            cur = []
        else:
            cur.append(c)
    parts.append("".join(cur))
    return parts


def iter_statements(raw_lines: list[str]) -> list[tuple[int, str]]:
    """行連結 -> コメント除去 -> コロン分割まで済ませた (行番号, 文) の列。"""
    stmts: list[tuple[int, str]] = []
    for lineno, ltext in merge_continuations(raw_lines):
        code = strip_comment(ltext)
        for piece in split_top_level(code, ":"):
            piece = piece.strip()
            if piece:
                stmts.append((lineno, piece))
    return stmts


# ==============================================================================
# モジュール情報
# ==============================================================================
@dataclass
class Finding:
    level: str   # "ERROR" | "WARN" | "SKIP"
    line: int
    message: str


@dataclass
class ModuleInfo:
    path: Path
    relpath: Path
    raw_text: str
    vb_name: str
    filename_stem: str
    statements: list = field(default_factory=list)
    public_names: dict = field(default_factory=dict)  # name -> kind
    layer: object = None
    findings: list = field(default_factory=list)

    def add(self, level: str, line: int, message: str) -> None:
        self.findings.append(Finding(level, line, message))


def load_module(path: Path, src_root: Path) -> ModuleInfo:
    raw_text = path.read_text(encoding="utf-8", errors="replace")
    raw_lines = raw_text.splitlines()
    stmts = iter_statements(raw_lines)

    vb_name = ""
    for line in raw_lines[:10]:
        m = ATTRIBUTE_VBNAME_PATTERN.match(line)
        if m:
            vb_name = m.group(1)
            break

    info = ModuleInfo(
        path=path,
        relpath=path.relative_to(src_root),
        raw_text=raw_text,
        vb_name=vb_name,
        filename_stem=path.stem,
        statements=stmts,
    )
    info.layer = classify_layer(info.relpath)

    for lineno, stmt in stmts:
        m = PUB_SUB_PATTERN.match(stmt)
        if m:
            info.public_names[m.group(1)] = "Sub"
            continue
        m = PUB_FUNC_PATTERN.match(stmt)
        if m:
            info.public_names[m.group(1)] = "Function"
            continue
        m = PUB_CONST_PATTERN.match(stmt)
        if m:
            info.public_names[m.group(1)] = "Const"
            continue
        m = PUB_TYPE_PATTERN.match(stmt)
        if m:
            info.public_names[m.group(1)] = "Type"
            continue

    return info


def module_name_for_display(info: ModuleInfo) -> str:
    return info.vb_name or info.filename_stem


# ==============================================================================
# 個別チェック
# ==============================================================================
def check_basics(info: ModuleInfo) -> None:
    if not any(OPTION_EXPLICIT_PATTERN.match(l) for l in info.raw_text.splitlines()):
        info.add("ERROR", 1, "Option Explicit がありません")

    if not info.vb_name:
        info.add("ERROR", 1, "Attribute VB_Name が見つかりません")
    elif info.vb_name != info.filename_stem:
        info.add(
            "ERROR", 1,
            f'Attribute VB_Name="{info.vb_name}" がファイル名"{info.filename_stem}"と不一致',
        )

    # 1モジュール30,000字契約(12章§2)。上限ちょうどまで使い切ると「バグを
    # 1行直すこともできない」状態になるため、残り2,000字で警告する。
    n = len(info.raw_text)
    if n > MAX_MODULE_CHARS:
        info.add("ERROR", 1, f"モジュールが{n}字で上限{MAX_MODULE_CHARS}字を超過")
    elif n > MODULE_WARN_CHARS:
        info.add(
            "WARN", 1,
            f"モジュールが{n}字で上限{MAX_MODULE_CHARS}字まで残り{MAX_MODULE_CHARS - n}字。"
            f"次の修正が入らなくなる前に凝集した機能を新モジュールへ切り出すこと",
        )


def check_module_registry(info: ModuleInfo) -> None:
    """12章§2 のモジュール一覧に無い名前が生えていないか(WARN)。"""
    name = module_name_for_display(info)
    if name in MODULE_REGISTRY or MODULE_REGISTRY_PREFIXES.match(name):
        return
    info.add(
        "WARN", 1,
        f"モジュール「{name}」は12章§2のモジュール一覧にありません"
        f"(仕様に無いモジュールを増やすときは司令塔の裁定が要ります)",
    )


# WindowsのCP932では0x8160はU+FF5Eであり、U+301C(WAVE DASH)には対応バイトが
# 無い。同様にU+2212/U+00A2等もWindows側では落ちる。実機で "?" 化を確認した
# 文字を明示的に拒否する(Pythonのコーデックだけに任せると見逃す)。
# 並びは U+301C WAVE DASH / U+2016 DOUBLE VERTICAL LINE / U+2212 MINUS SIGN /
# U+00A2 CENT / U+00A3 POUND / U+00AC NOT SIGN。この定義自体をCP932内の字で
# 書くことはできないため、エスケープで並べる。
_CP932_DENY = frozenset("\u301c\u2016\u2212\u00a2\u00a3\u00ac")


def _cp932_encodable(ch: str) -> bool:
    # Pythonのcp932コーデックはWindowsより寛容な文字(U+301C等)があるため、
    # Windowsの CP932 表に合わせて明示的に除外する。
    if ch in _CP932_DENY:
        return False
    try:
        ch.encode("cp932")
        return True
    except UnicodeEncodeError:
        return False


def check_cp932_safe(info: ModuleInfo) -> None:
    """CP932に無い文字がソースに混ざっていないか。

    配布ブックは開くたび vba_src シートのソースを VBE へ注入するが、VBE は
    コードを CP932 で保持する。CP932 に無い文字はリテラル "?"(0x3F)として
    保存されるため、実行時の文字列がそのまま化ける。PoC ではこれで
    プロンプト・画面文言の20箇所が化け、vbaProject.bin を覗くまで気付けなかった。

    RPNでは 17章 T-23 の受入条件(プロンプト・スキーマ本文にCP932外の文字が
    1文字も無いこと)がこの検査に直結する。検査するのはコメントを除いた
    実行文だけ(コメント側の注記まで落とすと警告が常駐して本命が埋もれる)。
    """
    for lineno, stmt in info.statements:
        bad = sorted({ch for ch in stmt if not _cp932_encodable(ch)})
        if not bad:
            continue
        shown = " ".join(f"{ch}(U+{ord(ch):04X})" for ch in bad[:6])
        info.add(
            "ERROR", lineno,
            f"CP932に無い文字が実行文に含まれる: {shown} 。"
            f"VBEへの注入時に'?'へ化けます。ChrW()で組むかCP932内の字へ置き換えてください",
        )


MODULE_DECL_RE = re.compile(
    r"^(?:Public|Private|Global|Dim)\s+(?:Const\s+|WithEvents\s+)?(\w+)", re.IGNORECASE)
PROC_HEAD_RE = re.compile(
    r"^(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function|Property)\s+\w+",
    re.IGNORECASE)
PROC_TAIL_RE = re.compile(r"^End\s+(Sub|Function|Property)\s*$", re.IGNORECASE)
TYPE_DECL_RE = re.compile(r"^(?:Public|Private)\s+(?:Type|Enum)\b", re.IGNORECASE)


def collect_module_level_names(raw_text: str) -> set[str]:
    """モジュールレベルで宣言された定数・変数の名前を集める。

    プロシージャの中の Dim は対象外(ローカルなので他モジュールから見えない)。
    Public Type / Public Enum は型名であって変数ではないので除外する。
    """
    names: set[str] = set()
    inside = False
    for raw in raw_text.split("\n"):
        s = strip_comment(raw).strip()
        if not s:
            continue
        if PROC_HEAD_RE.match(s):
            inside = True
        elif PROC_TAIL_RE.match(s):
            inside = False
        elif not inside:
            if TYPE_DECL_RE.match(s):
                continue
            m = MODULE_DECL_RE.match(s)
            if m:
                names.add(m.group(1))
    return names


# ==============================================================================
# OnAction配線ハンドラの再入保護(PoC R15-2b・実機第4報 RC8)
# ------------------------------------------------------------------------------
# Shape.OnAction に文字列で配線された Public Sub は、処理中の DoEvents で
# 発火したクリックから【処理の途中に入れ子で】呼び出され得る。入れ子で走った
# 画面遷移は、実行中の処理が掴んでいるシート状態と噛み合わずに失敗する。
# 保護は全ハンドラの先頭に1行入れるだけだが、手で入れている以上必ず抜ける。
#
# RPNでの読み替え: 関所は modUIProgress.TryEnterUiLock(14章§6・16章E-11)。
# 収集するのは .OnAction = "modX.Y" の【文字列リテラル代入】だけ(変数経由や
# 連結は静的には宛先が定まらないので対象にしない=誤検知ゼロ優先)。
# ==============================================================================
ONACTION_LITERAL_PATTERN = re.compile(
    r"\.OnAction\s*=\s*\"(mod[A-Za-z]\w*\.[A-Za-z_]\w*)\"\s*$", re.IGNORECASE)
ONACTION_GUARD_LOOKAHEAD = 4
# 保護を求めないハンドラの名簿("modX.Y" 形式)。増やすときは必ず理由を1行書くこと。
ONACTION_GUARD_ALLOWLIST: set[str] = set()
ONACTION_GUARD_CALLS = (
    re.compile(r"modUIProgress\s*\.\s*TryEnterUiLock", re.IGNORECASE),
)


def _proc_body_statements(info: ModuleInfo, proc_name: str) -> list:
    """指定Publicプロシージャの本体の実行文(コメント除去済み)を順に返す。"""
    body: list = []
    started = False
    head = re.compile(
        r"^Public\s+(?:Static\s+)?(?:Sub|Function)\s+%s\b" % re.escape(proc_name),
        re.IGNORECASE)
    end = re.compile(r"^End\s+(?:Sub|Function)\b", re.IGNORECASE)
    for lineno, stmt in info.statements:
        if not started:
            if head.match(stmt):
                started = True
            continue
        if end.match(stmt):
            break
        body.append((lineno, stmt))
    return body


def check_onaction_handler_guard(infos: list[ModuleInfo]) -> None:
    wired: dict[str, tuple[str, int]] = {}
    for info in infos:
        for lineno, stmt in info.statements:
            m = ONACTION_LITERAL_PATTERN.search(stmt)
            if m:
                wired.setdefault(m.group(1), (module_name_for_display(info), lineno))

    by_name = {module_name_for_display(i): i for i in infos}

    for target in sorted(wired):
        if target in ONACTION_GUARD_ALLOWLIST:
            continue
        mod_name, member = target.split(".", 1)
        owner = by_name.get(mod_name)
        wire_mod, wire_line = wired[target]
        wirer = by_name.get(wire_mod)
        if owner is None:
            if wirer is not None:
                wirer.add("WARN", wire_line,
                          f"OnAction配線先のモジュールが見つかりません: {target}")
            continue
        if member not in owner.public_names:
            if wirer is not None:
                wirer.add("WARN", wire_line,
                          f"OnAction配線先が見つかりません(Publicではない/改名?): {target}")
            continue

        body = _proc_body_statements(owner, member)
        if not body:
            continue
        head = body[:ONACTION_GUARD_LOOKAHEAD]
        if any(pat.search(stmt) for _, stmt in head for pat in ONACTION_GUARD_CALLS):
            continue
        owner.add(
            "WARN", head[0][0],
            f"OnActionで配線される {target} の先頭に多重実行の関所がありません"
            f"(modUIProgress.TryEnterUiLock を先頭へ。16章E-11。"
            f"関所が不要なハンドラなら vba_lint.py の "
            f"ONACTION_GUARD_ALLOWLIST に理由つきで登録すること)",
        )


# ==============================================================================
# MsgBox/InputBoxへの非BMP文字流出(PoC R18-6c・実機第5報)
# ------------------------------------------------------------------------------
# ChrW(&HD8xx)+ChrW(&HDCxx-DFxx)で組む非BMP絵文字は、Shape/セル値では正しく
# 描けるが、ネイティブMsgBox/InputBoxではサロゲート1単位ごとに "?" 化ける。
# check_cp932_safe とは別問題(こちらは実行時の描画限界)。
#
# RPNでの変更: PoC固有の関数名で構成された「MsgBox到達関数」許可リスト(D)は、
#   対応する関数がRPNに存在しないため空にした。同一文リテラル検出(A)は維持。
#   RPNで「複数のMsgBoxへ流れる文言生成関数」が生えたらここへ足す。
# ==============================================================================
SURROGATE_HIGH_PATTERN = re.compile(r"ChrW\s*\(\s*&H[Dd][89ABab][0-9A-Fa-f]{2}\s*\)")
MSGBOX_CALL_PATTERN = re.compile(r"\b(?:MsgBox|InputBox)\b", re.IGNORECASE)
MSGBOX_REACH_ALLOWLIST: tuple = ()


def check_msgbox_nonbmp(infos: list[ModuleInfo]) -> None:
    for info in infos:
        for lineno, stmt in info.statements:
            if MSGBOX_CALL_PATTERN.search(stmt) and SURROGATE_HIGH_PATTERN.search(stmt):
                info.add(
                    "ERROR", lineno,
                    "MsgBox/InputBoxと同じ文に非BMP絵文字(ChrWのサロゲートペア)が"
                    "直書きされています。ダイアログでは1単位ごとに'?'化けます。"
                    "絵文字を落として鍵括弧表記等へ置き換えてください",
                )

    by_name = {module_name_for_display(i): i for i in infos}
    for target in MSGBOX_REACH_ALLOWLIST:
        mod_name, member = target.split(".", 1)
        owner = by_name.get(mod_name)
        if owner is None or member not in owner.public_names:
            continue
        for lineno, stmt in _proc_body_statements(owner, member):
            if SURROGATE_HIGH_PATTERN.search(stmt):
                owner.add(
                    "ERROR", lineno,
                    f"MsgBox到達関数 {target} の本体に非BMP絵文字があります"
                    f"(呼ばれ方に関わらずMsgBoxで'?'化けます)",
                )


# ==============================================================================
# Split(/Filter(/Array( の具体配列型引数への直渡し(PoC R19-2b・実機第6報)
# ------------------------------------------------------------------------------
# Split()/Filter()/Array() は型システム上ただの Variant を返す式。
# ByRef arr() As String のような具体配列型の仮引数へ実引数位置で直渡しすると、
# 実Excel VBAはプロジェクト全体のコンパイルを拒否する。LibreOffice はこの型
# 検証をしないため素通りし、compileモードでも見逃す(=LOの死角)。
# ==============================================================================
PROC_SIG_HEAD_RE = re.compile(
    r"^(?:Public|Private|Friend)?\s*(?:Static\s+)?(Sub|Function)\s+([A-Za-z_]\w*)\s*\(",
    re.IGNORECASE,
)
ARRAY_PARAM_TYPE_RE = re.compile(r"\(\s*\)\s*As\s+([A-Za-z_]\w*)", re.IGNORECASE)
ARRAY_ARG_CALL_HEAD_RE = re.compile(r"^([A-Za-z_]\w*(?:\.[A-Za-z_]\w*)?)")
ARRAY_ARG_EXPR_RE = re.compile(r"^\s*(Split|Filter|Array)\s*\(", re.IGNORECASE)

ARRAY_ARG_CHECK_KEYWORDS = {
    "if", "elseif", "else", "end", "exit", "for", "each", "next", "while",
    "wend", "do", "loop", "until", "select", "case", "dim", "redim", "const",
    "static", "public", "private", "friend", "function", "sub", "property",
    "type", "enum", "with", "on", "resume", "goto", "gosub", "attribute",
    "option", "declare", "set", "let", "return", "stop", "erase",
    "randomize", "open", "close", "print", "input", "line", "width",
    "name", "kill", "mkdir", "rmdir", "chdir", "chdrive", "filecopy",
    "reset", "get", "put", "lock", "unlock", "debug", "err", "beep",
    "appactivate", "sendkeys", "wait", "implements", "event", "raiseevent",
    "rem", "true", "false", "nothing", "null", "me", "new",
}


def _find_matching_paren(s: str, open_idx: int) -> int:
    """s[open_idx] == '(' 前提。対応する ')' のインデックス(文字列リテラル考慮)。"""
    depth = 0
    in_str = False
    for i in range(open_idx, len(s)):
        c = s[i]
        if c == '"':
            in_str = not in_str
        elif not in_str:
            if c == "(":
                depth += 1
            elif c == ")":
                depth -= 1
                if depth == 0:
                    return i
    return -1


def _has_toplevel_assignment(stmt: str) -> bool:
    """文字列・カッコの外にある単独の '=' (":=" ではない)があれば代入とみなす。"""
    return _toplevel_assign_index(stmt) >= 0


def _toplevel_assign_index(stmt: str) -> int:
    depth = 0
    in_str = False
    for i, c in enumerate(stmt):
        if c == '"':
            in_str = not in_str
        elif not in_str:
            if c == "(":
                depth += 1
            elif c == ")":
                depth = max(0, depth - 1)
            elif c == "=" and depth == 0:
                if i > 0 and stmt[i - 1] == ":":
                    continue
                if i > 0 and stmt[i - 1] in "<>":
                    continue   # <= / >= は比較演算子
                if i + 1 < len(stmt) and stmt[i + 1] == "=":
                    continue
                return i
    return -1


def _is_whole_array_expr_call(arg: str):
    m = ARRAY_ARG_EXPR_RE.match(arg)
    if not m:
        return None
    close_idx = _find_matching_paren(arg, m.end() - 1)
    if close_idx == -1:
        return None
    if arg[close_idx + 1:].strip() != "":
        return None
    return m.group(1)


def _build_array_arg_signatures(infos: list[ModuleInfo]) -> dict:
    sigs: dict = {}
    for info in infos:
        mod_name = module_name_for_display(info)
        for _lineno, stmt in info.statements:
            m = PROC_SIG_HEAD_RE.match(stmt)
            if not m:
                continue
            proc_name = m.group(2)
            close_idx = _find_matching_paren(stmt, m.end() - 1)
            if close_idx == -1:
                continue
            param_str = stmt[m.end():close_idx]
            params = split_top_level(param_str, ",") if param_str.strip() else []
            flags = []
            for p in params:
                pm = ARRAY_PARAM_TYPE_RE.search(p)
                flags.append(bool(pm) and pm.group(1).strip().lower() != "variant")
            sigs[(mod_name, proc_name.lower())] = flags
    return sigs


ARRAY_ARG_CALL_IN_EXPR_RE = re.compile(r"([A-Za-z_]\w*(?:\.[A-Za-z_]\w*)?)\s*\(")
ARRAY_ARG_CALL_PREFIX_RE = re.compile(r"^Call\s+", re.IGNORECASE)


def _blank_string_literals(s: str) -> str:
    """文字列リテラルの中身を空白へ潰す(位置がずれないよう長さは保つ)。"""
    out = []
    in_str = False
    for c in s:
        if c == '"':
            in_str = not in_str
            out.append(c)
        elif in_str:
            out.append(" ")
        else:
            out.append(c)
    return "".join(out)


def _resolve_array_arg_sig(sigs: dict, infos: list[ModuleInfo], info: ModuleInfo,
                           mod_name: str, head: str):
    if "." in head:
        target_mod, proc_name = head.split(".", 1)
        return sigs.get((target_mod, proc_name.lower()))
    sig = sigs.get((mod_name, head.lower()))
    if sig is not None:
        return sig
    candidates = []
    for other in infos:
        if other is info:
            continue
        if head in other.public_names and other.public_names[head] in ("Sub", "Function"):
            s = sigs.get((module_name_for_display(other), head.lower()))
            if s is not None:
                candidates.append(s)
    return candidates[0] if len(candidates) == 1 else None


def _report_array_args(info: ModuleInfo, lineno: int, head: str,
                       arg_str: str, sig: list) -> None:
    args = split_top_level(arg_str, ",") if arg_str.strip() else []
    for idx, arg in enumerate(args):
        arg_s = arg.strip()
        if not arg_s or idx >= len(sig) or not sig[idx]:
            continue
        fn = _is_whole_array_expr_call(arg_s)
        if fn is None:
            continue
        info.add(
            "ERROR", lineno,
            f"{head} の第{idx + 1}引数が具体配列型(As T())なのに、"
            f"Variant配列を返す {fn}() を実引数位置に直接渡しています"
            f"(実Excelはコンパイル拒否。いったん Dim x() As T: x = {fn}(...) "
            f"で受けてから渡すこと)",
        )


def check_array_arg_variant_mismatch(infos: list[ModuleInfo]) -> None:
    sigs = _build_array_arg_signatures(infos)

    for info in infos:
        mod_name = module_name_for_display(info)
        for lineno, raw_stmt in info.statements:
            stmt = ARRAY_ARG_CALL_PREFIX_RE.sub("", raw_stmt, count=1)
            masked = _blank_string_literals(stmt)

            if _has_toplevel_assignment(masked):
                eq = _toplevel_assign_index(masked)
                if eq < 0:
                    continue
                rhs_masked = masked[eq + 1:]
                offset = eq + 1
                for m in ARRAY_ARG_CALL_IN_EXPR_RE.finditer(rhs_masked):
                    head = m.group(1)
                    if head.split(".", 1)[0].lower() in ARRAY_ARG_CHECK_KEYWORDS:
                        continue
                    open_idx = m.end() - 1
                    close_idx = _find_matching_paren(rhs_masked, open_idx)
                    if close_idx == -1:
                        continue
                    sig = _resolve_array_arg_sig(sigs, infos, info, mod_name, head)
                    if sig is None:
                        continue
                    _report_array_args(
                        info, lineno, head,
                        stmt[offset + open_idx + 1:offset + close_idx], sig)
                continue

            head_m = ARRAY_ARG_CALL_HEAD_RE.match(stmt)
            if not head_m:
                continue
            head = head_m.group(1)
            if head.split(".", 1)[0].lower() in ARRAY_ARG_CHECK_KEYWORDS:
                continue

            after = stmt[head_m.end():]
            after_lstrip = after.lstrip()
            if after_lstrip.startswith("("):
                open_idx = head_m.end() + (len(after) - len(after_lstrip))
                close_idx = _find_matching_paren(masked, open_idx)
                if close_idx == -1:
                    continue
                if stmt[close_idx + 1:].strip() != "":
                    continue
                arg_str = stmt[open_idx + 1:close_idx]
            elif after_lstrip:
                arg_str = after_lstrip
            else:
                continue

            if not (split_top_level(arg_str, ",") if arg_str.strip() else []):
                continue

            sig = _resolve_array_arg_sig(sigs, infos, info, mod_name, head)
            if sig is None:
                continue
            _report_array_args(info, lineno, head, arg_str, sig)


# ==============================================================================
# 修飾呼び出し `modX.Proc` の引数数照合(PoC R24-1b)
# ------------------------------------------------------------------------------
# シグネチャ変更の呼び出し側取り残しは、実Excelでは「引数は省略できません」で
# コンパイル不能になるのに、LibreOffice Basic は引数数をコンパイル時に照合
# しないため compile 検査をすり抜ける(LO死角)。
# 設計方針は偽陽性ゼロ最優先。数えられなかった呼び出しは件数だけ必ず表に出す。
# ==============================================================================
ARGC_SIG_RE = re.compile(
    r"^(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function)\s+"
    r"([A-Za-z_]\w*)\s*\(",
    re.IGNORECASE,
)
ARGC_PROPERTY_RE = re.compile(
    r"^(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?Property\s+"
    r"(?:Get|Let|Set)\s+([A-Za-z_]\w*)",
    re.IGNORECASE,
)
ARGC_DECLARE_RE = re.compile(r"^(?:Public\s+|Private\s+)?Declare\b", re.IGNORECASE)
ARGC_QUALIFIED_RE = re.compile(r"\b([A-Za-z_]\w*)\.([A-Za-z_]\w*)")
ARGC_CALL_PREFIX_RE = re.compile(r"^Call\s+", re.IGNORECASE)
ARGC_THEN_RE = re.compile(r"\bThen\b", re.IGNORECASE)
ARGC_ELSE_RE = re.compile(r"\bElse\b", re.IGNORECASE)
ARGC_NOT_ARG_START = set("=+-*/\\&<>^,)(.:;|")

ARGC_SKIPPED: list[tuple[str, int, str, str]] = []
ARGC_CHECKED = [0]
DUMP_ARGC_SKIPS = False


def _argc_statements(raw_lines: list[str]) -> list[tuple[int, str]]:
    """引数数照合用の文リスト。iter_statements と違い ":="(名前付き引数)では
    分割しない(分割すると引数列が壊れて数えられなくなる)。"""
    stmts: list[tuple[int, str]] = []
    for lineno, ltext in merge_continuations(raw_lines):
        code = strip_comment(ltext)
        masked = _blank_string_literals(code)
        depth = 0
        start = 0
        for i, c in enumerate(masked):
            if c == "(":
                depth += 1
            elif c == ")":
                depth = max(0, depth - 1)
            elif c == ":" and depth == 0:
                if i + 1 < len(masked) and masked[i + 1] == "=":
                    continue
                piece = code[start:i].strip()
                if piece:
                    stmts.append((lineno, piece))
                start = i + 1
        piece = code[start:].strip()
        if piece:
            stmts.append((lineno, piece))
    return stmts


def _argc_parse_params(param_str: str):
    """仮引数文字列 -> (必須数, 最大数 or None=無制限)。数えられなければ None。"""
    if not param_str.strip():
        return (0, 0)
    params = [p.strip() for p in split_top_level(param_str, ",")]
    if any(p == "" for p in params):
        return None
    required = 0
    optional_seen = False
    for p in params:
        low = p.lower()
        if low.startswith("paramarray") or low.startswith("byval paramarray"):
            return (required, None)
        if low.startswith("optional"):
            optional_seen = True
            continue
        if optional_seen:
            return None
        required += 1
    return (required, len(params))


def _argc_build_signatures(infos: list[ModuleInfo]) -> dict:
    sigs: dict = {}
    excluded: set = set()
    for info in infos:
        mod_name = module_name_for_display(info)
        for _lineno, stmt in _argc_statements(info.raw_text.splitlines()):
            if ARGC_DECLARE_RE.match(stmt):
                pm = re.search(r"\b(?:Sub|Function)\s+([A-Za-z_]\w*)", stmt, re.IGNORECASE)
                if pm:
                    excluded.add((mod_name, pm.group(1).lower()))
                continue
            pm = ARGC_PROPERTY_RE.match(stmt)
            if pm:
                excluded.add((mod_name, pm.group(1).lower()))
                continue
            m = ARGC_SIG_RE.match(stmt)
            if not m:
                continue
            key = (mod_name, m.group(2).lower())
            close_idx = _find_matching_paren(_blank_string_literals(stmt), m.end() - 1)
            if close_idx == -1:
                excluded.add(key)
                continue
            parsed = _argc_parse_params(stmt[m.end():close_idx])
            if parsed is None:
                excluded.add(key)
                continue
            if key in sigs and sigs[key] != parsed:
                excluded.add(key)
                continue
            sigs[key] = parsed
    for key in excluded:
        sigs.pop(key, None)
    return sigs


def _argc_count_args(arg_str: str):
    if not arg_str.strip():
        return 0
    if ":=" in arg_str:
        return None
    parts = split_top_level(arg_str, ",")
    if any(p.strip() == "" for p in parts):
        return None
    return len(parts)


def _argc_call_starts(masked: str) -> set:
    starts = {0}
    m = ARGC_CALL_PREFIX_RE.match(masked)
    if m:
        starts.add(m.end())
    if re.match(r"^\s*(?:If|ElseIf)\b", masked, re.IGNORECASE) and not ARGC_ELSE_RE.search(masked):
        tm = None
        for tm in ARGC_THEN_RE.finditer(masked):
            pass
        if tm is not None and masked[tm.end():].strip():
            starts.add(tm.end() + (len(masked[tm.end():]) - len(masked[tm.end():].lstrip())))
    return starts


def check_qualified_arg_count(infos: list[ModuleInfo]) -> None:
    ARGC_SKIPPED.clear()
    ARGC_CHECKED[0] = 0
    sigs = _argc_build_signatures(infos)

    for info in infos:
        rel = info.relpath.as_posix()
        for lineno, stmt in _argc_statements(info.raw_text.splitlines()):
            masked = _blank_string_literals(stmt)
            starts = _argc_call_starts(masked)
            for m in ARGC_QUALIFIED_RE.finditer(masked):
                mod_name, proc = m.group(1), m.group(2)
                key = (mod_name, proc.lower())
                if key not in sigs:
                    continue
                required, maxargs = sigs[key]
                tail = masked[m.end():]
                lead = len(tail) - len(tail.lstrip())
                nxt = tail.strip()[:1]

                if nxt == "(":
                    open_idx = m.end() + lead
                    close_idx = _find_matching_paren(masked, open_idx)
                    if close_idx == -1:
                        ARGC_SKIPPED.append((rel, lineno, f"{mod_name}.{proc}", "括弧が閉じていない"))
                        continue
                    after = masked[close_idx + 1:].lstrip()
                    if after[:1] == "(":
                        ARGC_SKIPPED.append((rel, lineno, f"{mod_name}.{proc}", "連鎖呼び出し/添字"))
                        continue
                    if after[:1] == "," and m.start() in starts:
                        ARGC_SKIPPED.append((rel, lineno, f"{mod_name}.{proc}", "括弧なし呼び出しの第1引数が括弧付き"))
                        continue
                    n = _argc_count_args(stmt[open_idx + 1:close_idx])
                    if n is None:
                        ARGC_SKIPPED.append((rel, lineno, f"{mod_name}.{proc}", "名前付き/省略引数"))
                        continue
                    if maxargs == 0 and n > 0:
                        ARGC_SKIPPED.append((rel, lineno, f"{mod_name}.{proc}", "0引数+括弧(添字と区別不能)"))
                        continue
                elif m.start() in starts and nxt and nxt not in ARGC_NOT_ARG_START:
                    rest = stmt[m.end():]
                    n = _argc_count_args(rest)
                    if n is None:
                        ARGC_SKIPPED.append((rel, lineno, f"{mod_name}.{proc}", "名前付き/省略引数"))
                        continue
                elif m.start() in starts and not nxt:
                    n = 0
                else:
                    if required > 0:
                        ARGC_SKIPPED.append((rel, lineno, f"{mod_name}.{proc}", "式中の括弧なし参照"))
                    continue

                ARGC_CHECKED[0] += 1
                if n < required or (maxargs is not None and n > maxargs):
                    if maxargs is None:
                        want = f"必須{required}個以上(ParamArray)"
                    elif required == maxargs:
                        want = f"{required}個"
                    else:
                        want = f"{required}から{maxargs}個"
                    info.add(
                        "ERROR", lineno,
                        f"{mod_name}.{proc} の実引数が{n}個ですが、宣言は{want}です"
                        f"(シグネチャ変更の呼び出し側取り残し。実Excelは"
                        f"「引数は省略できません」等でコンパイル拒否)",
                    )


def check_module_level_refs(infos: list[ModuleInfo]) -> None:
    """他モジュールのモジュールレベル定数・変数を、宣言せずに参照していないか。

    PoCの実機事故: モジュールを切り出したとき、そこで使っている定数を元の
    モジュールへ置いたままにした。実機Excelでは Option Explicit により
    「変数が定義されていません」のコンパイルエラーになりアプリが起動しなく
    なるが、vba_lint も LibreOffice も素通りした(LOは Option VBASupport 下で
    未定義参照を実行時まで遅延する)。
    """
    owners: dict[str, set[str]] = {}
    declared: dict[str, set[str]] = {}
    for info in infos:
        names = collect_module_level_names(info.raw_text)
        declared[info.vb_name] = {n.lower() for n in names}
        for n in names:
            owners.setdefault(n, set()).add(info.vb_name)

    for info in infos:
        mine = declared.get(info.vb_name, set())
        for lineno, stmt in info.statements:
            for ident in set(re.findall(r"(?<![\w.])([A-Za-z_]\w*)", stmt)):
                own = owners.get(ident)
                if not own:
                    continue
                if ident.lower() in mine or info.vb_name in own:
                    continue
                info.add(
                    "ERROR", lineno,
                    f"「{ident}」は {'/'.join(sorted(own))} のモジュールレベル宣言で、"
                    f"{info.vb_name} では宣言されていません"
                    f"(実機Excelで『変数が定義されていません』のコンパイルエラーになります)",
                )


PROC_DEF_RE = re.compile(
    r"^\s*(?:Public|Private|Friend)?\s*(?:Static\s+)?"
    r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)", re.IGNORECASE)
LABEL_DEF_RE = re.compile(r"^\s*([A-Za-z_]\w*):\s*$")
JUMP_RE = re.compile(r"\b(GoTo|GoSub|Resume)\s+\w+", re.IGNORECASE)


def _strip_strings(line: str) -> str:
    """文字列リテラルを空白へ潰す(コメントは strip_comment 済み前提)。"""
    out = []
    in_s = False
    for ch in line:
        if ch == '"':
            in_s = not in_s
            out.append(" ")
        else:
            out.append(" " if in_s else ch)
    return "".join(out)


def check_undefined_proc_refs(infos: list[ModuleInfo]) -> None:
    """他モジュールにしか定義が無いプロシージャを、修飾なしで呼んでいないか。

    PoCの実機事故: モジュールを切り出したとき、切り出した側が使う関数が範囲の
    外にあり持ってくるのを忘れた。実機Excelでは「Sub または Function が定義
    されていません」になり、機能が全滅した。
    """
    proc_owners: dict[str, set[str]] = {}
    proc_defs: dict[str, set[str]] = {}
    labels: dict[str, set[str]] = {}
    for info in infos:
        names, labs = set(), set()
        for raw in info.raw_text.split("\n"):
            m = PROC_DEF_RE.match(raw)
            if m:
                names.add(m.group(1))
            lm = LABEL_DEF_RE.match(strip_comment(raw))
            if lm:
                labs.add(lm.group(1))
        proc_defs[info.vb_name] = {n.lower() for n in names}
        labels[info.vb_name] = {n.lower() for n in labs}
        for n in names:
            proc_owners.setdefault(n, set()).add(info.vb_name)

    for info in infos:
        mine = proc_defs.get(info.vb_name, set())
        labs = labels.get(info.vb_name, set())
        for lineno, raw in merge_continuations(info.raw_text.split("\n")):
            code = _strip_strings(strip_comment(raw))
            if not code.strip():
                continue
            if PROC_DEF_RE.match(code):
                continue
            if LABEL_DEF_RE.match(code):
                continue
            code = JUMP_RE.sub(" ", code)
            for ident in set(re.findall(r"(?<![\w.])([A-Za-z_]\w*)", code)):
                own = proc_owners.get(ident)
                if not own:
                    continue
                if ident.lower() in mine or ident.lower() in labs:
                    continue
                if info.vb_name in own:
                    continue
                info.add(
                    "ERROR", lineno,
                    f"「{ident}」は {'/'.join(sorted(own))} でしか定義されていません。"
                    f"{info.vb_name} には無いので、修飾して呼ぶか、このモジュールへ持ってくること"
                    f"(実機Excelで『Sub または Function が定義されていません』になります)",
                )


ON_ERROR_GOTO_LABEL_RE = re.compile(
    r"^\s*On\s+Error\s+GoTo\s+([A-Za-z_]\w*)\s*$", re.IGNORECASE)
PROC_END_RE = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\b", re.IGNORECASE)
EXIT_PROC_RE = re.compile(r"^\s*(?:Exit\s+(?:Sub|Function|Property)|End)\s*$", re.IGNORECASE)
RESUME_RE = re.compile(r"^\s*Resume\b", re.IGNORECASE)
ERR_RAISE_RE = re.compile(r"\bErr\.Raise\b", re.IGNORECASE)


def check_handler_exit(info: ModuleInfo) -> None:
    """エラーハンドラから通常フローへ戻るとき Resume を使っているか。

    VBAは「エラーハンドラ実行中」という状態を持ち、これを解除できるのは Resume と
    プロシージャを抜けることだけ。On Error GoTo 0 はトラップの登録を消すだけで
    この状態は解除しないため、ハンドラから次のラベルへ落ちる書き方では
    【2回目以降のエラーが呼び出し元へ突き抜ける】。
    """
    lines = [(n, strip_comment(s)) for n, s in merge_continuations(info.raw_text.split("\n"))]
    proc_start = None
    for idx, (_lineno, code) in enumerate(lines):
        if PROC_DEF_RE.match(code):
            proc_start = idx
            continue
        if PROC_END_RE.match(code) and proc_start is not None:
            _check_handler_exit_in_proc(info, lines, proc_start, idx)
            proc_start = None


def _check_handler_exit_in_proc(info: ModuleInfo, lines, start: int, end: int) -> None:
    targets = set()
    for _, code in lines[start:end]:
        m = ON_ERROR_GOTO_LABEL_RE.match(code)
        if m and m.group(1) != "0":
            targets.add(m.group(1).lower())
    if not targets:
        return

    label_at = {}
    for idx in range(start, end):
        m = LABEL_DEF_RE.match(lines[idx][1])
        if m:
            label_at[m.group(1).lower()] = idx

    for name in sorted(targets):
        idx = label_at.get(name)
        if idx is None:
            continue
        stop = end
        for j in range(idx + 1, end):
            if LABEL_DEF_RE.match(lines[j][1]):
                stop = j
                break
        if stop == end:
            continue   # プロシージャ末尾まで届く = 抜けるので問題ない
        body = [lines[j][1] for j in range(idx + 1, stop)]
        if any(RESUME_RE.match(c) or EXIT_PROC_RE.match(c) or ERR_RAISE_RE.search(c) for c in body):
            continue
        info.add(
            "WARN", lines[idx][0],
            f"エラーハンドラ「{name}:」が Resume / Exit を使わずに次のラベルへ落ちている。"
            f"VBAは On Error GoTo 0 ではハンドラ実行中の状態を解除しないため、"
            f"2回目以降のエラーが素通りする(`Resume <ラベル>` で抜けること)",
        )


ON_ERROR_RESUME_NEXT_RE = re.compile(r"^\s*On\s+Error\s+Resume\s+Next\s*$", re.IGNORECASE)
TERMINATOR_RE = re.compile(
    r"^\s*(Exit\s+(Sub|Function|Property)\b|GoTo\s+\w+\s*$|Resume\b|End\s*$)", re.IGNORECASE)


def check_resume_on_fallthrough_label(info: ModuleInfo) -> None:
    """正常系からも落ちてくるラベルの中で Resume を使っていないか。

    Resume は「エラーが起きている」ことが前提の命令で、エラーなしで実行すると
    実行時エラー20「Resume にエラーがありません」になる。
    """
    lines = [(n, strip_comment(s)) for n, s in merge_continuations(info.raw_text.split("\n"))]
    proc_start = None
    for idx, (_lineno, code) in enumerate(lines):
        if PROC_DEF_RE.match(code):
            proc_start = idx
            continue
        if PROC_END_RE.match(code) and proc_start is not None:
            _check_fallthrough_resume_in_proc(info, lines, proc_start, idx)
            proc_start = None


def _check_fallthrough_resume_in_proc(info: ModuleInfo, lines, start: int, end: int) -> None:
    for idx in range(start + 1, end):
        if not LABEL_DEF_RE.match(lines[idx][1]):
            continue
        prev = next((lines[j][1] for j in range(idx - 1, start, -1) if lines[j][1].strip()), "")
        if TERMINATOR_RE.match(prev):
            continue
        for j in range(idx + 1, end):
            code = lines[j][1]
            if LABEL_DEF_RE.match(code):
                break
            if RESUME_RE.match(code):
                info.add(
                    "ERROR", lines[j][0],
                    f"ラベル「{LABEL_DEF_RE.match(lines[idx][1]).group(1)}:」は正常系からも"
                    f"落ちてくるのに Resume がある。エラーなしで Resume を実行すると"
                    f"実行時エラー20になる(ラベルの直前へ GoTo を置き、"
                    f"正常系が Resume を跨ぐようにすること)",
                )
                break


def check_resume_next_inside_handler(info: ModuleInfo) -> None:
    """稼働中のエラーハンドラの内側に On Error Resume Next を書いていないか。

    VBAには「有効(enabled)なハンドラ」と「稼働中(active)なハンドラ」の区別があり、
    ハンドラへ飛んだ瞬間から Resume / Exit / プロシージャ終了までの間に起きた
    エラーは【そのプロシージャでは一切捕まえられず、呼び出し元へ投げ返される】。
    PoCではこれで後始末の失敗が本当の原因を上書きし、フォールバックが一度も
    走らなかった。
    """
    lines = [(n, strip_comment(s)) for n, s in merge_continuations(info.raw_text.split("\n"))]
    proc_start = None
    for idx, (_lineno, code) in enumerate(lines):
        if PROC_DEF_RE.match(code):
            proc_start = idx
            continue
        if PROC_END_RE.match(code) and proc_start is not None:
            _check_resume_next_in_proc(info, lines, proc_start, idx)
            proc_start = None


def _check_resume_next_in_proc(info: ModuleInfo, lines, start: int, end: int) -> None:
    targets = set()
    for _, code in lines[start:end]:
        m = ON_ERROR_GOTO_LABEL_RE.match(code)
        if m and m.group(1) != "0":
            targets.add(m.group(1).lower())
    if not targets:
        return

    for idx in range(start, end):
        m = LABEL_DEF_RE.match(lines[idx][1])
        if not m or m.group(1).lower() not in targets:
            continue
        for j in range(idx + 1, end):
            code = lines[j][1]
            if RESUME_RE.match(code) or EXIT_PROC_RE.match(code):
                break
            if LABEL_DEF_RE.match(code):
                continue
            if ON_ERROR_RESUME_NEXT_RE.match(code):
                info.add(
                    "ERROR", lines[j][0],
                    f"エラーハンドラ「{m.group(1)}:」の稼働中に On Error Resume Next を書いている。"
                    f"VBAはハンドラ稼働中のエラーを同一プロシージャでは捕捉できないため無効で、"
                    f"ここで起きたエラーは呼び出し元へ飛び、本来の原因を上書きする"
                    f"(後始末は別Subへ切り出すか、Resume でハンドラを抜けてから行うこと)",
                )


def check_dim_type_drop_and_integer(info: ModuleInfo) -> None:
    for lineno, stmt in info.statements:
        if AS_INTEGER_PATTERN.search(stmt):
            info.add("ERROR", lineno, f"Integer型は禁止(Longを使う): 「{stmt.strip()[:80]}」")

        # ReDim x(0 To -1) 等の「上限<下限」は LibreOffice Basic では0要素配列と
        # して通るが、実機Excel VBAでは実行時エラー9になる(0要素配列を宣言でき
        # るのはVB.NETであってVBA/VB6ではない)。0件は count変数+ReDim(0 To 0) か
        # Split(vbNullString) で表現する。
        if NEGATIVE_REDIM_PATTERN.search(stmt):
            info.add(
                "ERROR", lineno,
                f"実機Excel VBAで実行時エラー9になる負範囲ReDim(To -1): 「{stmt.strip()[:80]}」",
            )

        m = DIM_STMT_PATTERN.match(stmt)
        if not m:
            continue
        segments = split_top_level(m.group(2), ",")
        if len(segments) < 2:
            continue
        for seg in segments[:-1]:
            if not AS_KEYWORD_PATTERN.search(seg):
                info.add(
                    "ERROR", lineno,
                    f"型落ち疑い(Dim a, b As T は先頭がVariantになる): 「{stmt.strip()[:80]}」",
                )
                break


PROC_DEF_PATTERN = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?"
    r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)",
    re.IGNORECASE,
)
LOCAL_DECL_PATTERN = re.compile(r"\b(?:Dim|Static|Const)\s+(\w+)", re.IGNORECASE)
PROC_PARAM_PATTERN = re.compile(
    r"(?:ByVal\s+|ByRef\s+|Optional\s+(?:ByVal\s+|ByRef\s+)?)(\w+)", re.IGNORECASE
)


def check_name_shadowing(info: ModuleInfo) -> None:
    """変数/引数名が同一モジュール内の手続き名と衝突していないか。

    PoCの実機事故: ローカル変数が同モジュールのSubを隠し、呼び出し行が実機
    Excel VBAで「Sub、Functionまたは Propertyが必要です」のコンパイルエラーに
    なった(LibreOffice Basicは同名解決を許すためLO側は素通り)。
    """
    procs: dict[str, int] = {}
    for lineno, stmt in info.statements:
        m = PROC_DEF_PATTERN.match(stmt)
        if m:
            procs[m.group(1).lower()] = lineno
    if not procs:
        return
    for lineno, stmt in info.statements:
        for m in LOCAL_DECL_PATTERN.finditer(stmt):
            name = m.group(1)
            if name.lower() in procs:
                info.add(
                    "ERROR", lineno,
                    f"変数名「{name}」が同一モジュール内の手続き名と衝突"
                    f"(実機VBAでコンパイルエラーの元): 「{stmt.strip()[:80]}」",
                )
        if PROC_DEF_PATTERN.match(stmt):
            for m in PROC_PARAM_PATTERN.finditer(stmt):
                name = m.group(1)
                if name.lower() in procs and procs[name.lower()] != lineno:
                    info.add(
                        "ERROR", lineno,
                        f"引数名「{name}」が同一モジュール内の手続き名と衝突"
                        f"(実機VBAでコンパイルエラーの元): 「{stmt.strip()[:80]}」",
                    )


MODULE_DECL_PATTERN = re.compile(
    r"^\s*(?:(?:Public|Private|Global)\s+)?(?:Const\s+\w|Declare\s)"
    r"|^\s*(?:Public|Private|Global|Dim)\s+\w+\s*(?:\(\s*\))?\s+As\s+",
    re.IGNORECASE,
)


def check_declaration_position(info: ModuleInfo) -> None:
    """モジュールレベル宣言が最初のプロシージャ定義より後に無いか。

    VBAは宣言部 -> プロシージャ部の順序を強制し、違反すると
    「End Sub、End Function、または End Property の後には、コメントのみが
    記述できます」というコンパイルエラーになる(LibreOffice Basicは途中宣言を
    許容するためLOゲートでは検出不能)。
    """
    depth = 0
    seen_proc = False
    for lineno, stmt in info.statements:
        if PROC_DEF_PATTERN.match(stmt):
            depth += 1
            seen_proc = True
            continue
        if re.match(r"^\s*End\s+(Sub|Function|Property)\b", stmt, re.IGNORECASE):
            depth -= 1
            continue
        if depth <= 0 and seen_proc and MODULE_DECL_PATTERN.match(stmt):
            info.add(
                "ERROR", lineno,
                f"モジュールレベル宣言がプロシージャ定義より後にある"
                f"(実機VBAでコンパイルエラー。宣言部はモジュール先頭へ): 「{stmt.strip()[:80]}」",
            )


RESERVED_IDENT_PATTERN = re.compile(
    r"\b(?:Dim|Const|Static|ByVal|ByRef)\s+base\b"
    r"|[(,]\s*base\s+As\b",
    re.IGNORECASE,
)


def check_reserved_identifiers(info: ModuleInfo) -> None:
    """StarBasic予約語 'Base' をVBAの識別子として使うと、LibreOffice構文チェック
    がコンパイルダイアログでサイレントにハングし、「タイムアウト=構文エラーの
    疑い」としてしか現れず原因特定に多大な時間を要する(PoCで二分探索で特定)。
    実機Excelでは有効な変数名だが、ツールチェーンの沈黙ハングを防ぐ恒久ガード。
    """
    for lineno, stmt in info.statements:
        if RESERVED_IDENT_PATTERN.search(stmt):
            info.add(
                "ERROR", lineno,
                "StarBasic予約語 'base' を識別子に使用(LO構文チェックが沈黙ハング)。"
                f"別名(basePath等)にしてください: 「{stmt.strip()[:80]}」",
            )


# VBA文法キーワードに加えてVBA/Excelの組み込み関数名も総ざらいでブロックリスト化
# する(PoCでは Dim fix As String が実機Windows Excelで「コンパイルエラー: 構文
# エラー」になり、プロジェクト全体が未コンパイル状態に陥って無関係な実行時
# エラー1004として表面化した。LibreOffice は 'Fix' を予約語として扱わない)。
VBA_RESERVED_BLOCKLIST = {
    w.lower() for w in (
        "Dim Static Const Public Private Friend As ByVal ByRef Optional ParamArray "
        "Sub Function Property Get Let Set End If Then Else ElseIf Select Case "
        "For Each In To Step Next Do While Wend Until Loop With Exit GoTo GoSub "
        "Return On Error Resume Call New Nothing Is Like Mod And Or Not Xor Eqv Imp "
        "True False Null Empty Me Option Explicit Compare Type Enum Declare "
        "Lib Alias Implements WithEvents Event RaiseEvent Class Attribute Rem "
        "ReDim Preserve Erase Stop Debug Variant Boolean Byte Integer Long "
        "LongLong LongPtr Single Double Currency Decimal Date String Object "
        "Abs Array Asc AscB AscW Atn CBool CByte CCur CDate CDbl CDec Chr ChrB ChrW "
        "CInt CLng CLngLng CLngPtr Cos CSng CStr CurDir CVar CVDate CVErr "
        "DateAdd DateDiff DatePart DateSerial DateValue Day DDB Dir DoEvents Environ "
        "EOF Exp FileAttr FileDateTime FileLen Filter Fix Format "
        "FormatCurrency FormatDateTime FormatNumber FormatPercent FreeFile FV "
        "GetAllSettings GetAttr GetObject GetSetting Hex Hour IIf IMEStatus Input "
        "InputB InputBox InStr InStrB InStrRev Int IPmt IRR IsArray IsDate IsEmpty "
        "IsError IsMissing IsNull IsNumeric IsObject Join LBound LCase Left LeftB "
        "Len LenB LoadPicture Loc LOF Log LTrim Mid MidB Minute MIRR MkDir Month "
        "MonthName MsgBox Now NPer NPV Oct Partition Pmt PPmt PV QBColor Rate RGB "
        "Right RightB RmDir Rnd Round RTrim Second Seek Sgn Shell Sin SLN Space Spc "
        "Split Sqr Str StrComp StrConv StrReverse Switch SYD Tab Tan Time "
        "Timer TimeSerial TimeValue Trim TypeName UBound UCase Val VarType Weekday "
        "WeekdayName Year Name Kill Width Height Top"
    ).split()
}

RESERVED_WORD_DECL_PATTERN = re.compile(
    r"\b(?:Dim|Private|Public|Static|ReDim(?:\s+Preserve)?)\s+([A-Za-z_]\w*)\s+As\b"
    r"|\b(?:ByVal|ByRef)\s+([A-Za-z_]\w*)\s+As\b",
    re.IGNORECASE,
)


def check_vba_reserved_words(info: ModuleInfo) -> None:
    for lineno, stmt in info.statements:
        for m in RESERVED_WORD_DECL_PATTERN.finditer(stmt):
            name = m.group(1) or m.group(2)
            if name.lower() == "base":
                continue   # check_reserved_identifiers 側で専用メッセージを出す
            if name.lower() in VBA_RESERVED_BLOCKLIST:
                info.add(
                    "ERROR", lineno,
                    f"識別子「{name}」がVBAの予約語/組み込み関数名と衝突"
                    f"(実機VBAでコンパイルエラー「構文エラー」の元): 「{stmt.strip()[:80]}」",
                )


# MS-VBAL 3.3.5.x <reserved-name> / <special-form> の逐語転記。
# これらは仕様上 IDENTIFIER として使えないが、LibreOffice Basic は通す
# (=LOゲートの死角)。PoCでは引数名 scale が実機Excelでのみ構文エラーになり、
# 当該プロシージャがコンパイル不能=「メソッドまたはデータメンバーが見つかり
# ません」として表面化して原因特定にクリーンインストール3回を要した。
MSVBAL_RESERVED_NAMES = {
    w.lower() for w in (
        "Abs CBool CByte CCur CDate CDbl CDec CInt CLng CLngLng CLngPtr CSng "
        "CStr CVar CVErr Date Debug DoEvents Fix Int Len LenB Me PSet Scale "
        "Sgn String "
        "Array Circle Input InputB LBound Scale UBound "
        "Point"
    ).split()
}

MSVBAL_DECL_PATTERNS = (
    re.compile(r"\b(?:Dim|Static|ReDim(?:\s+Preserve)?|Private|Public|Global)\s+"
               r"([A-Za-z_]\w*)\s*(?:\(|\bAs\b|$|,)", re.IGNORECASE),
    re.compile(r"\bConst\s+([A-Za-z_]\w*)\b", re.IGNORECASE),
    re.compile(r"\b(?:ByVal|ByRef|Optional|ParamArray)\s+([A-Za-z_]\w*)\b",
               re.IGNORECASE),
    re.compile(r"\b(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)\b",
               re.IGNORECASE),
)


def check_msvbal_reserved_names(info: ModuleInfo) -> None:
    for lineno, stmt in info.statements:
        hits = set()
        for pat in MSVBAL_DECL_PATTERNS:
            for m in pat.finditer(stmt):
                name = m.group(1)
                if name.lower() in MSVBAL_RESERVED_NAMES:
                    hits.add(name)
        for name in sorted(hits):
            info.add(
                "ERROR", lineno,
                f"識別子「{name}」はMS-VBAL仕様の reserved-name / special-form で、"
                "実機Excel VBAは識別子として拒否します"
                "(LibreOffice Basicは通すためLOゲートでは検出不能)"
                f": 「{stmt.strip()[:80]}」",
            )


def check_excel_tokens(info: ModuleInfo) -> None:
    """R4: Excelトークンは ui層 と R4_EXCEL_ALLOWED_MODULES のみ(12章§2)。"""
    name = module_name_for_display(info)
    if info.layer == LAYER_UI:
        return
    if info.layer == LAYER_TEST and not R4_PURE_REQUIRED_TEST.match(name):
        return   # test層はR4適用外(ただし純ロジック必須の2種は下へ流す)
    if info.layer != LAYER_TEST and name in R4_EXCEL_ALLOWED_MODULES:
        return
    reason = (
        "テストランナー/純ロジックテストはLibreOffice実行テスト(17章§1(c))の本体で、"
        "Excelトークンが入ると(c)が丸ごと動かなくなります"
        if info.layer == LAYER_TEST else
        "R4: Excelトークンはui層と指定モジュール(vba_lint.py の "
        "R4_EXCEL_ALLOWED_MODULES)のみ許可されます"
    )
    for lineno, stmt in info.statements:
        for pattern, label in FORBIDDEN_TOKEN_PATTERNS:
            if pattern.search(stmt):
                info.add(
                    "ERROR", lineno,
                    f"R4違反: 禁止トークン「{label}」使用({reason}): 「{stmt.strip()[:80]}」",
                )


APPLICATION_RUN_PATTERN = re.compile(
    r"Application\s*\.\s*Run\s*\(?\s*(?:\"(?P<lit>[^\"]*)\"|(?P<var>[A-Za-z_]\w*))"
)


def check_application_run_whitelist(info: ModuleInfo) -> None:
    """R3: Application.Run は modGatewayRPN のみ(12章§2)。"""
    name = module_name_for_display(info)
    for lineno, stmt in info.statements:
        for m in APPLICATION_RUN_PATTERN.finditer(stmt):
            lit = m.group("lit")
            var = m.group("var")
            if lit is not None:
                if lit not in RUN_LITERAL_WHITELIST_EXACT:
                    info.add(
                        "ERROR", lineno,
                        f'R3違反: Application.Run("{lit}", ...) はホワイトリスト外: '
                        f'「{stmt.strip()[:80]}」',
                    )
                    continue
                if lit in RUN_LITERAL_GATEWAY_ONLY and name != RUN_GATEWAY_MODULE:
                    info.add(
                        "ERROR", lineno,
                        f'R3違反: Application.Run("{lit}", ...) は{RUN_GATEWAY_MODULE}内のみ許可'
                        f'(このモジュールは{name}): 「{stmt.strip()[:80]}」',
                    )
            else:
                if name not in RUN_VARIABLE_ALLOWED_MODULES:
                    info.add(
                        "ERROR", lineno,
                        f"R3違反: 変数経由のApplication.Run({var}, ...)は"
                        f"{RUN_GATEWAY_MODULE}以外で禁止(このモジュールは{name}): "
                        f"「{stmt.strip()[:80]}」",
                    )


def check_cell_write_guard(info: ModuleInfo) -> None:
    """16章NFR-S7① / 17章T-46①: セルへの書込は modUtilText.SetCellSafe を通す。

    定数・ヘッダの書込は行末に `' SAFE:const` の注記を必須とし、注記の無い
    ヒットのみFAILにする(除外が実装者の裁量ではなく差分レビュー可能な痕跡として
    残る)。除外モジュールは防御関数の定義本体である modUtilText のみ。
    """
    name = module_name_for_display(info)
    if name in CELL_WRITE_EXEMPT_MODULES:
        return
    for lineno, raw in merge_continuations(info.raw_text.split("\n")):
        if SAFE_CONST_MARKER in raw:
            continue
        code = strip_comment(raw)
        for stmt in split_top_level(code, ":"):
            stmt = stmt.strip()
            if not stmt:
                continue
            eq = _toplevel_assign_index(_blank_string_literals(stmt))
            if eq < 0:
                continue
            head = stmt.split(None, 1)[0].lower() if stmt.split() else ""
            if head in ("if", "elseif", "while", "until", "case", "do", "for",
                        "dim", "const", "static", "public", "private", "friend",
                        "redim", "declare", "option", "attribute", "select"):
                continue
            # 代入の左辺 + "=" に検出パターンを当てる(比較 `If .Value = "" Then`
            # を書き込みと誤認しないため。書き込みは必ず代入の形になる)。
            lhs = stmt[:eq] + "="
            if not any(p.search(lhs) for p in CELL_WRITE_PATTERNS):
                continue
            if CELL_WRITE_SAFE_CALL.search(stmt):
                continue
            info.add(
                "ERROR", lineno,
                f"NFR-S7①違反: セルへの直接書込です。外部由来テキストは "
                f"modUtilText.SetCellSafe を通してください"
                f"(定数・ヘッダの書込なら行末に {SAFE_CONST_MARKER} を付けること): "
                f"「{stmt.strip()[:80]}」",
            )


HTML_BARE_IDENT_RE = re.compile(r"(?<![\w.\"])([A-Za-z_]\w*)")


def check_html_embed_guard(info: ModuleInfo) -> None:
    """16章NFR-S7③ / 17章T-46①: HTMLへ差し込む値は HtmlSafe / JsStringSafe を通す。

    対象は HTML を組み立てるモジュール(modExportHtml / modHtmlTemplate* /
    modHtmlTheme)のみ。`s = s & "..." & v & "..."` の形で、連結される素の識別子が
    「代入先のバッファ変数」でも「VBA組み込み定数」でも「HtmlSafe/JsStringSafe
    を通した式」でもない場合にFAILする。定数だけを連結する行や、意図的に
    エスケープ不要な行は行末に `' SAFE:html` を付けて明示する。
    """
    name = module_name_for_display(info)
    if not HTML_BUILDER_MODULE.match(name):
        return
    for lineno, raw in merge_continuations(info.raw_text.split("\n")):
        if SAFE_HTML_MARKER in raw:
            continue
        code = strip_comment(raw)
        for stmt in split_top_level(code, ":"):
            stmt = stmt.strip()
            masked = _blank_string_literals(stmt)
            if "&" not in masked:
                continue
            eq = _toplevel_assign_index(masked)
            if eq < 0:
                continue
            buf = stmt[:eq].strip().lower()
            if HTML_SAFE_CALL.search(stmt):
                continue
            rhs = stmt[eq + 1:]
            rhs_masked = _blank_string_literals(rhs)
            bad = []
            for m in HTML_BARE_IDENT_RE.finditer(rhs_masked):
                ident = m.group(1)
                low = ident.lower()
                if low == buf or low in HTML_SAFE_BARE_TOKENS:
                    continue
                if low in VBA_RESERVED_BLOCKLIST:
                    continue   # 組み込み関数呼び出し(CStr等)は識別子ではない
                bad.append(ident)
            if bad:
                info.add(
                    "ERROR", lineno,
                    f"NFR-S7③違反: HTML連結に素の値 {sorted(set(bad))[:4]} が混ざっています。"
                    f"modUtilText.HtmlSafe / JsStringSafe を通してください"
                    f"(エスケープ不要が確実なら行末に {SAFE_HTML_MARKER} を付けること): "
                    f"「{stmt.strip()[:80]}」",
                )


def check_core_product_vocab(info: ModuleInfo) -> None:
    """12章§4: core層のコードに製品固有の語彙(製品名・シート名)を書かない。"""
    if info.layer != LAYER_CORE:
        return
    for lineno, stmt in info.statements:
        for word in CORE_PRODUCT_VOCAB:
            if word in stmt:
                info.add(
                    "ERROR", lineno,
                    f"12章§4違反: core層に製品固有の語彙「{word}」があります"
                    f"(製品名・シート名はapp層から注入する。config app_tool_prefix 等): "
                    f"「{stmt.strip()[:80]}」",
                )
                break


def check_cross_module_references(info: ModuleInfo, known_modules: dict[str, ModuleInfo]) -> None:
    self_name = module_name_for_display(info)
    seen_skip: set[str] = set()
    for lineno, stmt in info.statements:
        for m in DOTTED_REF_PATTERN.finditer(stmt):
            prefix, member = m.group(1), m.group(2)
            if not MODULE_SHAPED_NAME.match(prefix):
                continue
            target = known_modules.get(prefix)
            if target is None:
                if prefix != self_name and prefix not in seen_skip:
                    seen_skip.add(prefix)
                    info.add(
                        "SKIP", lineno,
                        f"未実装モジュール参照のためスキップ: {prefix}.{member}",
                    )
                continue
            if member not in target.public_names:
                info.add(
                    "ERROR", lineno,
                    f"モジュール間参照エラー: {prefix}.{member} はPublicとして実在しない"
                    f"(現在の{prefix}のPublic一覧: {sorted(target.public_names) or 'なし'})",
                )


def check_layer_dependency(info: ModuleInfo, known_modules: dict[str, ModuleInfo]) -> None:
    """R1: 依存方向 ui -> app -> core の一方向(12章§4)。

    test層は全層を参照してよいので対象外。逆に、どの層からも test層を参照しては
    ならない(製品コードがテストに依存すると配布物からテストを外せなくなる)。
    唯一の例外は R1_TEST_LAYER_EXCEPTIONS の名指しペア(裁定書5 項目7)。
    """
    cur_layer = info.layer
    if cur_layer is None:
        return
    self_name = module_name_for_display(info)

    for lineno, stmt in info.statements:
        for m in DOTTED_REF_PATTERN.finditer(stmt):
            prefix, member = m.group(1), m.group(2)
            if not MODULE_SHAPED_NAME.match(prefix):
                continue
            target = known_modules.get(prefix)
            if target is None or target.layer is None:
                continue
            if prefix == self_name:
                continue

            if target.layer == LAYER_TEST and cur_layer != LAYER_TEST:
                if (self_name, prefix, member) in R1_TEST_LAYER_EXCEPTIONS:
                    # 仕様が明示的に命じた参照(12章§2・14章§4(a)・15章§8・T-14)。
                    # 例外表の根拠コメントは R1_TEST_LAYER_EXCEPTIONS の定義箇所。
                    continue
                info.add(
                    "ERROR", lineno,
                    f"R1違反: 製品コード({LAYER_LABEL[cur_layer]})からテスト層を参照: "
                    f"{prefix}.{member} 「{stmt.strip()[:80]}」",
                )
                continue
            if cur_layer == LAYER_TEST:
                continue
            if isinstance(cur_layer, int) and isinstance(target.layer, int):
                if target.layer > cur_layer:
                    info.add(
                        "ERROR", lineno,
                        f"R1違反: {LAYER_LABEL[cur_layer]}から上位の"
                        f"{LAYER_LABEL[target.layer]}を参照: {prefix}.{member}"
                        f"「{stmt.strip()[:80]}」",
                    )


def check_contract(info: ModuleInfo) -> None:
    name = module_name_for_display(info)
    contract = CONTRACT.get(name)
    if contract is None:
        return
    required = set(contract["required"])
    actual = set(info.public_names.keys())

    for nm in sorted(required - actual):
        info.add("ERROR", 1, f"契約違反: Public {nm} が契約にあるが実装されていない")

    if contract["closed"]:
        for nm in sorted(actual - required):
            info.add(
                "ERROR", 1,
                f"契約違反: Public {nm} は契約に無い(この表は閉じた一覧です)",
            )


ptn_find_call = re.compile(r"\.Find\s*\(", re.IGNORECASE)
ptn_find_lookin = re.compile(r"\bLookIn\s*:=", re.IGNORECASE)


def _paren_args(s: str, open_idx: int):
    depth = 0
    for i in range(open_idx, len(s)):
        c = s[i]
        if c == "(":
            depth += 1
        elif c == ")":
            depth -= 1
            if depth == 0:
                return s[open_idx + 1:i]
    return None


def check_find_lookin(info: ModuleInfo) -> None:
    """`.Find(` で LookIn を省略していたらERRORにする。

    Range.Find は LookIn / SearchOrder / MatchByte を省略すると「そのExcel
    セッションで最後に使われた値」を引き継ぐ(利用者がCtrl+Fの検索ダイアログで
    検索対象を切り替えるだけで更新される)。素の起動状態では xlFormulas なので
    通るが、一度切り替えられると What がセル値と一致していても Nothing が返る。
    LibreOffice の Find にはこの設定持ち越しが無いため run_lo_tests では原理的に
    再現しない(LO検査の死角)。
    """
    for lineno, stmt in info.statements:
        masked = _blank_string_literals(stmt)
        for m in ptn_find_call.finditer(masked):
            args = _paren_args(masked, m.end() - 1)
            if args is None:
                info.add(
                    "ERROR", lineno,
                    "`.Find(` の引数リストの閉じカッコを読み取れませんでした"
                    "(1つの文へ収まる書き方にしてください): "
                    f"「{stmt.strip()[:80]}」",
                )
                continue
            if ptn_find_lookin.search(args):
                continue
            info.add(
                "ERROR", lineno,
                "`.Find(` で LookIn を省略しています。Range.Find は省略すると"
                "『そのExcelセッションで最後に使われた値』を引き継ぐため、利用者の"
                "直前のCtrl+F操作次第で一致するはずのセルにNothingが返ります"
                "(LibreOffice では再現しないためテストで検知できません)。"
                "`LookIn:=xlValues, SearchOrder:=xlByRows, MatchByte:=False` を"
                f"明示してください: 「{stmt.strip()[:80]}」",
            )


RAW_ACTIVATE_PATTERN = re.compile(r"\b([A-Za-z_]\w*)\.Activate\b")
RAW_ACTIVATE_EXEMPT_PREFIXES = {"thisworkbook", "prevactive"}
RAW_ACTIVATE_ALLOW_MARKER = "lint:allow-raw-activate"


def check_raw_activate(info: ModuleInfo) -> None:
    """画面遷移(.Activate)を ui層 の外でやっていないか。

    PoCでは Activate 失敗の脱出路が無い箇所が複数あり、失敗しても利用者には
    何も伝わらず画面が固まったまま戻る手段が無い「無反応画面」になった。

    RPNでの読み替え: PoCは「modUI.ActivateSheetRobust 経由」を強制していたが、
    RPNにはその関数が仕様上存在しない。代わりに「画面遷移はui層の責務」という
    12章§4の依存規則そのものへ寄せ、ui層以外の素の .Activate をERRORにする。
    """
    if info.layer == LAYER_UI:
        return
    for lineno, raw in merge_continuations(info.raw_text.split("\n")):
        if RAW_ACTIVATE_ALLOW_MARKER in raw:
            continue
        code = _strip_strings(strip_comment(raw))
        for m in RAW_ACTIVATE_PATTERN.finditer(code):
            if m.group(1).lower() in RAW_ACTIVATE_EXEMPT_PREFIXES:
                continue
            info.add(
                "ERROR", lineno,
                f"ui層以外での素の.Activate使用(画面遷移はui層の責務。"
                f"失敗時の脱出路が無いと無反応画面になります): 「{code.strip()[:80]}」",
            )


# ==============================================================================
# メイン
# ==============================================================================
def discover_module_files(src_root: Path) -> list[Path]:
    files = sorted(src_root.rglob("*.bas")) + sorted(src_root.rglob("*.cls"))
    return sorted(set(files))


def run_lint(src_root: Path) -> int:
    if not src_root.exists():
        print(f"[vba_lint] 対象ディレクトリが存在しません: {src_root}")
        return 1

    modules = [load_module(f, src_root) for f in discover_module_files(src_root)]
    known_modules = {module_name_for_display(i): i for i in modules}

    for info in modules:
        check_basics(info)
        check_module_registry(info)
        check_cp932_safe(info)
        check_dim_type_drop_and_integer(info)
        check_name_shadowing(info)
        check_declaration_position(info)
        check_handler_exit(info)
        check_resume_next_inside_handler(info)
        check_resume_on_fallthrough_label(info)
        check_reserved_identifiers(info)
        check_vba_reserved_words(info)
        check_msvbal_reserved_names(info)
        check_excel_tokens(info)
        check_application_run_whitelist(info)
        check_cell_write_guard(info)
        check_html_embed_guard(info)
        check_core_product_vocab(info)
        check_cross_module_references(info, known_modules)
        check_layer_dependency(info, known_modules)
        check_contract(info)
        check_raw_activate(info)
        check_find_lookin(info)

    # モジュールをまたいだ検査は、全モジュールの宣言を集め終わってから1回だけ。
    check_module_level_refs(modules)
    check_undefined_proc_refs(modules)
    check_onaction_handler_guard(modules)
    check_msgbox_nonbmp(modules)
    check_array_arg_variant_mismatch(modules)
    check_qualified_arg_count(modules)

    implemented_names = set(known_modules.keys())
    not_yet = sorted(set(CONTRACT.keys()) - implemented_names)

    total_error = total_warn = total_skip = 0

    print("=" * 78)
    print("vba_lint レポート - リスク提案ナビ(RPN)")
    print("=" * 78)

    for info in modules:
        if not info.findings:
            continue
        rel = info.relpath.as_posix()
        total_error += sum(1 for x in info.findings if x.level == "ERROR")
        total_warn += sum(1 for x in info.findings if x.level == "WARN")
        total_skip += sum(1 for x in info.findings if x.level == "SKIP")

        print(f"\n[{rel}]")
        for f in sorted(info.findings,
                        key=lambda x: (x.level != "ERROR", x.level != "WARN", x.line)):
            print(f"  {f.level:<5} L{f.line}: {f.message}")

    if ARGC_SKIPPED:
        # 引数数照合の実効範囲の申告。数え切れなかった呼び出しは黙って通して
        # いるので、件数だけは必ず表に出す(偽陽性ゼロと引き換えの検出漏れ)。
        print("\n[引数数照合 - 数えられずスキップした修飾呼び出し]")
        reasons: dict[str, int] = {}
        for _rel, _ln, _name, why in ARGC_SKIPPED:
            reasons[why] = reasons.get(why, 0) + 1
        detail = " / ".join(f"{k}:{v}件" for k, v in sorted(reasons.items()))
        print(f"  WARN  L1: 引数数照合 {ARGC_CHECKED[0]} 件を照合、"
              f"{len(ARGC_SKIPPED)} 件をスキップしました({detail})")
        total_warn += 1
        if DUMP_ARGC_SKIPS:
            for rel, ln, nm, why in ARGC_SKIPPED:
                print(f"        - {rel}:{ln} {nm} ({why})")

    # CONTRACT完全性の自己検査(裁定書4 項目14): MODULE_REGISTRY(12章§2の一覧)の
    # 各モジュール名に CONTRACT 定義があるか。欠けている=公開契約が未整備のモジュールを
    # 起動時に1件のWARNで一覧化する(将来の契約追加漏れを機械で気付けるようにする)。
    # ERRORにはしない: 実装より契約が先に固まるとは限らず、また型モジュール等は
    # そもそも Public Function 契約を持たないため(それらは required 空で CONTRACT 済)。
    contract_missing = sorted(MODULE_REGISTRY - set(CONTRACT))
    if contract_missing:
        print("\n[CONTRACT完全性 - MODULE_REGISTRYにあるがCONTRACT未定義]")
        print(f"  WARN  L1: {len(contract_missing)}件のモジュールに公開契約(CONTRACT)が"
              "未定義です(14章§6/18章の契約が固まり次第 CONTRACT へ追加): "
              f"{contract_missing}")
        total_warn += 1

    if not_yet:
        print("\n[未実装モジュール(契約はあるがファイル無し) - SKIP]")
        for nm in not_yet:
            print(f"  SKIP  {nm}")
        total_skip += len(not_yet)

    print("\n" + "-" * 78)
    print(f"検査対象ファイル数: {len(modules)}")
    print(f"ERROR: {total_error} 件 / WARN: {total_warn} 件 / SKIP: {total_skip} 件")
    if total_error > 0:
        print("結果: NG(exit code 1) - ERRORを解消してください")
    else:
        print("結果: OK(exit code 0)")
    print("-" * 78)

    return 1 if total_error > 0 else 0


def main() -> int:
    parser = argparse.ArgumentParser(description="リスク提案ナビ VBA静的Lint")
    parser.add_argument(
        "--path", type=str, default=str(DEFAULT_SRC_ROOT),
        help="検査対象ディレクトリ(既定: <repo>/src)",
    )
    parser.add_argument(
        "--dump-argcount-skips", action="store_true",
        help="引数数照合でスキップした呼び出しを1件ずつ列挙する(調査用)",
    )
    args = parser.parse_args()
    global DUMP_ARGC_SKIPS
    DUMP_ARGC_SKIPS = bool(args.dump_argcount_skips)
    return run_lint(Path(args.path).resolve())


if __name__ == "__main__":
    sys.exit(main())
