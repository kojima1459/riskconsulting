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
    パス連結の一元化(裁定書29 W10.1): パスの連結は modUtilPath.JoinPath のみ
       (`& "\"` / `"\" &` の決め打ちを禁止)。一時フォルダは modUtilPath.TempDir
       のみ(`Environ$("TEMP")`/`("TMP")`/`("TMPDIR")` の決め打ちを禁止)

使い方:
    python3 tools/vba_lint.py                 # <repo>/src 配下を検査
    python3 tools/vba_lint.py --path <dir>    # 検査対象を変更(主にテスト用)
    exit code: 0 = ERROR 0件 / 1 = ERROR 1件以上
================================================================================
"""

from __future__ import annotations

import argparse
import io
import json
import re
import sys
import tokenize
from dataclasses import dataclass, field
from pathlib import Path

TOOLS_DIR = Path(__file__).resolve().parent
REPO_ROOT = TOOLS_DIR.parent
DEFAULT_SRC_ROOT = REPO_ROOT / "src"

MAX_MODULE_CHARS = 30000
# 残り2,000字を切ったら警告する(バグ修正1件ぶんの余裕がある状態を保つため)。
MODULE_WARN_CHARS = 28000

# 1物理行のCP932バイト長上限(W5.3 H7 / 裁定書19)。
# VBE の1物理行の上限 1,023 は**文字数ではなくCP932エンコード後のバイト数**で
# 数える。日本語1字=2バイトのため、733字でも1,047バイトになり上限を超える。
# 超えるとVBEが行を勝手に分断し、実機では「SubまたはFunctionが定義されて
# いません」という原因の分からないコンパイルエラーになる(W5.3実機事故:
# modMockLlm.BuildS1NewJson の282行=1,047バイト。テストモジュールは実行時に
# しかコンパイルされないため起動では出なかった)。
# 余白を残して 1,000 バイトで止める(残り23バイトは _ 行継続の追記ぶん)。
MAX_LINE_CP932_BYTES = 1000

# 18章§4.4「テンプレモジュールの分割規約」: modHtmlTemplate* は**25,000字**を
# 超えたら次番のモジュールへ関数単位で切り出す(30,000字の契約に対して余白を
# 持たせる)。17章 T-35 のDoDがこの検査を本ツールへ委ねているため、共通の
# 28,000字WARNではなく専用のERROR閾値として機械強制する。
# 経緯: W3では modHtmlTemplate1 が 25,358字 で§4.4に違反していたが、
#   28,000字のWARN帯に届かないため9ゲートすべてが緑のままだった(W3検証 MINOR1)。
TEMPLATE_MODULE_PATTERN = re.compile(r"^modHtmlTemplate\d+$")
TEMPLATE_MAX_CHARS = 25000

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
    # W3(T-30/T-31)で新設。12章§2のモジュール一覧に追記済み。
    #   modUISheet  = ui層のシート操作プリミティブ(名前付きレンジ・ブロック
    #                 アンカー・図形ボタン+OnAction配線)。11章§5と13章§2.9の
    #                 実装を1箇所へ集め、6つのui層モジュールが同じ12行を各自
    #                 持つのを防ぐ。
    #   modUICase2  = modUICase の分割先(30,000字契約)。S1～S4シートの描画と
    #                 編集の確定(13章§2.12-§2.15・§2.2 規約4)。
    #   modUICase5  = modUICase2 の分割先(30,000字契約。裁定書9 W4.1)。逆シリ
    #                 アライズ本体とブロックの幾何(ColIndexes/ColCount/RoomOf)。
    #                 列の引き当て方と部屋の数え方を読み書きで1本に保つ。
    #   modUICase3  = 同上。案件入力・フィードバック・判断台帳(13章§2.11/
    #                 §2.5/§2.7)。
    #   modUICase4  = 同上。フィードバック・判断台帳(13章§2.5/§2.7)。
    #   modUICaseFmt= 同上。13章§2.2 セル格納規約の変換(セル<->JSON値)の純関数。
    #                 modKnowledgeFmt が15章の整形規約を持つのと同じ切り口。
    #   modUIGuide  = 初回ガイドツアーと操作ガイドの図形ボタン(裁定書14 裁定6)。
    #                 起動からの結線は modBoot の1行(StartTourIfFirstRun)のみ。
    #   modUIToast  = トースト(図形カード+Application.OnTimeの自動消去。接頭辞
    #                 ts_)と hm_warning の1行組み立て(裁定書17 H2/H4)。文言の
    #                 値源をここへ寄せ、満杯の modUIHome へ文字を足さない。
    #   modUIHome2  = modUIHome の分割先(30,000字契約。17章§7 Z-13)。HOMEの
    #                 OnActionハンドラ群(画面は modUIHome・動作は modUIHome2)。
    "modUISheet", "modUICase2", "modUICase3", "modUICase4", "modUICase5",
    "modUICaseFmt", "modUIGuide", "modUIToast", "modUIHome2",
    # W6第1弾(1画面ナビ・17章 T-49)で新設。12章§2のモジュール一覧に追記済み。
    #   modUINav     = ナビの状態(STEPの自動決定)とOnActionハンドラ(11章§3.1.1)。
    #   modUINavDraw = ナビの描画(コーチ帯・区画パネル・強調枠・待ちカード)。
    #   modUIResearch= 区画①の調べる文8本の組み立てと[コピー](11章§3.2)。
    #   modUICase6   = 区画②の保管+プレビュー貼付(11章§3.3・§7.2(a))。
    #   modUICase7   = modUICase6 の分割先(30,000字契約。裁定書22 m2)。Ctrl+V
    #                  直貼り枠の取り込みと、その失敗の理由分け(はみ出し/表の線・
    #                  画像/個人情報)。
    "modUINav", "modUINavDraw", "modUIResearch", "modUICase6", "modUICase7",
    # W8.1(実機第1報・裁定書26 B)で新設。12章§2のモジュール一覧に追記済み。
    #   modUIViewport = 全画面表示(DisplayFullScreen・数式バー・罫線・行列見出し。
    #                   シートタブは触らない)。modBoot の起動シーケンスの最後で
    #                   1回だけ適用し、退避した元の値へ戻す口を1本持つ。
    "modUIViewport",
    # W10(裁定書28・17章 T-59)で新設。12章§2のモジュール一覧に追記済み。
    #   modBootData = 起動時の案件一覧の再構成(data_dir の企業フォルダを Dir$ で
    #                 走査し、企業ファイルの見出しから案件一覧を組み直す)。
    #                 modBoot は 30,000字契約を超えているため実体を置けず、
    #                 起動シーケンスからは手順(3)の1行だけで呼ぶ。
    "modBootData",
    # W12-A(裁定書34)で新設。12章§2のモジュール一覧は班2が追記する。
    #   modBootNavi = modBoot の分割先(30,000字契約)。HTML画面まわりの config
    #                 既定値・ui_mode=html の起動判断(LaunchIfHtml)・
    #                 case_data!data_key の入力規則(11章§5)。
    #   modNavi*    = HTML画面(モードレスの1枚窓)の実装。髙橋さん版 v0.3 を
    #                 当方規約へ寄せて取り込んだ。ui 層なので R4 の
    #                 Excelトークンは許可(許可モジュール表は増やさない)。
    #                 frmNaviHtml(UserForm)は .frm なので本表とは別扱い。
    "modBootNavi",
    "modNaviHost", "modNaviJson", "modNaviState", "modNaviState2",
    "modNaviStore", "modNaviChat", "modNaviActions", "modNaviActions2",
    # (W9.3 で clsAppEvents を撤去。ブックイベントは ThisWorkbook 文書モジュール
    #  が受ける。クラスの焼き方の不一致は 17章 Z-24 で解決済みだが、配布物は
    #  **可動部品を減らす**ためクラスを持たない。理由と経緯は 17章 Z-24 と
    #  build/build_rpn.py の _BAKED_THISWORKBOOK_TEXT のコメント)
    # ---- app 層 ----
    "modPipeline", "modPlayOps", "modSparring", "modCaseStore", "modCaseRead",
    "modInboxStore",
    "modJudgeStore", "modKnowledge", "modKnowledgeFmt", "modValidate",
    # modKnowledge2 = modKnowledge の分割先(30,000字契約。17章§7 Z-13)。
    #   シートを触らない純関数(絞込・列引き・E-34の純部)だけを持つため
    #   R4_EXCEL_ALLOWED_MODULES へは足さない(層(a)から直接叩ける)。
    "modKnowledge2",
    # modKnowledgeRank = 裁定書38 B-10。業種コード完全一致で上限に満たない
    #   ときの全業種補充を n-gram 重なり数で順位付けする純関数(NgramOverlap/
    #   RankRows)。呼ぶのは modKnowledge2 だけ。R4_EXCEL_ALLOWED_MODULES へは
    #   足さない(Excelトークンを一切持たない)。
    "modKnowledgeRank",
    # 分割・新設の追認は裁定書7 B-7/B-8(12章§2のモジュール一覧に追記済み)。
    #   modValidate2 / modCompanyFile2 = 30,000字契約による分割先。
    #   modCaseRead = 案件一覧の読取専用API(app層。R4許可も併せて追加)。
    # W15(裁定書38 班A)で新設。modValidate / modValidate2 が30,000字契約で
    #   満杯のため、V-S1-14 / V-S1-15(落とさない警告)だけを持つ。12章§2に追記済み。
    "modValidate2", "modValidate3", "modCompanyFile", "modCompanyFile2", "modPii",
    # W10(裁定書28・17章 T-59)で新設。12章§2のモジュール一覧に追記済み。
    #   modCompanyFile3 = 企業ファイルのスキーマ拡張(dossier_case/data/judge)・
    #                     自動保存・起動時再構成の読取口。modCompanyFile は
    #                     29,833字で満杯、modCompanyFile2 は下位I/Oだけを持つ
    #                     約束なので3本目を切った(R4許可も併せて追加)。
    "modCompanyFile3",
    # 裁定書8 A-1/A-2 が承認した分割先(12章§2のモジュール一覧に追記済み)。
    #   modPipeline2  = 入念モードの批判・改訂パイプ(T-28)。R4許可は与えない
    #                   (modPipeline と同じくシートに触れない)。
    #   modCaseStore2 = 案件2枚の下位シートI/O(R4許可も併せて追加)。
    #   modCaseStore3 = 同じ切り口の3本目(30,000字契約。17章§3の分割バックログ2番。
    #                   T-56)。data_key の一覧(DataKeys)とラウンド確定の付帯処理
    #                   (ApplyRoundFocus。13章§2.1 adopted_story_nos/focus_line_ids)。
    "modPipeline2", "modCaseStore2", "modCaseStore3",
    # W7(17章 T-55/T-57・裁定書25)で新設。12章§2のモジュール一覧に追記済み。
    #   modPipeline3 = 15章 v2.6 で増えた3プレースホルダ({{financeText}} /
    #                  {{incidentsText}} / {{focus_line_ids}})の値源組立。
    #                  modPipeline が30,000字契約で満杯のため分割した。
    #                  シートには store 経由でしか触れないので R4許可は与えない。
    #   modPipeline4 = 15章§0.7 の切詰め計画(TrimPlan)と関連する純関数。
    #                  modPipeline が30,000字契約で満杯のため分割した(T-57)。
    #                  シートに触れないので R4許可は与えない。
    "modPipeline3", "modPipeline4",
    # W14(裁定書37 B-03)で新設。12章§2のモジュール一覧に追記済み。
    #   modGround = evidence.quote の原文照合(NormalizeForMatch / QuoteFound /
    #               GroundNotes)。app層の**純文字列**モジュールで、案件データも
    #               config も読まない(値源の解決は modPipeline3.GroundHook)。
    #               R4許可は与えない(Excelトークンに触れない)。
    "modGround",
    "modExportHtml", "modExportPpt", "modExportHearing", "modAppTypes",
    # W15(裁定書38 班A)で新設。modPromptsCore が30,000字契約に達したため
    #   15章§2(S1 の system / user)を切り出した分割先。12章§2に追記済み。
    "modPromptsCore", "modPromptsCore2", "modPromptsBlocks", "modPromptsOps",
    "modSchemas",
    "modHtmlTheme",
    # ---- core 層 ----
    "modGatewayRPN", "modGatewayDirect", "modJsonLite", "modConfig", "modLog",
    # W11-a(裁定書30 裁定1(b))で新設。12章§2のモジュール一覧に追記済み。
    #   modGatewayLink = direct経路への薄い接続点。**同名で2ソース**を持ち、
    #                    prod(src/core/modGatewayLink.bas)は E0209 を返すだけ、
    #                    dev(src/core/dev/modGatewayLink.bas)は modGatewayDirect
    #                    へ転送する。どちらを使うかは build/modules.json の
    #                    dev_src が決める(実行時の名前ディスパッチはしない)。
    "modGatewayLink",
    # W12-A(裁定書34 §1.2)で新設。modGatewayRPN の分割先(30,000字契約)。
    #   引数だけで答えが決まる4本(BuildToolName / ResolveWaitSec /
    #   ResolveMaxTokens / TrimHistoryPairs)。挙動は移設前と同じ。
    "modGatewayRPN2",
    "modUtil", "modUtilText", "modTypes",
    # W10.1(裁定書29 T-60)で新設。12章§2のモジュール一覧に追記済み。
    #   modUtilPath = パスの連結(JoinPathWith/JoinPath)・分解(FileNameOf)・
    #                 一時フォルダ(TempDir)。区切り文字を知っているのはここだけ、
    #                 という状態を作る(下の check_path_join_literal が製品側の
    #                 決め打ちを禁止する)。modUtil が30,000字契約で満杯のため
    #                 分割した。**Excelトークンは1つも持たない**(Mac判定
    #                 IsMacExcel は core へ置かず modTestsExcel3 が持つ)。
    "modUtilPath",
    # W6第1弾(1画面ナビ・17章 T-49)で新設。12章§2のモジュール一覧に追記済み。
    #   modUIGeom  = 画面の幾何(帯・ボタンの並び・カードの高さ・表示時間)の純関数。
    #                Excelを1つも触らないので層(a)からテストできる(11章§8.6の流用表)。
    #   modNavText = 貼付テキストの純変換(StripDrFooter / PreviewLines /
    #                SplitFieldNotes / JoinFieldNotes。11章§7.2(a))。
    "modUIGeom", "modNavText",
    # W15(裁定書38 班C)で新設。12章§2のモジュール一覧に追記済み。
    #   modPromptsS5 / modSchemas2 = 15章§5.6 の S5 プロンプトとスキーマ
    #   modValidate4 = CheckS5(modValidate/2 は満杯・3 は班A が使用)
    #   modPipeline5 = S5 の実行(1呼び出し+修復1回)
    #   modProposalHtml1..4 = 顧客向け提案書 Wide 22枚のテンプレ(20章)
    #   modExportProposal   = 提案書の DATA 組立・書出(reviewedBy 必須)
    "modPromptsS5", "modValidate4", "modPipeline5",
    "modProposalHtml1", "modProposalHtml2", "modProposalHtml3",
    "modProposalHtml4", "modExportProposal",
    # W11-c(裁定書33 C-4)で新設。12章§2のモジュール一覧に追記済み。
    #   modRibbonWire = 経路側で本文が切られた疑いの判定(LooksRibbonCut)と
    #                   err_log detail の1語(CutNote)。16章 E-63。
    "modRibbonWire",
    # ---- test 層 ----
    # modTestsRunnerUi = ブック内テスト実行(17章 T-48・裁定書14 裁定5)。
    #   ターミナルの使えない社内PC向けに ps1 と同じ4条件をブック内で回す。
    "modTestRunner", "modTestsExcel", "modMockLlm", "modTestsRunnerUi",
    # W11-a(裁定書30 裁定1(d))で新設。12章§2のモジュール一覧に追記済み。
    #   modTestsPureHook = dev専用の純層テストを呼ぶ接続点(modGatewayLink と
    #                      同じ**同名2ソース**。prod版は1本も実行しない)。
    #   modTestsPureDev  = dev専用の純層テスト(modGatewayDirect の純関数17本)。
    #                      配布物には載らない(modules.json の ship:false)。
    "modTestsPureHook", "modTestsPureDev",
    # W11-c(裁定書33 C-3)で新設。リボンちゃんの応答抽出の逐語模擬(純関数)。
    "modRibbonSim",
    # W15(裁定書38 §1 班E・Z-45)で新設。12章§2のモジュール一覧に追記済み。
    #   modReportMail = 報告メールの下書き作成(Outlook.Application を late-bound
    #                   で.Display。自動送信はしない)。失敗時は報告文を返し
    #                   HTML画面側の[コピー]導線へフォールバックする。
    "modReportMail",
    # ---- フォーム(.frm) ----
    # W12-A(裁定書34)。HTML画面の器(UserForm + WebBrowser)。標準モジュールでは
    #   ないが、Attribute 行より後は同じ規則で検査する(discover_module_files)。
    "frmNaviHtml",
    # W12-A(裁定書34 §1.1)。HTML画面のテスト2本。純層は modTestRunner から、
    #   層(b)は modTestsExcel3 から結線する(modTestsExcel は 29,902字で満杯)。
    "modTestsPureNavi", "modTestsExcelNavi",
}
# 分割される可能性のあるモジュール名(末尾に1以上の数字が付く)。
# modMockLlm1..n は 12章§2(v2.4.1)が 30,000字契約による分割を明記している。
MODULE_REGISTRY_PREFIXES = re.compile(
    r"^(modTestsPure|modSchemas|modHtmlTemplate|modPromptsCore|modPromptsBlocks"
    r"|modPromptsOps|modMockLlm|modTestsExcel)\d*$"
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
    # W6第1弾(T-49)。14章§6へ登記した公開名。closed=False(内部の追加Publicは許す)。
    # W6第1弾(T-49)。ui層の新設4本。14章§6へ登記した公開名。
    "modUINav": {
        "closed": False,
        "required": ["DrawNav", "NavPrev", "NavNext", "ShowDrafts", "BackToNav",
                     "CurrentStep", "StepText", "StepAnchor", "FieldNotesWritten",
                     "StepFor", "StepRuleOf", "StepActionOf", "DrawOk"],
    },
    "modUINavDraw": {
        "closed": False,
        "required": ["DrawCoachBar", "DrawSections", "MoveFocusFrame",
                     "DrawWaitCard", "HideWaitCard", "RefreshAreas", "RefreshArea",
                     "EnsureAreaButtons", "ApplyAreaVisibility", "ShowMoreRows",
                     "DropNavShapes", "ResetForNewCase", "FieldNotesTemplate",
                     "LoadFieldNotesFor", "StyleFieldNotesHeads"],
    },
    "modUIResearch": {
        "closed": False,
        "required": ["BuildPrompts", "EnsureCopyButtons", "ToggleMore",
                     "CopyPrompt1", "CopyPrompt8", "FillTemplate",
                     "PlaceholderTable", "PlaceholderKeys"],
    },
    "modUICase6": {
        "closed": False,
        # W15 Round3(裁定書41 §1): SentinelCheck / MergedOrShapeCheck を撤去した
        # (呼出元0件の通過口。はみ出し判定は ReadDirectPaste の overflow、
        #  結合セル・図形は modUICase7.MergedOrShapeAt が本体)。
        "required": ["StoreArea", "LoadArea", "ReadDirectPaste",
                     "PasteIntoArea", "ShowArea", "ClearArea",
                     "SaveNav", "AreaKeys", "AreaField", "AreaBody",
                     "AreaTable", "HandlerName", "WriteFieldNotesArea"],
    },
    # modUICase7: 30,000字契約による modUICase6 の分割先(裁定書22 m2)。SaveNav の
    # 結線先そのものであり、直貼り枠の取り込みと失敗の理由分けを持つ。
    "modUICase7": {
        "closed": False,
        "required": ["ImportDirectPastes", "MergedOrShapeAt", "BlockedText",
                     "ShapeHitCount", "JoinExisting"],
    },
    "modUIGeom": {
        "closed": False,
        # W15 Round3(裁定書41 §1): 移植しただけで一度も呼ばれなかった
        # SumSpan / ClipToWidth / PillWidth / CardWaitMsFor の4本を撤去した。
        "required": ["FlowLeft", "ClipToChars",
                     "StepDots", "CardHeightFor",
                     "ToastSecondsFor"],
    },
    "modNavText": {
        "closed": False,
        "required": ["StripDrFooter", "PreviewLines", "SplitFieldNotes",
                     "JoinFieldNotes", "NormalizeEol", "FitsInRows"],
    },
    # ---- app 層 ----
    "modValidate": {
        "closed": False,
        "required": [
            "NormalizeLlmJson", "CheckS1", "CheckS2", "CheckS3", "CheckS4",
            "CheckS2C", "CheckS3C", "CheckPF",
        ],
    },
    # modValidate2: 30,000字契約による分割先(裁定書7 B-5 が §6 へ *Core 5本を宣言。
    #   「分割の継ぎ目」であり呼び出してよいのは modValidate だけ)。
    "modValidate2": {
        "closed": False,
        "required": [
            "NormalizeCore", "CheckS4Core", "CheckPFCore", "CheckS2CCore",
            "CheckS3CCore",
        ],
    },
    # modValidate3: S1の**落とさない警告**(V-S1-14 / V-S1-15)。modValidate /
    #   modValidate2 が満杯のため新設した(裁定書38 班A)。呼び出しは
    #   modPipeline3(注記経路)と modExportHtml(meta.s1_warn)だけ。
    "modValidate3": {
        "closed": True,
        "required": [
            "HeadOverlap","CheckS1Notes", "WarnNoteOf", "TrimUrl",
                     # 裁定書39 R1-09 / X-1: modValidate が30,000字契約に達した
                     # ため、V-S1-16 / V-S1-17 と S1 の後正規化(sources 補填)の
                     # 実体をこちらへ置いた(呼ぶのは modValidate だけ)。
                     "SoftNotesS1", "PostNormalize"],
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
    # modKnowledge2: 30,000字契約による modKnowledge の分割先(17章§7 Z-13)。
    # 移設した純関数8本は modKnowledge からの唯一の呼出先であり、改名・Private化
    # はその場で読込・絞込・E-34検査が落ちるため required で固定する。
    "modKnowledge2": {
        "closed": False,
        "required": [
            "PickAt", "CellAt", "CellRaw", "AddIdList", "ColOf", "SelectRows",
            "MissingColsOf", "BadRowsOf",
        ],
    },
    # modKnowledgeRank: 裁定書38 B-10。全業種補充の順位付け専用の純関数2本と、
    #   裁定書40 P-M3(c) の計測専用2本(IndexBuilds / ResetIndexBuilds。索引を
    #   作った累計回数。判定には使わない)。**closed=True** にして、以後この
    #   モジュールに本番経路から呼ばれない Public が黙って増えないようにする
    #   (裁定書41 §2。検証者 newIssues: 計測口が 14章§6 にも CONTRACT にも
    #   登記されておらず、どのゲートも赤くならなかった)。
    "modKnowledgeRank": {
        "closed": True,
        "required": ["NgramOverlap", "RankRows", "IndexBuilds", "ResetIndexBuilds"],
    },
    # 14章§6(裁定書6 B/C)。整形の純関数はここが唯一の実装。
    # 14章§6 modUtil節(裁定書6 項目8で契約化)。10関数。
    "modUtil": {
        "closed": False,
        "required": [
            "SplitForCells", "JoinCellChunks", "SplitKeepNonEmpty", "AppendIdList",
            "ClampLong", "SafeLeft", "BufInit", "BufAdd", "BufText", "FindHeaderCol",
        ],
    },
    # 裁定書29 W10.1(17章 T-60)。パスの連結・分解と実行環境の判定。
    #   純関数(JoinPathWith / FileNameOf)は層(a)の回帰網が叩くので required に
    #   載せる(Private へ落として層(a)を空振りさせる改変をここで止める)。
    "modUtilPath": {
        "closed": False,
        "required": [
            "JoinPathWith", "JoinPath", "TempDir", "FileNameOf",
        ],
    },
    "modKnowledgeFmt": {
        "closed": False,
        "required": [
            "FmtRiskLib", "FmtMenus", "FmtMenusSummary", "FmtLines", "FmtCases",
            "FmtSchemes", "FmtPatterns", "FmtRules", "FmtResearching", "FmtMechs",
            "TrimKbLine",
        ],
    },
    "modPromptsCore": {
        "closed": False,
        "required": [
            "BuildS2System", "BuildS2User",
            "BuildS3System", "BuildS3User", "BuildS4System", "BuildS4User",
        ],
    },
    # W15(裁定書38 班A): 15章§2 の2関数は modPromptsCore2 へ移した
    #   (modPromptsCore が30,000字契約に達したため。15章§10.2 の対応表も同時改訂)。
    "modPromptsCore2": {
        "closed": True,
        "required": ["BuildS1System", "BuildS1User"],
    },
    "modPromptsBlocks": {
        # 15章§10.2 末尾が「Block* の7関数」と明言しているため closed=True。
        "closed": True,
        "required": [
            "BlockCtx", "BlockRenewalS1", "BlockRenewalS2", "BlockRenewalS3",
            "BlockNewS2", "BlockRound2Focus",
            "BlockGuard", "BlockS4Proposal", "BlockS4Alliance",
        ],
    },
    "modPromptsOps": {
        "closed": False,
        "required": [
            "BuildS2CriticSystem", "BuildS2CriticUser", "BuildS3CriticSystem",
            "BuildS3CriticUser", "ReviseSuffix", "BuildSparringSystem",
            "BuildPFSystem", "BuildPFUser", "RepairSuffix",
            # 組立層(14章§6 (2)・裁定書6 B)。
            "Fill", "AsmS1User", "AsmS2User", "AsmS3User", "AsmS4System",
            "AsmS4User", "AsmS2CriticUser", "AsmS3CriticUser",
            "AsmSparringSystem", "AsmPFUser",
        ],
    },
    "modSchemas": {
        "closed": False,
        "required": [
            "SchemaS1", "SchemaS2", "SchemaS3", "SchemaS4",
            "SchemaS2C", "SchemaS3C", "SchemaPF",
        ],
    },
    # modPipeline: 14章§6が RunAll / RunStep と**判定核16本**を宣言(裁定書7 B-6)。
    #   16本は層(a)=modTestsPure から叩く純関数なので Private へ戻したら ERROR。
    "modPipeline": {
        "closed": False,
        "required": [
            "RunAll", "RunStep",
            "TrimInputPlan", "ProtectedOverBudget", "TruncField", "BudgetOf",
            "NeedsRepair", "ClassifyResult", "FailCodeOf", "CaseIdOfLine",
            "DeepEnabled", "ResolveQualityMode", "StepNameOf", "StatusForStep",
            "PlayIdOf", "KbRowCount", "UsesSlot", "S1SummaryOf",
        ],
    },
    # modPipeline2: 30,000字契約による分割先(裁定書8 A-1。入念モードの批判・改訂
    #   パイプ=T-28)。**分割の継ぎ目**であり呼んでよいのは modPipeline だけ。
    #   RunDeep は modPipeline からの1行フックの受け口なので required で固定する
    #   (改名・Private化は契約違反)。加えて14章§6が「Private へ戻すことは契約
    #   違反」と明記した**判定核8本**も required に載せる(裁定書9-3。§6の宣言文
    #   と lint の担保を一致させる。modSparring / modPipeline と同じ扱いで、
    #   Private 化すると層(a)の回帰網が黙って消えるため機械で検出する)。
    #   AdoptRevisionOf は E-36 の「改訂を破棄して改訂前を採用」の唯一の選択点
    #   (裁定書9-2)。closed=False は内部ヘルパの追加を許すため。
    "modPipeline2": {
        "closed": False,
        "required": [
            "RunDeep",
            "CritiqueStepOf", "ReviseStepOf", "NeedsRevision", "CritiqueDigest",
            "DeepOutcomeOf", "AdoptRevisionOf", "DeepWarningOf", "DeepRouteOf",
        ],
    },
    # modCaseRead: 案件一覧の読取専用API(裁定書7 B-7。14章§6が ReadCaseCtx を宣言)。
    "modCaseRead": {"closed": False, "required": ["ReadCaseCtx", "CaseColumnOf"]},
    # modPipeline3: 30,000字契約による modPipeline の分割先(17章 T-55・裁定書25)。
    #   15章 v2.6 の3プレースホルダの値源組立。純関数3本(FinanceBlockText /
    #   IncidentsBlockText / FocusLineIdsAttr)は層(a)から叩く回帰網の対象なので
    #   required に載せる(Private へ戻すと検査が黙って消える)。IncidentsFor は
    #   班C(T-56)の modKnowledge 実装へ差し替えるための**呼び口1本**。
    "modPipeline3": {
        "closed": False,
        "required": [
            "S1UserText", "S2UserText", "S3UserText",
            "FinanceBlockText", "IncidentsBlockText", "FocusLineIdsAttr",
            "IncidentsFor",
            # 裁定書37 B-03/B-05。原文照合の呼び口と充足度の1語化。
            # W15 Round3(裁定書41 §1): LastGroundNote / ResetGroundNote /
            # LastS1Notes / ResetS1Notes は読む者が0件だったため撤去した
            # (HTMLレポートは modExportHtml が自分で測り直している)。
            "DefendNotes", "GroundHook", "BuildHaystack",
            "SufficiencyNoteOf",
        ],
    },
    # modPipeline4: 30,000字契約による modPipeline の分割先(17章 T-57・裁定書25
    #   S6)。15章§0.7 の切詰めを6段化するにあたり、計画(TrimPlan。
    #   modKnowledgeFmt から移設)と適用(LoadKbSlots。modPipeline/modPipeline2 に
    #   写経されていた LoadKb/FetchKb を畳んだ唯一の実装)を集めた。TrimPlan は
    #   層(a)の回帰網が叩くので required に載せる(Private へ戻すと検査が黙って
    #   消える)。
    "modPipeline4": {
        "closed": False,
        "required": ["TrimPlan", "LoadKbSlots"],
    },
    # modPlayOps: プリフライト診断(T-25)。14章§6が宣言した判定核5本を required
    #   に載せる(裁定書9-3。純核を Private へ戻すと層(a)から検査できなくなる)。
    "modPlayOps": {
        "closed": False,
        "required": [
            "RunPreflight",
            "PfSurvivalOf", "PfPredTypesOf", "PfRefIds", "PfFailCodeOf",
            "CaseIdOfPfLine",
        ],
    },
    # modSparring: PL-04 商談の予行演習(T-27。裁定書8 B-9 で14章§6へ宣言)。
    #   実行制御3本(ResumeSparring / SendSparring / SendToInbox)＋履歴の読み出し
    #   (HistoryOf)に加え、**純核3本を required に載せる**。純核を Private へ
    #   戻すと層(a)の回帰網が消えるため、契約違反として機械で検出する
    #   (modPipeline / modPipeline2 の判定核と同じ扱い)。
    "modSparring": {
        "closed": False,
        "required": [
            "ResumeSparring", "SendSparring", "SendToInbox", "HistoryOf",
            "HistoryJoinOf", "TrimHistoryOf", "CanContinueSparring",
        ],
    },
    "modCaseStore": {
        "closed": False,
        "required": [
            "NewCase", "SaveData", "LoadData", "ResolveStepJson", "SetStatus",
            "InvalidateDownstream", "RepairStates", "FreezeRound",
            # 16章E-06の書込口(裁定書8 A-2で14章§6へ宣言)。
            "SetStepOutcome",
            # 純ロジックの公開(14章§6・裁定書6 項目7)。
            "BuildCaseId", "IsValidCaseId", "CanTransition", "ResolveDataKey",
        ],
    },
    # modCaseStore2: 30,000字契約による分割先(裁定書8 A-2。案件2枚の下位シート
    #   I/O)。modCompanyFile2 と同じく14章§6の公開契約面には載せないため
    #   required は空(closed=False で追加 Public を許容する)。
    "modCaseStore2": {"closed": False, "required": []},
    # modCaseStore3: 同上(T-56。data_key 一覧とラウンド確定の付帯処理)。
    "modCaseStore3": {"closed": False, "required": []},
    # modInboxStore: 受信箱シートの唯一の口(T-25)。実行制御2本に加え、14章§6が
    #   宣言した純ロジック7本を required に載せる(裁定書9-3)。InterestSummaryOf
    #   は FR-17 の集計規約(件数集計・降順・2件以上・上限件数)の唯一の値源で、
    #   Private へ戻すと G71 の回帰網がまるごと消える(裁定書9-1)。
    "modInboxStore": {
        "closed": False,
        "required": [
            "NewInboxItem", "SetInboxJudgement",
            "BuildInboxId", "IsValidInboxId", "CanInboxTransition",
            "JudgementError", "InterestKeyOf", "FmtInterestLine",
            "InterestSummaryOf",
        ],
    },
    # modJudgeStore: 判断台帳CRUD(13章§2.7・裁定書8 B-8)。NewJudgement は
    #   14章§6が予約した唯一の固定名(TJudgement受取・judge_id返却)。
    #   14章§6が宣言した純ロジック4本(採番・ID書式・19章§3の decision enum・
    #   13章§2.7の result enum)も required に載せる(裁定書9-3)。IsValidDecision
    #   は裁定書8 B-8 が名指しした「decision enum検証」そのもの。
    "modJudgeStore": {
        "closed": False,
        "required": [
            "NewJudgement",
            "BuildJudgeId", "IsValidJudgeId", "IsValidDecision",
            "IsValidJudgeResult",
        ],
    },
    "modExportHtml": {
        "closed": False,
        "required": ["GenerateHtmlReport", "GenerateHtmlReportEx"],
    },
    # modGround: 原文照合(裁定書37 B-03・14章§6)。3本とも層(a)の回帰網
    #   (modTestsPure24)が叩くので required に載せる(Private へ戻すと検査が
    #   黙って消える)。
    "modGround": {
        "closed": False,
        "required": ["NormalizeForMatch", "QuoteFound", "GroundNotes"],
    },
    "modExportHearing": {"closed": False, "required": ["BuildHearingSheet"]},
    # modExportPpt: 14章§6の GeneratePpt は **Phase 1.5**(§6の注記・§1の表)。
    # Phase 1 実装で required に入れると未実装ERRORになるため required は空にする
    # (CONTRACTの方針: Phase 1.5関数は required に入れない)。closed=False。
    "modExportPpt": {"closed": False, "required": []},
    # modAppTypes: TCaseCtx 等のドメイン型(Public Type)を持つ app層モジュール
    # (14章§6 の TCaseCtx 定義)。Public Function シグネチャは章に無いため required は空。
    "modAppTypes": {"closed": False, "required": []},
    # modPii: 保存・外部送信・レポート出力前のPII走査本体(16章 E-05・12章§2/§4)。
    # 公開5本は裁定書7 B-5 が14章§6へ宣言した(ScanReport は本文を返さない=NFR-S3)。
    # 呼び出し側(modUICase/modUIInbox/modSparring/modJudgeStore/modExportHtml/
    # modCompanyFile)が前段で必ず通す関係も§6が規定する。
    "modPii": {
        "closed": False,
        "required": ["HasPii", "DetectionCount", "KindsOf", "ScanReport", "MaskText",
                     "SharesLongFragment",
                     # SnippetsOf: 裁定書44 B-1新設。貼付警告文の断片を返す口
                     # (ScanReportと違い実文字列を返す。NFR-S3の対象外)。
                     "SnippetsOf"],
    },
    # modCompanyFile: 企業ドシエファイルの書出・取込(13章§2.8)。公開4本は裁定書7
    #   B-5 が14章§6へ宣言。下位I/Oの modCompanyFile2 は§6の公開契約面に載せない
    #   ため required は空(closed=False で追加 Public を許容する)。
    "modCompanyFile": {
        "closed": False,
        "required": [
            "CompanyFilePath", "ScanCaseForPii", "ExportCompanyFile",
            "ImportCompanyFile",
        ],
    },
    "modCompanyFile2": {"closed": False, "required": []},
    # modCompanyFile3: 14章§6 が W10(裁定書28)で宣言した公開口。純関数は
    #   層(a)の回帰網(modTestsPure20)が叩くので required に載せる(Private へ
    #   戻すと検査が黙って消えるため)。CaseSheetCols / PiiFlagOf は統合W10で
    #   足した pii_flag の写像(16章 E-05(7))。
    "modCompanyFile3": {
        "closed": False,
        "required": [
            "SchemaVersionCurrent", "SchemaVersionOf", "IsSchemaReadable",
            "CompanyFileNameOf", "CompanyDirOf", "CompanyDir",
            "HeaderToCaseRow", "HeaderValueOf", "PickNewerHeader",
            "IsFileNewer", "ReadFileHeader", "AutoSaveCase",
            "ExportExtensions", "ImportExtensions",
            "CaseFingerprint", "FileFingerprint",
            "CaseSheetCols", "PiiFlagOf",
        ],
    },
    # modBootData: 起動時再構成の入口1本(14章§6・裁定書28 W10)。
    "modBootData": {"closed": False, "required": ["RebuildCaseCache"]},
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
    # modUICase: 11章§5「日本語ラベル⇔enumの変換は modUICase の共通変換表
    #   (19章と一致必須)のみで行う」＋13章§2.2 逆シリアライズ規約1が名指しする
    #   `modUICase.SerializeSheet(stepNo)`。EnumPairsCsv は tools/enum_check.py が
    #   静的評価する**変換表そのもの**であり、Private化・改名すると17章§4-2の
    #   一致検査が対象を失う(=検査が無言で消える)ため required に載せる。
    #   W15 Round3(裁定書41 §1): SerializeSheet は modUICase2.SerializeStep を
    #   呼ぶだけの別名で呼出元0件だったため撤去した(本番は SerializeStep 直呼び)。
    "modUICase": {
        "closed": False,
        "required": ["EnumPairsCsv", "EnumJa", "EnumEn", "EnumLabels",
                     "ApplyEnumValidation"],
    },
    # 以下4本の公開口は 14章§6 が宣言していない(ui層の画面ハンドラは章の
    # 契約面に載っていない)。CONTRACT へは完全性自己検査
    # (MODULE_REGISTRY⇔CONTRACT)を満たすために required=[] で登録する。
    # closed=False なので追加 Public は許容する。
    "modUISheet": {"closed": False, "required": []},
    # modUIViewport: 裁定書26 B。modBoot からの結線先(ApplyFullScreen)と、
    # 退避値へ戻す唯一の口(RestoreScreen)を required で固定する。
    "modUIViewport": {
        "closed": False,
        "required": ["ApplyFullScreen", "RestoreScreen"],
    },
    # modUIGuide: 裁定書14 裁定6。起動の入口 StartTourIfFirstRun と再視聴の
    # RestartTour、操作ガイドのボタン EnsureGuideButtons を required で固定する
    # (modBoot / modUIHome / 図形の OnAction の結線先そのものであり、改名・
    #  Private化はその場で結線が切れる)。
    "modUIGuide": {
        "closed": False,
        "required": ["StartTourIfFirstRun", "RestartTour", "EnsureGuideButtons",
                     "AdvActionRow"],
    },
    # modUIToast: 裁定書17 H2/H4。modUIHome の ShowWarning と主要4ボタンの
    # 成功経路が呼ぶ結線先そのものであり、改名・Private化はその場で案内が
    # 消えるため required で固定する(HideToast は Application.OnTime の
    # コールバックなので Public 必須)。
    "modUIToast": {
        "closed": False,
        # ShowResearchPrompts は裁定書22 で**廃止**(14章§6に「廃止」として残す)。
        "required": ["ShowToast", "ShowNext", "HideToast", "CancelToast",
                     "WarnLine"],
    },
    "modUIHome": {"closed": False, "required": []},
    # modUIHome2: 30,000字契約による modUIHome の分割先(17章§7 Z-13)。HOMEと
    # S1～S4シートの図形ボタンの OnAction 結線先そのもの(改名・Private化はその場で
    # 結線が切れる)なので、移設した公開ハンドラを required で固定する。
    "modUIHome2": {
        "closed": False,
        "required": [
            # W15 Round3(裁定書41 §1): v3.2 で廃止した HOMEシートの
            # [くわしい操作]の結線先だった7本(HomeNewCase / HomeOpenCaseInput /
            # HomeOpenInbox / HomeOpenFeedback / HomeOpenJudgeLog /
            # HomeOpenSparring / HomePreflightAll)を撤去した。同じ操作は
            # ナビの新規モード・使い方タブ⑦上級・受信箱の一括診断に残っている。
            "HomeRunAll", "HomeRunS1", "HomeRunS2", "HomeRunS3", "HomeRunS4",
            "HomeFreezeRound", "HomeExportHtml", "HomeBuildHearing",
            "HomeCompanySave", "HomeCompanyOpen", "HomeReloadKnowledge",
        ],
    },
    "modUICase2": {"closed": False, "required": []},
    "modUICase3": {"closed": False, "required": []},
    "modUICase4": {"closed": False, "required": []},
    "modUICase5": {"closed": False, "required": []},
    "modUICaseFmt": {"closed": False, "required": []},
    "modUIInbox": {"closed": False, "required": []},
    "modUISparring": {"closed": False, "required": []},
    # ---- test 層 ----
    "modTestRunner": {
        # 本Lintと17章§4-1のランナー要件で公開口を確定させているため closed=True。
        "closed": True,
        "required": [
            "ResetTests", "Check", "Failures", "ReportText", "RunAllPureTests",
            "SetExpectedCount",
            # 裁定書14 裁定5: ブック内テスト実行が ps1 と同じ4条件を判定する
            # ための計数の読み出し口(集計の仕方は変えない)。
            # W15 Round3(裁定書41 §1): PassCount は4条件のどれにも使われず
            # 呼出元0件だったため撤去した(合格数は ExecutedCount - FailCount)。
            "FailCount", "SkipCount", "ExecutedCount",
        ],
    },
    # modTestsExcel: 14章§6のtest層契約(層(b)=実Excel E2Eスモークの入口。17章T-47)。
    "modTestsExcel": {"closed": False, "required": ["RunAllExcelTests"]},
    # modTestsRunnerUi: 17章 T-48(裁定書14 裁定5)。操作ガイドの[テストを実行]の
    # OnAction。引数なし公開1本のみ(Alt+F8からも実行できるようにするため)。
    "modTestsRunnerUi": {"closed": False, "required": ["RunAllTestsFromBook"]},
    # modTestsExcel2: 30,000字契約(12章§2)による modTestsExcel の分割先。
    # wintest からの入口は RunAllExcelTests のままで、本数だけ合流させる。
    "modTestsExcel2": {"closed": False, "required": ["RunExcelTests2"]},
    # ---- 裁定書35 §1: W12-A(HTML画面ホスト層)の22件登記。CONTRACT完全性
    #   (MODULE_REGISTRY⇔CONTRACT)を満たすために追加した。required に入れて
    #   よいのは14章§6に逐語で書かれた公開名だけ(「等」「系」表記は入れない)。
    #   全エントリ closed=False(内部の追加Publicは許容する)。
    # frmNaviHtml: UserForm(.frm)。14章§6に口の宣言なし(HTML画面の器そのもの)。
    "frmNaviHtml": {"closed": False, "required": []},
    # modBootNavi: 14章§6 W12-A表(1542行)。LaunchIfHtml は逐語。
    #   RegisterDefault系は「系」表記のため required に入れない。
    "modBootNavi": {"closed": False, "required": ["LaunchIfHtml"]},
    # modConfig: 14章§6に口の宣言なし(configキーの値源はconfigシートで、
    #   本モジュールの公開関数は§6の登記表に載っていない)。
    "modConfig": {"closed": False, "required": []},
    # modGatewayDirect: 14章§6本文(361-383行)。direct経路の6本は逐語のPublic宣言。
    "modGatewayDirect": {
        "closed": False,
        "required": ["CallDirect", "BackoffMs", "RetryBudgetFor",
                     "ParseKeyLine", "IsOSeriesModel", "BuildRequestBody"],
    },
    # modGatewayLink: 14章§6本文123行「`modGatewayRPN.DirectStep` が呼ぶのは接続
    #   モジュール `modGatewayLink.CallDirect`(14章§6と同一契約)」で逐語に名指し。
    "modGatewayLink": {"closed": False, "required": ["CallDirect"]},
    # modGatewayRPN2: 14章§6に口の宣言なし(W12-A裁定書34。modGatewayRPN の
    #   30,000字契約分割先で、§6は元モジュール名でしか登記していない)。
    "modGatewayRPN2": {"closed": False, "required": []},
    # modLog: 14章§6本文(353-359行)。TruncDetail / ShouldRotate は逐語。
    "modLog": {
        "closed": False,
        "required": ["TruncDetail", "ShouldRotate"],
    },
    # modMockLlm: 14章§6本文(1427-1443行)。4本は逐語のPublic宣言。
    "modMockLlm": {
        "closed": False,
        "required": ["MockResponse", "ResponseById", "FaultResponse",
                     "ResetFaultOnce"],
    },
    # modNaviActions: 14章§6 W12-A表(1539行)。4本は逐語(「等」の後続は入れない)。
    "modNaviActions": {
        "closed": False,
        "required": ["ActPasteMaterial", "ActSaveMaterials", "ActRunPipeline",
                     "ActExportReport"],
    },
    # modNaviActions2: 14章§6に口の宣言なし。
    "modNaviActions2": {"closed": False, "required": []},
    # modNaviChat: 14章§6 W12-A表(1540行)。4本は逐語。
    "modNaviChat": {
        "closed": False,
        "required": ["BuildChatSystem", "Ask", "History", "Clear"],
    },
    # modNaviHost: 14章§6 W12-A表(1537行)。5本は逐語。
    "modNaviHost": {
        "closed": False,
        "required": ["OpenNaviTool", "HostDispatchPending", "HostRequestJson",
                     "CloseNaviTool", "HostWorkbookClose"],
    },
    # modNaviJson: 14章§6に口の宣言なし。
    "modNaviJson": {"closed": False, "required": []},
    # modNaviState: 14章§6 W12-A表(1538行)。4本は逐語。
    "modNaviState": {
        "closed": False,
        "required": ["BuildAppState", "BuildCaseState", "BuildStageList",
                     "BuildPrompts"],
    },
    # modNaviState2: 14章§6に口の宣言なし。
    "modNaviState2": {"closed": False, "required": []},
    # modNaviStore: 14章§6 W12-A表(1541行)。9本は逐語。
    "modNaviStore": {
        "closed": False,
        "required": ["ListCases", "ListInbox", "ListJudgements", "LogRowsOf",
                     "AppendFeedback", "SetDisplayName", "SetArchived",
                     "ExportCaseJson", "ImportCaseJson"],
    },
    # modRibbonSim: 14章§6に口の宣言なし(W11-c裁定書33 C-3。リボン応答抽出の
    #   逐語模擬はMODULE_REGISTRYの説明にのみ載る)。
    "modRibbonSim": {"closed": False, "required": []},
    # modRibbonWire: 14章§6に口の宣言なし(W11-c裁定書33 C-4。同上)。
    "modRibbonWire": {"closed": False, "required": []},
    # modTestsExcelNavi: 14章§6に口の宣言なし(W12-A裁定書34 §1.1。層(b)の
    #   HTML画面テスト。§6は名前を登記していない)。
    "modTestsExcelNavi": {"closed": False, "required": []},
    # modTestsPureDev: 14章§6に口の宣言なし(dev専用の純層テスト。配布物には
    #   載らない=ship:false)。
    "modTestsPureDev": {"closed": False, "required": []},
    # modTestsPureHook: 14章§6に口の宣言なし(dev専用の純層テストの接続点。
    #   modGatewayLink と同じ同名2ソース)。
    "modTestsPureHook": {"closed": False, "required": []},
    # modTestsPureNavi: 14章§6に口の宣言なし(W12-A裁定書34 §1.1。純層のHTML
    #   画面テスト2本)。
    "modTestsPureNavi": {"closed": False, "required": []},
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
    # modGatewayDirect: 裁定書27 W9-B4 が kernel32 `Sleep`(`Declare PtrSafe`)の
    #   撤去と `Application.Wait` への置換を命じたため、バックオフ待ちの1行だけ
    #   Excelトークンが要る。**許可の幅はその1行**であり、シート・ブックには
    #   触れない(direct経路は 裁定書27 W9-B5 で配布ビルドからも外れ dev専用)。
    "modGatewayDirect",
    # modConfig: config シートの読み書きが責務そのもの(13章§2.3)。
    "modConfig",
    # modLog: err_log / usage_log / run_log シートへの記録が責務(13章§2.4)。
    "modLog",
    # 以下 store 系。案件・受信箱・判断台帳・ナレッジブックのシートI/Oが責務
    # (12章§2 の R4 但し書き「store系モジュール内は自身の責務範囲で可」)。
    "modCaseStore", "modCaseStore3", "modInboxStore", "modJudgeStore", "modKnowledge",
    # modCaseStore2: 上と同一責務の分割先(裁定書8 A-2。案件一覧 と case_data の
    #   下位シートI/Oだけを切り出したもの)。許可の幅は modCaseStore と同じ
    #   「本体ブックの案件2枚」で広がっていない。
    "modCaseStore2",
    # modCaseRead: 案件一覧の**読取専用**API(裁定書7 B-7・14章§6 ReadCaseCtx)。
    #   許可の幅は modCaseStore と同じ「案件一覧の1枚」で、書込は一切持たない
    #   (Range への代入が入ったら NFR-S7 の書込口検査と本注記の両方に違反する)。
    "modCaseRead",
    # modCompanyFile: 企業ドシエファイル(1社1.xlsx)の書出・取込(13章§2.8)。
    "modCompanyFile",
    # modCompanyFile2: 上と同一責務の分割先(T-29の実装が1本では30,000字契約を
    #   超えたため、下位のブック・シートI/Oだけを切り出したもの)。許可の幅は
    #   modCompanyFile と同じ「企業ドシエファイルのシートI/O」で広がっていない。
    #   分割の是非は12章§2のモジュール一覧に載っていないためWARNとして残る。
    "modCompanyFile2",
    # modCompanyFile3: 上と同一責務の3本目(裁定書28 W10)。許可の幅は
    #   modCompanyFile と同じ「企業ファイル側のシートI/O」で広がっていない
    #   (本体ブックの値は modCaseRead / modCaseStore 経由でしか読まない)。
    "modCompanyFile3",
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
# NFR-S7① の「セルへの直接書込」検査を掛けない代入先(裁定書34 §1.4・W12-A)。
# frmNaviHtml が触る `field` は **HTML(WebBrowser)の textarea 要素**であって
# Excel のセルではない。SetCellSafe(Excelのセルへ外部由来テキストを書くときの
# 唯一の口)を通しようがないので、変数名を名指しで外す。
# 変数名でしか区別できないため、**その変数が DOM 要素であることが読んで分かる
# モジュール**に限る(ここでは frmNaviHtml の2箇所だけ)。
CELL_WRITE_DOM_TARGETS = {("frmNaviHtml", "field")}
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
    "ヒアリングシート", "商談の予行演習",
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
#
# 2件目(裁定書14 裁定5・17章 T-48): 操作ガイドの[テストを実行]ボタンの OnAction
#   文字列 "modTestsRunnerUi.RunAllTestsFromBook" を ui層(modUIGuide)が図形へ
#   配線する。禁止理由(配布物からテストを外せなくなる)が当てはまらないことは
#   modMockLlm と同じで、テスト一式は build/modules.json に登録されて vba_src へ
#   焼き込まれる**配布物同梱の自己検査**であり(12章§2)、社内PCにはターミナルが
#   無いためブック内から回す口が唯一の検問になる。粒度は完全一致1件のみ。
R1_TEST_LAYER_EXCEPTIONS = {
    ("modGatewayRPN", "modMockLlm", "MockResponse"),
    ("modUIGuide", "modTestsRunnerUi", "RunAllTestsFromBook"),
    # 裁定書34 §1.2(W12-A)。HTML画面の[テストを実行](action="run_tests")。
    #   使い方タブの図形ボタン(modUIGuide の上の行)とまったく同じ性格の参照で、
    #   17章 T-48 が「利用者がブックの中だけで検問を回せること」を求めている以上、
    #   画面がテストの入口を1つ持つのは避けられない。
    #   幅を広げないための条件を2つ守っている:
    #     (1) 呼ぶ先は modTestsRunnerUi の2本だけ(modTestRunner / modTestsExcel を
    #         直に叩かせない。合否の4条件と集計はテスト層の中に閉じる)
    #     (2) 配布物から test 層を外したときは modTestsRunnerUi ごと消えるので、
    #         参照が残らない(ship 集合の判断は build/modules.json 側)
    ("modNaviActions2", "modTestsRunnerUi", "RunAllTestsHeadless"),
    ("modNaviActions2", "modTestsRunnerUi", "LastReportText"),
}

# 型落ち検出: Dim/Static文
DIM_STMT_PATTERN = re.compile(r"^(Dim|Static)\s+(.*)$", re.IGNORECASE)
AS_KEYWORD_PATTERN = re.compile(r"\bAs\b", re.IGNORECASE)
AS_INTEGER_PATTERN = re.compile(r"\bAs\s+Integer\b", re.IGNORECASE)
# UserForm のイベント署名(VBA が形を決めているので Integer を Long にできない)。
# 裁定書34 §1.4(W12-A)。ここに足すのは「VBAの仕様でそう書くしかない」ものだけ。
FORM_EVENT_INTEGER_SIG = re.compile(
    r"^(?:Public\s+|Private\s+)?Sub\s+UserForm_(QueryClose|Error)\s*\(", re.IGNORECASE)

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


# ==============================================================================
# .frm(UserForm)のデザイナヘッダ(裁定書34 §1.4)
# ------------------------------------------------------------------------------
# VBE の .frm エクスポートは
#     VERSION 5.00
#     Begin {GUID} frmXxx
#        Caption = "..."          <- 画面の属性。ソースではない
#     End
#     Attribute VB_Name = "frmXxx"
#     ...(属性行が続く)
#     Option Explicit             <- ここからがソース
# という形をしている。前半をそのまま lint に食わせると「宣言の位置」等の検査が
# 誤爆するので、**先頭の連続した非ソース行だけ**を空行へ置き換えて落とす
# (行番号がずれないよう、削らずに空にする)。
# 本文の途中に現れる `Attribute mBrowser.VB_VarHelpID = -1` のような行は
# 「先頭の連続」ではないので残る(.bas でも同じ扱い)。
FORM_HEADER_LINE = re.compile(
    r"^\s*(VERSION\s|Begin\b|End\s*$|Attribute\s|[A-Za-z_]\w*\s*=\s|\})",
    re.IGNORECASE)


def strip_form_header(raw_text: str) -> str:
    lines = raw_text.split("\n")
    last_attr = -1
    for i, line in enumerate(lines):
        s = line.strip()
        if s == "":
            continue
        if not FORM_HEADER_LINE.match(line):
            break
        if s.lower().startswith("attribute "):
            last_attr = i
    if last_attr < 0:
        return raw_text
    return "\n".join([""] * (last_attr + 1) + lines[last_attr + 1:])


def load_module(path: Path, src_root: Path) -> ModuleInfo:
    raw_text = path.read_text(encoding="utf-8", errors="replace")
    if path.suffix.lower() == ".frm":
        raw_text = strip_form_header(raw_text)
    raw_lines = raw_text.splitlines()
    stmts = iter_statements(raw_lines)

    vb_name = ""
    for line in path.read_text(encoding="utf-8", errors="replace").split("\n")[:20]:
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

    # 18章§4.4 の専用閾値(17章 T-35 のDoDが本ツールへ委ねた検査)。
    if TEMPLATE_MODULE_PATTERN.match(module_name_for_display(info)) and n > TEMPLATE_MAX_CHARS:
        info.add(
            "ERROR", 1,
            f"テンプレモジュールが{n}字で18章§4.4の上限{TEMPLATE_MAX_CHARS}字を超過"
            f"(次番の modHtmlTemplateN へ**関数単位で**切り出すこと。関数名は変えない)",
        )


# 行長バイト検査の実行痕跡(骨抜き防止)。check_line_cp932_bytes が実際に
# 走ったモジュールを記録し、run_lint の最後で全モジュール分そろっているかを
# 照合する。run_lint の呼び出し行を消すと「1件も走っていない」が検出され、
# lint は緑にならずERRORで落ちる(検査を外して赤を消す抜け道を塞ぐ)。
LINE_BYTES_SCANNED: set[str] = set()


def check_line_cp932_bytes(info: ModuleInfo) -> None:
    """1物理行のCP932バイト長 <= MAX_LINE_CP932_BYTES(ERROR)。

    12章§2の30,000字契約が「文字数」なのに対し、こちらは**バイト**である点が
    肝(裁定書19 H7)。CP932へ変換できない文字は check_cp932_safe が別途ERROR
    で捕捉済みのため、ここでは errors="replace" で1バイトに丸めて数えるだけに
    留める(同じ文字で2つのERRORを出して本命を埋もれさせない)。
    コメント行も対象: VBEの行分断はコメントでも同じように起きる。
    """
    LINE_BYTES_SCANNED.add(info.relpath.as_posix())
    for lineno, line in enumerate(info.raw_text.splitlines(), 1):
        nbytes = len(line.encode("cp932", errors="replace"))
        if nbytes <= MAX_LINE_CP932_BYTES:
            continue
        info.add(
            "ERROR", lineno,
            f"1物理行が{nbytes}バイト(CP932)で上限{MAX_LINE_CP932_BYTES}バイトを超過"
            f"(文字数は{len(line)}字)。VBEの1行上限1,023は文字数ではなくバイト数です。"
            f"VBEが行を分断し「SubまたはFunctionが定義されていません」になります。"
            f"文字列リテラルの境界で s = s & \"...\" を複数行へ分けてください",
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


# ==============================================================================
# 仕様側(docs)のプロンプト本文に対する同一基準の検問(裁定書6 A-2)
# ------------------------------------------------------------------------------
# check_cp932_safe は .bas しか見ない。しかし 15章・docs/08 は「本文の正」であり、
# 実装は そこから一字一句 写す(17章 T-23)。したがって仕様側にCP932外の文字が
# 混ざると、写した瞬間に .bas 側が赤くなる=原因が下流に出る。W2aでは実際に
# 15章のコードフェンス内のU+301Cが23件のERRORとして実装側に現れた。
# 再混入を仕様側で止めるため、プロンプト本文を持つ章のコードフェンス内を
# .bas と同じ _cp932_encodable で検査する。
# 対象はフェンス内のみ(フェンス外の解説文は LLM へ送らないため対象外。
# 15章§10.1(a) の抽出規約と同じ境界を使う)。
# ==============================================================================
DOCS_PROMPT_FILES = (
    "docs/spec/15_プロンプトとJSONスキーマ.md",
    "docs/08_ドシエ収集プロンプト集.md",
)
_DOCS_FENCE_RE = re.compile(r"^\s*```")


def check_docs_prompt_cp932(repo_root: Path) -> list[tuple[str, int, str]]:
    """15章・docs/08 のコードフェンス内をCP932基準で検査する。

    戻り値は (相対パス, 行番号, メッセージ) のリスト(すべてERROR相当)。
    """
    out: list[tuple[str, int, str]] = []
    for rel in DOCS_PROMPT_FILES:
        path = repo_root / rel
        if not path.exists():
            out.append((rel, 1, "プロンプト本文の正である章が見つかりません"
                                "(パスを変えたら DOCS_PROMPT_FILES を更新してください)"))
            continue
        in_fence = False
        for lineno, line in enumerate(path.read_text(encoding="utf-8").split("\n"), 1):
            if _DOCS_FENCE_RE.match(line):
                in_fence = not in_fence
                continue
            if not in_fence:
                continue
            bad = sorted({ch for ch in line if not _cp932_encodable(ch)})
            if not bad:
                continue
            shown = " ".join(f"{ch}(U+{ord(ch):04X})" for ch in bad[:6])
            out.append((
                rel, lineno,
                f"プロンプト本文(コードフェンス内)にCP932に無い文字: {shown} 。"
                f"15章§0 原則7。ここを写した .bas がVBEで '?' に化けます"))
    return out


# ==============================================================================
# 製品へ焼く文字列の危険6字検査(裁定書22 m6・11章§7.1)
# ------------------------------------------------------------------------------
# build/build_rpn.py と build/sheets_main.json の**文字列**は、そのまま配布ブックの
# セルへ焼かれる。CP932 に無い字(とりわけ U+301C WAVE DASH)が混ざると、Windows で
# 開いた利用者の画面に化けた字が出る。.bas と同じ _CP932_DENY(危険6字)で検査する。
#   ・.py は Python の字句解析で**文字列リテラルだけ**を見る(コメントは対象外だが、
#     写した瞬間に化けるので実運用ではコメントも揃えてある)
#   ・.json は全ての文字列(キーと値)を見る
# 全文の CP932 検査にしない理由: この2ファイルはコード生成器であり、日本語の
# 説明文に CP932 外の記号(┈ 等)を意図して使う箇所がある。**化けると実害が出る
# 6字だけ**を ERROR にする(検査を骨抜きにせず、偽陽性も出さない境界)。
# ==============================================================================
BUILD_TEXT_FILES = ("build/build_rpn.py", "build/sheets_main.json")


def _deny_hits(text: str) -> list[str]:
    return sorted({ch for ch in text if ch in _CP932_DENY})


def check_build_strings(repo_root: Path) -> list[tuple[str, int, str]]:
    """build/ の製品文字列に危険6字が混ざっていないか(すべてERROR相当)。"""
    out: list[tuple[str, int, str]] = []
    for rel in BUILD_TEXT_FILES:
        path = repo_root / rel
        if not path.exists():
            out.append((rel, 1, "製品へ焼く文字列を持つファイルが見つかりません"
                                "(パスを変えたら BUILD_TEXT_FILES を更新してください)"))
            continue
        text = path.read_text(encoding="utf-8")
        if rel.endswith(".py"):
            try:
                toks = list(tokenize.generate_tokens(io.StringIO(text).readline))
            except (tokenize.TokenError, IndentationError, SyntaxError) as exc:
                out.append((rel, 1, f"字句解析できませんでした: {exc}"))
                continue
            for tok in toks:
                if tok.type != tokenize.STRING:
                    continue
                bad = _deny_hits(tok.string)
                if not bad:
                    continue
                shown = " ".join(f"{ch}(U+{ord(ch):04X})" for ch in bad[:6])
                out.append((
                    rel, tok.start[0],
                    f"製品へ焼く文字列に CP932 の危険字: {shown} 。"
                    f"11章§7.1。波ダッシュは ～(U+FF5E)で書いてください"))
        else:
            try:
                data = json.loads(text)
            except ValueError as exc:
                out.append((rel, 1, f"JSONとして読めませんでした: {exc}"))
                continue
            found: set[str] = set()

            def walk(node):
                if isinstance(node, dict):
                    for k, v in node.items():
                        found.update(_deny_hits(k))
                        walk(v)
                elif isinstance(node, list):
                    for v in node:
                        walk(v)
                elif isinstance(node, str):
                    found.update(_deny_hits(node))

            walk(data)
            for ch in sorted(found):
                lineno = 1
                for i, line in enumerate(text.split("\n"), 1):
                    if ch in line:
                        lineno = i
                        break
                out.append((
                    rel, lineno,
                    f"製品へ焼く文字列に CP932 の危険字: {ch}(U+{ord(ch):04X}) 。"
                    f"11章§7.1。波ダッシュは ～(U+FF5E)で書いてください"))
    return out


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
# W3(T-30): ui層は図形の生成とOnAction配線を modUISheet.EnsureButton に一本化した
# ため、`.OnAction = "modX.Y"` の直接代入はその関数の中の【変数経由】1箇所だけに
# なった。配線先の文字列リテラルは呼び出し側の引数に現れるので、そちらも収集
# しないと本検査が丸ごと空振りする(=関所の検査が無言で消える)。
# 収集対象は EnsureButton 呼び出しの引数に現れる "modX.Y" 形の文字列リテラルのみ。
ONACTION_WIRING_PATTERN = re.compile(
    r"\bEnsureButton\b.*?\"(mod[A-Za-z]\w*\.[A-Za-z_]\w*)\"", re.IGNORECASE)
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
            for m2 in ONACTION_WIRING_PATTERN.finditer(stmt):
                wired.setdefault(m2.group(1), (module_name_for_display(info), lineno))

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


# ==============================================================================
# W9.4: UDT(ユーザー定義型)・配列を ByVal で渡す禁止(実機事故)
# ------------------------------------------------------------------------------
# 実機Mac Excelで[テストを実行]を押すと modPromptsOps.AsmS1User(ByVal ctx As
# TCaseCtx, ...) で「ユーザー定義型を ByVal で渡すことはできません」の
# コンパイルエラーになった。VBAの規則: ユーザー定義型(Type宣言した型)と配列は
# 手続き引数として ByVal で渡せない(ByRef 必須)。LibreOffice Basic はこれを
# 許すため lo-compile(層(c))が見逃す — 17章 T-58 W9.4。
#
# 検査の作り: src 全体から Type 宣言名(Public/Private 問わず)を集めてから、
# 全モジュールの「ByVal <name> As <型>」を当たり、型名がその集合にあれば
# ERROR。「ByVal <name>() As ...」(配列引数)は型を問わず常にERROR(VBAは
# どんな要素型の配列も ByVal で渡せない)。
# ==============================================================================
TYPE_DECL_PATTERN = re.compile(
    r"^(?:Public\s+|Private\s+)?Type\s+([A-Za-z_]\w*)", re.IGNORECASE
)
BYVAL_PARAM_PATTERN = re.compile(
    r"\bByVal\s+(\w+)\s*(\(\s*\))?\s+As\s+(\w+)", re.IGNORECASE
)


def _collect_type_names(infos: list[ModuleInfo]) -> set[str]:
    names: set[str] = set()
    for info in infos:
        for _lineno, stmt in info.statements:
            m = TYPE_DECL_PATTERN.match(stmt)
            if m:
                names.add(m.group(1).lower())
    return names


def check_udt_byval_param(infos: list[ModuleInfo]) -> None:
    """VBAが禁止する「UDT/配列の ByVal 引数」を全モジュールから検出する。"""
    type_names = _collect_type_names(infos)
    for info in infos:
        for lineno, stmt in info.statements:
            masked = _blank_string_literals(stmt)
            for m in BYVAL_PARAM_PATTERN.finditer(masked):
                pname, arr_mark, tname = m.group(1), m.group(2), m.group(3)
                if arr_mark:
                    info.add(
                        "ERROR", lineno,
                        f"配列引数「{pname}()」を ByVal で宣言しています。VBAは配列を"
                        f"ByVal で渡せません(実機Excelでコンパイルエラー。LibreOffice"
                        f"はこれを許すため lo-compile では検出できません)。"
                        f"ByRef へ直してください: 「{stmt.strip()[:80]}」",
                    )
                elif tname.lower() in type_names:
                    info.add(
                        "ERROR", lineno,
                        f"ユーザー定義型「{tname}」の引数「{pname}」を ByVal で宣言"
                        f"しています。VBAは Type 宣言した型(UDT)を ByVal で渡せません"
                        f"(実機Excelで「ユーザー定義型を ByVal で渡すことはできません」の"
                        f"コンパイルエラー。LibreOffice はこれを許すため lo-compile では"
                        f"検出できません)。ByRef へ直してください: 「{stmt.strip()[:80]}」",
                    )


# 上のルールの自己テスト(骨抜き防止)。正例=ByRef(UDT)/ByVal(非UDTスカラー)は
# findingが出てはいけない。負例=ByVal(UDT)/ByVal(配列、型は問わない)は必ず
# ERRORが出なければならない。run_lint の冒頭で毎回走らせる。
_UDT_BYVAL_TYPE_SRC = "Public Type TFoo\n    x As Long\nEnd Type\n"
_UDT_BYVAL_SELFTEST_OK = [
    "Public Function Bar(ByRef f As TFoo) As String\nEnd Function",
    "Public Function Baz(ByVal n As Long, ByVal s As String) As String\nEnd Function",
]
_UDT_BYVAL_SELFTEST_NG = [
    "Public Function Bar(ByVal f As TFoo) As String\nEnd Function",
    "Public Function Arr(ByVal xs() As Long) As String\nEnd Function",
    "Public Function ArrUdt(ByVal fs() As TFoo) As String\nEnd Function",
]


def _udt_byval_probe(type_src: str, proc_src: str) -> list[ModuleInfo]:
    type_info = ModuleInfo(
        path=Path("selftest_types.bas"), relpath=Path("selftest_types.bas"),
        raw_text=type_src, vb_name="selftest_types", filename_stem="selftest_types",
        statements=iter_statements(type_src.splitlines()),
    )
    proc_info = ModuleInfo(
        path=Path("selftest_proc.bas"), relpath=Path("selftest_proc.bas"),
        raw_text=proc_src, vb_name="selftest_proc", filename_stem="selftest_proc",
        statements=iter_statements(proc_src.splitlines()),
    )
    return [type_info, proc_info]


def _selftest_udt_byval_param() -> list[str]:
    """正例/負例を check_udt_byval_param に通し、食い違いを文字列で返す。"""
    problems: list[str] = []
    for src in _UDT_BYVAL_SELFTEST_OK:
        infos = _udt_byval_probe(_UDT_BYVAL_TYPE_SRC, src)
        check_udt_byval_param(infos)
        found = [f for i in infos for f in i.findings if f.level == "ERROR"]
        if found:
            problems.append(
                f"UDT/配列ByVal禁止ルールの正例を誤検知しました: {src!r} -> "
                f"{[f.message for f in found]}"
            )
    for src in _UDT_BYVAL_SELFTEST_NG:
        infos = _udt_byval_probe(_UDT_BYVAL_TYPE_SRC, src)
        check_udt_byval_param(infos)
        found = [f for i in infos for f in i.findings if f.level == "ERROR"]
        if not found:
            problems.append(f"UDT/配列ByVal禁止ルールの負例を検出できませんでした: {src!r}")
    return problems


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
            if FORM_EVENT_INTEGER_SIG.match(stmt.strip()):
                # 裁定書34 §1.4(W12-A): UserForm のイベントは**VBAが署名を決めて
                # いる**。`UserForm_QueryClose(Cancel As Integer, CloseMode As
                # Integer)` を Long にすると VBE がイベントとして結びつけず、
                # 閉じる操作を捕まえられなくなる(=画面が勝手に閉じる)。
                # 直せない形なので、**この2つのイベント署名だけ**を通す。
                continue
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
        "Point "
        # 裁定書27 W9-B7(b) で追加した主要語。MS-VBAL 3.3.5.x の
        # <reserved-name> / <statement-keyword> / <rem-keyword> に現れ、
        # 実機Excel VBA が識別子として拒否する(LibreOffice Basic は通すため
        # LOゲートでは検出できない=同じ死角)。既存コードに該当は無かった。
        "Name Type Select Time Left Right Mid Format Error Object Property "
        "Step Text Value Width Height Circle Line Print Put Get Write Close "
        "Open Seek Lock Unlock Reset"
    ).split()
}

MSVBAL_DECL_PATTERNS = (
    re.compile(r"\b(?:Dim|Static|ReDim(?:\s+Preserve)?|Private|Public|Global)\s+"
               r"([A-Za-z_]\w*)\s*(?:\(|\bAs\b|$|,)", re.IGNORECASE),
    re.compile(r"\bConst\s+([A-Za-z_]\w*)\b", re.IGNORECASE),
    # `Optional ByVal name As String` のように修飾子が続く形は、finditer が
    # 「Optional ByVal」を1回で食べて name を見落とす。修飾子の並びをまとめて
    # 読み飛ばしてから識別子を捕まえる(裁定書27 W9-B7(b))。
    re.compile(r"\b(?:ByVal|ByRef|Optional|ParamArray)"
               r"(?:\s+(?:ByVal|ByRef|ParamArray))*\s+([A-Za-z_]\w*)\b",
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


# ==============================================================================
# 16進リテラルの Long 接尾辞(裁定書8 A-5。W2bの実バグの再発防止)
# ------------------------------------------------------------------------------
# VBAは **16進リテラルの型を桁数で決める**。1～4桁は符号付き Integer なので
# &H8000 以上は負値へ化ける(&H9FFF = -24577)。W2b では modPii の漢字域判定
# `cp >= &H4E00 And cp <= &H9FFF` がこれで常に False になり、漢字姓の検知が
# 全滅した(黒箱テストが発見。実害はPII走査の素通り=16章E-05の骨抜き)。
# 検査を置く理由は「LOでは捕まらないから」ではない: LibreOffice も同じ型付け
# をするため、**当該の値を通るテストが1本でもあれば** run_lo_tests は落ちる
# (実測: modPii の `&H9FFF&` から接尾辞を外すと純ロジック3本が FAIL)。
# つまり LO ゲートの検出は**テストの網掛かり次第の間接検出**であり、網の無い
# 定数・新規コードは素通りする。この検査はその依存を断ち、**静的に確実へ**
# 変えるために置く(接尾辞 `&` を付ければ Long として解釈され値が保たれる)。
#
# 規則: 値が &H8000 以上の16進リテラルは接尾辞 `&` を必須とする。
#   ・1～4桁 … `&` が無いと **値が変わる**(上記の実バグそのもの)
#   ・5～8桁 … VBAは既に Long として読むので値は変わらないが、桁数で型が
#              変わる読み手泣かせを残さないため同じ規則を課す
#              (modPii の注記「codebase 全体で &H....& に統一」が house rule)。
#   ・`%`(Integer)等の別の型接尾辞は `&` ではないので違反として扱う。
# 例外: **ChrW() の唯一の引数**に置かれた素の16進リテラルだけは WARN に落とす。
#   ChrW の引数は -32768～65535 を受け、負値は符号なし16bit相当として解釈される
#   ため、この位置に限り値が保たれる(化けない)。とはいえ桁数依存の読みにくさは
#   同じなので、黙って通さず WARN として毎回表に出す。
# ==============================================================================
HEX_SUFFIX_MIN_VALUE = 0x8000
HEX_LITERAL_PATTERN = re.compile(r"&H([0-9A-Fa-f]+)(&?)")
HEX_IN_CHRW_PATTERN = re.compile(
    r"\bChrW\s*\(\s*&H[0-9A-Fa-f]+\s*\)", re.IGNORECASE)


def check_hex_literal_suffix(info: ModuleInfo) -> None:
    for lineno, stmt in info.statements:
        masked = _blank_string_literals(stmt)
        chrw_spans = [m.span() for m in HEX_IN_CHRW_PATTERN.finditer(masked)]
        for m in HEX_LITERAL_PATTERN.finditer(masked):
            digits, suffix = m.group(1), m.group(2)
            if suffix == "&":
                continue
            value = int(digits, 16)
            if value < HEX_SUFFIX_MIN_VALUE:
                continue
            if any(s <= m.start() and m.end() <= e for s, e in chrw_spans):
                info.add(
                    "WARN", lineno,
                    f"16進リテラル &H{digits} に Long接尾辞 '&' がありません"
                    f"(ChrW の引数位置なので値は保たれますが、桁数で型が変わる"
                    f"読みにくさは同じです。&H{digits}& と書いてください): "
                    f"「{stmt.strip()[:80]}」",
                )
                continue
            if len(digits.lstrip("0")) <= 4:
                why = (f"VBAは4桁以下の16進を符号付きIntegerとして読むため "
                       f"&H{digits} は {value - 0x10000} に化けます"
                       f"(W2bの実バグ: &H9FFF が -24577 になり漢字姓の検知が全滅)")
            else:
                why = ("5桁以上は既に Long として読まれますが、桁数で型が変わる"
                       "書き方を残さないため接尾辞は統一で必須です")
            info.add(
                "ERROR", lineno,
                f"16進リテラル &H{digits} に Long接尾辞 '&' がありません。{why}。"
                f"&H{digits}& と書いてください: 「{stmt.strip()[:80]}」",
            )


# ==============================================================================
# 整数除算演算子 `\` の禁止(裁定書23 C-3)
# ------------------------------------------------------------------------------
# Mac版Excelで vba_src からの AddFromString を通すと、`\`(U+005C)がMacの日本語
# コードページで別バイトへ化け、コンパイルエラーになる(実機で8行・10ファイルが
# 全滅)。演算子としての `\` は 0 方向切り捨ての除算なので、`Fix(a / b)`(Long へ
# 入れるなら `CLng(Fix(a / b))`)へ書き換えれば同値。文字列リテラルの中とコメント
# の中の `\`(パス・JSONエスケープの説明)は化けても実害が無いので対象外。
# ==============================================================================
def check_integer_division_operator(info: ModuleInfo) -> None:
    r"""R-BS: 文字列リテラル・コメント外の `\` 演算子を禁止(裁定書23 C-3)。"""
    for lineno, stmt in info.statements:
        masked = _blank_string_literals(stmt)
        if "\\" not in masked:
            continue
        info.add(
            "ERROR", lineno,
            "整数除算演算子 `\\` は禁止です(裁定書23 C-3: Mac版Excelの日本語"
            "コードページで別バイトに化けてコンパイルエラーになります)。"
            "`a \\ b` は `Fix(a / b)`(Longへ入れるなら `CLng(Fix(a / b))`)へ"
            f"書き換えてください: 「{stmt.strip()[:80]}」",
        )


# 上のルールの自己テスト(骨抜き防止)。正例=findingが出てはいけない書き方、
# 負例=必ずERRORが出なければならない書き方。run_lint の冒頭で毎回走らせる。
_INTDIV_SELFTEST_OK = [
    'mins = CLng(Fix(waitSec / 60))',
    'p = "C:\\work\\out.txt"',
    "s = Replace$(s, \"\\\\\", \"/\")",
]
_INTDIV_SELFTEST_NG = [
    'mins = waitSec \\ 60',
    'n = (Len(t) + cols - 1) \\ cols',
]


def _selftest_integer_division_operator() -> list[str]:
    """正例/負例を check_integer_division_operator に通し、食い違いを文字列で返す。"""
    problems: list[str] = []
    for src in _INTDIV_SELFTEST_OK:
        probe = ModuleInfo(path=Path("selftest.bas"), relpath=Path("selftest.bas"),
                           raw_text=src, vb_name="selftest",
                           filename_stem="selftest", statements=[(1, src)])
        check_integer_division_operator(probe)
        if probe.findings:
            problems.append(f"正例が誤検知されました: {src!r}")
    for src in _INTDIV_SELFTEST_NG:
        probe = ModuleInfo(path=Path("selftest.bas"), relpath=Path("selftest.bas"),
                           raw_text=src, vb_name="selftest",
                           filename_stem="selftest", statements=[(1, src)])
        check_integer_division_operator(probe)
        if not probe.findings:
            problems.append(f"負例が検知されませんでした: {src!r}")
    return problems


# ==============================================================================
# パス連結の一元化(裁定書29 W10.1・裁定5)
# ------------------------------------------------------------------------------
# 実機事故: 2026-09-05 のMac実測で層(b)が7本落ちた。原因の1つが、製品側に残って
# いた `dirText & "\" & name` というパス連結の**区切り決め打ち**である
# (modBootData 2箇所 / modCompanyFile 1 / modCompanyFile3 1 / modExportHtml 2)。
# W9.2 で modUtil.PathSep() を用意したのに、そこを通らない連結が残っていた。
# Windowsでは同じ文字列になるので**どのゲートも赤くならず**、Macの実Excel
# (区切りが "/")でだけ企業ファイルの往復とレポート出力が壊れる。
#
# 規則: 連結の形(`& "\"` / `"\" &`)を書いてはならない。連結は
#   modUtilPath.JoinPath(dirText, tail) だけを使う(区切りは PathSep が決める)。
#
# 検出の形: **連結形だけ**を見る(裁定書29 裁定5 の逐語)。
#   ・`... & "\"` / `"\" & ...`      -> ERROR(パス連結の決め打ち)
#   ・`If Right$(t, 1) = "\" Then`   -> 対象外(文字の比較。連結していない)
#   ・`Replace$(s, "\\", "/")`       -> 対象外(2文字のリテラルは別物)
#
# 許可(理由を必ず1行書くこと):
#   ・src/test/ 配下 … 層(a)/層(b)のテストは「Windowsの区切りで組んだパスを
#     分解できるか」を確かめるために `"\"` そのものを式に持つ必要がある
#     (modTestsPure11 の JsStringSafe 検査も同じ形を持つ)。裁定書29 裁定5 が
#     名指しで src/test を除外している。
#   ・modUtil / modUtilPath … PathSep / TrimTrailingSep / JoinPathWith の
#     **実装本体**。区切りを知ってよい唯一の場所であり、ここを禁止すると
#     規則そのものが書けない。
#   ・modJsonLite … `"\" & nx` は JSONのエスケープ(RFC 8259 の逆斜線)であって
#     パス連結ではない。裁定書29 裁定5 が「JSONエスケープは対象外」と定めて
#     いるが、検出パターン(連結形)には掛かってしまうため、モジュール単位で
#     外す。**modJsonLite はパスを一切組み立てない**(JSONの読み書きだけ)ので
#     この除外でパス連結の検出漏れは生じない。
# ==============================================================================
PATH_JOIN_EXEMPT_MODULES = {"modUtil", "modUtilPath", "modJsonLite"}
PATH_JOIN_EXEMPT_LAYER_DIRS = ("test",)
# `& "\"`(前が連結演算子)と `"\" &`(後ろが連結演算子)の2形。`"\\"`(2文字)は
# 別リテラルなので、リテラルは**ちょうど1文字の逆斜線**に限定する。
PATH_JOIN_PATTERNS = (
    re.compile(r'&\s*"\\"(?!")'),
    re.compile(r'(?<!")"\\"\s*&'),
)


def _path_join_exempt(info: ModuleInfo) -> bool:
    if module_name_for_display(info) in PATH_JOIN_EXEMPT_MODULES:
        return True
    parts = info.relpath.as_posix().split("/")
    return len(parts) > 1 and parts[0] in PATH_JOIN_EXEMPT_LAYER_DIRS


def check_path_join_literal(info: ModuleInfo) -> None:
    r"""裁定書29 裁定5: パス連結の `& "\"` / `"\" &` を禁止(JoinPath へ寄せる)。"""
    if _path_join_exempt(info):
        return
    for lineno, stmt in info.statements:
        for pat in PATH_JOIN_PATTERNS:
            if pat.search(stmt):
                info.add(
                    "ERROR", lineno,
                    "裁定書29 W10.1: パス区切りの決め打ち連結は禁止です"
                    "(Mac版Excelの区切りは \"/\" なので往復が成立しません)。"
                    "`modUtilPath.JoinPath(dirText, tail)` へ寄せてください: "
                    f"「{stmt.strip()[:80]}」",
                )
                break


# ==============================================================================
# 一時フォルダの環境変数の決め打ち(裁定書29 W10.1・裁定5)
# ------------------------------------------------------------------------------
# 同じMac実測で、テストが `Environ$("TEMP")` を使っていたために一時ファイルの
# 置き場が空になり、SaveAs失敗(Err=1004)・UTF8書出失敗・企業ファイル往復の
# 前提不成立が連鎖した。Macの実Excel は TEMP を持たず TMPDIR を持つ。
# 値源は modUtilPath.TempDir()(TEMP -> TMP -> TMPDIR -> 本体と同じフォルダ)の
# 1本に寄せ、他所からの決め打ちを禁止する。
#
# 対象の語: TEMP / TMP / TMPDIR の3つ(裁定書29 の逐語は TEMP だけだが、どれを
#   決め打ちしても同じ壊れ方をするので司令塔の裁定 2026-09-05 で3語へ広げた)。
# 許可: modUtil / modUtilPath のみ(TempDir の実装本体。裁定書29 は「modUtil
#   以外で禁止」と書いたが、30,000字契約により実装は modUtilPath へ置いた)。
#   **src/test も対象**にする(今回の事故はテスト側の決め打ちが原因だった)。
# ==============================================================================
ENV_TEMP_EXEMPT_MODULES = {"modUtil", "modUtilPath"}
# 裁定書29 の逐語は `Environ$("TEMP")` だけだが、Macの実Excel が持つのは TMPDIR
# であり、TMP も含めた3語のどれを決め打ちしても同じ壊れ方をする(司令塔の裁定
# 2026-09-05 で3語へ拡張)。値源は modUtilPath.TempDir() の1本に寄せる。
ENV_TEMP_PATTERN = re.compile(
    r'\bEnviron\$?\s*\(\s*"(?:TEMP|TMP|TMPDIR)"\s*\)', re.IGNORECASE)


def check_env_temp_literal(info: ModuleInfo) -> None:
    """裁定書29 裁定5: `Environ$("TEMP")` を禁止(modUtilPath.TempDir へ寄せる)。"""
    if module_name_for_display(info) in ENV_TEMP_EXEMPT_MODULES:
        return
    for lineno, stmt in info.statements:
        if ENV_TEMP_PATTERN.search(stmt):
            info.add(
                "ERROR", lineno,
                "裁定書29 W10.1: 一時フォルダの環境変数の決め打ちは禁止です"
                "(Mac版Excelは TEMP を持たず TMPDIR を持ちます)。"
                "`modUtilPath.TempDir()` を使ってください: "
                f"「{stmt.strip()[:80]}」",
            )


# ==============================================================================
# 保存先(data_dir)の環境変数の参照(裁定書31 裁定1・司令塔裁定 2026-09-05)
# ------------------------------------------------------------------------------
# 会社PC実測で `%OneDrive%` が**別の利用者(win11admin)のフォルダ**を指し、
# `%OneDriveCommercial%` は未定義だった。環境変数から OneDrive を探すと、
# 他人のフォルダへ企業データを書きに行く。よって data_dir の解決からこれらの
# 参照を全撤去した(17章 Z-32)。純関数テストは引数しか見ないので、「常に空を
# 返す Environ$」を足し戻す退行は層(a)では捕まらない(W11-b の変異注入(a)で
# 実測)。ここで**全モジュール(modUtil / modUtilPath / src/test を含む。除外
# なし)**を対象に機械で止める。
#
# 対象の語: OneDrive / OneDriveCommercial / USERPROFILE の3つ。
#   TEMP / TMP / TMPDIR の規則(上)とその除外は別物で、そのまま据え置く。
# ==============================================================================
ENV_DATADIR_PATTERN = re.compile(
    r'\bEnviron\$?\s*\(\s*"(?:OneDriveCommercial|OneDrive|USERPROFILE)"\s*\)',
    re.IGNORECASE)


def check_env_datadir_literal(info: ModuleInfo) -> None:
    """裁定書31 裁定1: OneDrive/USERPROFILE の環境変数参照を全モジュールで禁止。"""
    for lineno, stmt in info.statements:
        if ENV_DATADIR_PATTERN.search(stmt):
            info.add(
                "ERROR", lineno,
                "裁定書31 W11-b: 保存先の環境変数参照は禁止です"
                "(会社PCの %OneDrive% は別の利用者のフォルダを指します)。"
                "data_dir の値源は data_dir.txt -> config -> 本体と同じフォルダ"
                "\\データ の3段だけです(17章 Z-32): "
                f"「{stmt.strip()[:80]}」",
            )


# 上の3ルールの自己テスト(骨抜き防止)。正例=findingが出てはいけない書き方、
# 負例=必ずERRORが出なければならない書き方。run_lint の末尾で毎回走らせる。
_PATHJOIN_SELFTEST_OK = [
    'CompanyFilePath = modUtilPath.JoinPath(dirText, baseName & CF_EXT)',
    'If Right$(t, 1) = "\\" Or Right$(t, 1) = "/" Then',
]
_PATHJOIN_SELFTEST_NG = [
    'CompanyFilePath = dirText & "\\" & baseName & CF_EXT',
    'nameText = Dir$(dirText & "\\" & BD_PATTERN)',
]
_ENVTEMP_SELFTEST_OK = [
    'd = modUtilPath.TempDir()',
]
_ENVTEMP_SELFTEST_NG = [
    'd = Environ$("TEMP")',
    'd = Environ("TMPDIR")',
]
# 裁定書31: OneDrive/USERPROFILE は除外なしで禁止。TMPDIR は modUtilPath で
#   許可のまま(上の規則の管轄であり、この規則は反応してはいけない)。
_ENVDATADIR_SELFTEST_OK = [
    'd = Environ$("TMPDIR")',
]
_ENVDATADIR_SELFTEST_NG = [
    'outText = AppendCandidate(outText, Environ$("OneDrive"))',
]


def _selftest_path_rules() -> list[str]:
    """裁定書29 の2ルールへ正例/負例を通し、食い違いを文字列で返す。"""
    problems: list[str] = []
    cases = (
        ("パス連結禁止", check_path_join_literal,
         _PATHJOIN_SELFTEST_OK, _PATHJOIN_SELFTEST_NG),
        ('Environ$("TEMP")禁止', check_env_temp_literal,
         _ENVTEMP_SELFTEST_OK, _ENVTEMP_SELFTEST_NG),
        ('Environ$("OneDrive")禁止', check_env_datadir_literal,
         _ENVDATADIR_SELFTEST_OK, _ENVDATADIR_SELFTEST_NG),
    )
    for label, fn, ok_list, ng_list in cases:
        for src in ok_list:
            probe = _probe_module([src])
            fn(probe)
            if probe.findings:
                problems.append(f"{label}: 正例が誤検知されました: {src!r}")
        for src in ng_list:
            probe = _probe_module([src])
            fn(probe)
            if not probe.findings:
                problems.append(f"{label}: 負例が検知されませんでした: {src!r}")
    return problems


# ==============================================================================
# (a) AV表面積: 配布物から消すAPIの「形」(裁定書27 W9-B7(a))
# ------------------------------------------------------------------------------
# 2026-09-02、同じ社内環境でマクロ型マルウェアの「形」が社内AVのAMSIに検知され、
# VDIが強制停止して情シスチケットになった(裁定書27 事実)。AVが重く見るのは
# **何をしたか**ではなく**どう書いてあるか**なので、機能を保ったまま形だけを
# 配布物から消す。消したものが戻ってこないよう、ここで機械的に止める。
#
# 検査対象は「コメントを除いたコード」であり、**文字列リテラルの中も見る**。
# `CreateObject("ADODB.Stream")` や `GetObject("New:{CLSID}")` のように、
# 危険な形はまさに文字列の中に書かれるためである(文字列を伏せたら空振りする)。
# 説明のためのコメントは対象外(コメントは配布binのソースにも残るが、AMSIが
# 見るのは実行される形であり、記録として残す価値のほうが大きい)。
#
# 許可リスト:
#   ・src/test/ 配下 … 実機層(b)のテストは「危険な形が消えていること」を
#     確かめるために語そのものを持つことがある。配布ビルドの ship 判定は
#     tools/ship_check.py が vbaProject.bin を直接見るので、ここを緩めても
#     配布物の検査は緩まない。
#   ・modGatewayDirect … direct経路は 裁定書27 W9-B5 で**配布ビルドから外れ
#     dev専用**になった(build/modules.json の ship:false)。dev専用モジュール
#     に配布物の基準を課すと、開発用の経路を維持できなくなる。
FORBIDDEN_API_ALLOW_MODULES = {"modGatewayDirect"}
FORBIDDEN_API_ALLOW_LAYER_DIRS = ("test",)
FORBIDDEN_API_PATTERNS = (
    (re.compile(r"\bVBComponents\b", re.IGNORECASE),
     "VBComponents(VBAプロジェクトへの書込。自己インストーラの形)"),
    (re.compile(r"\bVBProject\b", re.IGNORECASE),
     "VBProject(VBAプロジェクトへの参照。自己インストーラの形)"),
    (re.compile(r"\bAddFromString\b", re.IGNORECASE),
     "AddFromString(ソースの実行時注入)"),
    (re.compile(r"\bExecuteExcel4Macro\b", re.IGNORECASE),
     "ExecuteExcel4Macro(XLM経由の実行)"),
    (re.compile(r"WScript\s*\.\s*Shell", re.IGNORECASE),
     "WScript.Shell(外部コマンドの実行)"),
    (re.compile(r"\bADSystemInfo\b", re.IGNORECASE),
     "ADSystemInfo(ドメイン情報の収集)"),
    (re.compile(r"\bADODB\s*\.\s*Stream\b", re.IGNORECASE),
     "ADODB.Stream(ファイル書出。modUtil.WriteUtf8File へ寄せること)"),
    (re.compile(r"Scripting\s*\.\s*FileSystemObject", re.IGNORECASE),
     "Scripting.FileSystemObject(ファイル操作。modUtil.EnsureFolder 等へ寄せること)"),
    (re.compile(r"\bGetObject\s*\(\s*\"[Nn][Ee][Ww]\s*:"),
     'GetObject("New:{CLSID}")(参照設定なしのCOM生成)'),
    (re.compile(r"[Nn][Ee][Ww]\s*:\s*\{"),
     "new:{CLSID}(参照設定なしのCOM生成)"),
    (re.compile(r"\bDeclare\s+PtrSafe\b", re.IGNORECASE),
     "Declare PtrSafe(Win32 APIの宣言)"),
    (re.compile(r"\bDeclare\s+(?:Sub|Function)\b", re.IGNORECASE),
     "Declare Sub/Function(Win32 APIの宣言)"),
    (re.compile(r"\bHyperlinks\s*\.\s*Add\b", re.IGNORECASE),
     "Hyperlinks.Add(11章§8.6 禁忌1: 図形のOnActionを殺す)"),
)


def _forbidden_api_exempt(info: ModuleInfo) -> bool:
    if module_name_for_display(info) in FORBIDDEN_API_ALLOW_MODULES:
        return True
    parts = info.relpath.as_posix().split("/")
    return len(parts) > 1 and parts[0] in FORBIDDEN_API_ALLOW_LAYER_DIRS


def check_forbidden_api_tokens(info: ModuleInfo) -> None:
    """裁定書27 W9-B7(a): 社内AVが重く見るAPIの形を配布ソースから禁止する。"""
    if _forbidden_api_exempt(info):
        return
    for lineno, raw in merge_continuations(info.raw_text.split("\n")):
        code = strip_comment(raw)
        if not code.strip():
            continue
        for pat, why in FORBIDDEN_API_PATTERNS:
            if pat.search(code):
                info.add(
                    "ERROR", lineno,
                    f"裁定書27 W9-B7(a) 禁止API: {why}。"
                    f"社内AVのAMSIがマクロ型マルウェアの特徴として検知するため、"
                    f"配布ソースには書けません: 「{code.strip()[:80]}」",
                )


# ==============================================================================
# (d) Workbooks.Open の前に DisplayAlerts を退避しているか(裁定書27 W9-B7(d))
# ------------------------------------------------------------------------------
# `Workbooks.Open` は、ファイルが他者にロックされている・読取専用推奨・リンクの
# 更新確認・修復の確認といった**モーダルダイアログ**を出す。無人の起動シーケンス
# や一括実行の途中でこれが出ると、画面は固まったように見えて誰も操作できない
# (16章 E-51 の「モーダルを出さない」に反する)。開く前に `DisplayAlerts` を
# 退避して False にし、開いた後で必ず戻すこと。
# 「直前5行」に退避があるかだけを見る(構文解析はしない。5行は Dim と存在確認を
#  挟む現実の書き方に足りる幅)。
WORKBOOKS_OPEN_PATTERN = re.compile(r"\bWorkbooks\s*\.\s*Open\b", re.IGNORECASE)
DISPLAY_ALERTS_PATTERN = re.compile(r"\bDisplayAlerts\b", re.IGNORECASE)
WORKBOOKS_OPEN_LOOKBACK = 5


def check_workbooks_open_alerts(info: ModuleInfo) -> None:
    """裁定書27 W9-B7(d): Workbooks.Open の直前5行に DisplayAlerts 退避が要る。"""
    lines = [(n, strip_comment(s)) for n, s in merge_continuations(info.raw_text.split("\n"))]
    for idx, (lineno, code) in enumerate(lines):
        if not WORKBOOKS_OPEN_PATTERN.search(code):
            continue
        window = [c for _, c in lines[max(0, idx - WORKBOOKS_OPEN_LOOKBACK):idx]]
        if any(DISPLAY_ALERTS_PATTERN.search(c) for c in window):
            continue
        info.add(
            "ERROR", lineno,
            f"裁定書27 W9-B7(d): Workbooks.Open の直前{WORKBOOKS_OPEN_LOOKBACK}行に "
            f"DisplayAlerts の退避がありません。ロック・読取専用推奨・リンク更新の"
            f"モーダルが無人実行を固めます(開く前に退避して False にし、"
            f"開いた後で必ず戻すこと): 「{code.strip()[:80]}」",
        )


# ------------------------------------------------------------------------------
# 上の2ルールの自己テスト(骨抜き防止)。正例=findingが出てはいけない書き方、
# 負例=必ずERRORが出なければならない書き方。run_lint の末尾で毎回走らせる。
# ------------------------------------------------------------------------------
_FORBIDDEN_API_SELFTEST_OK = [
    '    Set st = CreateObject("MSXML2.ServerXMLHTTP.6.0")',
    "    ' ADODB.Stream は 裁定書27 W9-B2 で撤去した(この行はコメント)",
    '    ok = modUtil.WriteUtf8File(pathText, bodyText, True)',
    '    ws.Shapes(shapeKey).OnAction = "modUINav.BackToNav"',
]
_FORBIDDEN_API_SELFTEST_NG = [
    '    Set st = CreateObject("ADODB.Stream")',
    '    Set fso = CreateObject("Scripting.FileSystemObject")',
    '    Set dobj = GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")',
    "    ThisWorkbook.VBProject.VBComponents.Add(1)",
    '    md.CodeModule.AddFromString srcText',
    '    ExecuteExcel4Macro "CALL(...)"',
    '    Set sh = CreateObject("WScript.Shell")',
    '    Set inf = CreateObject("ADSystemInfo")',
    '    Private Declare PtrSafe Function Foo Lib "user32" () As Long',
    '    ws.Hyperlinks.Add anchor, "https://example.invalid"',
]
_WBOPEN_SELFTEST_OK = [
    ["    prevAlerts = Application.DisplayAlerts",
     "    Application.DisplayAlerts = False",
     "    Set wb = Application.Workbooks.Open(pathText, 0, readOnlyMode)"],
]
_WBOPEN_SELFTEST_NG = [
    ["    If Not FileExists(pathText) Then Exit Function",
     "    Set wb = Application.Workbooks.Open(pathText, 0, readOnlyMode)"],
]


def _probe_module(lines: list[str]) -> ModuleInfo:
    src = "\n".join(lines)
    return ModuleInfo(path=Path("selftest.bas"), relpath=Path("selftest.bas"),
                      raw_text=src, vb_name="selftest", filename_stem="selftest",
                      statements=list(iter_statements(src.split("\n"))))


def _selftest_forbidden_api_tokens() -> list[str]:
    problems: list[str] = []
    for src in _FORBIDDEN_API_SELFTEST_OK:
        probe = _probe_module([src])
        check_forbidden_api_tokens(probe)
        if probe.findings:
            problems.append(f"禁止API検査: 正例が誤検知されました: {src!r}")
    for src in _FORBIDDEN_API_SELFTEST_NG:
        probe = _probe_module([src])
        check_forbidden_api_tokens(probe)
        if not probe.findings:
            problems.append(f"禁止API検査: 負例が検知されませんでした: {src!r}")
    # 許可リストが効くこと(効かないと dev専用モジュールが直せなくなる)。
    exempt = _probe_module([_FORBIDDEN_API_SELFTEST_NG[0]])
    exempt.vb_name = "modGatewayDirect"
    check_forbidden_api_tokens(exempt)
    if exempt.findings:
        problems.append("禁止API検査: modGatewayDirect の許可リストが効いていません")
    return problems


def _selftest_workbooks_open_alerts() -> list[str]:
    problems: list[str] = []
    for lines in _WBOPEN_SELFTEST_OK:
        probe = _probe_module(lines)
        check_workbooks_open_alerts(probe)
        if probe.findings:
            problems.append(f"Workbooks.Open検査: 正例が誤検知されました: {lines!r}")
    for lines in _WBOPEN_SELFTEST_NG:
        probe = _probe_module(lines)
        check_workbooks_open_alerts(probe)
        if not probe.findings:
            problems.append(f"Workbooks.Open検査: 負例が検知されませんでした: {lines!r}")
    return problems


def _selftest_msvbal_reserved_names() -> list[str]:
    """(b) MS-VBAL 予約名の識別子使用(裁定書27 W9-B7(b))の正負例。"""
    problems: list[str] = []
    ok_cases = ["Dim widthPt As Double", "ByVal formatName As String",
                "Public Function TextOf(ByVal s As String) As String"]
    ng_cases = ["Dim name As String", "Optional ByVal format As String",
                "Public Function Value(ByVal s As String) As String",
                "Private Const Step As Long = 1", "Dim height As Double"]
    for src in ok_cases:
        probe = _probe_module([src])
        check_msvbal_reserved_names(probe)
        if probe.findings:
            problems.append(f"MS-VBAL予約名検査: 正例が誤検知されました: {src!r}")
    for src in ng_cases:
        probe = _probe_module([src])
        check_msvbal_reserved_names(probe)
        if not probe.findings:
            problems.append(f"MS-VBAL予約名検査: 負例が検知されませんでした: {src!r}")
    return problems


def _selftest_hex_literal_suffix() -> list[str]:
    """(c) &H8000-&HFFFF の Long接尾辞(裁定書8 A-5 / 裁定書27 W9-B7(c))の正負例。"""
    problems: list[str] = []
    ok_cases = ["cp = &H9FFF&", "cp = &H7FFF", "s = \"&H9FFF\""]
    ng_cases = ["cp = &H8000", "cp = &H9FFF", "cp = &HFFFF"]
    for src in ok_cases:
        probe = _probe_module([src])
        check_hex_literal_suffix(probe)
        if [f for f in probe.findings if f.level == "ERROR"]:
            problems.append(f"16進接尾辞検査: 正例が誤検知されました: {src!r}")
    for src in ng_cases:
        probe = _probe_module([src])
        check_hex_literal_suffix(probe)
        if not [f for f in probe.findings if f.level == "ERROR"]:
            problems.append(f"16進接尾辞検査: 負例が検知されませんでした: {src!r}")
    return problems


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
            lhs_head = lhs.strip().split(".")[0].strip()
            if (name, lhs_head) in CELL_WRITE_DOM_TARGETS:
                continue
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


# ------------------------------------------------------------------------------
# W9.2: ブック名で修飾した手続き名を組み立てない
# ------------------------------------------------------------------------------
# `Application.OnTime Procedure:="'" & ThisWorkbook.Name & "'!modX.Foo"` や
# `Application.Run "'" & ThisWorkbook.Name & "'!modBoot.Boot"` は、**ホスト側
# (Excel)が実行時に文字列でブックと手続きを引き当てる**形である。ブック名に
# 非ASCII(日本語)が入る配布物では、この解決が環境によって失敗し、利用者には
# 「実行時エラー 5: プロシージャの呼び出し、または引数が無効です」の生ダイアログ
# しか出ない(2026-09 Mac実機・W9.2)。呼び先が**自分のプロジェクト内**にある
# 限り、修飾する必要はまったく無い(VBA が自分で名前解決する)。
# したがって「ブック名 + '!' で手続き名を組み立てる」書き方そのものを禁じる。
WORKBOOK_QUALIFIED_NAME = "ThisWorkbook.Name"
ptn_bang_in_literal = re.compile(r'"[^"]*!')


def check_workbook_qualified_proc(info: ModuleInfo) -> None:
    """`ThisWorkbook.Name` と `!` を含む文字列リテラルの連結を ERROR にする。"""
    for lineno, stmt in info.statements:
        if WORKBOOK_QUALIFIED_NAME not in stmt:
            continue
        if not ptn_bang_in_literal.search(stmt):
            continue
        info.add(
            "ERROR", lineno,
            "ブック名で修飾した手続き名を組み立てています"
            "(`\"'\" & ThisWorkbook.Name & \"'!modX.Foo\"` の形)。"
            "Application.OnTime / Application.Run はこの文字列を**ホスト側が**"
            "解決するため、ブック名に日本語が入る配布物では解決に失敗し、利用者に"
            "は「実行時エラー 5」の生ダイアログしか出ません(W9.2・Mac実機)。"
            "呼び先が自プロジェクト内なら修飾は不要です。"
            f"モジュール名だけを書いてください: 「{stmt.strip()[:80]}」",
        )


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
    """検査対象のソース。

    裁定書34 §1.4(W12-A)で **.frm(UserForm)** を足した。.frm は先頭に
    VBE のデザイナ情報(VERSION / Begin...End / Attribute 行)が付いた形で
    エクスポートされる。そこはソースではないので `strip_form_header` が
    落とし、**Attribute 行より後だけ**を .bas と同じ規則で検査する。
    対になる .frx はバイナリなので触らない。
    """
    files = (sorted(src_root.rglob("*.bas"))
             + sorted(src_root.rglob("*.cls"))
             + sorted(src_root.rglob("*.frm")))
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
        check_line_cp932_bytes(info)
        check_dim_type_drop_and_integer(info)
        check_name_shadowing(info)
        check_declaration_position(info)
        check_handler_exit(info)
        check_resume_next_inside_handler(info)
        check_resume_on_fallthrough_label(info)
        check_reserved_identifiers(info)
        check_vba_reserved_words(info)
        check_msvbal_reserved_names(info)
        check_hex_literal_suffix(info)
        check_integer_division_operator(info)
        check_path_join_literal(info)
        check_env_temp_literal(info)
        check_env_datadir_literal(info)
        check_excel_tokens(info)
        check_forbidden_api_tokens(info)
        check_workbooks_open_alerts(info)
        check_application_run_whitelist(info)
        check_cell_write_guard(info)
        check_html_embed_guard(info)
        check_core_product_vocab(info)
        check_cross_module_references(info, known_modules)
        check_layer_dependency(info, known_modules)
        check_contract(info)
        check_raw_activate(info)
        check_find_lookin(info)
        check_workbook_qualified_proc(info)

    # モジュールをまたいだ検査は、全モジュールの宣言を集め終わってから1回だけ。
    check_module_level_refs(modules)
    check_undefined_proc_refs(modules)
    check_onaction_handler_guard(modules)
    check_msgbox_nonbmp(modules)
    check_array_arg_variant_mismatch(modules)
    check_qualified_arg_count(modules)
    check_udt_byval_param(modules)

    # 骨抜き防止の自己検査(裁定書19 H7): 行長バイト検査が全モジュールを
    # 実際に走ったか。呼び出しを消す/条件で握り潰すと、ここが赤で止まる。
    _bytes_unscanned = [i.relpath.as_posix() for i in modules
                        if i.relpath.as_posix() not in LINE_BYTES_SCANNED]

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

    # `\\` 禁止ルールの自己テスト(裁定書23 C-3): 正例/負例が期待どおりに
    # 判定できているか。ルールを空振りさせる改変はここが赤で止める。
    intdiv_selftest = _selftest_integer_division_operator()
    if intdiv_selftest:
        print("\n[整数除算演算子 `\\` 禁止ルールの自己テスト]")
        for msg in intdiv_selftest:
            print(f"  ERROR L1: {msg}")
        total_error += len(intdiv_selftest)

    # 裁定書27 W9-B7 の4ルール(a)(b)(c)(d)の自己テスト。正負例を毎回通し、
    # ルールを空振りさせる改変(語の削除・パターンの緩和)をここで赤にする。
    w9_selftest = (_selftest_forbidden_api_tokens()
                   + _selftest_msvbal_reserved_names()
                   + _selftest_hex_literal_suffix()
                   + _selftest_workbooks_open_alerts())
    if w9_selftest:
        print("\n[裁定書27 W9-B7 の4ルールの自己テスト]")
        for msg in w9_selftest:
            print(f"  ERROR L1: {msg}")
        total_error += len(w9_selftest)

    # W10.1(裁定書29 T-60): パス連結禁止 / Environ$("TEMP")禁止 の自己テスト。
    # 正負例が期待どおり判定できているか。ルールを空振りさせる改変(パターンの
    # 削除・許可リストの拡大)はここが赤で止める。
    path_selftest = _selftest_path_rules()
    if path_selftest:
        print("\n[W10.1: パス連結 / 一時フォルダ 禁止ルールの自己テスト]")
        for msg in path_selftest:
            print(f"  ERROR L1: {msg}")
        total_error += len(path_selftest)

    # W9.4(17章 T-58): UDT/配列 ByVal 禁止ルールの自己テスト。正負例が期待どおり
    # 判定できているか。ルールを空振りさせる改変はここが赤で止める。
    udt_byval_selftest = _selftest_udt_byval_param()
    if udt_byval_selftest:
        print("\n[W9.4: UDT/配列ByVal禁止ルールの自己テスト]")
        for msg in udt_byval_selftest:
            print(f"  ERROR L1: {msg}")
        total_error += len(udt_byval_selftest)

    if _bytes_unscanned:
        print("\n[行長バイト検査の自己検査 - 検査が走っていないモジュール]")
        for rel in _bytes_unscanned:
            print(f"  ERROR L1: [{rel}] 1物理行CP932バイト長検査"
                  f"(<= {MAX_LINE_CP932_BYTES}バイト)が実行されていません"
                  f"(run_lint の check_line_cp932_bytes 呼び出しを復活させてください)")
        total_error += len(_bytes_unscanned)

    # 仕様側(15章・docs/08)のプロンプト本文の検問(裁定書6 A-2)。
    docs_bad = check_docs_prompt_cp932(REPO_ROOT)
    if docs_bad:
        print("\n[仕様側プロンプト本文のCP932検査 - 15章§0 原則7]")
        for rel, ln, msg in docs_bad:
            print(f"  ERROR L{ln}: [{rel}] {msg}")
        total_error += len(docs_bad)

    # 製品へ焼く文字列の危険6字検査(裁定書22 m6・11章§7.1)。
    build_bad = check_build_strings(REPO_ROOT)
    if build_bad:
        print("\n[製品へ焼く文字列の危険6字検査 - 11章§7.1]")
        for rel, ln, msg in build_bad:
            print(f"  ERROR L{ln}: [{rel}] {msg}")
        total_error += len(build_bad)

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
