Attribute VB_Name = "modTestsPureHook"
Option Explicit

' ============================================================================
' modTestsPureHook - dev専用の純層テストを呼ぶ接続点(配布版)。裁定書30 裁定1(d)。
' ----------------------------------------------------------------------------
' なぜ在るのか:
'   modGatewayDirect の純関数16本を検査する G11 群は、direct経路ごと配布物から
'   外れる(modules.json の ship:false)ため dev専用テストモジュール
'   modTestsPureDev へ移した。しかし modTestRunner.RunAllPureTests(配布物に
'   載る)が modTestsPureDev を名前で呼ぶと、VBA はプロジェクト全体を
'   コンパイルするので配布物が実機で壊れる。
'   そこで modGatewayLink と同じ「モードでソースを差し替える薄い接続モジュール」
'   をもう1本置き、ランナー側は常に modTestsPureHook.RunAll を呼ぶ形にする。
'     prod = 本ファイル(src/test/modTestsPureHook.bas)。1本も実行しない。
'     dev  = src/test/dev/modTestsPureHook.bas。modTestsPureDev.RunAll を呼ぶ。
'   どちらを使うかは build/modules.json の dev_src が唯一の値源である
'   (Application.Run や文字列ディスパッチによる切替はしない)。
'
' 期待本数(wintest/tests_expected.txt)は prod と dev_only の2行に分かれており、
'   配布ブックは prod、dev ブックは prod+dev_only を config へ焼く。本ファイルが
'   1本も Check を呼ばないことが prod 側の本数の前提である。
' ============================================================================

Public Sub RunAll()
    ' 配布版では dev専用テストが存在しないので、何も実行しない(0本)。
End Sub
