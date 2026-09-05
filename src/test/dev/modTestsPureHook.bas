Attribute VB_Name = "modTestsPureHook"
Option Explicit

' ============================================================================
' modTestsPureHook(dev版) - dev専用の純層テストを呼ぶ接続点。裁定書30 裁定1(d)。
' ----------------------------------------------------------------------------
' 開発ビルド(--dev)だけがこのソースを使う(build/modules.json の dev_src)。
' 配布ビルド(--prod)は src/test/modTestsPureHook.bas(何も実行しない版)を使う。
' グループ単位の失敗隔離は modTestsPure* と同じ作法で、ここで1段だけ張る
' (dev専用テストが落ちても、ランナー全体は結果を返しきる)。
' ============================================================================

Public Sub RunAll()
    On Error GoTo Failed
    modTestsPureDev.RunAll
    Exit Sub
Failed:
    modTestRunner.Check "GDev modTestsPureDev(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & _
        ") ※未実装/未注入の可能性"
End Sub
