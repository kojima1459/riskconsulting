Attribute VB_Name = "modUtilPath"
Option Explicit

' ============================================================================
' modUtilPath - パスの連結・分解と実行環境の判定(裁定書29 W10.1)
' ----------------------------------------------------------------------------
' 役割:
'   **パスの区切り文字を知っているのはここだけ**という状態を作る。W9.2 で
'   modUtil.PathSep() を用意したのに、製品側には `dirText & "\" & name` の
'   決め打ちが6箇所残っていた(modBootData 2 / modCompanyFile 1 /
'   modCompanyFile3 1 / modExportHtml 2)。Windowsでは無害だが、Macの実Excel は
'   区切りが "/" なので企業ファイルの往復が成立しない(裁定書29 原因(b))。
'   連結の式をここへ寄せ、tools/vba_lint.py の禁止規則で戻れないようにする。
'
' なぜ modUtil ではないのか: modUtil は 30,000字契約(12章§2)に対して残りが
'   1,000字ほどしかなく、本モジュールの4関数を足せない。同じ core 層の姉妹
'   モジュールとして分ける(文字列の無害化・整形は modUtilText、パスは
'   modUtilPath、それ以外の道具は modUtil、という切り分け)。
'
' 純と外皮の分け方(17章§1):
'   純   = JoinPathWith / FileNameOf … 引数だけで答えが決まる。層(a)で検査する
'   外皮 = JoinPath(PathSep を読む)/ TempDir(環境変数) … 環境依存なので
'          層(a)では検査できない。中身は上の純関数に委ねる
'
' R4準拠: Excelトークン(Worksheets/Range/Application/ThisWorkbook/MsgBox/
'   ActiveSheet)には**1つも触れない**。本体と同じフォルダの値も
'   modUtil.BookDirHint()(modBoot が起動手順(2)で預けたもの)から借りる。
'   Mac版Excelかどうかの判定(`Application.OperatingSystem`)は **core には
'   置かず、テスト層の modTestsExcel3.IsMacExcel が持つ**(司令塔の裁定
'   2026-09-05。層(b)のSKIP判定にしか使わないので core を広げない)。
'
' 12章§4準拠: 製品固有の語彙(製品名・シート名・ドメインenum)を持たない。
' ============================================================================

' ============================================================================
' JoinPathWith - フォルダとファイル名を区切り sep で連結する(純関数)。
' ----------------------------------------------------------------------------
'   規則(裁定書29 裁定1):
'     ・dirText の末尾の区切り("\" と "/")は落としてから連結する(重ねない)
'     ・tail の先頭の区切りも落とす(重ねない)
'     ・dirText が空(空白のみを含む)なら "" を返す。**空フォルダに対して
'       "\name" のような根っこ直下のパスを組み立てない**(裁定書29 原因(a)で
'       実際に起きた壊れ方がこれである)
'     ・sep が空なら "" を返す(区切りを決められないなら組み立てない。
'       modUtil.SplitPathParts と同じ fail-closed の作法)
'     ・tail が空(または区切りだけ)なら、末尾を落とした dirText を返す
' ============================================================================
Public Function JoinPathWith(ByVal dirText As String, ByVal tail As String, _
                             ByVal sep As String) As String
    If LenB(sep) = 0 Then Exit Function

    Dim d As String
    d = modUtil.TrimTrailingSep(dirText)
    If LenB(d) = 0 Then Exit Function

    Dim t As String
    t = tail
    Do While Len(t) > 0
        If Left$(t, 1) = "\" Or Left$(t, 1) = "/" Then
            t = Mid$(t, 2)
        Else
            Exit Do
        End If
    Loop

    If LenB(t) = 0 Then
        JoinPathWith = d
        Exit Function
    End If
    JoinPathWith = d & sep & t
End Function

' このプラットフォームの区切りで連結する(製品コードが使う唯一の口)。
'   区切りの決め方の正は modUtil.PathSep()(CurDir$ の先頭で判定)。
Public Function JoinPath(ByVal dirText As String, ByVal tail As String) As String
    JoinPath = JoinPathWith(dirText, tail, modUtil.PathSep())
End Function

' ============================================================================
' TempDir - 一時フォルダ(末尾の区切りは落とす)。
' ----------------------------------------------------------------------------
'   順に見る(裁定書29 裁定1): TEMP -> TMP -> TMPDIR -> 本体と同じフォルダ。
'   Windowsは TEMP/TMP を持ち、Macの実Excel は **TMPDIR しか持たない**
'   (裁定書29 原因(a): テストが TEMP 決め打ちで、Mac では空になっていた)。
'   最後の逃げ場は本体と同じフォルダだが、core層は ThisWorkbook を参照できない
'   規約(12章§4)のため、modBoot が起動手順(2)で預けた値を
'   modUtil.BookDirHint() から借りる。どれも取れなければ ""。
' ============================================================================
Public Function TempDir() As String
    On Error GoTo NoEnv

    Dim t As String
    t = Environ$("TEMP")
    If LenB(t) = 0 Then t = Environ$("TMP")
    If LenB(t) = 0 Then t = Environ$("TMPDIR")
    If LenB(t) = 0 Then t = modUtil.BookDirHint()
    TempDir = modUtil.TrimTrailingSep(t)
    Exit Function
NoEnv:
    TempDir = modUtil.TrimTrailingSep(modUtil.BookDirHint())
End Function

' ============================================================================
' FileNameOf - パスの末尾(ファイル名)を返す(純関数)。
' ----------------------------------------------------------------------------
'   最後の "\" または "/" より後ろ。**どちらの区切りも見る**ので、Windowsで
'   作ったパスもMacで作ったパスも同じ関数で分解できる。区切りが1つも無ければ
'   全体をそのまま返す(相対名として扱う)。
' ============================================================================
Public Function FileNameOf(ByVal pathText As String) As String
    Dim p As Long
    p = InStrRev(pathText, "\")

    Dim q As Long
    q = InStrRev(pathText, "/")
    If q > p Then p = q

    If p <= 0 Then
        FileNameOf = pathText
    Else
        FileNameOf = Mid$(pathText, p + 1)
    End If
End Function

' ============================================================================
' BuildVersionedFileName - W12-c(裁定書38 §1 班C 3)の版つきファイル名(純関数)。
' ----------------------------------------------------------------------------
'   テンプレート: `<headWord>_<Sanitize(company)>_<yyyymmdd>_v<app_version>`
'   (拡張子は付けない。呼び出し側が付ける)
'     headWord    : 用途を表す先頭語(例 "提案書" / "レポート")
'     company     : 案件入力の自由記述。**生では使わず必ず SanitizeFileName**
'     dateCompact : modUtilText.IsoDateCompact(Date) の8桁
'     appVersion  : config app_version。空なら "v" 以降を付けない
'     dirPath/ext : 13章§2.8 手順5(最終パス240字超)の判定にだけ使う
'   長すぎるときは company 部を Fnv1a64Hex(16桁)へ**置換**して再構成する
'   (途中で切ると別の会社が同名になり得るため。13章§2.8 手順5と同じ考え方)。
' ============================================================================
Public Function BuildVersionedFileName(ByVal headWord As String, ByVal company As String, _
                                       ByVal dateCompact As String, ByVal appVersion As String, _
                                       ByVal dirPath As String, ByVal ext As String) As String
    Dim head As String
    head = modUtilText.SanitizeFileName(company)
    If LenB(head) = 0 Then head = "no_name"

    Dim nameText As String
    nameText = VerJoin(headWord, head, dateCompact, appVersion)

    If Len(dirPath) + 1 + Len(nameText) + Len(ext) > 240 Then
        nameText = VerJoin(headWord, modUtilText.Fnv1a64Hex(company), dateCompact, appVersion)
    End If
    ' 全体へは SanitizeFileName を掛け直さない(同関数は32字で切り詰めるため、
    ' 掛けると日付と版が落ちる)。無害化が要るのは company 部だけで、それは
    ' 上で済ませてある。headWord / dateCompact / appVersion は当方の値。
    BuildVersionedFileName = nameText
End Function

' 4つの部品を "_" でつなぐ(空の部品は飛ばす)。版は "v" を前置する。
Private Function VerJoin(ByVal headWord As String, ByVal nameText As String, _
                         ByVal dateCompact As String, ByVal appVersion As String) As String
    Dim t As String
    t = headWord
    If LenB(t) > 0 Then
        t = t & "_" & nameText
    Else
        t = nameText
    End If
    If LenB(dateCompact) > 0 Then t = t & "_" & dateCompact
    If LenB(appVersion) > 0 Then t = t & "_v" & appVersion
    VerJoin = t
End Function
