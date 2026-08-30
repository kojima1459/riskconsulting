Attribute VB_Name = "modTestsPure5"
Option Explicit

' ============================
' modTestsPure5 - 15章§11 検証ルール表の全64ケース(17章 T-22 DoD・層(a))
' ----------------------------
' 役割: 裁定書6 D-11「テスト独立執筆」。**実装を1行も読まずに**、15章§11の
'   検証ルール表と15章の各スキーマだけを根拠に modValidate の入出力を固定する。
'   入口は Public Sub RunAll()。mock応答の受入条件・障害注入・NormalizeLlmJson・
'   SanitizeFileName は姉妹モジュール modTestsPure6 が持つ(30,000字契約のため
'   2分割。分割の裁定は司令塔へ申請済み)。
'
' 素材の作り方: 壊れたJSONを書き下ろさず、**15章§8.1のmock正常応答を素材に
'   最小改変**して作る(modMockLlm.ResponseById -> modJsonLite.ExtractJsonBlock)。
'   狙ったルール1本だけを踏み抜くので、期待値が「その行が出るか」に収斂する。
'   改変は本モジュール内の純文字列ヘルパ(PutVal/PutStr/SizeArr等)で行い、mockの
'   本文を書き写さない(同じ本文を2箇所に置かない)。
'
' 判定の形: 1ケース1本。共通形 ChkFire は
'     (1) 壊した素材では当該ケースの行が出る
'     (2) **素材のままでは出ない**(同じ文脈で誤発火しない)
'   の両方を1本で見る。行の照合は15章§11の規約どおり**行頭の [ケースID]**
'   (17章 T-22「戻り値の各行が [ケースID] で始まる」)。ケースIDはテスト名の
'   先頭から取り出すので(CaseIdOf)、**ケースIDの文字列はソース中に1本につき
'   1回しか現れない**(§4-2の照合スクリプトの1ケース1本カウントを乱さない)。
'
' 統合者へ(本ファイル単体では1本も実行されない。3点セットで結線すること):
'   (1) modTestsPure4.RunAll の末尾から modTestsPure5.RunAll を呼ぶ
'       (modTestsPure -> 2 -> 3 -> 4 と同じ数珠つなぎ)。**5 -> 6 は本ファイルの
'       RunAll 末尾で結線済み**なので、統合者が足すのはこの1本だけ。
'   (2) build/modules.json へ2件追加(modTestsPure5 / modTestsPure6。ともに
'       src/test/ 配下・role=test・type=std・wave=T-22)。
'   (3) wintest/tests_expected.txt を本ファイル70本＋modTestsPure6 の29本で更新
'       (裁定書7 A-3のID実在群の書き直しで 64 -> 70本になった)。
'   run_lo_tests.py の PURE_ALLOWLIST には modTestsPure5/6 と modValidate が既にある。
'   **modKnowledge は登録しない**。裁定書7 A-1/A-2 により実在ID一覧は Check系の
'   引数(15章§3/§4/§6.1 の1行書式テキスト)になったので、K付きの群は層(a)で
'   ホワイトリストを自給する。一覧の実体(Wl* ヘルパ)は30,000字契約の都合で
'   modTestsPure6 に置き、本モジュールは modTestsPure6.Wl* を呼ぶ。
'   modKnowledge には一切触れない。
'
' テスト本数: 70本(15章§11の全64ケース＋ID実在群のfail-closed面6本)
'   G30 CheckS1 11 / G31 CheckS2 16 / G31K CheckS2(ID実在) 3 /
'   G32 CheckS3 9 / G32K CheckS3(ID実在) 6 / G33 CheckS4 6 /
'   G34 CheckPF 6 / G34K CheckPF(ID実在) 3 / G35 CheckS2C 5 / G36 CheckS3C 5
'
' 裁定書7 A-1で実在ID一覧とS2Cの審査対象s2Jsonは引数へ昇格したが、次は依然と
' して層(a)から一意に作れないため、該当テストに前提をコメントで明記してある:
'   現場メモ提供の有無 / ラウンドの別と前ラウンド件数 / dossier_tier / S3のcase_type。
'
' 設計判断(R4準拠): Excelトークン不使用。改行は vbLf 基準。乱数・時刻不使用。
' ============================

' 15章§8.1の11 mock ID(ResponseById のキーの正)。
Private Const MK_S1N As String = "MK-S1-NEW"
Private Const MK_S1R As String = "MK-S1-RNW"
Private Const MK_S2N As String = "MK-S2-NEW"
Private Const MK_S2R As String = "MK-S2-RNW"
Private Const MK_S3 As String = "MK-S3"
Private Const MK_S4 As String = "MK-S4"
Private Const MK_PF As String = "MK-PF"
Private Const MK_C2H As String = "MK-S2C-HIT"
Private Const MK_C2C As String = "MK-S2C-CLEAN"
Private Const MK_C3H As String = "MK-S3C-HIT"
Private Const MK_C3C As String = "MK-S3C-CLEAN"

' 13章§2.1 case_type の2値。
Private Const CT_NEW As String = "new"
Private Const CT_RNW As String = "renewal"

' ----------------------------
' RunAll: グループ単位で隔離実行する。1グループが実行時エラーで落ちても
'   残りのグループは走る(未実装/未注入の事実は GroupFail で1件の失敗として
'   可視化し、無かったことにしない)。
' ----------------------------
Public Sub RunAll()
    Dim i As Long
    Dim grpName As String

    For i = 1 To 10
        grpName = "G?" & i
        On Error Resume Next
        Err.Clear
        RunGroup i, grpName
        If Err.Number <> 0 Then
            GroupFail grpName
            Err.Clear
        End If
        On Error GoTo 0
    Next i

    ' 姉妹モジュール(30,000字契約による分割)を同じ隔離作法で続けて回す。
    ' 統合者が結線するのは modTestsPure4 -> 5 の1本だけで済む。
    On Error Resume Next
    Err.Clear
    modTestsPure6.RunAll
    If Err.Number <> 0 Then
        GroupFail "modTestsPure6.RunAll"
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub RunGroup(ByVal grpNo As Long, ByRef grpName As String)
    Select Case grpNo
    Case 1
        grpName = "G30 CheckS1"
        T_S1
    Case 2
        grpName = "G31 CheckS2"
        T_S2
    Case 3
        grpName = "G31K CheckS2 ID実在"
        T_S2Ids
    Case 4
        grpName = "G32 CheckS3"
        T_S3
    Case 5
        grpName = "G32K CheckS3 ID実在"
        T_S3Ids
    Case 6
        grpName = "G33 CheckS4"
        T_S4
    Case 7
        grpName = "G34 CheckPF"
        T_PF
    Case 8
        grpName = "G34K CheckPF ID実在"
        T_PFIds
    Case 9
        grpName = "G35 CheckS2C"
        T_S2C
    Case 10
        grpName = "G36 CheckS3C"
        T_S3C
    End Select
End Sub

' ============================
' 共通ヘルパ(modTestsPure3/4 と同じ作法。各 modTestsPure* が自前で持つ)
' ============================
Private Sub GroupFail(ByVal grpName As String)
    modTestRunner.Check grpName & "(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & _
        ") ※未実装/未注入の可能性"
End Sub

Private Function HeadOf(ByVal s As String) As String
    If Len(s) <= 120 Then
        HeadOf = s
    Else
        HeadOf = Left$(s, 120) & "...(全" & Len(s) & "字)"
    End If
End Function

' ch を n 個ならべる(倍々に伸ばして切る)。
Private Function RepChar(ByVal ch As String, ByVal n As Long) As String
    Dim s As String
    If n <= 0 Or Len(ch) <= 0 Then Exit Function
    s = ch
    Do While Len(s) < n
        s = s & s
    Loop
    RepChar = Left$(s, n)
End Function

' ----------------------------
' 判定の共通形
' ----------------------------
' テスト名の先頭(最初の "_" まで)が15章§11のケースID。
Private Function CaseIdOf(ByVal nm As String) As String
    Dim p As Long
    p = InStr(nm, "_")
    If p > 1 Then
        CaseIdOf = Left$(nm, p - 1)
    Else
        CaseIdOf = nm
    End If
End Function

' 検証結果の中に「行頭が [ケースID]」の行があるか(15章§11・17章 T-22)。
Private Function HasCase(ByVal outText As String, ByVal caseId As String) As Boolean
    Dim lines() As String
    Dim i As Long
    Dim mark As String

    mark = "[" & caseId & "]"
    lines = Split(outText, vbLf)
    For i = 0 To UBound(lines)
        If Left$(LTrim$(lines(i)), Len(mark)) = mark Then
            HasCase = True
            Exit Function
        End If
    Next i
End Function

' 1ケース1本の共通形: 壊した素材で発火し、素材のままでは発火しないこと。
Private Sub ChkFire(ByVal nm As String, ByVal ngOut As String, ByVal okOut As String)
    Dim cid As String
    cid = CaseIdOf(nm)
    modTestRunner.Check nm, _
        (HasCase(ngOut, cid) And Not HasCase(okOut, cid)), _
        "壊した素材の結果=[" & HeadOf(ngOut) & "] 素材のままの結果=[" & HeadOf(okOut) & "]"
End Sub

' 「この文脈では発火してはならない」を固定する形(ケースIDはテスト名から取り出す)。
Private Sub ChkNoFire(ByVal nm As String, ByVal outText As String)
    modTestRunner.Check nm, Not HasCase(outText, CaseIdOf(nm)), _
        "発火してはならない文脈で発火しています。実際=[" & HeadOf(outText) & "]"
End Sub

' 合格判定のケース(戻り値 ""=合格。14章§6 modValidate の戻り値規約)。
Private Sub ChkPass(ByVal nm As String, ByVal outText As String)
    modTestRunner.Check nm, (outText = ""), "合格(空文字)を期待。実際=[" & HeadOf(outText) & "]"
End Sub

' fail-closed の検問(裁定書7 A-2・15章§11・16章E-07)。行の中身まで見て、
'   本物の幻覚検知と「一覧未提供」の申告を分ける。**ケースIDは頭と尾に割って
'   渡す**(テスト名に書くと17章§4-2の1ケース1本計数が狂う。ChkFaultと同作法)。
Private Sub ChkFailClosed(ByVal nm As String, ByVal outText As String, _
                          ByVal idHead As String, ByVal idTail As String)
    Dim cid As String
    cid = idHead & idTail
    modTestRunner.Check nm, _
        (InStr(outText, "[" & cid & "] ID実在検査が実行できません") > 0), _
        "期待=" & cid & " の一覧未提供による不合格(同一行)。実際=[" & HeadOf(outText) & "]"
End Sub

' 当該ケースIDが出てはならない文脈の固定(ケースIDは頭と尾に割って渡す)。
Private Sub ChkNoCase(ByVal nm As String, ByVal outText As String, _
                      ByVal idHead As String, ByVal idTail As String)
    modTestRunner.Check nm, Not HasCase(outText, idHead & idTail), _
        "発火してはならない文脈で " & idHead & idTail & _
        " が出ています。実際=[" & HeadOf(outText) & "]"
End Sub

' ============================
' JSON最小改変ヘルパ(純文字列。mock正常応答から狙った1ケースだけを壊す)
'   キーは**最初に現れた1件**を対象にする。15章のスキーマ順に読むと、
'   risks/stories/slides 等の先頭要素の項目が最初に現れる。
' ============================
Private Function MockJson(ByVal mockId As String) As String
    MockJson = modJsonLite.ExtractJsonBlock(modMockLlm.ResponseById(mockId))
End Function

Private Function SkipWs(ByVal js As String, ByVal p As Long) As Long
    Dim c As String
    Do While p <= Len(js)
        c = Mid$(js, p, 1)
        If c <> " " And c <> vbTab And c <> vbCr And c <> vbLf Then Exit Do
        p = p + 1
    Loop
    SkipWs = p
End Function

' キー k の値の開始位置(見つからなければ0)。値としての同名文字列は読み飛ばす。
Private Function KeyValStart(ByVal js As String, ByVal k As String, ByVal fromPos As Long) As Long
    Dim p As Long
    Dim q As Long
    Dim token As String

    token = """" & k & """"
    Do
        p = InStr(fromPos, js, token)
        If p = 0 Then Exit Function
        q = SkipWs(js, p + Len(token))
        If Mid$(js, q, 1) = ":" Then
            KeyValStart = SkipWs(js, q + 1)
            Exit Function
        End If
        fromPos = p + 1
    Loop
End Function

' 位置 startPos から始まる値の直後の位置。文字列・配列・オブジェクト・素値に対応。
Private Function ValueEnd(ByVal js As String, ByVal startPos As Long) As Long
    Dim p As Long
    Dim c As String
    Dim depth As Long
    Dim inQuote As Boolean

    p = startPos
    c = Mid$(js, p, 1)
    If c = """" Then
        p = p + 1
        Do While p <= Len(js)
            c = Mid$(js, p, 1)
            If c = "\" Then
                p = p + 2
            ElseIf c = """" Then
                ValueEnd = p + 1
                Exit Function
            Else
                p = p + 1
            End If
        Loop
        ValueEnd = p
    ElseIf c = "[" Or c = "{" Then
        Do While p <= Len(js)
            c = Mid$(js, p, 1)
            If inQuote Then
                If c = "\" Then
                    p = p + 1
                ElseIf c = """" Then
                    inQuote = False
                End If
            Else
                If c = """" Then
                    inQuote = True
                ElseIf c = "[" Or c = "{" Then
                    depth = depth + 1
                ElseIf c = "]" Or c = "}" Then
                    depth = depth - 1
                    If depth = 0 Then
                        ValueEnd = p + 1
                        Exit Function
                    End If
                End If
            End If
            p = p + 1
        Loop
        ValueEnd = p
    Else
        Do While p <= Len(js)
            c = Mid$(js, p, 1)
            If c = "," Or c = "}" Or c = "]" Then Exit Do
            p = p + 1
        Loop
        ValueEnd = p
    End If
End Function

' キー k の生の値テキスト(引用符・括弧を含む)。
Private Function GetVal(ByVal js As String, ByVal k As String) As String
    Dim p As Long
    Dim e As Long
    p = KeyValStart(js, k, 1)
    If p = 0 Then Exit Function
    e = ValueEnd(js, p)
    GetVal = Mid$(js, p, e - p)
End Function

' 最初のキー k の値を newText(生のJSONテキスト)へ差し替える。
Private Function PutVal(ByVal js As String, ByVal k As String, ByVal newText As String) As String
    Dim p As Long
    Dim e As Long
    PutVal = js
    p = KeyValStart(js, k, 1)
    If p = 0 Then Exit Function
    e = ValueEnd(js, p)
    PutVal = Left$(js, p - 1) & newText & Mid$(js, e)
End Function

' 全てのキー k の値を差し替える。
Private Function PutAllVal(ByVal js As String, ByVal k As String, ByVal newText As String) As String
    Dim s As String
    Dim p As Long
    Dim e As Long
    Dim fromPos As Long

    s = js
    fromPos = 1
    Do
        p = KeyValStart(s, k, fromPos)
        If p = 0 Then Exit Do
        e = ValueEnd(s, p)
        s = Left$(s, p - 1) & newText & Mid$(s, e)
        fromPos = p + Len(newText)
    Loop
    PutAllVal = s
End Function

' 文字列値の差し替え(値に引用符・逆斜線を含めない前提で使う)。
Private Function PutStr(ByVal js As String, ByVal k As String, ByVal v As String) As String
    PutStr = PutVal(js, k, """" & v & """")
End Function

Private Function PutAllStr(ByVal js As String, ByVal k As String, ByVal v As String) As String
    PutAllStr = PutAllVal(js, k, """" & v & """")
End Function

' 配列テキストの最上位 idx 番目の要素(1始まり)。
Private Function ItemAt(ByVal arrText As String, ByVal idx As Long) As String
    Dim p As Long
    Dim e As Long
    Dim n As Long

    p = InStr(arrText, "[")
    If p = 0 Then Exit Function
    p = SkipWs(arrText, p + 1)
    Do While p <= Len(arrText)
        If Mid$(arrText, p, 1) = "]" Then Exit Function
        e = ValueEnd(arrText, p)
        If e <= p Then Exit Function
        n = n + 1
        If n = idx Then
            ItemAt = Mid$(arrText, p, e - p)
            Exit Function
        End If
        p = SkipWs(arrText, e)
        If Mid$(arrText, p, 1) = "," Then p = SkipWs(arrText, p + 1)
    Loop
End Function

Private Function ItemCount(ByVal arrText As String) As Long
    Dim p As Long
    Dim e As Long
    Dim n As Long

    p = InStr(arrText, "[")
    If p = 0 Then Exit Function
    p = SkipWs(arrText, p + 1)
    Do While p <= Len(arrText)
        If Mid$(arrText, p, 1) = "]" Then Exit Do
        e = ValueEnd(arrText, p)
        If e <= p Then Exit Do
        n = n + 1
        p = SkipWs(arrText, e)
        If Mid$(arrText, p, 1) = "," Then p = SkipWs(arrText, p + 1)
    Loop
    ItemCount = n
End Function

' 配列を n 要素にそろえる(多ければ先頭から切り、少なければ先頭要素で埋める)。
Private Function SizeArr(ByVal arrText As String, ByVal n As Long) As String
    Dim i As Long
    Dim k As Long
    Dim c As Long
    Dim s As String

    c = ItemCount(arrText)
    If c = 0 Or n <= 0 Then
        SizeArr = "[]"
        Exit Function
    End If
    For i = 1 To n
        k = i
        If k > c Then k = 1
        If i > 1 Then s = s & ","
        s = s & ItemAt(arrText, k)
    Next i
    SizeArr = "[" & s & "]"
End Function

' キー k の配列を n 要素にそろえた JSON を返す。
Private Function ResizeKey(ByVal js As String, ByVal k As String, ByVal n As Long) As String
    ResizeKey = PutVal(js, k, SizeArr(GetVal(js, k), n))
End Function

' キー k の配列の2番目の要素を1番目で置き換える(件数は保ったまま値を重複させる)。
Private Function DupItem(ByVal js As String, ByVal k As String) As String
    Dim arrText As String
    Dim one As String
    Dim two As String

    arrText = GetVal(js, k)
    one = ItemAt(arrText, 1)
    two = ItemAt(arrText, 2)
    If Len(one) = 0 Or Len(two) = 0 Then
        DupItem = js
        Exit Function
    End If
    DupItem = PutVal(js, k, Replace(arrText, two, one, 1, 1))
End Function

' キー k とその値をまるごと落とす(必須キー欠落の素材づくり)。
Private Function DropKey(ByVal js As String, ByVal k As String) As String
    Dim p As Long
    Dim e As Long
    Dim ks As Long
    Dim q As Long

    DropKey = js
    p = KeyValStart(js, k, 1)
    If p = 0 Then Exit Function
    ks = InStrRev(Left$(js, p), """" & k & """")
    If ks = 0 Then Exit Function
    e = ValueEnd(js, p)
    q = SkipWs(js, e)
    If Mid$(js, q, 1) = "," Then
        e = q + 1
    ElseIf ks > 1 Then
        q = ks - 1
        Do While q > 1
            If Mid$(js, q, 1) = "," Then
                ks = q
                Exit Do
            ElseIf Mid$(js, q, 1) = " " Then
                q = q - 1
            Else
                Exit Do
            End If
        Loop
    End If
    DropKey = Left$(js, ks - 1) & Mid$(js, e)
End Function

' ============================
' G30 CheckS1(15章§11 CheckS1検証ルール表。素材=MK-S1-NEW / MK-S1-RNW)
' ============================
Private Sub T_S1()
    Dim s1n As String
    Dim s1r As String
    Dim okNew As String
    Dim okRnw As String

    s1n = MockJson(MK_S1N)
    s1r = MockJson(MK_S1R)
    okNew = modValidate.CheckS1(s1n, CT_NEW)
    okRnw = modValidate.CheckS1(s1r, CT_RNW)

    ChkFire "V-S1-01_必須キーmissing_infoの欠落_15章§11", _
        modValidate.CheckS1(DropKey(s1n, "missing_info"), CT_NEW), okNew

    ChkFire "V-S1-02_locationsのtypeがenum外_15章§11", _
        modValidate.CheckS1(PutStr(s1n, "type", "ガレージ"), CT_NEW), okNew

    ChkFire "V-S1-03_更新案件でcurrent_coverageが0件_15章§11", _
        modValidate.CheckS1(PutVal(s1r, "current_coverage", "[]"), CT_RNW), okRnw

    ' 新規文脈へ更新案件の応答(現契約3件)を流す=新規なのに現契約がある。
    ChkFire "V-S1-04_新規案件でcurrent_coverageが1件以上_15章§11", _
        modValidate.CheckS1(s1r, CT_NEW), okRnw

    ChkFire "V-S1-05_missing_infoが0件_15章§11", _
        modValidate.CheckS1(PutVal(s1n, "missing_info", "[]"), CT_NEW), okNew

    ChkFire "V-S1-06_input_qualityのcoverageが14件でない_15章§11", _
        modValidate.CheckS1(ResizeKey(s1n, "coverage", 13), CT_NEW), okNew

    ' 2件目のaspectを1件目で置換=重複1件と欠落1件が同時に立つ。
    ChkFire "V-S1-07_coverageのaspectに重複と欠落_15章§11", _
        modValidate.CheckS1(DupItem(s1n, "coverage"), CT_NEW), okNew

    ' MK-S1-NEW は overall=mid(15章§8.1)。researchを空にすると警告条件を満たす。
    ChkFire "V-S1-08_overallがhigh以外でresearch_requestsが0件_15章§11", _
        modValidate.CheckS1(PutVal(s1n, "research_requests", "[]"), CT_NEW), okNew

    ChkFire "V-S1-09_prompt_textが1800字超_15章§11", _
        modValidate.CheckS1(PutStr(s1n, "prompt_text", RepChar("a", 1801)), CT_NEW), okNew

    ChkFire "V-S1-10_research_requestsが7件_15章§11", _
        modValidate.CheckS1(ResizeKey(s1n, "research_requests", 7), CT_NEW), okNew

    ' 前提: 現場メモ提供有無は引数で渡らないため、field_insightsが0件であること
    '   自体を発火条件と読む(15章§11の補足「現場メモ提供時のみ判定」を層(a)で
    '   区別する手段が現行の公開シグネチャに無い)。
    ChkFire "V-S1-11_現場メモありでfield_insightsが0件_15章§11", _
        modValidate.CheckS1(PutVal(s1n, "field_insights", "[]"), CT_NEW), okNew
End Sub

' ============================
' G31 CheckS2(素材=MK-S2-NEW / MK-S2-RNW)
' ============================
Private Sub T_S2()
    Dim s2n As String
    Dim s2r As String
    Dim okNew As String
    Dim okRnw As String

    s2n = MockJson(MK_S2N)
    s2r = MockJson(MK_S2R)
    okNew = modValidate.CheckS2(s2n, CT_NEW)
    okRnw = modValidate.CheckS2(s2r, CT_RNW)

    ChkFire "V-S2-01_risksが5件未満_15章§11", _
        modValidate.CheckS2(ResizeKey(s2n, "risks", 4), CT_NEW), okNew

    ChkFire "V-S2-02_risk_noの重複_15章§11", _
        modValidate.CheckS2(DupItem(s2n, "risks"), CT_NEW), okNew

    ChkFire "V-S2-03_categoryがenum外_15章§11", _
        modValidate.CheckS2(PutStr(s2n, "category", "quality"), CT_NEW), okNew

    ChkFire "V-S2-04_evidenceのquoteが空_15章§11", _
        modValidate.CheckS2(PutStr(s2n, "quote", ""), CT_NEW), okNew

    ChkFire "V-S2-05_preventionsが0件_15章§11", _
        modValidate.CheckS2(PutVal(s2n, "preventions", "[]"), CT_NEW), okNew

    ' frequency=low のバンドは1-2。scoreを4にしてバンド不整合を作る。
    ChkFire "V-S2-07_frequency_scoreがバンド不整合_15章§11", _
        modValidate.CheckS2(PutVal(PutStr(s2n, "frequency", "low"), "frequency_score", "4"), CT_NEW), okNew

    ' impact=small のバンドは1-2。scoreを5にしてバンド不整合を作る。
    ChkFire "V-S2-08_impact_scoreがバンド不整合_15章§11", _
        modValidate.CheckS2(PutVal(PutStr(s2n, "impact", "small"), "impact_score", "5"), CT_NEW), okNew

    ' 前提: ラウンドの別(prevS2Jsonの有無)は引数で渡らないため、statusが
    '   proposed以外であること自体を初回ラウンドの不合格と読む。
    ChkFire "V-S2-09_初回ラウンドでstatusがproposed以外_15章§11", _
        modValidate.CheckS2(PutStr(s2n, "status", "confirmed"), CT_NEW), okNew

    ' 前提: 前ラウンドの件数が引数で渡らないため層(a)では「減った」を宣言できない。
    '   本テストは「渡せない以上、第2ラウンド専用の警告を初回文脈で誤発火
    '   させないこと」を固定する(誤発火は改訂運用を止めるため実害が大きい)。
    ChkNoFire "V-S2-10_前ラウンドより減少_初回文脈では誤発火しない_15章§11", _
        modValidate.CheckS2(ResizeKey(s2n, "risks", 7), CT_NEW)

    ChkFire "V-S2-11_更新案件でgapsが0件_15章§11", _
        modValidate.CheckS2(PutVal(s2r, "gaps", "[]"), CT_RNW), okRnw

    ChkFire "V-S2-12_新規案件でgapsが1件以上_15章§11", _
        modValidate.CheckS2(s2r, CT_NEW), okRnw

    ChkFire "V-S2-13_gapsのgap_typeがenum外_15章§11", _
        modValidate.CheckS2(PutStr(s2r, "gap_type", "none"), CT_RNW), okRnw

    ChkFire "V-S2-14_inference比率が50%超_15章§11", _
        modValidate.CheckS2(PutAllStr(s2n, "source", "inference"), CT_NEW), okNew

    ChkFire "V-S2-15_transferabilityがhardのリスク0件_15章§11", _
        modValidate.CheckS2(PutAllStr(s2n, "transferability", "cover"), CT_NEW), okNew

    ' MK-S2-NEW の emerging_risks は1件(15章§8.1)。4件へ増やして上限超過。
    ChkFire "V-S2-16_emerging_risksが3件超_15章§11", _
        modValidate.CheckS2(ResizeKey(s2n, "emerging_risks", 4), CT_NEW), okNew

    ChkFire "V-S2-17_emerging_risksのhorizonがenum外_15章§11", _
        modValidate.CheckS2(PutStr(s2n, "horizon", "far_future"), CT_NEW), okNew
End Sub

' ============================
' G31K CheckS2 のID実在検査(裁定書7 A-1/A-2/A-3)。3面を1群で張る:
'   (1)幽霊ID素材=発火 (2)実ID素材(一覧あり)=不発火 (3)一覧未提供+非空ID=不合格。
'   加えて「値が "" のIDは検査対象外」を張り、(3)が「一覧が無ければ何でも
'   落とす」へ振れていないことを分ける。
' ============================
Private Sub T_S2Ids()
    Dim s2n As String
    Dim wl As String

    s2n = MockJson(MK_S2N)
    ' 実在ID一覧は層(a)で自給する(15章§3の1行書式。姉妹モジュールが持つ)。
    wl = modTestsPure6.WlMenusSummary()

    ' (1)(2) 一覧を渡した上で、幽霊IDだけが発火する。
    ChkFire "V-S2-06_related_menu_idが実在しない_15章§11", _
        modValidate.CheckS2(PutStr(s2n, "related_menu_id", "M-9999"), CT_NEW, wl), _
        modValidate.CheckS2(s2n, CT_NEW, wl)

    ' (3) menusText 未提供では非空の related_menu_id を検査できないので不合格
    '     (fail-open へ戻すとここが落ちる)。
    ChkFailClosed "G31K_menusText未提供で実在検査が不合格になる_15章§11", _
        modValidate.CheckS2(s2n, CT_NEW), "V-S2-", "06"

    ' 値が "" のIDは検査対象外。一覧未提供でも発火してはならない。
    ChkNoCase "G31K_related_menu_idが空なら一覧未提供でも検査対象外_15章§11", _
        modValidate.CheckS2(PutAllStr(s2n, "related_menu_id", ""), CT_NEW), "V-S2-", "06"
End Sub

' ============================
' G32 CheckS3(素材=MK-S3。S2側は MK-S2-RNW=risksとgapsの双方を持つ)
' ============================
Private Sub T_S3()
    Dim s3 As String
    Dim s2r As String
    Dim okRnw As String
    Dim dupText As String

    s3 = MockJson(MK_S3)
    s2r = MockJson(MK_S2R)
    okRnw = modValidate.CheckS3(s3, s2r)
    dupText = "重複するトピック"

    ChkFire "V-S3-01_storiesが3件でない_15章§11", _
        modValidate.CheckS3(ResizeKey(s3, "stories", 2), s2r), okRnw

    ChkFire "V-S3-02_proposal_kindがenum外_15章§11", _
        modValidate.CheckS3(PutStr(s3, "proposal_kind", "bundle"), s2r), okRnw

    ChkFire "V-S3-07_target_risk_noがS2に存在しない_15章§11", _
        modValidate.CheckS3(PutVal(s3, "target_risk_nos", "[999]"), s2r), okRnw

    ChkFire "V-S3-08_target_gap_noがS2に存在しない_15章§11", _
        modValidate.CheckS3(PutVal(s3, "target_gap_nos", "[999]"), s2r), okRnw

    ChkFire "V-S3-09_メニューも型もひとつも使われていない_15章§11", _
        modValidate.CheckS3(PutAllStr(PutAllVal(s3, "menu_ids", "[]"), "scheme_id", ""), s2r), okRnw

    ' story先頭のheadlineと do_not_propose 先頭の topic を同一文言にする。
    ChkFire "V-S3-10_do_not_proposeのtopicと重複_15章§11", _
        modValidate.CheckS3(PutStr(PutStr(s3, "headline", dupText), "topic", dupText), s2r), okRnw

    ChkFire "V-S3-11_story_noが1から3の連番でない_15章§11", _
        modValidate.CheckS3(PutVal(s3, "story_no", "5"), s2r), okRnw

    ChkFire "V-S3-12_unmatched_risksのrisk_noがS2に存在しない_15章§11", _
        modValidate.CheckS3(PutVal(s3, "risk_no", "999"), s2r), okRnw

    ' 前提: case_type は引数で渡らないため、更新案件であることは s2Json の
    '   gaps が非空であること(新規はgaps 0件が不合格条件)から導く。
    ChkFire "V-S3-13_更新案件でupsellもcross_sellも0本_15章§11", _
        modValidate.CheckS3(PutAllStr(s3, "proposal_kind", "scheme"), s2r), okRnw
End Sub

' ============================
' G32K CheckS3 のID実在検査(最重要検証。幽霊IDの黙殺除去は禁止)
'   4つの一覧(menus/lines/schemes/cases)を15章§4の書式で自給し、G31Kと同じ
'   3面＋空値の対象外を張る(KPI「すり抜け0件」を層(a)で支える群)。
' ============================
Private Sub T_S3Ids()
    Dim s3 As String
    Dim s2r As String
    Dim okRnw As String
    Dim noOptId As String
    Dim wm As String
    Dim wn As String
    Dim ws As String
    Dim wc As String

    s3 = MockJson(MK_S3)
    s2r = MockJson(MK_S2R)
    ' 実在ID一覧は層(a)で自給する(15章§4の1行書式。姉妹モジュールが持つ)。
    wm = modTestsPure6.WlMenus()
    wn = modTestsPure6.WlLines()
    ws = modTestsPure6.WlSchemes()
    wc = modTestsPure6.WlCases()
    okRnw = modValidate.CheckS3(s3, s2r, wm, wn, ws, wc)

    ChkFire "V-S3-03_menu_idが実在しない_15章§11", _
        modValidate.CheckS3(PutVal(s3, "menu_ids", "[""M-9999""]"), s2r, wm, wn, ws, wc), okRnw

    ChkFire "V-S3-04_line_idが実在しない_15章§11", _
        modValidate.CheckS3(PutVal(s3, "line_ids", "[""L-99""]"), s2r, wm, wn, ws, wc), okRnw

    ChkFire "V-S3-05_scheme_idが実在しない_15章§11", _
        modValidate.CheckS3(PutStr(s3, "scheme_id", "S-9999"), s2r, wm, wn, ws, wc), okRnw

    ChkFire "V-S3-06_similar_case_idが実在しない_15章§11", _
        modValidate.CheckS3(PutStr(s3, "similar_case_id", "K-9999"), s2r, wm, wn, ws, wc), okRnw

    ' 一覧を1本も渡さない呼び出し。menu_ids は mock でも非空なので不合格。
    ChkFailClosed "G32K_menusText未提供で実在検査が不合格になる_15章§11", _
        modValidate.CheckS3(s3, s2r), "V-S3-", "03"

    ' 残る3一覧も個別にfail-closed(裁定書7 A-2。MK-S3のline_idsは非空・
    ' scheme_id/similar_case_idは実IDへ差し替えて非空にする)。
    ChkFailClosed "G32K_linesText未提供で実在検査が不合格になる_15章§11", _
        modValidate.CheckS3(PutAllStr(PutAllStr(s3, "scheme_id", ""), "similar_case_id", ""), _
            s2r, wm), "V-S3-", "04"
    ChkFailClosed "G32K_schemesText未提供で実在検査が不合格になる_15章§11", _
        modValidate.CheckS3(PutStr(PutAllStr(s3, "similar_case_id", ""), "scheme_id", "S-0004"), _
            s2r, wm, wn), "V-S3-", "05"
    ChkFailClosed "G32K_casesText未提供で実在検査が不合格になる_15章§11", _
        modValidate.CheckS3(s3, s2r, wm, wn, ws), "V-S3-", "06"

    ' 空許容のID(scheme_id / similar_case_id)を "" にすれば一覧未提供でも
    ' 検査対象外(15章§11)。落ちるなら fail-closed が当たり過ぎている。
    noOptId = PutAllStr(PutAllStr(s3, "scheme_id", ""), "similar_case_id", "")
    ChkNoCase "G32K_scheme_idが空なら一覧未提供でも検査対象外_15章§11", _
        modValidate.CheckS3(noOptId, s2r), "V-S3-", "05"
End Sub

' ============================
' G33 CheckS4(素材=MK-S4。slides 5枚・hearing_questions 8問)
' ============================
Private Sub T_S4()
    Dim s4 As String
    Dim okOut As String

    s4 = MockJson(MK_S4)
    okOut = modValidate.CheckS4(s4)

    ' 前提: dossier_tier は引数で渡らない。t1_quick(5枚固定)の文脈で4枚。
    ChkFire "V-S4-01_クイック案件でslidesが5枚でない_15章§11", _
        modValidate.CheckS4(ResizeKey(s4, "slides", 4)), okOut

    ' 前提: t2_full(5枚から既定10枚)の文脈で11枚=上限超過。
    ChkFire "V-S4-02_通常案件でslidesが上限超_15章§11", _
        modValidate.CheckS4(ResizeKey(s4, "slides", 11)), okOut

    ChkFire "V-S4-03_slide_noが連番でない_15章§11", _
        modValidate.CheckS4(PutVal(s4, "slide_no", "3")), okOut

    ChkFire "V-S4-04_bulletsが1点未満_15章§11", _
        modValidate.CheckS4(PutVal(s4, "bullets", "[]")), okOut

    ChkFire "V-S4-05_hearing_questionsが1問未満_15章§11", _
        modValidate.CheckS4(PutVal(s4, "hearing_questions", "[]")), okOut

    ChkFire "V-S4-06_file_titleが空_15章§11", _
        modValidate.CheckS4(PutStr(s4, "file_title", "")), okOut
End Sub

' ============================
' G34 CheckPF(素材=MK-PF)
' ============================
Private Sub T_PF()
    Dim pf As String
    Dim okOut As String

    pf = MockJson(MK_PF)
    okOut = modValidate.CheckPF(pf)

    ChkFire "V-PF-01_principle_checksが5件でない_15章§11", _
        modValidate.CheckPF(ResizeKey(pf, "principle_checks", 4)), okOut

    ChkFire "V-PF-02_grammar_checksが4件でない_15章§11", _
        modValidate.CheckPF(ResizeKey(pf, "grammar_checks", 3)), okOut

    ChkFire "V-PF-04_rework_suggestionsが3件超_15章§11", _
        modValidate.CheckPF(ResizeKey(pf, "rework_suggestions", 4)), okOut

    ChkFire "V-PF-05_predicted_drop_typesがenum外_15章§11", _
        modValidate.CheckPF(PutVal(pf, "predicted_drop_types", "[""T99""]")), okOut

    ChkFire "V-PF-06_pattern_idが不正_15章§11", _
        modValidate.CheckPF(PutStr(pf, "pattern_id", "P99")), okOut

    ChkFire "V-PF-07_survivalがenum外_15章§11", _
        modValidate.CheckPF(PutStr(pf, "survival", "medium")), okOut
End Sub

' ============================
' G34K CheckPF のID実在検査(M- / S- / K- で始まるのに一覧に無い)
'   refIdsText は menusSummary+patternsText+rulesText+researchingText の連結
'   (14章§6)。G31K/G32Kと同じ3面＋「ID形式でない ref_id は対象外」を張る。
' ============================
Private Sub T_PFIds()
    Dim pf As String
    Dim wl As String

    pf = MockJson(MK_PF)
    ' 実在ID一覧は層(a)で自給する(15章§6.1の1行書式。姉妹モジュールが持つ)。
    wl = modTestsPure6.WlPfRefIds()

    ChkFire "V-PF-03_duplicatesのref_idが実在しない_15章§11", _
        modValidate.CheckPF(PutStr(pf, "ref_id", "M-9999"), wl), _
        modValidate.CheckPF(pf, wl)

    ChkFailClosed "G34K_refIdsText未提供で実在検査が不合格になる_15章§11", _
        modValidate.CheckPF(pf), "V-PF-", "03"

    ' V-PF-03 の対象は「M- / S- / K- で始まるID形式」の ref_id だけ。
    ' 形式でない値(既存提案の言い回し)は一覧が無くても検査対象外。
    ChkNoCase "G34K_ID形式でないref_idは一覧未提供でも検査対象外_15章§11", _
        modValidate.CheckPF(PutStr(pf, "ref_id", "同種の既存提案あり")), "V-PF-", "03"
End Sub

' ============================
' G35 CheckS2C(素材=MK-S2C-HIT / MK-S2C-CLEAN)
' ============================
Private Sub T_S2C()
    Dim hit As String
    Dim clean As String
    Dim s2r As String
    Dim okOut As String

    hit = MockJson(MK_C2H)
    clean = MockJson(MK_C2C)
    ' 審査対象のS2(V-S2C-03の番号実在判定。裁定書7 A-1で引数へ昇格)。
    ' risks 8件とgaps 3件の双方を持つ MK-S2-RNW を審査対象に置く。
    s2r = MockJson(MK_S2R)
    okOut = modValidate.CheckS2C(hit, s2r)
    ' 審査対象S2未提供のfail-closed(裁定書7 A-2適用拡張・15章§11「03は
    ' 審査対象S2の未提供時も不合格」)。
    ChkFailClosed "G35_審査対象S2未提供で不合格になる_15章§11", _
        modValidate.CheckS2C(hit), "V-S2C-", "03"

    ChkFire "V-S2C-01_issue_typeがenum外_15章§11", _
        modValidate.CheckS2C(PutStr(hit, "issue_type", "typo"), s2r), okOut

    ChkFire "V-S2C-02_targetの書式が不正_15章§11", _
        modValidate.CheckS2C(PutStr(hit, "target", "risk 3"), s2r), okOut

    ' 書式は正しく番号だけが審査対象のS2に無い値(999)を素材とする。
    ChkFire "V-S2C-03_targetが指す番号がS2に存在しない_15章§11", _
        modValidate.CheckS2C(PutStr(hit, "target", "risk_no:999"), s2r), okOut

    ChkFire "V-S2C-04_verdict_summaryが空_15章§11", _
        modValidate.CheckS2C(PutStr(hit, "verdict_summary", ""), s2r), okOut

    ChkPass "V-S2C-05_指摘0件は合格で改訂スキップ_15章§11", modValidate.CheckS2C(clean, s2r)
End Sub

' ============================
' G36 CheckS3C(素材=MK-S3C-HIT / MK-S3C-CLEAN)
' ============================
Private Sub T_S3C()
    Dim hit As String
    Dim clean As String
    Dim okOut As String

    hit = MockJson(MK_C3H)
    clean = MockJson(MK_C3C)
    okOut = modValidate.CheckS3C(hit)

    ChkFire "V-S3C-01_executive_reactionsが3件でない_15章§11", _
        modValidate.CheckS3C(ResizeKey(hit, "executive_reactions", 2)), okOut

    ChkFire "V-S3C-02_story_noが1から3各1回でない_15章§11", _
        modValidate.CheckS3C(PutVal(hit, "story_no", "5")), okOut

    ChkFire "V-S3C-03_issue_typeがenum外_15章§11", _
        modValidate.CheckS3C(PutStr(hit, "issue_type", "typo")), okOut

    ChkFire "V-S3C-04_targetの書式が不正_15章§11", _
        modValidate.CheckS3C(PutStr(hit, "target", "story 1")), okOut

    ChkPass "V-S3C-05_lands全trueかつ指摘0件は合格_15章§11", modValidate.CheckS3C(clean)
End Sub

