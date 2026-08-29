Attribute VB_Name = "modPii"
Option Explicit

' ============================================================================
' modPii - PII(個人情報らしき文字列)走査の一元化された本体(app層・T-29)
' ----------------------------------------------------------------------------
' 責務(16章 E-05・12章§2/§4・18章§1.1(3)・13章§2.8):
'   保存・外部送信・レポート出力・企業ドシエ書出の【直前】に通す検知関数群。
'   走査対象の7経路(案件入力の貼付欄 / 受信箱body / 壁打ちの発話 / 判断台帳の
'   situation・key_reason / フィードバックの customer_quote / HTMLレポート出力 /
'   企業ドシエファイルの書出・共有前)は modUICase・modUIInbox・modSparring・
'   modJudgeStore・modExportHtml・modCompanyFile がそれぞれ前段で必ず通す。
'
' 本モジュールが持たない責務(重要):
'   ブロックするか警告に留めるかの【判断】は呼び出し側にある(16章 E-05 が経路
'   ごとに違う扱いを定めているため)。ここは「検知したか」「どこか」「伏字にした
'   文面」だけを返し、MsgBox もログもシートも触らない完全な純関数群である。
'     (1)-(4) ブロック / (5) customer_quote は警告＋伏字案 / (6) HTMLは警告のみ
'     (7) 企業ドシエは「確認した」を選ばない限り書き出さない
'
' R4(12章§4): Excelトークンを一切使わない純ロジック。tools/run_lo_tests.py の
'   PURE_ALLOWLIST に登録済みで、層(a)から直接叩ける。
'
' NFR-S3: 本モジュールが返す「検知レポート」には【本文を一切含めない】。
'   含めるのは検知種別(機械値)と文字位置、および呼び出し側が渡した箇所名
'   (列名など)だけである。err_log の detail へそのまま載せられる形にしてある。
'
' 検知パターンの出典:
'   16章 E-05 の3種(@付き=メール / 電話番号 / 「様」付き人名パターン)に、
'   docs/21 §2 の modPiiGuard(gemini-demo)から契約番号・証券番号らしき英数列を
'   足した4種。E-31(匿名化ボタン)の {{POLICY_NO}} 置換とは別物で、あちらは
'   modUICase の実行前置換、こちらは「危ないものが在ることの検知」である。
'
' 誤検知の設計方針:
'   日本語のビジネス文書では「仕様」「お客様」「皆様」「同様」「様式」などが
'   高頻度で現れる。人名検知を素朴に書くとこれらが全部当たり、警告が日常化して
'   誰も読まなくなる(警告の空洞化=検知しないのと同じ)。そこで敬称の直前1字と
'   直後1字に除外表を置き、疑わしきは【検知しない】側へ倒している。
'   取りこぼす例(既知): ひらがな名(「ゆき様」)・英字名(「Smith様」)・国際電話
'   表記(「+81-90-...」)。ここを広げると誤検知が跳ね上がるため意図的に外す。
' ============================================================================

' --- 検知種別(err_log detail に載る機械値。本文ではない) ---
Private Const PII_KIND_EMAIL As String = "email"
Private Const PII_KIND_PHONE As String = "phone"
Private Const PII_KIND_PERSON As String = "person"
Private Const PII_KIND_POLICY As String = "policy_no"

' 伏字の差し替え文字列(16章 E-05(5)の逐語)。
Private Const PII_MASK As String = "{{PERSON}}"

' 内部表現の区切り。1件=「種別:開始位置:長さ」で、件と件は vbLf(modUtil.BufText)。
Private Const PII_FLD As String = ":"

' 保持する検知の上限。巨大な貼付テキストで配列が際限なく伸びるのを止める安全弁で、
' これを超えた検知は記録しない(件数にも位置にも現れない)。上限に当たるような
' 文書はそもそもブロック対象なので、正確な件数より暴走回避を優先する。
Private Const PII_MAX_SPANS As Long = 500

' 検知レポートに並べる件数の上限と、レポート全体の字数上限(err_log detail の
' 400字に合わせる。13章§2.4)。超過分は "+N件" とだけ書く。
Private Const PII_REPORT_MAX_ITEMS As Long = 20
Private Const PII_REPORT_MAX_CHARS As Long = 400

' 人名として遡る最大字数(姓名で4字。「田中太郎様」を拾える長さ)。
Private Const PII_NAME_MAX As Long = 4

' 電話番号の桁数(国内の固定・携帯・フリーダイヤルを覆う10-11桁)と
' 許容するハイフン本数。日付や案件IDを拾わないため先頭は必ず 0 に限る。
Private Const PII_PHONE_MIN_DIGITS As Long = 10
Private Const PII_PHONE_MAX_DIGITS As Long = 11
Private Const PII_PHONE_MAX_HYPHENS As Long = 3

' 契約番号・証券番号らしい英数列(英字2-6字＋数字6桁以上)。
' M-0012 / RL-09-003 / C-20260901-001 のような本製品のIDに当たらない下限。
Private Const PII_POLICY_MIN_LETTERS As Long = 2
Private Const PII_POLICY_MAX_LETTERS As Long = 6
Private Const PII_POLICY_MIN_DIGITS As Long = 6

' 敬称の直前1字がこれらなら人名ではない(「仕様」「お客様」「同様」…)。
' 敬称ごとに分けてあるのは、同じ字でも敬称によって危険度が違うため。
Private Const PII_DENY_BEFORE_SAMA As String = "仕客皆同多模異一様王神殿社御貴各奥若子姫人有何那種者長"
Private Const PII_DENY_BEFORE_SHI As String = "摂華源両同彼旧各諸"
Private Const PII_DENY_BEFORE_DONO As String = "御神宮本拝各貴社長者"
Private Const PII_DENY_BEFORE_SAN As String = "皆客兄姉父母娘奥坊沢者長"

' 敬称の直後1字がこれらなら人名ではない(「様々」「様式」「様子」「氏名」…)。
Private Const PII_DENY_AFTER_SAMA As String = "々子式態相"
Private Const PII_DENY_AFTER_SHI As String = "名族"
Private Const PII_DENY_AFTER_DONO As String = "様"
Private Const PII_DENY_AFTER_SAN As String = ""

' 名前の直前がこれらなら敬語の定型(「お客様」「ご子息様」)なので人名ではない。
Private Const PII_DENY_HONORIFIC_HEAD As String = "おご御"

' 半角/全角の数字とハイフン。全角のまま貼られた電話番号を取りこぼさない。
Private Const PII_DIGITS_HALF As String = "0123456789"
Private Const PII_DIGITS_FULL As String = "０１２３４５６７８９"
Private Const PII_HYPHENS As String = "-－"

' 英字とメールのローカル部で追加に許す記号。英大小と数字を1本の定数へ並べると
' 「連続40文字以上のBase64様文字列」に見えて出荷前検問(17章 T-46②のキー走査)が
' 誤検知するため、26字ずつに分けて持つ(判定は Is*Char の各関数が行う)。
Private Const PII_LOWER As String = "abcdefghijklmnopqrstuvwxyz"
Private Const PII_UPPER As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Private Const PII_MAIL_EXTRA As String = "._%+-"

' ============================================================================
' HasPii - 1件でも検知したか(呼び出し側の「ブロックする/警告する」の起点)。
' ============================================================================
Public Function HasPii(ByVal sText As String) As Boolean
    HasPii = (LenB(ScanSpans(sText)) > 0)
End Function

' ============================================================================
' DetectionCount - 検知件数。0=検知なし。PII_MAX_SPANS で頭打ちになる。
' ============================================================================
Public Function DetectionCount(ByVal sText As String) As Long
    Dim spansText As String
    spansText = ScanSpans(sText)
    If LenB(spansText) = 0 Then Exit Function
    DetectionCount = UBound(Split(spansText, vbLf)) + 1
End Function

' ============================================================================
' KindsOf - 検知した種別を重複なく ";" 区切りで返す(検知なしは "")。
' ----------------------------------------------------------------------------
'   利用者向けメッセージの組み立て(「メールアドレスらしき文字列があります」)に
'   使う。日本語ラベルへの変換は ui層 の責務なので、ここは機械値のまま返す。
' ============================================================================
Public Function KindsOf(ByVal sText As String) As String
    Dim spansText As String
    spansText = ScanSpans(sText)
    If LenB(spansText) = 0 Then Exit Function

    Dim rows() As String
    rows = Split(spansText, vbLf)

    Dim acc As String
    Dim i As Long
    For i = LBound(rows) To UBound(rows)
        Dim kindText As String
        kindText = FieldAt(rows(i), 1)
        If InStr(1, ";" & acc & ";", ";" & kindText & ";", vbBinaryCompare) = 0 Then
            If LenB(acc) = 0 Then
                acc = kindText
            Else
                acc = acc & ";" & kindText
            End If
        End If
    Next i
    KindsOf = acc
End Function

' ============================================================================
' ScanReport - err_log の detail へそのまま載せられる検知レポート(NFR-S3)。
' ----------------------------------------------------------------------------
'   whereNote : 検知箇所の名前(列名・欄名。例 "案件入力/input_memo")。
'   戻り値    : 検知なしは ""。検知ありは
'               "案件入力/input_memo|person@37,email@120,+3件" のような形。
'   【本文は1文字も含めない】。含めるのは箇所名・種別・文字位置だけである。
' ============================================================================
Public Function ScanReport(ByVal sText As String, ByVal whereNote As String) As String
    Dim spansText As String
    spansText = ScanSpans(sText)
    If LenB(spansText) = 0 Then Exit Function

    Dim rows() As String
    rows = Split(spansText, vbLf)

    Dim acc As String
    Dim i As Long
    Dim shown As Long
    For i = LBound(rows) To UBound(rows)
        If shown >= PII_REPORT_MAX_ITEMS Then Exit For
        If LenB(acc) > 0 Then acc = acc & ","
        acc = acc & FieldAt(rows(i), 1) & "@" & FieldAt(rows(i), 2)
        shown = shown + 1
    Next i

    Dim total As Long
    total = UBound(rows) - LBound(rows) + 1
    If total > shown Then acc = acc & ",+" & CStr(total - shown) & "件"

    Dim headText As String
    headText = Trim$(whereNote)
    If LenB(headText) > 0 Then acc = headText & "|" & acc

    ScanReport = modUtil.SafeLeft(acc, PII_REPORT_MAX_CHARS)
End Function

' ============================================================================
' MaskText - 検知箇所を {{PERSON}} へ置換した文面を返す(16章 E-05(5))。
' ----------------------------------------------------------------------------
'   フィードバックの customer_quote だけは逐語引用がFR-11の要件のためブロック
'   せず、この差し替え案を利用者へ提示して採否を選ばせる。採用されたときだけ
'   呼び出し側が保存する(採否の判断はここではしない)。
'   置換は【後ろから】行う。前から置換すると以降の文字位置がずれる。
' ============================================================================
Public Function MaskText(ByVal sText As String) As String
    MaskText = sText

    Dim spansText As String
    spansText = ScanSpans(sText)
    If LenB(spansText) = 0 Then Exit Function

    Dim rows() As String
    rows = Split(spansText, vbLf)

    Dim outText As String
    outText = sText
    Dim i As Long
    For i = UBound(rows) To LBound(rows) Step -1
        Dim posStart As Long, spanLen As Long
        posStart = ToLong(FieldAt(rows(i), 2))
        spanLen = ToLong(FieldAt(rows(i), 3))
        If posStart >= 1 And spanLen >= 1 And posStart + spanLen - 1 <= Len(outText) Then
            outText = Left$(outText, posStart - 1) & PII_MASK & _
                      Mid$(outText, posStart + spanLen)
        End If
    Next i
    MaskText = outText
End Function

' ============================================================================
' ScanSpans - 走査の本体。"種別:開始位置:長さ" を vbLf 区切りで返す(内部表現)。
' ----------------------------------------------------------------------------
'   先頭から1文字ずつ前進し、その位置から始まる検知を順に試す。当たったら
'   その分だけ位置を飛ばすので、1つの文字列が2種別で二重に数えられない。
'   人名だけは敬称を起点に【前】へ遡るため、開始位置が現在位置より手前になる
'   (敬称の直前は漢字・カタカナで、メール/電話/証券番号の文字集合と重ならない
'    ため、遡っても既出の検知と重ならない)。
' ============================================================================
Private Function ScanSpans(ByVal sText As String) As String
    Dim nLen As Long
    nLen = Len(sText)
    If nLen = 0 Then Exit Function

    Dim buf() As String
    Dim cnt As Long
    modUtil.BufInit buf, cnt

    Dim i As Long
    i = 1
    Do While i <= nLen
        Dim spanLen As Long
        Dim posStart As Long
        Dim kindText As String
        spanLen = 0
        posStart = i
        kindText = vbNullString

        spanLen = MatchEmail(sText, i)
        If spanLen > 0 Then kindText = PII_KIND_EMAIL

        If spanLen = 0 Then
            spanLen = MatchPhone(sText, i)
            If spanLen > 0 Then kindText = PII_KIND_PHONE
        End If

        If spanLen = 0 Then
            spanLen = MatchPolicy(sText, i)
            If spanLen > 0 Then kindText = PII_KIND_POLICY
        End If

        If spanLen = 0 Then
            spanLen = MatchPerson(sText, i, posStart)
            If spanLen > 0 Then kindText = PII_KIND_PERSON
        End If

        If spanLen > 0 Then
            If cnt < PII_MAX_SPANS Then
                modUtil.BufAdd buf, cnt, _
                    kindText & PII_FLD & CStr(posStart) & PII_FLD & CStr(spanLen)
            End If
            If kindText = PII_KIND_PERSON Then
                ' 敬称の直後から再開する(遡った名前部を二度読まない)。
                i = i + (posStart + spanLen - i)
            Else
                i = i + spanLen
            End If
        Else
            i = i + 1
        End If
    Loop

    ScanSpans = modUtil.BufText(buf, cnt)
End Function

' ----------------------------------------------------------------------------
' MatchEmail - 位置 i から始まるメールアドレスらしき文字列の長さ(0=不一致)。
'   ローカル部の途中から始まる誤検知を避けるため、直前が同じ文字集合なら不一致。
'   ドメインは「.を1つ以上含み、最後の.の右が英字2字以上」を必須にする
'   (「@」だけ・「a@b」だけの文字列でいちいち警告を出さないため)。
' ----------------------------------------------------------------------------
Private Function MatchEmail(ByVal sText As String, ByVal i As Long) As Long
    Dim nLen As Long
    nLen = Len(sText)
    If i > 1 Then
        If IsMailLocalChar(Mid$(sText, i - 1, 1)) Then Exit Function
    End If

    Dim p As Long
    p = i
    Do While p <= nLen
        If Not IsMailLocalChar(Mid$(sText, p, 1)) Then Exit Do
        p = p + 1
    Loop
    If p = i Then Exit Function
    If p > nLen Then Exit Function
    If Mid$(sText, p, 1) <> "@" Then Exit Function

    Dim q As Long
    q = p + 1
    Do While q <= nLen
        Dim c As String
        c = Mid$(sText, q, 1)
        If c <> "." Then
            If Not IsMailDomainChar(c) Then Exit Do
        End If
        q = q + 1
    Loop

    Dim domainText As String
    domainText = Mid$(sText, p + 1, q - p - 1)
    Do While Len(domainText) > 0
        If Right$(domainText, 1) <> "." And Right$(domainText, 1) <> "-" Then Exit Do
        domainText = Left$(domainText, Len(domainText) - 1)
    Loop
    If Not LooksLikeDomain(domainText) Then Exit Function

    MatchEmail = (p - i) + 1 + Len(domainText)
End Function

' ドメインらしさ: 先頭が英数、"." を含み、最後の "." の右が英字2字以上。
Private Function LooksLikeDomain(ByVal domainText As String) As Boolean
    If Len(domainText) < 4 Then Exit Function
    If InStr(1, domainText, ".", vbBinaryCompare) = 0 Then Exit Function
    If Not IsMailDomainChar(Left$(domainText, 1)) Then Exit Function

    Dim lastDot As Long
    lastDot = InStrRev(domainText, ".")
    Dim tld As String
    tld = Mid$(domainText, lastDot + 1)
    If Len(tld) < 2 Then Exit Function

    Dim i As Long
    For i = 1 To Len(tld)
        If Not IsAlphaChar(Mid$(tld, i, 1)) Then Exit Function
    Next i
    LooksLikeDomain = True
End Function

' ----------------------------------------------------------------------------
' MatchPhone - 位置 i から始まる電話番号らしき文字列の長さ(0=不一致)。
'   先頭は 0(市外局番・携帯・フリーダイヤル)に限る。この1条件で「20260901」
'   のような日付連番や案件ID内の数字列を全て外せる。
' ----------------------------------------------------------------------------
Private Function MatchPhone(ByVal sText As String, ByVal i As Long) As Long
    Dim nLen As Long
    nLen = Len(sText)
    If Not IsPhoneHead(sText, i) Then Exit Function

    Dim p As Long
    p = i
    Dim digitCount As Long, hyphenCount As Long
    Do While p <= nLen
        Dim c As String
        c = Mid$(sText, p, 1)
        If IsDigitChar(c) Then
            digitCount = digitCount + 1
        ElseIf InStr(1, PII_HYPHENS, c, vbBinaryCompare) > 0 Then
            hyphenCount = hyphenCount + 1
        Else
            Exit Do
        End If
        p = p + 1
    Loop

    ' 末尾のハイフンは番号の一部ではないので削る。
    Do While p > i
        If IsDigitChar(Mid$(sText, p - 1, 1)) Then Exit Do
        hyphenCount = hyphenCount - 1
        p = p - 1
    Loop

    If digitCount < PII_PHONE_MIN_DIGITS Then Exit Function
    If digitCount > PII_PHONE_MAX_DIGITS Then Exit Function
    If hyphenCount > PII_PHONE_MAX_HYPHENS Then Exit Function
    MatchPhone = p - i
End Function

' 電話番号の走査開始点か(直前が数字・ハイフンなら途中なので開始点ではない)。
Private Function IsPhoneHead(ByVal sText As String, ByVal i As Long) As Boolean
    Dim c As String
    c = Mid$(sText, i, 1)
    If c <> "0" And c <> ChrW(&HFF10) Then Exit Function
    If i > 1 Then
        Dim prevCh As String
        prevCh = Mid$(sText, i - 1, 1)
        If IsDigitChar(prevCh) Then Exit Function
        If InStr(1, PII_HYPHENS, prevCh, vbBinaryCompare) > 0 Then Exit Function
    End If
    IsPhoneHead = True
End Function

Private Function IsDigitChar(ByVal c As String) As Boolean
    If InStr(1, PII_DIGITS_HALF, c, vbBinaryCompare) > 0 Then
        IsDigitChar = True
        Exit Function
    End If
    IsDigitChar = (InStr(1, PII_DIGITS_FULL, c, vbBinaryCompare) > 0)
End Function

' ----------------------------------------------------------------------------
' MatchPolicy - 契約番号・証券番号らしき英数列の長さ(0=不一致。docs/21 §2)。
'   「英字2-6字 + 任意のハイフン1本 + 数字6桁以上」。本製品のID(M-0012 /
'   RL-09-003 / C-20260901-001)はいずれも数字が6桁未満か英字が1字なので当たらない。
' ----------------------------------------------------------------------------
Private Function MatchPolicy(ByVal sText As String, ByVal i As Long) As Long
    Dim nLen As Long
    nLen = Len(sText)
    If Not IsAlphaChar(Mid$(sText, i, 1)) Then Exit Function
    If i > 1 Then
        If IsAlnumChar(Mid$(sText, i - 1, 1)) Then Exit Function
    End If

    Dim p As Long
    p = i
    Do While p <= nLen
        If Not IsAlphaChar(Mid$(sText, p, 1)) Then Exit Do
        p = p + 1
    Loop

    Dim letterCount As Long
    letterCount = p - i
    If letterCount < PII_POLICY_MIN_LETTERS Then Exit Function
    If letterCount > PII_POLICY_MAX_LETTERS Then Exit Function
    If p > nLen Then Exit Function

    If InStr(1, PII_HYPHENS, Mid$(sText, p, 1), vbBinaryCompare) > 0 Then p = p + 1

    Dim q As Long
    q = p
    Do While q <= nLen
        If InStr(1, PII_DIGITS_HALF, Mid$(sText, q, 1), vbBinaryCompare) = 0 Then Exit Do
        q = q + 1
    Loop
    If q - p < PII_POLICY_MIN_DIGITS Then Exit Function

    MatchPolicy = q - i
End Function

Private Function IsAlnumChar(ByVal c As String) As Boolean
    If IsAlphaChar(c) Then
        IsAlnumChar = True
        Exit Function
    End If
    IsAlnumChar = (InStr(1, PII_DIGITS_HALF, c, vbBinaryCompare) > 0)
End Function

Private Function IsAlphaChar(ByVal c As String) As Boolean
    If InStr(1, PII_LOWER, c, vbBinaryCompare) > 0 Then
        IsAlphaChar = True
        Exit Function
    End If
    IsAlphaChar = (InStr(1, PII_UPPER, c, vbBinaryCompare) > 0)
End Function

' メールのローカル部に許す文字(英数＋ . _ % + -)。
Private Function IsMailLocalChar(ByVal c As String) As Boolean
    If IsAlnumChar(c) Then
        IsMailLocalChar = True
        Exit Function
    End If
    IsMailLocalChar = (InStr(1, PII_MAIL_EXTRA, c, vbBinaryCompare) > 0)
End Function

' ドメイン部のラベルに許す文字(英数とハイフン。"." は呼び出し側で別扱い)。
Private Function IsMailDomainChar(ByVal c As String) As Boolean
    If IsAlnumChar(c) Then
        IsMailDomainChar = True
        Exit Function
    End If
    IsMailDomainChar = (c = "-")
End Function

' ----------------------------------------------------------------------------
' MatchPerson - 位置 i の敬称から【前へ遡って】人名らしき並びを見る(16章 E-05)。
'   outStart には名前の開始位置を返す(戻り値は名前＋敬称の合計字数)。
'   除外は3段: 名前の直前がお/ご/御 / 敬称の直前1字が除外表 / 直後1字が除外表。
' ----------------------------------------------------------------------------
Private Function MatchPerson(ByVal sText As String, ByVal i As Long, _
                             ByRef outStart As Long) As Long
    Dim honLen As Long
    honLen = HonorificLen(sText, i)
    If honLen = 0 Then Exit Function
    If i = 1 Then Exit Function

    Dim honText As String
    honText = Mid$(sText, i, honLen)

    ' 敬称の直前1字が除外表にあれば人名ではない。
    Dim prevCh As String
    prevCh = Mid$(sText, i - 1, 1)
    If InStr(1, DenyBeforeOf(honText), prevCh, vbBinaryCompare) > 0 Then Exit Function
    If Not IsNameChar(prevCh) Then Exit Function

    ' 敬称の直後1字が除外表にあれば熟語(様々・様式・氏名…)。
    Dim nextCh As String
    nextCh = Mid$(sText, i + honLen, 1)
    If LenB(nextCh) > 0 Then
        If InStr(1, DenyAfterOf(honText), nextCh, vbBinaryCompare) > 0 Then Exit Function
    End If

    Dim j As Long
    j = i - 1
    Do While j >= 1
        If Not IsNameChar(Mid$(sText, j, 1)) Then Exit Do
        If (i - j) >= PII_NAME_MAX Then
            j = j - 1
            Exit Do
        End If
        j = j - 1
    Loop

    Dim nameStart As Long
    nameStart = j + 1
    If nameStart > i - 1 Then Exit Function

    ' 「お客様」「ご子息様」の類は名前の直前で見分ける。
    If nameStart > 1 Then
        If InStr(1, PII_DENY_HONORIFIC_HEAD, Mid$(sText, nameStart - 1, 1), vbBinaryCompare) > 0 Then Exit Function
    End If

    outStart = nameStart
    MatchPerson = (i - nameStart) + honLen
End Function

' 位置 i から始まる敬称の字数(0=敬称ではない)。
Private Function HonorificLen(ByVal sText As String, ByVal i As Long) As Long
    Dim c As String
    c = Mid$(sText, i, 1)
    If c = "様" Or c = "氏" Or c = "殿" Then
        HonorificLen = 1
        Exit Function
    End If
    If Mid$(sText, i, 2) = "さん" Then HonorificLen = 2
End Function

Private Function DenyBeforeOf(ByVal honText As String) As String
    Select Case honText
        Case "様"
            DenyBeforeOf = PII_DENY_BEFORE_SAMA
        Case "氏"
            DenyBeforeOf = PII_DENY_BEFORE_SHI
        Case "殿"
            DenyBeforeOf = PII_DENY_BEFORE_DONO
        Case Else
            DenyBeforeOf = PII_DENY_BEFORE_SAN
    End Select
End Function

Private Function DenyAfterOf(ByVal honText As String) As String
    Select Case honText
        Case "様"
            DenyAfterOf = PII_DENY_AFTER_SAMA
        Case "氏"
            DenyAfterOf = PII_DENY_AFTER_SHI
        Case "殿"
            DenyAfterOf = PII_DENY_AFTER_DONO
        Case Else
            DenyAfterOf = PII_DENY_AFTER_SAN
    End Select
End Function

' 人名を構成しうる文字か(漢字・カタカナ・長音・中黒・々)。ひらがなと英字は
' 誤検知が多すぎるため名前構成要素にしない(本モジュール冒頭の既知の取りこぼし)。
Private Function IsNameChar(ByVal c As String) As Boolean
    Dim cp As Long
    cp = CodePointOf(c)
    If cp >= &H4E00 And cp <= &H9FFF Then
        IsNameChar = True
        Exit Function
    End If
    If cp >= &H30A1 And cp <= &H30FA Then
        IsNameChar = True
        Exit Function
    End If
    If cp = &H30FC Or cp = &H30FB Or cp = &H3005 Then IsNameChar = True
End Function

' AscW は符号付き16bitを返すため、U+8000 以上が負になる。Long のコードポイントへ
' 直してから範囲判定する(漢字域 U+4E00-U+9FFF はこの境界をまたぐ)。
Private Function CodePointOf(ByVal c As String) As Long
    If LenB(c) = 0 Then Exit Function
    Dim v As Long
    v = AscW(c)
    If v < 0 Then v = v + 65536
    CodePointOf = v
End Function

' "kind:start:len" の idx 番目(1始まり)の値。
Private Function FieldAt(ByVal rowText As String, ByVal idx As Long) As String
    Dim parts() As String
    parts = Split(rowText, PII_FLD)
    Dim k As Long
    k = LBound(parts) + idx - 1
    If k < LBound(parts) Or k > UBound(parts) Then Exit Function
    FieldAt = parts(k)
End Function

Private Function ToLong(ByVal s As String) As Long
    If LenB(s) = 0 Then Exit Function
    If Not IsNumeric(s) Then Exit Function
    ToLong = CLng(s)
End Function
