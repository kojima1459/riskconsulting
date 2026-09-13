Attribute VB_Name = "modExportProposal"
Option Explicit

' ==========================================================
' modExportProposal - 顧客向け提案書(Wide 22枚)の生成(20章§1・裁定書38 班C)
' ------------------------------------------------
' 責務(20章§1の表): 案件から S2 と S5 のJSONを解決し、匿名化復元・PII走査・
'   DATA組立・テンプレ組立・ファイル書出を行う。**HTML本文・CSS・JSの文字列を
'   ここに書かない**(テンプレ本体は modProposalHtml1..4)。
'
' 顧客向けであるがゆえの4つの掟(裁定書38 §1 班C 2):
'   (1) **reviewedBy が空なら生成しない**。担当者が内容を確認していない資料を
'       お客さまへ出す経路を**構造的に作らない**(社内IT・AI環境 v1.1 §7.3)。
'   (2) 内部の値を DATA に**入れない**: dossier_tier / quality_mode / case_type /
'       round_no / s4_variant / status / report_path、S3 の talk_script(話法と
'       taboo=SEC-18 相当)、S1 の field_insights(現場メモの原文)。
'       「入れてから隠す」のではなく、DATAに置く枝を持たない。
'   (3) S2 由来の文字列は、DATAへ入れる前に**対訳表で機械置換する**
'       (modValidate4.SoftenTaboo)。S2 は営業向けの語彙で書かれており、
'       LLMの言い換えを通らずに顧客の目に触れる唯一の経路であるため。
'   (4) 免責は**1本だけ**(modProposalHtml4.DisclaimerText)。AI生成の表示は
'       入れない(確認を経た顧客提示物であり、(1)により未確認では出力できない)。
'
' ファイル名(W12-c): `提案書_<会社名>_<yyyymmdd>_v<app_version>.html`。
'   会社名は modUtilText.SanitizeFileName を通し、組立は
'   modUtilPath.BuildVersionedFileName が唯一持つ(HTMLレポートと同じ1本)。
'
' R4(12章§2): シート・ブックに触れない。案件データは modCaseStore / modCaseRead
'   経由、config は modConfig 経由。CP932準拠(15章§0 原則7)。
' ==========================================================

Private Const EP_SRC As String = "modExportProposal"
Private Const EP_EXT As String = ".html"
Private Const EP_HEAD As String = "提案書"
Private Const EP_SEP As String = vbTab
Private Const EP_PH_COMPANY As String = "{{COMPANY}}"
Private Const EP_VER_DEFAULT As String = "2.0.0"
Private Const EP_CODE_FAIL As String = "E0502"
Private Const EP_SERIAL_MAX As Long = 9999
' 裁定書38 §1 班C: 確認していない資料は出さない。この文言が唯一の値源。
Private Const EP_NEED_REVIEW As String = "内容を確認してから出力してください。"
' 19章§3 の10分類を**顧客語**に置き換えたラベル(20章§4の5枚目)。
'   施設・自然災害・BCP の "BCP" は略語なので本文では日本語で書く(対訳表)。
Private Const EP_CAT_KEYS As String = "strategy_market;supply_chain;manufacturing_quality;" & _
    "sales_customer;facility_bcp;hr_labor;digital_info;legal_regulatory;" & _
    "finance_counterparty;brand_social"
Private Const EP_CAT_LABELS As String = "戦略・市場;調達・供給;製造・品質;販売・お客さま;" & _
    "施設・自然災害・事業継続;人材・労務;デジタル・情報;法務・規制;財務・取引先;" & _
    "ブランド・社会"
Private Const EP_SEMI As String = ";"
' 15章§5.6 Schema-S5 の最外 required(16キー)。`s5_edited` を人が直した場合は
'   スキーマ検査を通らないため、**出力の直前にここで必ず数える**
'   (裁定書39 R2-12。欠けたまま描くと22枚が黙って保留文に化ける)。
Private Const EP_S5_REQUIRED As String = "title|subtitle|themes|premise|structure|" & _
    "business|categories_note|headline|hard_risks|ideas|four|theme_table|steps|" & _
    "decisions|share_items|notes"

' ==========================================================
' GenerateProposalHtml - 14章§6の契約。""=成功 / 非空=失敗理由。
'   outPath には成功時の確定パスを返す(失敗時は空のまま)。
'   reviewedBy が空なら**生成せず** EP_NEED_REVIEW を返す(裁定書38 班C)。
' ==========================================================
Public Function GenerateProposalHtml(ByVal caseId As String, ByRef outPath As String, _
                                     ByVal reviewedBy As String) As String
    On Error GoTo Failed
    outPath = vbNullString

    ' 確認者名は**入口で1回だけ**整える(レポート側 modExportHtml.ReviewerOf と
    ' 同じ順序: 無害化 -> 区切りを落とす -> 前後の空白を落とす)。ここで落として
    ' おくと、表紙(coverFields)にもDATAの meta.reviewed_by にも TAB/LF が
    ' 入らない(裁定書40 Q-M1)。
    Dim reviewer As String
    reviewer = Trim$(StripFieldSeps(modUtilText.SanitizeInput(reviewedBy)))
    If LenB(NeedsReviewMessage(reviewer)) > 0 Then
        modLog.LogUsage "proposal_skipped", caseId, "not_reviewed"
        GenerateProposalHtml = NeedsReviewMessage(reviewer)
        Exit Function
    End If

    If Not modCaseStore.IsValidCaseId(caseId) Then
        modLog.LogError "E0101", EP_SRC & ".GenerateProposalHtml", "invalid_case_id"
        GenerateProposalHtml = "案件IDが不正です。"
        Exit Function
    End If

    Dim ctx As TCaseCtx
    Dim roundNo As Long
    Dim qualityMode As String, s4Variant As String, tierText As String
    If Not modCaseRead.ReadCaseCtx(caseId, ctx, roundNo, qualityMode, s4Variant, tierText) Then
        GenerateProposalHtml = "案件一覧からこの案件を読めませんでした。"
        Exit Function
    End If

    Dim s2Text As String, s5Text As String
    s2Text = modCaseStore.ResolveStepJson(caseId, 2)
    s5Text = ResolveProposalJson(caseId)
    ' 空判定は Trim$ ではなく modUtilText.HasVisibleText(裁定書39 R2-04 と
    ' 同型。Trim$ は Chr(32) しか落とさないので、TAB や全角空白だけの
    ' s5_edited を「作成済み」と誤認して E0502 の分かりにくい失敗になる)。
    If Not modUtilText.HasVisibleText(s5Text) Then
        modLog.LogUsage "proposal_skipped", caseId, "s5_empty"
        GenerateProposalHtml = "先にお客さま向け提案書の作成を実行してください。"
        Exit Function
    End If

    ' 必須キーの検査(裁定書39 R2-12)。欠けたまま描くと描画関数が例外を投げ、
    ' その枚が20章§4.1 の1行に化けて「わざと保留した項目」に見える。
    ' **黙って化けさせない**: ここで止めて E0502 を立てる。
    Dim missKeys As String
    missKeys = MissingProposalKeys(s5Text)
    If LenB(missKeys) > 0 Then
        modLog.LogError EP_CODE_FAIL, EP_SRC & ".GenerateProposalHtml", "s5_missing:" & missKeys
        GenerateProposalHtml = "提案書の内容がそろっていません(" & missKeys & _
            ")。お客さま向け提案書の作成をやり直してください。"
        Exit Function
    End If

    ' 匿名化の復元(16章 E-31)。**エスケープより前**に行う。
    s2Text = Replace(s2Text, EP_PH_COMPANY, ctx.company)
    s5Text = Replace(s5Text, EP_PH_COMPANY, ctx.company)

    ' PII走査(16章 E-05(6))。顧客向け資料には警告帯を出さない(お客さまに
    ' 当方の内部警告を見せない)ので、**記録だけ**を残して生成は続ける。
    PiiNote caseId, s2Text, s5Text

    Dim metaJson As String
    metaJson = BuildProposalMetaJson(ctx.company, s5Text, modUtil.NowStamp(), _
                   modConfig.GetStr("app_version", EP_VER_DEFAULT), reviewer, modUtil.NowStamp())

    Dim docText As String
    Dim softened As Long
    Dim tabooLeft As String
    Dim strictLeft As String
    docText = BuildProposalHtmlEx(metaJson, s2Text, s5Text, softened, tabooLeft, strictLeft)
    If LenB(docText) = 0 Then
        modLog.LogError EP_CODE_FAIL, EP_SRC & ".GenerateProposalHtml", "build_failed"
        GenerateProposalHtml = "提案書の組み立てに失敗しました。"
        Exit Function
    End If

    Dim dirText As String
    dirText = modUtil.ResolveDataDir(OutDirRaw())
    If LenB(dirText) = 0 Then
        modLog.LogError EP_CODE_FAIL, EP_SRC & ".GenerateProposalHtml", "no_out_dir"
        GenerateProposalHtml = "出力先フォルダを用意できませんでした。"
        Exit Function
    End If

    Dim pathText As String
    pathText = UniqueOutPath(dirText, ProposalFileBase(ctx.company, dirText))
    If LenB(pathText) = 0 Then
        modLog.LogError EP_CODE_FAIL, EP_SRC & ".GenerateProposalHtml", "no_free_filename"
        GenerateProposalHtml = "出力ファイル名の空きを見つけられませんでした。"
        Exit Function
    End If
    If Not modUtil.WriteUtf8File(pathText, docText, True) Then
        modLog.LogError EP_CODE_FAIL, EP_SRC & ".GenerateProposalHtml", "write_failed"
        GenerateProposalHtml = "ファイルの書き出しに失敗しました。"
        Exit Function
    End If

    modLog.LogUsage "proposal_html", caseId, "chars=" & CStr(Len(docText))
    ' 対訳表の機械置換を**必ず1行残す**(裁定書39 R2-11。docs/29 §5.3
    ' 「黙って直さない」。0件のときも記録して「1件も直していない」を示す)。
    modLog.LogUsage "proposal_taboo_softened", caseId, "n=" & CStr(softened)
    ' 置換しても残った社内語は**警告**として残す(裁定書40 S-M1 / 43 §1-4。
    ' mode=warn の対と終端でない文脈の語を置換しない代償。警告なので生成は
    ' 止めず、戻り値(=呼出側が失敗として扱う経路)にも載せない)。
    TabooLeftNote caseId, tabooLeft, strictLeft
    outPath = pathText
    Exit Function

Failed:
    modLog.LogError EP_CODE_FAIL, EP_SRC & ".GenerateProposalHtml", "unexpected", Err.Number
    GenerateProposalHtml = "提案書の生成中にエラーが発生しました。"
End Function

' ==========================================================
' NeedsReviewMessage - 顧客向けは確認必須(裁定書38 §1 班C・20章§8-3)の
'   **唯一の判断点**(純関数)。確認者名が空なら案内文、非空なら ""。
'   画面(区画④)でも先に弾くが、判断の正はここであり、画面を書き換えても
'   抜けられないようにするための1関数である。層(a)から直接叩ける。
'   空判定は Trim$ ではなく **modUtilText.HasVisibleText を直接呼ぶ**
'   (裁定書39 R2-04・42 §2-1)。Windows の Trim$ は Chr(32) しか落とさない
'   ため、TAB / LF / CR / 全角空白だけの文字列が「確認済み」として通っていた。
'   ここに**独自の前処理を挟まない**ことが肝で、裁定書40 S-m では提案書側
'   だけが NBSP を落としていたためレポート側と判定が食い違い、NBSP だけの
'   確認者名でレポートが「担当者が確認・編集したもの」になっていた
'   (=「片方だけ直す」型)。空白類の一覧は modUtilText.HasVisibleText が
'   唯一持ち、レポート(modExportHtml.ReviewerOf)と同じ1本を通る。
' ==========================================================
Public Function NeedsReviewMessage(ByVal reviewedBy As String) As String
    If Not modUtilText.HasVisibleText(reviewedBy) Then NeedsReviewMessage = EP_NEED_REVIEW
End Function

' ==========================================================
' StripFieldSeps - coverFields(EP_SEP=vbTab 区切り)へ入れる前の後始末。
' ----------------------------------------------------------
'   フィールドの**中身**が区切りを騙ると、modProposalHtml1.FieldAt が返す値が
'   1つずつずれる(会社名に vbTab を混ぜるだけで <title> と <noscript> の題が
'   会社名の後半に化け、日付・確認者まで全部ずれる。裁定書39 R2-05 と同型で、
'   顧客提示物である提案書側に残っていた=裁定書40 Q-M1)。
'   modUtilText.SanitizeInput は TAB と LF を「本文の構造」として**意図的に
'   残す**ので、区切りを使う側が落とす。落とすのは vbTab / vbLf / vbCr の
'   3つだけで、本文の他の文字には触れない。
'   **handoff**: 実装はレポート側 modExportHtml.StripFieldSeps(Private)と
'   逐語で同じである。司令塔が modUtilText へ1本に寄せたら、両方から
'   そちらを呼ぶこと(裁定書39 R2-04 が確認者名の判定でやったのと同じ整理。
'   modUtilText / modExportHtml は班Q2 の担当ファイルなので本波では触らない)。
' ==========================================================
Private Function StripFieldSeps(ByVal s As String) As String
    Dim t As String
    t = Replace(s, EP_SEP, vbNullString)
    t = Replace(t, vbCr, vbNullString)
    StripFieldSeps = Replace(t, vbLf, vbNullString)
End Function

' ==========================================================
' MissingProposalKeys - S5 の最外 required(EP_S5_REQUIRED)のうち無いキーを
'   ";" 区切りで返す(全部あれば "")。純関数なので層(a)から直接叩ける。
' ==========================================================
Public Function MissingProposalKeys(ByVal s5Json As String) As String
    MissingProposalKeys = modValidate4.MissingTopKeys( _
        modJsonLite.ExtractJsonBlock(s5Json), EP_S5_REQUIRED)
End Function

' ==========================================================
' ProposalFileBase - W12-c のファイル名(拡張子を除く)。組立の実体は
'   modUtilPath.BuildVersionedFileName(HTMLレポートと同じ1本)。
' ==========================================================
Public Function ProposalFileBase(ByVal company As String, ByVal dirText As String) As String
    ProposalFileBase = modUtilPath.BuildVersionedFileName(EP_HEAD, company, _
        modUtilText.IsoDateCompact(Date), _
        modConfig.GetStr("app_version", EP_VER_DEFAULT), dirText, EP_EXT)
End Function

' ==========================================================
' BuildProposalHtml - DATA組立からHTML全文までの**純組立関数**。
'   シート・ファイル・configに触れないので mock 素材だけで実行できる
'   (tools/render_proposal.py が実物のサンプルを出すのに使う)。
'   戻り値が空文字列 = 組立失敗。
' ==========================================================
Public Function BuildProposalHtml(ByVal metaJson As String, ByVal s2Json As String, _
                                  ByVal s5Json As String) As String
    Dim softened As Long
    Dim tabooLeft As String
    Dim strictLeft As String
    BuildProposalHtml = BuildProposalHtmlEx(metaJson, s2Json, s5Json, softened, _
                                            tabooLeft, strictLeft)
End Function

' ==========================================================
' BuildProposalHtmlEx - 上と同じだが、呼出側が記録に使う2つを返す。
'   softened  : 対訳表の機械置換の**箇所数**(裁定書39 R2-11)
'   tabooLeft : 置換しても顧客向け本文に**残った**社内語(";"区切り。裁定書40
'               S-M1)。mode=warn の対と、終端集合でない文脈の語は機械置換の
'               対象外なので残りうる。これは**警告であって不合格ではない**ので
'               戻り値には載せず、呼出側が usage_log へ記録する(裁定書40 §0)。
'   strictLeft: そのうち置換の取りこぼし(=実装の欠陥)だけ(裁定書43 §1-4)。
' ==========================================================
Public Function BuildProposalHtmlEx(ByVal metaJson As String, ByVal s2Json As String, _
                                    ByVal s5Json As String, _
                                    ByRef softened As Long, _
                                    ByRef tabooLeft As String, _
                                    ByRef strictLeft As String) As String
    On Error GoTo Failed
    softened = 0
    tabooLeft = vbNullString
    strictLeft = vbNullString
    If LenB(Trim$(metaJson)) = 0 Then Exit Function

    Dim dataJson As String
    dataJson = BuildProposalDataEx(metaJson, s2Json, s5Json, softened, tabooLeft, strictLeft)
    If LenB(dataJson) = 0 Then Exit Function

    ' 5項目とも**必ず** StripFieldSeps を通してから並べる(裁定書40 Q-M1)。
    ' 通し忘れが1項目でもあると、そこから後ろが全部1つずつずれる。
    Dim coverFields As String
    coverFields = StripFieldSeps(modJsonLite.GetStr(metaJson, "company")) & EP_SEP
    coverFields = coverFields & StripFieldSeps(modJsonLite.GetStr(metaJson, "title")) & EP_SEP
    coverFields = coverFields & StripFieldSeps(modJsonLite.GetStr(metaJson, "subtitle")) & EP_SEP
    coverFields = coverFields & StripFieldSeps(modJsonLite.GetStr(metaJson, "date")) & EP_SEP
    coverFields = coverFields & StripFieldSeps(modJsonLite.GetStr(metaJson, "reviewed_by"))

    BuildProposalHtmlEx = modProposalHtml1.BuildProposalDocument(dataJson, coverFields)
    Exit Function
Failed:
    ' **握りつぶさない**(裁定書43 司令塔)。ここで番号と説明を残さないと、
    '   呼出側は「提案書の組み立てに失敗しました。」しか出せず、原因が
    '   1件も記録に残らない。実際に LibreOffice の
    '   「Variable not defined: modValidate3」を丸1本の調査で突き止めた。
    modLog.LogError EP_CODE_FAIL, EP_SRC & ".BuildProposalHtmlEx", _
        "err=" & CStr(Err.Number) & " " & Err.Description
    BuildProposalHtmlEx = vbNullString
End Function

' ==========================================================
' BuildProposalMetaJson - 20章§3 の meta を組み立てる純関数。
'   **内部の値(tier / quality_mode / case_type / round_no / status)は
'   引数にも戻り値にも無い**。持たなければ漏れない。
' ==========================================================
Public Function BuildProposalMetaJson(ByVal company As String, ByVal s5Json As String, _
                                      ByVal generatedAt As String, ByVal appVersion As String, _
                                      ByVal reviewedBy As String, ByVal reviewedAt As String) As String
    Dim body As String
    body = modJsonLite.ExtractJsonBlock(s5Json)

    Dim s As String
    s = "{" & JStr("company", company)
    s = s & "," & JStr("title", modJsonLite.GetStr(body, "title"))
    s = s & "," & JStr("subtitle", modJsonLite.GetStr(body, "subtitle"))
    s = s & "," & JStr("date", DateHeadOf(generatedAt))
    s = s & "," & JStr("app_version", appVersion)
    s = s & "," & JStr("reviewed_by", reviewedBy)
    s = s & "," & JStr("reviewed_at", reviewedAt) & "}"
    BuildProposalMetaJson = s
End Function

' ==========================================================
' BuildProposalData - 20章§3 のDATA(ページに埋め込む唯一のJSON)。
'   {"meta":..,"stats":..,"categories":[..],"risks":[..],"p":{..}}
'   risks は影響度 x 起こりやすさ の降順(同点は番号の昇順)。
' ==========================================================
Public Function BuildProposalData(ByVal metaJson As String, ByVal s2Json As String, _
                                  ByVal s5Json As String) As String
    Dim softened As Long
    Dim tabooLeft As String
    Dim strictLeft As String
    BuildProposalData = BuildProposalDataEx(metaJson, s2Json, s5Json, softened, _
                                            tabooLeft, strictLeft)
End Function

' ==========================================================
' BuildProposalDataEx - 上と同じだが、呼出側が記録に使う2つを返す。
'   softened  : S2 由来の自由文に掛けた対訳表の機械置換の**箇所数**
'               (裁定書39 R2-11。置換したのに誰も数えていなかったため、
'                顧客文面が黙って書き換わっていた)
'   tabooLeft : 組み上がったDATA(=顧客の目に触れる本文のすべて)に**残った**
'               社内語(";"区切り。空=1語も残っていない。裁定書40 S-M1)。
'               mode=warn の対と、終端集合でない文脈にあった語は機械置換の
'               対象外なので残りうるし、`s5_edited` を人が直した経路は CheckS5 を
'               通らないので S5 側の本文にも残りうる。**DATA全体**を見るのは
'               この2経路を1箇所で押さえるためである。
'   strictLeft: そのうち**置換の取りこぼし**(=実装の欠陥)だけ(裁定書43 §1-4)。
'               判定が位置に依るので、**語の一覧ではなく本文**から数える
'               (一覧は語のうしろが ";" になり、";" は終端集合の字である)。
' ==========================================================
Public Function BuildProposalDataEx(ByVal metaJson As String, ByVal s2Json As String, _
                                    ByVal s5Json As String, _
                                    ByRef softened As Long, _
                                    ByRef tabooLeft As String, _
                                    ByRef strictLeft As String) As String
    Dim p As String
    softened = 0
    tabooLeft = vbNullString
    strictLeft = vbNullString
    p = modJsonLite.ExtractJsonBlock(s5Json)
    If LenB(Trim$(p)) = 0 Then Exit Function

    Dim items() As String
    Dim n As Long
    n = LoadRisks(s2Json, items)

    Dim s As String
    s = "{""meta"":" & metaJson
    s = s & ",""stats"":" & StatsJson(items, n, p)
    s = s & ",""categories"":[" & CategoriesBody(items, n) & "]"
    s = s & ",""risks"":[" & RisksBody(items, n, softened) & "]"
    s = s & ",""p"":" & p & "}"
    ' 顧客向け本文に残った社内語を**最後に1回だけ**数える(裁定書40 S-M1)。
    ' 不合格にはしない(ここで止めると語1つで提案書が作れなくなる)。
    tabooLeft = modValidate4.TabooHit(s)
    strictLeft = modValidate4.TabooHitStrict(s)
    BuildProposalDataEx = s
End Function

' ==========================================================
' StatsText - 15章§5.6 user の {{statsText}}(1行1項目)。**AIに数えさせない**
'   ための実数であり、値源はここ1箇所(20章§3)。
' ==========================================================
Public Function StatsText(ByVal s2Json As String, ByVal s3Json As String) As String
    Dim items() As String
    Dim n As Long
    n = LoadRisks(s2Json, items)

    Dim easy As Long, design As Long, nonIns As Long
    CountClasses items, n, easy, design, nonIns

    Dim ideas As Long
    ideas = modJsonLite.GetArrayItems(s3Json, "growth_ideas").Count

    Dim pri As Long
    pri = ideas
    If pri > 4 Then pri = 4

    Dim s As String
    s = "リスク総数=" & CStr(n) & vbLf
    s = s & "保険で備えやすい=" & CStr(easy) & vbLf
    s = s & "補償条件の設計が必要=" & CStr(design) & vbLf
    s = s & "保険以外の対策が中心=" & CStr(nonIns) & vbLf
    s = s & "成長支援アイデア=" & CStr(ideas) & vbLf
    s = s & "優先候補=" & CStr(pri)
    StatsText = s
End Function

' ==========================================================
' 内部(純関数)
' ==========================================================

' s2Json の risks[] を「影響 x 頻度 の降順・同点は番号の昇順」に並べて返す。
'   戻り値=件数。items は 0 から n-1 に1件ぶんのJSONが入る。
Private Function LoadRisks(ByVal s2Json As String, ByRef items() As String) As Long
    Dim col As Collection
    Set col = modJsonLite.GetArrayItems(s2Json, "risks")

    Dim n As Long
    n = col.Count
    ReDim items(0 To 0)
    If n = 0 Then
        LoadRisks = 0
        Exit Function
    End If

    ReDim items(0 To n - 1)
    Dim i As Long
    For i = 1 To n
        items(i - 1) = CStr(col(i))
    Next i

    ' 単純な挿入整列(件数は数十件のため十分)。
    Dim j As Long
    Dim keyText As String
    For i = 1 To n - 1
        keyText = items(i)
        j = i - 1
        Do While j >= 0
            If Not Precedes(items(j), keyText) Then
                items(j + 1) = items(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        items(j + 1) = keyText
    Next i
    LoadRisks = n
End Function

' a が b より前に並ぶか(スコア降順 -> 番号昇順)。
Private Function Precedes(ByVal a As String, ByVal b As String) As Boolean
    Dim sa As Long, sb As Long
    sa = ScoreOf(a)
    sb = ScoreOf(b)
    If sa <> sb Then
        Precedes = (sa > sb)
        Exit Function
    End If
    Precedes = (modJsonLite.GetLong(a, "risk_no", 0) <= modJsonLite.GetLong(b, "risk_no", 0))
End Function

Private Function ScoreOf(ByVal rj As String) As Long
    ScoreOf = modJsonLite.GetLong(rj, "impact_score", 0) * _
              modJsonLite.GetLong(rj, "frequency_score", 0)
End Function

' 3分類の件数(19章§3 transferability)。
Private Sub CountClasses(ByRef items() As String, ByVal n As Long, _
                         ByRef easy As Long, ByRef design As Long, ByRef nonIns As Long)
    Dim i As Long
    Dim t As String
    easy = 0
    design = 0
    nonIns = 0
    For i = 0 To n - 1
        t = Trim$(modJsonLite.GetStr(items(i), "transferability"))
        If t = "cover" Then
            easy = easy + 1
        ElseIf t = "partial" Then
            design = design + 1
        ElseIf t = "hard" Then
            nonIns = nonIns + 1
        End If
    Next i
End Sub

Private Function StatsJson(ByRef items() As String, ByVal n As Long, ByVal p As String) As String
    Dim easy As Long, design As Long, nonIns As Long
    CountClasses items, n, easy, design, nonIns

    Dim ideas As Long, pri As Long
    Dim it As Variant
    ' S5 の `headline` は `ideas` と**同名の文字列キー**を先に持つため、
    ' modJsonLite.GetArrayItems を直接呼ぶと0件と数えてしまう。値が [ で
    ' 始まる出現だけを採る modValidate4.ArrItems を通す(20章§3)。
    For Each it In modValidate4.ArrItems(p, "ideas")
        ideas = ideas + 1
        If modJsonLite.GetBoolJ(CStr(it), "priority", False) Then pri = pri + 1
    Next it

    Dim s As String
    s = "{""risk_total"":" & CStr(n)
    s = s & ",""easy"":" & CStr(easy)
    s = s & ",""design"":" & CStr(design)
    s = s & ",""non_ins"":" & CStr(nonIns)
    s = s & ",""ideas"":" & CStr(ideas)
    s = s & ",""ideas_priority"":" & CStr(pri) & "}"
    StatsJson = s
End Function

Private Function CategoriesBody(ByRef items() As String, ByVal n As Long) As String
    Dim keys() As String, labels() As String
    keys = Split(EP_CAT_KEYS, EP_SEMI)
    labels = Split(EP_CAT_LABELS, EP_SEMI)

    Dim acc As String
    Dim k As Long, i As Long, c As Long
    For k = LBound(keys) To UBound(keys)
        c = 0
        For i = 0 To n - 1
            If Trim$(modJsonLite.GetStr(items(i), "category")) = keys(k) Then c = c + 1
        Next i
        If LenB(acc) > 0 Then acc = acc & ","
        acc = acc & "{" & JStr("label", labels(k)) & ",""count"":" & CStr(c) & "}"
    Next k
    CategoriesBody = acc
End Function

' 顧客向けに**選んだ列だけ**を書き出す(20章§3の表)。S2 由来の自由文は
'   modValidate4.SoftenTaboo を通してから入れる(社内語をお客さまに見せない)。
Private Function RisksBody(ByRef items() As String, ByVal n As Long, _
                           ByRef changed As Long) As String
    Dim acc As String
    Dim i As Long
    Dim rj As String

    changed = 0
    For i = 0 To n - 1
        rj = items(i)
        If LenB(acc) > 0 Then acc = acc & ","
        acc = acc & "{""risk_no"":" & CStr(modJsonLite.GetLong(rj, "risk_no", 0))
        acc = acc & "," & JStr("category_label", CatLabelOf(modJsonLite.GetStr(rj, "category")))
        acc = acc & "," & JStr("risk_name", Soft(modJsonLite.GetStr(rj, "risk_name"), changed))
        acc = acc & ",""impact"":" & CStr(modJsonLite.GetLong(rj, "impact_score", 0))
        acc = acc & ",""frequency"":" & CStr(modJsonLite.GetLong(rj, "frequency_score", 0))
        acc = acc & "," & JStr("rank_label", RankLabelOf(ScoreOf(rj)))
        acc = acc & "," & JStr("class_label", ClassLabelOf(modJsonLite.GetStr(rj, "transferability")))
        acc = acc & "," & JStr("line_note", Soft(modJsonLite.GetStr(rj, "line_note"), changed))
        acc = acc & "," & JStr("check_note", Soft(modJsonLite.GetStr(rj, "gap_note"), changed))
        acc = acc & "," & JStr("control_note", Soft(modJsonLite.GetStr(rj, "control_note"), changed))
        acc = acc & "}"
    Next i
    RisksBody = acc
End Function

' 対訳表による機械置換(顧客向けの最後の砦。裁定書38 班C 2)。
'   SoftenTaboo は呼ぶたび第2引数を 0 に戻すので、**ここで足し込む**
'   (裁定書39 R2-11。以前は最後の1回ぶんしか残らず、しかも誰も読まなかった)。
Private Function Soft(ByVal t As String, ByRef total As Long) As String
    Dim n As Long

    Soft = modValidate4.SoftenTaboo(t, n)
    total = total + n
End Function

Private Function CatLabelOf(ByVal keyText As String) As String
    Dim keys() As String, labels() As String
    keys = Split(EP_CAT_KEYS, EP_SEMI)
    labels = Split(EP_CAT_LABELS, EP_SEMI)

    Dim k As Long
    For k = LBound(keys) To UBound(keys)
        If keys(k) = Trim$(keyText) Then
            CatLabelOf = labels(k)
            Exit Function
        End If
    Next k
    CatLabelOf = "その他"
End Function

' 評価ラベル(20章§4の6枚目・9枚目・20枚目で共通に使う)。
Public Function RankLabelOf(ByVal score As Long) As String
    If score >= 16 Then
        RankLabelOf = "最優先"
    ElseIf score >= 12 Then
        RankLabelOf = "高"
    ElseIf score >= 6 Then
        RankLabelOf = "中"
    Else
        RankLabelOf = "低"
    End If
End Function

' 備えやすさ(3分類)の顧客向けラベル。
Public Function ClassLabelOf(ByVal transferability As String) As String
    Select Case Trim$(transferability)
    Case "cover"
        ClassLabelOf = "保険で備えやすい"
    Case "partial"
        ClassLabelOf = "補償条件の設計が必要"
    Case "hard"
        ClassLabelOf = "保険以外の対策が中心"
    Case Else
        ClassLabelOf = "確認中"
    End Select
End Function

' 13章§2.2 の解決順(s5_edited > s5_json)。**空白類だけの s5_edited は
'   「直していない」として扱い s5_json へ落とす**(裁定書39 R2-04 と同型。
'   Trim$ だけだと TAB/全角空白だけの編集が本文を上書きしてしまう)。
Private Function ResolveProposalJson(ByVal caseId As String) As String
    Dim t As String
    t = modCaseStore.LoadData(caseId, "s5_edited")
    If modUtilText.HasVisibleText(t) Then
        ResolveProposalJson = t
        Exit Function
    End If
    ResolveProposalJson = modCaseStore.LoadData(caseId, "s5_json")
End Function

' 生成時刻 "yyyy/mm/dd hh:nn:ss" から表紙の日付 "yyyy/mm/dd" を切り出す。
Private Function DateHeadOf(ByVal stampText As String) As String
    Dim p As Long
    p = InStr(1, stampText, " ", vbBinaryCompare)
    If p > 1 Then
        DateHeadOf = Left$(stampText, p - 1)
    Else
        DateHeadOf = stampText
    End If
End Function

' ==========================================================
' TabooLeftNote - 顧客向け本文に**残った**社内語を usage_log へ1行残す
'   (裁定書40 S-M1 / 43 §1-4)。mode=warn の対と、終端集合でない文脈にあった
'   ために置換しなかった語が対象で、**件数**(proposal_jargon_left=N)と
'   V-S5-12 の警告文(語の一覧)の両方を残す。出力は止めない: ここは検証では
'   なく注記のチャネルである(裁定書40 §0)。営業向けレポートには出さない。
'   置換の取りこぼし(strictLeft。本文から数えたもの)があれば、それは実装の
'   欠陥なので**別の行**で区別して残す(strict=)。
' ==========================================================
Private Sub TabooLeftNote(ByVal caseId As String, ByVal tabooLeft As String, _
                          ByVal strictLeft As String)
    If LenB(tabooLeft) = 0 Then Exit Sub

    modLog.LogUsage "proposal_taboo_left", caseId, _
        "proposal_jargon_left=" & CStr(SepCount(tabooLeft)) & " " & _
        modValidate4.TabooWarnLine(tabooLeft)

    If LenB(strictLeft) > 0 Then
        modLog.LogUsage "proposal_taboo_left", caseId, "strict=" & strictLeft
    End If
End Sub

' ";" 区切りの語数(空なら0)。Split() の戻り値へ直接添字を付けない
'   (LibreOffice Basic が解さない。他のモジュールと同じ書き方)。
Private Function SepCount(ByVal listText As String) As Long
    Dim parts() As String

    If LenB(listText) = 0 Then Exit Function
    parts = Split(listText, EP_SEMI)
    SepCount = UBound(parts) - LBound(parts) + 1
End Function

' PII走査(16章 E-05(6))。検知種別と箇所だけを記録する(本文は残さない
'   =NFR-S3)。顧客向け資料なので画面にも紙にも警告を出さない。
Private Sub PiiNote(ByVal caseId As String, ByVal s2Text As String, ByVal s5Text As String)
    On Error GoTo Failed
    Dim acc As String
    acc = modPii.ScanReport(s2Text, "proposal/s2")
    If LenB(modPii.ScanReport(s5Text, "proposal/s5")) > 0 Then
        acc = acc & ";" & modPii.ScanReport(s5Text, "proposal/s5")
    End If
    If LenB(acc) > 0 Then modLog.LogError "E0103", EP_SRC & ".PiiNote", acc
    Exit Sub
Failed:
    modLog.LogUsage "proposal_pii_scan_failed", caseId, "err=" & CStr(Err.Number)
End Sub

Private Function OutDirRaw() As String
    Dim t As String
    t = Trim$(modConfig.GetStr("html_out_dir", vbNullString))
    If LenB(t) = 0 Then t = Trim$(modConfig.GetStr("data_dir", vbNullString))
    OutDirRaw = t
End Function

' 同名ファイルは上書きせず _2 _3 と連番で空きを探す(13章§2.8)。
Private Function UniqueOutPath(ByVal dirText As String, ByVal baseName As String) As String
    If LenB(Trim$(baseName)) = 0 Then Exit Function

    Dim candidate As String
    candidate = modUtilPath.JoinPath(dirText, baseName & EP_EXT)
    If Not OutFileExists(candidate) Then
        UniqueOutPath = candidate
        Exit Function
    End If

    Dim n As Long
    For n = 2 To EP_SERIAL_MAX
        candidate = modUtilPath.JoinPath(dirText, baseName & "_" & CStr(n) & EP_EXT)
        If Not OutFileExists(candidate) Then
            UniqueOutPath = candidate
            Exit Function
        End If
    Next n
End Function

Private Function OutFileExists(ByVal pathText As String) As Boolean
    On Error GoTo Unknown0
    OutFileExists = (LenB(Dir$(pathText)) > 0)
    Exit Function
Unknown0:
    OutFileExists = True
End Function

Private Function JStr(ByVal keyName As String, ByVal valueText As String) As String
    JStr = """" & keyName & """:""" & modJsonLite.EscapeJsonStr(valueText) & """"
End Function
