Attribute VB_Name = "modExportHtml"
Option Explicit

' ==========================================================
' modExportHtml - HTMLリスクレポートの生成(10章FR-37・14章§6・18章)
' ------------------------------------------------
' 責務(18章§1の表): 案件からJSONを解決し、匿名化復元・PII走査・DATA組立・
'   テンプレ組立・ファイル書出・report_path 記録を行う。
'   **HTML本文・CSS・JSの文字列をここに書かない**(12章§4。30,000字契約に
'   抵触するため)。テンプレ本体は modHtmlTemplate1..n、テーマCSSは modHtmlTheme。
'
' 生成手順(18章§1.1。この順序で行う):
'   (1) modCaseStore.ResolveStepJson で S1・S2・S3 を解決。S1が空なら生成しない
'       (エラーコードは立てず「先にStep1を実行してください」と案内して中止)。
'       S2・S3が空でも生成は続行し、該当セクションは登録表の empty で処理する。
'   (2) 匿名化の復元(16章 E-31)。**必ずエスケープより前**に行う。
'   (3) modPii 走査(16章 E-05(6))。検知しても生成はブロックせず警告に回し、
'       run_log/err_log には検知種別と箇所のみを記録する(本文は残さない=NFR-S3)。
'   (4) DATA(18章§2)を1本のJSON文字列として組み立てる。失敗は E0502。
'   (5) modHtmlTemplate1.BuildDocument でHTML全文を組み立てる。失敗は E0502。
'   (6) UTF-8(BOMあり)で書き出す。失敗は E0502。書き出しの実体は純VBAの
'       modUtil.WriteUtf8File(裁定書27 W9-B2 で ADODB.Stream を撤去した)。
'       出力先不存在は先に16章 E-21 のフォールバック(自動作成 -> Documents直下)。
'   (7) 確定パスを案件一覧の report_path に記録する。失敗は警告のみ。
'
' 生成にLLMを使わない(10章FR-37・16章NFR-S4)。案件の status は生成の成否で
'   変化させない(16章 E-48)。
'
' R4(12章§2・§4): 本モジュールは Excelトークン許可リストに**入っていない**。
'   シート・ブックには一切触れず、案件データは modCaseStore / modCaseRead 経由、
'   configは modConfig 経由、案件一覧への書込は modCaseStore 経由で行う。
' CP932準拠: 本文・注釈ともに CP932 内の文字だけで書く(絵文字不可)。
' ==========================================================

Private Const EX_SRC As String = "modExportHtml"
Private Const EX_EXT As String = ".html"
Private Const EX_TAIL As String = "リスクレポート"
Private Const EX_SEP As String = vbTab
Private Const EX_WARN_SEP As String = vbLf
Private Const EX_PH_COMPANY As String = "{{COMPANY}}"
Private Const EX_PH_POLICY As String = "{{POLICY_NO}}"
Private Const EX_THEME_DEFAULT As String = "standard"
' 保存先(裁定書27 W9-C2)。html_out_dir が空なら data_dir に従う。解決(候補の
' 並べ方・実在確認・作成)は modUtil.ResolveDataDir が唯一持つ。**既定は空**
' (裁定書31 裁定1・司令塔裁定: コード上に環境変数のリテラルを残さない。空の
' まま渡せば data_dir.txt -> config -> 本体と同じフォルダ \データ の順に決まる)。
Private Const EX_OUT_DEFAULT As String = vbNullString
Private Const EX_VER_DEFAULT As String = "2.0.0"
Private Const EX_CODE_FAIL As String = "E0502"

' 裁定書9 B4(13章§2.8): 同名ファイルは上書きせず _2 _3 と連番で空きを探す。
' 探索の上限(防御。実運用で到達しない)。超えたら書出失敗として扱う。
Private Const EX_SERIAL_MAX As Long = 9999

' ==========================================================
' GenerateHtmlReport - 14章§6の契約。""=成功 / 非空=失敗理由。
'   outPath には成功時の確定パスを返す(失敗時は空のまま)。
'   **担当者の確認を名乗らない既定の書き出し**(裁定書37 B-06)。確認済みで
'   書き出す口は GenerateHtmlReportEx。既存の呼出側を壊さないため、本関数は
'   reviewedBy="" で Ex へ委譲するだけにする(実装を2本に分けない)。
' ==========================================================
Public Function GenerateHtmlReport(ByVal caseId As String, ByRef outPath As String) As String
    GenerateHtmlReport = GenerateHtmlReportEx(caseId, outPath, vbNullString)
End Function

' ==========================================================
' GenerateHtmlReportEx - 確認者つきの書き出し(裁定書37 B-06。18章§3.5)。
'   reviewedBy: レポートを客先へ出す前に内容を確認・編集した担当者の表示名。
'     **空なら「AI生成・担当者確認前」の既定文**が出る(16章NFR-S5・社内IT環境
'     v1.1 §7.3「顧客提示は全件、利用者による確認が必須」に対し、担保の無い
'     遵守宣言をしない)。非空なら meta.reviewed_by / reviewed_at が入り、
'     免責1行目と表紙チップが「確認済」へ切り替わる。
'   reviewedAt は**ここで打つ**(呼出側が時刻を作らない)。
' ==========================================================
Public Function GenerateHtmlReportEx(ByVal caseId As String, ByRef outPath As String, _
                                     ByVal reviewedBy As String) As String
    On Error GoTo Failed
    outPath = vbNullString

    Dim reviewer As String, reviewedAt As String
    reviewer = Trim$(modUtilText.SanitizeInput(reviewedBy))
    If LenB(reviewer) > 0 Then reviewedAt = modUtil.NowStamp()

    If Not modCaseStore.IsValidCaseId(caseId) Then
        modLog.LogError "E0101", EX_SRC & ".GenerateHtmlReportEx", "invalid_case_id"
        GenerateHtmlReportEx = "案件IDが不正です。"
        Exit Function
    End If

    Dim ctx As TCaseCtx
    Dim roundNo As Long
    Dim qualityMode As String, s4Variant As String, tierText As String
    If Not modCaseRead.ReadCaseCtx(caseId, ctx, roundNo, qualityMode, s4Variant, tierText) Then
        GenerateHtmlReportEx = "案件一覧からこの案件を読めませんでした。"
        Exit Function
    End If

    ' (1) S1・S2・S3の解決。S1が空ならエラーコードを立てずに中止する
    '     (データ未整備は障害ではない。18章§1.1(1))。
    Dim s1Text As String, s2Text As String, s3Text As String
    s1Text = modCaseStore.ResolveStepJson(caseId, 1)
    s2Text = modCaseStore.ResolveStepJson(caseId, 2)
    s3Text = modCaseStore.ResolveStepJson(caseId, 3)
    If LenB(Trim$(s1Text)) = 0 Then
        modLog.LogUsage "html_report_skipped", caseId, "s1_empty"
        GenerateHtmlReportEx = "先にStep1を実行してください。"
        Exit Function
    End If

    ' 裁定書37 B-03: 原文照合は**匿名化の復元より前**に行う(貼付原文もS2も
    ' 保存時の表記のまま突き合わせる。復元後に照合すると社名の置換で全件が
    ' 未照合になる)。config ground_check=FALSE なら測らない=空のまま。
    Dim groundNote As String
    If modConfig.GetBool("ground_check", True) Then
        groundNote = modGround.GroundNotes(s2Text, _
            modPipeline3.BuildHaystack(caseId, s1Text), _
            modConfig.GetLong("ground_head_chars", modGround.GR_HEAD_DEFAULT))
    End If

    ' (2) 匿名化の復元(E-31)。**エスケープより前**に行う=復元後の実名が
    '     エスケープ対象になるようにする。復元表が無い {{POLICY_NO}} は
    '     プレースホルダのまま出し、警告に載せる(黙って消さない)。
    Dim warnText As String
    s1Text = RestoreAnonymized(s1Text, ctx.company, warnText)
    s2Text = RestoreAnonymized(s2Text, ctx.company, warnText)
    s3Text = RestoreAnonymized(s3Text, ctx.company, warnText)

    ' JSONとして読めない S2 / S3 は §2 により null になり該当セクションが
    ' 消える。**黙って消さない**ため、読めなかったことを警告に載せる。
    warnText = WarnIfUnreadable(s2Text, "Step2", warnText)
    warnText = WarnIfUnreadable(s3Text, "Step3", warnText)

    ' (3) PII走査。**生成はブロックしない**(18章§1.1(3)・16章 E-05(6))。
    '     裁定書9 B17: 走査自体の失敗も黙らせない(warnText へ警告を積んで続行)。
    Dim piiText As String
    piiText = PiiNote(caseId, s1Text, s2Text, s3Text, warnText)
    If LenB(piiText) > 0 Then
        modLog.LogError "E0103", EX_SRC & ".GenerateHtmlReportEx", piiText
        warnText = AddWarn(warnText, "個人情報らしき記述を検知しました。配布前に本文をご確認ください。")
    End If

    ' テーマ解決(18章§5.2)。未知のテーマ名は standard へ落とし、黙って戻さず
    ' run_log の detail に theme_fallback を記録する。
    Dim themeName As String
    themeName = ResolveTheme(modConfig.GetStr("html_theme", EX_THEME_DEFAULT), caseId)

    ' 裁定書38 班A(V-S1-14 / V-S1-15): S1の出典と接頭辞の点検。**落とさない**。
    ' 照合する原文は貼付原文だけ(BuildHaystack の第2引数に s1 を渡さない)。
    ' 匿名化の復元より前の s1Text を使う(復元でURLは変わらないが、社名の置換で
    ' 前後が動くため、run_log と同じ土俵で数えるには復元前が正しい)。
    Dim s1WarnNote As String
    s1WarnNote = modValidate3.WarnNoteOf(modValidate3.CheckS1Notes(s1Text, _
                     modPipeline3.BuildHaystack(caseId, vbNullString)))

    Dim metaJson As String
    metaJson = BuildMetaJson(caseId, ctx.company, ctx.industry_code, ctx.industry_name, _
                             ctx.case_type, tierText, qualityMode, roundNo, s4Variant, _
                             modUtil.NowStamp(), modConfig.GetStr("app_version", EX_VER_DEFAULT), _
                             themeName, warnText, reviewer, reviewedAt, groundNote, s1WarnNote)

    ' (4)(5) DATA組立とテンプレ組立。どちらの失敗も E0502(16章 E-48)。
    Dim docText As String
    docText = BuildReportHtml(metaJson, s1Text, s2Text, s3Text, themeName)
    If LenB(docText) = 0 Then
        modLog.LogError EX_CODE_FAIL, EX_SRC & ".GenerateHtmlReportEx", "build_failed"
        GenerateHtmlReportEx = "レポートの組み立てに失敗しました。"
        Exit Function
    End If

    ' (6) 書き出し。出力先不存在は先に E-21 のフォールバックを試す。
    Dim dirText As String
    dirText = ResolveOutDir(OutDirRaw())
    If LenB(dirText) = 0 Then
        modLog.LogError EX_CODE_FAIL, EX_SRC & ".GenerateHtmlReportEx", "no_out_dir"
        GenerateHtmlReportEx = "出力先フォルダを用意できませんでした。"
        Exit Function
    End If

    ' ファイル名は modUtilText.BuildFileNameSafe(13章§2.8の規則)を通す。
    ' HTMLへは入らない値の連結なので NFR-S7(3) の対象外。
    ' 裁定書9 B4: 末尾に _yyyymmdd を付与し、同名が既に在れば _2 _3 と連番を
    ' 探して**新規ファイルとして作る**(過去の出力を消さない。13章§2.8)。
    Dim pathText As String
    pathText = UniqueOutPath(dirText, FileNameOf(ctx.company, caseId, dirText))
    If LenB(pathText) = 0 Then
        modLog.LogError EX_CODE_FAIL, EX_SRC & ".GenerateHtmlReportEx", "no_free_filename"
        GenerateHtmlReportEx = "出力ファイル名の空きを見つけられませんでした。"
        Exit Function
    End If
    If Not WriteUtf8Bom(pathText, docText) Then
        modLog.LogError EX_CODE_FAIL, EX_SRC & ".GenerateHtmlReportEx", "write_failed"
        GenerateHtmlReportEx = "ファイルの書き出しに失敗しました。"
        Exit Function
    End If

    ' (7) 確定パスの記録。失敗しても生成済みファイルは残す(警告のみ)。
    If Not modCaseStore.SetReportPath(caseId, pathText) Then
        modLog.LogUsage "html_report_path_unrecorded", caseId, "set_report_path_failed"
    End If

    modLog.LogUsage "html_report", caseId, "theme=" & themeName & ";chars=" & CStr(Len(docText))
    outPath = pathText
    Exit Function

Failed:
    modLog.LogError EX_CODE_FAIL, EX_SRC & ".GenerateHtmlReportEx", "unexpected", Err.Number
    GenerateHtmlReportEx = "レポート生成中にエラーが発生しました。"
End Function

' ==========================================================
' BuildMetaJson - 18章§2の meta を組み立てる純関数。
'   reviewedBy / reviewedAt は裁定書37 B-06 の確認フラグ(未確認は両方空文字)。

'   s1WarnNote は裁定書38 班A の s1_warn("V-S1-14:2;V-S1-15:1" 形式・";" 区切り。
'   空=指摘なし)。値源は modValidate3.WarnNoteOf。
'   groundNote は裁定書37 B-03 の未照合 risk_no(";" 区切り。空=全件照合できた
'   か、検査していない)。
'   warnText は EX_WARN_SEP 区切りの警告文(空なら warnings は空配列)。
'   warnings は18章§4.1が静的HTMLの差込口を3箇所に限っているため、警告バナーの
'   本文もDATA経由でJS側へ渡す(textContent で描くのでエスケープ経路が増えない)。
' ==========================================================
Public Function BuildMetaJson(ByVal caseId As String, ByVal company As String, _
                              ByVal industryCode As String, ByVal industryName As String, _
                              ByVal caseType As String, ByVal dossierTier As String, _
                              ByVal qualityMode As String, ByVal roundNo As Long, _
                              ByVal s4Variant As String, ByVal generatedAt As String, _
                              ByVal appVersion As String, ByVal themeName As String, _
                              ByVal warnText As String, ByVal reviewedBy As String, _
                              ByVal reviewedAt As String, _
                              ByVal groundNote As String, _
                              Optional ByVal s1WarnNote As String = vbNullString) As String
    Dim s As String
    s = s & "{" & JStr("case_id", caseId) ' SAFE:html
    s = s & "," & JStr("company", company) ' SAFE:html
    s = s & "," & JStr("industry_code", industryCode) ' SAFE:html
    s = s & "," & JStr("industry_name", industryName) ' SAFE:html
    s = s & "," & JStr("case_type", caseType) ' SAFE:html
    s = s & "," & JStr("dossier_tier", dossierTier) ' SAFE:html
    s = s & "," & JStr("quality_mode", qualityMode) ' SAFE:html
    s = s & ",""round_no"":" & CStr(roundNo) ' SAFE:html
    s = s & "," & JStr("s4_variant", s4Variant) ' SAFE:html
    s = s & "," & JStr("generated_at", generatedAt) ' SAFE:html
    s = s & "," & JStr("app_version", appVersion) ' SAFE:html
    s = s & "," & JStr("theme", themeName) ' SAFE:html
    ' 裁定書37 B-06(18章§2): 担当者の確認。未確認なら**両方とも空文字**にする
    ' (キー自体は必ず置く。テンプレ側は reviewed_by の空/非空だけで分岐する)。
    s = s & "," & JStr("reviewed_by", reviewedBy) ' SAFE:html
    s = s & "," & JStr("reviewed_at", reviewedAt) ' SAFE:html
    ' 裁定書37 B-03(18章§2): 原文と照合できなかった risk_no の一覧。
    s = s & ",""ground_unmatched"":[" & modGround.NoteJsonArray(groundNote) & "]" ' SAFE:html
    ' 裁定書38 班A(18章§2): S1の警告("V-S1-14:2" 等)。無ければ空配列。
    s = s & ",""s1_warn"":[" & modGround.NoteJsonArray(s1WarnNote) & "]" ' SAFE:html
    s = s & ",""warnings"":[" & WarnArrayBody(warnText) & "]}" ' SAFE:html
    BuildMetaJson = s
End Function

' ==========================================================
' BuildReportHtml - DATA組立からHTML全文までの**純組立関数**(18章§1.1(4)(5))。
'   シート・ファイル・configに触れないので、mock素材だけでそのまま実行できる
'   (tools/render_report.py が実物のサンプルを出すのに使う)。
'   引数のJSONは modJsonLite.ExtractJsonBlock を通してから埋める(前後の説明文・
'   コードフェンス混入を落とす。14章§5)。s2/s3 が空なら null を置く(§2)。
'   戻り値が空文字列 = 組立失敗(呼び出し側が E0502 を立てる)。
' ==========================================================
Public Function BuildReportHtml(ByVal metaJson As String, ByVal s1Json As String, _
                                ByVal s2Json As String, ByVal s3Json As String, _
                                ByVal themeName As String) As String
    On Error GoTo Failed

    If LenB(Trim$(metaJson)) = 0 Then Exit Function

    ' 18章§2: s1/s2/s3 は15章スキーマの出力を**キー名を変えずそのまま**入れる。
    ' 未実行のStepは null を置く(キー自体は必ず置く)。
    Dim dataJson As String
    dataJson = "{""meta"":" & metaJson ' SAFE:html
    dataJson = dataJson & ",""s1"":" & OrNull(s1Json) ' SAFE:html
    dataJson = dataJson & ",""s2"":" & OrNull(s2Json) ' SAFE:html
    dataJson = dataJson & ",""s3"":" & OrNull(s3Json) & "}" ' SAFE:html
    ' ここで生のJSONを連結してよいのは、この文字列が modHtmlTemplate1.BuildDocument
    ' の中で **modUtilText.JsStringSafe を通ってから** 1本のJS文字列リテラルへ
    ' 入るためである(18章§5.3(1)・16章 E-47)。DOMへはJS側が textContent で
    ' 置くので、HTMLとして解釈される経路はここにも下流にも存在しない。

    ' 表紙の3値。HTMLへ差し込むのは BodyShellHtml / HeadHtml の中であり、
    ' いずれも modUtilText.HtmlSafe を通ってから入る(18章§4.1・§5.3(2))。
    Dim coverFields As String
    coverFields = modJsonLite.GetStr(metaJson, "company") & EX_SEP ' SAFE:html
    coverFields = coverFields & modJsonLite.GetStr(metaJson, "case_id") & EX_SEP ' SAFE:html
    coverFields = coverFields & modJsonLite.GetStr(metaJson, "generated_at") & EX_SEP ' SAFE:html
    ' 裁定書37 B-06: <noscript>(JSが動かない環境)の免責も同じ3項分岐にするため、
    ' 確認者と確認日時を静的HTML側へも渡す(差込は HtmlSafe を通る。18章§4.1)。
    coverFields = coverFields & modJsonLite.GetStr(metaJson, "reviewed_by") & EX_SEP ' SAFE:html
    coverFields = coverFields & modJsonLite.GetStr(metaJson, "reviewed_at") ' SAFE:html

    BuildReportHtml = modHtmlTemplate1.BuildDocument(themeName, dataJson, coverFields)
    Exit Function

Failed:
    BuildReportHtml = vbNullString
End Function

' ==========================================================
' ResolveTheme - config html_theme を modHtmlTheme.ThemeNames() と突き合わせ、
'   未知のテーマ名なら standard へフォールバックして run_log に記録する
'   (18章§5.2「黙って既定に戻さない」)。
' ==========================================================
Private Function ResolveTheme(ByVal wanted As String, ByVal caseId As String) As String
    Dim want As String
    want = LCase$(Trim$(wanted))
    If LenB(want) = 0 Then want = EX_THEME_DEFAULT

    Dim names() As String
    names = modUtil.SplitKeepNonEmpty(modHtmlTheme.ThemeNames(), ";")
    Dim i As Long
    For i = LBound(names) To UBound(names)
        If LCase$(Trim$(names(i))) = want Then
            ResolveTheme = want
            Exit Function
        End If
    Next i

    modLog.LogUsage "html_theme_fallback", caseId, "theme_fallback=" & want
    ResolveTheme = EX_THEME_DEFAULT
End Function

' 匿名化の復元(16章 E-31)。{{COMPANY}} は案件一覧の実名へ戻す。{{POLICY_NO}}
'   の復元表は本製品が保持していないため、残っていればプレースホルダのまま
'   出して警告に載せる(18章§1.1(2)の「復元できない場合」の扱い)。
Private Function RestoreAnonymized(ByVal jsonText As String, ByVal company As String, _
                                   ByRef warnText As String) As String
    RestoreAnonymized = jsonText
    If LenB(jsonText) = 0 Then Exit Function

    Dim t As String
    t = jsonText
    If InStr(1, t, EX_PH_COMPANY, vbBinaryCompare) > 0 Then
        If LenB(Trim$(company)) > 0 Then
            t = Replace(t, EX_PH_COMPANY, company)
        Else
            warnText = AddWarn(warnText, "匿名化の復元ができませんでした(会社名)。")
        End If
    End If
    If InStr(1, t, EX_PH_POLICY, vbBinaryCompare) > 0 Then
        warnText = AddWarn(warnText, "匿名化の復元ができませんでした(証券番号)。")
    End If
    RestoreAnonymized = t
End Function

' PII走査(16章 E-05(6))。検知箇所と種別だけを返す(本文は返さない=NFR-S3)。
'   裁定書9 B17(18章§1.1(3)): On Error GoTo 方式。走査が落ちたときは警告を
'   warnText へ積んで続行する(冒頭の On Error Resume Next で3本まとめて覆うと
'   警告も E0103 も出ないまま合格に見える経路が残るため)。
Private Function PiiNote(ByVal caseId As String, ByVal s1Text As String, _
                         ByVal s2Text As String, ByVal s3Text As String, _
                         ByRef warnText As String) As String
    On Error GoTo Failed
    Dim acc As String
    acc = JoinNote(acc, modPii.ScanReport(s1Text, "html/s1"))
    acc = JoinNote(acc, modPii.ScanReport(s2Text, "html/s2"))
    acc = JoinNote(acc, modPii.ScanReport(s3Text, "html/s3"))
    PiiNote = acc
    Exit Function

Failed:
    modLog.LogUsage "html_pii_scan_failed", caseId, "err=" & CStr(Err.Number)
    warnText = AddWarn(warnText, "個人情報の走査に失敗しました。配布前に本文をご確認ください")
    PiiNote = acc
End Function

Private Function JoinNote(ByVal acc As String, ByVal added As String) As String
    JoinNote = acc
    If LenB(added) = 0 Then Exit Function
    If LenB(acc) = 0 Then
        JoinNote = added
    Else
        JoinNote = acc & ";" & added ' SAFE:html
    End If
End Function

' 本文はあるのにJSONとして取り出せない Step を警告に載せる。DATAでは null に
'   なり該当セクションが消える(18章§2・§3)ので、消えた理由を読み手へ残す。
Private Function WarnIfUnreadable(ByVal jsonText As String, ByVal stepLabel As String, _
                                  ByVal acc As String) As String
    WarnIfUnreadable = acc
    If LenB(Trim$(jsonText)) = 0 Then Exit Function
    If LenB(Trim$(modJsonLite.ExtractJsonBlock(jsonText))) > 0 Then Exit Function
    WarnIfUnreadable = AddWarn(acc, stepLabel & "の結果をJSONとして読み取れなかったため、" & _
                                    "該当する節を表示していません。") ' SAFE:html
End Function

Private Function AddWarn(ByVal acc As String, ByVal added As String) As String
    If LenB(acc) = 0 Then
        AddWarn = added
    Else
        AddWarn = acc & EX_WARN_SEP & added ' SAFE:html
    End If
End Function

' warnText(EX_WARN_SEP区切り) を JSON配列の中身へ。
Private Function WarnArrayBody(ByVal warnText As String) As String
    If LenB(Trim$(warnText)) = 0 Then Exit Function
    Dim rows() As String
    rows = modUtil.SplitKeepNonEmpty(warnText, EX_WARN_SEP)
    Dim acc As String
    Dim i As Long
    For i = LBound(rows) To UBound(rows)
        If LenB(acc) > 0 Then
            acc = acc & ","
        End If
        acc = acc & """" & modJsonLite.EscapeJsonStr(rows(i)) & """" ' SAFE:html
    Next i
    WarnArrayBody = acc
End Function

' "key":"escaped value" の1組。keyName は本モジュール内のリテラル、valueText は
'   modJsonLite.EscapeJsonStr 済み。HTMLへは JsStringSafe を通って入る。
Private Function JStr(ByVal keyName As String, ByVal valueText As String) As String
    JStr = """" & keyName & """:""" & modJsonLite.EscapeJsonStr(valueText) & """" ' SAFE:html
End Function

' ファイル名(拡張子を除く)。規則の正は 13章§2.8 / modUtilText 側。
'   裁定書9 B4: 末尾は「リスクレポート_<yyyymmdd>」(IsoDateCompact)。日付を
'   落とすと同一案件の再生成が同名になる(13章§2.8のテンプレートの一部)。
Private Function FileNameOf(ByVal company As String, ByVal caseId As String, _
                            ByVal dirText As String) As String
    FileNameOf = modUtilText.BuildFileNameSafe(company, caseId, _
                     EX_TAIL & "_" & modUtilText.IsoDateCompact(Date), dirText, EX_EXT) ' SAFE:html ファイル名(HTMLへは入らない)
End Function

' 裁定書9 B4(13章§2.8): 存在しないパスが見つかるまで _2 _3 と連番を探す。
'   基本名が空・上限まで全て埋まっているときは ""(呼び出し側が失敗として扱う)。
Private Function UniqueOutPath(ByVal dirText As String, ByVal baseName As String) As String
    If LenB(Trim$(baseName)) = 0 Then Exit Function

    Dim candidate As String
    candidate = modUtilPath.JoinPath(dirText, baseName & EX_EXT) ' SAFE:html ファイルパス(HTMLへは入らない)
    If Not OutFileExists(candidate) Then
        UniqueOutPath = candidate
        Exit Function
    End If

    Dim n As Long
    For n = 2 To EX_SERIAL_MAX
        candidate = modUtilPath.JoinPath(dirText, baseName & "_" & CStr(n) & EX_EXT) ' SAFE:html ファイルパス(HTMLへは入らない)
        If Not OutFileExists(candidate) Then
            UniqueOutPath = candidate
            Exit Function
        End If
    Next n
End Function

' 存在検査。判定に失敗したときは True(=その名前を避ける)側へ倒し、既存
'   ファイルを上書きする方向へは倒さない(B4の趣旨=過去の出力を消さない)。
Private Function OutFileExists(ByVal pathText As String) As Boolean
    On Error GoTo Unknown0
    OutFileExists = (LenB(Dir$(pathText)) > 0)
    Exit Function
Unknown0:
    OutFileExists = True
End Function

' 空なら JSON の null、非空なら ExtractJsonBlock を通した本体。
Private Function OrNull(ByVal jsonText As String) As String
    OrNull = "null"
    If LenB(Trim$(jsonText)) = 0 Then Exit Function
    Dim body As String
    body = modJsonLite.ExtractJsonBlock(jsonText)
    If LenB(Trim$(body)) = 0 Then Exit Function
    OrNull = body
End Function

' 出力先の設定値(裁定書27 W9-C2)。html_out_dir を優先し、空なら data_dir。
Private Function OutDirRaw() As String
    Dim t As String
    t = Trim$(modConfig.GetStr("html_out_dir", vbNullString))
    If LenB(t) = 0 Then t = Trim$(modConfig.GetStr("data_dir", EX_OUT_DEFAULT))
    If LenB(t) = 0 Then t = EX_OUT_DEFAULT
    OutDirRaw = t
End Function

' 出力先の解決(16章 E-21・裁定書27 W9-C2 → 裁定書31 裁定1)。候補の並べ方と
'   実在確認は modUtil.ResolveDataDir が唯一持つ(data_dir.txt -> config ->
'   本体と同じフォルダ \データ)。第1候補以外へ落ちたら記録を残す。
Private Function ResolveOutDir(ByVal rawDir As String) As String
    Dim dirText As String
    dirText = modUtil.ResolveDataDir(rawDir)
    If LenB(dirText) = 0 Then Exit Function
    If StrComp(dirText, FirstCandidate(rawDir), vbTextCompare) <> 0 Then
        modLog.LogError "E0501", EX_SRC & ".ResolveOutDir", "out_dir_fallback"
    End If
    ResolveOutDir = dirText
End Function

' 第1候補(config が指した場所そのもの)。取れないときは ""。
Private Function FirstCandidate(ByVal rawDir As String) As String
    On Error GoTo Failed
    Dim listText As String
    ' 裁定書28: 第1候補は data_dir.txt(ランチャーが書いたポインタ)を含む
    ' 並びの先頭。ResolveOutDir が「第1候補以外へ落ちた」を判定する相手なので、
    ' 並べ方は ResolveDataDir と同じ引数で作る(片方だけ古いと誤検知になる)。
    ' 裁定書31 裁定1: 環境変数は渡さない(第3候補は本体と同じフォルダ \データ)。
    listText = modUtil.DataDirCandidates( _
        modUtil.ReadDataDirPointer(modUtil.BookDirHint()), rawDir, _
        modUtil.BookDirHint())
    If LenB(listText) = 0 Then Exit Function
    FirstCandidate = Split(listText, vbLf)(0)
    Exit Function
Failed:
    FirstCandidate = vbNullString
End Function

' 末尾の区切りを落とす。答えを2箇所に持たないため modUtil へ委ねる
'   (裁定書27 W9-B2 で modUtil.TrimTrailingSep を新設した)。
Private Function TrimTrailingSep(ByVal pathText As String) As String
    TrimTrailingSep = modUtil.TrimTrailingSep(pathText)
End Function

' フォルダの用意。裁定書27 W9-B2 で Scripting.FileSystemObject を撤去し、
'   MkDir だけで階層を作る modUtil.EnsureFolder へ寄せた(実装は1箇所)。
Private Function EnsureDir(ByVal dirText As String) As Boolean
    EnsureDir = modUtil.EnsureFolder(dirText)
End Function

' UTF-8(BOMあり)で書き出す(18章§5.3(3)・16章 E-47(3))。VBAの Open/Print # は
'   CP932で書かれ非CP932文字が "?" 化するため使わない。
'   裁定書27 W9-B2: ADODB.Stream を撤去し、**純VBA**(`Open For Binary` +
'   modUtilText.Utf8Bytes の自前エンコード)の modUtil.WriteUtf8File へ寄せた。
'   符号化の正しさ(BOM・3バイト・サロゲート・CP932外)は層(a)で固定してある。
Private Function WriteUtf8Bom(ByVal pathText As String, ByVal bodyText As String) As Boolean
    WriteUtf8Bom = modUtil.WriteUtf8File(pathText, bodyText, True)
End Function
