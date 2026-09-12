Attribute VB_Name = "modSchemas2"
Option Explicit

' ============================================================================
' modSchemas2 - JSONスキーマを返す純関数の2本目(30,000字契約による分割先)
' ----------------------------------------------------------------------------
' modSchemas が 26,000字を超えて Schema-S5 を収められないため、12章§2 の
'   分割規約(1モジュール30,000字)に従って新設した。関数名は変えずに置く
'   (同名の Public Function を2つ以上のモジュールに置かない)。
' 正は15章§5.6 の Schema-S5 フェンス。tools/prompt_diff.py が
'   (a)15章との逐語一致 (b)JSONとしてパースできること
'   (c)14章§3 の strict 要件(全オブジェクトで properties と required が一致し
'      additionalProperties: false が付く)を機械で見る。
'
' R4準拠: Excelトークンを持たない純文字列。CP932準拠。
' ============================================================================

' SchemaS5 - 15章§5.6 Schema-S5。顧客向け提案書(22枚)の文章だけを返させる。
'   件数・スコア・順位は VBA が S2/S3 から数えるため stats は持たない。
Public Function SchemaS5() As String
    Dim s As String
    s = "{" & vbLf
    s = s & "  ""type"": ""object""," & vbLf
    s = s & "  ""properties"": {" & vbLf
    s = s & "    ""title"": {""type"": ""string""}," & vbLf
    s = s & "    ""subtitle"": {""type"": ""string""}," & vbLf
    s = s & "    ""themes"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""name"": {""type"": ""string""}," & vbLf
    s = s & "      ""headline"": {""type"": ""string""}," & vbLf
    s = s & "      ""body"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""name"", ""headline"", ""body""], ""additionalProperties"": false}}," & vbLf
    s = s & "    ""premise"": {""type"": ""string""}," & vbLf
    s = s & "    ""structure"": {""type"": ""array"", ""items"": {""type"": ""string""}}," & vbLf
    s = s & "    ""business"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""areas"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "        ""name"": {""type"": ""string""}," & vbLf
    s = s & "        ""desc"": {""type"": ""string""}" & vbLf
    s = s & "      }, ""required"": [""name"", ""desc""], ""additionalProperties"": false}}," & vbLf
    s = s & "      ""factors"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "        ""name"": {""type"": ""string""}," & vbLf
    s = s & "        ""tag"": {""type"": ""string"", ""enum"": [""公開情報より"", ""当社の想定""]}," & vbLf
    s = s & "        ""desc"": {""type"": ""string""}" & vbLf
    s = s & "      }, ""required"": [""name"", ""tag"", ""desc""], ""additionalProperties"": false}}," & vbLf
    s = s & "      ""facts"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "        ""label"": {""type"": ""string""}," & vbLf
    s = s & "        ""value"": {""type"": ""string""}," & vbLf
    s = s & "        ""note"": {""type"": ""string""}" & vbLf
    s = s & "      }, ""required"": [""label"", ""value"", ""note""], ""additionalProperties"": false}}" & vbLf
    s = s & "    }, ""required"": [""areas"", ""factors"", ""facts""], ""additionalProperties"": false}," & vbLf
    s = s & "    ""categories_note"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""heavy"": {""type"": ""string""}," & vbLf
    s = s & "      ""meaning"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""heavy"", ""meaning""], ""additionalProperties"": false}," & vbLf
    s = s & "    ""headline"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""riskmap"": {""type"": ""string""}," & vbLf
    s = s & "      ""classes"": {""type"": ""string""}," & vbLf
    s = s & "      ""priority"": {""type"": ""string""}," & vbLf
    s = s & "      ""hard"": {""type"": ""string""}," & vbLf
    s = s & "      ""ideas"": {""type"": ""string""}," & vbLf
    s = s & "      ""four"": {""type"": ""string""}," & vbLf
    s = s & "      ""themes"": {""type"": ""string""}," & vbLf
    s = s & "      ""steps"": {""type"": ""string""}," & vbLf
    s = s & "      ""decide"": {""type"": ""string""}," & vbLf
    s = s & "      ""share"": {""type"": ""string""}," & vbLf
    s = s & "      ""appendix"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""riskmap"", ""classes"", ""priority"", ""hard"", ""ideas"", ""four"", ""themes"", ""steps"", ""decide"", ""share"", ""appendix""], ""additionalProperties"": false}," & vbLf
    s = s & "    ""hard_risks"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""risk_no"": {""type"": ""integer""}," & vbLf
    s = s & "      ""horizon"": {""type"": ""string"", ""enum"": [""現在"", ""3年"", ""5年"", ""10年""]}," & vbLf
    s = s & "      ""background"": {""type"": ""string""}," & vbLf
    s = s & "      ""approach"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""risk_no"", ""horizon"", ""background"", ""approach""], ""additionalProperties"": false}}," & vbLf
    s = s & "    ""ideas"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""title"": {""type"": ""string""}," & vbLf
    s = s & "      ""aim"": {""type"": ""string""}," & vbLf
    s = s & "      ""effect"": {""type"": ""integer""}," & vbLf
    s = s & "      ""difficulty"": {""type"": ""string"", ""enum"": [""低"", ""中"", ""高""]}," & vbLf
    s = s & "      ""priority"": {""type"": ""boolean""}" & vbLf
    s = s & "    }, ""required"": [""title"", ""aim"", ""effect"", ""difficulty"", ""priority""], ""additionalProperties"": false}}," & vbLf
    s = s & "    ""four"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""title"": {""type"": ""string""}," & vbLf
    s = s & "      ""aim"": {""type"": ""string""}," & vbLf
    s = s & "      ""mechanism"": {""type"": ""string""}," & vbLf
    s = s & "      ""insurance"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""title"", ""aim"", ""mechanism"", ""insurance""], ""additionalProperties"": false}}," & vbLf
    s = s & "    ""theme_table"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""theme"": {""type"": ""string""}," & vbLf
    s = s & "      ""issues"": {""type"": ""string""}," & vbLf
    s = s & "      ""insurance"": {""type"": ""string""}," & vbLf
    s = s & "      ""kpi"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""theme"", ""issues"", ""insurance"", ""kpi""], ""additionalProperties"": false}}," & vbLf
    s = s & "    ""steps"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""title"": {""type"": ""string""}," & vbLf
    s = s & "      ""who"": {""type"": ""string"", ""enum"": [""当社"", ""当社と貴社""]}," & vbLf
    s = s & "      ""desc"": {""type"": ""string""}," & vbLf
    s = s & "      ""ref"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""title"", ""who"", ""desc"", ""ref""], ""additionalProperties"": false}}," & vbLf
    s = s & "    ""decisions"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""title"": {""type"": ""string""}," & vbLf
    s = s & "      ""options"": {""type"": ""string""}," & vbLf
    s = s & "      ""note"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""title"", ""options"", ""note""], ""additionalProperties"": false}}," & vbLf
    s = s & "    ""share_items"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""text"": {""type"": ""string""}," & vbLf
    s = s & "      ""group"": {""type"": ""string""}," & vbLf
    s = s & "      ""priority"": {""type"": ""boolean""}" & vbLf
    s = s & "    }, ""required"": [""text"", ""group"", ""priority""], ""additionalProperties"": false}}," & vbLf
    s = s & "    ""notes"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""slide_no"": {""type"": ""integer""}," & vbLf
    s = s & "      ""read"": {""type"": ""string""}," & vbLf
    s = s & "      ""ask"": {""type"": ""string""}," & vbLf
    s = s & "      ""probe"": {""type"": ""array"", ""items"": {""type"": ""string""}}," & vbLf
    s = s & "      ""follow"": {""type"": ""array"", ""items"": {""type"": ""string""}}" & vbLf
    s = s & "    }, ""required"": [""slide_no"", ""read"", ""ask"", ""probe"", ""follow""], ""additionalProperties"": false}}" & vbLf
    s = s & "  }," & vbLf
    s = s & "  ""required"": [""title"", ""subtitle"", ""themes"", ""premise"", ""structure"", ""business"", ""categories_note"", ""headline"", ""hard_risks"", ""ideas"", ""four"", ""theme_table"", ""steps"", ""decisions"", ""share_items"", ""notes""]," & vbLf
    s = s & "  ""additionalProperties"": false" & vbLf
    s = s & "}"
    SchemaS5 = s
End Function
