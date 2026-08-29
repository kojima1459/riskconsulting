Attribute VB_Name = "modSchemas"
Option Explicit

' ============================================================================
' modSchemas - JSONスキーマを返す純関数(14章§7 のスキーマ・レジストリ)
' ----------------------------------------------------------------------------
' 本文の正は 15章(§2 / §3 / §4 / §4.5 / §4.6 / §5 / §6)。**本ファイルは docs/spec/15_プロンプトとJSONスキーマ.md
' から機械生成した写しであり、ここを手で書き換えてはならない**。文言を変える
' ときは 15章を先に改訂し、再生成する(17章 T-23。tools/prompt_diff.py が
' 15章とこの戻り値の diff ゼロを受入条件にしている)。
'
' 実装方式(14章§7・15章§10 冒頭):
'   Const は使わない。VBAの Const は1論理行1,023字・行継続25本の制約に当たり、
'   3,700字級のスキーマ本体を1宣言に収められないため、すべて
'   `s = s & "..." & vbLf` 方式の純関数で組み立てて返す。
'   改行は vbLf に統一し、末尾改行は付けない(15章§10.1 の正規化規則)。
'   本文中の二重引用符は VBA の文字列規則どおり "" で二重化してある。
'
' 戻り値は {{...}} プレースホルダを含んだ**テンプレート**である:
'   tools/prompt_diff.py は本モジュールの Public Function を「文字列リテラルと
'   vbLf 等の組込定数の連結」だけで評価するため(制御構文・関数呼び出し・引数
'   参照はいずれも評価不能として差分になる)、プレースホルダの実値埋め込みと
'   ブロック差し込みを本関数の中で行うことはできない。置換責務の所在
'   (本関数の中か modPipeline 側か)は 15章§10.2 の記述と prompt_diff.py の
'   評価器が食い違っており、司令塔の裁定待ちである。14章§6のシグネチャは
'   そのまま保ってあるので、裁定が「本関数の中」になれば引数はここで使える。
'
' R4準拠(12章§2): Worksheets / Range( / Application. / ThisWorkbook / MsgBox /
'   ActiveSheet には一切触れない純文字列モジュール。config やシートも読まない。
' CP932準拠(15章§0 原則7): 本文・注釈ともに CP932 内の文字だけで書く。
' ============================================================================

' --------------------------------------------------------------------------
' SchemaS1 - 15章§2 Schema-S1(step=s1)
' --------------------------------------------------------------------------
Public Function SchemaS1() As String
    Dim s As String
    s = ""
    s = s & "{" & vbLf
    s = s & "  ""type"": ""object""," & vbLf
    s = s & "  ""properties"": {" & vbLf
    s = s & "    ""company_name"": {""type"": ""string""}," & vbLf
    s = s & "    ""business_summary"": {""type"": ""string""}," & vbLf
    s = s & "    ""main_products"": {""type"": ""array"", ""items"": {""type"": ""string""}}," & vbLf
    s = s & "    ""processes"": {""type"": ""array"", ""items"": {""type"": ""string""}}," & vbLf
    s = s & "    ""locations"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""name"": {""type"": ""string""}," & vbLf
    s = s & "      ""type"": {""type"": ""string"", ""enum"": [""工場"", ""本社"", ""店舗"", ""倉庫"", ""その他""]}," & vbLf
    s = s & "      ""address"": {""type"": ""string""}," & vbLf
    s = s & "      ""hazard_note"": {""type"": ""string""}," & vbLf
    s = s & "      ""notes"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""name"", ""type"", ""address"", ""hazard_note"", ""notes""], ""additionalProperties"": false}}," & vbLf
    s = s & "    ""supply_chain"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""key_materials"": {""type"": ""array"", ""items"": {""type"": ""string""}}," & vbLf
    s = s & "      ""notes"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""key_materials"", ""notes""], ""additionalProperties"": false}," & vbLf
    s = s & "    ""customers"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""segments"": {""type"": ""array"", ""items"": {""type"": ""string""}}," & vbLf
    s = s & "      ""channels"": {""type"": ""array"", ""items"": {""type"": ""string""}}" & vbLf
    s = s & "    }, ""required"": [""segments"", ""channels""], ""additionalProperties"": false}," & vbLf
    s = s & "    ""workforce_notes"": {""type"": ""string""}," & vbLf
    s = s & "    ""management_notes"": {""type"": ""string""}," & vbLf
    s = s & "    ""strategy_outlook"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""mvv"": {""type"": ""string""}," & vbLf
    s = s & "      ""aspirations"": {""type"": ""array"", ""items"": {""type"": ""string""}}," & vbLf
    s = s & "      ""market_context"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""mvv"", ""aspirations"", ""market_context""], ""additionalProperties"": false}," & vbLf
    s = s & "    ""current_coverage"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""line_name"": {""type"": ""string""}," & vbLf
    s = s & "      ""coverage_summary"": {""type"": ""string""}," & vbLf
    s = s & "      ""limit_note"": {""type"": ""string""}," & vbLf
    s = s & "      ""special_note"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""line_name"", ""coverage_summary"", ""limit_note"", ""special_note""], ""additionalProperties"": false}}," & vbLf
    s = s & "    ""field_insights"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""note"": {""type"": ""string""}," & vbLf
    s = s & "      ""tag"": {""type"": ""string"", ""enum"": [""risk_clue"", ""relationship"", ""competitor"", ""constraint"", ""opportunity"", ""other""]}" & vbLf
    s = s & "    }, ""required"": [""note"", ""tag""], ""additionalProperties"": false}}," & vbLf
    s = s & "    ""missing_info"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""item"": {""type"": ""string""}," & vbLf
    s = s & "      ""why_needed"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""item"", ""why_needed""], ""additionalProperties"": false}}," & vbLf
    s = s & "    ""input_quality"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""coverage"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "        ""aspect"": {""type"": ""string"", ""enum"": [""profile"", ""business"", ""sites"", ""history"", ""news"", ""hr"", ""finance_risk"", ""sales_memo"", ""sns"", ""competitors"", ""market"", ""finance"", ""insurance_ctx"", ""hazard""]}," & vbLf
    s = s & "        ""status"": {""type"": ""string"", ""enum"": [""ok"", ""partial"", ""missing""]}" & vbLf
    s = s & "      }, ""required"": [""aspect"", ""status""], ""additionalProperties"": false}}," & vbLf
    s = s & "      ""overall"": {""type"": ""string"", ""enum"": [""high"", ""mid"", ""low""]}," & vbLf
    s = s & "      ""advice"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""coverage"", ""overall"", ""advice""], ""additionalProperties"": false}," & vbLf
    s = s & "    ""research_requests"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""purpose"": {""type"": ""string""}," & vbLf
    s = s & "      ""prompt_text"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""purpose"", ""prompt_text""], ""additionalProperties"": false}}" & vbLf
    s = s & "  }," & vbLf
    s = s & "  ""required"": [""company_name"", ""business_summary"", ""main_products"", ""processes"", ""locations""," & vbLf
    s = s & "               ""supply_chain"", ""customers"", ""workforce_notes"", ""management_notes"", ""strategy_outlook""," & vbLf
    s = s & "               ""current_coverage"", ""field_insights"", ""missing_info"", ""input_quality"", ""research_requests""]," & vbLf
    s = s & "  ""additionalProperties"": false" & vbLf
    s = s & "}"
    SchemaS1 = s
End Function

' --------------------------------------------------------------------------
' SchemaS2 - 15章§3 Schema-S2(step=s2。s2rでも再利用)
' --------------------------------------------------------------------------
Public Function SchemaS2() As String
    Dim s As String
    s = ""
    s = s & "{" & vbLf
    s = s & "  ""type"": ""object""," & vbLf
    s = s & "  ""properties"": {" & vbLf
    s = s & "    ""risks"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""risk_no"": {""type"": ""integer""}," & vbLf
    s = s & "      ""category"": {""type"": ""string"", ""enum"": [""strategy_market"", ""supply_chain"", ""manufacturing_quality"", ""sales_customer"", ""facility_bcp"", ""hr_labor"", ""digital_info"", ""legal_regulatory"", ""finance_counterparty"", ""brand_social""]}," & vbLf
    s = s & "      ""risk_name"": {""type"": ""string""}," & vbLf
    s = s & "      ""scenario"": {""type"": ""string""}," & vbLf
    s = s & "      ""status"": {""type"": ""string"", ""enum"": [""proposed"", ""confirmed"", ""rejected"", ""new""]}," & vbLf
    s = s & "      ""frequency"": {""type"": ""string"", ""enum"": [""high"", ""mid"", ""low""]}," & vbLf
    s = s & "      ""impact"": {""type"": ""string"", ""enum"": [""large"", ""mid"", ""small""]}," & vbLf
    s = s & "      ""frequency_score"": {""type"": ""integer"", ""minimum"": 1, ""maximum"": 5}," & vbLf
    s = s & "      ""impact_score"": {""type"": ""integer"", ""minimum"": 1, ""maximum"": 5}," & vbLf
    s = s & "      ""evidence"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "        ""quote"": {""type"": ""string""}," & vbLf
    s = s & "        ""source"": {""type"": ""string"", ""enum"": [""hp"", ""yuho"", ""memo"", ""contract"", ""prev_renewal"", ""knowledge"", ""inference""]}" & vbLf
    s = s & "      }, ""required"": [""quote"", ""source""], ""additionalProperties"": false}," & vbLf
    s = s & "      ""insurability"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "        ""transferability"": {""type"": ""string"", ""enum"": [""cover"", ""partial"", ""hard""]}," & vbLf
    s = s & "        ""line_note"": {""type"": ""string""}," & vbLf
    s = s & "        ""control_note"": {""type"": ""string""}" & vbLf
    s = s & "      }, ""required"": [""transferability"", ""line_note"", ""control_note""], ""additionalProperties"": false}," & vbLf
    s = s & "      ""loss_scale_note"": {""type"": ""string""}," & vbLf
    s = s & "      ""check_points"": {""type"": ""array"", ""items"": {""type"": ""string""}}," & vbLf
    s = s & "      ""preventions"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "        ""measure"": {""type"": ""string""}," & vbLf
    s = s & "        ""related_menu_id"": {""type"": ""string""}" & vbLf
    s = s & "      }, ""required"": [""measure"", ""related_menu_id""], ""additionalProperties"": false}}" & vbLf
    s = s & "    }, ""required"": [""risk_no"", ""category"", ""risk_name"", ""scenario"", ""status"", ""frequency"", ""impact"", ""frequency_score"", ""impact_score"", ""evidence"", ""insurability"", ""loss_scale_note"", ""check_points"", ""preventions""]," & vbLf
    s = s & "       ""additionalProperties"": false}}," & vbLf
    s = s & "    ""gaps"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""gap_no"": {""type"": ""integer""}," & vbLf
    s = s & "      ""gap_type"": {""type"": ""string"", ""enum"": [""uninsured"", ""underinsured"", ""overlap""]}," & vbLf
    s = s & "      ""target"": {""type"": ""string""}," & vbLf
    s = s & "      ""description"": {""type"": ""string""}," & vbLf
    s = s & "      ""risk_evidence"": {""type"": ""string""}," & vbLf
    s = s & "      ""coverage_evidence"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""gap_no"", ""gap_type"", ""target"", ""description"", ""risk_evidence"", ""coverage_evidence""]," & vbLf
    s = s & "       ""additionalProperties"": false}}," & vbLf
    s = s & "    ""emerging_risks"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""risk_name"": {""type"": ""string""}," & vbLf
    s = s & "      ""category"": {""type"": ""string"", ""enum"": [""strategy_market"", ""supply_chain"", ""manufacturing_quality"", ""sales_customer"", ""facility_bcp"", ""hr_labor"", ""digital_info"", ""legal_regulatory"", ""finance_counterparty"", ""brand_social""]}," & vbLf
    s = s & "      ""horizon"": {""type"": ""string"", ""enum"": [""already"", ""near"", ""mid_long""]}," & vbLf
    s = s & "      ""scenario"": {""type"": ""string""}," & vbLf
    s = s & "      ""evidence_quote"": {""type"": ""string""}," & vbLf
    s = s & "      ""evidence_source"": {""type"": ""string"", ""enum"": [""hp"", ""yuho"", ""memo"", ""contract"", ""prev_renewal"", ""knowledge"", ""inference""]}," & vbLf
    s = s & "      ""proposal_hint"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""risk_name"", ""category"", ""horizon"", ""scenario"", ""evidence_quote"", ""evidence_source"", ""proposal_hint""]," & vbLf
    s = s & "       ""additionalProperties"": false}}," & vbLf
    s = s & "    ""open_questions"": {""type"": ""array"", ""items"": {""type"": ""string""}}" & vbLf
    s = s & "  }," & vbLf
    s = s & "  ""required"": [""risks"", ""gaps"", ""emerging_risks"", ""open_questions""]," & vbLf
    s = s & "  ""additionalProperties"": false" & vbLf
    s = s & "}"
    SchemaS2 = s
End Function

' --------------------------------------------------------------------------
' SchemaS3 - 15章§4 Schema-S3(step=s3。s3rでも再利用)
' --------------------------------------------------------------------------
Public Function SchemaS3() As String
    Dim s As String
    s = ""
    s = s & "{" & vbLf
    s = s & "  ""type"": ""object""," & vbLf
    s = s & "  ""properties"": {" & vbLf
    s = s & "    ""stories"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""story_no"": {""type"": ""integer""}," & vbLf
    s = s & "      ""proposal_kind"": {""type"": ""string"", ""enum"": [""upsell"", ""cross_sell"", ""scheme""]}," & vbLf
    s = s & "      ""headline"": {""type"": ""string""}," & vbLf
    s = s & "      ""hook_question"": {""type"": ""string""}," & vbLf
    s = s & "      ""target_risk_nos"": {""type"": ""array"", ""items"": {""type"": ""integer""}}," & vbLf
    s = s & "      ""target_gap_nos"": {""type"": ""array"", ""items"": {""type"": ""integer""}}," & vbLf
    s = s & "      ""menu_ids"": {""type"": ""array"", ""items"": {""type"": ""string""}}," & vbLf
    s = s & "      ""line_ids"": {""type"": ""array"", ""items"": {""type"": ""string""}}," & vbLf
    s = s & "      ""scheme_id"": {""type"": ""string""}," & vbLf
    s = s & "      ""pitch"": {""type"": ""string""}," & vbLf
    s = s & "      ""similar_case_id"": {""type"": ""string""}," & vbLf
    s = s & "      ""expected_objection"": {""type"": ""string""}," & vbLf
    s = s & "      ""objection_response"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""story_no"", ""proposal_kind"", ""headline"", ""hook_question"", ""target_risk_nos"", ""target_gap_nos""," & vbLf
    s = s & "                    ""menu_ids"", ""line_ids"", ""scheme_id"", ""pitch"", ""similar_case_id""," & vbLf
    s = s & "                    ""expected_objection"", ""objection_response""]," & vbLf
    s = s & "       ""additionalProperties"": false}}," & vbLf
    s = s & "    ""unmatched_risks"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""risk_no"": {""type"": ""integer""}," & vbLf
    s = s & "      ""risk_name"": {""type"": ""string""}," & vbLf
    s = s & "      ""why_unmatched"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""risk_no"", ""risk_name"", ""why_unmatched""], ""additionalProperties"": false}}," & vbLf
    s = s & "    ""do_not_propose"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""topic"": {""type"": ""string""}," & vbLf
    s = s & "      ""reason"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""topic"", ""reason""], ""additionalProperties"": false}}" & vbLf
    s = s & "  }," & vbLf
    s = s & "  ""required"": [""stories"", ""unmatched_risks"", ""do_not_propose""]," & vbLf
    s = s & "  ""additionalProperties"": false" & vbLf
    s = s & "}"
    SchemaS3 = s
End Function

' --------------------------------------------------------------------------
' SchemaS4 - 15章§5 Schema-S4(step=s4)
' --------------------------------------------------------------------------
Public Function SchemaS4() As String
    Dim s As String
    s = ""
    s = s & "{" & vbLf
    s = s & "  ""type"": ""object""," & vbLf
    s = s & "  ""properties"": {" & vbLf
    s = s & "    ""file_title"": {""type"": ""string""}," & vbLf
    s = s & "    ""slides"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""slide_no"": {""type"": ""integer""}," & vbLf
    s = s & "      ""title"": {""type"": ""string""}," & vbLf
    s = s & "      ""bullets"": {""type"": ""array"", ""items"": {""type"": ""string""}}," & vbLf
    s = s & "      ""notes"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""slide_no"", ""title"", ""bullets"", ""notes""], ""additionalProperties"": false}}," & vbLf
    s = s & "    ""hearing_questions"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""question"": {""type"": ""string""}," & vbLf
    s = s & "      ""purpose"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""question"", ""purpose""], ""additionalProperties"": false}}" & vbLf
    s = s & "  }," & vbLf
    s = s & "  ""required"": [""file_title"", ""slides"", ""hearing_questions""]," & vbLf
    s = s & "  ""additionalProperties"": false" & vbLf
    s = s & "}"
    SchemaS4 = s
End Function

' --------------------------------------------------------------------------
' SchemaS2C - 15章§4.5 Schema-S2C(step=s2c)
' --------------------------------------------------------------------------
Public Function SchemaS2C() As String
    Dim s As String
    s = ""
    s = s & "{" & vbLf
    s = s & "  ""type"": ""object""," & vbLf
    s = s & "  ""properties"": {" & vbLf
    s = s & "    ""verdict_summary"": {""type"": ""string""}," & vbLf
    s = s & "    ""issues"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""target"": {""type"": ""string""}," & vbLf
    s = s & "      ""issue_type"": {""type"": ""string"", ""enum"": [""missing"", ""generic"", ""weak_evidence"", ""inconsistent"", ""gap_error"", ""insurability_error""]}," & vbLf
    s = s & "      ""detail"": {""type"": ""string""}," & vbLf
    s = s & "      ""suggestion"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""target"", ""issue_type"", ""detail"", ""suggestion""], ""additionalProperties"": false}}," & vbLf
    s = s & "    ""additional_risks"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""risk_name"": {""type"": ""string""}," & vbLf
    s = s & "      ""why"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""risk_name"", ""why""], ""additionalProperties"": false}}" & vbLf
    s = s & "  }," & vbLf
    s = s & "  ""required"": [""verdict_summary"", ""issues"", ""additional_risks""]," & vbLf
    s = s & "  ""additionalProperties"": false" & vbLf
    s = s & "}"
    SchemaS2C = s
End Function

' --------------------------------------------------------------------------
' SchemaS3C - 15章§4.6 Schema-S3C(step=s3c)
' --------------------------------------------------------------------------
Public Function SchemaS3C() As String
    Dim s As String
    s = ""
    s = s & "{" & vbLf
    s = s & "  ""type"": ""object""," & vbLf
    s = s & "  ""properties"": {" & vbLf
    s = s & "    ""executive_reactions"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""story_no"": {""type"": ""integer""}," & vbLf
    s = s & "      ""reaction"": {""type"": ""string""}," & vbLf
    s = s & "      ""lands"": {""type"": ""boolean""}" & vbLf
    s = s & "    }, ""required"": [""story_no"", ""reaction"", ""lands""], ""additionalProperties"": false}}," & vbLf
    s = s & "    ""issues"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""target"": {""type"": ""string""}," & vbLf
    s = s & "      ""issue_type"": {""type"": ""string"", ""enum"": [""wont_land"", ""not_executable"", ""wrong_priority"", ""weak_hook"", ""context_mismatch"", ""uw_concern""]}," & vbLf
    s = s & "      ""detail"": {""type"": ""string""}," & vbLf
    s = s & "      ""suggestion"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""target"", ""issue_type"", ""detail"", ""suggestion""], ""additionalProperties"": false}}" & vbLf
    s = s & "  }," & vbLf
    s = s & "  ""required"": [""executive_reactions"", ""issues""]," & vbLf
    s = s & "  ""additionalProperties"": false" & vbLf
    s = s & "}"
    SchemaS3C = s
End Function

' --------------------------------------------------------------------------
' SchemaPF - 15章§6 Schema-PF(step=pf)
' --------------------------------------------------------------------------
Public Function SchemaPF() As String
    Dim s As String
    s = ""
    s = s & "{" & vbLf
    s = s & "  ""type"": ""object""," & vbLf
    s = s & "  ""properties"": {" & vbLf
    s = s & "    ""summary"": {""type"": ""string""}," & vbLf
    s = s & "    ""principle_checks"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""q_no"": {""type"": ""integer""}," & vbLf
    s = s & "      ""question"": {""type"": ""string""}," & vbLf
    s = s & "      ""answer"": {""type"": ""string""}," & vbLf
    s = s & "      ""ok"": {""type"": ""boolean""}" & vbLf
    s = s & "    }, ""required"": [""q_no"", ""question"", ""answer"", ""ok""], ""additionalProperties"": false}}," & vbLf
    s = s & "    ""grammar_checks"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""key"": {""type"": ""string"", ""enum"": [""a"", ""b"", ""c"", ""d""]}," & vbLf
    s = s & "      ""label"": {""type"": ""string""}," & vbLf
    s = s & "      ""ok"": {""type"": ""boolean""}," & vbLf
    s = s & "      ""note"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""key"", ""label"", ""ok"", ""note""], ""additionalProperties"": false}}," & vbLf
    s = s & "    ""duplicates"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""ref_id"": {""type"": ""string""}," & vbLf
    s = s & "      ""relation"": {""type"": ""string"", ""enum"": [""重複"", ""近接"", ""差分あり""]}," & vbLf
    s = s & "      ""note"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""ref_id"", ""relation"", ""note""], ""additionalProperties"": false}}," & vbLf
    s = s & "    ""predicted_drop_types"": {""type"": ""array"", ""items"": {""type"": ""string""," & vbLf
    s = s & "      ""enum"": [""T1"", ""T2"", ""T3"", ""T4"", ""T5"", ""T6"", ""T7"", ""T8"", ""T9"", ""T10""]}}," & vbLf
    s = s & "    ""rework_suggestions"": {""type"": ""array"", ""items"": {""type"": ""object"", ""properties"": {" & vbLf
    s = s & "      ""approach"": {""type"": ""string"", ""enum"": [""entry_path"", ""benefit_form"", ""underwriter""]}," & vbLf
    s = s & "      ""pattern_id"": {""type"": ""string"", ""enum"": [""P1"",""P2"",""P3"",""P4"",""P5"",""P6"",""P7"",""P8"",""P9"",""P10"",""P11"",""P12"",""P13"",""P14"",""P15"",""""]}," & vbLf
    s = s & "      ""suggestion"": {""type"": ""string""}" & vbLf
    s = s & "    }, ""required"": [""approach"", ""pattern_id"", ""suggestion""], ""additionalProperties"": false}}," & vbLf
    s = s & "    ""survival"": {""type"": ""string"", ""enum"": [""high"", ""mid"", ""low""]}," & vbLf
    s = s & "    ""advice_to_poster"": {""type"": ""string""}" & vbLf
    s = s & "  }," & vbLf
    s = s & "  ""required"": [""summary"", ""principle_checks"", ""grammar_checks"", ""duplicates""," & vbLf
    s = s & "               ""predicted_drop_types"", ""rework_suggestions"", ""survival"", ""advice_to_poster""]," & vbLf
    s = s & "  ""additionalProperties"": false" & vbLf
    s = s & "}"
    SchemaPF = s
End Function
