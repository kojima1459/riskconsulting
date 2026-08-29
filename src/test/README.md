# src/test - テスト層

純ロジックテスト（modTestsPure1..n）・実行と集計（modTestRunner）・Excel固有E2Eスモーク（modTestsExcel）・本体内mockトランスポート（modMockLlm）。全層を参照してよく、レイヤリング規則（R1）とExcelトークン規則（R4）の適用外（ただし modTestRunner / modTestsPure* は LibreOffice で実行できる必要があるため純ロジックを機械で強制する）。どの層からもテスト層を参照してはならない。
