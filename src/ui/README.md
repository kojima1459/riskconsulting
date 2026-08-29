# src/ui - 画面層

起動シーケンス（modBoot）と各シートの描画・読取・進捗／ui_lock を持つ最上位層。app / core を参照してよく、app・core から ui を参照してはならない（依存方向 ui → app → core・12章§4）。Excelトークン（Worksheets / Range( / Application. / ThisWorkbook / MsgBox / ActiveSheet）を書いてよいのは原則この層だけ（R4）。
