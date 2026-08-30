# ============================================================================
# run_excel_tests.ps1 - 「リスク提案ナビ」Windows実機 自動スモークテスト
# ----------------------------------------------------------------------------
# 位置づけ(17章§1): テスト3層の **層(b) 実機** を回す唯一の入口。
#   LibreOffice(tools/run_lo_tests.py)の合格は**出荷条件ではない**。
#   出荷条件はここが返すPASSである(姉妹PJはLO緑のまま出荷して34ラウンド手戻り)。
#
# 前提: Windows + デスクトップ版Excel(VBA必須。Web版/Storeアプリ版は不可)
# 使い方(PowerShellをそのまま実行):
#   powershell -ExecutionPolicy Bypass -File wintest\run_excel_tests.ps1
#   powershell -ExecutionPolicy Bypass -File wintest\run_excel_tests.ps1 -Target prod
#   powershell -ExecutionPolicy Bypass -File wintest\run_excel_tests.ps1 `
#       -MockAddinPath "C:\path\リボンちゃん(検証用).xlam"        # T-14b
#   powershell -ExecutionPolicy Bypass -File wintest\run_excel_tests.ps1 `
#       -ExcelLayerEntry ""       # 層(b)を明示スキップ(T-46④は未達扱いになる)
#
# 何をするか:
#   1. レジストリでVBA信頼設定(AccessVBOM/マクロ許可)を現ユーザーに設定
#   2. dist\リスク提案ナビ_dev.xlsm (または prod版) をCOMで開く
#      -> 初回起動の自己インストーラが走り、vba_srcから全モジュールが組み上がる
#   3. wintest\tests_expected.txt を読んで modTestRunner.SetExpectedCount へ渡す
#      (17章§4-1 ランナー要件(2)。0件実行の「全緑」を成立させない)
#   4. modTestRunner.RunAllPureTests を実行し PASS/FAIL/SKIP を取得
#   5. 層(b) modTestsExcel.RunAllExcelTests を既定で実行する(T-47。同じランナー
#      へ積み増すため、本数照合は「純層=tests_expected」+「層(b)>=1本」の2段)
#   6. 結果を wintest\result_*.log に保存。FAILありなら終了コード1
#
# 移植元: PoC「マイ本棚AI」 wintest/run_excel_tests.ps1。
#   COMの解放順序(子から先にClose/Release -> 最後にExcel本体をQuit/Release)と
#   信頼設定の復元はそのまま維持した。RPN向けの変更はブック名・
#   tests_expected の受け渡し・層(b)フックの3点。
# ============================================================================
param(
    [ValidateSet("dev", "prod")] [string]$Target = "dev",
    [string]$MockAddinPath = "",     # 「リボンちゃん(検証用).xlam」のフルパス(任意・T-14b)
    # 層(b)のエントリ。T-47実装済みのため既定で modTestsExcel を実行する
    # (17章 T-46④「modTestsPure と modTestsExcel の両方」)。"" で明示スキップ可。
    [string]$ExcelLayerEntry = "modTestsExcel.RunAllExcelTests"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$bookName = if ($Target -eq "dev") { "リスク提案ナビ_dev.xlsm" } else { "リスク提案ナビ.xlsm" }
$xlsm = Join-Path $root ("dist\" + $bookName)
$log  = Join-Path $PSScriptRoot ("result_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".log")
$expectedFile = Join-Path $PSScriptRoot "tests_expected.txt"

function Log([string]$msg) {
    $line = "[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $msg
    Write-Host $line
    Add-Content -Path $log -Value $line -Encoding UTF8
}

if (-not (Test-Path $xlsm)) {
    throw "ビルド成果物が見つかりません: $xlsm (先に python build\build_rpn.py --$Target)"
}
if (-not (Test-Path $expectedFile)) {
    throw "wintest\tests_expected.txt がありません(17章§4-1: 1行目に10進整数のみ)"
}
$expectedRaw = (Get-Content -Path $expectedFile -TotalCount 1).Trim()
if ($expectedRaw -notmatch '^\d+$') {
    throw "wintest\tests_expected.txt の1行目が10進整数ではありません: '$expectedRaw'"
}
$expected = [int]$expectedRaw
Log "tests_expected = $expected"

# --- 1) VBA信頼設定(現ユーザーのみ。テスト終了後に元の値へ復元する) ---
Log "VBA信頼設定(HKCU)を確認・設定します(終了時に元の値へ復元します)"
$officeVersions = @("16.0")   # Office 2016以降/365は16.0
$originalSecurity = @{}
foreach ($v in $officeVersions) {
    $key = "HKCU:\Software\Microsoft\Office\$v\Excel\Security"
    New-Item -Path $key -Force | Out-Null

    $existing = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
    $originalSecurity[$v] = @{
        AccessVBOM  = if ($existing -and ($existing.PSObject.Properties.Name -contains "AccessVBOM"))  { $existing.AccessVBOM }  else { $null }
        VBAWarnings = if ($existing -and ($existing.PSObject.Properties.Name -contains "VBAWarnings")) { $existing.VBAWarnings } else { $null }
    }

    Set-ItemProperty -Path $key -Name AccessVBOM -Value 1 -Type DWord    # VBAプロジェクトOMを信頼
    Set-ItemProperty -Path $key -Name VBAWarnings -Value 1 -Type DWord   # マクロを有効化
}

function Restore-VbaTrustSettings {
    foreach ($v in $officeVersions) {
        $key = "HKCU:\Software\Microsoft\Office\$v\Excel\Security"
        $orig = $originalSecurity[$v]
        if ($null -ne $orig.AccessVBOM) {
            Set-ItemProperty -Path $key -Name AccessVBOM -Value $orig.AccessVBOM -Type DWord
        } else {
            Remove-ItemProperty -Path $key -Name AccessVBOM -ErrorAction SilentlyContinue
        }
        if ($null -ne $orig.VBAWarnings) {
            Set-ItemProperty -Path $key -Name VBAWarnings -Value $orig.VBAWarnings -Type DWord
        } else {
            Remove-ItemProperty -Path $key -Name VBAWarnings -ErrorAction SilentlyContinue
        }
    }
    Log "VBA信頼設定(HKCU)を元の値へ復元しました"
}

$excel = $null
$wb = $null
$tmpWb = $null
$addin = $null
$exitCode = 1
try {
    # --- 2) Excel起動+(任意)ニセリボンちゃんの有効化 ----------------------
    Log "Excelを起動します"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true          # 目視も可能に。完全無人化するなら $false
    $excel.DisplayAlerts = $false

    if ($MockAddinPath -ne "") {
        if (-not (Test-Path $MockAddinPath)) { throw "アドインが見つかりません: $MockAddinPath" }
        Log "ニセリボンちゃんを有効化: $MockAddinPath"
        $tmpWb = $excel.Workbooks.Add()   # AddIns.Addはブックが1つ開いている必要がある
        $addin = $excel.AddIns.Add($MockAddinPath, $false)
        $addin.Installed = $true
        $tmpWb.Close($false)
    }

    # --- 3) 本体ブックを開く(初回は自己インストーラが走る) -----------------
    Log "開きます: $xlsm"
    $wb = $excel.Workbooks.Open($xlsm)
    Start-Sleep -Seconds 20   # 自己インストール+Boot完了待ち(遅いPCは伸ばす)

    # --- 4) 純ロジックテストを実機VBAで実行(層a) --------------------------
    Log "modTestRunner.SetExpectedCount($expected) を実行します"
    $excel.Run("'" + $wb.Name + "'!modTestRunner.SetExpectedCount", $expected) | Out-Null
    Log "modTestRunner.RunAllPureTests を実行します"
    $excel.Run("'" + $wb.Name + "'!modTestRunner.RunAllPureTests") | Out-Null

    # 純層の実行本数を層(b)実行前に控える。tests_expected は**純層の本数**であり
    # (裁定書10 §3)、層(b)の Check は同じランナーへ積み増すため、本数照合は
    # 「純層ぶん = tests_expected」「層(b)ぶん >= 1」の2段で行う。
    $reportPure = $excel.Run("'" + $wb.Name + "'!modTestRunner.ReportText")
    $executedPure = -1
    if ($reportPure -match 'PASS\s+(\d+)\s*/\s*FAIL\s+(\d+)\s*/\s*SKIP\s+(\d+)') {
        $executedPure = [int]$Matches[1] + [int]$Matches[2]
    }

    # --- 5) Excel固有テスト(層b・T-47)。既定で modTestsExcel を実行 --------
    if ($ExcelLayerEntry -ne "") {
        Log "層(b) $ExcelLayerEntry を実行します"
        $excel.Run("'" + $wb.Name + "'!" + $ExcelLayerEntry) | Out-Null
    } else {
        Log "層(b)スキップ: -ExcelLayerEntry `"`" が指定されたため modTestsExcel は実行していません(T-46④は未達扱い)"
    }

    $fails  = $excel.Run("'" + $wb.Name + "'!modTestRunner.Failures")
    $report = $excel.Run("'" + $wb.Name + "'!modTestRunner.ReportText")

    Log "---- テスト結果 ----"
    $report -split "`n" | ForEach-Object { Log $_ }

    # 出荷条件は「FAIL 0件 かつ SKIP 0件 かつ 純層の実行本数=tests_expected
    # かつ (層(b)実行時は)層(b)が1本以上実行されている」。層(b)の厳密な本数は
    # modTestsExcel 自身が TE_EXPECTED と自己照合し、ズレはFAILとして現れる。
    # SKIPはPASSにもFAILにも現れないため、件数そのものをここでも見張る。
    $skips = 0
    $executed = -1
    if ($report -match 'PASS\s+(\d+)\s*/\s*FAIL\s+(\d+)\s*/\s*SKIP\s+(\d+)') {
        $executed = [int]$Matches[1] + [int]$Matches[2]
        $skips = [int]$Matches[3]
    } else {
        Log "集計行(PASS n / FAIL m / SKIP k)を読み取れませんでした(合格側へ倒しません)"
    }

    $excelCount = 0
    $layerOk = $true
    if ($ExcelLayerEntry -ne "") {
        if (($executed -ge 0) -and ($executedPure -ge 0)) {
            $excelCount = $executed - $executedPure
        }
        $layerOk = ($excelCount -ge 1)   # 層(b)が実際に走ったこと(0本の全緑を防ぐ)
    }

    if (([int]$fails -eq 0) -and ($skips -eq 0) -and ($executedPure -eq $expected) -and $layerOk) {
        Log "実機テスト: 全PASS(FAIL 0 / SKIP 0 / 純層 $executedPure = tests_expected $expected / 層(b) $excelCount 本)"
        $exitCode = 0
    } else {
        Log "実機テスト: NG (FAIL=$fails / SKIP=$skips / 純層=$executedPure / tests_expected=$expected / 層(b)=$excelCount)"
    }

    $wb.Close($false)
}
catch {
    Log ("エラー: " + $_.Exception.Message)
}
finally {
    # 子COMオブジェクトを先にClose/Release -> 最後にExcel本体をQuit/Release。
    # 順序を守らないとExcelプロセスが残留し、次回実行が不安定になる。
    if ($tmpWb) {
        try { $tmpWb.Close($false) } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($tmpWb)
        $tmpWb = $null
    }
    if ($addin) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($addin)
        $addin = $null
    }
    if ($wb) {
        try { $wb.Close($false) } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
        $wb = $null
    }
    if ($excel) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        $excel = $null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Restore-VbaTrustSettings

    Log "ログ: $log"
}
exit $exitCode
