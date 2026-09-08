#Requires -Version 5.1
<#
==============================================================================
 import_navi_modules.ps1 - 2段ビルドの第2段(裁定書34 §0.3・W12-A)
------------------------------------------------------------------------------
 **開発PC専用**。社内の利用者PCでは動かない(PowerShell が使えない・VBAプロ
 ジェクトへのアクセスが許可されていない)。配布 zip には**絶対に入れない**。
 当方の CI でも実行しない(実Excel が要る)。CI は tools/ship_check.py が
 このファイルの存在と必須の文字列だけを見る。

 何をするのか:
   入力 = 第1段の産物 dist\リスク提案ナビ.xlsm(標準モジュールだけ。
          ui_mode=sheet で単体でも動く)
   出力 = dist\final\リスク提案ナビ.xlsm(**名前は変えない**)
   足すもの:
     (1) UserForm frmNaviHtml(.frm + .frx)
     (2) 参照設定 Microsoft Internet Controls(SHDocVw。WithEvents で使う)
     (3) config の行(ui_mode / tests_expected ほか)
     (4) 案件一覧の26列目 archived_at
     (5) case_data!data_key の入力規則(32値)
     (6) run_log!step の入力規則に ch を足す
     (7) ui\ の5本を出力と同じフォルダへ複写
   標準モジュールは第1段の bin に既に全部入っているので**入れ直さない**
   (入れ直すと bin-roundtrip のバイト一致が崩れる)。

 使い方(開発PCの PowerShell):
   powershell -ExecutionPolicy Bypass -File build\win\import_navi_modules.ps1 `
       -WorkbookPath dist\リスク提案ナビ.xlsm
 事前に一度だけ: Excel のトラストセンター >
   「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」をオン。

 終わったら当方へ返すもの:
   dist\final\リスク提案ナビ.xlsm と、
   python3 tools\ship_check.py --final の出力。
==============================================================================
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [string]$OutputPath,
    [ValidateSet('html','sheet')][string]$UiMode='html'
)
$ErrorActionPreference='Stop'

$repoRoot=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$bookName='リスク提案ナビ.xlsm'

$sourcePath=(Resolve-Path -LiteralPath $WorkbookPath).ProviderPath
if([IO.Path]::GetExtension($sourcePath) -ne '.xlsm'){throw '入力は第1段の .xlsm です。'}
if(-not $OutputPath){$OutputPath=Join-Path $repoRoot ('dist\final\'+$bookName)}
$outputFull=[IO.Path]::GetFullPath($OutputPath)
if([IO.Path]::GetFileName($outputFull) -ne $bookName){throw ('出力の名前は '+$bookName+' のままにしてください。')}
if([StringComparer]::OrdinalIgnoreCase.Equals($sourcePath,$outputFull)){throw '入力と出力は別のファイルにしてください。'}

# --- 入れるもの: UserForm 1本だけ(標準モジュールは第1段の bin にある) -------
$formPath=Join-Path $repoRoot 'src\ui\navi\frmNaviHtml.frm'
$frxPath=[IO.Path]::ChangeExtension($formPath,'.frx')
if(-not (Test-Path -LiteralPath $formPath)){throw 'src\ui\navi\frmNaviHtml.frm がありません。'}
if(-not (Test-Path -LiteralPath $frxPath)){throw 'src\ui\navi\frmNaviHtml.frx がありません(.frm と対で要ります)。'}

# --- 当方の規約の再確認(12章§2。30,000字 / 1物理行 1,000 CP932バイト) ------
# リポジトリの .frm は UTF-8/LF が正。ここで CP932 へ落として字数と行長を見る。
$cp932=[Text.Encoding]::GetEncoding(932,[Text.EncoderExceptionFallback]::new(),[Text.DecoderExceptionFallback]::new())
$formText=[IO.File]::ReadAllText($formPath,[Text.Encoding]::UTF8)
if($formText.Length -gt 30000){throw 'frmNaviHtml が 30,000字を超えています。'}
foreach($line in ($formText -split "`r?`n")){
    try{$n=$cp932.GetByteCount($line)}catch{throw ('CP932 で表せない文字があります: '+$line)}
    if($n -gt 1000){throw ('1物理行が 1,000 CP932バイトを超えています: '+$line.Substring(0,[Math]::Min(40,$line.Length)))}
}
# VBE の Import は CP932/CRLF を期待するので、作業用に写しを作って渡す。
$stageDir=Join-Path $env:TEMP ('RpnStage2_'+[Guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($stageDir)|Out-Null
$stageForm=Join-Path $stageDir 'frmNaviHtml.frm'
[IO.File]::WriteAllBytes($stageForm,$cp932.GetBytes(($formText -replace "`r?`n","`r`n")))
Copy-Item -LiteralPath $frxPath -Destination (Join-Path $stageDir 'frmNaviHtml.frx')

# --- HTML画面の資産(5本) --------------------------------------------------
$uiNames=@('index.html','style.css','markdown.js','views.js','app.js')
foreach($name in $uiNames){
    if(-not (Test-Path -LiteralPath (Join-Path $repoRoot ('ui\'+$name)))){throw ('ui\'+$name+' がありません。')}
}

# --- 期待本数は wintest\tests_expected.txt の prod= から読む(直書き禁止) ----
$expectedText=$null
foreach($line in (Get-Content -LiteralPath (Join-Path $repoRoot 'wintest\tests_expected.txt'))){
    if($line -match '^\s*prod\s*=\s*(\d+)\s*$'){$expectedText=$Matches[1]}
}
if(-not $expectedText){throw 'wintest\tests_expected.txt の prod= を読めません。'}

# --- 13章§2.2 data_key の32値(値源は src\ui\modBootNavi.bas の BN_DATA_KEYS) --
$dataKeysText=$null
$bootNavi=[IO.File]::ReadAllText((Join-Path $repoRoot 'src\ui\modBootNavi.bas'),[Text.Encoding]::UTF8)
$m=[Regex]::Match($bootNavi,'(?s)BN_DATA_KEYS\s+As\s+String\s*=\s*(.+?)\r?\n\r?\n')
if($m.Success){
    $joined=($m.Groups[1].Value -replace '"\s*&\s*_\s*\r?\n\s*"','')
    $dataKeysText=([Regex]::Match($joined,'"(.*)"')).Groups[1].Value
}
if(-not $dataKeysText){throw 'modBootNavi.bas の BN_DATA_KEYS を読めません。'}

$inputHash=(Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
$outDir=Split-Path -Parent $outputFull
[IO.Directory]::CreateDirectory($outDir)|Out-Null
$workingPath=Join-Path $stageDir $bookName
Copy-Item -LiteralPath $sourcePath -Destination $workingPath

$excel=$null;$book=$null;$project=$null
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false
    $excel.DisplayAlerts=$false
    $excel.EnableEvents=$false
    $excel.AutomationSecurity=3
    $book=$excel.Workbooks.Open($workingPath,0,$false)
    if($book.ReadOnly){throw '写しが読み取り専用で開きました。'}
    try {$project=$book.VBProject; $null=$project.VBComponents.Count}
    catch {throw 'VBE のプロジェクトへ触れません。開発PCのトラストセンターで「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」をオンにしてから再実行してください。'}
    if($project.Protection -ne 0){throw 'VBA プロジェクトがロックされています。'}

    # 第1段の産物であることの確認(取り違え防止)。
    foreach($required in @('modBoot','modBootNavi','modNaviHost','modNaviActions','modGatewayRPN2')){
        try{$null=$project.VBComponents.Item($required)}catch{throw ('第1段の産物ではありません('+$required+' が居ません)。')}
    }

    # (2) 参照設定 Microsoft Internet Controls(SHDocVw)。
    #     MSForms はフォームを1本でも入れると Excel が自動で足す。
    $internetGuid='{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}'
    $internetRef=$null
    foreach($reference in $project.References){if($reference.Guid -eq $internetGuid){$internetRef=$reference}}
    if(-not $internetRef){$null=$project.References.AddFromGuid($internetGuid,1,1)}

    # (1) UserForm の取り込み。
    $existing=$null
    try{$existing=$project.VBComponents.Item('frmNaviHtml')}catch{}
    if($existing){$project.VBComponents.Remove($existing)}
    $imported=$project.VBComponents.Import($stageForm)
    if($imported.Name -ne 'frmNaviHtml'){throw ('取り込みで名前が変わりました: '+$imported.Name)}
    foreach($reference in $project.References){if($reference.IsBroken){throw ('壊れた参照設定: '+$reference.Name)}}

    # (3) config の行。ui_mode と tests_expected は必ず上書き、他は空のときだけ。
    $cfg=$book.Worksheets.Item('config')
    $settings=[ordered]@{
        ui_mode=$UiMode; ui_font_scale='medium'; chat_max_turns='12';
        chat_include_materials='FALSE'; ch_effort='medium'; ch_verbosity='low';
        app_display_name='リスク提案ナビ'; tests_expected=$expectedText
    }
    foreach($key in $settings.Keys){
        $last=[int]$cfg.Cells($cfg.Rows.Count,1).End(-4162).Row
        $row=0
        for($n=2;$n -le $last;$n++){if([string]$cfg.Cells($n,1).Value2 -eq $key){$row=$n;break}}
        if($row -eq 0){$row=$last+1;$cfg.Cells($row,1).NumberFormat='@';$cfg.Cells($row,1).Value2=$key}
        if($key -in @('ui_mode','tests_expected') -or [string]::IsNullOrWhiteSpace([string]$cfg.Cells($row,2).Value2)){
            $cfg.Cells($row,2).NumberFormat='@';$cfg.Cells($row,2).Value2=[string]$settings[$key]
        }
    }

    # (4) 案件一覧の26列目 archived_at(13章§2.1。既存25列は動かさない)。
    $cases=$book.Worksheets.Item('案件一覧')
    $found=$false
    for($n=1;$n -le 32;$n++){if([string]$cases.Cells(1,$n).Value2 -eq 'archived_at'){$found=$true}}
    if(-not $found){
        if(-not [string]::IsNullOrEmpty([string]$cases.Cells(1,26).Value2)){throw '案件一覧 の26列目が埋まっています。列の並びを人が確かめてください。'}
        $cases.Cells(1,26).NumberFormat='@';$cases.Cells(1,26).Value2='archived_at'
    }

    # (5) case_data!data_key の入力規則(32値)。値源は modBootNavi の内蔵定数。
    $caseData=$book.Worksheets.Item('case_data')
    $keyCol=0
    for($n=1;$n -le 20;$n++){if([string]$caseData.Cells(1,$n).Value2 -eq 'data_key'){$keyCol=$n;break}}
    if($keyCol -eq 0){throw 'case_data に data_key 列がありません。'}
    $lastData=[Math]::Max(51,[int]$caseData.Cells($caseData.Rows.Count,1).End(-4162).Row)
    $keyRange=$caseData.Range($caseData.Cells(2,$keyCol),$caseData.Cells($lastData,$keyCol))
    $keyRange.Validation.Delete()
    $keyRange.Validation.Add(3,1,1,($dataKeysText -replace ';',','))
    $keyRange.Validation.IgnoreBlank=$true
    $keyRange.Validation.ShowError=$true

    # (6) run_log!step の入力規則に ch(案件チャット)を足す。19章§3。
    $run=$book.Worksheets.Item('run_log')
    $stepCol=0
    for($n=1;$n -le 20;$n++){if([string]$run.Cells(1,$n).Value2 -eq 'step'){$stepCol=$n;break}}
    if($stepCol -eq 0){throw 'run_log に step 列がありません。'}
    $lastRun=[Math]::Max(51,[int]$run.Cells($run.Rows.Count,1).End(-4162).Row)
    $stepRange=$run.Range($run.Cells(2,$stepCol),$run.Cells($lastRun,$stepCol))
    $stepRange.Validation.Delete()
    $stepRange.Validation.Add(3,1,1,'s1,s2,s3,s4,pf,sp,wt,fg,s2c,s3c,s2r,s3r,ch')
    $stepRange.Validation.IgnoreBlank=$true
    $stepRange.Validation.ShowError=$true

    $book.Save()
} catch {
    Write-Warning ('作業中の写し: '+$workingPath)
    throw
} finally {
    if($book){$book.Close($false)}
    if($excel){$excel.Quit();[Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)|Out-Null}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}

if((Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash -ne $inputHash){throw '入力ブックのハッシュが変わりました(入力を書き換えていないか確認してください)。'}
Copy-Item -LiteralPath $workingPath -Destination $outputFull -Force

# (7) ui\ の5本を出力と同じフォルダへ。
$uiTarget=Join-Path $outDir 'ui'
[IO.Directory]::CreateDirectory($uiTarget)|Out-Null
foreach($name in $uiNames){
    Copy-Item -LiteralPath (Join-Path $repoRoot ('ui\'+$name)) -Destination (Join-Path $uiTarget $name) -Force
}

Write-Output ('作成: '+$outputFull)
Write-Output ('入力の SHA256 は変わっていません: '+$inputHash)
Write-Output '次にやること: 出力を開いて VBE の [デバッグ]>[VBAProject のコンパイル] を通し、'
Write-Output '  python3 tools\ship_check.py --final を走らせて、その出力と一緒に返してください。'
