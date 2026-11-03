# toru-xlsm-wo-excel-ni-hirakaseru.ps1
#   -- ★うちで 組んだ `.xlsm` を 実Excel が 開くか★（101）（2026-09-22）
#
#  ★★なぜ★★
#    Exally1「マクロ入り（`.xlsm`）でも 板を 足したい。★材料は 実Excel が 要る★ので お願いします」
#    ⇒`tests/fixtures/vba-sample.xlsm` が ★既に 在ります★（`scripts/make-vba-fixture.mjs` が 作った）
#    ⇒★但し これは うちで 手で 組んだ 物★です。
#    ⇒★実Excel が 開けるかを 1度も 測って いません★
#    ⇒★開けない 物を 材料に すると 「板を 足したら 壊れた」と 見えます★（★嘘の 赤★）
#    ⇒★だから 材料に する 前に ここを 測ります★
#
#  ★★測る 物★★
#    ①`Open` が 投げないか
#    ②★マクロが 入って いると 見えるか★（`HasVBProject`）
#    ③板の 名／マスの 値
#    ④★`VBProject` の 中まで 読めるか★
#       ⇒★読めない 事が 多い★（「VBA プロジェクト オブジェクト モデルへの アクセスを 信頼する」が 既定で 切）
#       ⇒★読めなくても それは 欠陥では ありません★＝★そう 書きます★
#
#  ★★読むだけ★★（★1バイトも 書きません★）
#
#  ★門★
#    ①貝殻が 5.1（exit 8）／②走らせる 前の Excel が 0個（exit 3）／③材料が 在るか（exit 4）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具> [-道 <xlsm>]

param([string]$道 = '')

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$repo = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
if (-not $道) { $道 = Join-Path $repo ('tests' + [string][char]92 + 'fixtures' + [string][char]92 + 'vba-sample.xlsm') }
if (-not (Test-Path -LiteralPath $道)) { Write-Host ('★★在りません★★ ' + $道); exit 4 }
Write-Host ('★見る★ ' + $道 + ' ／ ' + (Get-Item $道).Length + ' バイト ／ sha256 ' + (Get-FileHash $道 -Algorithm SHA256).Hash.ToLower())

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$xl = New-Object -ComObject Excel.Application
$bk = $null
$投げた = ''
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  # ★マクロは 動かしません★（`AutomationSecurity` 3 ＝ 全部 止める）
  $xl.AutomationSecurity = 3
  try { $bk = $xl.Workbooks.Open($道, $false, $true) }
  catch { $投げた = $_.Exception.Message }

  Write-Host ''
  if ($投げた) {
    Write-Host ('★★①Open が 投げました★★ ' + $投げた)
    Write-Host '  ⇒★この 材料は 実Excel で 開けません★＝★板を 足す 前に 直す 必要が 在ります★'
  } else {
    Write-Host '★①Open は 投げません★'
    Write-Host ('★②マクロが 入って いるか（HasVBProject）★ ... ' + [string]$bk.HasVBProject)
    Write-Host ('★③板★ ... ' + [string]$bk.Sheets.Count + '枚')
    for ($i = 1; $i -le [int]$bk.Sheets.Count; $i++) {
      $sh = $bk.Sheets.Item($i)
      Write-Host ('   ' + $i + '枚目 名 ' + [string]$sh.Name + ' ／ A1 「' + [string]$sh.Range('A1').Text + '」 ／ 図形 ' + [string]$sh.Shapes.Count + '個')
      $sh = $null
    }
    Write-Host ''
    Write-Host '★④VBProject の 中まで 読めるか★（★読めなくても 欠陥では ありません★）'
    try {
      $n = [int]$bk.VBProject.VBComponents.Count
      Write-Host ('   ★読めました★ ... 部品 ' + $n + '個')
      for ($k = 1; $k -le $n; $k++) {
        $c = $bk.VBProject.VBComponents.Item($k)
        Write-Host ('     ' + [string]$c.Name + ' ／ 種類 ' + [string]$c.Type + ' ／ 行 ' + [string]$c.CodeModule.CountOfLines)
        $c = $null
      }
    } catch {
      Write-Host ('   ★読めません★ ... ' + $_.Exception.Message)
      Write-Host '   ⇒★「VBA プロジェクト オブジェクト モデルへの アクセスを 信頼する」が 切って あるから★'
      Write-Host '   ⇒★これは 守りの 設定です＝開けるかどうかとは 別★'
    }
    $bk.Close($false); $bk = $null
  }
} finally {
  if ($null -ne $bk) { $bk.Close($false) }
  $bk = $null
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $xl = $null
  $t = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 300)) { Start-Sleep -Milliseconds 250 }
  $t.Stop()
  $残り = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
  if ($残り -eq 0) { Write-Host ('★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒で ★消えました★★') }
  else { Write-Host ('★★Excel は 消えませんでした★★ ／ 残り ' + $残り + '個') }
}
if ($投げた) { exit 1 }
