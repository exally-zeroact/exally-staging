# toru-jitsu-excel-no-hashira-haba.ps1
#   -- ★実Excel の 列の 幅・行の 高さ・字の 大きさ★（107）（2026-09-24）
#
#  ★★なぜ★★
#    司さん 2026-09-24
#      「★Excel内で 見切れてない とこが 見切れとる けん いよんやろが★」
#    ＝★実Excel では 切れて いない のに Exally で 切れて いる★
#    ＝★イ（実Excel と 同じ 表示）の 不具合★
#    ⇒★合わせる 相手（実Excel の 数）が 無ければ 直せません★
#    ⇒★まず 実Excel の 幅を 取ります★
#
#  ★★取る 物★★（★1枚目の 板だけでは ありません★＝全部の 板）
#    ・列の 幅（`ColumnWidth`＝字の 数 ／ `Width`＝ポイント）
#    ・行の 高さ（`RowHeight`）
#    ・★既定の 字★（`Application.StandardFont` / `StandardFontSize`）
#    ・★板ごとの 既定の 幅★（`StandardWidth`）
#    ・★字の 大きさが 既定と 違う マスが 在るか★（列の 頭の 行で 見る）
#
#  ★★出さない 物★★ ... ★マスの 値／式の 字★
#    ＝★板の 名は 出します★（★どの 板の どの 列か が 分からないと 直せない★）
#      ⇒★司さんの 実物の 板の 名です★＝★紙に 残す 時は 司さんの 許しの 中★
#
#  ★★読むだけ★★（`SaveAs` を 1回も 呼びません／`Close($false)`）
#
#  ★門★
#    ①貝殻が 5.1（exit 8）／②走らせる 前の Excel が 0個（exit 3）／③材料が 在る（exit 4）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具> -道 <xlsb> [-列数 20]

param(
  [string]$道 = '',
  [int]$列数 = 20,
  [string]$印 = ''
)

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }
if (-not $道) { Write-Host '★★-道 を 渡して ください★★'; exit 4 }
if (-not (Test-Path -LiteralPath $道)) { Write-Host ('★★在りません★★ ' + $道); exit 4 }
$h = (Get-FileHash $道 -Algorithm SHA256).Hash.ToLower()
Write-Host ('★見る★ ' + (Split-Path $道 -Leaf) + ' ／ ' + (Get-Item $道).Length + ' バイト ／ sha256 ' + $h)
if ($印) {
  if ($h -ne $印.ToLower()) { Write-Host ('★★材料が 違います★★ 待ち ' + $印); exit 2 }
  Write-Host '★印は 渡された 物と 合って います★'
} else { Write-Host '★★印を 渡されて いません＝すり替わりを 見て いません★★' }

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$xl = New-Object -ComObject Excel.Application
$bk = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $xl.AutomationSecurity = 3      # ★マクロは 動かしません★
  Write-Host ''
  Write-Host ('★この Excel の 既定の 字★ ' + [string]$xl.StandardFont + ' ／ 大きさ ' + [string]$xl.StandardFontSize)
  $bk = $xl.Workbooks.Open($道, $false, $true)
  Write-Host ('★板★ ' + [string]$bk.Sheets.Count + '枚')
  Write-Host ''
  Write-Host '板	既定幅	列	ColumnWidth	Width(pt)	行1の高さ	字の大きさ	字の名'
  for ($i = 1; $i -le [int]$bk.Sheets.Count; $i++) {
    $sh = $bk.Sheets.Item($i)
    $な = [string]$sh.Name
    $既定 = [string]$sh.StandardWidth
    $used = $sh.UsedRange
    $最終列 = [int]$used.Column + [int]$used.Columns.Count - 1
    $見る = [math]::Min($最終列, $列数)
    for ($c = 1; $c -le $見る; $c++) {
      $col = $sh.Columns.Item($c)
      $cell = $sh.Cells.Item(1, $c)
      Write-Host ($な + "`t" + $既定 + "`t" + $c + "`t" +
        [string]$col.ColumnWidth + "`t" + [string][math]::Round([double]$col.Width, 2) + "`t" +
        [string]$sh.Rows.Item(1).RowHeight + "`t" +
        [string]$cell.Font.Size + "`t" + [string]$cell.Font.Name)
      $col = $null; $cell = $null
    }
    $used = $null; $sh = $null
  }
  $bk.Close($false); $bk = $null
} finally {
  if ($null -ne $bk) { $bk.Close($false) }
  $bk = $null
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $xl = $null
  $t = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 300)) { Start-Sleep -Milliseconds 250 }
  $t.Stop()
  Write-Host ''
  Write-Host ('★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒で 消えました★')
}
