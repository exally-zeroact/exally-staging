# toru-jitsu-excel-no-idageta.ps1
#   -- ★実Excel 自身が `####` に して いる マスを 数える★（112）（2026-09-24）
#
#  ★★なぜ★★
#    Exally で ★450マス★が `####`。★幅も 字の 大きさも 実Excel と 合って います★。
#    ⇒★では 実Excel では どう 出て いるのか★
#    ⇒★実Excel の `.Text` は 幅が 足りなければ `#####` を 返します★
#    ⇒★両方 `####` なら 直す 物は 在りません★／★Excel だけ 数字なら そこが 欠陥★
#
#  ★★ブラウザを 使いません★★
#    ＝実物を お客さんの 道で 開くと ★2〜5分★ かかり、★5分でも 開かない 事が 在ります★
#    ＝★Excel 側だけで 答えが 出る 問いです★
#
#  ★★出さない 物★★ ... ★マスの 値（数字そのもの）★
#    ＝出すのは ★板・行・列・桁数・列の 幅★ だけ
#
#  ★★読むだけ★★（`Close($false)`）
#
#  使い方: <この道具> -道 <xlsb> [-板数 3] [-何個 20] [-印 <sha256>]

param(
  [string]$道 = '',
  [int]$板数 = 3,
  [int]$何個 = 20,
  [string]$印 = ''
)

$版 = $PSVersionTable.PSVersion
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で★★'; exit 8 }
if (-not $道 -or -not (Test-Path -LiteralPath $道)) { Write-Host '★★材料が 在りません★★'; exit 4 }
$h = (Get-FileHash $道 -Algorithm SHA256).Hash.ToLower()
Write-Host ('★見る★ ' + (Split-Path $道 -Leaf) + ' ／ sha256 ' + $h)
if ($印) {
  if ($h -ne $印.ToLower()) { Write-Host '★★材料が 違います★★'; exit 2 }
  Write-Host '★印は 合って います★'
}
$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個★')
if ($数1 -ne 0) { Write-Host '★★Excel が 動いて います★★'; exit 3 }

$xl = New-Object -ComObject Excel.Application
$bk = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $xl.AutomationSecurity = 3
  $bk = $xl.Workbooks.Open($道, $false, $true)
  $みる = [math]::Min([int]$bk.Sheets.Count, $板数)
  Write-Host ('★板★ ' + [string]$bk.Sheets.Count + '枚（★先頭 ' + $みる + '枚を 見ます★）')
  Write-Host ''
  $全井桁 = 0; $全マス = 0
  $出した = 0
  for ($i = 1; $i -le $みる; $i++) {
    $sh = $bk.Sheets.Item($i)
    $な = [string]$sh.Name
    $used = $sh.UsedRange
    $r0 = [int]$used.Row; $c0 = [int]$used.Column
    $rn = [int]$used.Rows.Count; $cn = [int]$used.Columns.Count
    $井桁 = 0; $マス = 0
    for ($r = 0; $r -lt $rn; $r++) {
      for ($c = 0; $c -lt $cn; $c++) {
        $cell = $sh.Cells.Item($r0 + $r, $c0 + $c)
        $v = $cell.Value2
        if ($null -eq $v) { $cell = $null; continue }
        $マス++
        $t = [string]$cell.Text
        if ($t -match '^#+$') {
          $井桁++
          if ($出した -lt $何個) {
            # ══ ★★「0」を 1つの 窓だけで 取らない★★ ══
            #   ★桁数は `.Value2` から 作って います★＝`.Value2` が 0を 返すと ★桁が 1に なります★
            #   ⇒★2つ目の 窓（`=(そのマス)=0`）と 型を 一緒に 出します★
            #   ★本は 1マスも 触りません★（`Evaluate`）
            $しき = '=(' + "'" + $な + "'!" + [string]$cell.Address($false, $false) + ')=0'
            $零 = '(hakarenai)'
            try { $零 = [string]$xl.Evaluate($しき.Substring(1)) } catch { $零 = '(utenai)' }
            $かた = '(kara)'
            if ($v -is [double]) { $かた = 'Double' }
            elseif ($v -is [string]) { $かた = 'String' }
            elseif ($v -is [bool]) { $かた = 'Bool' }
            $桁 = ([string]$v).Length
            $w = [string]$sh.Columns.Item($c0 + $c).ColumnWidth
            Write-Host ('  ' + $な + "`t" + [string]$cell.Address($false, $false) + "`t幅 " + $w + "`t桁 " + $桁 + "`t井桁 " + $t.Length + "`t零 " + $零 + "`t型 " + $かた)
            $出した++
          }
        }
        $cell = $null
      }
    }
    Write-Host ('★板「' + $な + '」★ 字の 在る マス ' + $マス + '個 ／ ★実Excel が `####` に して いる ' + $井桁 + '個★')
    $全井桁 += $井桁; $全マス += $マス
    $used = $null; $sh = $null
  }
  Write-Host ''
  Write-Host ('★★合計★★ 見た マス ' + $全マス + '個 ／ ★実Excel の `####` ' + $全井桁 + '個★')
  $bk.Close($false); $bk = $null
} finally {
  if ($null -ne $bk) { $bk.Close($false) }
  $bk = $null
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $xl = $null
  $t2 = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t2.Elapsed.TotalSeconds -lt 300)) { Start-Sleep -Milliseconds 250 }
  Write-Host ('★Excel は ' + [math]::Round($t2.Elapsed.TotalSeconds, 1) + '秒で 消えました★')
}
