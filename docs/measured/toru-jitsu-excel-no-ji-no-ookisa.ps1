# toru-jitsu-excel-no-ji-no-ookisa.ps1
#   -- ★実Excel の マスごとの 字の 大きさ・書体★（109）（2026-09-24）
#
#  ★★なぜ★★
#    司さん「★Excel内で 見切れてない とこが 見切れとる★」
#    ⇒108で 列の 幅を 並べたら ★Exally の 方が 1.33倍 広い★＝★幅 不足では ない★
#    ⇒残るのは ★字の 側★
#    ⇒Exally は 字の 大きさを ★12 が 672個／18 が 96個★ と 持って います
#    ⇒★実Excel でも そうなのかを 1マスずつ 取ります★（★A1 だけ 見て 決めない★）
#
#  ★★出さない 物★★ ... ★マスの 値／式の 字★
#    ＝出すのは ★字の 大きさ・書体・太字・その マスが 空か★ だけ
#
#  ★★読むだけ★★（`Close($false)`／`SaveAs` を 呼びません）
#
#  使い方: <この道具> -道 <xlsb> -板 給料1 -行 24 -列 14 [-印 <sha256>]

param(
  [string]$道 = '',
  [string]$板 = '',
  [int]$行 = 24,
  [int]$列 = 14,
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
} else { Write-Host '★★印を 渡されて いません★★' }

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います★★'; exit 3 }

$xl = New-Object -ComObject Excel.Application
$bk = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $xl.AutomationSecurity = 3
  $bk = $xl.Workbooks.Open($道, $false, $true)
  $sh = $null
  if ($板) { $sh = $bk.Sheets.Item($板) } else { $sh = $bk.Sheets.Item(1) }
  Write-Host ('★板★ ' + [string]$sh.Name)
  Write-Host ''
  $数え = @{}
  $空 = 0
  $字たち = @{}
  for ($r = 1; $r -le $行; $r++) {
    for ($c = 1; $c -le $列; $c++) {
      $cell = $sh.Cells.Item($r, $c)
      $v = $cell.Value2
      # ══ ★★「0」を 1つの 窓だけで 取らない★★ ══（`tests/monosashi-mado.test.mjs`）
      #   `.Value2` は ★0で ない 値に 0 を 返す★事が 在ります。
      #   ここでは ★空か どうか★にしか 使って いませんが、
      #   ★「空」と「0」を 取り違えると 数える マスが ずれます★
      #   ⇒★2つ目の 窓（`=(そのマス)=0`）と 型も 取ります★
      #   ★本は 1マスも 触りません★（`Evaluate` は 書き込みません）
      $しき = '=(' + "'" + [string]$sh.Name + "'!" + [string]$cell.Address($false, $false) + ')=0'
      $零 = '(hakarenai)'
      try { $零 = [string]$xl.Evaluate($しき.Substring(1)) } catch { $零 = '(utenai)' }
      $かた = '(kara)'
      if ($v -is [double]) { $かた = 'Double' }
      elseif ($v -is [string]) { $かた = 'String' }
      elseif ($v -is [bool]) { $かた = 'Bool' }
      if ($null -eq $v) {
        # ★空だと 思った マスが 本当に 空か★＝★零が True で 型が (kara)★なら 空
        if ($零 -eq 'True' -and $かた -ne '(kara)') { Write-Host ('  ★空に 見えるが 0の マス★ ' + [string]$cell.Address($false, $false)) }
        $空++; $cell = $null; continue
      }
      $s = [string]$cell.Font.Size
      $f = [string]$cell.Font.Name
      $数え[$s] = [int]$数え[$s] + 1
      $字たち[$f] = [int]$字たち[$f] + 1
      $cell = $null
    }
  }
  Write-Host ('★見た 所★ ' + $行 + '行 x ' + $列 + '列 ／ ★空★ ' + $空 + '個')
  Write-Host '★★字の 大きさの 出方（実Excel）★★'
  $数え.GetEnumerator() | Sort-Object { [int]$_.Value } -Descending | ForEach-Object {
    Write-Host ('   ' + $_.Key.PadLeft(4) + ' pt ... ' + $_.Value + '個')
  }
  Write-Host '★★書体の 出方（実Excel）★★'
  $字たち.GetEnumerator() | Sort-Object { [int]$_.Value } -Descending | ForEach-Object {
    Write-Host ('   ' + $_.Key + ' ... ' + $_.Value + '個')
  }
  $sh = $null
  $bk.Close($false); $bk = $null
} finally {
  if ($null -ne $bk) { $bk.Close($false) }
  $bk = $null
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $xl = $null
  $t = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 300)) { Start-Sleep -Milliseconds 250 }
  Write-Host ('★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒で 消えました★')
}
