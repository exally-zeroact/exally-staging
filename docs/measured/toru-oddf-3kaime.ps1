# toru-oddf-3kaime.ps1 — ★ODDFPRICE の 切り分け 3回目★（2026-09-16）
#
#  ★★聞く 事は 1つです★★
#    ★ODDFPRICE の 中で basis 2・3 の「1期の 長さ」は 何か★
#
#  ★見込みは 聞く 前に 書いて あります★
#    … `docs/measured/kansuu46/oddf-kiku-koto-3.md`
#
#  ★★どうして 1つに 絞れたか★★
#    1枠目 ㋐-2 当たり … まるごとの 期間も basis で 割る
#    2枠目 ㋐-1 当たり … 傾き（A の 分母）は 合って いる
#          ㋑-1 当たり … 跳ぶのは 準利払日の ★当日★
#          ㋒-2 当たり … 端数3期・basis 3 は ぴたり
#    ★コードを 読んだ★（★枠を 使わずに★）
#      `比()` の 1期の 長さ … basis 2 は 180（定数）／basis 3 は 182.5（定数）
#      ⇒★定数だから 跳ばない★
#    ★実Excel の COUPDAYS★ … basis 2 … 180 ／ basis 3 … 182.5 ＝★うちと 同じ★
#    ⇒★★ODDFPRICE は COUPDAYS とは 別の 長さを 使って います★★
#
#  ★★決まり★★
#    ・走らせる ★その時に★ Excel を 2つの 道具で 数える
#    ・★新しい 空の ブックだけ★（★司さんの 実物は 開きません★）
#    ・Visible=$false / DisplayAlerts=$false / finally で 必ず Quit
#    ・★消えるまで 待って 秒数を 出す★／★BOM 必須★／★2つ目の 窓＋型★
#    ・★1本ずつ 受け止める★（★1本の 事故で 枠を 丸ごと 落とさない★）
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-oddf-3kaime.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-oddf-3kaime-2026-09-16.tsv'

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel … Get-Process ' + $数1 + '個 ／ Win32_Process ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Error '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$式たち = New-Object System.Collections.Generic.List[object]
$足す = { param($訳, $式) $式たち.Add([pscustomobject]@{ 訳 = $訳; 式 = $式 }) }

# ══ ㋐★端数が ちょうど 1期（NC=1）★で 決済を 1日ずつ ══
#   発行 2009-01-01 ／ 初回 2009-07-01 ⇒ 端数 ちょうど 1期（★跳びが 入りません★）
foreach ($d in @('DATE(2009,1,2)', 'DATE(2009,1,3)', 'DATE(2009,4,1)', 'DATE(2009,6,29)', 'DATE(2009,6,30)')) {
  foreach ($b in @(1, 2, 3)) {
    & $足す "㋐NC1で1日ずつ $d (basis=$b)" `
      "=ODDFPRICE($d,DATE(2013,1,1),DATE(2009,1,1),DATE(2009,7,1),0.06,0.05,100,2,$b)"
  }
}

# ══ ㋑★f を 変える（1回／2回／4回）★ ══
foreach ($b in @(1, 2, 3)) {
  & $足す "㋑f=1 (basis=$b)" `
    "=ODDFPRICE(DATE(2009,3,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2010,1,1),0.06,0.05,100,1,$b)"
  & $足す "㋑f=2 (basis=$b)" `
    "=ODDFPRICE(DATE(2009,3,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2009,7,1),0.06,0.05,100,2,$b)"
  & $足す "㋑f=4 (basis=$b)" `
    "=ODDFPRICE(DATE(2009,3,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2009,4,1),0.06,0.05,100,4,$b)"
}

# ══ ㋒★期の 実日数が 変わる 形★ ══
#   1月始まり … 2009-01-01→07-01 ＝ 181日
#   7月始まり … 2009-07-01→2010-01-01 ＝ 184日
foreach ($b in @(1, 2, 3)) {
  & $足す "㋒1月始まり(181日) (basis=$b)" `
    "=ODDFPRICE(DATE(2009,3,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2009,7,1),0.06,0.05,100,2,$b)"
  & $足す "㋒7月始まり(184日) (basis=$b)" `
    "=ODDFPRICE(DATE(2009,9,1),DATE(2013,1,1),DATE(2009,7,1),DATE(2010,1,1),0.06,0.05,100,2,$b)"
}

# ══ ★★COUPDAYS を ★同じ 組で★ 一緒に 打つ★★ ══
#   ★なぜ★ 別の 果で 測った COUPDAYS（210/210 合った）は
#         basis 2 … 180 ／ basis 3 … 182.5 で ★うちと 同じ★ でした。
#         ★なのに ODDFPRICE は basis 2・3 で 跳びます★
#   ⇒★★同じ 紙の 上で 並べれば ★1枚で 証せます★★
#     ＝★「ODDFPRICE は COUPDAYS と 別の 長さを 使って いる」★
foreach ($b in @(1, 2, 3)) {
  & $足す "COUPDAYS 決済 2009-04-01 満期 2013-01-01 f=2 (basis=$b)" `
    "=COUPDAYS(DATE(2009,4,1),DATE(2013,1,1),2,$b)"
  & $足す "COUPDAYBS 決済 2009-04-01 (basis=$b)" `
    "=COUPDAYBS(DATE(2009,4,1),DATE(2013,1,1),2,$b)"
  & $足す "COUPDAYSNC 決済 2009-04-01 (basis=$b)" `
    "=COUPDAYSNC(DATE(2009,4,1),DATE(2013,1,1),2,$b)"
  & $足す "COUPDAYS f=1 (basis=$b)" `
    "=COUPDAYS(DATE(2009,3,1),DATE(2013,1,1),1,$b)"
  & $足す "COUPDAYS f=4 (basis=$b)" `
    "=COUPDAYS(DATE(2009,3,1),DATE(2013,1,1),4,$b)"
  & $足す "COUPDAYS 1月始まり (basis=$b)" `
    "=COUPDAYS(DATE(2009,3,1),DATE(2013,1,1),2,$b)"
  & $足す "COUPDAYS 7月始まり (basis=$b)" `
    "=COUPDAYS(DATE(2009,9,1),DATE(2013,1,1),2,$b)"
}

# ══ ㋓★対照（★合わなければ そこで 止める★）★ ══
& $足す '㋓対照1(紙に在る・○のはず)' `
  '=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,2)'
& $足す '㋓対照2(紙に在る・○のはず)' `
  '=ODDFPRICE(DATE(2009,3,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2010,1,1),0.06,0.05,100,2,1)'
& $足す '㋓対照3(前の枠と同じ・○のはず)' `
  '=ODDFPRICE(DATE(2009,7,1),DATE(2013,1,1),DATE(2008,7,1),DATE(2010,1,1),0.06,0.05,100,2,3)'

Write-Host ('★聞く 式 … ' + $式たち.Count + '本★')

$xl = New-Object -ComObject Excel.Application
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★ODDFPRICE の 切り分け 3回目★（2026-09-16）')
  $行.Add('#')
  $行.Add('# ★聞く 事は 1つ★ … ODDFPRICE の 中で basis 2・3 の「1期の 長さ」は 何か')
  $行.Add('# ★見込みは 聞く 前に 書いて あります★ … kansuu46/oddf-kiku-koto-3.md')
  $行.Add('#')
  $行.Add('# ★2つ目の 窓★ `=(式)=0` … `.Value2` は 0 で ない 値に 0 を 返す')
  $行.Add('#')
  $行.Add('# ★どの Excel か★ … 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('#')
  $行.Add('# 訳' + "`t" + '式' + "`t" + '答え' + "`t" + '出る字' + "`t" + '=(式)=0' + "`t" + '型')

  $r = 1
  foreach ($x in $式たち) {
    $c = $sh.Range('D' + $r)
    $打てた = $true
    try { $c.Formula = $x.式 } catch {
      $打てた = $false
      $行.Add($x.訳 + "`t" + $x.式 + "`t" + '(★打てません★)' + "`t" +
        ('★Excel が 式を 受け付けません★ ' + $_.Exception.Message) + "`t" + '(★打てません★)' + "`t" + '(★打てません★)')
      $r++
    }
    if (-not $打てた) { continue }
    $v = $c.Value2
    $答 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    $型 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
    $字 = [string]$c.Text
    $w = $sh.Range('E' + $r)
    $ゼロか = '(★窓②が 打てません★)'
    try { $w.Formula = '=(' + $x.式.Substring(1) + ')=0'; $ゼロか = [string]$w.Value2 } catch { }
    $行.Add($x.訳 + "`t" + $x.式 + "`t" + $答 + "`t" + $字 + "`t" + $ゼロか + "`t" + $型)
    $r++
  }

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★書いた … ' + $出 + '（' + $式たち.Count + '行）★')
  $bk.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $t = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 300)) {
    Start-Sleep -Milliseconds 500
  }
  $t.Stop()
  Write-Host ('★Excel が 消えるまで ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒 ／ 残り ' +
    @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count + '個★')
}
