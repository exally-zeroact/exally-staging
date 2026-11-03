# hakaru-tsukanda-mono.ps1 - ★掴んだ 物を 離すと 消えるか★（2026-09-18・9枠目の 測り）
#
#  ★★なぜ★★
#    8枠目で ★Excel が 120秒 待っても 消えませんでした★（残り 1個）。
#    ★決着を 出した 測り（2.71秒）と 8枠目の 間で ★2つ 同時に 変わって います★★
#      ㋐★掴んだ 物★ … 8枠目は `$c = $sh.Range(...)` と `$w = $sh.Range(...)` を 持って いた
#      ㋑★貝殻の 版★ … 決着は powershell.exe 5.1 ／ 8枠目は pwsh 7.6
#
#  ★★2つ 同時に 変わる 紙から どちらが 元かは 出ません★★
#    ⇒★★片方を 止めます★★ ＝ ★2 × 2 の 4通り★を 測ります
#      ⑴5.1 × 掴んだまま      ⑵5.1 × 離す
#      ⑶7.x × 掴んだまま      ⑷7.x × 離す
#    ★この 道具は 1通りだけ 測ります★（★どの 通りかは 口で 渡します★）
#    ⇒★呼ぶ 側が 4回 呼びます★（★貝殻は 呼ぶ 側が 選びます★）
#
#  ★★離す とは★★
#    ★掴んだ まま★ … $sh / $bk / $xl だけ $null（★8枠目と 同じ★）
#    ★離す★ ……… ＋ $c = $null; $w = $null（★これだけの 違い★）
#    ★★ReleaseComObject を 物ごと／GC::Collect は どちらでも 足しません★★
#
#  ★★門★★
#    ①★自分が 立てた Excel だけ 見ます★（PID で）
#       ＝★他の 誰かの Excel を 数えて 嘘の 秒を 出さない★
#    ②★走らせる 前に Excel を 2つの 道具で 数える★（1個でも 居たら 走らせない）
#    ③★60秒で 諦める★（★「待った」と「消えた」を 書き分ける★）
#    ④★新しい 空の ブックだけ★（★司さんの ブックは 開きません★）
#    ⑤★押す 本数は 8枠目と 同じ 数え方★（★軽い 式を 115本★）
#       ＝★掴む 回数を 8枠目に 揃える★（★1本だけでは $c は 1回しか 掴みません★）
#
#  ★★見て いない 事★★
#    ・★答えは 見て いません★（★秒だけ★）
#    ・★1回ずつです★＝★揺れは 分かりません★（★呼ぶ 側が 何度も 呼べば 出ます★）
#
#  使い方:
#    powershell.exe -NoProfile -File docs/measured/hakaru-tsukanda-mono.ps1 -Hanasu
#    pwsh          -NoProfile -File docs/measured/hakaru-tsukanda-mono.ps1
param([switch]$Hanasu)

$ErrorActionPreference = 'Stop'
$版 = $PSVersionTable.PSVersion.ToString()
$型 = if ($Hanasu) { '離す（$c/$w も null）' } else { '掴んだ まま（8枠目と 同じ）' }
Write-Host ''
Write-Host ('★★測る 物 … 貝殻 ' + $版 + ' × ' + $型 + '★★')

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★直前の Excel … Get-Process ' + $数1 + '個 ／ Win32_Process ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$xl = New-Object -ComObject Excel.Application
$わたしのPID = 0
$bk = $null
$sh = $null
$c = $null
$w = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  # ★★自分が 立てた Excel の PID を 取る★★（★他人の Excel を 数えない★）
  $わたしのPID = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
  Write-Host ('★私が 立てた Excel の PID … ' + ($わたしのPID -join ',') + '★')

  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)
  for ($i = 1; $i -le 5; $i++) { $sh.Range('A' + $i).Value2 = $i }

  # ★★掴む 回数を 8枠目に 揃える★★（115本）
  $時計 = [Diagnostics.Stopwatch]::StartNew()
  for ($r = 1; $r -le 115; $r++) {
    $c = $sh.Range('H' + $r)
    $c.Formula = '=SUM(A1:A5)'
    $v = $c.Value2
    $w = $sh.Range('I' + $r)
    $w.Formula = '=(SUM(A1:A5))=0'
    $v2 = $w.Value2
  }
  $時計.Stop()
  Write-Host ('★押すのに かかった 秒 … ' + [math]::Round($時計.Elapsed.TotalSeconds, 2) + '秒★（115本）')
  $bk.Close($false)
} finally {
  if ($Hanasu) {
    # ★★これだけの 違い★★
    $c = $null
    $w = $null
  }
  $sh = $null
  $bk = $null
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $xl = $null
  # ★物ごとの Release と GC::Collect は 足しません★

  $t = [Diagnostics.Stopwatch]::StartNew()
  $のこり = 1
  while ($t.Elapsed.TotalSeconds -lt 60) {
    $のこり = @(Get-Process -Id $わたしのPID -ErrorAction SilentlyContinue).Count
    if ($のこり -eq 0) { break }
    Start-Sleep -Milliseconds 250
  }
  $t.Stop()
  $秒 = [math]::Round($t.Elapsed.TotalSeconds, 2)
  if ($のこり -eq 0) {
    Write-Host ('★★' + $版 + ' × ' + $型 + ' … ' + $秒 + '秒で ★消えました★★★')
  } else {
    Write-Host ('★★' + $版 + ' × ' + $型 + ' … ' + $秒 + '秒 待っても ★消えません★★★ ／ 残り ' + $のこり + '個')
    Write-Host '  ★「待った」だけです＝「消えない」とは 言えません★（★60秒で 諦めました★）'
  }
}
