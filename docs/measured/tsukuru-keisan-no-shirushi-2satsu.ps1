# tsukuru-keisan-no-shirushi-2satsu.ps1
#   -- ★「開いたら 全部 計算しろ」の 印だけが 違う 本を 2冊 作る★（118）（2026-09-25）
#
#  ★★なぜ★★
#    Exally1 が `.xlsb` の `xl/workbook.bin` から ★記録 157（26バイト）★を 出しました。
#    2冊で 違ったのは 2か所だけ:
#        1バイト目 ............... 34 / 35
#        21〜24バイト目（4バイトの 数）... ★1 / 16★（ビット0 と ビット4）
#    ★但し 2人とも ★どの ビットが `fullCalcOnLoad` かを 決めて いません★★
#    ＝★番号表を 記憶で 書かない★（記憶「役所の様式のコード表を記憶で書くな」と 同じ）
#    ⇒★同じ 中身で 印だけ 違う 本を 実Excel に 作らせます★
#    ⇒★差が 出た ビットが それ★ と 言い切れます。
#
#  ★★作る 物★★（★2本の 名だけ★）
#    %TEMP%\exally-keisan-shirushi-nashi.xlsb   ... 印 ★無し★
#    %TEMP%\exally-keisan-shirushi-ari.xlsb     ... 印 ★有り★
#    ★中身は 1文字も 同じ★（同じ 手順で 作り、★印を 立てるかどうかだけ★ 変える）
#
#  ★★どうやって 印を 立てるか★★
#    ★`Workbook.ForceFullCalculation` を 使います★
#    ★但し 「これが `fullCalcOnLoad` に なる」とは ★言って いません★★
#      ＝★立てた時と 立てない時で ★実物の バイトが どう 変わるか★ を 見るだけ★
#      ＝★変わらなければ 「この 道では 立たない」と 出します★（★それも 測りです★）
#    ⇒★意味を 決めるのは 出た バイトです★
#
#  ★★門★★
#    ①貝殻が powershell.exe（5.1）（exit 8）
#    ②走らせる 前の Excel が 0個（exit 3）★他の 席の Excel を 閉じません★
#    ③書いて よい 名は 2本だけ（exit 7）
#    ④2冊の 大きさが 同じ ±で 極端に 違わないか（exit 6）＝★中身が 変わって いない 証し★
#    ⑤★掴んだ物 全部 $null＋Excel が 消えるまで 待つ★（記憶「5.1 かつ 全部 $null」）
#
#  ★★1バイトも 本番に 触りません★★（%TEMP% だけ）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$無し = Join-Path $env:TEMP 'exally-keisan-shirushi-nashi.xlsb'
$有り = Join-Path $env:TEMP 'exally-keisan-shirushi-ari.xlsb'
$許す = @('exally-keisan-shirushi-nashi.xlsb', 'exally-keisan-shirushi-ari.xlsb')
foreach ($f in @($無し, $有り)) {
  if ($許す -notcontains (Split-Path $f -Leaf)) { Write-Host '★★書いて よい 名は 2本だけです★★'; exit 7 }
}

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false

  # ══ ★2冊とも 同じ 手順で 作ります★ ══
  #    ★式を 1本 入れます★（式が 無いと 「開いたら 計算」の 印に 意味が 無い為）
  $作る = {
    param($印を立てるか, $出す先)
    $b = $xl.Workbooks.Add()
    $s = $b.Sheets.Item(1)
    $s.Range('A1').Value2 = [double]1
    $s.Range('A2').Value2 = [double]2
    $s.Range('A3').Formula2 = '=SUM(A1:A2)'
    if ($印を立てるか) { $b.ForceFullCalculation = $true }
    if (Test-Path $出す先) { Remove-Item -LiteralPath $出す先 -Force }
    $b.SaveAs($出す先, 50)       # 50 = xlExcel12（.xlsb）
    $立った = $b.ForceFullCalculation
    $b.Close($false)
    return $立った
  }

  Write-Host ''
  Write-Host '★★作って います★★'
  $立0 = & $作る $false $無し
  Write-Host ('  印 無し ... ForceFullCalculation = ' + $立0)
  $立1 = & $作る $true  $有り
  Write-Host ('  印 有り ... ForceFullCalculation = ' + $立1)
  if ($立1 -ne $true) {
    Write-Host '★★印を 立てられませんでした＝この 道では 立ちません★★'
  }
} finally {
  # ★★掴んだ物 全部 $null★★
  $sh = $null
  if ($null -ne $bk) { $bk.Close($false) }
  $bk = $null
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $xl = $null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  $t = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 180)) {
    Start-Sleep -Milliseconds 250
  }
  $t.Stop()
  $残り = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
  if ($残り -eq 0) { Write-Host ('★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒で ★消えました★★') }
  else { Write-Host ('★★Excel は 消えませんでした★★ ／ 残り ' + $残り + '個') }
}

Write-Host ''
Write-Host '★★出来た 物★★'
$ok = $true
foreach ($f in @($無し, $有り)) {
  if (-not (Test-Path -LiteralPath $f)) { Write-Host ('  ★★在りません★★ ' + $f); $ok = $false; continue }
  $x = Get-Item -LiteralPath $f
  $h = (Get-FileHash -LiteralPath $f -Algorithm SHA256).Hash.ToLower()
  Write-Host ('  ' + (Split-Path $f -Leaf).PadRight(38) + $x.Length.ToString().PadLeft(8) + ' バイト ／ sha256 ' + $h)
}
if (-not $ok) { exit 5 }

$a = (Get-Item -LiteralPath $無し).Length
$b2 = (Get-Item -LiteralPath $有り).Length
$差 = [math]::Abs($a - $b2)
Write-Host ''
Write-Host ('★大きさの 差★ ' + $差 + ' バイト')
if ($差 -gt 2048) {
  Write-Host '★★差が 大きすぎます＝中身も 変わって いる 疑い★★'
  exit 6
}
$h1 = (Get-FileHash -LiteralPath $無し -Algorithm SHA256).Hash
$h2 = (Get-FileHash -LiteralPath $有り -Algorithm SHA256).Hash
if ($h1 -eq $h2) {
  Write-Host '★★2冊が 同じ 中身です＝印は 1ビットも 変わって いません★★'
  Write-Host '  ⇒★`ForceFullCalculation` では この 印は 立ちません★（★これも 測りの 結果です★）'
} else {
  Write-Host '★2冊は 違います＝★どこが 違うかは Exally1 の 道具で 読みます★★'
}
Write-Host ''
Write-Host '★★次に やる 事★★'
Write-Host '  node docs/measured/hakaru-hiraita-toki-keisan-no-shirushi.mjs "<この2冊>" --番号 157'
Write-Host '  ⇒★差が 出た ビットが `fullCalcOnLoad`★ と 言い切れます'
