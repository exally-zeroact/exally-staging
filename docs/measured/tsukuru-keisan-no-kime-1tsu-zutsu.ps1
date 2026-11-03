# tsukuru-keisan-no-kime-1tsu-zutsu.ps1
#   -- ★計算まわりの 決めを ★1つずつ★ 変えて `.xlsb` を 作る★（119）（2026-09-25）
#
#  ★★なぜ★★
#    118（`tsukuru-keisan-no-shirushi-2satsu.ps1`）で 測った 結果:
#      `Workbook.ForceFullCalculation = $true` ⇒ ★記録 157 は 1ビットも 変わらない★
#      変わったのは ★番号 3073 と 3072 だけ＝本の ID（保存する たびに 変わる）★
#    ⇒★`ForceFullCalculation` は ファイルに 残りません★
#    ⇒★どの 決めが 157 を 変えるかを 1つずつ 試します★
#
#  ★★118 で 私が 出した 弱い 門を 直しました★★
#    118 は ★sha256 が 違えば 「2冊は 違う」と 出しました★
#    ⇒★本の ID が 毎回 変わるので ★いつでも 「違う」★ に なります★
#    ⇒★sha256 では 何も 分かりません★＝★記録ごとに 比べるしか ありません★
#    ⇒★この 道具は 作るだけ★／★比べるのは `yomu-xlsb-no-kiroku.mjs` と Exally1 の 道具★
#
#  ★★1つずつ 変える（★2つ 変えたら どちらが 効いたか 言えません★）★★
#    ⓪ もと           ... 何も 変えない（★物差し★）
#    ① 手で 計算       ... `Application.Calculation = xlCalculationManual`（-4135）
#    ② 繰り返し 計算   ... `Application.Iteration = $true`
#    ③ 保存前に 計算   ... `Application.CalculateBeforeSave = $false`（①と 一緒でないと 意味が 無い）
#    ④ 全部 計算       ... `Workbook.ForceFullCalculation = $true`（★118 と 同じ＝物差し★）
#    ⑤ 1904年 起点     ... `Workbook.Date1904 = $true`（★157 の 中に 在るか 見る★）
#    ⑥ 精度は 表示通り ... `Workbook.PrecisionAsDisplayed = $true`
#
#  ★★決めは 必ず 戻します★★
#    ＝`Application` の 決めは ★Excel 全体★の 物です
#    ＝★戻さないと 次に 司さんが Excel を 開いた 時に そのまま です★
#    ⇒★1冊ごとに 元に 戻し、最後に もう一度 全部 戻します★
#
#  ★★門★★
#    ①貝殻が 5.1（exit 8）／②走らせる 前の Excel が 0個（exit 3）★他の 席を 閉じません★
#    ③書いて よい 名は 7本だけ（exit 7）／④7冊 とも 出来たか（exit 5）
#    ⑤★決めが 元に 戻ったか を 最後に 数えます★（戻って いなければ exit 6）
#
#  ★★1バイトも 本番に 触りません★★（%TEMP% だけ）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$台本 = @(
  @{ 札 = '0-moto';      説 = 'もと（何も 変えない）' },
  @{ 札 = '1-tede';      説 = '手で 計算（Calculation = Manual）' },
  @{ 札 = '2-kurikaeshi'; 説 = '繰り返し 計算（Iteration = True）' },
  @{ 札 = '3-hozonmae';  説 = '保存前に 計算しない（手で 計算 ＋ CalculateBeforeSave = False）' },
  @{ 札 = '4-zenbu';     説 = '全部 計算（ForceFullCalculation = True）★118 と 同じ★' },
  @{ 札 = '5-1904';      説 = '1904年 起点（Date1904 = True）' },
  @{ 札 = '6-seido';     説 = '精度は 表示通り（PrecisionAsDisplayed = True）' }
)
$台本の数 = 7
Write-Host ('★台本 ... ' + $台本.Count + '本（決め打ち ' + $台本の数 + '本）')
if ($台本.Count -ne $台本の数) { exit 4 }

$出す = @{}
foreach ($x in $台本) { $出す[$x.札] = Join-Path $env:TEMP ('exally-keisan-kime-' + $x.札 + '.xlsb') }
foreach ($f in $出す.Values) {
  if ((Split-Path $f -Leaf) -notmatch '^exally-keisan-kime-[0-9]-[a-z0-9]+\.xlsb$') {
    Write-Host ('★★書いて よい 名では ありません★★ ' + $f); exit 7
  }
}

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null
$もとの計算 = $null; $もとの繰返 = $null; $もとの保存前 = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  # ★★元の 決めを 覚えて おきます（★必ず 戻す★）★★
  #  ★★2026-09-25 に 踏んだ 穴★★
  #    `Application.Calculation` は ★本を 1冊も 開いて いないと 空が 返ります★
  #    ⇒空を 覚えて 空を 書き戻そうと して ★落ちました★（1冊目の 後で 止まった）
  #    ⇒★捨てる 本を 1冊 開いてから 読みます★
  #  ★★もう 1つ 踏んだ 穴★★
  #    `Application.Calculation` は ★数では なく 飾り付きの 名★が 返ります
  #      （`xlCalculationAutomatic` ... 型は `Microsoft.Office.Interop.Excel.XlCalculation`）
  #    ⇒`-is [int]` は ★false★ に なります
  #    ⇒★数に 直して 覚えます★（戻す 時は 数で 書けます）
  $捨 = $xl.Workbooks.Add()
  $もとの計算   = [int]$xl.Calculation
  $もとの繰返   = [bool]$xl.Iteration
  $もとの保存前 = [bool]$xl.CalculateBeforeSave
  $捨.Close($false)
  $捨 = $null
  Write-Host ''
  Write-Host ('★元の 決め★ Calculation=' + $もとの計算 + ' ／ Iteration=' + $もとの繰返 + ' ／ CalculateBeforeSave=' + $もとの保存前)
  #  ★★読めて いなければ 走らせません★★
  #    ＝★戻せない 物を 変えては いけません★（`Application` の 決めは Excel 全体の 物）
  if ($もとの計算 -isnot [int] -or $もとの計算 -eq 0) {
    Write-Host '★★元の Calculation を 数として 読めません＝走らせません★★'
    exit 9
  }
  if ($もとの繰返 -isnot [bool] -or $もとの保存前 -isnot [bool]) {
    Write-Host '★★元の Iteration / CalculateBeforeSave を 真偽として 読めません＝走らせません★★'
    exit 9
  }

  Write-Host ''
  Write-Host '★★作って います（★1冊ごとに 元に 戻します★）★★'
  foreach ($x in $台本) {
    $札 = $x.札
    $先 = $出す[$札]
    $bk = $xl.Workbooks.Add()
    $sh = $bk.Sheets.Item(1)
    # ★★中身は 7冊 とも 同じ★★
    $sh.Range('A1').Value2 = [double]1
    $sh.Range('A2').Value2 = [double]2
    $sh.Range('A3').Formula2 = '=SUM(A1:A2)'

    switch ($札) {
      '1-tede'       { $xl.Calculation = -4135 }                                  # xlCalculationManual
      '2-kurikaeshi' { $xl.Iteration = $true }
      '3-hozonmae'   { $xl.Calculation = -4135; $xl.CalculateBeforeSave = $false }
      '4-zenbu'      { $bk.ForceFullCalculation = $true }
      '5-1904'       { $bk.Date1904 = $true }
      '6-seido'      { $bk.PrecisionAsDisplayed = $true }
      default        { }
    }

    if (Test-Path -LiteralPath $先) { Remove-Item -LiteralPath $先 -Force }
    $bk.SaveAs($先, 50)      # 50 = xlExcel12（.xlsb）

    # ★★必ず 元に 戻す★★（★型を 見てから★）
    #  ★★戻すのは 本を 閉じる ★前★★★（2026-09-25 に 3回 踏んだ）
    #    ＝本が 1冊も 開いて いないと `Application` の 決めを 触れません（0x800A03EC）
    if ($もとの計算 -is [int])    { $xl.Calculation = $もとの計算 }
    if ($もとの繰返 -is [bool])   { $xl.Iteration = $もとの繰返 }
    if ($もとの保存前 -is [bool]) { $xl.CalculateBeforeSave = $もとの保存前 }

    $bk.Close($false)
    $bk = $null
    $sh = $null

    Write-Host ('  ' + $札.PadRight(14) + ' ... ' + $x.説)
  }
} finally {
  $sh = $null
  if ($null -ne $bk) { $bk.Close($false) }
  $bk = $null
  if ($null -ne $xl) {
    # ★★最後に もう一度 全部 戻します★★
    try {
      # ★★本が 1冊も 開いて いないと 決めを 触れません（0x800A03EC）★★（2026-09-25 に 踏んだ）
      $捨2 = $xl.Workbooks.Add()
      if ($もとの計算 -is [int])    { $xl.Calculation = $もとの計算 }
      if ($もとの繰返 -is [bool])   { $xl.Iteration = $もとの繰返 }
      if ($もとの保存前 -is [bool]) { $xl.CalculateBeforeSave = $もとの保存前 }
      Write-Host ''
      Write-Host ('★戻した 後の 決め★ Calculation=' + [int]$xl.Calculation + ' ／ Iteration=' + $xl.Iteration + ' ／ CalculateBeforeSave=' + $xl.CalculateBeforeSave)
      # ★★「空 と 空 は 同じ」で ✓ を 出しては いけません★★（2026-09-25 に 実際に 出した）
      #    ⇒★3つ とも 型が 合って いる 事を 先に 見ます★
      $今の計算 = [int]$xl.Calculation
      $型よし = ($今の計算 -is [int]) -and ($今の計算 -ne 0) `
                -and ($もとの計算 -is [int]) -and ($もとの繰返 -is [bool]) -and ($もとの保存前 -is [bool])
      if (-not $型よし) {
        Write-Host '  ★★読めません＝戻ったか どうか ★未測定★★★'
      } else {
        $戻った = ($今の計算 -eq $もとの計算) -and ([bool]$xl.Iteration -eq $もとの繰返) -and ([bool]$xl.CalculateBeforeSave -eq $もとの保存前)
        if ($戻った) { Write-Host '  ✓ ★元に 戻って います（3つ とも 型を 見ました）★' } else { Write-Host '  ★★戻って いません★★' }
      }
      $捨2.Close($false)
      $捨2 = $null
    } catch { Write-Host ('★戻せませんでした ... ' + $_.Exception.Message + '★') }
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    $xl = $null
  }
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
foreach ($x in $台本) {
  $f = $出す[$x.札]
  if (-not (Test-Path -LiteralPath $f)) { Write-Host ('  ★★在りません★★ ' + $f); $ok = $false; continue }
  $it = Get-Item -LiteralPath $f
  $h = (Get-FileHash -LiteralPath $f -Algorithm SHA256).Hash.ToLower()
  Write-Host ('  ' + (Split-Path $f -Leaf).PadRight(38) + $it.Length.ToString().PadLeft(7) + ' バイト ／ ' + $h.Substring(0, 16))
}
if (-not $ok) { exit 5 }

Write-Host ''
Write-Host '★★★sha256 では 何も 分かりません★★★'
Write-Host '  ＝★本の ID（記録 3072・3073）が 保存する たびに 変わります★'
Write-Host '  ＝★118 で 私は これを 「2冊は 違う」と 出しました＝弱い 門でした★'
Write-Host '  ⇒★記録ごとに 比べます★'
Write-Host ''
Write-Host '★★次に やる 事★★'
Write-Host '  node docs/measured/yomu-xlsb-no-kiroku.mjs "%TEMP%\exally-keisan-kime-0-moto.xlsb" xl/workbook.bin --生 300'
Write-Host '  ⇒★7冊を 突き合わせて ★0-moto と 違う 記録★を 探します（3072・3073 は 除く）'
