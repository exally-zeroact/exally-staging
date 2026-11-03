# toru-wrap-wo-excel-ni-hirakaseru.ps1
#   -- ★WRAPROWS / WRAPCOLS を 裸で 書き出した 物を 実Excel に 開かせる★（53）（2026-09-20）
#
#  ★★なぜ★★
#    51・52 で 実Excel 自身の 印を 取った 所:
#      実Excel ... `_xlfn.WRAPROWS(A1:A6,3)` ／ `_xlfn.WRAPCOLS(A1:A6,3)`
#      うち ..... `lib/xlsx-io.js` の `XLFN`（130本）に ★無い★ ⇒★裸で 書く★
#    ★そして お客さんは これを 打てます★
#      `lib/shiki-tsunagi.js:170`（台）／`lib/formula-extra-plug.js:260`（皮）
#    ★`LAMBDA` が 要らない ので `NEEDS_XLPM` の 門は 止めません★
#    ⇒★★「打てて 計算できて 書き出せて、相手の Excel で 割れる」形★★
#    ⇒★但し 「#NAME? に なる」は ★見立て★でした＝★ここで 割ります★
#
#  ★★2026-09-20 足し ── ★前と 後を ★同じ 道具・同じ 1回★で 比べます★★
#    ・`exally-kakidashi-wrap.xlsx`  ... ★直す 前★（裸）
#    ・`exally-kakidashi-wrap2.xlsx` ... ★直した 後★（`_xlfn.` 付き）
#    ⇒★引数は 取りません＝この 2本しか 開けません★（★門を 緩めない★）
#    ⇒★記憶「入れて 落ちないかで 測るな＝引いて 正しい 答えが 出るかで 測る」★
#
#  ★★開く 物★★ `%TEMP%\exally-kakidashi-wrap.xlsx` ほか 1本（★字で 書いた 2本だけ★）
#    16,705B ／ sha256 19fad67d915cf5288aca00d6cb4726824ce90f38f706fc2637945fd52606de11
#    ★お客さんの 道★で 作った 物（画面で 打つ ⇒ 本番の 書き出しの 2行）
#    ★生の 字（私が 読みました）★
#      C1  `<f t="array" ref="C1:E2">WRAPROWS(A1:A6,3)</f>`   ★裸★
#      C5  `<f t="array" ref="C5:D7">WRAPCOLS(A1:A6,3)</f>`   ★裸★
#      G1  `<f t="array" ref="G1:G3">_xlfn.SEQUENCE(3)</f>`   ★対照＝付く★
#      I1  `<f t="array" ref="I1:K1">TRANSPOSE(A1:A3)</f>`    ★対照＝裸が 正しい★
#      A10 `<f>SUM(A1:A6)</f>`                                ★対照＝溢れない★
#      `_xlfn.` の 数 ★1個だけ★／`cm="1"` 4個
#    ・★司さんの 実物の 名は 1文字も 在りません★
#    ・★読むだけ★（`Close($false)`／`SaveAs` は 1文字も 在りません）
#
#  ★★4問★★
#    ⑴C1 / C5 は 開いた 瞬間 ★#NAME? か★（`ISERROR` True か）
#    ⑵対照 G1（SEQUENCE）が 1/2/3 に 溢れて いるか＝★下がって いない★
#    ⑶対照 I1（TRANSPOSE・裸）が 1/2/3 に 溢れて いるか
#       ＝★裸でも 通る 名前が 在る★事の 証し
#    ⑷A10（SUM）が 21 か
#
#  ★★⑶が 要る 訳★★
#    ★もし C1 も I1 も 両方 割れたら 「裸だから」では なく 「別の 訳」です★
#    ⇒★裸で 通る 物（TRANSPOSE）が 通って いて、裸で 割れる 物（WRAPROWS）が 割れる★
#      ＝★その 時だけ 「名前ごとに 印が 要る／要らない」が 言えます★
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②名が 1文字でも 違えば 走らない（exit 7）
#    ③ファイルが 無ければ 走らない（exit 6）／④走らせる 前の Excel が 0個（exit 3）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-wrap-excel-2026-09-20.tsv'

$二本 = @(
  @{ 札 = 'mae(hadaka)';  名 = 'exally-kakidashi-wrap.xlsx' },
  @{ 札 = 'ato(_xlfn.)'; 名 = 'exally-kakidashi-wrap2.xlsx' }
)
if ($二本.Count -ne 2) { exit 4 }

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }
foreach ($x in $二本) {
  $p = Join-Path $env:TEMP $x.名
  if (-not (Test-Path $p)) { Write-Host ('★★在りません ... ' + $p + '★★'); exit 6 }
  Write-Host ('★開く 物 ... ' + $p + '（' + (Get-Item $p).Length + ' バイト）★')
}

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$見るマス = @('C1', 'D1', 'E1', 'C2', 'D2', 'E2', 'F1',
              'C5', 'D5', 'C6', 'D6', 'C7', 'D7', 'C8',
              'G1', 'G2', 'G3', 'G4',
              'I1', 'J1', 'K1', 'L1',
              'A10')
$訊くマス = @('C1', 'C5', 'G1', 'I1', 'A10')

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null; $c = $null; $w = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★WRAPROWS / WRAPCOLS を 裸で 書き出した 物を 実Excel に 開かせた★（53）（2026-09-20）')
  $行.Add('# ★読むだけ★（保存して いません）／★開けるのは 字で 書いた 2本だけ★')
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★どの 貝殻か★ ... PowerShell ' + $版.ToString())
  $行.Add('# ★生の 字★ C1 `WRAPROWS(A1:A6,3)`（裸・ref=C1:E2）／C5 `WRAPCOLS(A1:A6,3)`（裸・ref=C5:D7）')
  $行.Add('# ★対照★   G1 `_xlfn.SEQUENCE(3)`（付く）／I1 `TRANSPOSE(A1:A3)`（裸が 正しい）／A10 `SUM(A1:A6)`')
  $行.Add('# ★実Excel 自身は★ `_xlfn.WRAPROWS` と 書きます（51・52 で 実測）')

  foreach ($ほん in $二本) {
  $p = Join-Path $env:TEMP $ほん.名
  $bk = $xl.Workbooks.Open($p, 0, $true)
  $sh = $bk.Sheets.Item(1)
  $行.Add('#')
  $行.Add('# ═══ ★★' + $ほん.札 + '★★ ... ' + $ほん.名 + '（' + (Get-Item $p).Length + ' バイト）═══')

  # ═══ ★★①開いた 瞬間★★ ═══
  $行.Add('#')
  $行.Add('# ★★①開いた 瞬間★★（★まだ 何も 打って いません★）')
  $行.Add('# マス' + "`t" + '値' + "`t" + '出る字' + "`t" + '式' + "`t" + '型' + "`t" + '溢れの一部か')
  foreach ($ma in $見るマス) {
    $c = $sh.Range($ma)
    $v = $c.Value2
    $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    $型 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
    $一部か = '(?)'
    try { $一部か = [string]$c.HasArray } catch { }
    $行.Add($ma + "`t" + $値 + "`t" + [string]$c.Text + "`t" + [string]$c.Formula + "`t" + $型 + "`t" + $一部か)
  }

  # ═══ ★★②誤りか どうかを 直に 訊く★★ ═══
  $行.Add('#')
  $行.Add('# ★★②ISERROR / ISTEXT / ISNUMBER★★（★ここから 計算し直しが 起きます★）')
  $行.Add('# マス' + "`t" + 'ISERROR' + "`t" + 'ISTEXT' + "`t" + 'ISNUMBER' + "`t" + '窓に入っている式(3つ)')
  $窓行 = 40
  foreach ($ma in $訊くマス) {
    # ★★1つずつ 括弧で 囲む★★（`,` が `+` より 強い＝今日 実測で 踏んだ 所）
    $式たち = @(('=ISERROR(' + $ma + ')'), ('=ISTEXT(' + $ma + ')'), ('=ISNUMBER(' + $ma + ')'))
    if ($式たち.Count -ne 3) { Write-Host '★★式が 3本 在りません＝繋がって います★★'; exit 9 }
    for ($k = 0; $k -lt 3; $k++) {
      try { $w = $sh.Cells.Item($窓行, 15 + $k); $w.Formula2 = $式たち[$k] } catch { }
    }
    $窓行 = $窓行 + 1
  }
  try { $xl.CalculateFull() } catch { }
  $窓行 = 40
  foreach ($ma in $訊くマス) {
    $答 = New-Object System.Collections.Generic.List[string]
    $窓式 = New-Object System.Collections.Generic.List[string]
    for ($k = 0; $k -lt 3; $k++) {
      $w = $sh.Cells.Item($窓行, 15 + $k)
      $v2 = $w.Value2
      $x2 = '(kara)'
      if ($null -ne $v2) { $x2 = [string]$v2 }
      $答.Add($x2)
      $窓式.Add([string]$w.Formula)
    }
    $窓行 = $窓行 + 1
    $行.Add($ma + "`t" + $答[0] + "`t" + $答[1] + "`t" + $答[2] + "`t" + ($窓式 -join ' / '))
  }

  # ═══ ★★2つ目の 窓★★ ═══
  #   ★なぜ★ ... `.Value2` の 「0」は ★本物の 0★ とも ★空★ とも ★誤りの 番号★ とも
  #              区別が 付きません（記憶「意味の 無い 数は 一番 見つけにくい」）
  $行.Add('#')
  $行.Add('# ★★2つ目の 窓★★（★上の 読みの 後に 打って います★）')
  $行.Add('# マス' + "`t" + '値' + "`t" + '型' + "`t" + '=(マス)=0')
  $窓行 = 60
  foreach ($ma2 in $見るマス) {
    $c = $sh.Range($ma2)
    $v = $c.Value2
    $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    $型 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
    $ゼロか = '(★窓2が 打てません★)'
    try { $w = $sh.Range('S' + $窓行); $w.Formula2 = '=(' + $ma2 + ')=0'; $ゼロか = [string]$w.Value2 } catch { }
    $窓行 = $窓行 + 1
    $行.Add($ma2 + "`t" + $値 + "`t" + $型 + "`t" + $ゼロか)
  }

  $bk.Close($false)
  $bk = $null
  $sh = $null
  }

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ('★書いた ... ' + $出 + '★')
} finally {
  # ★★掴んだ物 全部 $null★★
  $c = $null; $w = $null; $sh = $null
  if ($null -ne $bk) { $bk.Close($false) }
  $bk = $null
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $xl = $null
  $t = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 180)) {
    Start-Sleep -Milliseconds 250
  }
  $t.Stop()
  $残り = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
  if ($残り -eq 0) { Write-Host ('★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒で ★消えました★★') }
  else { Write-Host ('★★Excel は 消えませんでした★★ ／ 残り ' + $残り + '個') }
}
