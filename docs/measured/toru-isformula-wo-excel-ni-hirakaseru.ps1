# toru-isformula-wo-excel-ni-hirakaseru.ps1
#   -- ★ISFORMULA の 前（裸）と 後（_xlfn.）を 実Excel に 開かせる★（63）（2026-09-20）
#
#  ★★なぜ★★
#    61（総当たり 319個）で 実Excel 自身に 聞いた 所 `_xlfn.ISFORMULA(A1:A3)`。
#    うちの 名簿には ★無かった★＝★裸で 書いて いた★＝★5個目の 穴★。
#    ⇒★「#NAME? に なる」は ★見立て★でした★ ⇒★ここで 割ります★
#
#  ★★★もう 1つ 見る 事（★Exally1 が 気付いた 別の 穴★）★★★
#    ★実Excel★ ... `_xlfn.ISFORMULA(A1:A3)` ★cm 在り★ ref="H1337:H1339"＝★3マスに 溢れる★
#    ★うち★ ..... `<c r="C1" t="b">` ★cm 無し★＝★1つの 値（FALSE）だけ★
#    ⇒★これは 「印」の 話では なく ★計算の 話★かも しれません★
#    ⇒★後の C2 C3 が 埋まるかを 見ます★
#       埋まる ⇒★書き出しは 正しい／画面だけが 1マス★
#       空 ....⇒★書き出しも 1マス分しか 運んで いない★
#
#  ★★開く 物（★字で 書いた 2本だけ★）★★
#    前 `%TEMP%\exally-kakidashi-isformula0.xlsx` 16,267B（`ISFORMULA(A1:A3)`＝裸）
#    後 `%TEMP%\exally-kakidashi-isformula.xlsx`  16,273B（`_xlfn.ISFORMULA(A1:A3)`）
#    ★材料★ A1 = 5（数）／A2 = `=1+1`（式）／A3 = abc（字）
#    ★対照（前後とも 同じ）★
#      E1  `_xlfn.FORMULATEXT(A2)` ⇒ `=1+1`（★もう 直した＝通る はず★）
#      G1  `SUMPRODUCT(A1:A1,A1:A1)` ⇒ 25（★裸で 正しい★）
#      I1  `_xlfn.SEQUENCE(3)` ⇒ I1:I3（★付く★）
#      A10 `SUM(A1:A1)` ⇒ 5（★溢れない★）
#    ・★司さんの 実物の 名は 1文字も 在りません★／★読むだけ★
#
#  ★★5問★★
#    ⑴前の C1 が ★#NAME?★ か／★溢れ先 C2 C3 も 空か★
#    ⑵後の C1 が ★FALSE★ か（A1 は 数＝式では ない）
#    ⑶対照 E1 が 前後とも `=1+1` か
#      ＝★ISFORMULA だけ 割れて FORMULATEXT が 通れば 直しが 1つずつ 効いて いる★
#    ⑷G1 25 ／ I1:I3 1/2/3 ／ A10 5 が 前後とも 同じか
#    ⑸★後の C2 C3 が 埋まるか★（上の 「もう 1つ」）
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②2本とも 無ければ 走らない（exit 6）
#    ③走らせる 前の Excel が 0個（exit 3）／④台本の 数 決め打ち（exit 4）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-isformula-excel-2026-09-20.tsv'

# ★★2026-09-20 足し ── ★前（裸）も 置かれた ので 2本立てに します★★
#   ＝★同じ 道具・同じ 1回で 比べる★（別々に 測ると 「道具が 変わった」が 混ざる）
$二本 = @(
  @{ 札 = 'mae(hadaka)';            名 = 'exally-kakidashi-isformula0.xlsx' },
  @{ 札 = 'ato(_xlfn. wo tashita)'; 名 = 'exally-kakidashi-isformula.xlsx' }
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

$見るマス = @('A1', 'A2', 'A3',
              'C1', 'C2', 'C3', 'C4', 'D1',
              'E1', 'E2',
              'G1', 'G2',
              'I1', 'I2', 'I3', 'I4',
              'A10')
$訊くマス = @('C1', 'C2', 'E1', 'G1', 'I1', 'A10')

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null; $c = $null; $w = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★ISFORMULA の 前（裸）と 後（_xlfn.）を 実Excel に 開かせた★（63）（2026-09-20）')
  $行.Add('# ★読むだけ★（保存して いません）／★開けるのは 字で 書いた 2本だけ★')
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★どの 貝殻か★ ... PowerShell ' + $版.ToString())
  $行.Add('# ★生の 字★ C1 `ISFORMULA(A1:A3)`（裸）⇒ `_xlfn.ISFORMULA(A1:A3)`')
  $行.Add('# ★材料★   A1 = 5（数）／A2 = =1+1（式）／A3 = abc（字）')
  $行.Add('# ★対照★   E1 `_xlfn.FORMULATEXT(A2)`＝`=1+1`／G1 `SUMPRODUCT`＝25（裸で 正しい）／I1 `_xlfn.SEQUENCE(3)`／A10 `SUM`＝5')
  $行.Add('# ★実Excel 自身は★ `_xlfn.ISFORMULA` と 書き ★cm 在り・ref=H1337:H1339＝3マスに 溢れます★（61 で 実測）')
  $行.Add('# ★うちの 書き出しは cm 無し＝1つの 値（FALSE）だけ★ ⇒★後の C2 C3 が 埋まるかを 見ます★')

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
