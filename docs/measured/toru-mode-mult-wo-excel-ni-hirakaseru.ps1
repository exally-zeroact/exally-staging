# toru-mode-mult-wo-excel-ni-hirakaseru.ps1
#   -- ★MODE.MULT（直した 後）を 実Excel に 開かせる★（56）（2026-09-20）
#
#  ★★なぜ★★
#    55 で 実Excel 自身に 聞いた 所:
#      `_xlfn.MODE.MULT(A1:A6)`  cm 在り  ref="H43:H44"（★溢れる★）
#    ⇒うちの `XLFN`（132本）に ★無かった★＝★裸で 書いて いた★
#    ⇒Exally1 が 133本に して commit（`0d1a987`）
#    ⇒★足したら 通るか は 開くまで 言えません★
#      ＝記憶「★入れて 落ちないかで 測るな＝引いて 正しい 答えが 出るかで 測る★」
#
#  ★★2026-09-20 足し ── ★前（裸）も 置かれました★★
#    `exally-kakidashi-mode0.xlsx` ... ★裸★（名簿から 一時的に 外して 作った 物）
#    `exally-kakidashi-mode.xlsx`  ... ★`_xlfn.` 付き★
#    ⇒★これで 3個目も 「前 割れて 後 通る」が 立ちます★
#
#  ★★開く 物★★ `%TEMP%\exally-kakidashi-mode.xlsx`（★1本の 名★）
#    16,471B ／ sha256 7c031eac743c7a4712b5555ad7baddeacead5994ba807ca9d4c7b953939988f7
#    ★お客さんの 道★で 作った 物（画面で 打つ ⇒ 本番の 書き出しの 2行）
#    ★材料★ A1:A6 = 1,2,2,3,3,4（★最頻値が 2つ＝2 と 3★）
#    ★生の 字★
#      C1  `<f t="array" ref="C1:C2">_xlfn.MODE.MULT(A1:A6)</f>`  ★直した 物★
#      E1  `<f>_xlfn.MODE.SNGL(A1:A6)</f>`                        ★対照＝名簿に 元から 在る★
#      G1  `<f t="array" ref="G1:G3">_xlfn.SEQUENCE(3)</f>`       ★対照＝付く★
#      I1  `<f t="array" ref="I1:K1">TRANSPOSE(A1:A3)</f>`        ★対照＝裸が 正しい★
#      A10 `<f>SUM(A1:A6)</f>`                                    ★対照＝溢れない★
#    ・★司さんの 実物の 名は 1文字も 在りません★
#    ・★読むだけ★（`Close($false)`／`SaveAs` は 1文字も 在りません）
#
#  ★★4問★★
#    ⑴C1:C2 が ★2 と 3 に 溢れるか★（★狙い＝通る★）
#    ⑵E1（MODE.SNGL）が 2 か＝★元から 在る 物が 下がって いないか★
#    ⑶G1・I1・A10 が 前の 測りと 同じか
#    ⑷C3 が 空か（★はみ出し★）
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②名が 1文字でも 違えば 走らない（exit 7）
#    ③ファイルが 無ければ 走らない（exit 6）／④走らせる 前の Excel が 0個（exit 3）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-mode-mult-excel-2026-09-20.tsv'

# ★★2026-09-20 足し ── ★前（裸）も 置かれた ので 2本立てに します★★
#   ＝★同じ 道具・同じ 1回で 比べる★（別々に 測ると 「道具が 変わった」が 混ざる）
$二本 = @(
  @{ 札 = 'mae(hadaka)';           名 = 'exally-kakidashi-mode0.xlsx' },
  @{ 札 = 'ato(_xlfn. wo tashita)'; 名 = 'exally-kakidashi-mode.xlsx' }
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

$見るマス = @('C1', 'C2', 'C3', 'D1',
              'E1', 'E2',
              'G1', 'G2', 'G3', 'G4',
              'I1', 'J1', 'K1', 'L1',
              'A10')
$訊くマス = @('C1', 'C2', 'E1', 'G1', 'I1', 'A10')

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null; $c = $null; $w = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★MODE.MULT（直した 後）を 実Excel に 開かせた★（56）（2026-09-20）')
  $行.Add('# ★読むだけ★（保存して いません）／★開けるのは 字で 書いた 2本だけ★')
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★どの 貝殻か★ ... PowerShell ' + $版.ToString())
  $行.Add('# ★生の 字★ C1 `_xlfn.MODE.MULT(A1:A6)`（ref=C1:C2・★直した 物★）')
  $行.Add('# ★対照★   E1 `_xlfn.MODE.SNGL`（名簿に 元から 在る）／G1 `_xlfn.SEQUENCE(3)`／I1 `TRANSPOSE`（裸）／A10 `SUM`')
  $行.Add('# ★実Excel 自身は★ `_xlfn.MODE.MULT` と 書きます（55 で 実測・ref=H43:H44）')
  $行.Add('# ★断り★ ... ★「直した 後」しか 在りません＝前が 割れて いたかは 未測定★')

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
