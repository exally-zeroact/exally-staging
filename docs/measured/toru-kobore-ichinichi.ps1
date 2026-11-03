# toru-kobore-ichinichi.ps1 -- ★溢れの 一生を 実Excel に 聞く★（㊵）（2026-09-20）
#
#  ★なぜ★ ... Exally1 が 数えました
#             `book.html` が 借り物を 呼ぶ 所 ★18か所★
#             ＝式に 答える 以外に ★溢れの 広さ★／★そのマスが 溢れの 一部か★ に 使って いる
#             ⇒★★外す 前に 「実Excel が どう するか」を 持って いないと 測れません★★
#
#  ★台本★ ... 毎回 D1 に `=SEQUENCE(3)` を 打つ（★D1:D3 が 1/2/3 に なる★）
#           ⇒その後 1つ 手を 加えて ★D1:D5 を 値と 式で 読む★
#    ㋐何も しない（★対照★）
#    ㋑★頭（D1）を 消す★
#    ㋒★溢れ先（D2）に 字を 打つ★（★#SPILL! に 変わるか★）
#    ㋓★溢れ先（D2）を 消す★
#    ㋔★頭（D1）を 別の 式に 書き換える★（`=SEQUENCE(2)` ⇒ ★D3 は 消えるか★）
#
#  ★おまけ★ ... 最後に `.xlsx` で 保存する（★中の 形は 別の 道具で 読みます★）
#              ＝★開き直しては いません★（`Workbooks.Open` は 1文字も 使いません）
#
#  ★門★
#    ①式に 外へ 出る 6個が 0件（exit 5）／②式に ASCII の 外の 字が 0件（exit 6）
#    ③走らせる 前の Excel が 0個（exit 3）／④台本の 数 決め打ち（exit 4）
#    ⑤貝殻が powershell.exe（5.1）（exit 8）
#    ⑥司さんの 実物を 開く 字は ★1文字も 在りません★
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-kobore-ichinichi-2026-09-20.tsv'
$保存先 = Join-Path $env:TEMP 'kobore-ichinichi.xlsx'

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$台本 = @(
  @{ 番 = '1'; 名 = 'nani mo shinai (taishou)'; 手 = 'none' },
  @{ 番 = '2'; 名 = 'atama D1 wo kesu';         手 = 'clearD1' },
  @{ 番 = '3'; 名 = 'koboresaki D2 ni ji wo utsu'; 手 = 'writeD2' },
  @{ 番 = '4'; 名 = 'koboresaki D2 wo kesu';    手 = 'clearD2' },
  @{ 番 = '5'; 名 = 'atama wo betsu no shiki ni'; 手 = 'rewriteD1' }
)
$台本の数 = 5
Write-Host ('★台本 ... ' + $台本.Count + '本★（決め打ち ' + $台本の数 + '本）')
if ($台本.Count -ne $台本の数) { exit 4 }

$使う式 = '=SEQUENCE(3)', '=SEQUENCE(2)'
$外へ出る = 'WEBSERVICE', 'STOCKHISTORY', 'TRANSLATE', 'DETECTLANGUAGE', 'IMAGE', 'RTD'
$見つけた = 0
foreach ($f in $使う式) {
  foreach ($n in $外へ出る) { if ($f.ToUpper().Contains($n)) { $見つけた++ } }
  foreach ($ch in $f.ToCharArray()) { if ([int]$ch -gt 127) { $見つけた++ } }
}
Write-Host ('★外へ 出る 6個 ＋ ASCII の 外の 字 ... ' + $見つけた + '件★')
if ($見つけた -ne 0) { exit 5 }

$xl = New-Object -ComObject Excel.Application
$bk = $null
$sh = $null
$c = $null
$w = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)
  $sh.Range('A1').Value2 = 3
  $sh.Range('A2').Value2 = 1
  $sh.Range('A3').Value2 = 2

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★溢れの 一生を 実Excel に 聞いた★（㊵）（2026-09-20）')
  $行.Add('# ★台本★ 毎回 D1 に =SEQUENCE(3) を 打つ（D1:D3 = 1/2/3）⇒ その後 1つ 手を 加える')
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★どの 貝殻か★ ... PowerShell ' + $版.ToString())
  $行.Add('# 番' + "`t" + '手' + "`t" + 'マス' + "`t" + '値' + "`t" + '出る字' + "`t" + '式' + "`t" + '溢れの一部か' + "`t" + '=(D1)=0' + "`t" + '型')

  $押し時計 = [Diagnostics.Stopwatch]::StartNew()
  foreach ($x in $台本) {
    # ★毎回 まっさらに してから 打ち直す★
    $sh.Range('D1:F10').ClearContents() | Out-Null
    $sh.Range('D1').Formula2 = '=SEQUENCE(3)'
    switch ($x.手) {
      'clearD1'   { $sh.Range('D1').ClearContents() | Out-Null }
      'writeD2'   { $sh.Range('D2').Value2 = 'jama' }
      'clearD2'   { $sh.Range('D2').ClearContents() | Out-Null }
      'rewriteD1' { $sh.Range('D1').Formula2 = '=SEQUENCE(2)' }
    }
    # ★2つ目の 窓★（★0 が 見せかけに ならないか★）
    $ゼロか = '(★窓2が 打てません★)'
    # ★他の 道具と ★同じ 書き方★に します★（`'=(' + ... + ')=0'`）
    #   ＝門は この 形を 探します／★中身は 同じ★
    $見る = 'D1'
    try { $w = $sh.Range('J1'); $w.Formula2 = '=(' + $見る + ')=0'; $ゼロか = [string]$w.Value2 } catch { }
    for ($r = 1; $r -le 5; $r++) {
      $c = $sh.Cells.Item($r, 4)
      $v = $c.Value2
      $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
      # ★★型も 見ます★★（`.Value2` の 0 が 見せかけか を 判じる 材料）
      $型 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
      $字 = [string]$c.Text
      $式 = [string]$c.Formula
      $一部か = '(?)'
      try { $一部か = [string]$c.HasArray } catch { }
      $行.Add($x.番 + "`t" + $x.名 + "`t" + ('D' + $r) + "`t" + $値 + "`t" + $字 + "`t" + $式 + "`t" + $一部か + "`t" + $ゼロか + "`t" + $型)
    }
  }
  $押し時計.Stop()
  $押し秒 = [math]::Round($押し時計.Elapsed.TotalSeconds, 2)
  $行.Add('# ★★押すのに かかった 秒 ... ' + $押し秒 + '★★')

  # ★最後に もう 一度 溢れさせて 保存する★（★開き直しては いません★）
  $sh.Range('D1:F10').ClearContents() | Out-Null
  $sh.Range('D1').Formula2 = '=SEQUENCE(3)'
  if (Test-Path $保存先) { Remove-Item $保存先 -Force }
  $bk.SaveAs($保存先, 51)
  $行.Add('# ★保存した★ ... ' + $保存先 + '（★中の 形は 別の 道具で 読みます★）')

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ('★★押すのに かかった 秒 ... ' + $押し秒 + '秒★★（台本 ' + $台本.Count + '本）')
  Write-Host ('★書いた ... ' + $出 + '★')
  Write-Host ('★保存した ... ' + $保存先 + '★')
  $bk.Close($false)
} finally {
  $c = $null
  $w = $null
  $sh = $null
  $bk = $null
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $xl = $null
  $t = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 120)) {
    Start-Sleep -Milliseconds 250
  }
  $t.Stop()
  $残り = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
  if ($残り -eq 0) {
    Write-Host ('★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒で ★消えました★★')
  } else {
    Write-Host ('★★Excel は 消えませんでした★★ ／ 残り ' + $残り + '個')
  }
}
