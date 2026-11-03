# tsukuru-xlsb-tema-iro-1tsu-zutsu.ps1
#   -- ★テーマの 色の 番号が 実際に 何色かを 実Excel に 言わせる★（88）（2026-09-21）
#
#  ★★なぜ★★
#    87 で ★色は どこでも 同じ 8バイト★ と 分かり、
#    ★kind=7 ＝テーマの 色★ まで 割れました。
#    ⇒★でも 「テーマの 1番が 何色か」は 1つも 測って いません★
#    ⇒テーマの 色の マスは ★画面に 出す 時に 実際の 色に 直さないと いけません★
#
#  ★★やり方★★
#    1マスに つき ★テーマの 色を 1つだけ★ 当てる（COM の `Font.ThemeColor`）
#    ⇒★同じ マスから 2つ 読みます★
#        ①`Font.Color` ... ★Excel が 実際の 色に 直した 数★（COM は BGR）
#        ②包みの `styles.bin` ... ★kind と idx★
#    ⇒★COM の 番号／包みの 番号／実際の 色★の 3つが 1行に 並びます
#    ★どれも 私が 決めて いません★＝Excel が 返した 物だけ
#
#  ★★ついでに 測る 物★★
#    `Theme.ThemeColorScheme.Colors(i).RGB` ... ★テーマそのものの 色★
#    ⇒`Font.Color` と 合うかを 見ます（合わなければ ★どちらかが 別の 物★）
#
#  ★★作る 物★★ `%TEMP%\exally-tema-iro.xlsb`（★1本の 名★）
#
#  ★門★
#    ①貝殻が 5.1（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③★10色 とも 当たって いるか（読み返す）★（exit 5）／④名が 決め打ち（exit 7）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$出す先 = Join-Path $env:TEMP 'exally-tema-iro.xlsb'
if ((Split-Path $出す先 -Leaf) -ne 'exally-tema-iro.xlsb') { exit 7 }

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

# ★COM の テーマ色の 番号★（1〜10）... ★名前は Excel の 物を そのまま★
$札 = @{ 1 = 'Dark1'; 2 = 'Light1'; 3 = 'Dark2'; 4 = 'Light2'; 5 = 'Accent1';
         6 = 'Accent2'; 7 = 'Accent3'; 8 = 'Accent4'; 9 = 'Accent5'; 10 = 'Accent6' }

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)

  Write-Host ''
  Write-Host '★★1マスに 1色ずつ 当てて 読み返します★★'
  $効いた = 0
  for ($n = 1; $n -le 10; $n++) {
    $sh.Cells.Item($n, 1).Value2 = $n
    $sh.Cells.Item($n, 1).Font.ThemeColor = $n
    $戻 = [string]$sh.Cells.Item($n, 1).Font.ThemeColor
    $色 = $sh.Cells.Item($n, 1).Font.Color
    # ★COM は BGR★ ⇒ ★RGB の 字に 直します★（直す 式も 紙に 残します）
    $b = [int]([math]::Floor($色 / 65536)) -band 255
    $g = [int]([math]::Floor($色 / 256)) -band 255
    $r = [int]$色 -band 255
    $rgb = '{0:x2}{1:x2}{2:x2}' -f $r, $g, $b
    if ($戻 -eq [string]$n) { $効いた++ }
    Write-Host ('  A' + $n + ' ... COM の 番号 ' + $n + '（' + $札[$n] + '）' +
                ' ／ 読み返し ' + $戻 + ' ／ Font.Color ' + $色 + ' ⇒ ★#' + $rgb + '★')
  }
  Write-Host ('★効いた 色★ ' + $効いた + ' / 10')
  if ($効いた -ne 10) { Write-Host '★★当たって いない 色が 在ります★★'; exit 5 }

  Write-Host ''
  Write-Host '★★テーマそのものの 色★★（`Theme.ThemeColorScheme`）'
  for ($i = 1; $i -le 12; $i++) {
    $c = $bk.Theme.ThemeColorScheme.Colors($i).RGB
    $b = [int]([math]::Floor($c / 65536)) -band 255
    $g = [int]([math]::Floor($c / 256)) -band 255
    $r = [int]$c -band 255
    Write-Host ('  ' + $i.ToString().PadLeft(2) + ' ... ' + $c.ToString().PadLeft(10) +
                ' ⇒ ★#' + ('{0:x2}{1:x2}{2:x2}' -f $r, $g, $b) + '★')
  }

  if (Test-Path $出す先) { Remove-Item $出す先 -Force }
  $bk.SaveAs($出す先, 50)
  # ★★2026-09-21 ── ★`.xlsx` でも 同じ 物を 出します★★
  #   Exally1「`.xlsx` の `styles.xml` は `<color theme="4"/>` と ★番号しか 書いて いない★」
  #   ⇒★同じ 盤面を 2つの 形で 保存して 突き合わせます★
  #     ＝★形が 違う なら 書き分けが 要ります★
  $出す先2 = Join-Path $env:TEMP 'exally-tema-iro.xlsx'
  if ((Split-Path $出す先2 -Leaf) -ne 'exally-tema-iro.xlsx') { exit 7 }
  if (Test-Path $出す先2) { Remove-Item $出す先2 -Force }
  $bk.SaveAs($出す先2, 51)
  $bk.Close($false); $sh = $null; $bk = $null
} finally {
  $sh = $null
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

if (-not (Test-Path $出す先)) { Write-Host '★★出来て いません★★'; exit 5 }
$x = Get-Item $出す先
Write-Host ''
Write-Host ('★★作りました★★ ... ' + $出す先)
Write-Host ('  ★大きさ★ ' + $x.Length + ' バイト ／ ★sha256★ ' + (Get-FileHash $出す先 -Algorithm SHA256).Hash.ToLower())
$出す先2 = Join-Path $env:TEMP 'exally-tema-iro.xlsx'
if (Test-Path $出す先2) {
  $x2 = Get-Item $出す先2
  Write-Host ('★★作りました（xlsx）★★ ... ' + $出す先2)
  Write-Host ('  ★大きさ★ ' + $x2.Length + ' バイト ／ ★sha256★ ' + (Get-FileHash $出す先2 -Algorithm SHA256).Hash.ToLower())
} else { Write-Host '★★xlsx が 出来て いません★★'; exit 5 }
