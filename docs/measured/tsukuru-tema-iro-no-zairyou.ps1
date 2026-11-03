# tsukuru-tema-iro-no-zairyou.ps1
#   -- ★テーマの 色／番号の 色で 塗った 材料を 実Excel に 作らせる★（78）（2026-09-21）
#
#  ★★なぜ★★
#    Exally1 が 飾りを 画面まで 出す 所を 作りました（7/13 ⇒ 12/13）。
#    ⇒★但し 「テーマの 色（`<color theme=`）」と 「番号の 色（`indexed`）」は 読んで いません★
#    ⇒★当て推量で 黒を 入れない★と 決めた＝★間違った 色より 色なしの 方が まし★
#    ⇒★だから 実Excel が どう 書くか／どう 見せるかを 先に 測ります★
#
#  ★★入れる 物★★
#    B2 ... ★テーマの 色で 塗る★（`Interior.ThemeColor`）
#    B4 ... ★テーマの 色の 字★（`Font.ThemeColor`）
#    B6 ... ★番号の 色で 塗る★（`Interior.ColorIndex`）
#    B8 ... ★濃さを 変えた テーマの 色★（`TintAndShade`）
#    ★対照★ B10 ... ★RGB で 直に 塗る★（★今 読めて いる 形★）
#
#  ★★置き場★★ ... B列の 2 4 6 8 10 行（★互いに 離す★）
#  ★★読む 物★★ ... 実Excel が 返す ★RGB の 数★ ＋ ★テーマの 番号★ ＋ ★濃さ★
#    ⇒★生の `xl/styles.xml` は 別の 道具で 読みます★（1つの 道具に 2つの 仕事を させない）
#
#  ★★置き場★★ `%TEMP%\exally-tema-iro.xlsx`（★1本の 名★）
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③★書いた 後に 読み返して 5つとも 色が 付いて いるか★（1つでも 欠けたら exit 5）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$出す先 = Join-Path $env:TEMP 'exally-tema-iro.xlsx'
$許す名 = 'exally-tema-iro.xlsx'
if ((Split-Path $出す先 -Leaf) -ne $許す名) { Write-Host '★★書いて よい 名は 1本だけです★★'; exit 7 }

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null
$出来た = $false
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)

  # ── ★色の 材料★
  $sh.Range('B2').Value2 = 'tema-nuri'
  $sh.Range('B2').Interior.ThemeColor = 5          # msoThemeColorAccent1
  $sh.Range('B4').Value2 = 'tema-ji'
  $sh.Range('B4').Font.ThemeColor = 6              # msoThemeColorAccent2
  $sh.Range('B6').Value2 = 'bangou-nuri'
  $sh.Range('B6').Interior.ColorIndex = 6          # ★番号の 色（黄）★
  $sh.Range('B8').Value2 = 'tema-koku'
  $sh.Range('B8').Interior.ThemeColor = 5
  $sh.Range('B8').Interior.TintAndShade = -0.5     # ★濃さを 変える★
  $sh.Range('B10').Value2 = 'rgb-nuri'
  $sh.Range('B10').Interior.Color = 65535          # ★対照＝RGB で 直に★

  if (Test-Path $出す先) { Remove-Item $出す先 -Force }
  $bk.SaveAs($出す先, 51)
  $bk.Close($false)
  $bk = $null

  # ══ ★★読み返して 数える★★ ══
  $bk = $xl.Workbooks.Open($出す先, 0, $true)
  $sh = $bk.Sheets.Item(1)
  $欠け = 0
  Write-Host ''
  Write-Host '★★読み返し★★'
  $見る = @(
    @{ 札 = 'B2  テーマ塗り';   マス = 'B2';  何 = 'nuri' },
    @{ 札 = 'B4  テーマ字';     マス = 'B4';  何 = 'ji' },
    @{ 札 = 'B6  番号塗り';     マス = 'B6';  何 = 'nuri' },
    @{ 札 = 'B8  テーマ濃さ';   マス = 'B8';  何 = 'nuri' },
    @{ 札 = 'B10 RGB塗り(対照)'; マス = 'B10'; 何 = 'nuri' }
  )
  foreach ($m in $見る) {
    $c = $sh.Range($m.マス)
    if ($m.何 -eq 'nuri') {
      $rgb = [string]$c.Interior.Color
      $tem = '(無し)'
      try { $tem = [string]$c.Interior.ThemeColor } catch { }
      $ti = [string]$c.Interior.TintAndShade
      $ci = [string]$c.Interior.ColorIndex
      Write-Host ('  ' + $m.札.PadRight(20) + ' RGB ' + $rgb.PadRight(10) + ' テーマ ' + $tem.PadRight(6) + ' 濃さ ' + $ti.PadRight(6) + ' 番号 ' + $ci)
      if ($rgb -eq '16777215') { $欠け++ }
    } else {
      $rgb = [string]$c.Font.Color
      $tem = '(無し)'
      try { $tem = [string]$c.Font.ThemeColor } catch { }
      $ti = [string]$c.Font.TintAndShade
      Write-Host ('  ' + $m.札.PadRight(20) + ' RGB ' + $rgb.PadRight(10) + ' テーマ ' + $tem.PadRight(6) + ' 濃さ ' + $ti)
      if ($rgb -eq '0') { $欠け++ }
    }
    $c = $null
  }
  $bk.Close($false)
  $bk = $null
  Write-Host ('★★欠け ... ' + $欠け + '個★★')
  if ($欠け -ne 0) { exit 5 }
  $出来た = $true
} finally {
  # ★★2026-09-20 ── ★図形の 持ち手も 手放す★★
  #   ＝`$かたち`（`Shapes.AddShape` が 返す 物）を $null に して いなかった ので
  #     ★Excel が 120秒 では 消えませんでした★（3本目で 実際に 出ました）
  #   ⇒記憶「実Excel を COM＝5.1 かつ ★掴んだ物 全部 $null★ の 2つ 揃った 時だけ 消える」
  $かたち = $null
  $sh = $null
  if ($null -ne $bk) { $bk.Close($false) }
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
  if ($残り -eq 0) { Write-Host ('★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒で ★消えました★★') }
  else { Write-Host ('★★Excel は 消えませんでした★★ ／ 残り ' + $残り + '個') }
}
if ($出来た) {
  $x = Get-Item $出す先
  $h = (Get-FileHash $出す先 -Algorithm SHA256).Hash.ToLower()
  Write-Host ''
  Write-Host ('★★作りました★★ ... ' + $出す先)
  Write-Host ('  ★大きさ★ ' + $x.Length + ' バイト')
  Write-Host ('  ★sha256★ ' + $h)
}
