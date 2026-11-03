# toru-tema-iro-wo-excel-ni-hirakaseru.ps1
#   -- ★テーマの 色は ファイルごとに 変わるか★（79）（2026-09-21）
#
#  ★★なぜ★★
#    78 で 実Excel が 書く 字を 見ました:
#      `<fgColor theme="4"/>` ／ `<fgColor theme="4" tint="-0.499984740745262"/>`
#      `<fgColor indexed="13"/>` ／ `<fgColor rgb="FFFFFF00"/>`
#    ⇒★`theme="4"` だけでは 何色か 分かりません★
#    ⇒★`xl/theme/theme1.xml` を 読んで 初めて 色に なります★
#    ⇒★★だから 「theme 4 は この 色」と 決め打ちに できるか★★を 割ります
#
#  ★★やり方★★
#    ①78 が 作った 物 ... `%TEMP%\exally-tema-iro.xlsx`
#    ②その 写しの ★`xl/theme/theme1.xml` の accent1 だけ★ を ★FF0000（赤）★に した 物
#       ... `%TEMP%\exally-tema-iro2.xlsx`（★他は 1バイトも 変えて いません★）
#    ⇒★同じ `theme="4"` が 2つの ファイルで 同じ 色か／違う 色か★
#
#  ★★分かれ道★★
#    ・同じ 色 ⇒★決め打ちに できる★
#    ・違う 色 ⇒★★決め打ちは 間違い＝`theme1.xml` を 読むしか ない★★
#
#  ★★開く 物★★（★字で 書いた 2本だけ★）
#    `%TEMP%\exally-tema-iro.xlsx` ／ `%TEMP%\exally-tema-iro2.xlsx`
#    ・★司さんの 実物の 名は 1文字も 在りません★／★読むだけ★
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②2本とも 無ければ 走らない（exit 6）
#    ③走らせる 前の Excel が 0個（exit 3）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-tema-iro-excel-2026-09-21.tsv'

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$二本 = @(
  @{ 札 = 'moto(そのまま)';        名 = 'exally-tema-iro.xlsx' },
  @{ 札 = 'accent1 を 赤に した'; 名 = 'exally-tema-iro2.xlsx' }
)
foreach ($x in $二本) {
  $p = Join-Path $env:TEMP $x.名
  if (-not (Test-Path $p)) { Write-Host ('★★在りません ... ' + $p + '★★'); exit 6 }
  Write-Host ('★開く 物 ... ' + $p + '（' + (Get-Item $p).Length + ' バイト）★')
}

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個★')
if ($数1 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$見る = @(
  @{ 札 = 'B2  テーマ塗り(theme4)'; マス = 'B2';  何 = 'nuri' },
  @{ 札 = 'B4  テーマ字(theme5)';   マス = 'B4';  何 = 'ji' },
  @{ 札 = 'B6  番号塗り(indexed13)'; マス = 'B6'; 何 = 'nuri' },
  @{ 札 = 'B8  テーマ濃さ';         マス = 'B8';  何 = 'nuri' },
  @{ 札 = 'B10 RGB塗り(対照)';      マス = 'B10'; 何 = 'nuri' }
)

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null; $c = $null; $w2 = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★テーマの 色は ファイルごとに 変わるか★（79）（2026-09-21）')
  $行.Add('# ★読むだけ★（保存して いません）／★開けるのは 字で 書いた 2本だけ★')
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★2本目は `xl/theme/theme1.xml` の accent1 だけ FF0000 に した 写し★（他は 同じ）')
  $行.Add('# 紙' + "`t" + '見た物' + "`t" + 'RGB' + "`t" + 'テーマ番号' + "`t" + '濃さ' + "`t" + '色の番号' + "`t" + '型' + "`t" + '=(マス)=0')
  foreach ($x in $二本) {
    $p = Join-Path $env:TEMP $x.名
    $bk = $xl.Workbooks.Open($p, 0, $true)
    $sh = $bk.Sheets.Item(1)
    # ══ ★★ブックの テーマ そのものを 訊きます★★ ══
    #   ＝★マスの 色が 同じ だった 時、「テーマが 効いて いない」のか
    #     「テーマは 効いて いるが 元々 同じ 色」なのかが 分かりません★
    #   ⇒★ブックの accent1 を 直に 読みます★
    $テ = '(読めません)'
    try { $テ = [string]$bk.Theme.ThemeColorScheme.Colors(5).RGB } catch { $テ = '★投げました★ ' + $_.Exception.Message }
    $行.Add($x.札 + "`t" + '★ブックの テーマ accent1★' + "`t" + $テ + "`t" + '' + "`t" + '' + "`t" + '')
    Write-Host ('  ' + $x.札.PadRight(20) + ' ★ブックの accent1★ ' + $テ)
    $窓行 = 40
    foreach ($m in $見る) {
      $c = $sh.Range($m.マス)
      if ($m.何 -eq 'nuri') {
        $rgb = [string]$c.Interior.Color
        $tem = '(無し)'; try { $tem = [string]$c.Interior.ThemeColor } catch { }
        $ti = [string]$c.Interior.TintAndShade
        $ci = [string]$c.Interior.ColorIndex
      } else {
        $rgb = [string]$c.Font.Color
        $tem = '(無し)'; try { $tem = [string]$c.Font.ThemeColor } catch { }
        $ti = [string]$c.Font.TintAndShade
        $ci = [string]$c.Font.ColorIndex
      }
      # ══ ★★2つ目の 窓★★（★`=(マス)=0` の 真偽と ★型★ を 一緒に 取る★）══
      #   ★なぜ★ ... `.Value2` の 「0」は ★本物の 0★ とも ★空★ とも ★誤りの 番号★ とも
      #              区別が 付きません（記憶「意味の 無い 数は 一番 見つけにくい」）
      #   ★置き場★ ... Z列（★色を 見る マスは B列＝ぶつかりません★）
      $v0 = $c.Value2
      $型 = if ($null -eq $v0) { '(kara)' } elseif ($v0 -is [double]) { 'Double' } elseif ($v0 -is [string]) { 'String' } elseif ($v0 -is [bool]) { 'Boolean' } else { 'Other' }
      $ゼロか = '(★窓2が 打てません★)'
      try { $w2 = $sh.Range('Z' + $窓行); $w2.Formula2 = '=(' + $m.マス + ')=0'; $ゼロか = [string]$w2.Value2 } catch { }
      $窓行 = $窓行 + 1
      $行.Add($x.札 + "`t" + $m.札 + "`t" + $rgb + "`t" + $tem + "`t" + $ti + "`t" + $ci + "`t" + $型 + "`t" + $ゼロか)
      Write-Host ('  ' + $x.札.PadRight(20) + ' ' + $m.札.PadRight(24) + ' RGB ' + $rgb)
      $c = $null
    }
    $bk.Close($false)
    $bk = $null
    $sh = $null
  }
  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ('★書いた ... ' + $出 + '★')
} finally {
  $w2 = $null; $c = $null; $sh = $null
  if ($null -ne $bk) { $bk.Close($false) }
  $bk = $null
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $xl = $null
  $t = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 180)) { Start-Sleep -Milliseconds 250 }
  $残り = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
  if ($残り -eq 0) { Write-Host '★Excel は 消えました★' } else { Write-Host ('★★Excel は 消えませんでした★★ ' + $残り + '個') }
}
