# toru-jitsu-excel-ga-ji-wo-dasu-ichiban-semai-haba.ps1
#   -- ★実Excel が # で なく 字を 出す 一番 狭い 列★を 1組ずつ 取る（134）（2026-10-04・経営者）
#
#  ★★なぜ★★
#    Exally は 列の 幅を ★お客さんの 機械の 字体★ で 測って いた（Linux の WebKit で 74点 → 93点・Exally1 実測）。
#    ⇒ 列の 幅を ★実Excel が 持つ 数の 表★ から 引く 直しに なる（Exally1 の 案①）。
#    ⇒ その 受け入れの 数＝★「どの 列の 幅で # から 字に 切り替わるか」が 1組ずつ Excel と 同じか★。
#    ★列の 幅 1つ（8.44 など）で 合う／合わない を 見るだけ では、境目から 遠い 組は どちらでも 合う★。
#
#  ★★取り方★★
#    ・物差しの tsv（道具132 の 出し）から 組・値の式・書式 を 読む（★字の 答えの 組 ="abc" は 外す★＝字は # に ならない）
#    ・★新しい 本★（司さんの 本は 使わない）・★Normal の 字体を 名指しで 揃える★（既定 游ゴシック 11）
#    ・1組 1列。★列の 幅を 二分で 探す★：入る（.Text が # だけで ない）／入らない
#      ＝ 0.1〜255字 の 間を 1/512字 まで 詰める ⇒ ★入る 一番 狭い 列（字・px）★と ★入らない 一番 広い 列（字・px）★
#    ・px ＝ 列の .Width（点）× 96/72（★100% の 時の 画面の px★）
#    ・★255字でも 入らない★＝「いつも #」（負の 時刻 等）／★0.1字（見える 一番 狭い 列）でも 入る★＝「いつも 字」
#
#  ★★0 を 2つの 窓で 見る★★（tests/monosashi-mado.test.mjs）
#    ★2つ目の 窓★ 隣の 列に `=(式)=0` を 打ち 本当に 0 か／★3つ目の 窓★ .Value2 の 型（-is [string] / -is [double]）
#
#  ★★門★★
#    ①5.1（exit 8）／②物差しが 在る（exit 4）／③走らせる 前の Excel が 0個（exit 3）
#    ④★頼んだ 数 ≠ 書いた 数 なら 赤（exit 6）★／⑤★書式が 付かなかった 行が 在れば 赤（exit 9）★
#
#  ★★2026-10-04 追記 ── -マスの大きさ★★（Exally1 の 問い「余白の 0の幅 は マスの 字か Normal か」）
#    渡すと ★Normal は -大きさ の まま・マスの 字だけ その 大きさ★ に する（渡さなければ 今までと 1行も 変わらない）
#
#  使い方:
#    powershell.exe -NoProfile -ExecutionPolicy Bypass -File docs/measured/toru-jitsu-excel-ga-ji-wo-dasu-ichiban-semai-haba.ps1 -物差し <tsv> -出す先 <tsv> [-字体 游ゴシック] [-大きさ 11]

param([string]$物差し = '', [string]$出す先 = '', [string]$字体 = '游ゴシック', [double]$大きさ = 11, [double]$マスの大きさ = 0)

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }
if (-not $物差し -or -not (Test-Path -LiteralPath $物差し)) { Write-Host '★★物差しが 在りません★★'; exit 4 }
if (-not $出す先) { Write-Host '★★-出す先 を 渡して ください★★'; exit 4 }

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

# ★物差しから 組を 読む★（頭の 行 2つ を 飛ばす）
$頼 = New-Object System.Collections.ArrayList
$行ら = [System.IO.File]::ReadAllLines($物差し, (New-Object System.Text.UTF8Encoding($false)))
foreach ($l in $行ら) {
  if ($l.StartsWith('#') -or $l.StartsWith('組')) { continue }
  $a = $l.Split("`t")
  if ($a.Length -lt 3) { continue }
  if ($a[1] -eq '="abc"') { continue }
  [void]$頼.Add(@($a[0], $a[1], $a[2]))
}
Write-Host ('★頼んだ 数★ ' + $頼.Count + '（字の 答えの 組は 外した）')

# ★★入るか＝「幅60 で 出る 字」と 同じ 字が 出たか★★（2026-10-04 に 直した）
#   ★前は 「# だけで なければ 入る」と して いた★ ⇒ ★狭い 列では Excel は # も 出さず .Text が 空★（0.1字・0.5字 で 実測）
#   ⇒ 空を 「入る」と 取り 158組を 「いつも 字」と 出した（★道具の 誤り★）
function 入るか($c, $全) { return ([string]$c.Text -ceq $全) }

$xl = New-Object -ComObject Excel.Application
$wb = $null; $ws = $null; $c = $null; $col = $null; $c2 = $null; $st = $null
$出 = New-Object System.Collections.ArrayList
$付かず = 0
$版xl = ''; $実字体 = ''
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $版xl = [string]$xl.Version + ' build ' + [string]$xl.Build
  $wb = $xl.Workbooks.Add()
  $st = $wb.Styles.Item('Normal')
  $st.Font.Name = $字体
  $st.Font.Size = $大きさ
  $実字体 = [string]$st.Font.Name + ' ' + [string]$st.Font.Size
  $st = $null
  $ws = $wb.Worksheets.Item(1)
  $i = 0
  foreach ($q in $頼) {
    $i++
    $colI = 2 * $i - 1
    $c = $ws.Cells.Item(1, $colI)
    $c.Formula = [string]$q[1]
    if ($マスの大きさ -gt 0) { $c.Font.Size = $マスの大きさ }
    $誤 = ''
    try { $c.NumberFormatLocal = [string]$q[2] } catch {
      $換 = ([string]$q[2]).Replace('[Red]', '[赤]')
      if ([string]$q[2] -eq 'General') { $換 = 'G/標準' }
      try { $c.NumberFormatLocal = $換; $誤 = '日本語の 名（' + $換 + '）で 付けた' } catch { $誤 = '★付かない★'; $付かず++ }
    }
    $v = $c.Value2
    $型 = '他'
    if ($v -is [string]) { $型 = 'String' } elseif ($v -is [double]) { $型 = 'Double' }
    $c2 = $ws.Cells.Item(1, $colI + 1)
    $c2.Formula = ('=(' + ([string]$q[1]).Substring(1) + ')=0')
    $真 = [string]$c2.Value2
    $c2 = $null
    $col = $ws.Columns.Item($colI)
    $col.ColumnWidth = [double]60
    $全 = [string]$c.Text
    # ★二分★
    $col.ColumnWidth = [double]255
    $広く入る = (入るか $c $全) -and -not ($全 -match '^#+$')
    # ★下の 端は 「見える 一番 狭い 列」★（2026-10-04 に 1回 踏んだ）
    #   ＝ 0.01字 に すると Excel は 列を 隠し（幅 0）、★隠れた 列の .Text は 字を そのまま 返す★
    #   ⇒ 158組が 「いつも 字」と 出た（幅4 で 98組が # の 真値と 合わない）
    $col.ColumnWidth = [double]0.1
    $下px = [double]$col.Width * 96 / 72
    if ($下px -le 0 -or $col.Hidden) { throw ('★下の 端の 列が 見えない★ px ' + $下px) }
    $狭く入る = 入るか $c $全
    $入る字 = ''; $入るpx = ''; $入らぬ字 = ''; $入らぬpx = ''; $判 = ''
    if (-not $広く入る) { $判 = 'いつも #' }
    elseif ($狭く入る) { $判 = 'いつも 字' }
    else {
      $lo = 0.1; $hi = 255.0
      while (($hi - $lo) -gt (1.0 / 512)) {
        $mid = ($lo + $hi) / 2
        $col.ColumnWidth = [double]$mid
        if (入るか $c $全) { $hi = $mid } else { $lo = $mid }
      }
      $col.ColumnWidth = [double]$hi
      $入る字 = [string]$col.ColumnWidth; $入るpx = [string]([math]::Round([double]$col.Width * 96 / 72, 2))
      $col.ColumnWidth = [double]$lo
      $入らぬ字 = [string]$col.ColumnWidth; $入らぬpx = [string]([math]::Round([double]$col.Width * 96 / 72, 2))
      $判 = '境目'
    }
    $字60 = $全
    [void]$出.Add(($q[0] + "`t" + $q[1] + "`t" + $q[2] + "`t" + $判 + "`t" + $入る字 + "`t" + $入るpx + "`t" + $入らぬ字 + "`t" + $入らぬpx + "`t" + $字60 + "`t" + $型 + "`t" + $真 + "`t" + $誤))
    $col = $null; $c = $null
  }
  $wb.Close($false)
} finally {
  $c = $null; $c2 = $null; $col = $null; $st = $null; $ws = $null; $wb = $null
  $xl.Quit()
  $xl = $null
}

$書 = New-Object System.IO.StreamWriter($出す先, $false, (New-Object System.Text.UTF8Encoding($false)))
$書.NewLine = "`n"
$書.WriteLine('# この 紙は 式の 答えの 紙では ありません（列の 幅の 境目の 物差し）')
$書.WriteLine('# 134 実Excel が 字を 出す 一番 狭い 列／Excel ' + $版xl + '／本の 既定の 字体 ' + $実字体 + '／マスの 大きさ ' + $(if ($マスの大きさ -gt 0) { [string]$マスの大きさ } else { '（Normal と 同じ）' }) + '／頼んだ ' + $頼.Count + '／書いた ' + $出.Count + '／px は 100% の 画面')
$書.WriteLine("組`t値の式`t書式`t判じ`t入る一番狭い(字)`t入る一番狭い(px)`t入らない一番広い(字)`t入らない一番広い(px)`t幅60の字`t値の型(.Value2)`t本当に0か(=(式)=0)`t誤り")
foreach ($l in $出) { $書.WriteLine($l) }
$書.Close()
$境 = @($出 | Where-Object { $_.Split("`t")[3] -eq '境目' }).Count
$常井 = @($出 | Where-Object { $_.Split("`t")[3] -eq 'いつも #' }).Count
$常字 = @($出 | Where-Object { $_.Split("`t")[3] -eq 'いつも 字' }).Count
Write-Host ('★Excel★ ' + $版xl + ' ／ ★字体★ ' + $実字体)
Write-Host ('★頼んだ ' + $頼.Count + ' ／ 書いた ' + $出.Count + ' ／ 境目 ' + $境 + ' ／ いつも # ' + $常井 + ' ／ いつも 字 ' + $常字 + '★')
Write-Host ('★書式が 付かなかった 行★ ' + $付かず)
if ($頼.Count -ne $出.Count) { Write-Host '★★頼んだ 数と 書いた 数が 違います★★'; exit 6 }
if ($付かず -ne 0) { exit 9 }
