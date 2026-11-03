# toru-shoshiki-dai.ps1 — ★書式の 台が 何を すれば よいか 実Excel に 聞く★（2026-09-15）
#
#  ★★なぜ★★
#    ★指示役1 の 決め（2026-09-15）★＝★書式を 読む 台を 1つ 作り、TEXT も 画面も それを 呼ぶ★。
#    ★測って 決めた 相手★（`docs/measured/toru-jitsubutsu-shoshiki.ps1` の 実測）
#      ㋐ TEXT 740回 … ★aaa（731）／m/d（9）の 2種類だけ★
#      ㋑ 画面 67,542マス … ★18種★（上位 6種で 95.6%）
#    ⇒★この 紙は「その 18種 × 代表の 値」を 実Excel に 打たせた 物★です。
#
#  ★★この 紙で 決まる 一番 大事な 事★★
#    ★同じ 書式で「マスに 付けた 時」と「TEXT() に 渡した 時」が 同じ 字に なるか★
#    ⇒★同じなら 台は 1つで 足ります★／★違うなら 違う 所を 名指しで 出す★
#
#  ★★踏まない 様に した 罠★★
#    ①★`.Text` は 列の 幅で 変わります★（狭いと `###`）⇒★先に 幅を 広げる★
#    ②★`NumberFormatLocal`（この国の 字）と `NumberFormat`（世界共通の 字）は 別物★
#       ⇒★両方 紙に 書きます★（TEXT に 渡すのは この国の 字）
#    ③前の 跡を 消してから 打ち、CalculateFull してから 読む
#    ④終わりに Excel が 残って いないか 数える
$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-shoshiki-dai-2026-09-15.tsv'

$誤りの番号 = @{
  -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!';
  -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!';
  -2146826246 = '#N/A';   -2146826243 = '#SPILL!'; -2146826238 = '#CALC!'; -2146826237 = '#BUSY!'
}

# ★実物が 使って いる 書式★（伏せ字の 4種を 除いた 14種＝★99.9% を 賄う 分★）
$書式たち = @(
  'G/標準',
  '#,##0_ ',
  '#,##0_);[赤](#,##0)',
  '#,##0.00_ ',
  '0.00_);[赤](0.00)',
  'm/d;@',
  '#,##0.00_);[赤](#,##0.00)',
  'm/d(aaa)',
  '#,##0.000_ ',
  'm"月"d"日";@',
  'yyyy"年"m"月"',
  '#,##0;[赤]-#,##0',
  '0_);[赤](0)',
  '#,##0.0_',
  # ★TEXT が 使う 2種★（実物 740回）
  'aaa',
  'm/d'
)

# ★代表の 値★（★0・負・大きい・日付・時刻・字・真偽・空★）
$値たち = @(
  @('数0',       0),
  @('数1234.5',  1234.5),
  @('数-1234.5', -1234.5),
  @('数1234567', 1234567),
  @('数0.5',     0.5),
  @('数-0.5',    -0.5),
  @('数0.004',   0.004),
  @('数-0.004',  -0.004),
  @('日45292',   45292),
  @('日45658',   45658),
  @('時0.5208',  0.520833333333333),
  @('数1.5',     1.5)
)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
$版 = $xl.Version; $ビルド = $xl.Build
$wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
# ★★列の 幅を 広げる★★（`.Text` は 狭いと `###` に なります）
$ws.Columns.Item(1).ColumnWidth = 40
$ws.Columns.Item(2).ColumnWidth = 40
$ws.Columns.Item(3).ColumnWidth = 40

$結果 = New-Object System.Collections.ArrayList
$j = 0
foreach ($書 in $書式たち) {
  $j++
  Write-Host ("  … {0}/{1}  {2}" -f $j, $書式たち.Count, $書)
  $ws.Range('A1:C40').Clear() | Out-Null
  $ws.Range('A1:C40').NumberFormatLocal = 'G/標準'
  $世界 = ''
  $i = 0
  foreach ($v in $値たち) {
    $i++
    $ws.Cells.Item($i, 1).Value2 = [double]$v[1]     # 素の 値（見るだけ）
    $ws.Cells.Item($i, 2).Value2 = [double]$v[1]     # ★書式を 付ける 方★
    $ws.Cells.Item($i, 3).Formula = ('=TEXT(B' + $i + ',"' + ($書 -replace '"', '""') + '")')
  }
  try { $ws.Range('B1:B' + $値たち.Count).NumberFormatLocal = $書 } catch { }
  try { $世界 = [string]$ws.Range('B1').NumberFormat } catch { $世界 = '★取れない★' }
  $xl.CalculateFull()
  $i = 0
  foreach ($v in $値たち) {
    $i++
    $見た目 = ''
    try { $見た目 = [string]$ws.Cells.Item($i, 2).Text } catch { $見た目 = '★取れない★' }
    # ══ ★★2つ目の 窓★★ ══（見張り `tests/monosashi-mado.test.mjs` の 決め）
    #   ★見た目が「0」に 見える 時、中も 本当に 0 か★を 別の 窓で 見ます。
    #   ★この 道具では 生の 数を 置いて いる★ので 見せかけの 0 は 出ない はずですが、
    #   ★「はず」で 免除に 逃げない★＝★測って 出します★。
    #   （`.Value2` は ★0 で ない 値に 0 を 返す★事が 在る＝足し引きで 終わる 式の 時）
    $窓２ = '-'
    if ($見た目 -match '^\s*0([.,]0+)?\s*$') {
      try {
        $ws.Range('BZ1').Clear() | Out-Null
        $ws.Range('BZ1').Formula = ('=(B' + $i + ')=0')
        $z = $ws.Range('BZ1').Value2
        $ws.Range('BZ1').Clear() | Out-Null
        $窓２ = if ($z -is [bool]) { $(if ($z) { 'TRUE' } else { 'FALSE' }) } else { '★判じられない★' }
      } catch { $窓２ = '★判じられない★' }
    }
    # ★型も 見ます★（`.Value2` が 何の 型で 返ったか）
    $型 = '-'
    try { $v2 = $ws.Cells.Item($i, 2).Value2
      $型 = if ($null -eq $v2) { 'Empty' } elseif ($v2 -is [string]) { 'String' }
            elseif ($v2 -is [bool]) { 'Boolean' } elseif ($v2 -is [double]) { 'Number' } else { 'Other' }
    } catch { $型 = '-' }
    $tv = $ws.Cells.Item($i, 3).Value2
    $t答 = ''
    if ($null -eq $tv) { $t答 = '(空)' }
    elseif (($tv -is [int] -or $tv -is [long]) -and $誤りの番号.ContainsKey([int]$tv)) { $t答 = $誤りの番号[[int]$tv] }
    elseif ($tv -is [double]) { $t答 = $tv.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture) }
    else { $t答 = [string]$tv }
    $同じ = if ($見た目 -eq $t答) { '同じ' } else { '★違う★' }
    # ★★1列目を 番号に する★★（2026-09-15 に 踏んだ）
    #   ★書式の 字は `#` で 始まります★（`#,##0` …）
    #   ⇒★紙の 決まり「# で 始まる 行は 覚え書き」と ぶつかります★
    #   ⇒★読む 道具が その 行を 黙って 飛ばしました★（16種の うち ★7種が 消えた★）
    #   ⇒★1列目を 番号に して 行の 頭を `#` に しない★
    [void]$結果.Add(("{0}`t{1}`t{2}`t{3}`t{4}`t{5}`t{6}`t{7}`t{8}`t{9}" -f `
      ($結果.Count + 1), $書, $世界, $v[0], ([double]$v[1]).ToString('R', [System.Globalization.CultureInfo]::InvariantCulture), $見た目, $t答, $同じ, $窓２, $型))
  }
}

$wb.Close($false); $xl.Quit()
foreach ($o in @($ws, $wb, $xl)) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null }
[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
$残り = @()
for ($t = 0; $t -lt 40; $t++) {
  Start-Sleep -Milliseconds 500
  $残り = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
  if ($残り.Count -eq 0) { break }
}
if ($残り.Count -eq 0) { Write-Host ('★Excel は 残って いません（0個・' + (($t + 1) * 0.5) + '秒で 消えた）★') }
else { Write-Host ('★★Excel が ' + $残り.Count + '個 残って います＝20秒 待っても 消えません★★') }

# ★★数え方★★ … 8列目が 「★違う★」の 行（★後ろに 列を 足したので `-like '*…'` では 数えられません★）
#   2026-09-15 … 列を 2本 足した 時 ★この 数が 30→0 に なりました★（★静かに 嘘に なる★）
$違い = @($結果 | Where-Object { ($_ -split "`t")[7] -eq '★違う★' }).Count
$頭 = @(
  '# ★書式の 台が 何を すれば よいか 実Excel に 聞いた★（2026-09-15）',
  '#   ★測って 選んだ 書式★＝司さんの 実物が 実際に 使って いる 物',
  '#     ㋐TEXT 740回 … aaa（731）／m/d（9）の ★2種類だけ★',
  '#     ㋑画面 67,542マス … 18種（★伏せ字の 4種を 除いた 14種＝99.9% を 賄う★）',
  "# ★どの Excel で 打ったか★ … 版 $版 ／ build $ビルド",
  '# ★★列の 幅を 40 に 広げて から 読んで います★★（`.Text` は 狭いと ### に なる）',
  '# ★★「この国の 字」と「世界共通の 字」を 並べて 在ります★★',
  '#   1列目 … NumberFormatLocal（★TEXT に 渡すのは こちら★）',
  '#   2列目 … NumberFormat（世界共通）',
  '# ★★一番 大事な 列＝7列目★★',
  '#   ★マスに 付けた 時の 見た目★と ★TEXT() に 渡した 時★が 同じか',
  "#   ⇒★違う 行 … $違い 本★（0 なら ★台は 1つで 足ります★）",
  '# ★★2つ目の 窓★★ … 見た目が 0 の 行だけ `=(B行)=0` の 真偽も 取って 在ります（9列目）',
  '#   ＝`.Value2` は ★0 で ない 値に 0 を 返す★事が 在る（足し引きで 終わる 式）',
  '#   ★この 道具は 生の 数を 置くだけ★＝見せかけの 0 は 出ない はずですが ★測って 出します★',
  '# 番号' + "`t" + '書式(この国)' + "`t" + '書式(世界)' + "`t" + '値の名' + "`t" + '値' + "`t" + 'マスの見た目' + "`t" + 'TEXT()の答え' + "`t" + '同じか' + "`t" + '窓②=0の真偽' + "`t" + '型'
)
[System.IO.File]::WriteAllText($出, ((($頭 + $結果) -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding $false))
Write-Host "★書いた … $出★（行 $($結果.Count) 本 ／ ★違う 行 $違い 本★）"
