# toru-cell5.ps1 — ★CELL("format") の 日付・時刻の 合図を 詰める★（2026-09-07）
#
#  ★3回目で 分かった 事★
#    D1 … 年+月+日 ／ D2 … 月+年 ／ D3 … 月+日 ／ D6〜D9 … 時刻
#  ★でも まだ 決まらない 形が 在る★（★どちらでも 合う 組で 決めない★）
#    ・年だけ（yyyy）／月だけ（mmm）／日だけ（dd）
#    ・秒だけ（ss）／分:秒（mm:ss）／経過時間（[h]:mm）
#    ・年+月（yyyy/m）／日+年（d-mmm-yyyy）
#    ・★D4・D5 が 出る 形が 在るのか★（3回目では 1本も 出なかった）
#
#  使い方: powershell -File docs/measured/kansuu46/toru-cell5.ps1

$ErrorActionPreference = 'Stop'

# ══ ★★2つ目の 窓（2026-09-08 に 足した）★★ ══════════════════
#  ★物差しの 欠陥★ .Value2 は ★0 で ない 値に 0 を 返す★
#    =0.1+0.2-0.3    … .Value2 ★0★ ／ =(式)=0 ★False★ ／ (式)*1e17 5.55
#    =11.1+22.2-33.3 … .Value2 0   ／ =(式)=0 ★True★  ／ (式)*1e17 0
#    ⇒★.Value2 では この 2つが どちらも 0 に 見える★
#    正体 …★最後の 演算が ＋か− の 時だけ 実Excel が ★見せる 時に★ 0 に する★
#  ★もう1つ★ =DEC2BIN(0.5) は ★文字列の "0"★＝数の 0 では ない
#    ⇒ ="0"=0 は FALSE ⇒★見せかけの 0 と 同じ 顔★⇒★型を 見ないと 分けられない★
#  ⇒★見張り tests/monosashi-mado.test.mjs が これを 入れて いない 道具を 赤に する★
function 窓２_型($v) {
  if ($null -eq $v) { return 'Empty' }
  if ($v -is [string]) { return 'String' }
  if ($v -is [bool]) { return 'Boolean' }
  if ($v -is [double] -or $v -is [int] -or $v -is [long]) { return 'Number' }
  return 'Other'
}
function 窓２_本当にゼロか($sh, [string]$式) {
  # ★『0』が 出た 時だけ 呼ぶ★ … =(式)=0 の 真偽を 返す
  $中 = $式 -replace '^=\s*', ''
  try {
    $sh.Range('BZ1').Clear() | Out-Null
    $sh.Range('BZ1').Formula = ('=(' + $中 + ')=0')
    $z = $sh.Range('BZ1').Value2
    $sh.Range('BZ1').Clear() | Out-Null
    if ($z -is [bool]) { return $(if ($z) { 'TRUE' } else { 'FALSE' }) }
    return '★判じられない★'
  } catch { return '★判じられない★' }
}

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-cell5-2026-09-07.tsv'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$版 = $xl.Version; $ビルド = $xl.Build
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)

$結果 = New-Object System.Collections.ArrayList
$形たち = @(
  'yyyy', 'yy', 'mmm', 'mmmm', 'dd', 'd', 'aaa', 'aaaa',
  'ss', 'mm:ss', 'h', '[h]:mm', '[mm]:ss',
  'yyyy/m', 'yyyy"年"m"月"', 'd-mmm-yyyy', 'yyyy/m/d h:mm:ss',
  'm/d/yy', 'mm/dd/yy', 'dd/mm/yy', 'yy/m/d',
  '0.00_ ', '#,##0"円"', '"("#,##0")"', '0"個"',
  '[>100]0;[<=100]0.00', '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-'
)
$r = 1
foreach ($f in $形たち) {
  $c = $ws.Cells.Item($r, 1)
  $c.Value2 = 45293.5
  $付いた = $true
  try { $c.NumberFormatLocal = $f } catch {
    $付いた = $false
    [void]$結果.Add(("CELL`t(表示形式を 付けられない)`t★付かない★`t表示形式 {0}" -f $f))
  }
  if ($付いた) {
    $実 = [string]$c.NumberFormatLocal
    foreach ($k in @('format', 'color', 'parentheses')) {
      $ws.Range('H1').Formula = '=CELL("' + $k + '",A' + $r + ')'
      $v = $ws.Range('H1').Value2
      $答 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { $v.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
      [void]$結果.Add(("CELL`t=CELL(""{0}"",A{1})`t{2}`t表示形式 {3}（実際に 付いた 形 {4}）" -f $k, $r, $答, $f, $実))
    }
  }
  $r++
}

$wb.Close($false); $xl.Quit()
foreach ($o in @($ws, $wb, $xl)) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null }

$頭 = @(
  "# ★実Excel に 打たせた CELL の 答え（5回目・日付/時刻の 合図を 詰める）★",
  "# 取った 日 … 2026-09-07",
  "# ★どの Excel で 打ったか★ … 版 $版 ／ build $ビルド",
  "# ★A列に 45293.5（2024/1/2 12:00）を 置いて 表示形式だけ 変えた★",
  "# 関数`t式`t実Excel の 答え`tどんな 場面か"
)
[System.IO.File]::WriteAllText($出, ((($頭 + $結果) -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding $false))
Write-Host "★書いた … $出（$($結果.Count) 本）★"
