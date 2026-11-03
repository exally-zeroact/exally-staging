# toru-cell6.ps1 — ★最後に 残った「どちらでも 合う」形を つぶす★（2026-09-07）
#
#  ★5回目で 分かった 事★
#    ・年/月/日 の うち ★2つ 以上★ 無いと D系に ならない（yyyy だけ は G）
#    ・時刻は ★時（h）が 無いと G★（ss も mm:ss も G）
#    ・かっこは ★" " の 中の ( は 数えない★（"("#,##0")" は 0）
#  ★まだ 決まらない★
#    ・年+日（月が 無い）は D1 か D3 か
#    ・後ろに ;@ が 付いた 形（yyyy/m/d;@）
#    ・時刻に 秒だけ 足した 形（h:ss）
#    ・★D4・D5 は 本当に 出ないのか★（出るなら どの 形か を 探す）
#
#  使い方: powershell -File docs/measured/kansuu46/toru-cell6.ps1

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
$出 = Join-Path $ここ 'golden-cell6-2026-09-07.tsv'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$版 = $xl.Version; $ビルド = $xl.Build
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)

$結果 = New-Object System.Collections.ArrayList
$形たち = @(
  'yyyy/d', 'yyyy"年"d"日"', 'd"日"yyyy"年"',
  'yyyy/m/d;@', 'yyyy/m/d;;', 'm/d;@',
  'h:ss', 'h AM/PM', 'h:mm:ss.0', 'mm:ss.0',
  'yyyy/m/d h:mm AM/PM', 'm/d/yy h:mm',
  'mmm d, yyyy', 'mmmm d', 'yyyy-mm', 'mmm yyyy'
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
    $ws.Range('H1').Formula = '=CELL("format",A' + $r + ')'
    $v = $ws.Range('H1').Value2
    $答 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { $v.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    [void]$結果.Add(("CELL`t=CELL(""format"",A{0})`t{1}`t表示形式 {2}（実際に 付いた 形 {3}）" -f $r, $答, $f, $実))
  }
  $r++
}

$wb.Close($false); $xl.Quit()
foreach ($o in @($ws, $wb, $xl)) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null }

$頭 = @(
  "# ★実Excel に 打たせた CELL の 答え（6回目・残りの あいまいな 形）★",
  "# 取った 日 … 2026-09-07",
  "# ★どの Excel で 打ったか★ … 版 $版 ／ build $ビルド",
  "# ★A列に 45293.5（2024/1/2 12:00）を 置いて 表示形式だけ 変えた★",
  "# 関数`t式`t実Excel の 答え`tどんな 場面か"
)
[System.IO.File]::WriteAllText($出, ((($頭 + $結果) -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding $false))
Write-Host "★書いた … $出（$($結果.Count) 本）★"
