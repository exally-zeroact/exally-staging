# toru-cell2.ps1 — ★CELL の「見た目」を もっと 細かく 実Excel に 聞く★（2026-09-07）
#
#  ★1回目（toru-cell.ps1）で 足りなかった 所★
#    ・`format` … 表示形式ごとの 合図（G / F2 / ,0 / C2 / P0 / D1 … ）を
#      ★1つしか 測っていなかった★ ⇒ ここで 24通り 測る
#    ・`color`  … ★負の 数を 赤で 出す 形★の 時だけ 1に なるはず ⇒ 確かめる
#    ・`parentheses` … ★かっこを 付ける 形★の 時だけ 1に なるはず ⇒ 確かめる
#    ・`prefix` … 左/中央/右/繰り返し で 何が 返るか（字と 数の 両方）
#    ・`type` `contents` … 式が 空文字を 返す時／誤りの 時
#
#  使い方: powershell -File docs/measured/kansuu46/toru-cell2.ps1

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
$出 = Join-Path $ここ 'golden-cell2-2026-09-07.tsv'

$誤りの番号 = @{
  -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!';
  -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!';
  -2146826246 = '#N/A'
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$版 = $xl.Version; $ビルド = $xl.Build; $言語 = $xl.LanguageSettings.LanguageID(2)
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)

$結果 = New-Object System.Collections.ArrayList
function 打つ($式) {
  $ws.Range('H1').Formula = $式
  $v = $ws.Range('H1').Value2
  if ($null -eq $v) { return @('(空)', 'null') }
  $t = $v.GetType().Name
  if ($v -is [double]) { return @($v.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture), $t) }
  if ($v -is [int] -and $誤りの番号.ContainsKey([int]$v)) { return @($誤りの番号[[int]$v], 'error値') }
  return @([string]$v, $t)
}

# ── ①表示形式ごとの format / color / parentheses ──────────────
#  ★表示形式は ★英語の 記号★で 入れる★（.NumberFormat は 英語・.NumberFormatLocal は 日本語）
$形たち = @(
  'General', '0', '0.00', '0.000',
  '#,##0', '#,##0.00',
  '$#,##0_);($#,##0)', '$#,##0.00_);($#,##0.00)',
  '$#,##0_);[Red]($#,##0)', '$#,##0.00_);[Red]($#,##0.00)',
  '0%', '0.00%',
  '0.00E+00', '# ?/?', '# ??/??',
  'm/d/yy', 'd-mmm-yy', 'd-mmm', 'mmm-yy', 'mm/dd',
  'h:mm AM/PM', 'h:mm:ss AM/PM', 'h:mm', 'h:mm:ss',
  'm/d/yy h:mm', '@',
  '0;[Red]0', '0;-0', '#,##0_);(#,##0)', '#,##0.00_);[Red](#,##0.00)'
)
$r = 1
foreach ($f in $形たち) {
  $c = $ws.Cells.Item($r, 1)
  $c.Value2 = 1234.5678
  $付いた = $true
  try { $c.NumberFormat = $f } catch {
    $付いた = $false
    [void]$結果.Add(("CELL`t(表示形式を 付けられない)`t★付かない★`t-`t表示形式 {0}" -f $f))
  }
  if (-not $付いた) { $r++; continue }
  foreach ($k in @('format', 'color', 'parentheses')) {
    $式 = '=CELL("' + $k + '",A' + $r + ')'
    $a = 打つ $式
    [void]$結果.Add(("CELL`t{0}`t{1}`t{2}`t表示形式 {3}" -f $式, $a[0], $a[1], $f))
  }
  $r++
}

# ── ②そろえ方ごとの prefix ────────────────────────────────
$そろえ = @{ '既定' = -4130; '左' = -4131; '中央' = -4108; '右' = -4152; '繰り返し' = 5; '両端' = -4130 }
$行 = 100
foreach ($名 in @('既定', '左', '中央', '右', '繰り返し')) {
  foreach ($中 in @('字', '数')) {
    $c = $ws.Cells.Item($行, 2)
    if ($中 -eq '字') { $c.Value2 = 'abc' } else { $c.Value2 = 12 }
    $c.HorizontalAlignment = $そろえ[$名]
    $式 = '=CELL("prefix",B' + $行 + ')'
    $a = 打つ $式
    [void]$結果.Add(("CELL`t{0}`t{1}`t{2}`tそろえ {3}／中身 {4}" -f $式, $a[0], $a[1], $名, $中))
    $行++
  }
}

# ── ③type / contents の 変わり種 ────────────────────────────
$変 = @(
  @('C1', '=""',        '式が 空の 字を 返す'),
  @('C2', '=1/0',       '式が 誤りに なる'),
  @('C3', '=TRUE',      '真偽'),
  @('C4', "='abc'",     '（打てないはず）'),
  @('C5', $null,        '本当に 空'),
  @('C6', '=A1',        '空でない セルを 指す 式')
)
foreach ($v in $変) {
  $場 = $v[0]
  try {
    if ($null -eq $v[1]) { $ws.Range($場).Value2 = $null } else { $ws.Range($場).Formula = $v[1] }
  } catch { }
  foreach ($k in @('type', 'contents')) {
    $式 = '=CELL("' + $k + '",' + $場 + ')'
    $a = 打つ $式
    [void]$結果.Add(("CELL`t{0}`t{1}`t{2}`t{3}（{4}）" -f $式, $a[0], $a[1], $v[2], $v[1]))
  }
}

# ── ④引数が 1つだけ（＝今 いる セル）／範囲を 渡す ─────────────
$ws.Range('E5').Value2 = 7
foreach ($式 in @('=CELL("row")', '=CELL("col")', '=CELL("address")',
                  '=CELL("address",A1:C3)', '=CELL("row",A1:C3)', '=CELL("contents",A1:C3)',
                  '=CELL("type",Sheet1!A1)', '=CELL("ADDRESS",A1)', '=CELL("Address",A1)')) {
  $a = 打つ $式
  [void]$結果.Add(("CELL`t{0}`t{1}`t{2}`t形の 変わり種" -f $式, $a[0], $a[1]))
}

$wb.Close($false); $xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

$頭 = @(
  "# ★実Excel に 打たせた CELL の 答え（2回目・見た目を 細かく）★",
  "# 取った 日 … 2026-09-07",
  "# ★どの Excel で 打ったか★ … 版 $版 ／ build $ビルド ／ UI の 言語 $言語",
  "# ①A1〜 に 1234.5678 を 置いて 表示形式を 1つずつ 変えた（式は 英語の 記号で 入れた）",
  "# ②B100〜 に 字/数を 置いて そろえ方を 変えた",
  "# ③C1〜C6 は type/contents の 変わり種",
  "# 関数`t式`t実Excel の 答え`t型`tどんな 場面か"
)
[System.IO.File]::WriteAllText($出, ((($頭 + $結果) -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding $false))
Write-Host "★書いた … $出（$($結果.Count) 本）★"
Write-Host "★Excel … 版 $版 ／ build $ビルド ／ UI $言語★"
