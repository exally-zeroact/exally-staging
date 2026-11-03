# toru-cell3.ps1 — ★CELL の「色」「かっこ」「円」「うちが 出す 表示形式」を 聞く★（2026-09-07）
#
#  ★2回目で 測れていなかった 所★（2回目の 紙に「付かない」と 残っている）
#    ・`[Red]` は ★日本語の Excel では 付かない★（`[赤]` と 書く）
#      ⇒★だから 2回目は color が ★1度も 1に ならなかった★＝物差しが 空洞★
#    ・`$` は 日本語の Excel では ★ただの 字★（C0/C2 に ならず ,0/,2 に なった）
#      ⇒★円（¥）で 測り直す★
#    ・`parentheses` は ★正の 数に かっこが 付く 形★でないと 1に ならないはず
#    ・★うちの 画面が 出す 表示形式★（yyyy年m月d日・h時mm分 など）は 未測定
#
#  ★日本語の 記号なので `.NumberFormatLocal` で 入れる★（英語の `.NumberFormat` では 付かない）
#
#  使い方: powershell -File docs/measured/kansuu46/toru-cell3.ps1

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
$出 = Join-Path $ここ 'golden-cell3-2026-09-07.tsv'

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
  if ($null -eq $v) { return '(空)' }
  if ($v -is [double]) { return $v.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture) }
  return [string]$v
}

$形たち = @(
  'G/標準',
  '\¥#,##0', '\¥#,##0.00',
  '\¥#,##0;[赤]-\¥#,##0', '\¥#,##0_);[赤](\¥#,##0)',
  '0;[赤]0', '0.00;[赤]0.00',
  '#,##0;[赤]#,##0', '#,##0;[赤](#,##0)',
  '(#,##0);(#,##0)', '(0);(0)', '(#,##0);[赤](#,##0)',
  '0_);(0)',
  'yyyy/m/d', 'yyyy/mm/dd', 'yyyy-m-d', 'yyyy-mm-dd',
  'yyyy年m月d日', 'yyyy年mm月dd日', 'm月d日', 'm月d日(aaa)', 'm/d', 'mm/dd',
  'h:mm', 'hh:mm', 'h時mm分', 'hh時mm分', 'h:mm AM/PM', 'a h:mm',
  '0.0%', '0.000%', '#,##0.000', '0.00E+00', '@'
)
$r = 1
foreach ($f in $形たち) {
  $c = $ws.Cells.Item($r, 1)
  $c.Value2 = 1234.5678
  $付いた = $true
  try { $c.NumberFormatLocal = $f } catch {
    $付いた = $false
    [void]$結果.Add(("CELL`t(表示形式を 付けられない)`t★付かない★`t表示形式 {0}" -f $f))
  }
  if ($付いた) {
    $実 = [string]$c.NumberFormatLocal
    foreach ($k in @('format', 'color', 'parentheses')) {
      $式 = '=CELL("' + $k + '",A' + $r + ')'
      [void]$結果.Add(("CELL`t{0}`t{1}`t表示形式 {2}（実際に 付いた 形 {3}）" -f $式, (打つ $式), $f, $実))
    }
  }
  $r++
}

$wb.Close($false); $xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

$頭 = @(
  "# ★実Excel に 打たせた CELL の 答え（3回目・色／かっこ／円／うちの 表示形式）★",
  "# 取った 日 … 2026-09-07",
  "# ★どの Excel で 打ったか★ … 版 $版 ／ build $ビルド ／ UI の 言語 $言語",
  "# ★入れ方★ `.NumberFormatLocal`（日本語の 記号）／`\¥` は 円の 記号",
  "# 関数`t式`t実Excel の 答え`tどんな 場面か"
)
[System.IO.File]::WriteAllText($出, ((($頭 + $結果) -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding $false))
Write-Host "★書いた … $出（$($結果.Count) 本）★"
