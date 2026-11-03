# toru-areas2.ps1 — ★AREAS に「範囲で ない 物」が 来たら どうなるか★（2026-09-07）
#
#  ★1回目で 測れなかった 所★
#    `=AREAS("A1")` `=AREAS(5)` `=AREAS(TRUE)` `=AREAS("")` `=AREAS(1/0)` は
#    ★実Excel が 式として 受け取りません★（0x800A03EC＝入力を 断られる）
#    ⇒★答えが 「無い」のでは なく ★聞けていない★★
#    ⇒★★ここで「#VALUE! だろう」と 決めたら それは 実測では なく 当て推量★★
#
#  ★そこで 別の 道で 聞く★
#    ★Excel が 受け取る 形★で、中身だけ 範囲で なく する。
#      `=AREAS(IF(TRUE,"x","y"))` … 関数が 字を 返す
#      `=AREAS(INDIRECT("zzz"))`  … 範囲を 指せない 字
#      `=AREAS(CHOOSE(1,"x","y"))`
#      `=AREAS(A1:B2 D5:E6)`      … 重ならない 交わり（#NULL! が 中に 入る）
#    ⇒★これなら 実Excel が 何を 返すか 見える★
#
#  ★もらった ファイルから 来た 式も 見る★
#    お客さんは ★他所で 作った .xlsx★を 開く。
#    その中に `=AREAS("A1")` が 入っていたら 実Excel は 何を 出すか。
#    ⇒ ここでは ★式を 字として セルに 置いてから 式に 変える★道で 試す。
#
#  使い方: powershell -File docs/measured/kansuu46/toru-areas2.ps1

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
$出 = Join-Path $ここ 'golden-areas2-2026-09-07.tsv'

$誤りの番号 = @{
  -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!';
  -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!';
  -2146826246 = '#N/A'
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$版 = $xl.Version; $ビルド = $xl.Build
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
for ($r = 1; $r -le 9; $r++) { for ($c = 1; $c -le 9; $c++) { $ws.Cells.Item($r, $c).Value2 = $r * 10 + $c } }

$式たち = @(
  '=AREAS(IF(TRUE,"x","y"))',
  '=AREAS(IF(TRUE,A1,B1))',
  '=AREAS(CHOOSE(1,"x","y"))',
  '=AREAS(INDIRECT("zzz"))',
  '=AREAS(INDIRECT("A1")&"")',
  '=AREAS(A1:B2 D5:E6)',
  '=AREAS((A1:B2 D5:E6,A1))',
  '=AREAS(OFFSET(A1,-5,0))',
  '=AREAS(INDEX(A1:C3,99,1))',
  '=AREAS(N(A1))',
  '=AREAS(T(A1))',
  '=AREAS(A1:A2^1)',
  '=AREAS(SUM(A1:A2))',
  '=AREAS(A1#)'
)

$結果 = New-Object System.Collections.ArrayList
function 足す($式, $答, $型, $道) {
  [void]$結果.Add(("AREAS`t{0}`t{1}`t{2}`t{3}" -f $式, $答, $型, $道))
}
function 読む() {
  $v = $ws.Range('K1').Value2
  if ($null -eq $v) { return @('(空)', 'null') }
  $t = $v.GetType().Name
  if ($v -is [double]) { return @($v.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture), $t) }
  if ($v -is [int] -and $誤りの番号.ContainsKey([int]$v)) { return @($誤りの番号[[int]$v], 'error値') }
  return @([string]$v, $t)
}

# ── ①ふつうに 打つ ──────────────────────────────
foreach ($式 in $式たち) {
  try {
    $ws.Range('K1').Formula = $式
    $a = 読む
    足す $式 $a[0] $a[1] 'ふつうに 打った'
  } catch {
    足す $式 '★Excel が 式として 受け取らない★' 'error' 'ふつうに 打った'
  }
}

# ── ②★もらった ファイルの 式★の 道 ────────────────
#    ★字として 置いてから 式に 変える★＝人が 打つ 時の 見張りを 通らない道
$別の道 = @('=AREAS("A1")', '=AREAS(5)', '=AREAS(TRUE)', '=AREAS("")', '=AREAS(1/0)', '=AREAS(A1&"")')
foreach ($式 in $別の道) {
  $入った = $true
  try {
    # 先に 字として 置く（頭に ' を 付けると 式に ならない）
    $ws.Range('K1').Value2 = "'" + $式
    # ★式に 変える★＝Formula に 入れ直す（ここで 断られるか どうか）
    $ws.Range('K1').Formula = $式
  } catch { $入った = $false }
  if ($入った) {
    $a = 読む
    足す $式 $a[0] $a[1] '字から 式に 変えた'
  } else {
    足す $式 '★Excel が 式として 受け取らない★' 'error' '字から 式に 変えた'
  }
}

# ── ③★.xlsx に 書いて 開き直す★＝お客さんが もらった ファイルと 同じ 道 ──
$仮 = Join-Path $env:TEMP ('areas-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.xlsx')
try {
  # ★式として 入らない 物は そもそも 書けない★ので、
  #   ★入る 物だけ★を 書いて 開き直し、★開いた 後の 答え★を 見る
  $ws.Range('M1').Formula = '=AREAS((A1,A2))'
  $ws.Range('M2').Formula = '=AREAS(A1:B2 D5:E6)'
  $wb.SaveAs($仮, 51)
  $wb.Close($false)
  $wb2 = $xl.Workbooks.Open($仮)
  $ws3 = $wb2.Worksheets.Item(1)
  foreach ($場 in @('M1', 'M2')) {
    $v = $ws3.Range($場).Value2
    $答 = if ($null -eq $v) { '(空)' }
          elseif ($v -is [double]) { $v.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture) }
          elseif ($v -is [int] -and $誤りの番号.ContainsKey([int]$v)) { $誤りの番号[[int]$v] }
          else { [string]$v }
    足す ([string]$ws3.Range($場).Formula) $答 '-' '★ファイルに 書いて 開き直した★'
  }
  $wb2.Close($false)
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws3) | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb2) | Out-Null
} catch {
  足す '（ファイルの 道）' ('★測れなかった★ ' + ($_.Exception.Message -replace "`r?`n", ' ')) 'error' '★ファイルに 書いて 開き直した★'
}

$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
try { Remove-Item $仮 -Force } catch { }

$頭 = @(
  "# ★実Excel に 打たせた AREAS の 答え（2回目・範囲で ない 物）★",
  "# 取った 日 … 2026-09-07",
  "# ★どの Excel で 打ったか★ … 版 $版 ／ build $ビルド",
  "# ★1回目で `=AREAS(""A1"")` などが 打てなかった★ので、",
  "#   ★Excel が 受け取る 形★（関数ごしに 字を 渡す 等）で 聞き直した。",
  "# ★置いた 中身★ A1:I9 に 数（行*10+列）",
  "# 関数`t式`t実Excel の 答え`t型`tどの 道で 入れたか"
)
[System.IO.File]::WriteAllText($出, ((($頭 + $結果) -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding $false))
Write-Host "★書いた … $出（$($結果.Count) 本）★"
