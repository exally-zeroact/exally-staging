# toru-areas.ps1 — ★AREAS が 何を 返すかを 実Excel に 聞く★（2026-09-07）
#
#  ★AREAS は「範囲が 幾つに 分かれているか」を 数える 関数★
#    `=AREAS(B2:D4)` … 1 ／ `=AREAS((B2:D4,E5,F6:I9))` … 3 のはず
#
#  ★★先に 分かっている 事（自分で 押して 確かめた）★★
#    うちの エンジン（HyperFormula）は
#    ★とびとびの 範囲を ★式を 読む 所★で 断る★
#      `=SUM((A1:B3,D1:D2))` … Parsing error（RParen を 待っていたら , が 来た）
#      `=AREAS((A1:B3,D1:D2))` … 同じ
#    ⇒★関数の 所まで 届かない★＝出すなら ★JS層（式の 字を 見る 段）★
#    ⇒ だから ★字の 形ごとに 実Excel が 何を 返すか★を 細かく 取る
#
#  ★聞く事★
#    ①ふつうの 範囲／1マス／列ぜんぶ／行ぜんぶ
#    ②とびとび（union の カンマ）… 2つ・3つ・4つ
#    ③重なり（intersection の 空白）
#    ④名前を 付けた 範囲（1つの 物・とびとびの 物）
#    ⑤別の シート
#    ⑥関数が 返す 範囲（INDEX／OFFSET／CHOOSE）
#    ⑦断るはずの 物（字・数・空）
#
#  使い方: powershell -File docs/measured/kansuu46/toru-areas.ps1

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
$出 = Join-Path $ここ 'golden-areas-2026-09-07.tsv'

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
$ws2 = $wb.Worksheets.Add()
$ws2.Name = '二枚目'
$ws.Activate()

# ★中身を 置く★（INDEX や OFFSET が 何かを 指せる ように）
for ($r = 1; $r -le 9; $r++) { for ($c = 1; $c -le 9; $c++) { $ws.Cells.Item($r, $c).Value2 = $r * 10 + $c } }
$ws2.Range('A1').Value2 = 7

# ★名前を 付ける★（1つの 物／とびとびの 物）
$wb.Names.Add('ひとつ', '=Sheet1!$B$2:$D$4') | Out-Null
$wb.Names.Add('とびとび', '=Sheet1!$B$2:$D$4,Sheet1!$F$6:$I$9') | Out-Null

$式たち = @(
  # ①ふつう
  '=AREAS(B2:D4)', '=AREAS(A1)', '=AREAS(A:A)', '=AREAS(1:1)', '=AREAS(A:C)', '=AREAS(1:3)',
  # ②とびとび
  '=AREAS((B2:D4,E5))', '=AREAS((B2:D4,E5,F6:I9))', '=AREAS((A1,A2,A3,A4))',
  '=AREAS((A1:A2,A2:A3))', '=AREAS((A:A,C:C))', '=AREAS((A1,A1))',
  # ③重なり（空白）
  '=AREAS(B2:D4 C3:E5)', '=AREAS(B2:D4 A1)', '=AREAS((B2:D4 C3:E5,A1))',
  # ④名前
  '=AREAS(ひとつ)', '=AREAS(とびとび)', '=AREAS((ひとつ,A1))',
  # ⑤別の シート
  '=AREAS(二枚目!A1:B2)', '=AREAS((Sheet1!A1,二枚目!A1))', '=AREAS((二枚目!A1,二枚目!C3))',
  # ⑥関数が 返す 範囲
  '=AREAS(INDEX(A1:C3,1,1))', '=AREAS(INDEX(A1:C3,,1))', '=AREAS(OFFSET(A1,0,0,2,2))',
  '=AREAS(CHOOSE(1,A1:B2,C1:D2))', '=AREAS(CHOOSE(2,A1:B2,C1:D2))',
  '=AREAS(INDIRECT("A1:B2"))', '=AREAS(A1:INDEX(A1:C3,2,1))',
  # ⑦断るはず
  '=AREAS("A1")', '=AREAS(5)', '=AREAS(TRUE)', '=AREAS("")', '=AREAS(1/0)',
  # ⑧入れ子（他の 関数の 中）
  '=SUM(AREAS((A1,A2)),1)', '=IF(AREAS((A1,A2))=2,"ふたつ","ちがう")',
  '=AREAS((A1,A2))+AREAS(B2:D4)'
)

$結果 = New-Object System.Collections.ArrayList
foreach ($式 in $式たち) {
  $答 = ''; $型 = ''
  try {
    $ws.Range('K1').Formula = $式
    $v = $ws.Range('K1').Value2
    if ($null -eq $v) { $答 = '(空)'; $型 = 'null' }
    else {
      $型 = $v.GetType().Name
      if ($v -is [double]) { $答 = $v.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture) }
      elseif ($v -is [int] -and $誤りの番号.ContainsKey([int]$v)) { $答 = $誤りの番号[[int]$v]; $型 = 'error値' }
      else { $答 = [string]$v }
    }
  } catch {
    $答 = '★打てない★ ' + ($_.Exception.Message -replace "`r?`n", ' ' -replace "`t", ' ')
    $型 = 'error'
  }
  [void]$結果.Add(("AREAS`t{0}`t{1}`t{2}" -f $式, $答, $型))
}

$wb.Close($false); $xl.Quit()
foreach ($o in @($ws2, $ws, $wb, $xl)) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null }

$頭 = @(
  "# ★実Excel に 打たせた AREAS の 答え★",
  "# 取った 日 … 2026-09-07",
  "# ★どの Excel で 打ったか★ … 版 $版 ／ build $ビルド",
  "# ★置いた 中身★ Sheet1 の A1:I9 に 数（行*10+列）／二枚目!A1=7",
  "# ★名前★ ひとつ = Sheet1!`$B`$2:`$D`$4 ／ とびとび = Sheet1!`$B`$2:`$D`$4,Sheet1!`$F`$6:`$I`$9",
  "# ★式は K1 に 打った★（数を 置いた 所と 重ならない 場所）",
  "# 関数`t式`t実Excel の 答え`t型"
)
[System.IO.File]::WriteAllText($出, ((($頭 + $結果) -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding $false))
Write-Host "★書いた … $出（$($結果.Count) 本）★"
Write-Host "★Excel … 版 $版 ／ build $ビルド★"
