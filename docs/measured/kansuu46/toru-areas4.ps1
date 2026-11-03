# toru-areas4.ps1 — ★AREAS で 残していた 穴を 埋める★（2026-09-07）
#
#  ★作る 前に 測る★
#    `AREAS-dou-dasu-ka.md` に「★未測定★」と 書いた 物が 在ります。
#      `AREAS((とびとび,A1))` … ★とびとびの 名前 ＋ もう1つ★
#    ⇒★『3 のはず』と 書いた＝★推し量り★★
#    ⇒★作る 前に 潰します★
#
#  ★ついでに 聞く（字だけで 数える 時に 迷う 形）★
#    ・空白が 前後に 入った 形 `( A1 , B2 )`
#    ・入れ子の かっこ `((A1,A2))`
#    ・とびとびの 中に とびとびの 名前が 2つ
#    ・`$` 付き `($A$1,$B$2)`
#    ・シート名に 空白が 入る `'二 枚目'!A1`
#    ・同じ 範囲を 3回
#    ・重なりと とびとびが 混ざる
#
#  使い方: powershell -File docs/measured/kansuu46/toru-areas4.ps1

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
$出 = Join-Path $ここ 'golden-areas4-2026-09-07.tsv'

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
$ws2.Name = '二 枚目'          # ★名前に 空白を 入れる★
$ws.Activate()
for ($r = 1; $r -le 9; $r++) { for ($c = 1; $c -le 9; $c++) { $ws.Cells.Item($r, $c).Value2 = $r * 10 + $c } }
$ws2.Range('A1').Value2 = 7

$wb.Names.Add('ひとつ', '=Sheet1!$B$2:$D$4') | Out-Null
$wb.Names.Add('とびとび', '=Sheet1!$B$2:$D$4,Sheet1!$F$6:$I$9') | Out-Null
$wb.Names.Add('みっつ', '=Sheet1!$A$1,Sheet1!$C$3,Sheet1!$E$5') | Out-Null

$式たち = @(
  # ★埋める 穴★
  '=AREAS((とびとび,A1))',
  '=AREAS((A1,とびとび))',
  '=AREAS((とびとび,とびとび))',
  '=AREAS(みっつ)',
  '=AREAS((みっつ,とびとび))',
  # ★字だけで 数える 時に 迷う 形★
  '=AREAS(( A1 , B2 ))',
  '=AREAS((A1,A2,A3))',
  '=AREAS(($A$1,$B$2))',
  '=AREAS((A1:B2,A1:B2,A1:B2))',
  '=AREAS(''二 枚目''!A1)',
  '=AREAS((''二 枚目''!A1,''二 枚目''!C3))',
  '=AREAS((A1:C3 B2:D4,E5))',
  '=AREAS((A1:C3,B2:D4 C3:E5))',
  '=AREAS((A:A,1:1))',
  '=AREAS((A1,B2,C3,D4,E5,F6,G7,H8))'
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
  } catch { $答 = '★Excel が 式として 受け取らない★'; $型 = 'error' }
  [void]$結果.Add(("AREAS`t{0}`t{1}`t{2}" -f $式, $答, $型))
}

$wb.Close($false); $xl.Quit()
foreach ($o in @($ws2, $ws, $wb, $xl)) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null }

$頭 = @(
  "# ★実Excel に 打たせた AREAS の 答え（4回目・残していた 穴）★",
  "# 取った 日 … 2026-09-07",
  "# ★どの Excel で 打ったか★ … 版 $版 ／ build $ビルド",
  "# ★名前★ ひとつ=1か所 ／ とびとび=2か所 ／ みっつ=3か所",
  "# ★シート★ Sheet1 ／ '二 枚目'（★名前に 空白★）",
  "# 関数`t式`t実Excel の 答え`t型"
)
[System.IO.File]::WriteAllText($出, ((($頭 + $結果) -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding $false))
Write-Host "★書いた … $出（$($結果.Count) 本）★"
