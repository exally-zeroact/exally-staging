# toru-areas3.ps1 — ★「字を 返す 関数」と「字の マスを 指す 関数」を 分けられるか★（2026-09-07）
#
#  ★2回目で 分かった 事★
#    `=AREAS(IF(TRUE,A1,B1))` … ★1★（マスを 指している）
#    `=AREAS(IF(TRUE,"x","y"))` … ★#VALUE!★（字そのもの）
#
#  ★★ここで 決めかけた 決まり★★
#    「★出てきた 値が 字なら #VALUE!★」
#  ★★でも それは ★どちらでも 合う 組★かも しれない★★
#    ⇒ 2回目の A1 には ★数★が 入っていた（11）。
#    ⇒★A1 に ★字★を 入れたら どうなるか★を 見ないと 決められない。
#      ・もし ★1★なら … 「値が 字なら #VALUE!」は ★間違い★
#      ・もし ★#VALUE!★なら … その 決まりで よい
#
#  ★答えが 割れる 組を わざと 作る★（物差しの 作り方 ①）
#
#  使い方: powershell -File docs/measured/kansuu46/toru-areas3.ps1

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
$出 = Join-Path $ここ 'golden-areas3-2026-09-07.tsv'

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

# ★A1 は 字／B1 は 数／C1 は 空／D1 は 誤り★（同じ 式で 中身だけ 変える）
$ws.Range('A1').Value2 = 'あいう'
$ws.Range('B1').Value2 = 11
$ws.Range('C1').Value2 = $null
$ws.Range('D1').Formula = '=1/0'
$ws.Range('A2').Value2 = 'かきく'

$式たち = @(
  # ★マスを 指す 関数（中身は 字）★ … 1 なら「値が 字」で 決めては いけない
  '=AREAS(IF(TRUE,A1,A2))',
  '=AREAS(IF(TRUE,A1:A2,B1))',
  '=AREAS(INDEX(A1:A2,1))',
  '=AREAS(CHOOSE(1,A1,A2))',
  '=AREAS(INDIRECT("A1"))',
  '=AREAS(OFFSET(A1,0,0))',
  # ★マスを 指す（中身は 空）★
  '=AREAS(IF(TRUE,C1,B1))',
  '=AREAS(INDEX(C1:C2,1))',
  # ★マスを 指す（中身は 誤り）★
  '=AREAS(IF(TRUE,D1,B1))',
  '=AREAS(INDEX(D1:D2,1))',
  # ★字そのものを 返す 関数★ … #VALUE! のはず
  '=AREAS(IF(TRUE,"x","y"))',
  '=AREAS(CHOOSE(1,"x","y"))',
  '=AREAS(IF(TRUE,A1&"",B1))',
  # ★数そのものを 返す 関数★
  '=AREAS(IF(TRUE,1,2))',
  '=AREAS(CHOOSE(1,1,2))',
  # ★真偽を 返す★
  '=AREAS(IF(TRUE,TRUE,FALSE))'
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
    $答 = '★Excel が 式として 受け取らない★'; $型 = 'error'
  }
  [void]$結果.Add(("AREAS`t{0}`t{1}`t{2}" -f $式, $答, $型))
}

$wb.Close($false); $xl.Quit()
foreach ($o in @($ws, $wb, $xl)) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null }

$頭 = @(
  "# ★実Excel に 打たせた AREAS の 答え（3回目・答えが 割れる 組）★",
  "# 取った 日 … 2026-09-07",
  "# ★どの Excel で 打ったか★ … 版 $版 ／ build $ビルド",
  "# ★置いた 中身★ A1='あいう'（字）／A2='かきく'（字）／B1=11（数）／C1=空／D1==1/0（誤り）",
  "# ★聞きたい 事★ 『出てきた 値が 字なら #VALUE!』で よいか",
  "#   ⇒ A1 が ★字★でも マスを 指していれば 1 なら、その 決まりは ★間違い★",
  "# 関数`t式`t実Excel の 答え`t型"
)
[System.IO.File]::WriteAllText($出, ((($頭 + $結果) -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding $false))
Write-Host "★書いた … $出（$($結果.Count) 本）★"
