# toru-cell.ps1 — ★CELL が 何を 返すかを 実Excel に 聞く★（2026-09-07）
#
#  ★普通の 取り方では 足りない★
#    CELL は ★そのセルの 中身や 見た目★を 返すので、
#    ★空の シートに 式だけ 置いても 何も 分かりません★。
#  ⇒ だから ここでは
#      ①A1:D8 に ★決めた 中身★を 置く（数・字・日付・空・式）
#      ②B列の 幅を 変える／1マスを 保護しない／表示形式を 付ける
#      ③その 上で CELL を 打たせる
#
#  使い方: powershell -File docs/measured/kansuu46/toru-cell.ps1

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
$出 = Join-Path $ここ 'golden-cell-2026-09-07.tsv'

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

# ★中身を 置く★（何を 置いたかは 紙の 頭に 書く）
$ws.Range('A1').Value2 = 123.456          # 数
$ws.Range('A2').Value2 = 'あいう'          # 字
$ws.Range('A3').Formula = '=1+2'          # 式
$ws.Range('A4').Value2 = $null            # 空
$ws.Range('B1').Value2 = -5               # 負の 数
$ws.Range('B2').Value2 = 'x'
$ws.Range('C1').Formula = '=DATE(2024,1,2)'
$ws.Range('C1').NumberFormatLocal = 'yyyy/m/d'
$ws.Range('D1').Value2 = 1000
$ws.Range('D1').NumberFormatLocal = '#,##0'
$ws.Range('B:B').ColumnWidth = 12         # ★幅を 変える★
$ws.Range('A2').HorizontalAlignment = -4152   # 右そろえ（prefix が 変わるか）

$種 = @('address','col','row','contents','type','width','prefix','protect','format',
        'color','parentheses','filename','こんなの')
$場 = @('A1','A2','A3','A4','B1','C1','D1')

$結果 = New-Object System.Collections.ArrayList
foreach ($k in $種) {
  foreach ($c in $場) {
    $式 = '=CELL("' + $k + '",' + $c + ')'
    $答 = ''; $型 = ''
    try {
      $ws.Range('F1').Formula = $式
      $v = $ws.Range('F1').Value2
      if ($null -eq $v) { $答 = '(空)'; $型 = 'null' }
      else {
        $型 = $v.GetType().Name
        if ($v -is [double]) { $答 = $v.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture) }
        elseif ($v -is [int] -and $誤りの番号.ContainsKey([int]$v)) { $答 = $誤りの番号[[int]$v]; $型 = 'error値' }
        else { $答 = [string]$v }
      }
    } catch {
      $答 = '★打てない★ ' + ($_.Exception.Message -replace "`r?`n", ' '); $型 = 'error'
    }
    [void]$結果.Add(("CELL`t{0}`t{1}`t{2}" -f $式, $答, $型))
  }
}

$wb.Close($false); $xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

$頭 = @(
  "# ★実Excel に 打たせた CELL の 答え★",
  "# 取った 日 … 2026-09-07",
  "# ★どの Excel で 打ったか★ … 版 $版 ／ build $ビルド ／ UI の 言語 $言語",
  "# ★置いた 中身★ A1=123.456 / A2='あいう'(右そろえ) / A3='=1+2' / A4=空 /",
  "#             B1=-5 / B2='x' / C1=DATE(2024,1,2) 表示 yyyy/m/d / D1=1000 表示 #,##0",
  "#             B列の 幅=12",
  "# 関数`t式`t実Excel の 答え`t型"
)
[System.IO.File]::WriteAllText($出, ((($頭 + $結果) -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding $false))
Write-Host "★書いた … $出（$($結果.Count) 本）★"
Write-Host "★Excel … 版 $版 ／ build $ビルド ／ UI $言語★"
