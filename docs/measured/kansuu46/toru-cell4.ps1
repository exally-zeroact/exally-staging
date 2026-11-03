# toru-cell4.ps1 — ★CELL の 残りの 問い 4つ★（2026-09-07）
#
#  ★まだ 決められない 事を 名指しで 聞く★
#    ①別の シートを 指した 時の `address` は どんな 形か
#    ②列を 隠した 時の `width` は 何を 返すか
#    ③鍵を 外した セル（ロックを 解除）の `protect` は 0に なるか
#    ④★保存した 後の `filename` は どんな 形か★
#      ⇒ うちは ★ファイルの 置き場を 持たない★ので 何を 返すか 決める 材料に する
#
#  使い方: powershell -File docs/measured/kansuu46/toru-cell4.ps1

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
$出 = Join-Path $ここ 'golden-cell4-2026-09-07.tsv'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$版 = $xl.Version; $ビルド = $xl.Build
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws2 = $wb.Worksheets.Add()
$ws2.Name = '二枚目'
$ws.Activate()

$結果 = New-Object System.Collections.ArrayList
function 打つ($式) {
  $ws.Range('H1').Formula = $式
  $v = $ws.Range('H1').Value2
  if ($null -eq $v) { return '(空)' }
  if ($v -is [double]) { return $v.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture) }
  return [string]$v
}
function 足す($式, $場面) { [void]$結果.Add(("CELL`t{0}`t{1}`t{2}" -f $式, (打つ $式), $場面)) }

# ①別の シート
$ws2.Range('B3').Value2 = 9
足す '=CELL("address",二枚目!B3)' '別の シートを 指す'
足す '=CELL("row",二枚目!B3)'     '別の シートを 指す'
足す '=CELL("col",二枚目!B3)'     '別の シートを 指す'
足す '=CELL("contents",二枚目!B3)' '別の シートを 指す'

# ②隠した 列
$ws.Range('D1').Value2 = 1
$ws.Range('D:D').EntireColumn.Hidden = $true
足す '=CELL("width",D1)' '列を 隠した'
$ws.Range('D:D').EntireColumn.Hidden = $false
足す '=CELL("width",D1)' '隠すのを やめた'

# ③鍵を 外す
$ws.Range('E1').Value2 = 1
$ws.Range('E1').Locked = $false
足す '=CELL("protect",E1)' '鍵を 外した セル（シートは 保護していない）'

# ④保存した 後の filename
足す '=CELL("filename",A1)' '保存する 前'
$仮 = Join-Path $env:TEMP ('cell-filename-' + [guid]::NewGuid().ToString('N').Substring(0,8) + '.xlsx')
$wb.SaveAs($仮, 51)
足す '=CELL("filename",A1)' '保存した 後'
足す '=CELL("filename",二枚目!B3)' '保存した 後・別の シート'

$wb.Close($false); $xl.Quit()
foreach ($o in @($ws2, $ws, $wb, $xl)) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null }
try { Remove-Item $仮 -Force } catch { }

$頭 = @(
  "# ★実Excel に 打たせた CELL の 答え（4回目・残りの 問い 4つ）★",
  "# 取った 日 … 2026-09-07",
  "# ★どの Excel で 打ったか★ … 版 $版 ／ build $ビルド",
  "# ★保存した 先は 仮の 場所★（測り終わったら 消している）",
  "# 関数`t式`t実Excel の 答え`tどんな 場面か"
)
[System.IO.File]::WriteAllText($出, ((($頭 + $結果) -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding $false))
Write-Host "★書いた … $出（$($結果.Count) 本）★"
