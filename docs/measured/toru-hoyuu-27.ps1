# toru-hoyuu-27.ps1 — ★保留 27行を ★元の 紙の 土台★で 取り直す★（2026-09-08）
#
#  ★★なぜ 保留に なったか★★
#    `toru-zero-torinaosu.ps1` は ★1つの 土台★で 697行を 押した。
#    ⇒ 紙ごとに 土台が 違うので、27行は ★窓①が 元の 紙と 違う 値★に なった
#    ⇒★『見せかけの 0』か どうかを 判じられない＝保留★に した
#    ⇒★この 紙で ★元の 紙の 土台★で 取り直す★
#
#  ★★27行の 中身（★全部 マスを 参照する 式★）★★
#    golden-346（20行）……… =BESSELK(D1,D2) / =CHIDIST(D1,2) / =PHI(D1) など
#      ⇒ 土台 … ★D1・D2 に 何が 入って いるか★で 答えが 変わる
#    golden-cell（4行）…… =CELL("contents",A4) など
#    golden-hikizan-zero（2行）… =A1+A2-0.3 / =0.1+0.2-A3
#    golden-cell4（1行）… =CELL("width",D1) / =CELL("protect",E1)
#
#  ★★元の 紙の 土台（★それぞれの 取り方の 頭から 写した★）★★
#    golden-346          … A1:A5=1..5 ／ B1:B5=2,4,6,8,10 ／ ★D1=2024/1/1 D2=2026/1/1★
#      ⇒★1回目は 私が 土台を ★書き写し間違えた★（D1=0.5 D2=1 と 書いた）★
#        ⇒ 20行が「元の 紙と 違う 値」に なった＝★取り直しに 失敗して いた★
#        ⇒★元の 取り方 `kansuu46/toru-346.ps1` の 106〜111行目から 写し直した★
#    golden-hikizan-zero … ★A1=0.1 A2=0.2 A3=0.3★
#    golden-cell 系       … 紙ごとに 違う（★この 紙では 346 と hikizan-zero だけ 取り直す★）
#  ⇒★★CELL 系 4行は ★別の 土台★が 要るので この 紙では 取り直しません★★
#    ⇒★「取り直して いない」と はっきり 書く（★測って いない を 緑に しない★）★
#
#  ★司さんの 実物には 触りません★＝新しい ブック・保存せず
#
#  使い方: powershell -File docs/measured/toru-hoyuu-27.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-hoyuu-27-2026-09-08.tsv'

$誤りの番号 = @{
  -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!';
  -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!';
  -2146826246 = '#N/A';   -2146826243 = '#SPILL!'; -2146826238 = '#CALC!'; -2146826237 = '#BUSY!'
}

# ★保留の 行を 読む★
$入 = Join-Path $ここ 'golden-zero-torinaoshi-2026-09-08.tsv'
if (-not (Test-Path $入)) { Write-Error '★取り直しの 紙が 無い★'; exit 2 }
$保留 = @()
foreach ($l in (Get-Content $入 -Encoding UTF8)) {
  if ($l.StartsWith('#')) { continue }
  $p = $l -split "`t"
  if ($p.Count -lt 7) { continue }
  if ($p[6] -like '*土台が 違う*') { $保留 += @{ 紙 = $p[0]; 式 = $p[1] } }
}
Write-Host ('★保留 … ' + $保留.Count + '行★')

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Worksheets.Item(1)

  function 生([string]$式) {
    $sh.Range('BZ1:CF5').Clear() | Out-Null
    try { $sh.Range('BZ1').Formula = $式; return $sh.Range('BZ1').Value2 } catch { return '★受け付けない★' }
  }
  function 型に($v) {
    if ($null -eq $v) { return 'Empty' }
    if ($v -is [string]) { return 'String' }
    if ($v -is [bool]) { return 'Boolean' }
    if ($v -is [double] -or $v -is [int] -or $v -is [long]) { return 'Number' }
    return 'Other'
  }
  function 字に($v) {
    if ($null -eq $v) { return '(空)' }
    if (($v -is [int] -or $v -is [long]) -and $誤りの番号.ContainsKey([int]$v)) { return $誤りの番号[[int]$v] }
    if ($v -is [bool]) { return $(if ($v) { 'TRUE' } else { 'FALSE' }) }
    if ($v -is [double]) { return $v.ToString('R') }
    return [string]$v
  }

  # ★土台を 敷く 関数（★元の 紙の 取り方の 頭から 写した★）★
  function 土台_346 {
    $sh.Cells.Clear() | Out-Null
    for ($r = 1; $r -le 5; $r++) {
      $sh.Range('A' + $r).Value2 = [double]$r
      $sh.Range('B' + $r).Value2 = [double]($r * 2)
    }
    # ★元の 取り方 kansuu46/toru-346.ps1 の 110〜111行目と ★同じ 書き方★★
    $sh.Range('D1').Formula = '=DATE(2024,1,1)'
    $sh.Range('D2').Formula = '=DATE(2026,1,1)'
  }
  function 土台_hikizan {
    $sh.Cells.Clear() | Out-Null
    $sh.Range('A1').Value2 = [double]0.1
    $sh.Range('A2').Value2 = [double]0.2
    $sh.Range('A3').Value2 = [double]0.3
  }

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★保留 27行を ★元の 紙の 土台★で 取り直した★（2026-09-08）')
  $行.Add('#')
  $行.Add('# ★なぜ 保留だったか★')
  $行.Add('#   `toru-zero-torinaosu.ps1` は ★1つの 土台★で 697行を 押した')
  $行.Add('#   ⇒ 紙ごとに 土台が 違うので 27行は ★窓①が 元の 紙と 違う 値★に なった')
  $行.Add('#   ⇒★判じられない＝保留★に した')
  $行.Add('#')
  $行.Add('# ★元の 紙の 土台（それぞれの 取り方の 頭から 写した）★')
  $行.Add('#   golden-346          … A1:A5=1..5 ／ B1:B5=2,4,6,8,10 ／ ★D1=2024/1/1 D2=2026/1/1★')
  $行.Add('#   ★1回目は 私が 土台を 書き写し間違え（D1=0.5 D2=1）、20行の 取り直しに 失敗した★')
  $行.Add('#   ⇒★元の 取り方 kansuu46/toru-346.ps1 の 106〜111行目から 写し直した★')
  $行.Add('#   golden-hikizan-zero … ★A1=0.1 A2=0.2 A3=0.3★')
  $行.Add('#')
  $行.Add('# ★★CELL 系 4行は この 紙では 取り直して いません★★')
  $行.Add('#   （紙ごとに 土台が 違い、`=CELL("width",D1)` は ★列の 幅★まで 要る）')
  $行.Add('#   ⇒★「取り直して いない」と はっきり 書く＝★測って いない を 緑に しない★')
  $行.Add('#')
  $行.Add('# ★3つの 窓＋型★ 窓①.Value2 ／ ★型★ ／ 窓②=(式)=0 ／ 窓③(式)*1e17')
  $行.Add('# 元の紙' + "`t" + '式' + "`t" + '窓①.Value2' + "`t" + '型' + "`t" + '窓②=0の真偽' + "`t" + '窓③*1e17' + "`t" + '判じ')

  $本数 = 0; $本当 = 0; $見せかけ = 0; $字 = 0; $未 = 0; $違 = 0; $判ぜず = 0
  foreach ($x in $保留) {
    if ($x.紙 -like 'golden-346*') { 土台_346 }
    elseif ($x.紙 -like 'golden-hikizan-zero*') { 土台_hikizan }
    else {
      # ★CELL 系＝取り直さない★
      $行.Add($x.紙 + "`t" + $x.式 + "`t" + '—' + "`t" + '—' + "`t" + '—' + "`t" + '—' + "`t" + '★取り直して いない（土台が 別）★')
      $本数++; $未++
      continue
    }
    $中 = $x.式 -replace '^=\s*', ''
    $生v = 生 $x.式
    $v1 = 字に $生v
    $型 = 型に $生v
    $v2 = 字に (生 ('=(' + $中 + ')=0'))
    $v3 = 字に (生 ('=(' + $中 + ')*1e17'))
    $判 = ''
    if ($v1 -ne '0') { $判 = '★元の 紙と 違う 値（' + $v1 + '）＝★保留の まま（土台が まだ 違う）★'; $違++ }
    elseif ($型 -eq 'String') { $判 = '★字の "0"（数では ない）★'; $字++ }
    elseif ($v2 -eq 'TRUE') { $判 = '本当に 0'; $本当++ }
    elseif ($v2 -eq 'FALSE') { $判 = '★★見せかけの 0★★'; $見せかけ++ }
    else { $判 = '★判じられない（窓②が 真偽で 返らない＝' + $v2 + '）★'; $判ぜず++ }
    $行.Add($x.紙 + "`t" + $x.式 + "`t" + $v1 + "`t" + $型 + "`t" + $v2 + "`t" + $v3 + "`t" + $判)
    $本数++
    Write-Host ('  ' + $x.式.PadRight(30) + ' 窓①=' + $v1.PadRight(10) + ' 型=' + $型.PadRight(8) + ' 窓②=' + $v2.PadRight(6) + ' → ' + $判)
  }

  [System.IO.File]::WriteAllLines($出, $行, (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★★取り直した … ' + $本数 + '行★★')
  Write-Host ('  本当に 0 …………………… ' + $本当 + '行')
  Write-Host ('  字の "0" ………………… ' + $字 + '行')
  Write-Host ('  ★★見せかけの 0 ………… ' + $見せかけ + '行★★')
  Write-Host ('  ★取り直して いない（CELL 系）… ' + $未 + '行★')
  Write-Host ('  ★判じられない（窓②が 真偽で 返らない）… ' + $判ぜず + '行★')
  Write-Host ('  ★★まだ 土台が 違う（取り直し 失敗）… ' + $違 + '行★★')
  $足し = $本当 + $字 + $見せかけ + $未 + $判ぜず + $違
  Write-Host ('  ―― 足すと ' + $足し + ' ／ 全 ' + $本数 + '行')
  if ($足し -ne $本数) { Write-Error '★数が 合わない＝どこかの 行を 数え落として いる★'; exit 3 }
  if ($違 -gt 0) { Write-Host '  ⇒★0行に なるまで 土台を 直す。0で ない まま 緑に しない★' }
  Write-Host ('★書いた … ' + $出 + '★')

  $bk.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
