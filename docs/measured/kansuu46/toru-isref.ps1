# toru-isref.ps1 — ★ISREF が 何を 返すかを 実Excel に 聞く★（2026-09-07）
#
#  ★なぜ 測るか★
#    `prompt/kansuu.md` の 13行の うち ★ISREF だけ★ 実測が 書いて ありませんでした
#    （指示役が 数えて 見つけた）。
#    ⇒ 1本 打ってみたら ★うちの 書き換えが 壊れていました★
#      `=ISREF(SUM(A1:A2))` → `=FALSE())` … ★かっこが 合わない★ ⇒ #ERROR!
#    ⇒★『実測が 書いていない 1行』が ★本当に 壊れていた★★
#
#  ★聞く事★
#    ①マスを 指す 物（A1／A1:B2／名前）… TRUE のはず
#    ②値そのもの（字・数・計算）……… FALSE のはず
#    ③★関数が 返す 物★
#       INDEX／OFFSET／INDIRECT は ★マスを 指す★ ⇒ TRUE のはず
#       SUM は ★値★ ⇒ FALSE のはず
#       IF は 中身しだい
#    ⇒★AREAS で 分かった「値では なく ★指しているか★」と 同じ 話かを 確かめる★
#
#  使い方: powershell -File docs/measured/kansuu46/toru-isref.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-isref-2026-09-07.tsv'

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
for ($r = 1; $r -le 5; $r++) { for ($c = 1; $c -le 5; $c++) { $ws.Cells.Item($r, $c).Value2 = $r * 10 + $c } }
$wb.Names.Add('なまえ', '=Sheet1!$A$1:$B$2') | Out-Null

$式たち = @(
  # ①マスを 指す
  '=ISREF(A1)', '=ISREF(A1:B2)', '=ISREF(A:A)', '=ISREF(なまえ)', '=ISREF(Sheet1!A1)',
  # ②値そのもの
  '=ISREF("a")', '=ISREF(1)', '=ISREF(TRUE)', '=ISREF(1+1)', '=ISREF(A1&"")',
  # ③関数が 返す 物（★ここが 肝★）
  '=ISREF(SUM(A1:A2))', '=ISREF(INDEX(A1:B2,1,1))', '=ISREF(INDEX(A1:B2,,1))',
  '=ISREF(OFFSET(A1,0,0))', '=ISREF(INDIRECT("A1"))', '=ISREF(INDIRECT("zzz"))',
  '=ISREF(IF(TRUE,A1,B1))', '=ISREF(IF(TRUE,"x","y"))', '=ISREF(CHOOSE(1,A1,B1))',
  '=ISREF(CHOOSE(1,1,2))', '=ISREF(A1:INDEX(A1:B2,2,1))',
  # ④誤りの マスを 指す
  '=ISREF(E5)', '=ISREF(INDEX(A1:B2,99,1))',
  # ⑤入れ子（他の 関数の 中）
  '=IF(ISREF(A1),"は","い")', '=SUM(ISREF(A1)*1,1)'
)

$結果 = New-Object System.Collections.ArrayList
$ws.Range('E5').Formula = '=1/0'
foreach ($式 in $式たち) {
  $答 = ''; $型 = ''
  try {
    $ws.Range('H1').Formula = $式
    $v = $ws.Range('H1').Value2
    if ($null -eq $v) { $答 = '(空)'; $型 = 'null' }
    else {
      $型 = $v.GetType().Name
      if ($v -is [double]) { $答 = $v.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture) }
      elseif ($v -is [int] -and $誤りの番号.ContainsKey([int]$v)) { $答 = $誤りの番号[[int]$v]; $型 = 'error値' }
      else { $答 = [string]$v }
    }
  } catch { $答 = '★Excel が 式として 受け取らない★'; $型 = 'error' }
  [void]$結果.Add(("ISREF`t{0}`t{1}`t{2}" -f $式, $答, $型))
}

$wb.Close($false); $xl.Quit()
foreach ($o in @($ws, $wb, $xl)) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null }

$頭 = @(
  "# ★実Excel に 打たせた ISREF の 答え★",
  "# 取った 日 … 2026-09-07",
  "# ★どの Excel で 打ったか★ … 版 $版 ／ build $ビルド",
  "# ★置いた 中身★ A1:E5 に 数（行*10+列）／E5 だけ =1/0（誤り）／名前 なまえ = Sheet1!`$A`$1:`$B`$2",
  "# 関数`t式`t実Excel の 答え`t型"
)
[System.IO.File]::WriteAllText($出, ((($頭 + $結果) -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding $false))
Write-Host "★書いた … $出（$($結果.Count) 本）★"
