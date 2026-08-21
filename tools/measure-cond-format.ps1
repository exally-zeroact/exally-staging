# measure-cf2.ps1 — ★条件付き書式の「当たり」を Excel自身に描かせて測る★
#   ★1回目の測り方は嘘だった★＝ =IF(A6>25,...) は「式の答え」であって
#   「条件付き書式が そのセルに当たったか」ではない。
#   ここでは ★Range.DisplayFormat.Interior.Color★（実際に描かれる色）を読む。
$ErrorActionPreference = 'Stop'
$out = [ordered]@{}
$RED = 255          # BGR: 赤
$GRN = 65280        # BGR: 緑

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$out['_excel'] = 'Excel ' + $xl.Version + ' build ' + $xl.Build + ' UI=' + $xl.LanguageSettings.LanguageID(2)
$out['_測り方'] = '★Range.DisplayFormat.Interior.Color を読む＝実際に描かれる色（条件付き書式込み）★'

$bk = $xl.Workbooks.Add()
$sh = $bk.Worksheets.Item(1)
function Put($a, $v) { $sh.Range($a).Formula = [string]$v }
function 色($a) { return [double]$sh.Range($a).DisplayFormat.Interior.Color }
function 色を出す($a) { return [string](色 $a) }
# ★PowerShellの罠★ 色($a) -eq $c と書くと ★色 に3つの引数を渡す★（比べていない）。必ず (色 $a) と書く
function 当たったか($a, $c) { if ((色 $a) -eq $c) { return '当たる' } else { return '当たらない' } }
function 全消し() { $sh.Cells.FormatConditions.Delete() }

Put 'A1' '10'; Put 'A2' '20'; Put 'A3' '30'; Put 'A4' '40'; Put 'A5' '50'
# A6 は空 ／ A7 は文字 ／ A8 は 空文字を返す式 ／ A9 はエラー
Put 'A7' 'あいう'; Put 'A8' '=IF(1=1,"","x")'; Put 'A9' '=1/0'

# ---- ① 「より大きい」25 ------------------------------------------------
全消し
$fc = $sh.Range('A1:A9').FormatConditions.Add(1, 5, '=25')
$fc.Interior.Color = $RED
$o = [ordered]@{}
foreach ($a in 'A1','A2','A3','A4','A5','A6','A7','A8','A9') {
  $o[$a + ' (' + [string]$sh.Range($a).Text + ')'] = 当たったか $a $RED
}
$out['①25より大きい'] = $o

# ---- ② 「より小さい」25（★空セルは当たるのか★）------------------------
全消し
$fc = $sh.Range('A1:A9').FormatConditions.Add(1, 6, '=25')   # xlLess=6
$fc.Interior.Color = $RED
$o = [ordered]@{}
foreach ($a in 'A1','A2','A6','A7','A8','A9') {
  $o[$a + ' (' + [string]$sh.Range($a).Text + ')'] = 当たったか $a $RED
}
$out['②25より小さい'] = $o

# ---- ③ 「次の値に等しい」空文字 ----------------------------------------
全消し
$fc = $sh.Range('A1:A9').FormatConditions.Add(1, 3, '=""')   # xlEqual=3
$fc.Interior.Color = $RED
$o = [ordered]@{}
foreach ($a in 'A6','A7','A8') { $o[$a] = 当たったか $a $RED }
$out['③「""に等しい」'] = $o

# ---- ④ 2本が重なった時（StopIfTrue の効き）-----------------------------
全消し
$a = $sh.Range('A1:A5').FormatConditions.Add(1, 5, '=15')  ; $a.Interior.Color = $RED  # >15
$b = $sh.Range('A1:A5').FormatConditions.Add(1, 5, '=35')  ; $b.Interior.Color = $GRN  # >35
$out['④COMで足した時の既定'] = [ordered]@{
  '>15 の Priority' = $a.Priority; '>15 の StopIfTrue' = $a.StopIfTrue
  '>35 の Priority' = $b.Priority; '>35 の StopIfTrue' = $b.StopIfTrue
  '★注意★' = '★COMのAddは 後ろに足す＋StopIfTrue=True になる。画面(UI)で足すと 先頭に入り StopIfTrue=False★＝測り方で答えが変わる所'
}
$o = [ordered]@{}
foreach ($x in 'A2','A4','A5') {
  $c = 色 $x
  $o[$x + ' (' + [string]$sh.Range($x).Text + ')'] = if ($c -eq $RED) { '赤(>15が勝ち)' } elseif ($c -eq $GRN) { '緑(>35が勝ち)' } else { '色なし' }
}
$out['④どちらが勝つか（>15が Priority 1）'] = $o
$b.SetFirstPriority()
$o2 = [ordered]@{}
foreach ($x in 'A2','A4','A5') {
  $c = 色 $x
  $o2[$x] = if ($c -eq $RED) { '赤(>15が勝ち)' } elseif ($c -eq $GRN) { '緑(>35が勝ち)' } else { '色なし' }
}
$out['④>35 を先頭にした後'] = $o2
$out['④結論'] = '★番号の小さいルールが勝つ★（同じ「塗り」を両方が持っている時）'

# ---- ⑤ 式のルール =$A1>25 を A1:C3 に当てる（相対参照のずれ）------------
全消し
Put 'B1' '99'; Put 'B2' '1'; Put 'C1' '5'
$sh.Range('A1').Select() | Out-Null
$fx = $sh.Range('A1:C3').FormatConditions.Add(2, 0, '=$A1>25')
$fx.Interior.Color = $RED
$o = [ordered]@{}
foreach ($x in 'A1','B1','C1','A2','B2','C2','A3','B3','C3') { $o[$x] = 当たったか $x $RED }
$out['⑤式 =$A1>25 を A1:C3 に当てた'] = $o
$out['⑤結論'] = '★列は $ で止まり、行だけ ずれる＝A列の値で その行 全部が当たる★'

# ---- ⑥ 上位2項目 / 重複 / 文字を含む -----------------------------------
全消し
$t = $sh.Range('A1:A5').FormatConditions.AddTop10(); $t.TopBottom = 0; $t.Rank = 2; $t.Interior.Color = $RED
$o = [ordered]@{}
foreach ($x in 'A1','A2','A3','A4','A5') { $o[$x + ' (' + [string]$sh.Range($x).Text + ')'] = 当たったか $x $RED }
$out['⑥上位2項目'] = $o

全消し
Put 'D1' 'あ'; Put 'D2' 'い'; Put 'D3' 'あ'; Put 'D4' 'う'
$d = $sh.Range('D1:D4').FormatConditions.AddUniqueValues(); $d.DupeUnique = 1; $d.Interior.Color = $RED
$o = [ordered]@{}
foreach ($x in 'D1','D2','D3','D4') { $o[$x + ' (' + [string]$sh.Range($x).Text + ')'] = 当たったか $x $RED }
$out['⑥重複する値'] = $o

全消し
Put 'E1' 'ABC'; Put 'E2' 'abc'; Put 'E3' 'xyz'; Put 'E4' 'xABCx'
$sh.Range('E1').Select() | Out-Null
$tx = $sh.Range('E1:E4').FormatConditions.Add(2, 0, '=NOT(ISERROR(SEARCH("abc",E1)))')
$tx.Interior.Color = $RED
$o = [ordered]@{}
foreach ($x in 'E1','E2','E3','E4') { $o[$x + ' (' + [string]$sh.Range($x).Text + ')'] = 当たったか $x $RED }
$out['⑥「abc を含む」（大文字小文字）'] = $o

# ---- ⑦ 手で塗った色より 条件付き書式が勝つか ---------------------------
全消し
$sh.Range('F1').Formula = '30'
$sh.Range('F1').Interior.Color = 65535       # 手で黄色
$f = $sh.Range('F1').FormatConditions.Add(1, 5, '=25'); $f.Interior.Color = $RED
$out['⑦手の塗り(黄) vs 条件付き書式(赤)'] = [ordered]@{
  'Interior.Color（そのセルが持つ色）'   = [double]$sh.Range('F1').Interior.Color
  'DisplayFormat（実際に描かれる色）'    = 色 'F1'
  '結論' = '★条件付き書式が勝つ。手の塗りは 消えずに 下に残る（条件を外すと戻る）★'
}

# ---- ⑧ 計算には影響しない ----------------------------------------------
$sh.Range('F3').Formula = '=SUM(A1:A5)'
$out['⑧合計'] = [ordered]@{ 'SUM(A1:A5)' = [string]$sh.Range('F3').Text; '結論' = '★見た目だけ。計算は1ミリも変わらない★' }

# ---- ⑨ ファイルにどう書かれるか ----------------------------------------
$tmp = Join-Path $env:TEMP 'cf-probe2.xlsx'
if (Test-Path $tmp) { Remove-Item $tmp -Force }
$bk.SaveAs($tmp, 51)
$out['⑨保存した見本'] = $tmp

$bk.Close($false); $xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

$json = $out | ConvertTo-Json -Depth 6
[System.IO.File]::WriteAllText((Join-Path $PSScriptRoot 'cond-format-measured2.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host $json
