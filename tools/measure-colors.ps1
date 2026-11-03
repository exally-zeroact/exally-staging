# ★色を 16進で 測り直す★（10進の 数字を そのまま 16進と 読み違えていないかの 確かめ）
$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add(); $ws=$wb.Worksheets.Item(1)
function 色を出す($名, $rgb){
  $n = [int]$rgb
  $B = ($n -shr 16) -band 0xFF
  $G = ($n -shr 8)  -band 0xFF
  $R = $n -band 0xFF
  "  {0,-16} 10進={1,-10} 16進=0x{2:X6}  → R={3} G={4} B={5}  ＝ #{3:X2}{4:X2}{5:X2}" -f $名,$n,$n,$R,$G,$B
}
"── ハイパーリンクの 字の色 ──"
# ★$null は 渡せない（実測）★＝要る物だけ 渡す
$ws.Hyperlinks.Add($ws.Range('A1'), 'https://example.com/') | Out-Null
$ws.Range('A1').Value2 = 'リンクの字'
色を出す 'リンクの字' $ws.Range('A1').Font.Color
"  下線 = {0}" -f $ws.Range('A1').Font.Underline
$ws.Range('A1').Clear() | Out-Null
"── 図形の 既定 ──"
$sh = $ws.Shapes.AddShape(1, 10, 10, 100, 50)
色を出す '図形の 塗り' $sh.Fill.ForeColor.RGB
色を出す '図形の 線'   $sh.Line.ForeColor.RGB
"  線の 太さ = {0}" -f $sh.Line.Weight
$tb = $ws.Shapes.AddTextbox(1, 10, 80, 120, 30)
色を出す 'テキスト箱の塗り' $tb.Fill.ForeColor.RGB
"── ふつうの セル ──"
色を出す '字の色' $ws.Range('B2').Font.Color
$wb.Close($false); $xl.Quit()
