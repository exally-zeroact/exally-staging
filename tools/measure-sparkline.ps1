# スパークラインの 既定を 実Excel で 測る（読むだけ・保存しない）
$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add(); $ws=$wb.Worksheets.Item(1)
$ws.Range('A1').Value2=10; $ws.Range('B1').Value2=-5; $ws.Range('C1').Value2=30; $ws.Range('D1').Value2=20
$種類 = @{ '折れ線'=1; '縦棒'=2; '勝敗'=3 }
foreach($k in $種類.Keys){
  $t=$種類[$k]
  $先 = $ws.Range('F1')
  $sg = $先.SparklineGroups.Add($t, 'A1:D1')
  $行 = "{0} 番号={1}" -f $k,$t
  try { $行 += "  線の太さ={0}" -f $sg.LineWeight } catch {}
  try { $行 += "  高い点={0} 低い点={1} 最初={2} 最後={3} マイナス={4} 印={5}" -f `
     $sg.Points.Highpoint.Visible,$sg.Points.Lowpoint.Visible,$sg.Points.Firstpoint.Visible,`
     $sg.Points.Lastpoint.Visible,$sg.Points.Negative.Visible,$sg.Points.Markers.Visible } catch { $行 += "  点=(読めない)" }
  try { $行 += "  縦軸の最小={0} 最大={1}" -f $sg.Axes.Vertical.MinScaleType,$sg.Axes.Vertical.MaxScaleType } catch {}
  try { $行 += "  横軸を出す={0}" -f $sg.Axes.Horizontal.Axis.Visible } catch {}
  try { $行 += "  空の扱い={0}" -f $sg.DisplayBlanksAs } catch {}
  $行
  $先.SparklineGroups.Clear() | Out-Null
}
"（縦軸の型 0=個別 1=同じ 2=決め打ち ／ 空の扱い 0=空 1=零 2=つなぐ）"
$wb.Close($false); $xl.Quit()
