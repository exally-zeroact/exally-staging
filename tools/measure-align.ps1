# ★そろえ方の 番号★を 1つずつ 実Excel で 測る（読むだけ）
$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add(); $ws=$wb.Worksheets.Item(1)
foreach($n in 0,1,2,3,4,5,6){
  while($ws.Shapes.Count -gt 0){ $ws.Shapes.Item(1).Delete() }
  $a = $ws.Shapes.AddShape(1, 10, 10, 60, 40)
  $b = $ws.Shapes.AddShape(1, 100, 80, 80, 20)
  try {
    $ws.Shapes.Range(@($a.Name,$b.Name)).Align($n, $false)
    "{0} … A(左={1} 上={2} 幅={3} 高={4}) B(左={5} 上={6} 幅={7} 高={8})" -f `
      $n,$a.Left,$a.Top,$a.Width,$a.Height,$b.Left,$b.Top,$b.Width,$b.Height
  } catch { "{0} … ★出来ない★" -f $n }
}
$wb.Close($false); $xl.Quit()
