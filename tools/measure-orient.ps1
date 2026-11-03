# ★系列を 縦に取るか 横に取るか★ を 実Excel で 測る（読むだけ）
$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add(); $ws=$wb.Worksheets.Item(1)
function 置く($左上, $行数, $列数) {
  $r0=$ws.Range($左上).Row; $c0=$ws.Range($左上).Column
  for ($r=0;$r -lt $行数;$r++){ for($c=0;$c -lt $列数;$c++){
    $ws.Cells.Item($r0+$r, $c0+$c).Value2 = ($r+1)*10 + ($c+1)
  }}
}
function 測る($名, $範囲, $種類) {
  $ws.Range($範囲).Select() | Out-Null
  $co=$ws.Shapes.AddChart2(-1,$種類)
  $ch=$co.Chart
  $n=$ch.SeriesCollection().Count
  $行 = "{0} 種類={1} … 系列={2}" -f $名,$種類,$n
  for ($i=1; $i -le [math]::Min($n,3); $i++) {
    $s=$ch.SeriesCollection($i)
    $y = try { ($s.Values -join ',') } catch { '?' }
    $行 += "  [{0}] Y={1}" -f $i,$y
  }
  $co.Delete()
  $行
}
置く 'A1' 3 3;  測る '3行3列' 'A1:C3' 51
置く 'E1' 4 2;  測る '4行2列' 'E1:F4' 51
置く 'H1' 2 4;  測る '2行4列' 'H1:K2' 51
置く 'A10' 5 5; 測る '5行5列' 'A10:E14' 51
$wb.Close($false); $xl.Quit()
