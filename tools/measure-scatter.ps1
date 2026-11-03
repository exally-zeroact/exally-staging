# 散布図が ★どの列を X に するか★ を 実Excel で 測る（読むだけ）
$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add(); $ws=$wb.Worksheets.Item(1)
function 測る($名, $範囲) {
  $ws.Range($範囲).Select() | Out-Null
  $co=$ws.Shapes.AddChart2(-1,-4169)
  $ch=$co.Chart
  $n=$ch.SeriesCollection().Count
  $行 = "{0} … 系列={1}" -f $名,$n
  for ($i=1; $i -le $n; $i++) {
    $s=$ch.SeriesCollection($i)
    $x = try { ($s.XValues -join ',') } catch { '(読めない)' }
    $y = try { ($s.Values  -join ',') } catch { '(読めない)' }
    $行 += "  [{0}] 名='{1}' X={2} Y={3}" -f $i,$s.Name,$x,$y
  }
  $co.Delete()
  $行
}
# ① 1列目が 字
$ws.Range('A1').Value2='あ'; $ws.Range('B1').Value2=1; $ws.Range('C1').Value2=10
$ws.Range('A2').Value2='い'; $ws.Range('B2').Value2=4; $ws.Range('C2').Value2=20
$ws.Range('A3').Value2='う'; $ws.Range('B3').Value2=9; $ws.Range('C3').Value2=30
測る '①1列目が字(A1:C3)' 'A1:C3'
# ② 数だけ 2列
$ws.Range('E1').Value2=1; $ws.Range('F1').Value2=10
$ws.Range('E2').Value2=4; $ws.Range('F2').Value2=20
$ws.Range('E3').Value2=9; $ws.Range('F3').Value2=30
測る '②数だけ2列(E1:F3)' 'E1:F3'
# ③ 数だけ 3列
$ws.Range('H1').Value2=1; $ws.Range('I1').Value2=10; $ws.Range('J1').Value2=100
$ws.Range('H2').Value2=4; $ws.Range('I2').Value2=20; $ws.Range('J2').Value2=200
$ws.Range('H3').Value2=9; $ws.Range('I3').Value2=30; $ws.Range('J3').Value2=300
測る '③数だけ3列(H1:J3)' 'H1:J3'
$wb.Close($false); $xl.Quit()
