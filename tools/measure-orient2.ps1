$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add(); $ws=$wb.Worksheets.Item(1)
function 測る($名,$範囲){
  $ws.Range($範囲).Select() | Out-Null
  $co=$ws.Shapes.AddChart2(-1,51); $ch=$co.Chart
  $n=$ch.SeriesCollection().Count
  $行="{0} … 系列={1}" -f $名,$n
  for($i=1;$i -le $n;$i++){ $s=$ch.SeriesCollection($i)
    $行 += "  [{0}] 名='{1}' Y={2}" -f $i,$s.Name,($s.Values -join ',') }
  $x = try { ($ch.SeriesCollection(1).XValues -join ',') } catch { '?' }
  $行 += "  横軸={0}" -f $x
  $co.Delete(); $行
}
# 見出し行＋見出し列＋数2x2
$ws.Range('A1').Value2='月'; $ws.Range('B1').Value2='売上'; $ws.Range('C1').Value2='原価'
$ws.Range('A2').Value2='1月'; $ws.Range('B2').Value2=100; $ws.Range('C2').Value2=60
$ws.Range('A3').Value2='2月'; $ws.Range('B3').Value2=150; $ws.Range('C3').Value2=80
測る '見出し付き 数2行2列(A1:C3)' 'A1:C3'
# 見出し行＋見出し列＋数3x2
$ws.Range('E1').Value2='月'; $ws.Range('F1').Value2='売上'; $ws.Range('G1').Value2='原価'
$ws.Range('E2').Value2='1月'; $ws.Range('F2').Value2=100; $ws.Range('G2').Value2=60
$ws.Range('E3').Value2='2月'; $ws.Range('F3').Value2=150; $ws.Range('G3').Value2=80
$ws.Range('E4').Value2='3月'; $ws.Range('F4').Value2=120; $ws.Range('G4').Value2=70
測る '見出し付き 数3行2列(E1:G4)' 'E1:G4'
# 見出し行＋見出し列＋数2x3
$ws.Range('A10').Value2='月'; $ws.Range('B10').Value2='売上'; $ws.Range('C10').Value2='原価'; $ws.Range('D10').Value2='粗利'
$ws.Range('A11').Value2='1月'; $ws.Range('B11').Value2=100; $ws.Range('C11').Value2=60; $ws.Range('D11').Value2=40
$ws.Range('A12').Value2='2月'; $ws.Range('B12').Value2=150; $ws.Range('C12').Value2=80; $ws.Range('D12').Value2=70
測る '見出し付き 数2行3列(A10:D12)' 'A10:D12'
$wb.Close($false); $xl.Quit()
