# ピボットテーブルの 既定を 実Excel で 測る（読むだけ・保存しない）
$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add(); $ws=$wb.Worksheets.Item(1)
# ★1行ずつ 素直に 置く（配列の 中の 数と 字が 混ざると PowerShell が 転ぶ）★
$ws.Range('A1').Value2='店'; $ws.Range('B1').Value2='品'; $ws.Range('C1').Value2='数'; $ws.Range('D1').Value2='金'
$ws.Range('A2').Value2='東'; $ws.Range('B2').Value2='A'; $ws.Range('C2').Value2=1; $ws.Range('D2').Value2=100
$ws.Range('A3').Value2='東'; $ws.Range('B3').Value2='B'; $ws.Range('C3').Value2=2; $ws.Range('D3').Value2=200
$ws.Range('A4').Value2='西'; $ws.Range('B4').Value2='A'; $ws.Range('C4').Value2=3; $ws.Range('D4').Value2=300
$ws.Range('A5').Value2='西'; $ws.Range('B5').Value2='B'; $ws.Range('C5').Value2=4; $ws.Range('D5').Value2=400
$ws.Range('A6').Value2='東'; $ws.Range('B6').Value2='A'; $ws.Range('C6').Value2=5; $ws.Range('D6').Value2=500
$ws2 = $wb.Worksheets.Add()
$cache = $wb.PivotCaches().Create(1, $ws.Range('A1:D6'))   # 1 = xlDatabase
$pt = $cache.CreatePivotTable($ws2.Range('A3'), 'ピボット1')
"名前 = {0}" -f $pt.Name
"　置いた所 = {0}" -f $pt.TableRange2.Address(0,0)
$pt.PivotFields('店').Orientation = 1     # xlRowField
$pt.PivotFields('品').Orientation = 2     # xlColumnField
# ★AddDataField に $null は 渡せない（実測）★＝名前を 決めて 渡す
$pt.AddDataField($pt.PivotFields('金'), '合計 / 金', -4157) | Out-Null   # xlSum
"── 出来た 形 ──"
"　範囲 = {0}" -f $pt.TableRange2.Address(0,0)
for($r=1;$r -le 8;$r++){
  $行 = @()
  for($c=1;$c -le 5;$c++){ $行 += "'" + $ws2.Cells.Item($r,$c).Text + "'" }
  "  {0}行目: {1}" -f $r, ($行 -join ' ')
}
"── 既定 ──"
"　合計の 名前 = '{0}'" -f $pt.DataFields(1).Name
"　総計（行）= {0} ／ 総計（列）= {1}" -f $pt.RowGrand, $pt.ColumnGrand
"　空セルの 見せ方 = '{0}'（NullString）／空を 出すか = {1}" -f $pt.NullString, $pt.DisplayNullString
"　小計 = {0}" -f $pt.PivotFields('店').Subtotals(1)

"── 列を 使わない 時（行と 値だけ）──"
$ws3 = $wb.Worksheets.Add()
$cache2 = $wb.PivotCaches().Create(1, $ws.Range('A1:D6'))
$pt2 = $cache2.CreatePivotTable($ws3.Range('A3'), 'ピボット2')
$pt2.PivotFields('店').Orientation = 1
$pt2.AddDataField($pt2.PivotFields('金'), '合計 / 金', -4157) | Out-Null
"　範囲 = {0}" -f $pt2.TableRange2.Address(0,0)
for($r=1;$r -le 8;$r++){
  $行 = @()
  for($c=1;$c -le 4;$c++){ $行 += "'" + $ws3.Cells.Item($r,$c).Text + "'" }
  "  {0}行目: {1}" -f $r, ($行 -join ' ')
}
$wb.Close($false); $xl.Quit()
