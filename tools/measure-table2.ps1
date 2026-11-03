$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add(); $ws=$wb.Worksheets.Item(1)
# ① 1行目が 字（見出しらしい）
$ws.Range('A1').Value2='月'; $ws.Range('B1').Value2='売上'
$ws.Range('A2').Value2='1月'; $ws.Range('B2').Value2=100
$lo=$ws.ListObjects.Add(1,$ws.Range('A1:B2'),$null,0)   # 0 = xlGuess
"①1行目が字 … 見出し行={0} 範囲={1} 列の名={2}" -f $lo.ShowHeaders,$lo.Range.Address(0,0),(($lo.ListColumns|%{$_.Name}) -join ',')
# ② 1行目も 数
$ws.Range('D1').Value2=1; $ws.Range('E1').Value2=2
$ws.Range('D2').Value2=3; $ws.Range('E2').Value2=4
$lo2=$ws.ListObjects.Add(1,$ws.Range('D1:E2'),$null,0)
"②1行目も数 … 見出し行={0} 範囲={1} 列の名={2}" -f $lo2.ShowHeaders,$lo2.Range.Address(0,0),(($lo2.ListColumns|%{$_.Name}) -join ',')
# ③ SUBTOTAL 109 は 何か
$ws.Range('H1').Value2=1;$ws.Range('H2').Value2=2;$ws.Range('H3').Value2=3
$ws.Range('I1').Formula='=SUBTOTAL(109,H1:H3)'
$ws.Range('I2').Formula='=SUBTOTAL(9,H1:H3)'
$ws.Rows.Item(2).Hidden=$true
"③SUBTOTAL 109={0} / 9={1}（2行目を 隠した時）" -f $ws.Range('I1').Value2,$ws.Range('I2').Value2
$ws.Rows.Item(2).Hidden=$false
# ④ 集計行の 既定は どの列に 入るか（2列 と 3列）
"④2列テーブルの 集計行:"
$lo.ShowTotals=$true
"   {0}" -f (($lo.TotalsRowRange.Cells | %{ "'"+$_.Text+"'/"+$_.Formula }) -join ' | ')
$wb.Close($false); $xl.Quit()
