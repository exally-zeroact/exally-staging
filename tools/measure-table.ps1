# テーブル（Ctrl+T）の 既定を 実Excel で 測る（読むだけ・保存しない）
$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add(); $ws=$wb.Worksheets.Item(1)
$ws.Range('A1').Value2='月'; $ws.Range('B1').Value2='売上'; $ws.Range('C1').Value2='原価'
$ws.Range('A2').Value2='1月'; $ws.Range('B2').Value2=100; $ws.Range('C2').Value2=60
$ws.Range('A3').Value2='2月'; $ws.Range('B3').Value2=150; $ws.Range('C3').Value2=80
$ws.Range('A4').Value2='3月'; $ws.Range('B4').Value2=120; $ws.Range('C4').Value2=70
$lo = $ws.ListObjects.Add(1, $ws.Range('A1:C4'), $null, 1)   # 1=xlSrcRange, 1=xlYes(見出しあり)
"名前          = {0}" -f $lo.Name
"範囲          = {0}" -f $lo.Range.Address(0,0)
"見出し行      = {0}" -f $lo.ShowHeaders
"集計行        = {0}" -f $lo.ShowTotals
"自動フィルタ  = {0}" -f $lo.ShowAutoFilter
"しま（行）    = {0}" -f $lo.ShowTableStyleRowStripes
"しま（列）    = {0}" -f $lo.ShowTableStyleColumnStripes
"最初の列      = {0}" -f $lo.ShowTableStyleFirstColumn
"最後の列      = {0}" -f $lo.ShowTableStyleLastColumn
"見た目の名    = {0}" -f $lo.TableStyle.Name
"列の名        = {0}" -f (($lo.ListColumns | ForEach-Object { $_.Name }) -join ' / ')
"行の数        = {0}" -f $lo.ListRows.Count
# 2つ目を 作ると 名前は どう なるか
$ws.Range('E1').Value2='あ'; $ws.Range('E2').Value2=1
$lo2 = $ws.ListObjects.Add(1, $ws.Range('E1:E2'), $null, 1)
"2つ目の名前  = {0}" -f $lo2.Name
# 集計行を 出すと 何が 入るか
$lo.ShowTotals = $true
"集計行を出すと 範囲 = {0}" -f $lo.Range.Address(0,0)
$r = $lo.TotalsRowRange
"  集計行の中身 = {0}" -f (($r.Cells | ForEach-Object { "'" + $_.Text + "'/" + $_.Formula }) -join ' | ')
$wb.Close($false); $xl.Quit()
