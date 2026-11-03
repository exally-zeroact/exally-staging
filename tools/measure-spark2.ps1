$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add(); $ws=$wb.Worksheets.Item(1)
$ws.Range('A1').Value2=10; $ws.Range('B1').Value2=-5; $ws.Range('C1').Value2=30; $ws.Range('D1').Value2=20
$先=$ws.Range('F1'); $sg=$先.SparklineGroups.Add(1,'A1:D1')
"既定 … 縦軸最小の型={0} 空の扱い={1}" -f $sg.Axes.Vertical.MinScaleType, $sg.DisplayBlanksAs
foreach($v in 1,2,3){ try { $sg.Axes.Vertical.MinScaleType=$v; "  {0} を入れたら {1}" -f $v,$sg.Axes.Vertical.MinScaleType } catch { "  {0} は 入らない" -f $v } }
foreach($v in 1,2,3){ try { $sg.DisplayBlanksAs=$v; "  空={0} を入れたら {1}" -f $v,$sg.DisplayBlanksAs } catch { "  空={0} は 入らない" -f $v } }
$先.SparklineGroups.Clear() | Out-Null
$wb.Close($false); $xl.Quit()
