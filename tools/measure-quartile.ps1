# 箱ひげ図の 四分位＝実Excel で 測る（読むだけ・保存しない）
$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add(); $ws=$wb.Worksheets.Item(1)
1..10 | ForEach-Object { $ws.Cells.Item($_,1).Value2 = $_ }
$ws.Range('C1').Formula='=QUARTILE.EXC(A1:A10,1)'
$ws.Range('C2').Formula='=QUARTILE.EXC(A1:A10,2)'
$ws.Range('C3').Formula='=QUARTILE.EXC(A1:A10,3)'
$ws.Range('D1').Formula='=QUARTILE.INC(A1:A10,1)'
$ws.Range('D2').Formula='=QUARTILE.INC(A1:A10,2)'
$ws.Range('D3').Formula='=QUARTILE.INC(A1:A10,3)'
"EXC 1,2,3 = {0} / {1} / {2}" -f $ws.Range('C1').Value2,$ws.Range('C2').Value2,$ws.Range('C3').Value2
"INC 1,2,3 = {0} / {1} / {2}" -f $ws.Range('D1').Value2,$ws.Range('D2').Value2,$ws.Range('D3').Value2
# 箱ひげ図の 既定の 四分位のやり方
$ws.Range('A1:A10').Select() | Out-Null
$co=$ws.Shapes.AddChart2(-1,121)
$sc=$co.Chart.SeriesCollection(1)
try { "箱ひげの既定 QuartileCalculation = {0}  (1=排他/2=包含 のどちらか)" -f $sc.QuartileCalculationType }
catch { "箱ひげの既定 QuartileCalculation … ★読めなかった★ {0}" -f $_.Exception.Message }
try { "  平均の線 ShowMeanLine={0} 内側の点 ShowInnerPoints={1}" -f $sc.ShowMeanLine,$sc.ShowInnerPoints } catch {}
$co.Delete()
$wb.Close($false); $xl.Quit()
