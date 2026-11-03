# 箱ひげ図の 既定を ★保存した中身★から 読む（COMの性質が読めなかったため）
$ErrorActionPreference='Stop'
$out = $args[0]
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add(); $ws=$wb.Worksheets.Item(1)
1..10 | ForEach-Object { $ws.Cells.Item($_,1).Value2 = $_ }
$ws.Range('A1:A10').Select() | Out-Null
$co=$ws.Shapes.AddChart2(-1,121) | Out-Null
$wb.SaveAs($out, 51)
$wb.Close($false); $xl.Quit()
"保存した: $out"
