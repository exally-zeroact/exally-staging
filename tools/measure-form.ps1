# 「フォーム」（データ入力フォーム）の 中身を 測る（読むだけ）
# ★COM から 窓の 中は 読めない★ので、ShowDataForm が 何を 相手に するかだけ 確かめる。
$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add(); $ws=$wb.Worksheets.Item(1)
$ws.Range('A1').Value2='月'; $ws.Range('B1').Value2='売上'
$ws.Range('A2').Value2='1月'; $ws.Range('B2').Value2=100
$ws.Range('A3').Value2='2月'; $ws.Range('B3').Value2=150
"ShowDataForm が 在るか = {0}" -f ($null -ne ($ws | Get-Member -Name ShowDataForm))
# 相手に する 範囲（CurrentRegion）
$ws.Range('A2').Select() | Out-Null
"選んだ所の まわり（CurrentRegion）= {0}" -f $ws.Range('A2').CurrentRegion.Address(0,0)
"  行数={0} 列数={1}" -f $ws.Range('A2').CurrentRegion.Rows.Count, $ws.Range('A2').CurrentRegion.Columns.Count
# 列が 33 を 超えると フォームは 出せない（Excelの 決め）＝実際に 試す
for($c=1;$c -le 34;$c++){ $ws.Cells.Item(10,$c).Value2 = "列$c" }
$ws.Cells.Item(11,1).Value2 = 'あ'
$ws.Range('A10').Select() | Out-Null
"34列の まわり = {0}（列数={1}）" -f $ws.Range('A10').CurrentRegion.Address(0,0), $ws.Range('A10').CurrentRegion.Columns.Count
$wb.Close($false); $xl.Quit()
