# CSV の 読み方を 実Excel で 測る（読むだけ・保存しない）
$ErrorActionPreference='Stop'
$dir = Join-Path $env:TEMP 'exally-csv-measure'
New-Item -ItemType Directory -Force -Path $dir | Out-Null
$f = Join-Path $dir 'a.csv'
# ★"" と 中の カンマ・中の 改行★
$中身 = "名,金`r`n" + '"山田, 太郎",100' + "`r`n" + '"あ""い",200' + "`r`n" + '"1行目' + "`r`n" + '2行目",300' + "`r`n"
[System.IO.File]::WriteAllText($f, $中身, [Text.Encoding]::UTF8)
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Open($f)
$ws=$wb.Worksheets.Item(1)
"── UTF-8 の CSV を 開いた ──"
for($r=1;$r -le 5;$r++){
  $行=@(); for($c=1;$c -le 3;$c++){ $行 += "'" + $ws.Cells.Item($r,$c).Text + "'" }
  "  {0}行目: {1}" -f $r, ($行 -join ' ')
}
"  使った範囲 = {0}" -f $ws.UsedRange.Address(0,0)
$wb.Close($false)
# Shift_JIS でも 試す
$f2 = Join-Path $dir 'b.csv'
[System.IO.File]::WriteAllText($f2, "名,金`r`nあいう,10`r`n", [Text.Encoding]::GetEncoding(932))
$wb2=$xl.Workbooks.Open($f2)
"── Shift_JIS の CSV ── A2='{0}'" -f $wb2.Worksheets.Item(1).Range('A2').Text
$wb2.Close($false)
$xl.Quit()
Remove-Item -Recurse -Force $dir
