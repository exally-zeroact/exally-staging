# 予測（FORECAST 系）と メモ の 真値を 実Excel で 測る（読むだけ・保存しない）
$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add(); $ws=$wb.Worksheets.Item(1)
# 1..6 の 月に 100,120,140,160,180,200（きれいな 直線）
for($i=1;$i -le 6;$i++){ $ws.Cells.Item($i,1).Value2 = $i; $ws.Cells.Item($i,2).Value2 = 80 + $i*20 }
"── 直線の 予測（FORECAST.LINEAR）──"
$ws.Range('D1').Formula='=FORECAST.LINEAR(7,B1:B6,A1:A6)'
$ws.Range('D2').Formula='=FORECAST(7,B1:B6,A1:A6)'
$ws.Range('D3').Formula='=TREND(B1:B6,A1:A6,7)'
$ws.Range('D4').Formula='=SLOPE(B1:B6,A1:A6)'
$ws.Range('D5').Formula='=INTERCEPT(B1:B6,A1:A6)'
"  FORECAST.LINEAR(7) = {0}" -f $ws.Range('D1').Value2
"  FORECAST(7)        = {0}" -f $ws.Range('D2').Value2
"  TREND(7)           = {0}" -f $ws.Range('D3').Value2
"  SLOPE              = {0}" -f $ws.Range('D4').Value2
"  INTERCEPT          = {0}" -f $ws.Range('D5').Value2
# でこぼこの 数でも 測る
$ws.Range('F1').Value2=100; $ws.Range('F2').Value2=90; $ws.Range('F3').Value2=130
$ws.Range('F4').Value2=120; $ws.Range('F5').Value2=160; $ws.Range('F6').Value2=150
$ws.Range('G1').Formula='=FORECAST.LINEAR(7,F1:F6,A1:A6)'
$ws.Range('G2').Formula='=SLOPE(F1:F6,A1:A6)'
$ws.Range('G3').Formula='=INTERCEPT(F1:F6,A1:A6)'
"  でこぼこ … 予測(7)={0} 傾き={1} 切片={2}" -f $ws.Range('G1').Value2, $ws.Range('G2').Value2, $ws.Range('G3').Value2
"── メモ（作者・大きさ）──"
$c = $ws.Range('A1').AddComment('めも')
"  作者は 空か = {0}（Excelの ユーザー名が 入る＝名前そのものは 書き残さない）" -f ([string]::IsNullOrEmpty($c.Author))
"  既定で 見えているか = {0}" -f $c.Visible
"  形 = {0} x {1}" -f [math]::Round($c.Shape.Width,1), [math]::Round($c.Shape.Height,1)
$ws.Range('A1').ClearComments()
$wb.Close($false); $xl.Quit()
