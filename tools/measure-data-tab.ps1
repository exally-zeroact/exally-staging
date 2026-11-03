# データタブの 真値を 実Excel で 測る（読むだけ・保存しない）
$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add(); $ws=$wb.Worksheets.Item(1)

"── フラッシュ フィル ──"
$ws.Range('A1').Value2='やまだ たろう'; $ws.Range('A2').Value2='すずき はなこ'; $ws.Range('A3').Value2='さとう じろう'
$ws.Range('B1').Value2='やまだ'
$ws.Range('B1').Select() | Out-Null
try { $ws.Range('B1:B3').Select() | Out-Null; $xl.Selection.FlashFill(); "  FlashFill を 呼べた" }
catch { "  ★FlashFill は 呼べない★ {0}" -f $_.Exception.Message }
"  結果 … B1='{0}' B2='{1}' B3='{2}'" -f $ws.Range('B1').Text,$ws.Range('B2').Text,$ws.Range('B3').Text

"── ゴールシーク（What-If）──"
$ws.Range('D1').Value2 = 10
$ws.Range('D2').Formula = '=D1*3+5'
"  はじめ … D1={0} D2={1}" -f $ws.Range('D1').Value2, $ws.Range('D2').Value2
$ok = $ws.Range('D2').GoalSeek(50, $ws.Range('D1'))
"  GoalSeek(50) = {0} → D1={1} D2={2}" -f $ok, [math]::Round($ws.Range('D1').Value2,10), $ws.Range('D2').Value2
"  ★どこまで 合わせるか（既定の 反復）★ 反復計算={0} 回数={1} 変化の下限={2}" -f `
  $wb.Application.Iteration, $wb.Application.MaxIterations, $wb.Application.MaxChange

"── 統合（Consolidate）──"
$ws.Range('F1').Value2='あ'; $ws.Range('G1').Value2=1
$ws.Range('F2').Value2='い'; $ws.Range('G2').Value2=2
$ws.Range('F3').Value2='あ'; $ws.Range('G3').Value2=3
"  関数の 既定（xlSum）= {0}" -f -4157

"── 再適用 ──"
$ws.Range('I1').Value2='名'; $ws.Range('I2').Value2='あ'; $ws.Range('I3').Value2='い'; $ws.Range('I4').Value2='あ'
$ws.Range('I1:I4').AutoFilter(1, 'あ') | Out-Null
"  絞った後 … 見えている行 = {0}" -f $ws.Range('I2:I4').SpecialCells(12).Count
$ws.Range('I3').Value2='あ'
"  値を 直した後（再適用の 前）= {0}" -f $ws.Range('I2:I4').SpecialCells(12).Count
$ws.AutoFilter.ApplyFilter()
"  ★ApplyFilter（再適用）の 後 = {0}★" -f $ws.Range('I2:I4').SpecialCells(12).Count
$wb.Close($false); $xl.Quit()
