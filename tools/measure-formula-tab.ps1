# 数式タブの 真値を 実Excel で 測る（読むだけ・保存しない）
$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add(); $ws=$wb.Worksheets.Item(1)

"── 計算方法 ──"
"既定 = {0}  （-4105=自動 / -4135=手動 / 2=データテーブル以外自動）" -f $xl.Calculation
foreach($v in -4105,-4135,2){ $xl.Calculation=$v; "  {0} を入れたら {1}" -f $v,$xl.Calculation }
$xl.Calculation=-4105

"── 選択範囲から名前を作る ──"
$ws.Range('A1').Value2='月'; $ws.Range('B1').Value2='売上'
$ws.Range('A2').Value2='1月'; $ws.Range('B2').Value2=100
$ws.Range('A3').Value2='2月'; $ws.Range('B3').Value2=150
$ws.Range('A1:B3').CreateNames($true,$false,$false,$false)   # 上端行から
"作られた名前の数 = {0}" -f $wb.Names.Count
foreach($n in $wb.Names){ "  {0} → {1}" -f $n.Name, $n.RefersTo }

"── 名前に 使えない字 ──"
foreach($nm in '売上','売 上','1月','A1','あ_い','あ.い'){
  try { $wb.Names.Add($nm, '=Sheet1!$B$2') | Out-Null; "  '{0}' → 入った" -f $nm }
  catch { "  '{0}' → ★入らない★" -f $nm }
}

"── トレース（参照元・参照先）──"
$ws.Range('D1').Formula='=B2+B3'
$ws.Range('D2').Formula='=D1*2'
$ws.Range('D1').ShowPrecedents() | Out-Null
"  矢印の数（参照元を出した後）= {0}" -f $ws.Range('D1').Worksheet.Parent.Application.ActiveSheet.Shapes.Count
$ws.ClearArrows()
"  ClearArrows のあと = {0}" -f $ws.Shapes.Count

"── 数式の 検証（1手ずつ）──"
"  Evaluate('=B2+B3') = {0}" -f $ws.Evaluate('=B2+B3')
"  Evaluate('B2+B3')  = {0}" -f $ws.Evaluate('B2+B3')

$wb.Close($false); $xl.Quit()
