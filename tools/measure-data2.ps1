# スライサー／統合／アウトライン の 真値を 実Excel で 測る（読むだけ・保存しない）
$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add(); $ws=$wb.Worksheets.Item(1)
$ws.Range('A1').Value2='店'; $ws.Range('B1').Value2='金'
$ws.Range('A2').Value2='東'; $ws.Range('B2').Value2=100
$ws.Range('A3').Value2='西'; $ws.Range('B3').Value2=200
$ws.Range('A4').Value2='東'; $ws.Range('B4').Value2=300
$lo = $ws.ListObjects.Add(1, $ws.Range('A1:B4'), $null, 1)

"── スライサー ──"
try {
  $sc = $wb.SlicerCaches.Add2($lo, '店')
  $sl = $sc.Slicers.Add($ws, $null, '店', '店', 10, 300, 144, 200)
  "  作れた … 名前='{0}' 大きさ={1}x{2}" -f $sl.Name, $sl.Width, $sl.Height
  "  中の 見出し = '{0}' 出すか={1}" -f $sl.Caption, $sl.DisplayHeader
  "  列の 数 = {0}" -f $sl.NumberOfColumns
  "  中身 = {0}" -f (($sc.SlicerItems | ForEach-Object { $_.Name + '(' + $_.Selected + ')' }) -join ' / ')
  $sl.Delete()
} catch { "  ★スライサーは 作れない★ {0}" -f $_.Exception.Message }

"── アウトライン（グループ化）──"
$ws.Range('5:7').Select() | Out-Null
$ws.Range('A5').Value2='あ'; $ws.Range('A6').Value2='い'; $ws.Range('A7').Value2='う'
$ws.Rows('5:7').Group() | Out-Null
"  5行目の 段 = {0}" -f $ws.Rows(5).OutlineLevel
"  8行目の 段 = {0}" -f $ws.Rows(8).OutlineLevel
"  たたむ前 … 5行目は 隠れているか = {0}" -f $ws.Rows(5).Hidden
$ws.Outline.ShowLevels(1) | Out-Null
"  ★1段だけ 出したら 5行目は 隠れているか = {0}★" -f $ws.Rows(5).Hidden
$ws.Outline.ShowLevels(2) | Out-Null
"  2段 出したら 5行目は 隠れているか = {0}" -f $ws.Rows(5).Hidden
"  まとめの 行は 下か = {0}（SummaryRow）" -f $ws.Outline.SummaryRow
"  まとめの 列は 右か = {0}（SummaryColumn）" -f $ws.Outline.SummaryColumn

"── 統合（Consolidate）の 関数の 番号 ──"
"  合計=-4157 / 個数=-4112 / 平均=-4106 / 最大=-4136 / 最小=-4139"
$wb.Close($false); $xl.Quit()
