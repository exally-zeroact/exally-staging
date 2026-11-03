# 表示タブの 真値を 実Excel で 測る（読むだけ・保存しない）
$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add(); $ws=$wb.Worksheets.Item(1)
$w=$xl.ActiveWindow
"── ブックの表示 ──"
"既定の View = {0}  （1=標準 / 2=改ページプレビュー / 3=ページレイアウト）" -f $w.View
foreach($v in 1,2,3){ $w.View=$v; "  {0} を入れたら {1}（ズーム={2}）" -f $v,$w.View,$w.Zoom }
$w.View=1
"── ズーム ──"
"既定の Zoom = {0}" -f $w.Zoom
foreach($z in 25,50,75,100,200,400,10,5,401){
  try { $w.Zoom=$z; "  {0} を入れたら {1}" -f $z,$w.Zoom } catch { "  {0} は ★入らない★" -f $z }
}
$w.Zoom=100
"── 選択範囲に合わせる ──"
$ws.Range('A1').Value2=1; $ws.Range('E10').Value2=2
$ws.Range('A1:E10').Select() | Out-Null
$w.Zoom = $true      # True = 選択範囲に合わせる
"  A1:E10 に 合わせたら Zoom={0}" -f $w.Zoom
$w.Zoom=100
"── 分割 ──"
$ws.Range('C5').Select() | Out-Null
$w.Split = $true
"  Split=True … SplitRow={0} SplitColumn={1}" -f $w.SplitRow,$w.SplitColumn
$w.Split = $false
"  Split=False … SplitRow={0} SplitColumn={1}" -f $w.SplitRow,$w.SplitColumn
"── 改ページ プレビューの 既定 ──"
$w.View=2
"  改ページ中の Zoom = {0}" -f $w.Zoom
$w.View=1
$wb.Close($false); $xl.Quit()
