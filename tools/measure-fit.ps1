# 拡大縮小印刷／配置（整列・グループ化・回転）の 真値を 実Excel で 測る（読むだけ）
$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add(); $ws=$wb.Worksheets.Item(1); $ps=$ws.PageSetup
"── 拡大縮小印刷 ──"
"既定 … Zoom={0} 横={1} 縦={2}" -f $ps.Zoom, $ps.FitToPagesWide, $ps.FitToPagesTall
$ps.Zoom = $false
"Zoom を False に すると … Zoom={0} 横={1} 縦={2}（ページに 合わせる 形）" -f $ps.Zoom, $ps.FitToPagesWide, $ps.FitToPagesTall
$ps.FitToPagesWide = 1; $ps.FitToPagesTall = 2
"横1・縦2 に すると … 横={0} 縦={1}" -f $ps.FitToPagesWide, $ps.FitToPagesTall
$ps.Zoom = 75
"Zoom=75 に 戻すと … Zoom={0} 横={1} 縦={2}" -f $ps.Zoom, $ps.FitToPagesWide, $ps.FitToPagesTall
foreach($z in 10,25,100,400,401,9){
  try { $ps.Zoom = $z; "  Zoom {0} → {1}" -f $z, $ps.Zoom } catch { "  Zoom {0} は ★入らない★" -f $z }
}
$ps.Zoom = 100
"── 配置（整列・グループ化・回転）──"
$a = $ws.Shapes.AddShape(1, 10, 10, 60, 40)
$b = $ws.Shapes.AddShape(1, 100, 80, 60, 40)
"作った … A(左={0} 上={1}) B(左={2} 上={3})" -f $a.Left,$a.Top,$b.Left,$b.Top
$ws.Shapes.Range(@($a.Name,$b.Name)).Align(1, $false)   # 1 = msoAlignLefts
"左そろえ の後 … A左={0} B左={1}" -f $a.Left, $b.Left
$ws.Shapes.Range(@($a.Name,$b.Name)).Align(5, $false)   # 5 = msoAlignTops
"上そろえ の後 … A上={0} B上={1}" -f $a.Top, $b.Top
$g = $ws.Shapes.Range(@($a.Name,$b.Name)).Group()
"グループ化 … 名前='{0}' 中の数={1}" -f $g.Name, $g.GroupItems.Count
$g.Ungroup() | Out-Null
"解除の後 … Shapes.Count={0}" -f $ws.Shapes.Count
$ws.Shapes.Item(1).Rotation = 90
"回転 … {0}度" -f $ws.Shapes.Item(1).Rotation
$ws.Shapes.Item(1).IncrementRotation(90)
"さらに 90度 … {0}度" -f $ws.Shapes.Item(1).Rotation
$wb.Close($false); $xl.Quit()
