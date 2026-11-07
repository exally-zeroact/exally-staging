# 隠した 行・列の 後ろに 判子（四角）を 置いた 材料を 実Excel（COM）で 作る（2026-11-07）
#   使い方: powershell -File tools/hanko-kakushi-excel.ps1 -Out <道\名前（拡張子なし）>
#   出る物: <名前>.xlsx ／ <名前>.xlsb ／ <名前>.txt（実Excel が 言う 判子の Left・Top＝答え）
param([string]$Out)
$ErrorActionPreference = 'Stop'
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  for ($r = 1; $r -le 12; $r++) { for ($c = 1; $c -le 6; $c++) { $ws.Cells.Item($r, $c).Value2 = "$r-$c" } }
  $ws.Rows.Item(3).Hidden = $true
  $ws.Rows.Item(5).RowHeight = 30
  $ws.Rows.Item(7).RowHeight = 24
  $ws.Rows.Item(7).Hidden = $true
  $ws.Columns.Item(2).Hidden = $true
  $ws.Columns.Item(3).ColumnWidth = 20
  $ws.Columns.Item(3).Hidden = $true
  $c = $ws.Cells.Item(10, 5)
  $s = $ws.Shapes.AddShape(1, $c.Left + 6, $c.Top + 3, 45, 45)
  $s.Name = 'hanko'
  $res = @()
  $res += "shape left=$($s.Left) top=$($s.Top) tl=$($s.TopLeftCell.Address(0,0))"
  foreach ($r in 1..10) { $res += "row $r h=$($ws.Rows.Item($r).RowHeight) hid=$($ws.Rows.Item($r).Hidden) top=$($ws.Rows.Item($r).Top)" }
  foreach ($k in 1..6) { $res += "col $k w=$($ws.Columns.Item($k).Width) hid=$($ws.Columns.Item($k).Hidden) left=$($ws.Columns.Item($k).Left)" }
  $res | Out-File -Encoding utf8 ($Out + '.txt')
  $wb.SaveAs($Out + '.xlsx', 51)
  $wb.SaveAs($Out + '.xlsb', 50)
  $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
# Excel は サインインの 名前と 保存した 道を 入れるので、閉じた 後に repo の 道具 scripts/sakusha-mihon.mjs で 替える（判じは 1か所・11-07 押す前の門が止めた）
& node (Join-Path $PSScriptRoot '../scripts/sakusha-mihon.mjs') ($Out + '.xlsx') ($Out + '.xlsb')
if ($LASTEXITCODE -ne 0) { throw 'sakusha-mihon が 赤' }
