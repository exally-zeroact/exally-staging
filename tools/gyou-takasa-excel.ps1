param([string]$Out)
$ErrorActionPreference = 'Stop'
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  for ($r = 1; $r -le 12; $r++) { for ($c = 1; $c -le 6; $c++) { $ws.Cells.Item($r, $c).Value2 = "$r-$c" } }
  $ws.Rows.Item(2).RowHeight = 6
  $ws.Rows.Item(3).Hidden = $true
  $ws.Rows.Item(5).RowHeight = 30
  $ws.Rows.Item(7).RowHeight = 0.75
  $ws.Columns.Item(2).Hidden = $true
  $ws.Columns.Item(4).ColumnWidth = 2
  $ws.Rows.Item(9).RowHeight = 30
  $ws.Rows.Item(9).Hidden = $true
  $ws.Columns.Item(6).ColumnWidth = 5
  $ws.Columns.Item(6).Hidden = $true
  $res = @()
  foreach ($r in 1..10) { $res += "row $r h=$($ws.Rows.Item($r).RowHeight) hid=$($ws.Rows.Item($r).Hidden) top=$($ws.Rows.Item($r).Top)" }
  foreach ($c in 1..7) { $res += "col $c w=$($ws.Columns.Item($c).Width) hid=$($ws.Columns.Item($c).Hidden)" }
  $res += "std h=$($ws.StandardHeight) std w=$($ws.StandardWidth)"
  $res | Out-File -Encoding utf8 ($Out + '.txt')
  $wb.SaveAs($Out + '.xlsx', 51)
  $wb.SaveAs($Out + '.xlsb', 50)
  $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
# Excel は サインインの 名前と 保存した 道を 入れるので、閉じた 後に repo の 道具 scripts/sakusha-mihon.mjs で 替える（判じは 1か所・11-07 押す前の門が止めた）
& node (Join-Path $PSScriptRoot '../scripts/sakusha-mihon.mjs') ($Out + '.xlsx') ($Out + '.xlsb')
if ($LASTEXITCODE -ne 0) { throw 'sakusha-mihon が 赤' }
