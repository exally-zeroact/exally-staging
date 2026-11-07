param([string]$Out)
$ErrorActionPreference = 'Stop'
# ★xlsb で Excel が 自動で 決めた 行の 高さ★（2026-11-07）＝字の 大きさ・折り返しで Excel が 決めた 高さ（customHeight 無し）
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  for ($r = 1; $r -le 10; $r++) { $ws.Cells.Item($r, 1).Value2 = "$r" }
  $ws.Range('A2').Font.Size = 24
  $ws.Range('A4').WrapText = $true; $ws.Range('A4').Value2 = "a`nb`nc"
  $ws.Rows.Item(6).RowHeight = 30
  $ws.Range('A8').Font.Size = 20; $ws.Rows.Item(8).Hidden = $true
  $ws.Range('A9').Font.Size = 6
  $res = @()
  foreach ($r in 1..10) { $res += "row $r h=$($ws.Rows.Item($r).RowHeight) hid=$($ws.Rows.Item($r).Hidden) top=$($ws.Rows.Item($r).Top)" }
  $res += "std h=$($ws.StandardHeight)"
  $res | Out-File -Encoding utf8 ($Out + '.txt')
  $wb.SaveAs($Out + '.xlsx', 51)
  $wb.SaveAs($Out + '.xlsb', 50)
  $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
# Excel は サインインの 名前と 保存した 道を 入れるので、閉じた 後に scripts/sakusha-mihon.mjs で 替える
& node (Join-Path $PSScriptRoot '../scripts/sakusha-mihon.mjs') ($Out + '.xlsx') ($Out + '.xlsb')
if ($LASTEXITCODE -ne 0) { throw 'sakusha-mihon が 赤' }
