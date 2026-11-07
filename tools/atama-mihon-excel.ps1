param([string]$Mode, [string]$Path, [string]$Rows = '1,3,5,7', [switch]$Hiroi)
$ErrorActionPreference = 'Stop'
# ★xlsb の 新しい 行の 頭★（2026-11-07）＝最初の 行が 手で 決めた 高さ・隠し・段の 組・行の 書式 を 持つ 本を 作る／直した 本を 開いて 測る
#   make  … $Path.xlsb を 作る（作った 時の Excel の 数を $Path.txt に）／-Hiroi で 3行目の 1501列目にも 打つ（★1024列を 越えると Excel は 行の 頭を 33バイトで 書く★）
#   hakaru … $Path（直した .xlsb）を 開き、$Rows の 行の 高さ・隠し・段・太字・A/C/BES の 値を 出す
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
try {
  if ($Mode -eq 'make') {
    $wb = $x.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Cells.Item(1, 1).Value2 = 1
    $ws.Rows.Item(1).Font.Bold = $true
    $ws.Rows.Item(1).RowHeight = 40
    $ws.Rows.Item(1).OutlineLevel = 3
    $ws.Rows.Item(1).Hidden = $true
    $ws.Cells.Item(3, 2).Value2 = 2
    if ($Hiroi) { $ws.Cells.Item(3, 1501).Value2 = 3 }
    $res = @()
    foreach ($r in 1..7) { $rr = $ws.Rows.Item($r); $res += "row $r h=$($rr.RowHeight) hid=$($rr.Hidden) lv=$($rr.OutlineLevel) bold=$($rr.Font.Bold)" }
    $res += "std h=$($ws.StandardHeight)"
    $res | Out-File -Encoding utf8 ($Path + '.txt')
    $wb.SaveAs($Path + '.xlsb', 50)
    $wb.Close($false)
  } elseif ($Mode -eq 'hakaru') {
    $wb = $x.Workbooks.Open($Path, 0, $true)
    $out = @()
    foreach ($ws in $wb.Worksheets) {
      $out += "sheet $($ws.Name) std=$($ws.StandardHeight)"
      foreach ($r in ($Rows -split ',')) {
        $rr = $ws.Rows.Item([int]$r)
        $out += "row $r h=$($rr.RowHeight) hid=$($rr.Hidden) lv=$($rr.OutlineLevel) bold=$($ws.Cells.Item([int]$r, 3).Font.Bold) A=$($ws.Cells.Item([int]$r, 1).Value2) C=$($ws.Cells.Item([int]$r, 3).Value2) BES=$($ws.Cells.Item([int]$r, 1501).Value2)"
      }
    }
    $wb.Close($false)
    $out
  } else { throw "Mode は make か hakaru" }
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
if ($Mode -eq 'make') {
  # Excel は サインインの 名前と 保存した 道を 入れるので、閉じた 後に scripts/sakusha-mihon.mjs で 替える
  & node (Join-Path $PSScriptRoot '../scripts/sakusha-mihon.mjs') ($Path + '.xlsb')
  if ($LASTEXITCODE -ne 0) { throw 'sakusha-mihon が 赤' }
}
