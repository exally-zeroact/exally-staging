# verify-workbook-excel.ps1 - 「うちが書き出した xlsx を実Excelで開いて再計算し、golden と一致するか」を全ケースで見る。
#   往復検証の(3)を 221 ケース全部でやる版。Windows + Excel が要るので CI では走らない。
#
#   手順: node tests/xlsx-harness/build-libre-input.mjs tests/xlsx-harness/tmp/cases.xlsx
#         pwsh -File tools/verify-workbook-excel.ps1
#
#   ここが通る = ①xlsx-io.js の書き出しが Excel に正しく読める(★新関数の _xlfn. も含めて)
#                ②golden が COM 経由だけでなく「ファイル経由」でも同じ答えになる、の二重確認。
$ErrorActionPreference = 'Stop'
$ROOT = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$xlsx = Join-Path $ROOT 'tests\xlsx-harness\tmp\cases.xlsx'
if (-not (Test-Path $xlsx)) { throw "先に build-libre-input.mjs を実行してください: $xlsx が無い" }
$order = [System.IO.File]::ReadAllText("$xlsx.order.json", [Text.Encoding]::UTF8) | ConvertFrom-Json
$goldPath = Get-ChildItem (Join-Path $ROOT 'tests\xlsx-harness\golden') -Filter 'excel-*.json' | Select-Object -First 1
$gold = [System.IO.File]::ReadAllText($goldPath.FullName, [Text.Encoding]::UTF8) | ConvertFrom-Json

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($xlsx)
$xl.CalculateFullRebuild()
$ws = $wb.Worksheets.Item(1)

$ng = 0; $bad = @()
for ($i = 0; $i -lt $order.Count; $i++) {
  $id = $order[$i]
  $v = $ws.Cells.Item($i + 1, 20).Value2
  $g = $gold.cases.PSObject.Properties[$id].Value
  $got = if ($null -eq $v) { '' } elseif ($v -is [bool]) { if ($v) { 'TRUE' } else { 'FALSE' } } else { [string]$v }
  $want = if ($g.t -eq 'b') { if ($g.v) { 'TRUE' } else { 'FALSE' } } elseif ($g.t -eq 'e') { [string]$g.code } else { [string]$g.v }
  $ok = if ($g.t -eq 'n' -or $g.t -eq 'e') { [math]::Abs([double]$got - [double]$want) -le 1e-9 } else { $got -eq $want }
  if (-not $ok) { $ng++; if ($bad.Count -lt 20) { $bad += "  $id : ファイル経由=$got  golden=$want (t=$($g.t))" } }
}
$wb.Close($false); $xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

"書き出したブックを実Excelで再計算 -> 一致 $($order.Count - $ng) / $($order.Count)  (golden: $($goldPath.Name))"
if ($ng) { "★不一致 $ng 件:"; $bad | ForEach-Object { $_ }; exit 1 }
exit 0
