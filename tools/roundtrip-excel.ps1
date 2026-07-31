# roundtrip-excel.ps1 - 往復検証の(3): 書き出した xlsx を「実Excel」で開いて再計算し、
#   ・式がそのまま残っているか
#   ・再計算した値が、書き出す時に入れておいた計算済みの値と一致するか
#   を確かめる。Windows + Excel が要るので CI では走らせない(CIは(1)(2)まで)。
#
#   前に node tests/xlsx-harness/roundtrip.test.mjs を走らせて tmp/roundtrip.xlsx を作っておくこと。
$ErrorActionPreference = 'Stop'
$ROOT = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$xlsx = Join-Path $ROOT 'tests\xlsx-harness\tmp\roundtrip.xlsx'
$expPath = Join-Path $ROOT 'tests\xlsx-harness\tmp\roundtrip-expected.json'
if (-not (Test-Path $xlsx)) { throw "先に node tests/xlsx-harness/roundtrip.test.mjs を実行してください: $xlsx が無い" }
$expected = [System.IO.File]::ReadAllText($expPath, [Text.Encoding]::UTF8) | ConvertFrom-Json

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($xlsx)
$xl.CalculateFullRebuild()          # ★キャッシュ値を信じず、実Excelに計算し直させる
$ws = $wb.Worksheets.Item(1)

"sheet: " + $ws.Name
"| セル | 式(Excelが読んだ物) | 再計算値 | 期待値 | 判定 |"
"|---|---|---|---|---|"
$ng = 0
foreach ($p in $expected.PSObject.Properties) {
  $addr = $p.Name
  $cell = $ws.Range($addr)
  $gotF = [string]$cell.Formula
  $gotV = $cell.Value2
  $wantF = [string]$p.Value.f
  $wantV = $p.Value.v
  $fOk = ($gotF -eq $wantF)
  $vOk = $false
  if ($wantV -is [string]) { $vOk = ([string]$gotV -eq [string]$wantV) }
  else { $vOk = ([math]::Abs([double]$gotV - [double]$wantV) -le 1e-9) }
  if (-not ($fOk -and $vOk)) { $ng++ }
  $mark = if ($fOk -and $vOk) { 'OK' } elseif (-not $fOk) { '★式が変わった' } else { '★値が違う' }
  "| $addr | $gotF | $gotV | $wantV | $mark |"
}
$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

""
if ($ng -eq 0) { "OK: 実Excelで開いて再計算しても、式も値も一致した ($(@($expected.PSObject.Properties).Count) 本)"; exit 0 }
else { "★NG: $ng 本が不一致"; exit 1 }
