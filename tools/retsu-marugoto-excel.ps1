param([string]$In, [string]$Out)
$ErrorActionPreference = 'Stop'
$list = Get-Content -LiteralPath $In -Encoding UTF8 | Where-Object { $_ -ne '' }
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $bk = $xl.Workbooks.Add(); $sh = $bk.Worksheets.Item(1); $sh.Name = 'Sheet1'
  $sh2 = $bk.Worksheets.Add([Type]::Missing, $sh); $sh2.Name = 'Calc'
  for ($i=1; $i -le 5; $i++) { $c = $sh.Cells.Item($i,1); $c.Value2 = $i; $c = $sh.Cells.Item($i,2); $c.Value2 = $i*10 }
  $rows = @()
  # ★式は 別の板 Calc の A 列に置く＝Sheet1 の A:A・1:1 を 自分で 汚さない（板の名の無い参照は Sheet1 に 書き換える）
  $n = 0
  foreach ($f in $list) {
    $n++
    $g = $f -replace '(?<![!A-Za-z$])(\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?\d+:\$?\d+)(?![0-9(])', 'Sheet1!$1'
    $g = $g -replace 'Sheet1!Sheet1!', 'Sheet1!'
    $c = $sh2.Cells.Item($n*3, 1)
    $c.Formula2 = $g
    $xl.Calculate()
    $v = $c.Value2; $t = $c.Text
    $spill = $false; try { $spill = [bool]$c.HasSpill } catch {}
    $sz = ''; if ($spill) { $sr = $c.SpillingToRange; $sz = '' + $sr.Rows.Count + 'x' + $sr.Columns.Count; $sr = $null }
    $rows += [pscustomobject]@{ 式=$f; 置いた式=$g; 値2=('' + $v); 字=$t; 溢れ=$sz }
  }
  $rows | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $Out -Encoding UTF8
  "Excel " + $xl.Version + " 本数 " + $rows.Count
} finally {
  $c = $null; $sh = $null; $sh2 = $null
  if ($bk) { $bk.Close($false) }; $bk = $null
  $xl.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl); $xl = $null
}
