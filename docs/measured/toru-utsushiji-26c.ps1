# toru-26: ken4f ireko no shoukei (A20) wo soto no SUBTOTAL/AGGREGATE ga hazusu ka
$ErrorActionPreference = 'Stop'; $out = $args[0]
$err = @{ -2146826288='#NULL!'; -2146826281='#DIV/0!'; -2146826273='#VALUE!'; -2146826265='#REF!'; -2146826259='#NAME?'; -2146826252='#NUM!'; -2146826246='#N/A' }
function V($c) { $v = $c.Value2; if ($v -is [int]) { return $err[$v] }; if ($v -eq $null) { return '(sora)' }; return [string]$v }
function T($c) { $v = $c.Value2; if ($v -is [string]) {'String'} elseif ($v -is [double]) {'Double'} elseif ($v -is [bool]) {'Boolean'} elseif ($v -is [int]) {'Error'} else {'Empty'} }
$naka = @('="SUBTOTAL(9,A1:A5)"','=LET(s,SUBTOTAL(9,A1:A5),s)','=_xlfn.AGGREGATE(9,6,A1:A5)','=SUBTOTAL(9,A1:A5)+AGGREGATE(9,6,A1:A5)','=ROWS(SUBTOTAL)')
$soto = @('=SUBTOTAL(3,A1:A30)','=SUBTOTAL(9,A1:A30)','=AGGREGATE(3,0,A1:A30)','=AGGREGATE(9,0,A1:A30)','=COUNTA(A1:A30)')
$x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  for ($i=0; $i -lt 5; $i++) { $ws.Cells.Item($i+1,1).Value2 = [double]($i+1) }
  $lines = @("# Excel $($x.Version) build $($x.Build) / A1:A5=1..5 / A20=naka / K=soto / toru-26c.ps1 2026-11-09 (ji no naka,LET,_xlfn)")
  foreach ($nk in $naka) {
    $ws.Range("A20").Formula = $nk
    foreach ($f in $soto) { $c = $ws.Range("K1"); $c.Formula = $f; $x.CalculateFull(); $w = $ws.Range("L1"); $w.Formula = ('=(' + $c.Address($false,$false) + ')=0'); $x.Calculate(); $z = [string]$w.Value2; $w.ClearContents(); $lines += "A20 $nk`t$f`t$(V $c)`tz0=$z`tkata=$(T $c)"; $c.ClearContents() }
  }
  [IO.File]::WriteAllText($out, ($lines -join "`n") + "`n", (New-Object Text.UTF8Encoding $false)); $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
'ok'
