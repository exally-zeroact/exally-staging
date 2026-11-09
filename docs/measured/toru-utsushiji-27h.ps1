# toru-27h: XLOOKUP no ookisa no chigai / AGGREGATE ni XLOOKUP no mitsukaranai / TAKE no 1 masu no ayamari / Nai to jika no ayamari no jun
$ErrorActionPreference = 'Stop'; $out = $args[0]
$err = @{ -2146826288='#NULL!'; -2146826281='#DIV/0!'; -2146826273='#VALUE!'; -2146826265='#REF!'; -2146826259='#NAME?'; -2146826252='#NUM!'; -2146826246='#N/A' }
function V($c) { $v = $c.Value2; if ($v -is [int]) { return $err[$v] }; if ($v -eq $null) { return '(sora)' }; return [string]$v }
function T($c) { $v = $c.Value2; if ($v -is [string]) {'String'} elseif ($v -is [double]) {'Double'} elseif ($v -is [bool]) {'Boolean'} elseif ($v -is [int]) {'Error'} else {'Empty'} }
$soto = @(
 '=XLOOKUP(1,C1:C2,B1:B5)','=XLOOKUP(7,C1:C2,B1:B2)','=XLOOKUP(1,C1:C2,B1:B5,9)','=XLOOKUP(1,B1:B5,C1:C2)','=XLOOKUP(99,C1:C2,B1:B5)','=XLOOKUP(2,B1:B5,C1:D5)','=XLOOKUP(7,C1:D1,B1:C1)','=XLOOKUP(7,C1:D1,B1:B2)',
 '=AGGREGATE(9,6,XLOOKUP(99,B1:B5,B1:B5))','=AGGREGATE(9,6,XLOOKUP(99,B1:B5,B1:B5,B4))','=AGGREGATE(9,6,XLOOKUP(1,C1:C2,B1:B5))','=AGGREGATE(9,6,TAKE(B3:B5,1))','=AGGREGATE(9,6,TAKE(B1:B5,,1))','=AGGREGATE(9,6,DROP(B1:B5,-4))',
 '=AGGREGATE(9,6,XLOOKUP(1,B1:B5,B1:B5),TAKE(B1:B5,1))','=AGGREGATE(9,6,TAKE(B1:B5,"2"))','=AGGREGATE(9,6,TAKE(B1:B5,2.7))','=ISREF(XLOOKUP(2,B1:B5,B1:B5))','=ISREF(TAKE(B1:B5,2))','=ISREF(DROP(B1:B5,1))'
)
$x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $ws.Range("B1").Value2 = [double]1; $ws.Range("B2").Value2 = [double]2; $ws.Range("B3").Formula = '=1/0'; $ws.Range("B4").Value2 = [double]4; $ws.Range("B5").Value2 = [double]5
  $ws.Range("C1").Value2 = [double]7; $ws.Range("C2").Value2 = [double]1
  $lines = @("# Excel $($x.Version) build $($x.Build) / B1:B5=1,2,=1/0,4,5 / C1:C2=7,1 / toru-27h.ps1 2026-11-09")
  foreach ($f in $soto) { $c = $ws.Range("K1"); try { $c.Formula2 = $f } catch { $lines += "-`t$f`t(ukenai)`tz0=-`tkata=-"; continue }; $x.CalculateFull(); $w = $ws.Range("L1"); $w.Formula = ('=(' + $c.Address($false,$false) + ')=0'); $x.Calculate(); $z = [string]$w.Value2; $w.ClearContents(); $lines += "-`t$f`t$(V $c)`tz0=$z`tkata=$(T $c)"; $c.ClearContents() }
  $s2 = $wb.Worksheets.Add([Type]::Missing, $ws); $s2.Name = 'Nai'; $s2.Range("B2").Value2 = [double]100
  $k = @('=AGGREGATE(9,6,1/0,Nai!B2)','=AGGREGATE(9,6,Nai!B2,1/0)')
  $cs = @(); $i = 2; foreach ($f in $k) { $c = $ws.Range("K$i"); $c.Formula2 = $f; $cs += ,@($c, $f); $i++ }
  $x.CalculateFull(); $mae = @{}; foreach ($p in $cs) { $mae[$p[1]] = V $p[0] }
  $s2.Delete(); $x.CalculateFull()
  foreach ($p in $cs) { $c = $p[0]; $lines += "-`t$($p[1]) (Nai wo kesu mae $($mae[$p[1]])) ato no shiki $($c.Formula)`t$(V $c)`tz0=-`tkata=$(T $c)" }
  [IO.File]::WriteAllText($out, ($lines -join "`n") + "`n", (New-Object Text.UTF8Encoding $false)); $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
'ok'
