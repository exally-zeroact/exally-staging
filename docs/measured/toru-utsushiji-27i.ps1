# toru-27i: XLOOKUP no sagasu tokoro ga 2 jigen / TAKE DROP no kazu wo masu ya TRUE de watasu
$ErrorActionPreference = 'Stop'; $out = $args[0]
$err = @{ -2146826288='#NULL!'; -2146826281='#DIV/0!'; -2146826273='#VALUE!'; -2146826265='#REF!'; -2146826259='#NAME?'; -2146826252='#NUM!'; -2146826246='#N/A' }
function V($c) { $v = $c.Value2; if ($v -is [int]) { return $err[$v] }; if ($v -eq $null) { return '(sora)' }; return [string]$v }
function T($c) { $v = $c.Value2; if ($v -is [string]) {'String'} elseif ($v -is [double]) {'Double'} elseif ($v -is [bool]) {'Boolean'} elseif ($v -is [int]) {'Error'} else {'Empty'} }
$soto = @(
 '=XLOOKUP(2,B1:C5,B1:C5)','=XLOOKUP(2,B1:C2,B1:C2)','=XLOOKUP(7,C1:D2,B1:C2)','=XLOOKUP(99,B1:C5,B1:C5)','=XLOOKUP(2,B1:C5,B1:B5)','=XLOOKUP(2,B1:C5,B1:C5,9)','=XLOOKUP(2,B1:C5,B1:C5,,0,-1)','=IFERROR(XLOOKUP(2,B1:C5,B1:C5),"x")','=AGGREGATE(9,6,XLOOKUP(2,B1:C5,B1:C5))','=XLOOKUP(2,B1:C2,B1:B2)','=AGGREGATE(9,6,TAKE(B1:B5,C2))','=AGGREGATE(9,6,DROP(B1:B5,C2))','=AGGREGATE(9,6,TAKE(B1:B5,TRUE))','=AGGREGATE(9,6,DROP(B1:B5,TRUE))','=ISREF(TAKE(B1:B5,C2))','=ISREF(DROP(B1:B5,C2))','=AGGREGATE(9,6,TAKE(B1:B5,,C2))','=AGGREGATE(9,6,TAKE(B1:B5,C2+1))','=SUM(TAKE(B1:B5,C2))','=AGGREGATE(9,6,XLOOKUP(99,B1:B5,B1:B5,,0))','=XLOOKUP(99,B1:B5,B1:B5,,0)'
)
$x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $ws.Range("B1").Value2 = [double]1; $ws.Range("B2").Value2 = [double]2; $ws.Range("B3").Formula = '=1/0'; $ws.Range("B4").Value2 = [double]4; $ws.Range("B5").Value2 = [double]5
  $ws.Range("C1").Value2 = [double]7; $ws.Range("C2").Value2 = [double]1
  $lines = @("# Excel $($x.Version) build $($x.Build) / B1:B5=1,2,=1/0,4,5 / C1:C2=7,1 / toru-27i.ps1 2026-11-09")
  foreach ($f in $soto) { $c = $ws.Range("K1"); try { $c.Formula2 = $f } catch { $lines += "-`t$f`t(ukenai)`tz0=-`tkata=-"; continue }; $x.CalculateFull(); $w = $ws.Range("L1"); $w.Formula = ('=(' + $c.Address($false,$false) + ')=0'); $x.Calculate(); $z = [string]$w.Value2; $w.ClearContents(); $lines += "-`t$f`t$(V $c)`tz0=$z`tkata=$(T $c)"; $c.ClearContents() }
  [IO.File]::WriteAllText($out, ($lines -join "`n") + "`n", (New-Object Text.UTF8Encoding $false)); $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
'ok'
