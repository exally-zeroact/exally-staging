# toru-28: XLOOKUP no 5 tsume / marugoto no 2 jigen / TAKE no kazu ga TRUE no masu
$ErrorActionPreference = 'Stop'; $out = $args[0]
$err = @{ -2146826238="#CALC!"; -2146826288='#NULL!'; -2146826281='#DIV/0!'; -2146826273='#VALUE!'; -2146826265='#REF!'; -2146826259='#NAME?'; -2146826252='#NUM!'; -2146826246='#N/A' }
function V($c) { $v = $c.Value2; if ($v -is [int]) { return $err[$v] }; if ($v -eq $null) { return '(sora)' }; return [string]$v }
function T($c) { $v = $c.Value2; if ($v -is [string]) {'String'} elseif ($v -is [double]) {'Double'} elseif ($v -is [bool]) {'Boolean'} elseif ($v -is [int]) {'Error'} else {'Empty'} }
$soto = @(
 '=XLOOKUP(3,B1:B5,B1:B5,"n",1)','=XLOOKUP(3,B1:B5,B1:B5,"n",-1)','=XLOOKUP(3,B1:B5,B1:B5,,1)','=XLOOKUP(3,B1:B5,B1:B5,,-1)','=XLOOKUP(9,B1:B5,B1:B5,"n",1)','=XLOOKUP(0,B1:B5,B1:B5,"n",-1)','=XLOOKUP(4,B1:B5,B1:B5,"n",1)','=XLOOKUP(3,B1:B5,B1:B5,"n",0)','=XLOOKUP(3,B1:B5,B1:B5,"n",3)','=XLOOKUP(3,B1:B5,B1:B5,"n",2)','=XLOOKUP("a*",G1:G3,B1:B3,"n",2)','=XLOOKUP(3,B1:B5,B1:B5,"n",1,-1)','=XLOOKUP(3,B1:B5,B1:B5,"n",-1,-1)','=AGGREGATE(9,6,XLOOKUP(3,B1:B5,B1:B5,"n",1))','=XLOOKUP(2,B:C,B:C)','=XLOOKUP(2,B:C,B:B)','=XLOOKUP(2,1:2,1:2)','=ISNA(XLOOKUP(2,B:C,B:C))','=ROWS(TAKE(B1:B5,F1))','=ROWS(TAKE(B1:B5,TRUE))','=ROWS(DROP(B1:B5,F1))','=TAKE(B1:B5,F1)'
)
$x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $ws.Range("B1").Value2 = [double]1; $ws.Range("B2").Value2 = [double]2; $ws.Range("B3").Formula = '=1/0'; $ws.Range("B4").Value2 = [double]4; $ws.Range("B5").Value2 = [double]5
  $ws.Range("C1").Value2 = [double]7; $ws.Range("C2").Value2 = [double]1
  $ws.Range("F1").Value2 = $true; $ws.Range("G1").Value2 = "x"; $ws.Range("G2").Value2 = "abc"; $ws.Range("G3").Value2 = "ab"
  $lines = @("# Excel $($x.Version) build $($x.Build) / B1:B5=1,2,=1/0,4,5 / C1:C2=7,1 / F1=TRUE / G1:G3=x,abc,ab / toru-28.ps1 2026-11-09")
  foreach ($f in $soto) { $c = $ws.Range("K1"); try { $c.Formula2 = $f } catch { $lines += "-`t$f`t(ukenai)`tz0=-`tkata=-"; continue }; $x.CalculateFull(); $w = $ws.Range("L1"); $w.Formula = ('=(' + $c.Address($false,$false) + ')=0'); $x.Calculate(); $z = [string]$w.Value2; $w.ClearContents(); $lines += "-`t$f`t$(V $c)`tz0=$z`tkata=$(T $c)"; $c.ClearContents() }
  [IO.File]::WriteAllText($out, ($lines -join "`n") + "`n", (New-Object Text.UTF8Encoding $false)); $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
'ok'
