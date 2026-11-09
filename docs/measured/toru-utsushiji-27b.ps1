# toru-27b: AGGREGATE no sansho no katachi (ref1,ref2,...) ni masu / tabane / kansuu no sansho / jika no atai wo mazeru
$ErrorActionPreference = 'Stop'; $out = $args[0]
$err = @{ -2146826288='#NULL!'; -2146826281='#DIV/0!'; -2146826273='#VALUE!'; -2146826265='#REF!'; -2146826259='#NAME?'; -2146826252='#NUM!'; -2146826246='#N/A' }
function V($c) { $v = $c.Value2; if ($v -is [int]) { return $err[$v] }; if ($v -eq $null) { return '(sora)' }; return [string]$v }
function T($c) { $v = $c.Value2; if ($v -is [string]) {'String'} elseif ($v -is [double]) {'Double'} elseif ($v -is [bool]) {'Boolean'} elseif ($v -is [int]) {'Error'} else {'Empty'} }
$soto = @()
foreach ($o in 0..7) { foreach ($t in 'B1,B2','B3','B1,B3','B3,B1:B5','B1:B5,B3','(B1,B3)','B1:B2,B4','B1,B2,B4,B5','B1,C1','B1,D1') { $soto += "=AGGREGATE(9,$o,$t)" } }
foreach ($o in 0,3,4,6) { foreach ($fn in 1,2,3,4,5,6,7,8,10,11,12,13) { $soto += "=AGGREGATE($fn,$o,B1,B2,B4)"; $soto += "=AGGREGATE($fn,$o,B1,B3,B5)" } }
foreach ($o in 0,6) { foreach ($t in 'B1,B2,1','(B1,B2),1','B3,1','B1:B5,B1,1','B1:B5,1,1') { $soto += "=AGGREGATE(14,$o,$t)" } }
foreach ($t in '"2"','TRUE','{1,2}','B1+0','INDEX(B1:B5,2)','OFFSET(B1,1,0)','IF(TRUE,B2)','Sheet1!B2','B:B','C1','B2:B2','(B2)','((B1,B2))','INDIRECT("B2")','-B2','B2:B3 B3:B4','#N/A','A1') { $soto += "=AGGREGATE(9,6,B1:B5,$t)"; $soto += "=AGGREGATE(9,0,B1,$t)" }
$x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $ws.Range("B1").Value2 = [double]1; $ws.Range("B2").Value2 = [double]2; $ws.Range("B3").Formula = '=1/0'; $ws.Range("B4").Value2 = [double]4; $ws.Range("B5").Value2 = [double]5
  $ws.Range("C1").Value2 = 'moji'; $ws.Range("D1").Value2 = $true
  $lines = @("# Excel $($x.Version) build $($x.Build) / B1:B5=1,2,=1/0,4,5 / C1='moji' / D1=TRUE / A1=sora / toru-27b.ps1 2026-11-09")
  foreach ($f in $soto) { $c = $ws.Range("K1"); try { $c.Formula = $f } catch { $lines += "-`t$f`t(ukenai)`tz0=-`tkata=-"; continue }; $x.CalculateFull(); $w = $ws.Range("L1"); $w.Formula = ('=(' + $c.Address($false,$false) + ')=0'); $x.Calculate(); $z = [string]$w.Value2; $w.ClearContents(); $lines += "-`t$f`t$(V $c)`tz0=$z`tkata=$(T $c)"; $c.ClearContents() }
  [IO.File]::WriteAllText($out, ($lines -join "`n") + "`n", (New-Object Text.UTF8Encoding $false)); $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
'ok'
