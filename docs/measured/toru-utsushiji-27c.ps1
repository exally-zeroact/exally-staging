# toru-27c: AGGREGATE no 1tsume no hikisuu to shiki no ayamari
$ErrorActionPreference = 'Stop'; $out = $args[0]
$err = @{ -2146826288='#NULL!'; -2146826281='#DIV/0!'; -2146826273='#VALUE!'; -2146826265='#REF!'; -2146826259='#NAME?'; -2146826252='#NUM!'; -2146826246='#N/A' }
function V($c) { $v = $c.Value2; if ($v -is [int]) { return $err[$v] }; if ($v -eq $null) { return '(sora)' }; return [string]$v }
function T($c) { $v = $c.Value2; if ($v -is [string]) {'String'} elseif ($v -is [double]) {'Double'} elseif ($v -is [bool]) {'Boolean'} elseif ($v -is [int]) {'Error'} else {'Empty'} }
$soto = @()
foreach ($o in 0,6) { foreach ($t in '{1,2}','B1:B5+0','5','"5"','1/0','#N/A','TRUE','B1:B5,1/0','B1:B5,B1:B2+0','B1:B5,"2",#N/A','B1:B5,#N/A,"2"','B3,"2"','(B1,B3),B3','B1:B5 C1:C5','B2:B3 B3:B4,B1','INDEX((B1,B2),1,1,2)','IF(FALSE,B1,B3)','CHOOSE(2,B1,B3)','B1:B5,(B1,B3)') { $soto += "=AGGREGATE(9,$o,$t)" } }
foreach ($o in 0,6) { foreach ($t in '{1,2},1','B1:B5+0,1','(B1,B3),1','B3,1','B1,1') { $soto += "=AGGREGATE(14,$o,$t)"; $soto += "=AGGREGATE(15,$o,$t)" } }
$x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $ws.Range("B1").Value2 = [double]1; $ws.Range("B2").Value2 = [double]2; $ws.Range("B3").Formula = '=1/0'; $ws.Range("B4").Value2 = [double]4; $ws.Range("B5").Value2 = [double]5
  $ws.Range("C1").Value2 = 'moji'; $ws.Range("D1").Value2 = $true
  $lines = @("# Excel $($x.Version) build $($x.Build) / B1:B5=1,2,=1/0,4,5 / C1='moji' / D1=TRUE / A1=sora / toru-27c.ps1 2026-11-09")
  foreach ($f in $soto) { $c = $ws.Range("K1"); try { $c.Formula = $f } catch { $lines += "-`t$f`t(ukenai)`tz0=-`tkata=-"; continue }; $x.CalculateFull(); $w = $ws.Range("L1"); $w.Formula = ('=(' + $c.Address($false,$false) + ')=0'); $x.Calculate(); $z = [string]$w.Value2; $w.ClearContents(); $lines += "-`t$f`t$(V $c)`tz0=$z`tkata=$(T $c)"; $c.ClearContents() }
  [IO.File]::WriteAllText($out, ($lines -join "`n") + "`n", (New-Object Text.UTF8Encoding $false)); $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
'ok'
