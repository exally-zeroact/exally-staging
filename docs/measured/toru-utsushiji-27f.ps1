# toru-27f: AGGREGATE ni tate no teisuu no hyou / #REF! wo jika ni kaku (ita wo kesita ato no katachi)
$ErrorActionPreference = 'Stop'; $out = $args[0]
$err = @{ -2146826288='#NULL!'; -2146826281='#DIV/0!'; -2146826273='#VALUE!'; -2146826265='#REF!'; -2146826259='#NAME?'; -2146826252='#NUM!'; -2146826246='#N/A' }
function V($c) { $v = $c.Value2; if ($v -is [int]) { return $err[$v] }; if ($v -eq $null) { return '(sora)' }; return [string]$v }
function T($c) { $v = $c.Value2; if ($v -is [string]) {'String'} elseif ($v -is [double]) {'Double'} elseif ($v -is [bool]) {'Boolean'} elseif ($v -is [int]) {'Error'} else {'Empty'} }
$soto = @('=AGGREGATE(9,6,{1;2;4;5})','=AGGREGATE(19,6,{1;2;4;5},1)','=AGGREGATE(9,6,B1:B5,#REF!)','=AGGREGATE(9,6,#REF!)','=AGGREGATE(9,6,(B1,#REF!))','=AGGREGATE(9,0,B1,#REF!)','=AGGREGATE(2,6,B1:B5,#REF!)','=AGGREGATE(14,6,#REF!,1)','=AGGREGATE(9,6,IF({TRUE},B1:B2))','=AGGREGATE(9,6,INDEX({1;2;3},0))','=AGGREGATE(2,0,5)','=AGGREGATE(9,0,1/0,"2")','=AGGREGATE(9,0,"2",1/0)','=AGGREGATE(9,0,B1:B2,A20)','=AGGREGATE(9,6,LET(x,B1:B2,x))','=AGGREGATE(9,6,LET(x,{1,2},x))','=AGGREGATE(9,6,CHOOSE({1},B1:B2))')
$x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $ws.Range("B1").Value2 = [double]1; $ws.Range("B2").Value2 = [double]2; $ws.Range("B3").Formula = '=1/0'; $ws.Range("B4").Value2 = [double]4; $ws.Range("B5").Value2 = [double]5
  $ws.Range("A20").Formula = '=SUBTOTAL(9,B1:B2)'
  $lines = @("# Excel $($x.Version) build $($x.Build) / B1:B5=1,2,=1/0,4,5 / A20=SUBTOTAL(9,B1:B2) / toru-27f.ps1 2026-11-09")
  foreach ($f in $soto) { $c = $ws.Range("K1"); try { $c.Formula2 = $f } catch { $lines += "-`t$f`t(ukenai)`tz0=-`tkata=-"; continue }; $x.CalculateFull(); $w = $ws.Range("L1"); $w.Formula = ('=(' + $c.Address($false,$false) + ')=0'); $x.Calculate(); $z = [string]$w.Value2; $w.ClearContents(); $lines += "-`t$f`t$(V $c)`tz0=$z`tkata=$(T $c)"; $c.ClearContents() }
  # ita wo kesu: Sheet2 wo tsukutte sansho shite kara kesu
  $s2 = $wb.Worksheets.Add([Type]::Missing, $ws); $s2.Name = 'Nai'; $s2.Range("B2").Value2 = [double]100
  $c = $ws.Range("K2"); $c.Formula = '=AGGREGATE(9,6,B1:B5,Nai!B2)'; $x.CalculateFull(); $mae = V $c
  $s2.Delete(); $x.CalculateFull(); $lines += "-`t=AGGREGATE(9,6,B1:B5,Nai!B2) (Nai wo kesu mae $mae) ato no shiki $($c.Formula)`t$(V $c)`tz0=-`tkata=$(T $c)"
  [IO.File]::WriteAllText($out, ($lines -join "`n") + "`n", (New-Object Text.UTF8Encoding $false)); $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
'ok'
