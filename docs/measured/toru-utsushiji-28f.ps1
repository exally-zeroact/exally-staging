# toru-28f: XLOOKUP/XMATCH no 5,6 tsume (ayamari no narabi, sora, seiki hyougen 3)
$ErrorActionPreference = 'Stop'; $out = $args[0]
$err = @{ -2146826238="#CALC!"; -2146826288='#NULL!'; -2146826281='#DIV/0!'; -2146826273='#VALUE!'; -2146826265='#REF!'; -2146826259='#NAME?'; -2146826252='#NUM!'; -2146826246='#N/A' }
function V($c) { $v = $c.Value2; if ($v -is [int]) { return $err[$v] }; if ($v -eq $null) { return '(sora)' }; return [string]$v }
function T($c) { $v = $c.Value2; if ($v -is [string]) {'String'} elseif ($v -is [double]) {'Double'} elseif ($v -is [bool]) {'Boolean'} elseif ($v -is [int]) {'Error'} else {'Empty'} }
$soto = @(
 '=XLOOKUP(2,1:2,1:2)','=XLOOKUP(2,B:C,B:C)','=XMATCH(2,B:C)','=XMATCH(4,H:H)','=XLOOKUP(4,H:H,I:I,"n",0,-1)','=XMATCH(Z1,M:M,0,-1)','=XLOOKUP(Z1,M:M,N:N,"n",0,-1)','=XMATCH(Z1,M:M)','=XMATCH(9,B:B,1,2)','=XMATCH(3,B:B,1,2)','=XMATCH(2,B:B,0,2)','=XMATCH(4,H:H,0,-2)','=XLOOKUP(3,B:B,C:C,"n",1)','=XLOOKUP("a*",G:G,B:B,"n",2)','=XMATCH("^$",M1:M3,3)','=XMATCH("^$",M:M,3,-1)','=XMATCH("a*",G1:G3,2,-2)','=XMATCH("a.c",G1:G3,3,-2)','=XMATCH(TRUE,S1:S2,1)','=XMATCH(TRUE,S1:S2,1,-1)','=XMATCH(Z1,M1:M3,1)','=XMATCH(Z1,M1:M3,0,2)','=XMATCH("(?i)ABC",G1:G3,3)','=XMATCH("[[:digit:]]",B1:B5,3)','=XMATCH("*",M1:M3,2,-1)','=XMATCH(9,B:B,1)','=XMATCH(0,B:B,-1,2)'
)
$x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $ws.Range("B1").Value2 = [double]1; $ws.Range("B2").Value2 = [double]2; $ws.Range("B3").Formula = '=1/0'; $ws.Range("B4").Value2 = [double]4; $ws.Range("B5").Value2 = [double]5
  $ws.Range("C1").Value2 = [double]7; $ws.Range("C2").Value2 = [double]1
  $ws.Range("F1").Value2 = $true; $ws.Range("G1").Value2 = "x"; $ws.Range("G2").Value2 = "abc"; $ws.Range("G3").Value2 = "ab"; $ws.Range("H1").Value2 = [double]2; $ws.Range("H2").Value2 = "a"; $ws.Range("H3").Value2 = [double]4; $ws.Range("H4").Value2 = [double]2; $ws.Range("H5").Value2 = "c"; $ws.Range("H6").Value2 = $true; foreach ($j in 1..6) { $ws.Range("I$j").Value2 = [double](10*$j) }
  $ws.Range("K1").Value2 = [double]2; $ws.Range("K2").Formula = '=NA()'; $ws.Range("K3").Value2 = "b"; $ws.Range("K4").Value2 = $true; foreach ($j in 1..4) { $ws.Range("L$j").Value2 = [double](10*$j) }; $ws.Range("M1").Value2 = [double]1; $ws.Range("M3").Value2 = [double]5; foreach ($j in 1..3) { $ws.Range("N$j").Value2 = [double](10*$j) }; $ws.Range("O1").Value2 = "a"; $ws.Range("O3").Value2 = "c"; $ws.Range("R1").Value2 = $false; $ws.Range("R3").Value2 = $true; $ws.Range("S1").Formula = '=1/0'; $ws.Range("S2").Formula = '=NA()'
  $lines = @("# Excel $($x.Version) build $($x.Build) / B1:B5=1,2,=1/0,4,5 / C1:C2=7,1 / F1=TRUE / G1:G3=x,abc,ab / H1:H6=2,a,4,2,c,TRUE / I1:I6=10..60 / K1:K4=2,=NA(),b,TRUE / L1:L4=10..40 / M1:M3=1,(sora),5 / N1:N3=10..30 / O1:O3=a,(sora),c / R1:R3=FALSE,(sora),TRUE / S1:S2==1/0,=NA() / shiki wa P10 / toru-28f.ps1 2026-11-09")
  foreach ($f in $soto) { $c = $ws.Range("P10"); try { $c.Formula2 = $f } catch { $lines += "-`t$f`t(ukenai)`tz0=-`tkata=-"; continue }; $x.CalculateFull(); $w = $ws.Range("Q10"); $w.Formula = ('=(' + $c.Address($false,$false) + ')=0'); $x.Calculate(); $z = [string]$w.Value2; $w.ClearContents(); $lines += "-`t$f`t$(V $c)`tz0=$z`tkata=$(T $c)"; $c.ClearContents() }
  [IO.File]::WriteAllText($out, ($lines -join "`n") + "`n", (New-Object Text.UTF8Encoding $false)); $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
'ok'
