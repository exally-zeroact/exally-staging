# toru-28c: XLOOKUP/XMATCH no 5,6 tsume (ayamari no narabi, sora, seiki hyougen 3)
$ErrorActionPreference = 'Stop'; $out = $args[0]
$err = @{ -2146826238="#CALC!"; -2146826288='#NULL!'; -2146826281='#DIV/0!'; -2146826273='#VALUE!'; -2146826265='#REF!'; -2146826259='#NAME?'; -2146826252='#NUM!'; -2146826246='#N/A' }
function V($c) { $v = $c.Value2; if ($v -is [int]) { return $err[$v] }; if ($v -eq $null) { return '(sora)' }; return [string]$v }
function T($c) { $v = $c.Value2; if ($v -is [string]) {'String'} elseif ($v -is [double]) {'Double'} elseif ($v -is [bool]) {'Boolean'} elseif ($v -is [int]) {'Error'} else {'Empty'} }
$soto = @(
 '=XLOOKUP(TRUE,K1:K3,L1:L3,"n",1)','=XLOOKUP("c",K1:K4,L1:L4,"n",1)','=XLOOKUP(9,K1:K4,L1:L4,"n",-1)','=XLOOKUP(TRUE,K1:K4,L1:L4,"n",-1)','=XLOOKUP(3,M1:M3,N1:N3,"n",-1)','=XLOOKUP(3,M1:M3,N1:N3,"n",1)','=XLOOKUP(0.5,M1:M3,N1:N3,"n",-1)','=XLOOKUP(-1,M1:M3,N1:N3,"n",1)','=XLOOKUP(0,M1:M3,N1:N3,"n",0)','=XMATCH(9,B1:B5,1)','=XMATCH("c",K1:K4,1)','=XMATCH(3,H1:H6,-1,-1)','=XMATCH(9,H1:H6,1)','=XMATCH(TRUE,K1:K3,1)','=XMATCH(-1,M1:M3,1)','=XLOOKUP("a.c",G1:G3,B1:B3,"n",3)','=XLOOKUP("^ab$",G1:G3,B1:B3,"n",3)','=XLOOKUP("A.C",G1:G3,B1:B3,"n",3)','=XLOOKUP("[",G1:G3,B1:B3,"n",3)','=XLOOKUP(2,B1:B5,B1:B5,"n",3)','=XLOOKUP("2",B1:B5,B1:B5,"n",3)','=XMATCH("a.c",G1:G3,3)','=XLOOKUP("a",G1:G3,B1:B3,"n",3,-1)','=XLOOKUP(3,B1:B5,B1:B5,"n",0,-2)','=XLOOKUP(30,I1:I6,H1:H6,"n",0,2)','=XLOOKUP(2,H1:H6,I1:I6,"n",0,-1.5)','=XLOOKUP(3,B1:B5,B1:B5,"n",TRUE)','=XLOOKUP(3,B1:B5,B1:B5,"n","x")','=XLOOKUP(3,B1:B5,B1:B5,"n",1,0)','=XLOOKUP(2,B1:B5,B1:B5,"n",2)','=XLOOKUP(3,B1:B5,B1:B5,"n",4)','=XLOOKUP(3,B1:B5,B1:B5,"n",-1.5)','=XLOOKUP(3,B1:B5,B1:B5,"n",1,-1.5)'
)
$x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $ws.Range("B1").Value2 = [double]1; $ws.Range("B2").Value2 = [double]2; $ws.Range("B3").Formula = '=1/0'; $ws.Range("B4").Value2 = [double]4; $ws.Range("B5").Value2 = [double]5
  $ws.Range("C1").Value2 = [double]7; $ws.Range("C2").Value2 = [double]1
  $ws.Range("F1").Value2 = $true; $ws.Range("G1").Value2 = "x"; $ws.Range("G2").Value2 = "abc"; $ws.Range("G3").Value2 = "ab"; $ws.Range("H1").Value2 = [double]2; $ws.Range("H2").Value2 = "a"; $ws.Range("H3").Value2 = [double]4; $ws.Range("H4").Value2 = [double]2; $ws.Range("H5").Value2 = "c"; $ws.Range("H6").Value2 = $true; foreach ($j in 1..6) { $ws.Range("I$j").Value2 = [double](10*$j) }
  $ws.Range("K1").Value2 = [double]2; $ws.Range("K2").Formula = '=NA()'; $ws.Range("K3").Value2 = "b"; $ws.Range("K4").Value2 = $true; foreach ($j in 1..4) { $ws.Range("L$j").Value2 = [double](10*$j) }; $ws.Range("M1").Value2 = [double]1; $ws.Range("M3").Value2 = [double]5; foreach ($j in 1..3) { $ws.Range("N$j").Value2 = [double](10*$j) }
  $lines = @("# Excel $($x.Version) build $($x.Build) / B1:B5=1,2,=1/0,4,5 / C1:C2=7,1 / F1=TRUE / G1:G3=x,abc,ab / H1:H6=2,a,4,2,c,TRUE / I1:I6=10..60 / K1:K4=2,=NA(),b,TRUE / L1:L4=10..40 / M1:M3=1,(sora),5 / N1:N3=10..30 / toru-28c.ps1 2026-11-09")
  foreach ($f in $soto) { $c = $ws.Range("P1"); try { $c.Formula2 = $f } catch { $lines += "-`t$f`t(ukenai)`tz0=-`tkata=-"; continue }; $x.CalculateFull(); $w = $ws.Range("Q1"); $w.Formula = ('=(' + $c.Address($false,$false) + ')=0'); $x.Calculate(); $z = [string]$w.Value2; $w.ClearContents(); $lines += "-`t$f`t$(V $c)`tz0=$z`tkata=$(T $c)"; $c.ClearContents() }
  [IO.File]::WriteAllText($out, ($lines -join "`n") + "`n", (New-Object Text.UTF8Encoding $false)); $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
'ok'
