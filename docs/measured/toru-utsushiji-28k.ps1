# toru-28k: kenn 4h tataki naoshi 3 (1 masu no afure no moto no sanshou)
# shiki wa 1 hon zutsu atarashii sheet no P10 ni oku (tonari wo fusaganai) / HasSpill to SpillingToRange wo toru
$ErrorActionPreference = 'Stop'; $out = $args[0]
$err = @{ -2146826238="#CALC!"; -2146826288='#NULL!'; -2146826281='#DIV/0!'; -2146826273='#VALUE!'; -2146826265='#REF!'; -2146826259='#NAME?'; -2146826252='#NUM!'; -2146826246='#N/A'; -2146826243='#SPILL!' }
function V($v) { if ($v -is [int]) { return $err[$v] }; if ($v -eq $null) { return '(sora)' }; return [string]$v }
function T($v) { if ($v -is [string]) {'String'} elseif ($v -is [double]) {'Double'} elseif ($v -is [bool]) {'Boolean'} elseif ($v -is [int]) {'Error'} else {'Empty'} }
$soto = @(
 '=ISREF(XLOOKUP(1,B1,C1:C5))','=AGGREGATE(9,6,XLOOKUP(1,B1,B1:B5))','=ROW(XLOOKUP(1,B1,C3:C5))','=SUM(XLOOKUP(1,B1,C1:C2):C5)','=XLOOKUP(2,B1,C1:C5,"n")','=INDEX(XLOOKUP(1,B1,C1:C5),2)','=SUM(XLOOKUP(1,B1,C1:C5))','=ISREF(XLOOKUP(1,B1,{5;6}))','=COLUMN(XLOOKUP(1,B1,C1:D1))','=ROWS(XLOOKUP(1,B1,C1:C5))'
)
$x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add(); $lines = @("# Excel " + $x.Version + " build " + $x.Build + " / B1:B5=1,2,=1/0,4,5 / C1:C2=7,1 / S1==1/0 / shiki wa P10 / 1 hon zutsu atarashii sheet / toru-28k.ps1 2026-11-09")
  foreach ($f in $soto) {
    $ws = $wb.Worksheets.Add()
    $ws.Range("B1").Value2 = [double]1; $ws.Range("B2").Value2 = [double]2; $ws.Range("B3").Formula = '=1/0'; $ws.Range("B4").Value2 = [double]4; $ws.Range("B5").Value2 = [double]5
    $ws.Range("C1").Value2 = [double]7; $ws.Range("C2").Value2 = [double]1; $ws.Range("S1").Formula = '=1/0'
    $c = $ws.Range("P10"); $c.Formula2 = $f; $x.Calculate()
    $v = $c.Value2; $sp = 'nashi'; $p11 = V $ws.Range("P11").Value2; $q10 = V $ws.Range("Q10").Value2
    if ($c.HasSpill) { $sp = $c.SpillingToRange.Address($false, $false) }
    $lines += ("-`t" + $f + "`t" + (V $v) + "`tafure=" + $sp + "`tP11=" + $p11 + "`tQ10=" + $q10 + "`tkata=" + (T $v))
  }
  [IO.File]::WriteAllLines($out, $lines, (New-Object Text.UTF8Encoding($false))); 'ok'
} finally { $wb.Close($false); $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
