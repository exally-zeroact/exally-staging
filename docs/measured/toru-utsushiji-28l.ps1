# toru-28l: kenn 4h modoshita sekkei no settei-mae taiketsu (tome1 = 4tsume ga hanni/hairetsu, tome2 = n>1 no ban no hashi)
# shiki wa 1 hon zutsu atarashii sheet no 'basho' ni oku / HasSpill to SpillingToRange wo toru
$ErrorActionPreference = 'Stop'; $out = $args[0]
$err = @{ -2146826238="#CALC!"; -2146826288='#NULL!'; -2146826281='#DIV/0!'; -2146826273='#VALUE!'; -2146826265='#REF!'; -2146826259='#NAME?'; -2146826252='#NUM!'; -2146826246='#N/A'; -2146826243='#SPILL!' }
function V($v) { if ($v -is [int]) { return $err[$v] }; if ($v -eq $null) { return '(sora)' }; return [string]$v }
function T($v) { if ($v -is [string]) {'String'} elseif ($v -is [double]) {'Double'} elseif ($v -is [bool]) {'Boolean'} elseif ($v -is [int]) {'Error'} else {'Empty'} }
$soto = @(
 @('P10','=XLOOKUP(99,B1:B5,B1:B5,B1:B3)'),
 @('E2','=XLOOKUP(99,B1:B5,B1:B5,B1:B3)'),
 @('E3','=XLOOKUP(99,B1:B5,B1:B5,B1:B3)'),
 @('P10','=XLOOKUP(99,B1:B5,C1:C5,B1:B3)'),
 @('P10','=XLOOKUP(99,B1:B5,C1:C5,{8;9})'),
 @('P10','=XLOOKUP(99,B1:B5,C1:D5,{8,9})'),
 @('P10','=XLOOKUP(99,B1:B5,C1:C5,{8,9})'),
 @('P10','=ROWS(XLOOKUP(99,B1:B5,C1:C5,B1:B3))'),
 @('P10','=ISREF(XLOOKUP(99,B1:B5,C1:C5,B1:B3))'),
 @('P10','=XLOOKUP(2,B1:B5,C1:XFD5)'),
 @('P10','=SUM(XLOOKUP(2,B1:B5,C1:XFD5))'),
 @('P10','=XLOOKUP(2,B1:B5,C1:XFC5)')
)
$x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add(); $lines = @("# Excel " + $x.Version + " build " + $x.Build + " / B1:B5=1,2,=1/0,4,5 / C1:C2=7,1 / XFD2=9 / shiki wa basho no masu / 1 hon zutsu atarashii sheet / toru-28l.ps1 2026-11-09")
  foreach ($p in $soto) {
    $at = $p[0]; $f = $p[1]
    $ws = $wb.Worksheets.Add()
    $ws.Range("B1").Value2 = [double]1; $ws.Range("B2").Value2 = [double]2; $ws.Range("B3").Formula = '=1/0'; $ws.Range("B4").Value2 = [double]4; $ws.Range("B5").Value2 = [double]5
    $ws.Range("C1").Value2 = [double]7; $ws.Range("C2").Value2 = [double]1; $ws.Range("XFD2").Value2 = [double]9
    $c = $ws.Range($at); $c.Formula2 = $f; $x.Calculate()
    $v = $c.Value2; $sp = 'nashi'
    $sita = V $c.Offset(1, 0).Value2; $migi = V $c.Offset(0, 1).Value2
    if ($c.HasSpill) { $sp = $c.SpillingToRange.Address($false, $false) }
    $lines += ($at + "`t" + $f + "`t" + (V $v) + "`tafure=" + $sp + "`tshita=" + $sita + "`tmigi=" + $migi + "`tkata=" + (T $v))
  }
  [IO.File]::WriteAllLines($out, $lines, (New-Object Text.UTF8Encoding($false))); 'ok'
} finally { $wb.Close($false); $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
