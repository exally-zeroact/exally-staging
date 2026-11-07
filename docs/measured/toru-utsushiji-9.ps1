# toru-utsushiji-9.ps1 - *IFS no hairetsu no jouken (.Formula / .Formula2) + hani no katachi chigai + heikin suru hani no ayamari/shingi/ji (2026-11-07 atai-naoshi-3 sekkei-mae taiketsu tome1 / nin-i 2,3)
# ban12 A1:A5=1..5 B1:B5=10,20,30,40,50 C1:C5=4,=NA(),TRUE,="5",'7 D1:D5=1,=NA(),3,4,5 E1:E2='>0,'>2 A7:E7=1..5
$ErrorActionPreference = 'Stop'
$out = $args[0]
$err = @{ -2146826288='#NULL!'; -2146826281='#DIV/0!'; -2146826273='#VALUE!'; -2146826265='#REF!'; -2146826259='#NAME?'; -2146826252='#NUM!'; -2146826246='#N/A'; -2146826238='#SPILL!'; -2146826215='#CALC!' }
$fs = @(
 '=AVERAGEIFS(B1:B5,A1:A5,{">0",">2"})',
 '=SUM(AVERAGEIFS(B1:B5,A1:A5,{">0",">2"}))',
 '=SUMPRODUCT(AVERAGEIFS(B1:B5,A1:A5,{">0",">2"}))',
 '=AVERAGEIFS(B1:B5,A1:A5,">"&{0,2})',
 '=AVERAGEIFS(B1:B5,A1:A5,{">0";">2"})',
 '=SUMIFS(B1:B5,A1:A5,{">0",">2"})',
 '=SUM(SUMIFS(B1:B5,A1:A5,{">0",">2"}))',
 '=SUM(COUNTIFS(A1:A5,{">0",">2"}))',
 '=SUM(MAXIFS(B1:B5,A1:A5,{"<3",">3"}))',
 '=SUM(MINIFS(B1:B5,A1:A5,{"<3",">3"}))',
 '=SUM(COUNTIF(A1:A5,{">0",">2"}))',
 '=SUM(SUMIF(A1:A5,{">0",">2"},B1:B5))',
 '=SUM(AVERAGEIF(A1:A5,{">0",">2"},B1:B5))',
 '=AVERAGEIFS(B1:B5,A1:A5,E1:E2)',
 '=SUM(AVERAGEIFS(B1:B5,A1:A5,E1:E2))',
 '=SUM(SUMIFS(B1:B5,A1:A5,E1:E2))',
 '=AVERAGEIFS(B1:B3,A1:A5,">0")',
 '=AVERAGEIFS(A7:E7,A1:A5,">0")',
 '=AVERAGEIFS(B1:B5,A1:A5,">0",A7:E7,">0")',
 '=SUMIFS(B1:B3,A1:A5,">0")',
 '=SUMIFS(A7:E7,A1:A5,">0")',
 '=COUNTIFS(A1:A5,">0",A7:E7,">0")',
 '=MAXIFS(B1:B3,A1:A5,">0")',
 '=AVERAGEIFS(C1:C5,A1:A5,">0")',
 '=AVERAGEIFS(C1:C5,A1:A5,"<>2")',
 '=SUMIFS(C1:C5,A1:A5,">0")',
 '=SUMIFS(C1:C5,A1:A5,"<>2")',
 '=MAXIFS(C1:C5,A1:A5,">0")',
 '=MINIFS(C1:C5,A1:A5,"<>2")',
 '=AVERAGEIF(A1:A5,">0",C1:C5)',
 '=SUMIF(A1:A5,">0",C1:C5)',
 '=COUNTIFS(D1:D5,">0")',
 '=AVERAGEIFS(B1:B5,D1:D5,">0")',
 '=SUMIFS(B1:B5,D1:D5,">0")'
)
function V($c) { $v = $c.Value2; if ($v -is [int]) { return $err[$v] }; if ($v -eq $null) { return '(sora)' }; return [string]$v }
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $c5 = @('4','=NA()','TRUE','="5"',"'7"); $d5 = @('1','=NA()','3','4','5')
  for ($i=0; $i -lt 5; $i++) {
    $ws.Cells.Item($i+1,1).Value2 = [double]($i+1); $ws.Cells.Item($i+1,2).Value2 = [double](10*($i+1))
    $ws.Cells.Item($i+1,3).Formula = $c5[$i]; $ws.Cells.Item($i+1,4).Formula = $d5[$i]; $ws.Cells.Item(7,$i+1).Value2 = [double]($i+1)
  }
  $ws.Cells.Item(1,5).Formula = "'>0"; $ws.Cells.Item(2,5).Formula = "'>2"
  $lines = @("# Excel $($x.Version) build $($x.Build) / ban12 A1:A5=1..5 B1:B5=10..50 C1:C5=4,=NA(),TRUE,=""5"",'7 D1:D5=1,=NA(),3,4,5 E1:E2='>0,'>2 A7:E7=1..5 / F=.Formula(kyuu) F2=.Formula2(ima; migi/shita=koboreta tonari)")
  $r = 10
  foreach ($f in $fs) {
    $c = $ws.Cells.Item($r, 11); $c.Formula = $f; $x.CalculateFull(); $a = V $c; $c.ClearContents()
    $c.Formula2 = $f; $x.CalculateFull(); $b = V $c; $m = V $ws.Cells.Item($r, 12); $s = V $ws.Cells.Item($r+1, 11)
    $w = $ws.Cells.Item($r, 20); $w.Formula = ('=(' + $c.Address($false,$false) + ')=0'); $x.Calculate(); $z = [string]$w.Value2; $w.ClearContents()
    $cv = $c.Value2; $t = if ($cv -is [string]) {'String'} elseif ($cv -is [double]) {'Double'} elseif ($cv -is [bool]) {'Boolean'} elseif ($cv -is [int]) {'Error'} else {'Empty'}
    $c.ClearContents()
    $lines += "$f`tF=$a`tF2=$b`tmigi=$m`tshita=$s`tz0=$z`tkata=$t"; $r += 3
  }
  [IO.File]::WriteAllLines($out, $lines, (New-Object Text.UTF8Encoding $false))
  $wb.Close($false)
} finally { $x.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($x) | Out-Null }
'ok'
