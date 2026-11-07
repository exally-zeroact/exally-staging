# toru-utsushiji-6.ps1 - kazu no jouken (kata ga kazu / shingi) + hizuke/% no migi + MATCH no ~~ (2026-11-07 atai-naoshi-2 honban-mae taiketsu nin'i A B C)
# ban8 A1:A6 = 0 / =" " / 16 / 'x / 5 / '5 ; B1:B6 = 1..6 ; Z1 = 5 (kazu) ; Z2 = TRUE
# ban9 C1:C3 = 45000 / 46000 / 0.5 ; D1:D3 = 1..3
# ban10 E1:E10 = apple,#N/A,TRUE,="",a~b," x",x,5,#DIV/0!,(kara)
# F1 = 'a~b
$ErrorActionPreference = 'Stop'
$out = $args[0]
$err = @{ -2146826288='#NULL!'; -2146826281='#DIV/0!'; -2146826273='#VALUE!'; -2146826265='#REF!'; -2146826259='#NAME?'; -2146826252='#NUM!'; -2146826246='#N/A' }
$fs = @(
 '=COUNTIF(A1:A6,5)','=COUNTIF(A1:A6,Z1)','=SUMIF(A1:A6,5,B1:B6)','=COUNTIF(A1:A6,0)','=COUNTIF(A1:A6,16)',
 '=COUNTIF(E1:E10,TRUE)','=COUNTIF(E1:E10,Z2)','=COUNTIF(E1:E10,5)','=COUNTIF(E1:E10,FALSE)',
 '=COUNTIF(C1:C3,"<=2025/1/1")','=COUNTIF(C1:C3,"=2025/1/1")','=COUNTIF(C1:C3,"2025/1/1")','=COUNTIF(C1:C3,">=5%")','=COUNTIF(C1:C3,"50%")','=COUNTIF(C1:C3,"<12:00")','=COUNTIF(C1:C3,"12:00")',
 '=MATCH("a~~b",F1:F1,0)','=MATCH("a~b",F1:F1,0)'
)
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $a = @('0','=" "','16',"'x",'5',"'5")
  for ($i=0; $i -lt 6; $i++) { $ws.Cells.Item($i+1,1).Formula = $a[$i]; $ws.Cells.Item($i+1,2).Value2 = [double]($i+1) }
  $ws.Range('Z1').Value2 = [double]5; $ws.Range('Z2').Formula = 'TRUE'
  $c3 = @(45000, 46000, 0.5)
  for ($i=0; $i -lt 3; $i++) { $ws.Cells.Item($i+1,3).Value2 = [double]$c3[$i]; $ws.Cells.Item($i+1,4).Value2 = [double]($i+1) }
  $e = @('apple','=NA()','TRUE','=""',"'a~b","' x",'x','5','=1/0','')
  for ($i=0; $i -lt 10; $i++) { if ($e[$i] -ne '') { $ws.Cells.Item($i+1,5).Formula = $e[$i] } }
  $ws.Range('F1').Formula = "'a~b"
  $lines = @("# Excel $($x.Version) build $($x.Build) / ban8 A1:A6=0,"" "",16,x,5(kazu),'5(ji) B1:B6=1..6 Z1=5 Z2=TRUE / ban9 C1:C3=45000,46000,0.5 / ban10 E1:E10=apple,#N/A,TRUE,="""",a~b,"" x"",x,5,#DIV/0!,kara / F1='a~b")
  $r = 1
  foreach ($f in $fs) {
    $c = $ws.Cells.Item($r, 8); $c.Formula = $f; $x.CalculateFull()
    $val = $c.Value2; if ($val -is [int]) { $val = $err[$val] }
    $w = $ws.Cells.Item($r, 9); $w.Formula = ('=(' + $c.Address($false,$false) + ')=0'); $x.Calculate(); $z = [string]$w.Value2; $w.ClearContents()
    $cv = $c.Value2; $t = if ($cv -is [string]) {'String'} elseif ($cv -is [double]) {'Double'} elseif ($cv -is [bool]) {'Boolean'} elseif ($cv -is [int]) {'Error'} else {'Empty'}
    $lines += "$f`t$val`tz0=$z`tkata=$t"; $c.ClearContents(); $r++
  }
  [IO.File]::WriteAllLines($out, $lines, (New-Object Text.UTF8Encoding $false))
  $wb.Close($false)
} finally { $x.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($x) | Out-Null }
'ok'
