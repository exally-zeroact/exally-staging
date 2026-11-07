# toru-utsushiji-5.ps1 - kanma no yomi + > < ga ji/shingi no masu ni ataru ka (2026-11-07 atai-naoshi-2)
# ban7 A1:A10 = '1,00 / '1,2,3 / ',5 / '1,000.5 / '-1,000 / '1,0000 / TRUE / 5 / 'b / '5,
# B1:B10 = 1..10
$ErrorActionPreference = 'Stop'
$out = $args[0]
$err = @{ -2146826281='#DIV/0!'; -2146826273='#VALUE!'; -2146826246='#N/A' }
$fs = @(
 '=COUNTIF(A1:A10,"100")','=COUNTIF(A1:A10,"123")','=COUNTIF(A1:A10,"0.5")','=COUNTIF(A1:A10,"1000.5")',
 '=COUNTIF(A1:A10,"-1000")','=COUNTIF(A1:A10,"10000")','=COUNTIF(A1:A10,"1,00")','=COUNTIF(A1:A10,",5")',
 '=COUNTIF(A1:A10,"5,")','=COUNTIF(A1:A10,"5")',
 '=COUNTIF(A1:A10,">4")','=COUNTIF(A1:A10,"<4")','=COUNTIF(A1:A10,">=1000")','=COUNTIF(A1:A10,">a")','=COUNTIF(A1:A10,"<a")',
 '=COUNTIF(A1:A10,"<z")','=COUNTIF(A1:A10,">FALSE")','=COUNTIF(A1:A10,"<>1,000")','=SUMIF(A1:A10,">0",B1:B10)'
)
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $v = @("'1,00","'1,2,3","',5","'1,000.5","'-1,000","'1,0000",'TRUE','5',"'b","'5,")
  for ($i=0; $i -lt 10; $i++) { $ws.Cells.Item($i+1,1).Formula = $v[$i]; $ws.Cells.Item($i+1,2).Value2 = $i+1 }
  $t = @(); for ($i=1; $i -le 10; $i++) { $t += [string]$ws.Cells.Item($i,1).Value2.GetType().Name }
  $lines = @("# Excel $($x.Version) build $($x.Build) / ban7 A1:A10='1,00 '1,2,3 ',5 '1,000.5 '-1,000 '1,0000 TRUE(shingi) 5(kazu) 'b '5, / B1:B10=1..10 / kata=" + ($t -join ','))
  $r = 1
  foreach ($f in $fs) {
    $c = $ws.Cells.Item($r, 4); $c.Formula = $f; $x.CalculateFull()
    $val = $c.Value2; if ($val -is [int]) { $val = $err[$val] }
    $w = $ws.Cells.Item($r, 5); $w.Formula = ('=(' + $c.Address($false,$false) + ')=0'); $x.Calculate(); $z = [string]$w.Value2; $w.ClearContents()
    $cv = $c.Value2; $t = if ($cv -is [string]) {'String'} elseif ($cv -is [double]) {'Double'} elseif ($cv -is [bool]) {'Boolean'} elseif ($cv -is [int]) {'Error'} else {'Empty'}
    $lines += "$f`t$val`tz0=$z`tkata=$t"; $c.ClearContents(); $r++
  }
  [IO.File]::WriteAllLines($out, $lines, (New-Object Text.UTF8Encoding $false))
  $wb.Close($false)
} finally { $x.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($x) | Out-Null }
'ok'
