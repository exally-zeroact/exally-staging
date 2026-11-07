# toru-utsushiji.ps1 — 条件の写し字（* ? ~）を実Excelに聞く（2026-11-07・値の直し2の下調べ）
# 盤 A1:A8 = apple / Apple / ab / a*b / 5 / (空) / a?c / 15 ・ B1:B8 = 1..8
$ErrorActionPreference = 'Stop'
$出 = $args[0]
$誤 = @{ -2146826288='#NULL!'; -2146826281='#DIV/0!'; -2146826273='#VALUE!'; -2146826265='#REF!'; -2146826259='#NAME?'; -2146826252='#NUM!'; -2146826246='#N/A'; -2146826238='#GETTING_DATA'; -2146826229='#SPILL!'; -2146826215='#CALC!' }
$式たち = @(
 '=COUNTIF(A1:A8,"*")','=COUNTIF(A1:A8,"<>*")','=COUNTIF(A1:A8,"=*")','=COUNTIF(A1:A8,"a*")','=COUNTIF(A1:A8,"A*")',
 '=COUNTIF(A1:A8,"?")','=COUNTIF(A1:A8,"??")','=COUNTIF(A1:A8,"a?c")','=COUNTIF(A1:A8,"a~?c")','=COUNTIF(A1:A8,"a~*b")',
 '=COUNTIF(A1:A8,"~*")','=COUNTIF(A1:A8,"*b")','=COUNTIF(A1:A8,"<>a*")','=COUNTIF(A1:A8,">a*")','=COUNTIF(A1:A8,"1*")',
 '=COUNTIF(A1:A8,"5*")','=COUNTIF(A1:A8,"<>")','=COUNTIF(A1:A8,"~")','=COUNTIF(A1:A8,"a~")','=COUNTIF(A1:A8,"*~*")',
 '=SUMIF(A1:A8,"a*",B1:B8)','=SUMIF(A1:A8,"<>*",B1:B8)','=COUNTIFS(A1:A8,"a*",B1:B8,">2")','=AVERAGEIF(A1:A8,"?*",B1:B8)',
 '=COUNTIF(A:A,"<>*")','=COUNTIF(A:A,"*")','=COUNTIF(A:A,"a*")'
)
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $値 = @('apple','Apple','ab',"'a*b",'5','','a?c','15')
  for ($i=0; $i -lt 8; $i++) { if ($値[$i] -ne '') { $ws.Cells.Item($i+1,1).Formula = $値[$i] }; $ws.Cells.Item($i+1,2).Value2 = $i+1 }
  $行 = @("# 実Excel $($x.Version) build $($x.Build) / 盤 A1:A8=apple,Apple,ab,a*b,5(数),空,a?c,15(数) B1:B8=1..8")
  $r = 1
  foreach ($f in $式たち) {
    $c = $ws.Cells.Item($r, 4); $c.Formula = $f; $x.CalculateFull()
    $v = $c.Value2
    if ($v -is [int]) { $v = $誤[$v] }
    $w = $ws.Cells.Item($r, 5); $w.Formula = ('=(' + $c.Address($false,$false) + ')=0'); $x.Calculate(); $z = [string]$w.Value2; $w.ClearContents()
    $cv = $c.Value2; $t = if ($cv -is [string]) {'String'} elseif ($cv -is [double]) {'Double'} elseif ($cv -is [bool]) {'Boolean'} elseif ($cv -is [int]) {'Error'} else {'Empty'}
    $行 += "$f`t$v`tz0=$z`tkata=$t"; $c.ClearContents(); $r++
  }
  [IO.File]::WriteAllLines($出, $行, (New-Object Text.UTF8Encoding $false))
  $wb.Close($false)
} finally { $x.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($x) | Out-Null }
'ok'
