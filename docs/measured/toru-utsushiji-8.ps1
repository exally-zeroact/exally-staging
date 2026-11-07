# toru-utsushiji-8.ps1 - AVERAGEIFS to AVERAGEIF no hikaku + kazu no jouken ga ji no masu ni ataru ka (2026-11-07 atai-naoshi-3 taiketsu nin-i 1)
# ban11 G1:G10 = 45658 / '2025/1/1 / 0.5 / '12:00 / '50% / 0 / =25/24 / =1/24 / '25:00 / 0.05 ; H1:H10 = 1,2,4,...,512 (SUMIF de dono masu ka wakaru)
$ErrorActionPreference = 'Stop'
$out = $args[0]
$err = @{ -2146826288='#NULL!'; -2146826281='#DIV/0!'; -2146826273='#VALUE!'; -2146826265='#REF!'; -2146826259='#NAME?'; -2146826252='#NUM!'; -2146826246='#N/A' }
$cr = @('2025/1/1','=2025/1/1','<>2025/1/1','>=2025/1/1','<2025/1/1','2025-1-1','12:00','=12:00','<>12:00','50%','=50%','<>50%','5%','5 %','50 %','24:00','=24:00','>=24:00','25:00','=25:00','<>25:00','1:00','0:00','<12:00')
$fs = @()
$fs = @("=AVERAGEIFS(H1:H10,G1:G10,""12:00"")","=AVERAGEIF(G1:G10,""12:00"",H1:H10)","=AVERAGEIFS(H1:H10,G1:G10,""50%"")","=AVERAGEIFS(H1:H10,G1:G10,""<>2025/1/1"")","=AVERAGEIFS(H1:H10,G1:G10,""2025/1/1"")","=AVERAGEIFS(H1:H10,G1:G10,""25:00"")","=AVERAGEIFS(H1:H10,G1:G10,"">=24:00"")","=AVERAGEIFS(H1:H10,G1:G10,""zzz"")","=COUNTIF(G1:G10,0.5)","=SUMIF(G1:G10,0.5,H1:H10)","=COUNTIF(G1:G10,""45658"")","=SUMIF(G1:G10,""45658"",H1:H10)","=SUMIF(G1:G10,45658,H1:H10)","=SUMIF(G1:G10,""0.5"",H1:H10)")
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $g = @('45658',"'2025/1/1",'0.5',"'12:00","'50%",'0','=25/24','=1/24',"'25:00",'0.05')
  for ($i=0; $i -lt 10; $i++) { $ws.Cells.Item($i+1,7).Formula = $g[$i]; $ws.Cells.Item($i+1,8).Value2 = [double][math]::Pow(2,$i) }
  $kata = @(); for ($i=0; $i -lt 10; $i++) { $v = $ws.Cells.Item($i+1,7).Value2; $kata += $(if ($v -is [string]) {'ji'} else {'kazu'}) }
  $lines = @("# Excel $($x.Version) build $($x.Build) / ban11 G1:G10=45658,'2025/1/1,0.5,'12:00,'50%,0,=25/24,=1/24,'25:00,0.05 (kata: $($kata -join ',')) H1:H10=2^0..2^9")
  $r = 1
  foreach ($f in $fs) {
    $c = $ws.Cells.Item($r, 11); $c.Formula = $f; $x.CalculateFull()
    $val = $c.Value2; if ($val -is [int]) { $val = $err[$val] }
    $w = $ws.Cells.Item($r, 12); $w.Formula = ('=(' + $c.Address($false,$false) + ')=0'); $x.Calculate(); $z = [string]$w.Value2; $w.ClearContents()
    $cv = $c.Value2; $t = if ($cv -is [string]) {'String'} elseif ($cv -is [double]) {'Double'} elseif ($cv -is [bool]) {'Boolean'} elseif ($cv -is [int]) {'Error'} else {'Empty'}
    $lines += "$f`t$val`tz0=$z`tkata=$t"; $c.ClearContents(); $r++
  }
  [IO.File]::WriteAllLines($out, $lines, (New-Object Text.UTF8Encoding $false))
  $wb.Close($false)
} finally { $x.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($x) | Out-Null }
'ok'
