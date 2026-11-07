# toru-utsushiji-11.ps1 - *IF(S) no hani ni yomenai sanshou (B1# INDIRECT OFFSET IF CHOOSE LET) (2026-11-07 atai-naoshi-3 retsu-x no machine-chi wo hakaru)
# tsukaikata: toru-utsushiji-11.ps1 <shiki no ichiran.txt> <out.tsv>
# ban14 A1:A5=1..5 B1:B5=10,20,1,1,50 (retsu-x no zairyou + B3=1 B4=1) / .Formula to .Formula2 no ryouhou de utsu (chigaeba tomeru; utenai shiki wa "# (utenai)" no chuu ni nokosu) / ushiro 2 ran = 2tsume no mado (z0= kata=)
$ErrorActionPreference = 'Stop'
$fsfile = $args[0]; $out = $args[1]
$err = @{ -2146826288='#NULL!'; -2146826281='#DIV/0!'; -2146826273='#VALUE!'; -2146826265='#REF!'; -2146826259='#NAME?'; -2146826252='#NUM!'; -2146826246='#N/A'; -2146826238='#SPILL!'; -2146826215='#CALC!' }
$fs = Get-Content -Encoding UTF8 $fsfile | ? { $_ -ne '' }
function V($c) { $v = $c.Value2; if ($v -is [int]) { return $err[$v] }; if ($v -eq $null) { return '(sora)' }; return [string]$v }
function T($c) { $v = $c.Value2; if ($v -is [string]) {'String'} elseif ($v -is [double]) {'Double'} elseif ($v -is [bool]) {'Boolean'} elseif ($v -is [int]) {'Error'} else {'Empty'} }
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  for ($i=0; $i -lt 5; $i++) { $ws.Cells.Item($i+1,1).Value2 = [double]($i+1); $ws.Cells.Item($i+1,2).Value2 = [double](@(10,20,1,1,50)[$i]) }
  $lines = @("# Excel $($x.Version) build $($x.Build) / ban14 A1:A5=1..5 B1:B5=10,20,1,1,50 / yomenai sanshou / .Formula to .Formula2 onaji / toru-utsushiji-11.ps1")
  $r = 10
  foreach ($f in $fs) {
    $c = $ws.Cells.Item($r, 11)
    try { $c.Formula = $f; $x.CalculateFull(); $a = V $c } catch { $a = "(utenai)" }; $c.ClearContents()
    try { $c.Formula2 = $f; $x.CalculateFull(); $b = V $c } catch { $b = "(utenai)" }
    if ($a -ne "(utenai)" -and $b -ne "(utenai)" -and $a -ne $b) { throw "Formula to Formula2 ga chigau: $f F=$a F2=$b" }
    if ($b -eq "(utenai)") { $lines += "# (utenai) $f"; $r += 3; continue }
    $w = $ws.Cells.Item($r, 12); $w.Formula = ('=(' + $c.Address($false,$false) + ')=0'); $x.Calculate(); $z = [string]$w.Value2; $w.ClearContents()
    $t = T $c; $c.ClearContents()
    $lines += "$f`t$b`tz0=$z`tkata=$t"; $r += 3
  }
  [IO.File]::WriteAllText($out, ($lines -join "`n") + "`n", (New-Object Text.UTF8Encoding $false))
  $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
'ok'
