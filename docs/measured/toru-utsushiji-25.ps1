# toru-utsushiji-25.ps1 - ken4d: IF CHOOSE IFS SWITCH wo tooru retsu-marugoto + AGGREGATE + TOCOL no 3tsume ga 1masu 2026-11-08
# tsukaikata: toru-utsushiji-16.ps1 <shiki no ichiran.txt> <out.tsv>
# ban15 = ban14 (A1:A5=1..5 B1:B5=10,20,1,1,50) + E1 ="5" (shiki no ji) + E2=7 + C2=TRUE
# .Formula to .Formula2 no ryouhou de utsu: chigaeba "# (F to F2 ga chigau)" no chuu ni nokosu (kazoenai = ken5 enzan no afure) / utenai shiki wa "# (utenai)" / ushiro 2 ran = 2tsume no mado (z0= kata=)
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
  $ws.Range("E1").Formula = '="5"'; $ws.Range("E2").Value2 = [double]7; $ws.Range("C2").Value2 = $true
  $lines = @("# Excel $($x.Version) build $($x.Build) / ban15 A1:A5=1..5 B1:B5=10,20,1,1,50 E1=shiki-moji5 E2=7 C2=TRUE / LET LAMBDA no sanshou / .Formula to .Formula2 onaji dake / toru-utsushiji-25.ps1")
  $r = 10
  foreach ($f in $fs) {
    $c = $ws.Cells.Item($r, 11)
    try { $c.Formula = $f; $x.CalculateFull(); $a = V $c } catch { $a = "(utenai)" }; $c.ClearContents()
    try { $c.Formula2 = $f; $x.CalculateFull(); $b = V $c } catch { $b = "(utenai)" }
    if ($b -eq "(utenai)") { $c.ClearContents(); $lines += "# (utenai) $f"; $r += 8; continue }
    if ($a -ne $b -or $c.HasSpill) { $c.ClearContents(); $lines += "# (F to F2 ga chigau) $f F=$a F2=$b"; $r += 8; continue }
    $w = $ws.Cells.Item($r, 12); $w.Formula = ('=(' + $c.Address($false,$false) + ')=0'); $x.Calculate(); $z = [string]$w.Value2; $w.ClearContents()
    $t = T $c; $c.ClearContents()
    $lines += "$f`t$b`tz0=$z`tkata=$t"; $r += 8
  }
  [IO.File]::WriteAllText($out, ($lines -join "`n") + "`n", (New-Object Text.UTF8Encoding $false))
  $wb.Close($false)
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
'ok'
