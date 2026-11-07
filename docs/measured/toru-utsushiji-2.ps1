# toru-utsushiji-2.ps1 — 条件の写し字の測り足し（2026-11-07・値の直し2 設計前 taiketsu 止め1）
# 盤2 A1:A10 = apple / #N/A / TRUE / ="" / a~b / " x" / x / 5 / #DIV/0! / (空) ・ B1:B10 = 1..10
# 盤3（retsu-x の盤）H1:H5 = 1..5・I1:I5 = 10..50・L1 = "*"（帯の式は H:H・I:I で聞く）
# 使い方 `powershell -NoProfile -ExecutionPolicy Bypass -File <この道> <出す先>`＝自分で New-Object した Excel だけを使う
$ErrorActionPreference = 'Stop'
$出 = $args[0]
$誤 = @{ -2146826288='#NULL!'; -2146826281='#DIV/0!'; -2146826273='#VALUE!'; -2146826265='#REF!'; -2146826259='#NAME?'; -2146826252='#NUM!'; -2146826246='#N/A'; -2146826238='#GETTING_DATA'; -2146826229='#SPILL!'; -2146826215='#CALC!' }
$式たち = @(
 '=COUNTIF(A1:A10,"*")','=COUNTIF(A1:A10,"<>*")','=COUNTIF(A1:A10,"=*")','=COUNTIF(A1:A10,"?*")','=COUNTIF(A1:A10,"<>?*")',
 '=COUNTIF(A1:A10,"<>x")','=COUNTIF(A1:A10,"=x")','=COUNTIF(A1:A10,"x")','=COUNTIF(A1:A10,"<>")','=COUNTIF(A1:A10,"=")','=COUNTIF(A1:A10,"")',
 '=COUNTIF(A1:A10,"a~b")','=COUNTIF(A1:A10,"a~~b")','=COUNTIF(A1:A10,"<>a~b")',
 '=COUNTIF(A1:A10," x*")','=COUNTIF(A1:A10,"= x*")','=COUNTIF(A1:A10,"<> x*")','=COUNTIF(A1:A10,"x*")','=COUNTIF(A1:A10,"=x*")','=COUNTIF(A1:A10,"<>x*")',
 '=COUNTIF(A1:A10,"5")','=COUNTIF(A1:A10,"=5")','=COUNTIF(A1:A10,"TRUE")','=COUNTIF(A1:A10,"<>TRUE")','=COUNTIF(A1:A10,"#N/A")','=COUNTIF(A1:A10,"<>#N/A")',
 '=SUMIF(A1:A10,"<>*",B1:B10)','=SUMIF(A1:A10,"*",B1:B10)','=COUNTIFS(A1:A10,"<>*",B1:B10,">2")',
 '=COUNTIF(H:H,"<>*")','=SUMIF(H:H,"<>*",I:I)','=COUNTIF(H:H,"a*")','=COUNTIF(H:H,"?")','=COUNTIF(H:H,"~*")',
 '=COUNTIFS(H:H,"",I:I,"*")','=SUMIFS(I:I,H:H,"<>*")','=COUNTIF(H:H,L1)','=COUNTIF(H:H,"<>3")','=COUNTIF(H:H,">2")'
)
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $値 = @('apple','=NA()','TRUE','=""',"'a~b","' x",'x','5','=1/0','')
  for ($i=0; $i -lt 10; $i++) { if ($値[$i] -ne '') { $ws.Cells.Item($i+1,1).Formula = $値[$i] }; $ws.Cells.Item($i+1,2).Value2 = $i+1 }
  for ($i=1; $i -le 5; $i++) { $ws.Cells.Item($i,8).Value2 = $i; $ws.Cells.Item($i,9).Value2 = $i*10 }
  $ws.Cells.Item(1,12).Formula = "'*"
  $行 = @("# 実Excel $($x.Version) build $($x.Build) / 盤2 A1:A10=apple,#N/A,TRUE,="""",a~b,"" x"",x,5(数),#DIV/0!,空 B1:B10=1..10 / 盤3 H1:H5=1..5 I1:I5=10..50 L1=*")
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
