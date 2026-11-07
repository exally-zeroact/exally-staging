# toru-utsushiji-3.ps1 — 条件の 数の 形 と D系の 条件（2026-11-07・値の直し2 叩き直しの 止め1）
# 盤4 A1:A6 = 0 / " " / 16 / x / 5 / '5（字）
# 盤5（D系）F1="名"・F2:F11 = apple,#N/A,TRUE,="",a~b," x",x,5,#DIV/0!,空・G1="数"・G2:G11=1..10
#   条件の マス J1="名"・J2 に 条件の 字を 1つずつ 入れる（="..." の 式で 字に する）
# 使い方 `powershell -NoProfile -ExecutionPolicy Bypass -File <この道> <出す先>`＝自分で New-Object した Excel だけを使う
$ErrorActionPreference = 'Stop'
$出 = $args[0]
$誤 = @{ -2146826288='#NULL!'; -2146826281='#DIV/0!'; -2146826273='#VALUE!'; -2146826265='#REF!'; -2146826259='#NAME?'; -2146826252='#NUM!'; -2146826246='#N/A'; -2146826238='#GETTING_DATA'; -2146826229='#SPILL!'; -2146826215='#CALC!' }
$式たち = @(
 '=COUNTIF(A1:A6," ")','=COUNTIF(A1:A6,"0x10")','=COUNTIF(A1:A6,"0b0")','=COUNTIF(A1:A6,"=0x10")','=COUNTIF(A1:A6,"16")','=COUNTIF(A1:A6,"1.6e1")',
 '=COUNTIF(A1:A6,"5")','=COUNTIF(A1:A6,"=5")','=COUNTIF(A1:A6,"<>5")','=COUNTIF(A1:A6," 5")','=COUNTIF(A1:A6,"+5")','=COUNTIF(A1:A6,"0")','=COUNTIF(A1:A6,"=0")'
)
$D条件 = @('apple','<>apple','#N/A','<>','a*','*','<>*','x','5','TRUE','=','?*','a~~b')
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $値4 = @('0','=" "','16',"'x",'5',"'5")
  for ($i=0; $i -lt 6; $i++) { $ws.Cells.Item($i+1,1).Formula = $値4[$i] }
  $ws.Cells.Item(1,6).Formula = "'名"; $ws.Cells.Item(1,7).Formula = "'数"
  $値5 = @('apple','=NA()','TRUE','=""',"'a~b","' x",'x','5','=1/0','')
  for ($i=0; $i -lt 10; $i++) { if ($値5[$i] -ne '') { $ws.Cells.Item($i+2,6).Formula = $値5[$i] }; $ws.Cells.Item($i+2,7).Value2 = $i+1 }
  $ws.Cells.Item(1,10).Formula = "'名"
  $行 = @("# 実Excel $($x.Version) build $($x.Build) / 盤4 A1:A6=0,"" "",16,x,5(数),'5(字) / 盤5 F1:G11 名=apple,#N/A,TRUE,="""",a~b,"" x"",x,5,#DIV/0!,空 数=1..10 / 条件 J1:J2（J2 は =""..."" の 式）")
  $r = 1
  foreach ($f in $式たち) {
    $c = $ws.Cells.Item($r, 4); $c.Formula = $f; $x.CalculateFull()
    $v = $c.Value2; if ($v -is [int]) { $v = $誤[$v] }
    $w = $ws.Cells.Item($r, 5); $w.Formula = ('=(' + $c.Address($false,$false) + ')=0'); $x.Calculate(); $z = [string]$w.Value2; $w.ClearContents()

    $cv = $c.Value2; $t = if ($cv -is [string]) {'String'} elseif ($cv -is [double]) {'Double'} elseif ($cv -is [bool]) {'Boolean'} elseif ($cv -is [int]) {'Error'} else {'Empty'}

    $行 += "$f`t$v`tz0=$z`tkata=$t"; $c.ClearContents(); $r++
  }
  foreach ($k in $D条件) {
    $ws.Cells.Item(2,10).Formula = '="' + $k + '"'
    foreach ($f in @('=DCOUNTA(F1:G11,1,J1:J2)','=DSUM(F1:G11,2,J1:J2)')) {
      $c = $ws.Cells.Item($r, 4); $c.Formula = $f; $x.CalculateFull()
      $v = $c.Value2; if ($v -is [int]) { $v = $誤[$v] }
      $w = $ws.Cells.Item($r, 5); $w.Formula = ('=(' + $c.Address($false,$false) + ')=0'); $x.Calculate(); $z = [string]$w.Value2; $w.ClearContents()

      $cv = $c.Value2; $t = if ($cv -is [string]) {'String'} elseif ($cv -is [double]) {'Double'} elseif ($cv -is [bool]) {'Boolean'} elseif ($cv -is [int]) {'Error'} else {'Empty'}

      $行 += "$f`tJ2=$k`t$v`tz0=$z`tkata=$t"; $c.ClearContents(); $r++
    }
  }
  [IO.File]::WriteAllLines($出, $行, (New-Object Text.UTF8Encoding $false))
  $wb.Close($false)
} finally { $x.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($x) | Out-Null }
'ok'
