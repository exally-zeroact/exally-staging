# toru-utsushiji-4.ps1 — 字の マスを 数と 読むか（2026-11-07・値の直し2 叩き直し3回目 任意1）
# 盤6 A1:A7 = '5 / ' 5 / '5. / '1E3 / '1,000 / 5(数) / '-0 ・ B1:B7 = 1..7
$ErrorActionPreference = 'Stop'
$出 = $args[0]
$誤 = @{ -2146826281='#DIV/0!'; -2146826273='#VALUE!'; -2146826246='#N/A' }
$式たち = @('=COUNTIF(A1:A7,"5")','=COUNTIF(A1:A7,"=5")','=COUNTIF(A1:A7,"<>5")','=SUMIF(A1:A7,"5",B1:B7)','=COUNTIF(A1:A7,"1000")','=COUNTIF(A1:A7,"1E3")','=COUNTIF(A1:A7,"0")','=COUNTIF(A1:A7,">4")','=COUNTIF(A1:A7,"1,000")')
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $値 = @("'5","' 5","'5.","'1E3","'1,000",'5',"'-0")
  for ($i=0; $i -lt 7; $i++) { $ws.Cells.Item($i+1,1).Formula = $値[$i]; $ws.Cells.Item($i+1,2).Value2 = $i+1 }
  $行 = @("# 実Excel $($x.Version) build $($x.Build) / 盤6 A1:A7='5,' 5','5.','1E3','1,000',5(数),'-0（字） B1:B7=1..7")
  $r = 1
  foreach ($f in $式たち) {
    $c = $ws.Cells.Item($r, 4); $c.Formula = $f; $x.CalculateFull()
    $v = $c.Value2; if ($v -is [int]) { $v = $誤[$v] }
    $w = $ws.Cells.Item($r, 5); $w.Formula = ('=(' + $c.Address($false,$false) + ')=0'); $x.Calculate(); $z = [string]$w.Value2; $w.ClearContents()

    $cv = $c.Value2; $t = if ($cv -is [string]) {'String'} elseif ($cv -is [double]) {'Double'} elseif ($cv -is [bool]) {'Boolean'} elseif ($cv -is [int]) {'Error'} else {'Empty'}

    $行 += "$f`t$v`tz0=$z`tkata=$t"; $c.ClearContents(); $r++
  }
  [IO.File]::WriteAllLines($出, $行, (New-Object Text.UTF8Encoding $false))
  $wb.Close($false)
} finally { $x.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($x) | Out-Null }
'ok'
