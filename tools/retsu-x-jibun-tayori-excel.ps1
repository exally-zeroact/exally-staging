param([string]$Out)
$ErrorActionPreference = 'Stop'
# taiketsu 止め2（11-07）：広く読む関数が自分を頼る式を範囲に含む時、実Excel は輪か・値は何か
# 材料 T1:T5=1..5 / U1:U4=10..40。式を置いて計算→U3=99 にしてもう一度。打つ順 2通り。
$fs = @(
 @{ 名='VLOOKUP'; X='=VLOOKUP(3,T1:U5,2,FALSE)' },
 @{ 名='SUMIF'; X='=SUMIF(T1:T5,3,U1:U5)' },
 @{ 名='IF 両枝'; X='=IF(TRUE,U3,U5)' },
 @{ 名='INDEX MATCH'; X='=INDEX(U1:U5,MATCH(3,T1:T5,0))' },
 @{ 名='SUM 丸ごと'; X='=SUM(U1:U4)+0*SUM(U1:U5)' }
)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$rows = @()
try {
  foreach ($k in $fs) { foreach ($jun in @('X1から','U5から')) {
    $bk = $xl.Workbooks.Add(); $sh = $bk.Worksheets.Item(1)
    for ($i=1; $i -le 5; $i++) { $sh.Cells.Item($i,20).Value2 = $i }
    for ($i=1; $i -le 4; $i++) { $sh.Cells.Item($i,21).Value2 = $i*10 }
    if ($jun -eq 'X1から') { $sh.Range('X1').Formula2 = $k.X; $sh.Range('U5').Formula2 = '=X1+1' } else { $sh.Range('U5').Formula2 = '=X1+1'; $sh.Range('X1').Formula2 = $k.X }
    $xl.Calculate()
    $a1 = 'X1=' + $sh.Range('X1').Text + ' U5=' + $sh.Range('U5').Text
    $cr1 = ''; try { $r = $sh.CircularReference; if ($r) { $cr1 = $r.Address($false,$false) } } catch {}
    $sh.Range('U3').Value2 = 99
    $xl.Calculate()
    $a2 = 'X1=' + $sh.Range('X1').Text + ' U5=' + $sh.Range('U5').Text
    $cr2 = ''; try { $r = $sh.CircularReference; if ($r) { $cr2 = $r.Address($false,$false) } } catch {}
    $rows += [pscustomobject]@{ 名=$k.名; 式=$k.X; 順=$jun; 前=$a1; 輪前=$cr1; U3は99=$a2; 輪後=$cr2 }
    $bk.Close($false)
  } }
  "Excel " + $xl.Version + " build " + $xl.Build + " 本数 " + $rows.Count
} finally { $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }
$rows | ConvertTo-Json | Out-File -Encoding utf8 $Out
