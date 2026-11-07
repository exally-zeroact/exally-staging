param([string]$Out)
$ErrorActionPreference = 'Stop'
# 頼りの件：広く頼る形が実Excel で輪（循環参照）になるか（11-07）
# 材料 A1:A5=1..5 / B1:B5=10..50 / C1:C5=100..500。式を置いて Calculate、B3=7 にして もう一度。
$cases = @(
 @{ 名='C+R 両端 H3 J1'; 式=@(@('H3','=SUM(INDEX(A:A,5):INDEX(1:1,3))'),@('J1','=H3*2')); 見る=@('H3','J1') },
 @{ 名='C+R 鎖 K1'; 式=@(@('H3','=SUM(INDEX(A:A,5):INDEX(1:1,3))'),@('J1','=H3*2'),@('K1','=SUM(INDEX(A:A,5):INDEX(1:1,3))+J1')); 見る=@('H3','J1','K1') },
 @{ 名='四角 INDEX(A1:K5) J1'; 式=@(@('H3','=SUM(INDEX(A1:K5,1,1):C5)'),@('J1','=H3*2')); 見る=@('H3','J1') },
 @{ 名='A10 A11'; 式=@(@('A10','=SUM(INDEX(A:A,5):INDEX(1:1,3))'),@('A11','=A10+1')); 見る=@('A10','A11') },
 @{ 名='帯C 端 B10'; 式=@(@('H3','=SUM(INDEX(A:A,5):C3)'),@('B10','=H3*2')); 見る=@('H3','B10') },
 @{ 名='四角 A1:A9 B8'; 式=@(@('H3','=SUM(INDEX(A1:A9,2):C3)'),@('B8','=H3*2')); 見る=@('H3','B8') },
 @{ 名='INDEX 1マス 列 A7'; 式=@(@('H3','=INDEX(A:A,2)'),@('A7','=H3*2')); 見る=@('H3','A7') },
 @{ 名='対照 輪なし'; 式=@(@('H3','=SUM(INDEX(A:A,5):INDEX(1:1,3))'),@('J9','=H3*2')); 見る=@('H3','J9') }
)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$rows = @()
try {
  foreach ($k in $cases) {
    $bk = $xl.Workbooks.Add(); $sh = $bk.Worksheets.Item(1)
    for ($i=1; $i -le 5; $i++) { $sh.Cells.Item($i,1).Value2 = $i; $sh.Cells.Item($i,2).Value2 = $i*10; $sh.Cells.Item($i,3).Value2 = $i*100 }
    foreach ($p in $k.式) { $sh.Range($p[0]).Formula2 = $p[1] }
    $xl.Calculate()
    $a1 = @(); foreach ($c in $k.見る) { $a1 += ($c + '=' + $sh.Range($c).Text) }
    $cr1 = ''; try { $r = $sh.CircularReference; if ($r) { $cr1 = $r.Address($false,$false) } } catch {}
    $sh.Range('B3').Value2 = 7
    $xl.Calculate()
    $a2 = @(); foreach ($c in $k.見る) { $a2 += ($c + '=' + $sh.Range($c).Text) }
    $cr2 = ''; try { $r = $sh.CircularReference; if ($r) { $cr2 = $r.Address($false,$false) } } catch {}
    $rows += [pscustomobject]@{ 名=$k.名; 前=($a1 -join ' '); 輪前=$cr1; B3は7=($a2 -join ' '); 輪後=$cr2 }
    $sh = $null; $bk.Close($false); $bk = $null
  }
  $rows | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $Out -Encoding UTF8
  "Excel " + $xl.Version + " build " + $xl.Build + " 本数 " + $rows.Count
} finally {
  $xl.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl); $xl = $null
}
