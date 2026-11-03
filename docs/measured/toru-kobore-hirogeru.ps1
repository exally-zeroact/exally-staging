# toru-kobore-hirogeru.ps1 -- ★溢れの 物差しを 広げる★（2026-09-19）
#
#  ★なぜ★ ... 今の 紙は ★13本／材料 1組だけ★
#         ⇒★直した 後に 「戻って いないか」が 見えません★
#
#  ★材料★ A1:A3 = 3 / 1 / 2 　 B1:B3 = 10 / 20 / 30
#  ★読む★ D1:H10（50マス）
#  ★邪魔★ ‸18 は E1 に 先に 字を 置きます（★#SPILL! に なるか★）
#  ★２つ目の 窓★ J1 に `=(式)=0`
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-kobore-hirogeru-2026-09-19.tsv'

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$式たち = @(
  @{ 番 = '1'; 式 = '=SEQUENCE(1,4)'; 邪魔 = $null; 訳 = 'yoko 1x4' },
  @{ 番 = '2'; 式 = '=SEQUENCE(4,1)'; 邪魔 = $null; 訳 = 'tate 4x1' },
  @{ 番 = '3'; 式 = '=SEQUENCE(3,3)'; 邪魔 = $null; 訳 = 'shikaku 3x3' },
  @{ 番 = '4'; 式 = '=TRANSPOSE(SEQUENCE(2,3))'; 邪魔 = $null; 訳 = 'tenchi 3x2' },
  @{ 番 = '5'; 式 = '=MAP(A1:A3,LAMBDA(x,x*10))'; 邪魔 = $null; 訳 = 'MAP tate' },
  @{ 番 = '6'; 式 = '=MAP(A1:B3,LAMBDA(x,x+1))'; 邪魔 = $null; 訳 = 'MAP shikaku 3x2' },
  @{ 番 = '7'; 式 = '=BYROW(A1:B3,LAMBDA(r,MAX(r)))'; 邪魔 = $null; 訳 = 'BYROW' },
  @{ 番 = '8'; 式 = '=BYCOL(A1:B3,LAMBDA(c,MIN(c)))'; 邪魔 = $null; 訳 = 'BYCOL yoko' },
  @{ 番 = '9'; 式 = '=SCAN(1,A1:A3,LAMBDA(a,c,a*c))'; 邪魔 = $null; 訳 = 'SCAN kakezan' },
  @{ 番 = '10'; 式 = '=MAKEARRAY(3,2,LAMBDA(r,c,r+c))'; 邪魔 = $null; 訳 = 'MAKEARRAY 3x2' },
  @{ 番 = '11'; 式 = '=MAKEARRAY(1,3,LAMBDA(r,c,c*c))'; 邪魔 = $null; 訳 = 'MAKEARRAY yoko' },
  @{ 番 = '12'; 式 = '=REDUCE(0,A1:A3,LAMBDA(a,c,a+c))'; 邪魔 = $null; 訳 = 'REDUCE = koboremasen' },
  @{ 番 = '13'; 式 = '=SORT(B1:B3,1,-1)'; 邪魔 = $null; 訳 = 'SORT gyaku' },
  @{ 番 = '14'; 式 = '=UNIQUE(A1:A3)'; 邪魔 = $null; 訳 = 'UNIQUE' },
  @{ 番 = '15'; 式 = '=FILTER(A1:A3,A1:A3>5)'; 邪魔 = $null; 訳 = 'FILTER = atehamaru mono nashi' },
  @{ 番 = '16'; 式 = '=TEXTSPLIT("a,b;c,d",",",";")'; 邪魔 = $null; 訳 = 'TEXTSPLIT 2jigen' },
  @{ 番 = '17'; 式 = '=MAP(A1:A3,LAMBDA(x,IF(x>2,"dai","shou")))'; 邪魔 = $null; 訳 = 'MAP ji wo kaesu' },
  @{ 番 = '18'; 式 = '=SEQUENCE(2,2)'; 邪魔 = 'E1'; 訳 = '★溢れ先が 塞がれて いる（#SPILL! に なるか）★' },
  @{ 番 = '19'; 式 = '=SEQUENCE(2,2)'; 邪魔 = $null; 訳 = '★上と 同じ 式／邪魔が 無い 時★（対照）' },
  @{ 番 = '20'; 式 = '=SUM(A1:A3)'; 邪魔 = $null; 訳 = '★対照★ 6 に なる はず' }
)

$外へ出る = 'WEBSERVICE', 'STOCKHISTORY', 'TRANSLATE', 'DETECTLANGUAGE', 'IMAGE', 'RTD'
$見つけた = New-Object System.Collections.Generic.List[string]
foreach ($x in $式たち) {
  foreach ($n in $外へ出る) { if ($x.式.ToUpper().Contains($n)) { $見つけた.Add($x.番) } }
}
Write-Host ('★外へ 出る 6個 ... ' + $見つけた.Count + '件★')
if ($見つけた.Count -ne 0) { exit 5 }

$式の本数 = 20
Write-Host ('★聞く 式 ... ' + $式たち.Count + '本★（決め打ち ' + $式の本数 + '本）')
if ($式たち.Count -ne $式の本数) { exit 4 }

$xl = New-Object -ComObject Excel.Application
$bk = $null
$sh = $null
$c = $null
$w = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)
  $sh.Range('A1').Value2 = 3
  $sh.Range('A2').Value2 = 1
  $sh.Range('A3').Value2 = 2
  $sh.Range('B1').Value2 = 10
  $sh.Range('B2').Value2 = 20
  $sh.Range('B3').Value2 = 30

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★溢れの 物差しを 広げた★（2026-09-19）')
  $行.Add('# ★材料★ A1:A3 = 3/1/2 ／ B1:B3 = 10/20/30')
  $行.Add('# ★邪魔★ ⑱は E1 に 先に 字を 置く（#SPILL! に なるか）')
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★どの 貝殻か★ ... PowerShell ' + $版.ToString())
  $行.Add('# 番' + "`t" + '式' + "`t" + '行数' + "`t" + '列数' + "`t" + '並び(行は / 列は |)' + "`t" + '型' + "`t" + '=(式)=0 の 左上' + "`t" + '邪魔のマス' + "`t" + '訳')

  $押し時計 = [Diagnostics.Stopwatch]::StartNew()
  foreach ($x in $式たち) {
    $sh.Range('D1:N10').ClearContents() | Out-Null
    if ($null -ne $x.邪魔) { $sh.Range($x.邪魔).Value2 = "jama" }
    $打てた = $true
    # ★`.Formula` では 溢れません★（暗黙の 交差）⇒ ★`.Formula2`★
    try { $sh.Range('D1').Formula2 = $x.式 } catch { $打てた = $false }
    if (-not $打てた) {
      $行.Add($x.番 + "`t" + $x.式 + "`t0`t0`t" + '(★打てません★)' + "`t`t`t" + [string]$x.邪魔 + "`t" + $x.訳)
      continue
    }
    $ゼロか = "(★窓2が 打てません★)"
    try { $w = $sh.Range('J1'); $w.Formula2 = '=(' + $x.式.Substring(1) + ')=0'; $ゼロか = [string]$w.Value2 } catch { }
    $行たち = New-Object System.Collections.Generic.List[string]
    $最終行 = 0
    $最終列 = 0
    $型 = ""
    for ($r = 1; $r -le 10; $r++) {
      $一行 = New-Object System.Collections.Generic.List[string]
      for ($cc = 4; $cc -le 8; $cc++) {
        $c = $sh.Cells.Item($r, $cc)
        $v = $c.Value2
        if ($null -eq $v) { $一行.Add(""); continue }
        if ($r -gt $最終行) { $最終行 = $r }
        if (($cc - 3) -gt $最終列) { $最終列 = $cc - 3 }
        if ($型 -eq "") {
          if ($v -is [string]) { $型 = "String" }
          elseif ($v -is [double]) { $型 = "Double" }
          elseif ($v -is [bool]) { $型 = "Boolean" }
          else { $型 = "Other" }
        }
        if ($v -is [double]) { $一行.Add($v.ToString("R", [Globalization.CultureInfo]::InvariantCulture)) }
        else { $一行.Add([string]$v) }
      }
      $行たち.Add(($一行 -join '|'))
    }
    $並び = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $最終行; $i++) { $並び.Add($行たち[$i]) }
    $文 = ($並び -join ' / ')
    $行.Add($x.番 + "`t" + $x.式 + "`t" + $最終行 + "`t" + $最終列 + "`t" + $文 + "`t" + $型 + "`t" + $ゼロか + "`t" + [string]$x.邪魔 + "`t" + $x.訳)
  }
  $押し時計.Stop()
  $押し秒 = [math]::Round($押し時計.Elapsed.TotalSeconds, 2)
  $行.Add('# ★★押すのに かかった 秒 ... ' + $押し秒 + '★★')

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ('★★押すのに かかった 秒 ... ' + $押し秒 + '秒★★（' + $式たち.Count + '本）')
  Write-Host ('★書いた ... ' + $出 + '★')
  $bk.Close($false)
} finally {
  $c = $null
  $w = $null
  $sh = $null
  $bk = $null
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $xl = $null
  $t = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 120)) {
    Start-Sleep -Milliseconds 250
  }
  $t.Stop()
  $残り = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
  if ($残り -eq 0) {
    Write-Host ('★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒で ★消えました★★')
  } else {
    Write-Host ('★★Excel は 消えませんでした★★ ／ 残り ' + $残り + '個')
  }
}
