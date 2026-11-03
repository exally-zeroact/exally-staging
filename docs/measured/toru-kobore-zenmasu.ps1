# toru-kobore-zenmasu.ps1 -- ★溢れた 先の 全マスを 実Excel に 聞く★（2026-09-19）
#
#  ★なぜ 要るか★
#    今の 分母は ★溢れの 左上しか 見て いません★
#    （`osu-kami-webkit.mjs` 24行目の 断り）
#    ⇒★★2,105 / 2,116 の「合った」は ★左上が 合った★という 意味です★★
#
#  ★材料★ A1:A3 = 3 / 1 / 2 　 B1:B3 = 10 / 20 / 30
#  ★読む 範囲★ D1:H10（50マス）を 全部
#  ★２つ目の 窓★ J1 に `=(式)=0` を 打ち J1:N10 を 読む
#  ★対照★ `=SUM(A1:A3)` ⇒ ★6★（材料が 入ったか）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-kobore-zenmasu-2026-09-19.tsv'

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) {
  Write-Host '★★この 道具は powershell.exe（5.1）で 走らせて ください★★'
  exit 8
}

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$式たち = @(
  @{ 番 = '1'; 式 = '=SORT(A1:A3)' },
  @{ 番 = '2'; 式 = '=UNIQUE(A1:A3)' },
  @{ 番 = '3'; 式 = '=FILTER(A1:A3,A1:A3>1)' },
  @{ 番 = '4'; 式 = '=SEQUENCE(3)' },
  @{ 番 = '5'; 式 = '=SEQUENCE(2,3)' },
  @{ 番 = '6'; 式 = '=TRANSPOSE(A1:A3)' },
  @{ 番 = '7'; 式 = '=SORTBY(A1:A3,B1:B3)' },
  @{ 番 = '8'; 式 = '=TEXTSPLIT("a,b,c",",")' },
  @{ 番 = '9'; 式 = '=MAP(A1:A3,LAMBDA(x,x*2))' },
  @{ 番 = '10'; 式 = '=BYROW(A1:B3,LAMBDA(r,SUM(r)))' },
  @{ 番 = '11'; 式 = '=SCAN(0,A1:A3,LAMBDA(a,c,a+c))' },
  @{ 番 = '12'; 式 = '=MAKEARRAY(2,2,LAMBDA(r,c,r*c))' },
  @{ 番 = '13'; 式 = '=BYCOL(A1:B3,LAMBDA(c,SUM(c)))' },
  @{ 番 = '14'; 式 = '=SUM(A1:A3)' }
)

$外へ出る = 'WEBSERVICE', 'STOCKHISTORY', 'TRANSLATE', 'DETECTLANGUAGE', 'IMAGE', 'RTD'
$見つけた = New-Object System.Collections.Generic.List[string]
foreach ($x in $式たち) {
  foreach ($n in $外へ出る) { if ($x.式.ToUpper().Contains($n)) { $見つけた.Add($x.番) } }
}
Write-Host ('★外へ 出る 6個 ... ' + $見つけた.Count + '件★')
if ($見つけた.Count -ne 0) { exit 5 }

$式の本数 = 14
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

  # ★材料★
  $sh.Range('A1').Value2 = 3
  $sh.Range('A2').Value2 = 1
  $sh.Range('A3').Value2 = 2
  $sh.Range('B1').Value2 = 10
  $sh.Range('B2').Value2 = 20
  $sh.Range('B3').Value2 = 30

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★溢れた 先の 全マスを 実Excel に 聞いた★（2026-09-19）')
  $行.Add('# ★材料★ A1:A3 = 3/1/2 ／ B1:B3 = 10/20/30')
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★どの 貝殻か★ ... PowerShell ' + $版.ToString())
  $行.Add('# 番' + "`t" + '式' + "`t" + '行数' + "`t" + '列数' + "`t" + '並び(行は / 列は |)' + "`t" + '型' + "`t" + '=(式)=0 の 左上')

  $押し時計 = [Diagnostics.Stopwatch]::StartNew()
  foreach ($x in $式たち) {
    $sh.Range('D1:N10').ClearContents() | Out-Null
    $打てた = $true
    # ★★`.Formula` では 溢れません★★（★1回目で 14本とも 1x1 に なりました★）
    #   ＝`.Formula` は ★昔の 形★で 書くので ★暗黙の 交差★が 効きます
    #   ⇒★`.Formula2` が ★溢れる 形★★
    try { $sh.Range('D1').Formula2 = $x.式 } catch { $打てた = $false }
    if (-not $打てた) {
      $行.Add($x.番 + "`t" + $x.式 + "`t0`t0`t" + '(★打てません★)' + "`t`t")
      continue
    }
    # ★★2つ目の 窓★★ `=(式)=0`（★溢れるので これも 並びに なります★）
    $ゼロか = "(★窓2が 打てません★)"
    try { $w = $sh.Range('J1'); $w.Formula2 = '=(' + $x.式.Substring(1) + ')=0'; $ゼロか = [string]$w.Value2 } catch { }
    # ★D1:H10 を 全部 読む★
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
    $行.Add($x.番 + "`t" + $x.式 + "`t" + $最終行 + "`t" + $最終列 + "`t" + $文 + "`t" + $型 + "`t" + $ゼロか)
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
