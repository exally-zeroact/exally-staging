# toru-phonetic-lambda.ps1 -- ★PHONETIC と 裸の LAMBDA を 実Excel に 聞く★（㊴）（2026-09-19）
#
#  ★なぜ★ ... Exally1 が 台帳（動く/動かない の 棚）を 直すのに
#             ★PHONETIC が 判じられない★と 出しました
#             ⇒★実Excel に 聞かないと 決められません★
#
#  ★材料★
#    A1 = 1 ／ A2 = 2      ... ★ふりがなは 付けません★
#    B1 = 「山田」          ... ★ふりがなを 付けます★（付いたかを ★読み返して 紙に 書きます★）
#
#  ★門★
#    ①式に 外へ 出る 6個が 0件（exit 5）
#    ②★式★に ASCII の 外の 字が 0件（exit 6）... ★材料は 別★（日本語を 使います）
#    ③走らせる 前の Excel が 0個（exit 3）
#    ④本数 決め打ち（exit 4）
#    ⑤貝殻が powershell.exe（5.1）（exit 8）
#    ⑥司さんの 実物を 開く 字は ★1文字も 在りません★
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-phonetic-lambda-2026-09-19.tsv'

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) {
  Write-Host '★★この 道具は powershell.exe（5.1）で 走らせて ください★★'
  exit 8
}

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... Get-Process ' + $数1 + '個 ／ Win32_Process ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$式たち = @(
  @{ 番 = '1'; 式 = '=PHONETIC(A1)' },
  @{ 番 = '2'; 式 = '=PHONETIC(A1:A2)' },
  @{ 番 = '3'; 式 = '=PHONETIC(B1)' },
  @{ 番 = '4'; 式 = '=LAMBDA(x,x+1)' },
  @{ 番 = '5'; 式 = '=LAMBDA(x,x+1)(2)' },
  @{ 番 = '6'; 式 = '=SUM(A1:A2)' }
)

$外へ出る = 'WEBSERVICE', 'STOCKHISTORY', 'TRANSLATE', 'DETECTLANGUAGE', 'IMAGE', 'RTD'
$見つけた = New-Object System.Collections.Generic.List[string]
foreach ($x in $式たち) {
  foreach ($n in $外へ出る) { if ($x.式.ToUpper().Contains($n)) { $見つけた.Add($x.番 + ' ' + $n) } }
}
Write-Host ('★外へ 出る 6個 ... ' + $見つけた.Count + '件★（0件で ないと 走りません）')
if ($見つけた.Count -ne 0) { exit 5 }

$外の字 = New-Object System.Collections.Generic.List[string]
foreach ($x in $式たち) {
  foreach ($ch in $x.式.ToCharArray()) { if ([int]$ch -gt 127) { $外の字.Add($x.番); break } }
}
Write-Host ('★式に ASCII の 外の 字 ... ' + $外の字.Count + '件★（0件で ないと 走りません）')
if ($外の字.Count -ne 0) { exit 6 }

$式の本数 = 6
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

  # ══ ★材料★ ══
  $sh.Range('A1').Value2 = 1
  $sh.Range('A2').Value2 = 2
  # ★B1 に 日本語を 入れて ふりがなを 付ける★
  $名 = [char]0x5C71 + [char]0x7530          # 山田
  $読 = [char]0x30E4 + [char]0x30DE + [char]0x30C0   # ヤマダ
  $sh.Range('B1').Value2 = $名
  $ふりがな付いた = 'いいえ'
  $ふりがなの字 = ''
  try {
    [void]$sh.Range('B1').Phonetics.Add(1, 2, $読)
    $ふりがなの字 = [string]$sh.Range('B1').Phonetics.Item(1).Text
    if ($ふりがなの字.Length -gt 0) { $ふりがな付いた = 'はい' }
  } catch {
    $ふりがなの字 = '(★付けられません★ ' + $_.Exception.Message + ')'
  }
  Write-Host ('★B1 の ふりがな ... ' + $ふりがな付いた + '（読み返した 字 ... ' + $ふりがなの字 + '）★')

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★PHONETIC と 裸の LAMBDA を 実Excel に 聞いた（㊴）★（2026-09-19）')
  $行.Add('# ★材料★ A1=1 ／ A2=2（ふりがな 無し） ／ B1=山田')
  $行.Add('# ★B1 の ふりがな★ ... ' + $ふりがな付いた + ' ／ 読み返した 字 ... ' + $ふりがなの字)
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★どの 貝殻か★ ... PowerShell ' + $版.ToString())
  $行.Add('# 番' + "`t" + '式' + "`t" + '答え' + "`t" + '出る字' + "`t" + '=(式)=0' + "`t" + '型')

  $押し時計 = [Diagnostics.Stopwatch]::StartNew()
  $r = 1
  foreach ($x in $式たち) {
    $c = $sh.Range('D' + $r)
    $打てた = $true
    try { $c.Formula2 = $x.式 } catch {
      $打てた = $false
      $行.Add($x.番 + "`t" + $x.式 + "`t" + '(★打てません★)' + "`t" + $_.Exception.Message + "`t" + '(★打てません★)' + "`t")
      $r++
    }
    if (-not $打てた) { continue }
    $v = $c.Value2
    $答 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    $型 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
    $字 = [string]$c.Text
    # ★★2つ目の 窓★★ `=(式)=0`（★.Value2 は 0で ない 値にも 0 を 返す 事が 在ります★）
    $w = $sh.Range('Z' + $r)
    $ゼロか = '(★窓2が 打てません★)'
    try { $w.Formula2 = '=(' + $x.式.Substring(1) + ')=0'; $ゼロか = [string]$w.Value2 } catch { }
    $行.Add($x.番 + "`t" + $x.式 + "`t" + $答 + "`t" + $字 + "`t" + $ゼロか + "`t" + $型)
    $r++
  }
  $押し時計.Stop()
  $押し秒 = [math]::Round($押し時計.Elapsed.TotalSeconds, 2)
  $行.Add('# ★★押すのに かかった 秒 ... ' + $押し秒 + '★★')

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
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
