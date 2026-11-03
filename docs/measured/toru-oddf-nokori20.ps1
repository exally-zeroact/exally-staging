# toru-oddf-nokori20.ps1 -- ★ 2009-03-25 から 04-20 まで 1日ずつ 27本★（㈙の次）（2026-09-19）
#
#  ★ 04-15 は 同じ 紙の 中に 2本 入れて います★
#  ★順は 散らして います★（種 420）
#  ★対照★ ... golden-oddf-nokori-2026-09-18.tsv = ★96.1767213744142★
#
#
#  ★問いは 紙から 機械で 拾いました★ ... kansuu46/oddf-nokori-1pon-kiku-koto.md
#    ＝★手で 書き写して いません★
#    ＝`basis` と 書いて ある 行は 0〜4 に 広げて います
#
#  ★材料は 要りません★ ... 全部 その場の 日付（★司さんの 実物は 使いません★）
#
#  ★門★
#    ①式に 外へ 出る 6個が 0件（exit 5）
#    ②式に ASCII の 外の 字が 0件（exit 6）
#    ③走らせる 前の Excel が 0個（2つの 道具・exit 3）
#    ④本数 決め打ち（exit 4）
#    ⑤貝殻が powershell.exe（5.1）（exit 8）
#    ⑥司さんの 実物を 開く 字が 0件（★この 道具には 1文字も 在りません★）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-oddf-nokori20-2026-09-19.tsv'

# ══ ★⑤貝殻の 版★ ══
$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) {
  Write-Host '★★この 道具は `powershell.exe`（5.1）で 走らせて ください★★'
  Write-Host '  ★訳★ ... 5.1 と 7.x は ★同じ 数を 違う 字で 書きます★'
  exit 8
}

# ══ ★③Excel を 2つの 道具で★ ══
$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... Get-Process ' + $数1 + '個 ／ Win32_Process ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

# ══ ★紙から 拾った 問い★ ══
$式たち = @(
  @{ 番 = '1'; 式 = '=ODDFPRICE(DATE(2009,4,14), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '2'; 式 = '=ODDFPRICE(DATE(2009,4,11), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '3'; 式 = '=ODDFPRICE(DATE(2009,4,1), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '4'; 式 = '=ODDFPRICE(DATE(2009,3,26), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '5'; 式 = '=ODDFPRICE(DATE(2009,4,7), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '6'; 式 = '=ODDFPRICE(DATE(2009,4,20), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '7'; 式 = '=ODDFPRICE(DATE(2009,4,12), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '8'; 式 = '=ODDFPRICE(DATE(2009,3,29), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '9'; 式 = '=ODDFPRICE(DATE(2009,3,31), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '10'; 式 = '=ODDFPRICE(DATE(2009,3,30), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '11'; 式 = '=ODDFPRICE(DATE(2009,4,18), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '12'; 式 = '=ODDFPRICE(DATE(2009,3,27), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '13'; 式 = '=ODDFPRICE(DATE(2009,4,3), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '14'; 式 = '=ODDFPRICE(DATE(2009,4,8), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '15'; 式 = '=ODDFPRICE(DATE(2009,4,17), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '16'; 式 = '=ODDFPRICE(DATE(2009,4,13), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '17'; 式 = '=ODDFPRICE(DATE(2009,4,9), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '18'; 式 = '=ODDFPRICE(DATE(2009,4,4), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '19'; 式 = '=ODDFPRICE(DATE(2009,4,10), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '20'; 式 = '=ODDFPRICE(DATE(2009,3,10), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0.045, 0.05, 100, 4, 1)' },
  @{ 番 = '21'; 式 = '=ODDFPRICE(DATE(2009,4,16), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '22'; 式 = '=ODDFPRICE(DATE(2009,4,15), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '23'; 式 = '=ODDFPRICE(DATE(2009,3,28), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '24'; 式 = '=ODDFPRICE(DATE(2009,4,6), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '25'; 式 = '=ODDFPRICE(DATE(2009,4,2), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '26'; 式 = '=ODDFPRICE(DATE(2009,4,5), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '27'; 式 = '=ODDFPRICE(DATE(2009,4,19), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '28'; 式 = '=ODDFPRICE(DATE(2009,4,15), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' },
  @{ 番 = '29'; 式 = '=ODDFPRICE(DATE(2009,3,25), DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)' }
)

# ══ ★①外へ 出る 6個★ ══
$外へ出る = 'WEBSERVICE', 'STOCKHISTORY', 'TRANSLATE', 'DETECTLANGUAGE', 'IMAGE', 'RTD'
$見つけた = New-Object System.Collections.Generic.List[string]
foreach ($x in $式たち) {
  foreach ($n in $外へ出る) { if ($x.式.ToUpper().Contains($n)) { $見つけた.Add($x.番 + ' ' + $n) } }
}
Write-Host ('★外へ 出る 6個 ... ' + $見つけた.Count + '件★（0件で ないと 走りません）')
if ($見つけた.Count -ne 0) { foreach ($s in $見つけた) { Write-Host ('    ' + $s) }; exit 5 }

# ══ ★②式に ASCII の 外の 字★ ══
$外の字 = New-Object System.Collections.Generic.List[string]
foreach ($x in $式たち) {
  foreach ($ch in $x.式.ToCharArray()) { if ([int]$ch -gt 127) { $外の字.Add($x.番 + ' ' + $ch); break } }
}
Write-Host ('★式に ASCII の 外の 字 ... ' + $外の字.Count + '件★（0件で ないと 走りません）')
if ($外の字.Count -ne 0) { foreach ($s in $外の字) { Write-Host ('    ' + $s) }; exit 6 }

# ══ ★④本数の 門★ ══
$式の本数 = 29
Write-Host ('★聞く 式 ... ' + $式たち.Count + '本★（決め打ち ' + $式の本数 + '本）')
if ($式たち.Count -ne $式の本数) {
  Write-Host ('★★' + $式の本数 + '本の はずが ' + $式たち.Count + '本です★★')
  exit 4
}

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

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★ODDF の 残り 1本 と ODDL の 6本を 実Excel に 聞いた★（2026-09-18）')
  $行.Add('#')
  $行.Add('# ★問いは 紙から 機械で 拾いました★ ... kansuu46/oddf-nokori-1pon-kiku-koto.md')
  $行.Add('# ★材料は 要りません★（全部 その場の 日付）')
  $行.Add('#')
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★どの 貝殻か★ ... PowerShell ' + $版.ToString())
  $コード = [string][System.Text.Encoding]::Default.CodePage
  $土地 = [System.Globalization.CultureInfo]::CurrentCulture.Name
  $行.Add('# ★どの 文字コードか★ ... ANSI ' + $コード + ' ／ ' + $土地)
  $行.Add('#')
  $行.Add('# 番' + "`t" + '式' + "`t" + '答え' + "`t" + '出る字' + "`t" + '=(式)=0' + "`t" + '型')

  $押し時計 = [Diagnostics.Stopwatch]::StartNew()
  $r = 1
  foreach ($x in $式たち) {
    $c = $sh.Range('A' + $r)
    $打てた = $true
    try { $c.Formula = $x.式 } catch {
      $打てた = $false
      $行.Add($x.番 + "`t" + $x.式 + "`t" + '(★打てません★)' + "`t" + ('★Excel が 式を 受け付けません★ ' + $_.Exception.Message) + "`t" + '(★打てません★)')
      $r++
    }
    if (-not $打てた) { continue }
    $v = $c.Value2
    $答 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    $型 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
    $字 = [string]$c.Text
    # ★★2つ目の 窓★★（`monosashi-mado` の 門）
    #   ★訳★ ... `.Value2` は ★0 で ない 値にも 0 を 返す 事が 在ります★
    #            ⇒★`=(式)=0` を 別の マスで 取り、真偽と 型を 一緒に 見る★
    $w = $sh.Range('Z' + $r)
    $ゼロか = '(★窓2が 打てません★)'
    try { $w.Formula = '=(' + $x.式.Substring(1) + ')=0'; $ゼロか = [string]$w.Value2 } catch { }
    $行.Add($x.番 + "`t" + $x.式 + "`t" + $答 + "`t" + $字 + "`t" + $ゼロか + "`t" + $型)
    $r++
  }
  $押し時計.Stop()
  $押し秒 = [math]::Round($押し時計.Elapsed.TotalSeconds, 2)
  $行.Add('# ★★押すのに かかった 秒 ... ' + $押し秒 + '★★')

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★★押すのに かかった 秒 ... ' + $押し秒 + '秒★★（' + $式たち.Count + '本）')
  Write-Host ('★書いた ... ' + $出 + '（' + $式たち.Count + '行）★')
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
    Write-Host ('★★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒 待っても ★消えませんでした★★ ／ 残り ' + $残り + '個')
  }
}
