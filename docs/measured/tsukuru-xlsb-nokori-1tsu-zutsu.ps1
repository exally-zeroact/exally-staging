# tsukuru-xlsb-nokori-1tsu-zutsu.ps1
#   -- ★残って いた 未測定を 1回で 測る 材料★（89）（2026-09-21）
#
#  ★★なぜ★★
#    87 で 線を 測った 時、★A5 の 下線が A6 の 上へ 移って★
#    ★太さの バイトを A5 から 読めません でした★。
#    ⇒★今度は 1行 おきに 置きます★（隣り合わせない＝移り先が 空の マス）
#    ⇒ついでに 残りの 未測定も 同じ 包みで 測ります。
#
#  ★★置き場（★先に 隙間を 計算しました★）★★
#    A1  下線 細い（Weight 既定）
#    A3  下線 ★中くらい★（Weight -4138）
#    A5  下線 ★太い★（Weight 4）
#    A7  下線 ★極細★（Weight 1）
#    A9  ★斜めの 線★（Borders.Item(5)＝左上から 右下）
#    A11 ★真ん中 揃え★（HorizontalAlignment -4108）
#    A13 ★鍵を 外す★（Locked = False）
#    A15 ★文字★（'abc'）... ★数の マスと 記録の 番号が 違う はず★
#    ⇒★1行 おき★＝★隣へ 移っても 空の 行に 移る★ので A列の 記録が 空に ならない
#
#  ★門★
#    ①貝殻が 5.1（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③★8つ とも 読み返して 効いて いるか★（exit 5）／④名が 決め打ち（exit 7）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$出す先 = Join-Path $env:TEMP 'exally-nokori-1tsu.xlsb'
if ((Split-Path $出す先 -Leaf) -ne 'exally-nokori-1tsu.xlsb') { exit 7 }

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)

  foreach ($r in 1, 3, 5, 7, 9, 11, 13) { $sh.Cells.Item($r, 1).Value2 = $r }
  $sh.Range('A15').Value2 = 'abc'

  $sh.Range('A1').Borders.Item(9).LineStyle = 1
  $sh.Range('A3').Borders.Item(9).LineStyle = 1
  $sh.Range('A3').Borders.Item(9).Weight = -4138          # ★中くらい★
  $sh.Range('A5').Borders.Item(9).LineStyle = 1
  $sh.Range('A5').Borders.Item(9).Weight = 4              # ★太い★
  $sh.Range('A7').Borders.Item(9).LineStyle = 1
  $sh.Range('A7').Borders.Item(9).Weight = 1              # ★極細★
  $sh.Range('A9').Borders.Item(5).LineStyle = 1           # ★斜め★
  $sh.Range('A11').HorizontalAlignment = -4108            # ★真ん中★
  $sh.Range('A13').Locked = $false                        # ★鍵を 外す★

  $確 = @(
    ('A1 太さ=' + [string]$sh.Range('A1').Borders.Item(9).Weight),
    ('A3 太さ=' + [string]$sh.Range('A3').Borders.Item(9).Weight),
    ('A5 太さ=' + [string]$sh.Range('A5').Borders.Item(9).Weight),
    ('A7 太さ=' + [string]$sh.Range('A7').Borders.Item(9).Weight),
    ('A9 斜め=' + [string]$sh.Range('A9').Borders.Item(5).LineStyle),
    ('A11 揃え=' + [string]$sh.Range('A11').HorizontalAlignment),
    ('A13 鍵=' + [string]$sh.Range('A13').Locked),
    ('A15 字=' + [string]$sh.Range('A15').Value2)
  )
  Write-Host ''
  foreach ($c in $確) { Write-Host ('  ' + $c) }
  if ($確.Count -ne 8) { Write-Host '★★確かめが 8本 在りません★★'; exit 5 }
  foreach ($c in $確) { if ($c -match '=$|=0$|鍵=True$') { Write-Host ('★★効いて いません★★ ' + $c); exit 5 } }

  if (Test-Path $出す先) { Remove-Item $出す先 -Force }
  $bk.SaveAs($出す先, 50)
  $bk.Close($false); $sh = $null; $bk = $null
} finally {
  $sh = $null
  if ($null -ne $bk) { $bk.Close($false) }
  $bk = $null
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $xl = $null
  $t = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 180)) {
    Start-Sleep -Milliseconds 250
  }
  $t.Stop()
  $残り = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
  if ($残り -eq 0) { Write-Host ('★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒で ★消えました★★') }
  else { Write-Host ('★★Excel は 消えませんでした★★ ／ 残り ' + $残り + '個') }
}

if (-not (Test-Path $出す先)) { exit 5 }
$x = Get-Item $出す先
Write-Host ''
Write-Host ('★★作りました★★ ... ' + $出す先)
Write-Host ('  ★大きさ★ ' + $x.Length + ' バイト ／ ★sha256★ ' + (Get-FileHash $出す先 -Algorithm SHA256).Hash.ToLower())
