# toru-jitsu-excel-no-e.ps1
#   -- ★実Excel 自身に 絵を 描かせる★（86）（2026-09-21）
#
#  ★★なぜ★★
#    Exally1「★今まで 一度も 『実Excel の 絵』と 並べて いません★」
#    ⇒飾りが「ファイルの 中に 在る」のは 測った。
#      「Exally の 画面に 出る」のも 測った。
#      ★でも 「実Excel では どう 見えるか」は 1枚も 無い★
#    ⇒★合わせる 相手が 無い まま 「同じように 表示」は 言えません★
#
#  ★★なぜ 画面の 写真では ないか★★
#    ・窓の 縁・リボン・行列の 見出し・選んだ 印が ★一緒に 写ります★
#      ⇒★中身の 違いか 窓の 違いか 分からなく なります★
#    ・窓の 大きさ／画面の 倍率で ★同じ 物が 違う 絵に なります★
#    ⇒★Excel 自身に マスを 絵に させます★（`CopyPicture` ⇒ 図に 貼る ⇒ 書き出す）
#      ＝★出る 絵は マスの 中身だけ★
#
#  ★★測って いない 事（先に 書きます）★★
#    ・★これは 「印刷の 絵」でも 「画面の 絵」でも ありません★
#      `CopyPicture` の 1つ目に `1`（画面の 見た目）を 渡して います。
#      ⇒★選んだ 印・カーソルは 入りません★
#    ・★行列の 見出し（A B C / 1 2 3）は 入りません★
#      ⇒Exally の 絵と 並べる 時は ★見出しを 外して 比べて ください★
#
#  ★★作る 物★★ `%TEMP%\exally-jitsu-excel-no-e.png`（★1本の 名★）
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③開く 物が 在るか（exit 4）／④★書いた 名が 決め打ちか★（exit 7）
#    ⑤★出来た 絵の 頭が PNG の 印か★（exit 9）／⑥★大きさが 0では ないか★（exit 5）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具> [-元 <xlsx>] [-所 A1:F12]

param(
  [string]$元 = (Join-Path $env:TEMP 'exally-tameshi-hiraku3.xlsx'),
  [string]$所 = 'A1:F12'
)

$出す先 = Join-Path $env:TEMP 'exally-jitsu-excel-no-e.png'
$許す名 = 'exally-jitsu-excel-no-e.png'
if ((Split-Path $出す先 -Leaf) -ne $許す名) { Write-Host '★★書いて よい 名は 1本だけです★★'; exit 7 }

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

if (-not (Test-Path $元)) { Write-Host ('★★開く 物が 在りません★★ ... ' + $元); exit 4 }
Write-Host ('★開く 物★ ... ' + $元)
Write-Host ('★所★ ... ' + $所)

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null; $ma = $null; $枠 = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  # ★読むだけ★（★お客さんの ファイルには 1文字も 書きません★）
  $bk = $xl.Workbooks.Open($元, $false, $true)
  $sh = $bk.Sheets.Item(1)
  $ma = $sh.Range($所)

  Write-Host ('  ★マスの 大きさ★ 幅 ' + [math]::Round($ma.Width, 1) + ' ／ 高さ ' + [math]::Round($ma.Height, 1) + '（ポイント）')
  Write-Host ('  ★繋げた マス★ A5 ... ' + [string]$sh.Range('A5').MergeCells)

  # ★★Excel 自身に 絵に させます★★
  #   `CopyPicture(1, 2)` ＝ 1:画面の 見た目 ／ 2:ビットマップ
  #
  # ★★2026-09-21 ── ★1回目は 真っ白が 出ました★★
  #   ＝`Paste` は ★投げずに 何も しない★ 事が 在ります（写し場が まだ 出来て いない）
  #   ⇒★貼れたかは 「投げなかったか」では なく ★図が 1つ 載ったか★ で 見ます★
  #   ⇒板を 前に 出して から 写し、★載るまで 何回か 試します★
  $sh.Activate() | Out-Null
  # ★★2026-09-21 ── ★枠を マスの 上に 置いたら 枠自身が 写りました★★
  #   ＝`0, 0` に 置くと 写す 所を ★白い 枠が 覆います★
  #   ⇒★投げず・真っ白でも なく ★線だけの 絵★が 出ます（一番 見つけにくい）★
  #   ⇒★写す 所の ★外★ に 置きます★
  $よけ = $ma.Left + $ma.Width + 200
  $枠 = $sh.ChartObjects().Add($よけ, 0, $ma.Width, $ma.Height)
  $枠.Chart.ChartArea.Border.LineStyle = 0     # ★枠の 線を 消す★（中身だけ に する）
  $のった = $false
  for ($k = 1; $k -le 20; $k++) {
    $ma.CopyPicture(1, 2) | Out-Null
    Start-Sleep -Milliseconds 200
    # ★★2026-09-21 ── ★枠を 選んで から でないと 貼れません★★
    #   （どけた 途端に 0個に なりました＝前は 覆って いた ので たまたま 通って いた）
    try { $枠.Activate() | Out-Null } catch { }
    try { $枠.Chart.Paste() } catch { }
    if ([int]$枠.Chart.Shapes.Count -gt 0) { $のった = $true; break }
  }
  Write-Host ('  ★貼れたか★ ... ' + $のった + '（載った 図 ' + [string]$枠.Chart.Shapes.Count + '個）')
  if (-not $のった) { Write-Host '★★貼れませんでした★★'; exit 6 }
  if (Test-Path $出す先) { Remove-Item $出す先 -Force }
  $枠.Chart.Export($出す先, 'PNG') | Out-Null
  $枠.Delete() | Out-Null
  $枠 = $null

  $bk.Close($false)
  $bk = $null
} finally {
  # ★★掴んだ物 全部 $null★★
  if ($null -ne $枠) { $枠.Delete() | Out-Null }
  $枠 = $null
  $ma = $null
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

if (-not (Test-Path $出す先)) { Write-Host '★★絵が 出来て いません★★'; exit 5 }
$x = Get-Item $出す先
if ($x.Length -le 0) { Write-Host '★★絵の 大きさが 0です★★'; exit 5 }
$fs = [IO.File]::OpenRead($出す先)
$b = New-Object byte[] 8
[void]$fs.Read($b, 0, 8)
$fs.Close()
# ★PNG の 印★ 89 50 4E 47 0D 0A 1A 0A
$印 = ($b | ForEach-Object { $_.ToString('x2') }) -join ''
Write-Host ('★頭 8バイト★ ' + $印)
if ($印 -ne '89504e470d0a1a0a') { Write-Host '★★PNG の 印では ありません★★'; exit 9 }

# ══ ★★絵の 中身を 数えます★★ ══
#   ★頭が PNG★ だけでは ★真っ白でも 通ります★（2026-09-21 に 1回 通しました）
#   ⇒★白では ない 点を 数えます★＝0個なら 赤
Add-Type -AssemblyName System.Drawing
$絵 = [Drawing.Bitmap]::FromFile($出す先)
$白でない = 0
$色たち = New-Object 'System.Collections.Generic.HashSet[int]'
for ($y = 0; $y -lt $絵.Height; $y += 2) {
  for ($x2 = 0; $x2 -lt $絵.Width; $x2 += 2) {
    $c = $絵.GetPixel($x2, $y)
    if ($c.R -lt 250 -or $c.G -lt 250 -or $c.B -lt 250) { $白でない++ }
    [void]$色たち.Add($c.ToArgb())
  }
}
$幅 = $絵.Width; $高 = $絵.Height
$絵.Dispose()
Write-Host ('★絵の 大きさ★ ' + $幅 + ' x ' + $高 + ' 点')
Write-Host ('★白では ない 点★ ' + $白でない + '個（2点おきに 数えました）／ ★色の 種類★ ' + $色たち.Count + '色')
if ($白でない -eq 0) { Write-Host '★★真っ白です＝中身が 在りません★★'; exit 10 }
if ($色たち.Count -lt 3) { Write-Host '★★色が 2色 以下＝飾りが 出て いません★★'; exit 11 }

Write-Host ''
Write-Host ('★★描かせました★★ ... ' + $出す先)
Write-Host ('  ★大きさ★ ' + $x.Length + ' バイト ／ ★sha256★ ' + (Get-FileHash $出す先 -Algorithm SHA256).Hash.ToLower())
