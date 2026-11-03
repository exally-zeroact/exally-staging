# tsukuru-tameshi-hiraku.ps1
#   -- ★「ファイルを 開く 道」を 測る 為の 材料を ★実Excel に 作らせる★★（2026-09-20）
#
#  ★★なぜ 要るか★★
#    Exally1 が 借り物の 呼び口 16回を 数えた 時、`setSheetContent`（`_pushGrid`）だけ
#    ★0件の まま★でした。訳は ★ファイルを 食わせる 道を 通して いない★から です。
#    ⇒★0件を「使って いない」と 読んでは いけません★
#    ⇒★だから 食わせる 物が 要ります★
#
#  ★★実Excel に 作らせる 訳★★
#    ・うちの 書き出しで 作ると ★うちの 癖が 混ざります★（今日 それで #SPILL! の 穴が 出た）
#    ・★お客さんが 持って くる ファイルは 実Excel が 作った 物★です
#    ⇒★測る 材料は 実物に 近い 方が よい★
#
#  ★★中身（★式を 3通り 入れます★）★★
#    A1:A3 = 3 / 1 / 2      ＝ ただの 数
#    B1    = `=SUM(A1:A3)`  ＝ ★溢れない 式★
#    C1    = `=SEQUENCE(3)` ＝ ★溢れる 式★（C1:C3）
#    D1    = `abc`          ＝ 字
#    E1    = `=D1&"!"`      ＝ ★字を 使う 式★
#
#  ★★置き場★★ `%TEMP%\exally-tameshi-hiraku.xlsx`（★1本の 名★）
#    ・★司さんの 実物の 名は 1文字も 在りません★
#    ・★既に 在れば 消してから 作ります★（★消すのは この 1本だけ★）
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③書いた 後に 読み返して ★式が 3つとも 在るか★ 見る（exit 5）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$出す先 = Join-Path $env:TEMP 'exally-tameshi-hiraku.xlsx'
$許す名 = 'exally-tameshi-hiraku.xlsx'
if ((Split-Path $出す先 -Leaf) -ne $許す名) { Write-Host '★★書いて よい 名は 1本だけです★★'; exit 7 }

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null
$出来た = $false
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)
  $sh.Range('A1').Value2 = 3
  $sh.Range('A2').Value2 = 1
  $sh.Range('A3').Value2 = 2
  $sh.Range('D1').Value2 = 'abc'
  # ★`Formula2` を 使います★＝`Formula` では 溢れません（2026-09-19 実測）
  $sh.Range('B1').Formula2 = '=SUM(A1:A3)'
  $sh.Range('C1').Formula2 = '=SEQUENCE(3)'
  $sh.Range('E1').Formula2 = '=D1&"!"'

  if (Test-Path $出す先) { Remove-Item $出す先 -Force }
  $bk.SaveAs($出す先, 51)
  $bk.Close($false)
  $bk = $null

  # ★★読み返して 確かめます★★（★書いた だけでは 「在る」に なりません★）
  $bk = $xl.Workbooks.Open($出す先, 0, $true)
  $sh = $bk.Sheets.Item(1)
  $式の数 = 0
  foreach ($ma in @('B1', 'C1', 'E1')) {
    $f = [string]$sh.Range($ma).Formula
    Write-Host ('  ' + $ma + ' ... 値 ' + [string]$sh.Range($ma).Value2 + ' ／ 式 ' + $f)
    if ($f.StartsWith('=')) { $式の数++ }
  }
  Write-Host ('  C1:C3 ... ' + [string]$sh.Range('C1').Value2 + ' / ' + [string]$sh.Range('C2').Value2 + ' / ' + [string]$sh.Range('C3').Value2)
  $bk.Close($false)
  $bk = $null
  if ($式の数 -ne 3) { Write-Host ('★★式が 3つ 在りません ... ' + $式の数 + '個★★'); exit 5 }
  $出来た = $true
} finally {
  $sh = $null
  if ($null -ne $bk) { $bk.Close($false) }
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
  if ($残り -eq 0) { Write-Host ('★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒で ★消えました★★') }
  else { Write-Host ('★★Excel は 消えませんでした★★ ／ 残り ' + $残り + '個') }
}
if ($出来た) {
  $x = Get-Item $出す先
  $h = (Get-FileHash $出す先 -Algorithm SHA256).Hash.ToLower()
  Write-Host ''
  Write-Host ('★★作りました★★ ... ' + $出す先)
  Write-Host ('  ★大きさ★ ' + $x.Length + ' バイト')
  Write-Host ('  ★sha256★ ' + $h)
}
