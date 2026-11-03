# tsukuru-xlsb-ita-2mai.ps1
#   -- ★板が 2枚 在る `.xlsb` を 作る★（82）（2026-09-21）
#
#  ★★なぜ★★
#    Exally1 が `.xlsb` の ★板の 名前 ⇒ 部品名★ を ★「並び」で 当てて います★
#    （この repo が 前から そう して います＝`saveXlsb`）
#    ⇒★板が 2枚 以上で 合うかは ★未測定★★
#    ⇒★合わないと 判子が ★別の 板★に 出ます★
#      ＝★消えるのでは なく ずれる＝一番 気付きにくい★
#
#  ★★中身★★
#    ★Sheet1★ ... 図形 `hanko1`（四角）／A1 = 'ita1'
#    ★Sheet2★ ... 図形 `hanko2`（丸）  ／A1 = 'ita2'
#    ⇒★どちらに 出るかで 割れます★
#    ⇒★形も 変えます★（四角／丸）＝★名前だけでなく 形でも 見分けられる★
#
#  ★★置き場★★ `%TEMP%\exally-tameshi-xlsb2.xlsb`（★1本の 名★）
#    ・★司さんの 実物の 名は 1文字も 在りません★／★作り物です★
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③★読み返して 板が 2枚／図形が 各1個／名が hanko1 hanko2★（1つでも 違えば exit 5）
#    ④拡張子が `.xlsb`（exit 7）／⑤包みの 頭が `PK`（exit 9）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

# ★★`-いれかえ` ＝ ★並びと 部品の 番号を わざと ずらす★★
#   ＝`Ita2` を 前へ 動かす ⇒ 板の 並びは Ita2, Ita1／部品は sheet1=Ita1, sheet2=Ita2
#   ⇒★「並びで 当てる」やり方だと ★判子が 別の 板に 出ます★★
#   ⇒★これが 本当の 試しです★（並びが 揃った 物では 割れません）
param([switch]$いれかえ)

$出す先 = if ($いれかえ) { Join-Path $env:TEMP 'exally-tameshi-xlsb3.xlsb' }
          else { Join-Path $env:TEMP 'exally-tameshi-xlsb2.xlsb' }
$許す名 = if ($いれかえ) { 'exally-tameshi-xlsb3.xlsb' } else { 'exally-tameshi-xlsb2.xlsb' }
if ((Split-Path $出す先 -Leaf) -ne $許す名) { Write-Host '★★書いて よい 名は 1本だけです★★'; exit 7 }

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null; $sh2 = $null; $か1 = $null; $か2 = $null
$出来た = $false
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)

  # ── ★板 1枚目★
  $sh.Name = 'Ita1'
  $sh.Range('A1').Value2 = 'ita1'
  $か1 = $sh.Shapes.AddShape(1, 100, 20, 60, 60)   # 1 ＝ 四角
  $か1.Name = 'hanko1'
  # ── ★板 2枚目★
  # ★★2026-09-21 ── `Sheets.Add()` は ★今の 板の 前★ に 入ります★★
  #   ＝1回目は Ita2 が 1枚目に なりました（★並びが 逆★）
  #   ⇒★後ろに 入れる★＝第2引数（After）に 1枚目を 渡します
  #   ⇒★「作った 順」と「並びの 順」は 別物★＝★これ 自体が 落とし穴★
  $sh2 = $bk.Sheets.Add([Type]::Missing, $sh)
  $sh2.Name = 'Ita2'
  $sh2.Range('A1').Value2 = 'ita2'
  $か2 = $sh2.Shapes.AddShape(9, 200, 40, 50, 50)  # 9 ＝ 丸（msoShapeOval）
  $か2.Name = 'hanko2'
  # ★★`-いれかえ` の 時だけ Ita2 を 前へ 動かします★★
  if ($いれかえ) { $sh2.Move($sh) | Out-Null }
  $bk.Sheets.Item(1).Activate() | Out-Null

  if (Test-Path $出す先) { Remove-Item $出す先 -Force }
  # ★★50 ＝ xlExcel12（.xlsb）★★（51 は xlsx＝取り違えると 別の 形）
  $bk.SaveAs($出す先, 50)
  $bk.Close($false)
  $bk = $null

  # ══ ★★読み返して 数える★★ ══
  $bk = $xl.Workbooks.Open($出す先, 0, $true)
  $sh = $bk.Sheets.Item(1)
  $欠け = 0
  Write-Host ''
  Write-Host '★★読み返し★★'
  $枚 = [int]$bk.Sheets.Count
  Write-Host ('  板の 数 ... ' + $枚 + '（待つ 2）')
  if ($枚 -ne 2) { $欠け++ }
  $まち1 = if ($いれかえ) { 'Ita2' } else { 'Ita1' }
  $まち2 = if ($いれかえ) { 'Ita1' } else { 'Ita2' }
  Write-Host ('  ★待つ 並び★ 1枚目 ' + $まち1 + ' ／ 2枚目 ' + $まち2)
  foreach ($n in 1, 2) {
    $s3 = $bk.Sheets.Item($n)
    $名 = [string]$s3.Name
    $数 = [string]$s3.Shapes.Count
    $図 = '(無し)'
    try { if ($s3.Shapes.Count -ge 1) { $図 = [string]$s3.Shapes.Item(1).Name } } catch { }
    $値 = [string]$s3.Range('A1').Value2
    Write-Host ('  ' + $n + '枚目 ... 名 ' + $名.PadRight(6) + ' A1 ' + $値.PadRight(6) + ' 図形 ' + $数 + '個 名 ' + $図)
    if ($数 -ne '1') { $欠け++ }
    $s3 = $null
  }
  $い1 = [string]$bk.Sheets.Item(1).Shapes.Item(1).Name
  $い2 = [string]$bk.Sheets.Item(2).Shapes.Item(1).Name
  $ま1 = if ($いれかえ) { 'hanko2' } else { 'hanko1' }
  $ま2 = if ($いれかえ) { 'hanko1' } else { 'hanko2' }
  if ($い1 -ne $ま1) { $欠け++ }
  if ($い2 -ne $ま2) { $欠け++ }
  $bk.Close($false)
  $bk = $null
  Write-Host ('★★欠け ... ' + $欠け + '個★★')
  if ($欠け -ne 0) { exit 5 }
  $出来た = $true
} finally {
  # ★★2026-09-20 ── ★図形の 持ち手も 手放す★★
  #   ＝`$かたち`（`Shapes.AddShape` が 返す 物）を $null に して いなかった ので
  #     ★Excel が 120秒 では 消えませんでした★（3本目で 実際に 出ました）
  #   ⇒記憶「実Excel を COM＝5.1 かつ ★掴んだ物 全部 $null★ の 2つ 揃った 時だけ 消える」
  $か1 = $null
  $か2 = $null
  $sh2 = $null
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
  # ══ ★★門④⑤＝出来た 物が 本当に `.xlsb` か★★ ══
  if ((Split-Path $出す先 -Leaf) -notlike '*.xlsb') { Write-Host '★★.xlsb では ありません★★'; exit 7 }
  $頭 = [System.IO.File]::ReadAllBytes($出す先)[0..1]
  if ($頭[0] -ne 80 -or $頭[1] -ne 75) { Write-Host '★★包みの 頭が PK では ありません★★'; exit 9 }
  Write-Host '★包みの 頭 ... PK（zip の 形）★'
  $x = Get-Item $出す先
  $h = (Get-FileHash $出す先 -Algorithm SHA256).Hash.ToLower()
  Write-Host ''
  Write-Host ('★★作りました★★ ... ' + $出す先)
  Write-Host ('  ★大きさ★ ' + $x.Length + ' バイト')
  Write-Host ('  ★sha256★ ' + $h)
}
