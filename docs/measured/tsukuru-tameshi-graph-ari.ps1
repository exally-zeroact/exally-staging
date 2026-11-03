# tsukuru-tameshi-graph-ari.ps1
#   -- ★グラフ入りの 材料を 実Excel に 作らせる★（94）（2026-09-21）
#
#  ★★なぜ★★
#    Exally1「★グラフが 減って いないか★を 数えたい。でも 材料に グラフが 在りません」
#    ⇒★数えると 書いたのに 材料に 無い＝数えた 事に できません★
#    ⇒★私が 作ります★
#
#  ★★中身（★1つずつ 別の マスに 置きます★）★★
#    A1:B5 ... グラフの 元に なる 数（見出し 1行 ＋ 4行）
#    D1     ... ★太字・赤★
#    D2     ... ★黄色の 塗り★
#    D3     ... ★下に 太い 線★
#    D5     ... `=SEQUENCE(3)`（★縦に 溢れる★）
#    A7:C7  ... ★繋げた マス★
#    判子   ... 四角 1つ（`hanko`）
#    グラフ ... ★棒グラフ 1つ★（A1:B5 から）
#
#  ★★「図形の 数」に 気を つけます★★
#    ★グラフも 図形（Shapes）に 数えられます★
#    ⇒`Shapes.Count` は ★2★ に なります（判子 ＋ グラフ）
#    ⇒★門では `ChartObjects().Count` と 分けて 数えます★
#      （★ここを 分けずに 「図形 1つ」と 書くと 嘘に なります★）
#
#  ★★置き場★★ `%TEMP%\exally-tameshi-graph.xlsx`（★1本の 名★）
#    ・★司さんの 実物の 名は 1文字も 在りません★／中身は ★作り物の 数だけ★
#
#  ★門★
#    ①貝殻が 5.1（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③★書いた 後に 読み返して 7つ 全部 在るか★（1つでも 欠けたら exit 5）
#    ④★包みの 頭が PK★（exit 9）／⑤★包みに グラフの 部品が 在るか★（exit 6）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$出す先 = Join-Path $env:TEMP 'exally-tameshi-graph.xlsx'
if ((Split-Path $出す先 -Leaf) -ne 'exally-tameshi-graph.xlsx') { exit 7 }

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

  # ── ★グラフの 元★
  $sh.Range('A1').Value2 = [string]'tsuki'
  $sh.Range('B1').Value2 = [string]'uriage'
  $元 = 10, 30, 20, 40
  for ($i = 0; $i -lt 4; $i++) {
    $sh.Cells.Item($i + 2, 1).Value2 = [string]('m' + ($i + 1))
    $sh.Cells.Item($i + 2, 2).Value2 = [double]$元[$i]
  }

  # ── ★飾り（1つずつ 別の マス）★
  $sh.Range('D1').Value2 = [double]1
  $sh.Range('D1').Font.Bold = $true
  $sh.Range('D1').Font.Color = [int]255                 # ★COM は BGR＝赤★
  $sh.Range('D2').Value2 = [double]2
  $sh.Range('D2').Interior.Color = [int]65535           # ★黄色★
  $sh.Range('D3').Value2 = [double]3
  $sh.Range('D3').Borders.Item(9).LineStyle = 1
  $sh.Range('D3').Borders.Item(9).Weight = 4            # ★太い★
  $sh.Range('D5').Formula2 = '=SEQUENCE(3)'             # ★縦に 溢れる★
  $sh.Range('A7:C7').Merge() | Out-Null
  $sh.Range('A7').Value2 = [string]'tsunageta masu'

  # ── ★判子★
  $sh.Shapes.AddShape(1, 330, 20, 60, 60).Name = 'hanko'

  # ── ★グラフ★（★棒★）
  $枠 = $sh.ChartObjects().Add(330, 110, 240, 150)
  $枠.Name = 'graph1'
  $枠.Chart.SetSourceData($sh.Range('A1:B5'))
  $枠.Chart.ChartType = 51                              # ★51 ＝ 縦棒★
  $枠 = $null

  # ── ★★読み返して 数えます★★
  $図全部 = [int]$sh.Shapes.Count
  $グラフ = [int]$sh.ChartObjects().Count
  $判子 = $図全部 - $グラフ
  $確 = @(
    ('太字 D1 = ' + [string]$sh.Range('D1').Font.Bold),
    ('字の色 D1 = ' + [string]$sh.Range('D1').Font.Color),
    ('塗り D2 = ' + [string]$sh.Range('D2').Interior.Color),
    ('線 D3 = ' + [string]$sh.Range('D3').Borders.Item(9).Weight),
    ('溢れ D5:D7 = ' + [string]$sh.Range('D5').Value2 + '/' + [string]$sh.Range('D7').Value2),
    ('繋げた A7 = ' + [string]$sh.Range('A7').MergeCells),
    ('判子 = ' + $判子 + '個 ／ グラフ = ' + $グラフ + '個（図形 全部 ' + $図全部 + '個）')
  )
  Write-Host ''
  foreach ($c in $確) { Write-Host ('  ' + $c) }
  $欠け = 0
  if ([string]$sh.Range('D1').Font.Bold -ne 'True') { $欠け++ }
  if ([int]$sh.Range('D1').Font.Color -ne 255) { $欠け++ }
  if ([int]$sh.Range('D2').Interior.Color -ne 65535) { $欠け++ }
  if ([int]$sh.Range('D3').Borders.Item(9).Weight -ne 4) { $欠け++ }
  if ([string]$sh.Range('D7').Value2 -ne '3') { $欠け++ }
  if ([string]$sh.Range('A7').MergeCells -ne 'True') { $欠け++ }
  if ($判子 -ne 1) { $欠け++ }
  if ($グラフ -ne 1) { $欠け++ }
  Write-Host ('★★欠け ... ' + $欠け + '個★★')
  if ($欠け -ne 0) { Write-Host '★★材料が 揃って いません★★'; exit 5 }

  if (Test-Path $出す先) { Remove-Item $出す先 -Force }
  $bk.SaveAs($出す先, 51)
  $bk.Close($false); $sh = $null; $bk = $null
} finally {
  $sh = $null
  if ($null -ne $bk) { $bk.Close($false) }
  $bk = $null
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $xl = $null
  $t = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 300)) { Start-Sleep -Milliseconds 250 }
  $t.Stop()
  $残り = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
  if ($残り -eq 0) { Write-Host ('★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒で ★消えました★★') }
  else { Write-Host ('★★Excel は 消えませんでした★★ ／ 残り ' + $残り + '個') }
}

if (-not (Test-Path $出す先)) { exit 5 }
$fs = [IO.File]::OpenRead($出す先)
$b = New-Object byte[] 2
[void]$fs.Read($b, 0, 2)
$fs.Close()
if (([char]$b[0] + [string][char]$b[1]) -ne 'PK') { Write-Host '★★包みの 頭が PK では ありません★★'; exit 9 }

Add-Type -AssemblyName System.IO.Compression.FileSystem
$z = [IO.Compression.ZipFile]::OpenRead($出す先)
$名たち = $z.Entries | ForEach-Object { $_.FullName }
$z.Dispose()
$グラフ部品 = @($名たち | Where-Object { $_ -like 'xl/charts/*' }).Count
$図部品 = @($名たち | Where-Object { $_ -like 'xl/drawings/*' }).Count
Write-Host ''
Write-Host ('★包みの 部品★ ' + $名たち.Count + '本 ／ グラフ ' + $グラフ部品 + '本 ／ 図 ' + $図部品 + '本')
if ($グラフ部品 -lt 1) { Write-Host '★★包みに グラフの 部品が 在りません★★'; exit 6 }

Write-Host ('★★作りました★★ ... ' + $出す先)
Write-Host ('  ★大きさ★ ' + (Get-Item $出す先).Length + ' バイト ／ ★sha256★ ' + (Get-FileHash $出す先 -Algorithm SHA256).Hash.ToLower())
