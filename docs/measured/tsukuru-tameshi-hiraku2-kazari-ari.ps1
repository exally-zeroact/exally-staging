# tsukuru-tameshi-hiraku2-kazari-ari.ps1
#   -- ★「判子・罫線が 残るか」を 測れる 材料を 実Excel に 作らせる★（2026-09-20）
#
#  ★★なぜ 作り直すか（★私の 落ち度★）★★
#    1本目（`tsukuru-tameshi-hiraku.ps1`）で 作った 材料は
#      罫線の 型 ★1★（＝空の 既定だけ）／塗りの 型 ★2★（＝既定だけ）／絵・判子 ★0本★
#    ⇒★飾りが 1つも 入って いません★
#    ⇒★だから 「受け取った ファイルを 保存し直すと 判子・罫線が 残るか」は
#      あの 材料では ★測れません★★（★守る 物が 無いから 「守れた」と 言えない★）
#    ⇒記憶「★掃く 窓は 押す前に 隙間を 計算して から 置く★」の 形
#
#  ★★入れる 飾り（★1つずつ 別の 種類★）★★
#    ㋐罫線 ......... A1:C3 の 四方（`BorderAround`）＋ B2 の 下線だけ 太く
#    ㋑塗り ......... B1 を 黄色
#    ㋒字の 飾り ..... A1 を 太字＋赤
#    ㋓表示の 形 ..... B1 を 通貨（`"\",\"#,##0"`）／A3 を パーセント
#    ㋔繋げた マス ... A5:C5 を 1つに して 字を 入れる
#    ㋕★判子の 代わり★ ... 四角の 図形 1個（`Shapes.AddShape`）
#    ㋖列の 幅 ....... A 列を 30
#    ㋗溢れる 式 ..... C1 に `=SEQUENCE(3)`（★C1:C3★）
#    ㋘ただの 式 ..... B1 に `=SUM(A1:A3)`
#
#  ★★置き場★★ `%TEMP%\exally-tameshi-hiraku2.xlsx`（★1本の 名★）
#    ・★司さんの 実物の 名は 1文字も 在りません★
#    ・★既に 在れば 消してから 作ります★（★消すのは この 1本だけ★）
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③★書いた 後に 読み返して 飾りが 全部 在るか 数える★（1つでも 欠けたら exit 5）
#      ＝★「作った」は 「在る」では ありません★
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$出す先 = Join-Path $env:TEMP 'exally-tameshi-hiraku2.xlsx'
$許す名 = 'exally-tameshi-hiraku2.xlsx'
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

  # ── ★中身★
  $sh.Range('A1').Value2 = 3
  $sh.Range('A2').Value2 = 1
  $sh.Range('A3').Value2 = 0.25
  $sh.Range('D1').Value2 = 'abc'
  $sh.Range('B1').Formula2 = '=SUM(A1:A3)'
  $sh.Range('C1').Formula2 = '=SEQUENCE(3)'
  $sh.Range('E1').Formula2 = '=D1&"!"'

  # ── ㋐罫線
  $sh.Range('A1:C3').BorderAround(1, 3) | Out-Null   # xlContinuous ／ xlThick
  $sh.Range('B2').Borders.Item(9).LineStyle = 1      # xlEdgeBottom
  $sh.Range('B2').Borders.Item(9).Weight = 4         # xlThick
  # ── ㋑塗り
  $sh.Range('B1').Interior.Color = 65535             # 黄
  # ── ㋒字の 飾り
  $sh.Range('A1').Font.Bold = $true
  $sh.Range('A1').Font.Color = 255                   # 赤
  # ── ㋓表示の 形
  $sh.Range('B1').NumberFormatLocal = '#,##0'
  $sh.Range('A3').NumberFormatLocal = '0.0%'
  # ── ㋔繋げた マス
  $sh.Range('A5:C5').Merge() | Out-Null
  $sh.Range('A5').Value2 = 'kazari no tame no musubi'
  # ── ㋕判子の 代わり（図形 1個）
  $かたち = $sh.Shapes.AddShape(1, 200, 20, 60, 60)  # msoShapeRectangle
  $かたち.Name = 'hanko'
  # ── ㋖列の 幅
  $sh.Columns.Item(1).ColumnWidth = 30

  if (Test-Path $出す先) { Remove-Item $出す先 -Force }
  $bk.SaveAs($出す先, 51)
  $bk.Close($false)
  $bk = $null

  # ══ ★★読み返して 数える★★（★「作った」は 「在る」では ない★）══
  $bk = $xl.Workbooks.Open($出す先, 0, $true)
  $sh = $bk.Sheets.Item(1)
  $欠け = 0
  $見た = New-Object System.Collections.Generic.List[string]

  $四方 = [string]$sh.Range('A1').Borders.Item(7).LineStyle   # xlEdgeLeft
  $見た.Add('罫線(A1 左) LineStyle=' + $四方)
  if ($四方 -eq '-4142') { $欠け++ }                          # xlLineStyleNone

  $下線 = [string]$sh.Range('B2').Borders.Item(9).Weight
  $見た.Add('罫線(B2 下) Weight=' + $下線)
  if ($下線 -ne '4') { $欠け++ }

  $塗り = [string]$sh.Range('B1').Interior.Color
  $見た.Add('塗り(B1) Color=' + $塗り)
  if ($塗り -ne '65535') { $欠け++ }

  $太 = [string]$sh.Range('A1').Font.Bold
  $色 = [string]$sh.Range('A1').Font.Color
  $見た.Add('字(A1) Bold=' + $太 + ' Color=' + $色)
  if ($太 -ne 'True' -or $色 -ne '255') { $欠け++ }

  $形1 = [string]$sh.Range('B1').NumberFormatLocal
  $形2 = [string]$sh.Range('A3').NumberFormatLocal
  $見た.Add('表示の形 B1=' + $形1 + ' A3=' + $形2)
  if ($形1 -notmatch '#') { $欠け++ }
  if ($形2 -notmatch '%') { $欠け++ }

  $結び = [string]$sh.Range('A5').MergeCells
  $見た.Add('繋げたマス A5=' + $結び)
  if ($結び -ne 'True') { $欠け++ }

  $図 = [string]$sh.Shapes.Count
  $見た.Add('図形の数=' + $図)
  if ($図 -ne '1') { $欠け++ }

  $幅 = [string]$sh.Columns.Item(1).ColumnWidth
  $見た.Add('A列の幅=' + $幅)

  $溢れ = [string]$sh.Range('C1').Value2 + '/' + [string]$sh.Range('C2').Value2 + '/' + [string]$sh.Range('C3').Value2
  $見た.Add('C1:C3=' + $溢れ + ' 式=' + [string]$sh.Range('C1').Formula)
  if ($溢れ -ne '1/2/3') { $欠け++ }

  $見た.Add('B1=' + [string]$sh.Range('B1').Value2 + ' 式=' + [string]$sh.Range('B1').Formula)

  foreach ($l in $見た) { Write-Host ('  ' + $l) }
  $bk.Close($false)
  $bk = $null

  if ($欠け -ne 0) { Write-Host ('★★飾りが ' + $欠け + '個 欠けて います★★'); exit 5 }
  $出来た = $true
} finally {
  # ★★2026-09-20 ── ★図形の 持ち手も 手放す★★
  #   ＝`$かたち`（`Shapes.AddShape` が 返す 物）を $null に して いなかった ので
  #     ★Excel が 120秒 では 消えませんでした★（3本目で 実際に 出ました）
  #   ⇒記憶「実Excel を COM＝5.1 かつ ★掴んだ物 全部 $null★ の 2つ 揃った 時だけ 消える」
  $かたち = $null
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
  Write-Host '  ★飾り 欠け 0個★'
}
