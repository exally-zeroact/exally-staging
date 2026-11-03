# tsukuru-tameshi-hiraku3-kazari-to-kobore3.ps1
#   -- ★飾り 8種 ＋ 溢れ 3種 を 1枚に した 材料を 実Excel に 作らせる★（2026-09-20）
#
#  ★★なぜ★★
#    Exally1 が 測った 所:
#      受け取った ファイルを 読むと 溢れ先に ★`_台溢れ元`★ が 付く
#      ＝★実Excel が 作った `cm="1"` 付きの 溢れを ★台★ が 引き受けて いる★
#      ＝★借り物では なく うちの 台が 答えて いる★
#    ⇒★但し 今 それが 付いたのは ★1本の 式（=SEQUENCE(3)）だけ★★
#    ⇒★縦・横・2次元 の 3種で 付けば 「読む 側は もう 台で 立って いる」が 強く なります★
#    ⇒★借り物外しの 本丸に 直に 効きます★
#
#  ★★置く 所を 先に 決めました（★ぶつからない 隙間を 計算して から★）★★
#    ★記憶「掃く 窓は 押す前に 隙間を 計算して から 置く」★
#    ★今日 Exally1 が 横3 の 行き先を 2次元で 塞いで 1回 踏んだ 所です★
#
#      列        A        B        C        D        E   F   G
#      行1     3(太赤)  SUM      ┌溢れ縦   abc      ┌── 溢れ横 ──┐
#      行2     1                 │                 （空）
#      行3     0.25(%)           └                 （空）
#      行5     ├──── 繋げた マス ────┤   =D1&"!"
#      行10    ┌──── 溢れ 2次元 ────┐
#      行11    └─────────────────┘
#
#    ★重なりを 1つだけ わざと 残しました★
#      罫線 A1:C3 ★の 中に★ 溢れ縦 C1:C3 が 在る
#      ⇒★実務に 近い★（罫線で 囲った 表の 中に 式が 在る）
#
#  ★★中身★★
#    ★飾り 8種★（2本目の 材料と 同じ）
#      ㋐罫線 A1:C3 の 四方（太線）＋ B2 の 下に 太い 線
#      ㋑塗り B1 を 黄        ㋒字 A1 を 太字＋赤
#      ㋓表示の 形 B1 `#,##0` ／ A3 `0.0%`
#      ㋔繋げた マス A5:C5    ㋕図形 1個（名 `hanko`）
#      ㋖A列の 幅 30
#    ★溢れ 3種★
#      ㋗縦   C1  `=SEQUENCE(3)`   ⇒ C1:C3
#      ㋘横   E1  `=SEQUENCE(1,3)` ⇒ E1:G1
#      ㋙2次元 A10 `=SEQUENCE(2,3)` ⇒ A10:C11
#    ★対照（★溢れない 式★）★
#      ㋚B1 `=SUM(A1:A3)` ⇒ 4.25
#      ㋛D5 `=D1&"!"`     ⇒ abc!
#      ⇒★Exally1 の お願い＝「溢れだけ 壊れた」と「全部 壊れた」を 見分ける 為★
#
#  ★★置き場★★ `%TEMP%\exally-tameshi-hiraku3.xlsx`（★1本の 名★）
#    ・★司さんの 実物の 名は 1文字も 在りません★
#    ・★既に 在れば 消してから 作ります★（★消すのは この 1本だけ★）
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③★書いた 後に 読み返して 飾り 8種 ＋ 溢れ 3種 ＋ 対照 2種 が 在るか 数える★
#      ⇒★1つでも 欠けたら exit 5★（★「作った」は 「在る」では ない★）
#    ④★溢れが 互いを 塞いで いないか★＝3種 とも 待つ 形に 広がって いるか
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$出す先 = Join-Path $env:TEMP 'exally-tameshi-hiraku3.xlsx'
$許す名 = 'exally-tameshi-hiraku3.xlsx'
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

  # ── ★ただの 数と 字★
  $sh.Range('A1').Value2 = 3
  $sh.Range('A2').Value2 = 1
  $sh.Range('A3').Value2 = 0.25
  $sh.Range('D1').Value2 = 'abc'
  # ── ★対照（溢れない 式）★
  $sh.Range('B1').Formula2 = '=SUM(A1:A3)'
  $sh.Range('D5').Formula2 = '=D1&"!"'
  # ── ★溢れ 3種★（★`Formula` では 溢れません／`Formula2` です★）
  $sh.Range('C1').Formula2 = '=SEQUENCE(3)'
  $sh.Range('E1').Formula2 = '=SEQUENCE(1,3)'
  $sh.Range('A10').Formula2 = '=SEQUENCE(2,3)'
  # ── ★飾り★
  $sh.Range('A1:C3').BorderAround(1, 3) | Out-Null
  $sh.Range('B2').Borders.Item(9).LineStyle = 1
  $sh.Range('B2').Borders.Item(9).Weight = 4
  $sh.Range('B1').Interior.Color = 65535
  $sh.Range('A1').Font.Bold = $true
  $sh.Range('A1').Font.Color = 255
  $sh.Range('B1').NumberFormatLocal = '#,##0'
  $sh.Range('A3').NumberFormatLocal = '0.0%'
  $sh.Range('A5:C5').Merge() | Out-Null
  $sh.Range('A5').Value2 = 'kazari no tame no musubi'
  $かたち = $sh.Shapes.AddShape(1, 320, 20, 60, 60)
  $かたち.Name = 'hanko'
  $sh.Columns.Item(1).ColumnWidth = 30

  if (Test-Path $出す先) { Remove-Item $出す先 -Force }
  $bk.SaveAs($出す先, 51)
  $bk.Close($false)
  $bk = $null

  # ══ ★★読み返して 数える★★ ══
  $bk = $xl.Workbooks.Open($出す先, 0, $true)
  $sh = $bk.Sheets.Item(1)
  $欠け = 0
  $見た = New-Object System.Collections.Generic.List[string]
  function 見る {
    param($名, $待, $出)
    $判 = 'ok'
    if ([string]$出 -ne [string]$待) { $判 = '★違う★' }
    return @($名, [string]$待, [string]$出, $判)
  }
  $表 = New-Object System.Collections.Generic.List[object]
  # ★飾り 8種★
  $表.Add((見る '罫線(A1 左)'      1        ([string]$sh.Range('A1').Borders.Item(7).LineStyle)))
  $表.Add((見る '罫線(B2 下) 太さ' 4        ([string]$sh.Range('B2').Borders.Item(9).Weight)))
  $表.Add((見る '塗り(B1)'         65535    ([string]$sh.Range('B1').Interior.Color)))
  $表.Add((見る '字(A1) 太字'      'True'   ([string]$sh.Range('A1').Font.Bold)))
  $表.Add((見る '字(A1) 色'        255      ([string]$sh.Range('A1').Font.Color)))
  $表.Add((見る '繋げたマス(A5)'   'True'   ([string]$sh.Range('A5').MergeCells)))
  $表.Add((見る '図形の数'         1        ([string]$sh.Shapes.Count)))
  $表.Add((見る 'A列の幅'          30       ([string]$sh.Columns.Item(1).ColumnWidth)))
  # ★溢れ 3種（★広がった 形を 数で 見ます★）★
  $表.Add((見る '溢れ縦 C1:C3'     '1/2/3'  (([string]$sh.Range('C1').Value2) + '/' + ([string]$sh.Range('C2').Value2) + '/' + ([string]$sh.Range('C3').Value2))))
  $表.Add((見る '溢れ横 E1:G1'     '1/2/3'  (([string]$sh.Range('E1').Value2) + '/' + ([string]$sh.Range('F1').Value2) + '/' + ([string]$sh.Range('G1').Value2))))
  $表.Add((見る '溢れ2次元 A10:C11' '1/2/3/4/5/6' (([string]$sh.Range('A10').Value2) + '/' + ([string]$sh.Range('B10').Value2) + '/' + ([string]$sh.Range('C10').Value2) + '/' + ([string]$sh.Range('A11').Value2) + '/' + ([string]$sh.Range('B11').Value2) + '/' + ([string]$sh.Range('C11').Value2))))
  # ★式の 字（★頭だけに 式が 在る＝溢れて いる 証し★）★
  $表.Add((見る 'C1 の式'   '=SEQUENCE(3)'   ([string]$sh.Range('C1').Formula)))
  $表.Add((見る 'C2 の式'   ''               ([string]$sh.Range('C2').Formula)))
  $表.Add((見る 'E1 の式'   '=SEQUENCE(1,3)' ([string]$sh.Range('E1').Formula)))
  $表.Add((見る 'F1 の式'   ''               ([string]$sh.Range('F1').Formula)))
  $表.Add((見る 'A10 の式'  '=SEQUENCE(2,3)' ([string]$sh.Range('A10').Formula)))
  $表.Add((見る 'B10 の式'  ''               ([string]$sh.Range('B10').Formula)))
  # ★対照 2種★
  $表.Add((見る '対照 B1 の式' '=SUM(A1:A3)' ([string]$sh.Range('B1').Formula)))
  $表.Add((見る '対照 B1 の値' 4.25          ([string]$sh.Range('B1').Value2)))
  $表.Add((見る '対照 D5 の値' 'abc!'        ([string]$sh.Range('D5').Value2)))
  # ★溢れが 互いを 塞いで いないか＝#SPILL! が 0件★
  $塞ぎ = 0
  foreach ($ma in @('C1', 'E1', 'A10')) {
    $v = $sh.Range($ma).Value2
    if ($v -isnot [double]) { $塞ぎ++ }
  }
  $表.Add((見る '塞がった 溢れの 数' 0 ([string]$塞ぎ)))

  foreach ($t2 in $表) {
    if ($t2[3] -ne 'ok') { $欠け++ }
    Write-Host ('  ' + $t2[0].PadRight(20) + ' 待つ ' + $t2[1].PadRight(14) + ' 出た ' + $t2[2].PadRight(14) + ' ' + $t2[3])
  }
  # ★表示の 形は 「含むか」で 見ます★
  $形1 = [string]$sh.Range('B1').NumberFormatLocal
  $形2 = [string]$sh.Range('A3').NumberFormatLocal
  if ($形1 -notmatch '#') { $欠け++ }
  if ($形2 -notmatch '%') { $欠け++ }
  Write-Host ('  表示の形             B1 ' + $形1 + ' ／ A3 ' + $形2)
  $図名 = '(無し)'
  try { if ($sh.Shapes.Count -ge 1) { $図名 = [string]$sh.Shapes.Item(1).Name } } catch { }
  if ($図名 -ne 'hanko') { $欠け++ }
  Write-Host ('  図形の名             ' + $図名)

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
}
