# tsukuru-xlsb-ji-to-sen-1tsu-zutsu.ps1
#   -- ★`43`（字）と `46`（線）の 中身を 割る 材料★（87）（2026-09-21）
#
#  ★★なぜ★★
#    85（`xlsb-styles-no-ji.md`）で ★どの 番号が 何か★ までは 割れました。
#    ⇒でも `43`（字・35バイト）と `46`（線・51バイト）の ★中の どこに 何が 在るか★ は
#      ★1つも 測って いません★。
#    ⇒`.xlsb` の 飾りが 画面に 出ない 残り 5つは ここです。
#
#  ★★やり方★★ ... 85と 同じ ★1つだけ 変える★
#    字  ... 1マスに つき ★1つだけ★ 変える（太字／斜体／下線／大きさ／色／字の 名）
#    線  ... 1マスに つき ★1辺だけ★（上／下／左／右）＋ ★太さ★ ＋ ★色★ ＋ ★形★
#    ⇒★全部 別の マス★＝同じ 飾りに まとめられて 1本しか 増えない のを 避けます
#
#  ★★作る 物★★（★名は 2本 決め打ち★）
#    %TEMP%\exally-ji-1tsu.xlsb ／ %TEMP%\exally-sen-1tsu.xlsb
#
#  ★門★
#    ①貝殻が 5.1（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③★打った 飾りを 読み返して 効いて いるか★（exit 5）
#    ④名が 決め打ちか（exit 7）／⑤包みの 頭が PK（exit 9）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$許す名 = 'exally-ji-1tsu.xlsb', 'exally-sen-1tsu.xlsb'

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null
$出来 = @()
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false

  # ══ ★★字★★ ══（★1マスに つき 1つだけ★）
  $名 = 'exally-ji-1tsu.xlsb'
  if ($許す名 -notcontains $名) { exit 7 }
  $先 = Join-Path $env:TEMP $名
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)
  for ($r = 1; $r -le 7; $r++) { $sh.Cells.Item($r, 1).Value2 = $r }
  $sh.Range('A1').Font.Bold = $true            # ★太字だけ★
  $sh.Range('A2').Font.Italic = $true          # ★斜体だけ★
  $sh.Range('A3').Font.Underline = 2           # ★下線だけ★（xlUnderlineStyleSingle）
  $sh.Range('A4').Font.Size = 20               # ★大きさだけ★
  $sh.Range('A5').Font.Color = 255             # ★色だけ★（COM は BGR ＝ 赤）
  $sh.Range('A6').Font.Name = 'Arial'          # ★字の 名だけ★
  $sh.Range('A7').Font.Strikethrough = $true   # ★取り消し線だけ★
  $確 = @(
    ('太字=' + [string]$sh.Range('A1').Font.Bold),
    ('斜体=' + [string]$sh.Range('A2').Font.Italic),
    ('下線=' + [string]$sh.Range('A3').Font.Underline),
    ('大きさ=' + [string]$sh.Range('A4').Font.Size),
    ('色=' + [string]$sh.Range('A5').Font.Color),
    ('名=' + [string]$sh.Range('A6').Font.Name),
    ('取消=' + [string]$sh.Range('A7').Font.Strikethrough)
  )
  Write-Host ('  字 ... ' + ($確 -join ' / '))
  if ($確.Count -ne 7) { Write-Host '★★確かめが 7本 在りません★★'; exit 5 }
  foreach ($c in $確) { if ($c -match '=False$|=$') { Write-Host ('★★効いて いません★★ ' + $c); exit 5 } }
  if (Test-Path $先) { Remove-Item $先 -Force }
  $bk.SaveAs($先, 50)
  $bk.Close($false); $sh = $null; $bk = $null
  $出来 += $先

  # ══ ★★線★★ ══（★1マスに つき 1辺／1つだけ★）
  $名 = 'exally-sen-1tsu.xlsb'
  if ($許す名 -notcontains $名) { exit 7 }
  $先 = Join-Path $env:TEMP $名
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)
  for ($r = 1; $r -le 7; $r++) { $sh.Cells.Item($r, 1).Value2 = $r }
  # ★7=左 / 8=上 / 9=下 / 10=右★
  $sh.Range('A1').Borders.Item(8).LineStyle = 1                      # ★上だけ★
  $sh.Range('A2').Borders.Item(9).LineStyle = 1                      # ★下だけ★
  $sh.Range('A3').Borders.Item(7).LineStyle = 1                      # ★左だけ★
  $sh.Range('A4').Borders.Item(10).LineStyle = 1                     # ★右だけ★
  $sh.Range('A5').Borders.Item(9).LineStyle = 1
  $sh.Range('A5').Borders.Item(9).Weight = 4                         # ★下＋太さ★
  $sh.Range('A6').Borders.Item(9).LineStyle = 1
  $sh.Range('A6').Borders.Item(9).Color = 255                        # ★下＋色（赤）★
  $sh.Range('A7').Borders.Item(9).LineStyle = -4115                  # ★下＋形（破線）★
  $確2 = @(
    ('上=' + [string]$sh.Range('A1').Borders.Item(8).LineStyle),
    ('下=' + [string]$sh.Range('A2').Borders.Item(9).LineStyle),
    ('左=' + [string]$sh.Range('A3').Borders.Item(7).LineStyle),
    ('右=' + [string]$sh.Range('A4').Borders.Item(10).LineStyle),
    ('太さ=' + [string]$sh.Range('A5').Borders.Item(9).Weight),
    ('色=' + [string]$sh.Range('A6').Borders.Item(9).Color),
    ('形=' + [string]$sh.Range('A7').Borders.Item(9).LineStyle)
  )
  Write-Host ('  線 ... ' + ($確2 -join ' / '))
  if ($確2.Count -ne 7) { Write-Host '★★確かめが 7本 在りません★★'; exit 5 }
  foreach ($c in $確2) { if ($c -match '=0$|=$') { Write-Host ('★★効いて いません★★ ' + $c); exit 5 } }
  if (Test-Path $先) { Remove-Item $先 -Force }
  $bk.SaveAs($先, 50)
  $bk.Close($false); $sh = $null; $bk = $null
  $出来 += $先
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

Write-Host ''
foreach ($f in $出来) {
  $fs = [IO.File]::OpenRead($f)
  $b = New-Object byte[] 2
  [void]$fs.Read($b, 0, 2)
  $fs.Close()
  if (([char]$b[0] + [string][char]$b[1]) -ne 'PK') { Write-Host '★★包みの 頭が PK では ありません★★'; exit 9 }
  $x = Get-Item $f
  Write-Host ('★出来ました★ ' + (Split-Path $f -Leaf).PadRight(24) + ' ' + $x.Length + ' バイト ／ sha256 ' + (Get-FileHash $f -Algorithm SHA256).Hash.ToLower())
}
Write-Host ('★本数★ ' + $出来.Count + '本（決め打ち 2本）')
if ($出来.Count -ne 2) { exit 6 }
