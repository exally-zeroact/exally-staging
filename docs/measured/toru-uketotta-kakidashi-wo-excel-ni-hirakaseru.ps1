# toru-uketotta-kakidashi-wo-excel-ni-hirakaseru.ps1
#   -- ★受け取った ファイルを 保存し直した 物を 実Excel に 開かせる★（㊼）（2026-09-20）
#
#  ★★なぜ 要るか★★
#    Exally1 が 測りました:
#      ★お客さんが 画面で `=SEQUENCE(3)` を `=SEQUENCE(2)` に 直して 保存すると
#        ★保存された 物は `=SEQUENCE(3)` の まま／溢れ先 C3 も 古い 3 の まま★★
#      ★窓（DiffPreview）にも 出ない＝断りも 知らせも 無い★
#    ⇒生の 字でも 私が 確かめました:
#      `<c r="C1" cm="1"><f t="array" ref="C1:C3">_xlfn.SEQUENCE(3)</f><v>1</v></c>`
#      `<c r="C3"><v>3</v></c>`
#    ⇒★でも 「お客さんが 実Excel で 開いた 時 何を 見るか」は まだ 誰も 見て いません★
#
#  ★★開く 物★★ `%TEMP%\exally-uketotta-kakidashi.xlsx`（★1本の 名★）
#    ・元は `exally-tameshi-hiraku.xlsx`（9,923B・実Excel が 作った 物）
#    ・お客さんの 道で 開き ★A1 を 3 ⇒ 99★／★C1 の SEQUENCE(3) ⇒ SEQUENCE(2)★ に 直して 保存
#    ・★司さんの 実物の 名は 1文字も 在りません★
#    ・★読むだけ★（`Close($false)`／`SaveAs` は 1文字も 在りません）
#
#  ★★この 材料で 測れない 事（★先に 書きます★）★★
#    ★判子・罫線・塗り・図形は ★元の 材料に 1つも 入って いません★★
#      （罫線の 型 1＝空の 既定だけ／塗りの 型 2＝既定だけ／絵 0本）
#    ⇒★だから 「判子が 残るか」は この 紙では 答えられません★
#    ⇒飾り入りの 材料を 別に 作りました（`tsukuru-tameshi-hiraku2-kazari-ari.ps1`）
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②名が 1文字でも 違えば 走らない（exit 7）
#    ③ファイルが 無ければ 走らない（exit 6）／④走らせる 前の Excel が 0個（exit 3）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-uketotta-kakidashi-excel-2026-09-20.tsv'

$許す名 = 'exally-uketotta-kakidashi.xlsx'
$開く = Join-Path $env:TEMP $許す名

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }
if ((Split-Path $開く -Leaf) -ne $許す名) { Write-Host '★★開いて よい ファイルは 1本だけです★★'; exit 7 }
if (-not (Test-Path $開く)) { Write-Host ('★★在りません ... ' + $開く + '★★'); exit 6 }
Write-Host ('★開く 物 ... ' + $開く + '（' + (Get-Item $開く).Length + ' バイト）★')

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$見るマス = @('A1', 'A2', 'A3', 'B1', 'C1', 'C2', 'C3', 'C4', 'D1', 'E1')

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null; $c = $null; $w = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Open($開く, 0, $true)
  $sh = $bk.Sheets.Item(1)

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★受け取った ファイルを 保存し直した 物を 実Excel に 開かせた★（㊼）（2026-09-20）')
  $行.Add('# ★開いた 物★ ... ' + $開く + '（' + (Get-Item $開く).Length + ' バイト）')
  $行.Add('# ★読むだけ★（保存して いません）')
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★どの 貝殻か★ ... PowerShell ' + $版.ToString())
  $行.Add('# ★お客さんが した 事★ A1 を 3 ⇒ 99 ／ C1 の =SEQUENCE(3) を ⇒ =SEQUENCE(2) に')
  $行.Add('# ★画面では★ C1=1 C2=2 ★C3 は 空★ に なって いました')
  $行.Add('# ★生の 字では★ C1 の 式は ★=SEQUENCE(3) の まま★／C3 に ★3 が 残って います★')
  $行.Add('# ★この 材料に 飾り（判子・罫線・塗り・図形）は 1つも 在りません＝そこは 測れません★')

  # ═══ ★★①開いた 瞬間★★ ═══
  $行.Add('#')
  $行.Add('# ★★①開いた 瞬間★★（★まだ 何も 打って いません★）')
  $行.Add('# マス' + "`t" + '値' + "`t" + '出る字' + "`t" + '式' + "`t" + '型' + "`t" + '溢れの一部か')
  foreach ($ma in $見るマス) {
    $c = $sh.Range($ma)
    $v = $c.Value2
    $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    $型 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
    $一部か = '(?)'
    try { $一部か = [string]$c.HasArray } catch { }
    $行.Add($ma + "`t" + $値 + "`t" + [string]$c.Text + "`t" + [string]$c.Formula + "`t" + $型 + "`t" + $一部か)
  }

  # ═══ ★★②飾りが 在るか（★この 材料には 元から 無い＝裏取り★）★★ ═══
  $行.Add('#')
  $行.Add('# ★★②飾り★★（★元の 材料に 無い＝「消えた」では なく「元から 無い」★）')
  $行.Add('# 見た物' + "`t" + '数')
  $行.Add('罫線(A1 左) LineStyle' + "`t" + [string]$sh.Range('A1').Borders.Item(7).LineStyle)
  $行.Add('塗り(B1) Color' + "`t" + [string]$sh.Range('B1').Interior.Color)
  $行.Add('字(A1) Bold' + "`t" + [string]$sh.Range('A1').Font.Bold)
  $行.Add('図形の数' + "`t" + [string]$sh.Shapes.Count)
  $行.Add('A列の幅' + "`t" + [string]$sh.Columns.Item(1).ColumnWidth)

  # ═══ ★★2つ目の 窓★★（★`=(マス)=0` の 真偽と 型を 一緒に 取る★）═══
  #   ★なぜ★ ... `.Value2` の 「0」は ★本物の 0★ とも ★空★ とも ★誤りの 番号★ とも
  #              区別が 付きません（記憶「意味の 無い 数は 一番 見つけにくい」）
  #   ★上の 読みが 済んだ 後に 打ちます★（★式を 打つと 計算し直しが 起きる★）
  $行.Add('#')
  $行.Add('# ★★2つ目の 窓★★（★上の 読みの 後に 打って います★）')
  $行.Add('# マス' + "`t" + '値' + "`t" + '型' + "`t" + '=(マス)=0')
  $窓行 = 40
  foreach ($ma2 in $見るマス) {
    $c = $sh.Range($ma2)
    $v = $c.Value2
    $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    $型 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
    $ゼロか = '(★窓2が 打てません★)'
    try { $w = $sh.Range('N' + $窓行); $w.Formula2 = '=(' + $ma2 + ')=0'; $ゼロか = [string]$w.Value2 } catch { }
    $窓行 = $窓行 + 1
    $行.Add($ma2 + "`t" + $値 + "`t" + $型 + "`t" + $ゼロか)
  }

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ('★書いた ... ' + $出 + '★')
  # ★★保存しません★★
  $bk.Close($false)
  $bk = $null
} finally {
  $c = $null; $w = $null; $sh = $null
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
