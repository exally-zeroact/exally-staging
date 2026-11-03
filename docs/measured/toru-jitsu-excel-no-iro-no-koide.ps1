# toru-jitsu-excel-no-iro-no-koide.ps1
#   -- ★濃さ（tint / shade）で 色が どう 変わるかを 実Excel に 言わせる★（92）（2026-09-21）
#
#  ★★なぜ★★
#    Exally1 の お願い 2つ
#      ⑴`<color theme="4" tint="-0.5"/>` の ★実際の 色★ が 分からない
#         ⇒今は ★元の 色を そのまま 出して いる★（色味は 合う／★明るさが ずれる★）
#      ⑵判子の 線の 色 `<a:lnRef><a:schemeClr val="accent1"><a:shade val="15000"/>`
#         ⇒★濃さの 計算を 測って いない★ので ★付けて いない★
#    ⇒★どちらも 「当て推量で 付けない」と 言って 止まって います＝正しい★
#    ⇒★だから 私が 実Excel に 聞きます★
#
#  ★★測り方★★
#    ①マスの 塗りに ★テーマの accent1★ を 当て、`TintAndShade` を 1つずつ 変えて
#      ★Excel が 返す 実際の 色★（`Interior.Color`）を 読む
#      ⇒★私が 計算しません★＝Excel の 答えだけ 並べます
#    ②判子を 1つ 置いて ★線と 塗りの 色★を そのまま 読む
#      （`ObjectThemeColor` ／ `TintAndShade` ／ `RGB`）
#    ③`.xlsx` と `.xlsb` の 両方で 保存し、★包みに 何と 書かれるか★は 別の 道具で 読む
#
#  ★★先に 書いて おく 見立て（★外れたら そう 書きます★）★★
#    ・`TintAndShade` が ★正★ ... 白へ 寄る（明るく）
#    ・`TintAndShade` が ★負★ ... 黒へ 寄る（暗く）
#    ・`shade val="15000"` は ★15%★ の 意味で、★0.15 を 掛ける★のでは ないか
#      ⇒★掛け算なら R21 G96 B130 は ほぼ 真っ黒に なる★ので、
#        ★実際の 線の 色が 真っ黒で なければ この 見立ては 外れ★です
#
#  ★★作る 物★★ `%TEMP%\exally-iro-no-koide.xlsx` と `.xlsb`（★名は 2本 決め打ち★）
#
#  ★門★
#    ①貝殻が 5.1（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③★濃さの 数だけ 行が 出たか★（exit 4）／④★判子が 1つ 置けたか★（exit 5）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$出1 = Join-Path $env:TEMP 'exally-iro-no-koide.xlsx'
$出2 = Join-Path $env:TEMP 'exally-iro-no-koide.xlsb'
if ((Split-Path $出1 -Leaf) -ne 'exally-iro-no-koide.xlsx') { exit 7 }
if ((Split-Path $出2 -Leaf) -ne 'exally-iro-no-koide.xlsb') { exit 7 }

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

# ★COM は BGR★ ⇒ ★RGB の 字に 直す 式も 紙に 残します★
function 字にする($c) {
  $b = [int]([math]::Floor($c / 65536)) -band 255
  $g = [int]([math]::Floor($c / 256)) -band 255
  $r = [int]$c -band 255
  return ('#{0:x2}{1:x2}{2:x2}' -f $r, $g, $b)
}

# ★★`0` を ★`0.0`★ と 書きます★★＝並びに Int32 が 混ざると その 行で 落ちます
$濃さたち = -0.9, -0.75, -0.5, -0.25, 0.0, 0.25, 0.5, 0.75, 0.9
$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null; $判 = $null
$出た = 0
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)

  Write-Host ''
  Write-Host '★★①マスの 塗り＝テーマの accent1（COM の 番号 5）に 濃さを 掛ける★★'
  Write-Host '  ★Excel が 返した 色だけ 並べます（私は 1つも 計算して いません）★'
  $行 = 1
  foreach ($t in $濃さたち) {
    $ma = 'A' + $行
    # ★★2026-09-21 ── ★`Value2` に ★整数★ を 渡すと 落ちます★★
    #   `指定された キャストは 有効では ありません`（今日 2回目）
    #   ＝`0` は Int32、`-0.5` は Double。★同じ 並びに 混ざると そこで 止まります★
    #   ⇒★`[double]` と 名指しで 書きます★（型を 並びに 決めさせない）
    $sh.Range($ma).Value2 = [double]$t
    $sh.Range($ma).Interior.ThemeColor = 5          # ★accent1★（実測で 5＝#156082）
    $sh.Range($ma).Interior.TintAndShade = [double]$t   # ★ここも 整数で 落ちます★
    $戻 = $sh.Range($ma).Interior.TintAndShade
    $色 = $sh.Range($ma).Interior.Color
    Write-Host ('  濃さ ' + ([string]$t).PadLeft(6) + ' ⇒ 読み返し ' + ([string]$戻).PadLeft(6) +
                ' ／ Interior.Color ' + ([string]$色).PadLeft(10) + ' ⇒ ★' + (字にする $色) + '★')
    $出た++
    $行 = $行 + 2
  }
  if ($出た -ne $濃さたち.Count) { Write-Host '★★濃さの 数だけ 出て いません★★'; exit 4 }

  Write-Host ''
  Write-Host '★★②判子（既定の 四角）の 線と 塗り★★'
  $判 = $sh.Shapes.AddShape(1, 300, 20, 60, 60)    # ★1 ＝ 四角★
  $判.Name = 'hanko-iro'
  if ([int]$sh.Shapes.Count -ne 1) { Write-Host '★★判子が 置けて いません★★'; exit 5 }
  $塗 = $判.Fill.ForeColor
  $線 = $判.Line.ForeColor
  Write-Host ('  塗り ... テーマの 番号 ' + ([string]$塗.ObjectThemeColor).PadLeft(3) +
              ' ／ 濃さ ' + ([string]$塗.TintAndShade).PadLeft(8) +
              ' ／ 明るさ ' + ([string]$塗.Brightness).PadLeft(8) +
              ' ／ RGB ' + ([string]$塗.RGB).PadLeft(10) + ' ⇒ ★' + (字にする $塗.RGB) + '★')
  Write-Host ('  線   ... テーマの 番号 ' + ([string]$線.ObjectThemeColor).PadLeft(3) +
              ' ／ 濃さ ' + ([string]$線.TintAndShade).PadLeft(8) +
              ' ／ 明るさ ' + ([string]$線.Brightness).PadLeft(8) +
              ' ／ RGB ' + ([string]$線.RGB).PadLeft(10) + ' ⇒ ★' + (字にする $線.RGB) + '★')
  Write-Host ('  線の 太さ ... ' + [string]$判.Line.Weight + ' ／ 線の 形 ' + [string]$判.Line.Style)

  Write-Host ''
  Write-Host '★★③テーマの accent1 そのもの★★'
  $a1 = $bk.Theme.ThemeColorScheme.Colors(5).RGB
  Write-Host ('  Colors(5) ... ' + ([string]$a1).PadLeft(10) + ' ⇒ ★' + (字にする $a1) + '★')

  foreach ($先 in @(@($出1, 51), @($出2, 50))) {
    if (Test-Path $先[0]) { Remove-Item $先[0] -Force }
    $bk.SaveAs($先[0], $先[1])
  }
  $bk.Close($false); $判 = $null; $sh = $null; $bk = $null
} finally {
  $判 = $null
  $sh = $null
  if ($null -ne $bk) { $bk.Close($false) }
  $bk = $null
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $xl = $null
  $t2 = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t2.Elapsed.TotalSeconds -lt 300)) {
    Start-Sleep -Milliseconds 250
  }
  $t2.Stop()
  $残り = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
  if ($残り -eq 0) { Write-Host ('★Excel は ' + [math]::Round($t2.Elapsed.TotalSeconds, 1) + '秒で ★消えました★★') }
  else { Write-Host ('★★Excel は 消えませんでした★★ ／ 残り ' + $残り + '個') }
}

Write-Host ''
foreach ($f in $出1, $出2) {
  if (-not (Test-Path $f)) { Write-Host ('★★出来て いません★★ ' + $f); exit 5 }
  Write-Host ('★作りました★ ' + (Split-Path $f -Leaf).PadRight(26) + (Get-Item $f).Length + ' バイト ／ sha256 ' + (Get-FileHash $f -Algorithm SHA256).Hash.ToLower())
}
Write-Host '⇒★包みに 何と 書かれて いるかは 別に 読みます★'
