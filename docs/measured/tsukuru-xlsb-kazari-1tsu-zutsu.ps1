# tsukuru-xlsb-kazari-1tsu-zutsu.ps1
#   -- ★飾りを 1つだけ 変えた `.xlsb` を 何本も 作る★（84）（2026-09-21）
#
#  ★★なぜ★★
#    `xl/styles.bin` の 記録の 番号（43 / 45 / 46 / 47 / 44 ...）に
#    ★名前を 付ける 根拠★が 要ります。
#    ⇒★仕様書の 名前（`BrtFont` 等）を 写すのは 「そう 呼ばれている」だけ★
#      ＝★実物の 証しでは ありません★
#    ⇒★飾りを 1つだけ 増やして ★どの 番号が 増えるか★ を 見ます★
#
#  ★★掃く 窓の 隙間を 先に 計算しました★★
#    ・何も 飾らない 物 ... 元の 数
#    ・太字だけ 2マス（★字だけ 変える★）      ⇒ 増えるなら ★字の 記録★
#    ・塗りだけ 2マス（★塗りだけ 変える★）    ⇒ 増えるなら ★塗りの 記録★
#    ・罫線だけ 2マス（★線だけ 変える★）      ⇒ 増えるなら ★線の 記録★
#    ・表示形式だけ 2マス（★書式だけ 変える★）⇒ 増えるなら ★書式の 記録★
#    ⇒★1本ずつ ★別の 色を 使う★＝同じ 飾りに まとめられて 増えないのを 避けます★
#      （例 太字は 2マスとも 太字だと 記録は 1本しか 増えない）
#      ⇒★だから 2マスは ★別の 濃さ／別の 色★に します★
#
#  ★★作る 物★★（★名は 5本 決め打ち★）
#    %TEMP%\exally-1tsu-nashi.xlsb / -ji.xlsb / -nuri.xlsb / -sen.xlsb / -kaki.xlsb
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③★書いた 後に 読み返して 飾りが 付いて いるか★（exit 5）
#    ④★名が 5本の 決め打ちの 中か★（exit 7）／⑤★包みの 頭が PK★（exit 9）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$許す名 = 'exally-1tsu-nashi.xlsb', 'exally-1tsu-ji.xlsb', 'exally-1tsu-nuri.xlsb',
          'exally-1tsu-sen.xlsb', 'exally-1tsu-kaki.xlsb'

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

  foreach ($種 in 'nashi', 'ji', 'nuri', 'sen', 'kaki') {
    $名 = 'exally-1tsu-' + $種 + '.xlsb'
    if ($許す名 -notcontains $名) { Write-Host '★★書いて よい 名では ありません★★'; exit 7 }
    $先 = Join-Path $env:TEMP $名

    $bk = $xl.Workbooks.Add()
    $sh = $bk.Sheets.Item(1)
    $sh.Range('A1').Value2 = 1
    $sh.Range('A2').Value2 = 2

    # ★★1つだけ 変えます★★（★2マスは ★別の 値★＝まとめられない ように★）
    $確かめ = ''
    switch ($種) {
      'nashi' { $確かめ = '(何も しない)' }
      'ji' {
        $sh.Range('A1').Font.Bold = $true
        $sh.Range('A2').Font.Italic = $true
        $確かめ = '太字=' + [string]$sh.Range('A1').Font.Bold + ' 斜体=' + [string]$sh.Range('A2').Font.Italic
      }
      'nuri' {
        # ★★色は ★3つ★ 要ります★★（2026-09-21 に 踏みました）
        #   赤と 緑の 2つだけだと ★違った バイトが 2つ★ 出るだけで
        #   「R G B が どの 順に 並ぶか」が ★決まりません★
        #   ⇒★3つ目（青）で 割ります★（★窓の 隙間を 先に 計算する★）
        #   ★COM の `Interior.Color` は BGR です★ ... 255=赤 / 65280=緑 / 16711680=青
        $sh.Range('A1').Interior.Color = 255
        $sh.Range('A2').Interior.Color = 65280
        $sh.Range('A3').Value2 = 3
        $sh.Range('A3').Interior.Color = 16711680
        $確かめ = '塗1=' + [string]$sh.Range('A1').Interior.Color + ' 塗2=' + [string]$sh.Range('A2').Interior.Color + ' 塗3=' + [string]$sh.Range('A3').Interior.Color
      }
      'sen' {
        $sh.Range('A1').Borders.Item(9).LineStyle = 1   # 下（xlEdgeBottom）
        $sh.Range('A2').Borders.Item(7).LineStyle = 1   # 左（xlEdgeLeft）
        $確かめ = '線1=' + [string]$sh.Range('A1').Borders.Item(9).LineStyle + ' 線2=' + [string]$sh.Range('A2').Borders.Item(7).LineStyle
      }
      'kaki' {
        $sh.Range('A1').NumberFormat = '0.000'
        $sh.Range('A2').NumberFormat = '0.00000'
        $確かめ = '書式1=' + [string]$sh.Range('A1').NumberFormat + ' 書式2=' + [string]$sh.Range('A2').NumberFormat
      }
    }
    Write-Host ('  ' + $種.PadRight(6) + ' ... ' + $確かめ)
    if ($種 -ne 'nashi' -and $確かめ -match 'False|^線1=0|^$') {
      Write-Host '★★飾りが 付いて いません★★'; exit 5
    }

    if (Test-Path $先) { Remove-Item $先 -Force }
    $bk.SaveAs($先, 50)     # ★50 ＝ xlExcel12（.xlsb）★
    $bk.Close($false)
    $sh = $null
    $bk = $null
    $出来 += $先
  }
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
  $x = Get-Item $f
  $fs = [IO.File]::OpenRead($f)
  $b = New-Object byte[] 2
  [void]$fs.Read($b, 0, 2)
  $fs.Close()
  $頭 = [char]$b[0] + [string][char]$b[1]
  if ($頭 -ne 'PK') { Write-Host '★★包みの 頭が PK では ありません★★'; exit 9 }
  if ([IO.Path]::GetExtension($f) -ne '.xlsb') { Write-Host '★★.xlsb では ありません★★'; exit 7 }
  Write-Host ('★出来ました★ ' + (Split-Path $f -Leaf).PadRight(24) + ' ' + $x.Length + ' バイト ／ sha256 ' + (Get-FileHash $f -Algorithm SHA256).Hash.ToLower())
}
Write-Host ('★本数★ ' + $出来.Count + '本（決め打ち 5本）')
if ($出来.Count -ne 5) { exit 6 }
