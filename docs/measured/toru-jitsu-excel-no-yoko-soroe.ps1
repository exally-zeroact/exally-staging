# toru-jitsu-excel-no-yoko-soroe.ps1
#   -- ★実Excel は 何を どちらへ 寄せるか★（91）（2026-09-21）
#
#  ★★なぜ★★
#    Exally1「真偽（TRUE / FALSE）と 誤り（#VALUE! 等）の 横の 揃えを 測って ほしい。
#             ★私が 当て推量で 中央に しない 為★」
#    ⇒★当て推量で 直さずに 訊いて きたのは 正しい★ ので 測ります。
#
#  ★★測り方の 落とし穴（先に 書きます）★★
#    `HorizontalAlignment` は ★打った 設定★を 返します。
#    何も 設定して いない マスは ★`xlGeneral`（1）★＝「中身で 決める」。
#    ⇒★`xlGeneral` のままでは 「実際に どちらへ 出るか」は 分かりません★
#    ⇒★だから 「設定」と 「実際の 見た目」を 2つ 測ります★
#        ①`HorizontalAlignment`   ... 打った 設定（ほぼ 全部 1＝General）
#        ②★絵に して 字の 出る 場所を 見る★ ... ★これが 本当の 答え★
#      ここでは ①を 出し、②は `toru-jitsu-excel-no-e.ps1` で 絵を 撮ります。
#
#  ★★盤面★★（★1行 おき★＝隣の 飾りに 巻き込まれない）
#    A1  数          123
#    A3  字          'abc'
#    A5  真          =TRUE()
#    A7  偽          =FALSE()
#    A9  誤り        =1/0        （#DIV/0!）
#    A11 誤り(名)    =NOSUCHFN() （#NAME?）
#    A13 日付        =DATE(2024,1,1)
#    A15 空白        （何も 打たない）
#
#  ★門★
#    ①貝殻が 5.1（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③★8つ とも 読み返して 中身が 入って いるか★（exit 5）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$出す先 = Join-Path $env:TEMP 'exally-yoko-soroe.xlsx'
if ((Split-Path $出す先 -Leaf) -ne 'exally-yoko-soroe.xlsx') { exit 7 }

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

# ★XlHAlign の 数★（★名前は Excel の 物★）
$揃え札 = @{ 1 = 'General（中身で 決める）'; -4131 = 'Left'; -4108 = 'Center'; -4152 = 'Right';
             5 = 'Fill'; -4130 = 'Justify'; 7 = 'CenterAcrossSelection'; 8 = 'Distributed' }

$台本 = @(
  @{ 行 = 1;  札 = '数';        式 = $null },
  @{ 行 = 3;  札 = '字';        式 = $null },
  @{ 行 = 5;  札 = '真';        式 = '=TRUE()' },
  @{ 行 = 7;  札 = '偽';        式 = '=FALSE()' },
  @{ 行 = 9;  札 = '誤り(除0)'; 式 = '=1/0' },
  @{ 行 = 11; 札 = '誤り(名)';  式 = '=NOSUCHFN()' },
  @{ 行 = 13; 札 = '日付';      式 = '=DATE(2024,1,1)' },
  @{ 行 = 15; 札 = '空白';      式 = $null }
)
if ($台本.Count -ne 8) { Write-Host '★★台本が 8本 在りません★★'; exit 4 }

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)
  $sh.Columns.Item(1).ColumnWidth = 14

  foreach ($x in $台本) {
    $ma = 'A' + $x.行
    # ★★2026-09-21 ── ハッシュ表に 型の 違う 値を 混ぜて 入れたら
    #   `Value2 = $x.値` が ★指定された キャストは 有効では ありません★ で 落ちました★★
    #   ⇒★型を 名指しで 書きます★（表から 型を 当てさせない）
    if ($null -ne $x.式) { $sh.Range($ma).Formula2 = $x.式 }
    elseif ($x.札 -eq '数') { $sh.Range($ma).Value2 = [double]123 }
    elseif ($x.札 -eq '字') { $sh.Range($ma).Value2 = [string]'abc' }
  }

  Write-Host ''
  Write-Host '★★打った 設定（HorizontalAlignment）★★'
  Write-Host '  ★これは 「打った 設定」です＝General なら 実際の 寄りは 絵で 見ます★'
  $入った = 0
  foreach ($x in $台本) {
    $ma = 'A' + $x.行
    $c = $sh.Range($ma)
    $h = [int]$c.HorizontalAlignment
    $札 = if ($揃え札.ContainsKey($h)) { $揃え札[$h] } else { '(知らない 数 ' + $h + ')' }
    $字 = [string]$c.Text
    $v = $c.Value2
    $型 = if ($null -eq $v) { '空' } elseif ($v -is [bool]) { '真偽' } elseif ($v -is [double]) { '数' } else { $v.GetType().Name }
    if ($x.札 -eq '空白' -or $字 -ne '') { $入った++ }
    Write-Host ('  A' + $x.行.ToString().PadRight(3) + ' ' + $x.札.PadRight(11) +
                ' 字「' + $字.PadRight(12) + '」 型 ' + $型.PadRight(6) + ' 揃え ' + $h.ToString().PadLeft(6) + ' = ' + $札)
  }
  Write-Host ('★中身が 入った マス★ ' + $入った + ' / 8')
  if ($入った -ne 8) { Write-Host '★★入って いない マスが 在ります★★'; exit 5 }

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
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 180)) {
    Start-Sleep -Milliseconds 250
  }
  $t.Stop()
  $残り = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
  if ($残り -eq 0) { Write-Host ('★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒で ★消えました★★') }
  else { Write-Host ('★★Excel は 消えませんでした★★ ／ 残り ' + $残り + '個') }
}

if (-not (Test-Path $出す先)) { exit 5 }
Write-Host ''
Write-Host ('★★作りました★★ ... ' + $出す先)
Write-Host ('  ★sha256★ ' + (Get-FileHash $出す先 -Algorithm SHA256).Hash.ToLower())
Write-Host '  ⇒★次に この 本を `toru-jitsu-excel-no-e.ps1` で 絵に して ★実際の 寄り★を 見ます★'
