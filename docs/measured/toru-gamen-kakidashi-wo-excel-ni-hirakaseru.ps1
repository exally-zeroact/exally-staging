# toru-gamen-kakidashi-wo-excel-ni-hirakaseru.ps1
#   -- ★お客さんの 道で 書き出した .xlsx を 実Excel に 開かせる★（㊹）（2026-09-20）
#
#  ★★なぜ 要るか★★
#    ㊷㊸で 測ったのは ★道具で 作った 物★でした（Exally1 が 材料を 手で 組んだ）
#    ⇒★画面で 打った 時に 本当に 同じ 物が 出るか★は ★未測定★でした
#    ⇒★しかも 縦3マス 1本だけ★＝★横・2次元は 1本も 測って いません★
#
#  ★★この 道具が 測る 4つ★★
#    ㋐縦3     D1  `=SEQUENCE(3)`   ⇒ D1:D3
#    ㋑横3     F1  `=SEQUENCE(1,3)` ⇒ F1:H1
#    ㋒2次元   A20 `=SEQUENCE(2,3)` ⇒ A20:C21
#    ㋓溢れない A10 `=SUM(1,2)`     ⇒ ★対照★（`cm=` が 付いて いない 事の 裏取り）
#
#  ★★開く 門（★ここが 大事★）★★
#    ★この 道具は ★1本の ファイルしか 開けません★★
#      `%TEMP%\exally-kakidashi-gamen.xlsx`（★画面の 道で 書き出した 物★）
#    ・★名前が 1文字でも 違えば 走りません（exit 7）★
#    ・★`.xlsb` や 司さんの 実物の 名は 1文字も 在りません★
#    ・★保存しません★（`$bk.Close($false)`／`SaveAs` は 1文字も 在りません）
#
#  ★他の 門★
#    ①走らせる 前の Excel が 0個（exit 3）／②貝殻が powershell.exe（5.1）（exit 8）
#    ③開く ファイルが 無ければ 走らない（exit 6）／④台本の 数 決め打ち（exit 4）
#
#  ★★「変わったか」の 判じ方★★（★2回 嘘を 吐いた 後の 形★）
#    ・★投げなかった＝書けた、では ありません★（COM は 黙って 何も しない）
#    ・★1マスだけ 見ては いけません★（同じ 字を 書くと 見分けが 付かない）
#    ⇒★★見る 範囲 ★全部★ の 字と 値を 繋げて 比べます★★
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具> [-書ける]

param([switch]$書ける)

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = if ($書ける) { Join-Path $ここ 'golden-gamen-kakidashi-excel-kakeru-2026-09-20.tsv' }
      else { Join-Path $ここ 'golden-gamen-kakidashi-excel-2026-09-20.tsv' }

# == ★開いて よい ただ 1本★ ==
$許す名 = 'exally-kakidashi-gamen.xlsx'
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

# == ★台本★ ==
#   ★見る 範囲は 溢れ先より 1マス 広く 取ります★（★はみ出して いないかを 見る 為★）
$台本 = @(
  @{ 番='1'; 名='tate3';     頭='D1';  行1=1;  行2=5;  列1=4; 列2=5; 新='=SEQUENCE(2)' },
  @{ 番='2'; 名='yoko3';     頭='F1';  行1=1;  行2=2;  列1=6; 列2=9; 新='=SEQUENCE(1,2)' },
  @{ 番='3'; 名='2jigen';    頭='A20'; 行1=20; 行2=22; 列1=1; 列2=4; 新='=SEQUENCE(1,2)' },
  @{ 番='4'; 名='koborenai'; 頭='A10'; 行1=10; 行2=10; 列1=1; 列2=2; 新='=SUM(1,3)' }
)
$台本の数 = 4
if ($台本.Count -ne $台本の数) { exit 4 }

$あ = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
function 名前 { param($r, $c) ; return ([string]$あ[$c - 1] + [string]$r) }

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null; $c = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $読むだけ = -not $書ける
  Write-Host ('★開き方 ... ReadOnly=' + $読むだけ + '★（★保存は しません★）')
  $bk = $xl.Workbooks.Open($開く, 0, $読むだけ)
  $sh = $bk.Sheets.Item(1)

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★お客さんの 道で 書き出した .xlsx を 実Excel に 開かせた★（㊹）（2026-09-20）')
  $行.Add('# ★開いた 物★ ... ' + $開く + '（' + (Get-Item $開く).Length + ' バイト）')
  $行.Add('# ★開き方★ ... ReadOnly=' + $読むだけ + '（★保存して いません★）')
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★どの 貝殻か★ ... PowerShell ' + $版.ToString())
  $行.Add('# ★台本 ' + $台本.Count + '本★（縦3／横3／2次元／溢れない＝対照）')

  # ★★①開いた 直後★★
  $行.Add('#')
  $行.Add('# ★★①開いた 直後★★')
  $行.Add('# 番' + "`t" + '名' + "`t" + 'マス' + "`t" + '値' + "`t" + '出る字' + "`t" + '式' + "`t" + '型' + "`t" + '溢れの一部か')
  $控え = @{}
  foreach ($x in $台本) {
    $束 = New-Object System.Collections.Generic.List[string]
    for ($r = $x.行1; $r -le $x.行2; $r++) {
      for ($cc = $x.列1; $cc -le $x.列2; $cc++) {
        $c = $sh.Cells.Item($r, $cc)
        $v = $c.Value2
        $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
        $型 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
        $出字 = [string]$c.Text
        $式2 = [string]$c.Formula
        $一部か = '(?)'
        try { $一部か = [string]$c.HasArray } catch { }
        $ma = 名前 $r $cc
        $束.Add($ma + '=' + $値 + ':' + $式2)
        $行.Add($x.番 + "`t" + $x.名 + "`t" + $ma + "`t" + $値 + "`t" + $出字 + "`t" + $式2 + "`t" + $型 + "`t" + $一部か)
      }
    }
    $控え[$x.番] = ($束 -join '|')
  }

  # ★★②頭を 小さい 式に 書き換える★★（★生きて いるか★）
  $行.Add('#')
  $行.Add('# ★★②頭を 小さい 式に 書き換えた 後★★（★保存して いません★）')
  $行.Add('# 番' + "`t" + '名' + "`t" + 'マス' + "`t" + '値' + "`t" + '出る字' + "`t" + '式' + "`t" + '型' + "`t" + '溢れの一部か')
  foreach ($x in $台本) {
    $判じ = '(★書き換えられません★)'
    try {
      $sh.Range($x.頭).Formula2 = $x.新
      # ★★見る 範囲 全部の 字と 値を 繋げて 比べます★★
      $束 = New-Object System.Collections.Generic.List[string]
      for ($r = $x.行1; $r -le $x.行2; $r++) {
        for ($cc = $x.列1; $cc -le $x.列2; $cc++) {
          $c = $sh.Cells.Item($r, $cc)
          $v = $c.Value2
          $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
          $束.Add((名前 $r $cc) + '=' + $値 + ':' + [string]$c.Formula)
        }
      }
      $後 = ($束 -join '|')
      if ($後 -eq $控え[$x.番]) { $判じ = '★投げないが 1マスも 変わりません★' }
      else { $判じ = 'ok（★変わりました★）' }
    } catch { $判じ = '★投げました★ ' + $_.Exception.Message }
    $行.Add('# ★' + $x.名 + ' ... ' + $x.頭 + ' を ' + $x.新 + ' に★ ⇒ ' + $判じ)
    for ($r = $x.行1; $r -le $x.行2; $r++) {
      for ($cc = $x.列1; $cc -le $x.列2; $cc++) {
        $c = $sh.Cells.Item($r, $cc)
        $v = $c.Value2
        $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
        $型 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
        $出字 = [string]$c.Text
        $式2 = [string]$c.Formula
        $一部か = '(?)'
        try { $一部か = [string]$c.HasArray } catch { }
        $行.Add($x.番 + "`t" + $x.名 + "`t" + (名前 $r $cc) + "`t" + $値 + "`t" + $出字 + "`t" + $式2 + "`t" + $型 + "`t" + $一部か)
      }
    }
  }

  $窓行 = 40
    # ═══ ★★2つ目の 窓★★（★`=(マス)=0` の 真偽と 型を 一緒に 取る★）═══
    #   ★なぜ★ ... `.Value2` の 「0」は ★本物の 0★ とも ★空★ とも ★誤りの 番号★ とも
    #              区別が 付きません（記憶「意味の 無い 数は 一番 見つけにくい」）
    #   ⇒★別の 口（式）で もう 一度 0 かを 訊いて 型と 並べます★
    #   ★この 窓は ★上の 読みが 済んだ 後★に 打ちます★
    #     ＝式を 打つと 計算し直しが 起きる ので ★「開いた 瞬間」を 汚さない★
    $行.Add('#')
    $行.Add('# ★★2つ目の 窓★★（★上の 読みの 後に 打って います★）')
    $行.Add('# マス' + "`t" + '値' + "`t" + '型' + "`t" + '=(マス)=0')
    foreach ($ma2 in @('D1','D2','D3','F1','G1','H1','A20','B20','C20','A21','A10')) {
      $c = $sh.Range($ma2)
      $v = $c.Value2
      $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
      $型 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
      $ゼロか = '(★窓2が 打てません★)'
      # ★他の 道具と ★同じ 書き方★（`'=(' + マス + ')=0'`）★
      try { $w = $sh.Range('N' + $窓行); $w.Formula2 = '=(' + $ma2 + ')=0'; $ゼロか = [string]$w.Value2 } catch { }
      $窓行 = $窓行 + 1
      $行.Add($ma2 + "`t" + $値 + "`t" + $型 + "`t" + $ゼロか)
    }

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ('★書いた ... ' + $出 + '★')
  # ★★保存しません★★
  $bk.Close($false)
} finally {
  $c = $null; $sh = $null; $bk = $null
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
