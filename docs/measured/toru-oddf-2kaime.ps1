# toru-oddf-2kaime.ps1 — ★ODDFPRICE の 切り分け 2回目★（2026-09-16）
#
#  ★★見込みは 聞く 前に 書いて あります★★
#    … `docs/measured/kansuu46/oddf-kiku-koto-2.md`
#    ⇒★後から 答えに 寄せられません★
#
#  ★1回目で 分かった 事★
#    ㋐★まるごとの 期間も basis で 割る★（見込み ㋐-2 が 当たり）
#    ㋑★準利払日を 越えた 所で 差が 約10倍に 跳ぶ★（見込み ㋑-2 が 当たり）
#    ⇒★★2つの 別々の 違いが 在ります★★
#       ①第1準期間の 中でも 残る 小さい 差（basis2 で 6.7e-3）
#       ②★期を またぐ と 出る 大きい 差★（basis2 で 6.5e-2）
#    ★端数3期・basis 3 は ★ぴたり 一致★★＝★合う 形が 在る★
#
#  ★★この 回で 聞く 事★★
#    ㋐★段の 中の 傾き★ … 第2準期間の 中で 決済を 1日ずつ
#    ㋑★段の 境目の 1日★ … 準利払日の 前日／当日／翌日
#    ㋒★ぴたり 合った 形（端数3期・basis 3）の 隣★
#    ㋓★ODDLPRICE の 対照を 作り直す★（1回目は ★私が 日付を 間違えました★）
#    ㋔★対照★（★紙に 在って ○ の 物★）
#
#  ★★決まり★★
#    ・走らせる ★その時に★ Excel を 2つの 道具で 数える（1個でも 居たら 走らせない）
#    ・★新しい 空の ブックだけ★（★司さんの 実物は 開きません★）
#    ・Visible=$false / DisplayAlerts=$false / finally で 必ず Quit
#    ・★消えるまで 待って 秒数を 出す★／★.ps1 は BOM 必須★
#    ・★2つ目の 窓（=(式)=0）と 型★も 取る
#    ・★1本ずつ 受け止める★（★1本の 事故で 枠を 丸ごと 落とさない★）
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-oddf-2kaime.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-oddf-2kaime-2026-09-16.tsv'

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel … Get-Process ' + $数1 + '個 ／ Win32_Process ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Error '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$式たち = New-Object System.Collections.Generic.List[object]
$足す = { param($訳, $式) $式たち.Add([pscustomobject]@{ 訳 = $訳; 式 = $式 }) }

# ══ ㋐★段の 中の 傾き★ ══
#   発行 2009-01-01 ／ 初回 2010-01-01 ／ 年2回 ⇒ 準利払日 2009-01-01・2009-07-01・2010-01-01
#   ★第2準期間（2009-07-01 → 2010-01-01）の 中で 決済を 1日ずつ 動かす★
foreach ($d in @('DATE(2009,7,2)', 'DATE(2009,7,3)', 'DATE(2009,7,4)', 'DATE(2009,7,5)')) {
  foreach ($b in @(1, 2, 3)) {
    & $足す "㋐段の中の傾き $d (basis=$b)" `
      "=ODDFPRICE($d,DATE(2013,1,1),DATE(2009,1,1),DATE(2010,1,1),0.06,0.05,100,2,$b)"
  }
}

# ══ ㋑★段の 境目の 1日★ ══（準利払日 2009-07-01 の 前日／当日／翌日）
foreach ($d in @('DATE(2009,6,30)', 'DATE(2009,7,1)', 'DATE(2009,7,2)')) {
  foreach ($b in @(0, 1, 2, 3, 4)) {
    & $足す "㋑境目の1日 $d (basis=$b)" `
      "=ODDFPRICE($d,DATE(2013,1,1),DATE(2009,1,1),DATE(2010,1,1),0.06,0.05,100,2,$b)"
  }
}

# ══ ㋒★ぴたり 合った 形（端数3期・basis 3）の 隣★ ══
#   ★合う 物だけ 増やしても 何も 分かりません★＝★合う 形と 合わない 形の 境目を 見ます★
#   発行 2008-07-01 ／ 初回 2010-01-01 ／ 年2回 ⇒ 端数 3期
foreach ($d in @('DATE(2009,6,30)', 'DATE(2009,7,1)', 'DATE(2009,7,2)', 'DATE(2009,10,1)')) {
  foreach ($b in @(2, 3)) {
    & $足す "㋒合った形の隣 $d (basis=$b)" `
      "=ODDFPRICE($d,DATE(2013,1,1),DATE(2008,7,1),DATE(2010,1,1),0.06,0.05,100,2,$b)"
  }
}

# ══ ㋓★ODDLPRICE の 対照を 作り直す★ ══
#   ★1回目は 最終利払日を 決済日より 後に して しまい #NUM! に しました（★私の 作り間違い★）
#   ★正しい 並び★ … 最終利払日 ＜ 決済日 ＜ 満期
foreach ($b in @(0, 1, 2, 3, 4)) {
  & $足す "㋓ODDLPRICE対照(作り直し basis=$b)" `
    "=ODDLPRICE(DATE(2009,3,10),DATE(2009,8,31),DATE(2008,11,30),0.045,0.05,100,4,$b)"
}

# ══ ㋔★対照（★紙に 在って ○ の 物★）★ ══
& $足す '㋔対照(紙に在る・○のはず)' `
  '=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,2)'
& $足す '㋔対照(紙に在る・○のはず)' `
  '=ODDFPRICE(DATE(2009,3,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2010,1,1),0.06,0.05,100,2,1)'
& $足す '㋔対照(前の枠と同じ・○のはず)' `
  '=ODDFPRICE(DATE(2009,7,1),DATE(2013,1,1),DATE(2008,7,1),DATE(2010,1,1),0.06,0.05,100,2,3)'

Write-Host ('★聞く 式 … ' + $式たち.Count + '本★')

$xl = New-Object -ComObject Excel.Application
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★ODDFPRICE の 切り分け 2回目★（2026-09-16）')
  $行.Add('#')
  $行.Add('# ★見込みは 聞く 前に 書いて あります★ … kansuu46/oddf-kiku-koto-2.md')
  $行.Add('#   ⇒★後から 答えに 寄せられません★')
  $行.Add('#')
  $行.Add('# ★1回目で 分かった 事★')
  $行.Add('#   ㋐まるごとの 期間も basis で 割る（見込み ㋐-2 が 当たり）')
  $行.Add('#   ㋑準利払日を 越えた 所で 差が 約10倍に 跳ぶ（見込み ㋑-2 が 当たり）')
  $行.Add('#   ★端数3期・basis 3 は ぴたり 一致★')
  $行.Add('#')
  $行.Add('# ★2つ目の 窓★ `=(式)=0` … `.Value2` は 0 で ない 値に 0 を 返す')
  $行.Add('#')
  $行.Add('# ★どの Excel か★ … 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('#')
  $行.Add('# 訳' + "`t" + '式' + "`t" + '答え' + "`t" + '出る字' + "`t" + '=(式)=0' + "`t" + '型')

  $r = 1
  foreach ($x in $式たち) {
    $c = $sh.Range('D' + $r)
    $打てた = $true
    try { $c.Formula = $x.式 } catch {
      $打てた = $false
      $行.Add($x.訳 + "`t" + $x.式 + "`t" + '(★打てません★)' + "`t" +
        ('★Excel が 式を 受け付けません★ ' + $_.Exception.Message) + "`t" + '(★打てません★)' + "`t" + '(★打てません★)')
      $r++
    }
    if (-not $打てた) { continue }
    $v = $c.Value2
    $答 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    $型 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
    $字 = [string]$c.Text
    $w = $sh.Range('E' + $r)
    $ゼロか = '(★窓②が 打てません★)'
    try { $w.Formula = '=(' + $x.式.Substring(1) + ')=0'; $ゼロか = [string]$w.Value2 } catch { }
    $行.Add($x.訳 + "`t" + $x.式 + "`t" + $答 + "`t" + $字 + "`t" + $ゼロか + "`t" + $型)
    $r++
  }

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★書いた … ' + $出 + '（' + $式たち.Count + '行）★')
  $bk.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $t = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 300)) {
    Start-Sleep -Milliseconds 500
  }
  $t.Stop()
  Write-Host ('★Excel が 消えるまで ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒 ／ 残り ' +
    @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count + '個★')
}
