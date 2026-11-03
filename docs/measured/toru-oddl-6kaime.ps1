# toru-oddl-6kaime.ps1 — ★ODDL の DSC の 元を 割る★（2026-09-18）
#
#  ★★なぜ この 形か（★前の 枠の 誤りから 作り直しました★）★★
#    5枠目 … DC／A／DSC の ★3つの 未知数★を ★3つの 利回り★から 解き、
#            ★その 3点で「差 0」と 書きました★
#    ⇒★★3つを 3点で 解けば どんな 数でも 差 0＝検めに なりません★★
#    ⇒今回は ★★未知を 1つずつ 直に 取り出します★★
#
#    ODDLPRICE ＝ (100 ＋ DC×券) ÷ (1 ＋ DSC×y/f) − A×券   （券 ＝ 100×利率÷f）
#      ㋐★利率 ＝ 0★ ⇒ 券 ＝ 0 ⇒ 価格 ＝ 100 ÷ (1 ＋ DSC×y/f)
#         ⇒★★DSC ＝ (100÷価格 − 1) × f ÷ y★★（★DSC だけが 出ます★）
#      ㋑★利回り ＝ 0★ ⇒ 価格 ＝ 100 ＋ 券×(DC − A)
#         ⇒★★DC − A ＝ (価格 − 100) ÷ 券★★
#      ㋒★決済 ＝ 最終利払日★（A ＝ 0）＋ 利回り 0 ⇒★DC が 直に★ ⇒ A ＝ DC −(DC−A)
#
#    ★★判じ方は ㋐と ㋑を 並べるだけ★★
#      DSC(㋐) ＝ DC−A(㋑) ⇒★決まりは DC − A★
#      DSC(㋐) ＜ DC−A(㋑) ⇒★決まりは NC − A★（はみ出しを 落として いる）
#
#  ★★何を 割るか★★（★見込みは 聞く 前に commit 済み … kansuu46/oddl-rokuwakume-no-an.md 4520cb0★）
#    組  f  はみ出し  NC   今 分かって いる 事
#    1   4  3日      3    ★NC − A（落とす）★   … 既に 2つの 決済で 確かめ済み
#    2   2  61日     1    ★DC − A（落とさない）★
#    3   2  1日      1    これから
#    4   4  48日     2    これから
#    5   1  28日     1    これから（★f=1 は まだ 見立て★）
#    6   2  1日      3    ★これが 決め手★（f=2 なのに NC が 大きい）
#      ⇒★NC が 元★なら 組6 は ★落とす★
#      ⇒★f が 元★なら 組6 は ★落とさない★
#      ⇒★はみ出しの 大きさが 元★なら 組3・組6 が 落とし 組4 が 落とさない
#
#  ★★★この 道具が 通した 門（8つ）★★★
#    ★なぜ 頭に 書くか★ … ★★門は 引き継がれません★★
#    ①★字の 誤り★ ……… ★0件★（PSParser::Tokenize で 数える）
#    ②★本数の 門★ ……… ★在り★（$式の本数 = 93／違えば exit 4）
#        ★わざと 1本 減らして 終わり値 4 を 確かめます★
#    ③★Excel の 門★ …… ★在り★（Get-Process と Win32_Process の ★2つで★）
#    ④★1本ずつ 受け止め★ … ★在り★（try/catch）
#        ★訳★ 2026-09-16 に CALL の 式 1本で ★307秒の 枠が 丸ごと 消えました★
#    ⑤★★司さんの 実物を 開く 字★★ … ★0件★
#        ＝司さんの ブックを ★1文字も 書きません★／★新しい 空の ブックだけ★
#    ⑥★BOM★ ………… ★在り★（efbbbf）
#    ⑦★2つ目の 窓＋型★ … ★在り★（.Value2 は 0 で ない 値に 0 を 返す）
#    ⑧★消えるまで 待って 秒数を 出す★ … ★在り★
#
#  ★★repo が 公開に なりました（2026-09-18）★★
#    ⇒★紙に 司さんの 商売の 中身を 書きません★
#    ⇒★この 紙の 日付・金額は 全部 作り物です★（教科書の 債券）
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-oddl-6kaime.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-oddl-6kaime-2026-09-18.tsv'

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel … Get-Process ' + $数1 + '個 ／ Win32_Process ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$式たち = New-Object System.Collections.Generic.List[object]
$足す = { param($訳, $式) $式たち.Add([pscustomobject]@{ 訳 = $訳; 式 = $式 }) }

# ══ ★6組★ ══ （★日付・金額は 全部 作り物の 債券です★）
$組 = @(
  @{ 名 = '組1(f4 はみ3日 NC3)';  f = 4; 最 = 'DATE(2008,11,30)'; 満 = 'DATE(2009,8,31)'; 決 = 'DATE(2009,3,10)' },
  @{ 名 = '組2(f2 はみ61日 NC1)'; f = 2; 最 = 'DATE(2007,10,15)'; 満 = 'DATE(2008,6,15)'; 決 = 'DATE(2008,2,7)' },
  @{ 名 = '組3(f2 はみ1日 NC1)';  f = 2; 最 = 'DATE(2008,11,30)'; 満 = 'DATE(2009,5,31)'; 決 = 'DATE(2009,1,15)' },
  @{ 名 = '組4(f4 はみ48日 NC2)'; f = 4; 最 = 'DATE(2008,11,30)'; 満 = 'DATE(2009,7,15)'; 決 = 'DATE(2009,1,15)' },
  @{ 名 = '組5(f1 はみ28日 NC1)'; f = 1; 最 = 'DATE(2009,1,31)';  満 = 'DATE(2010,2,28)'; 決 = 'DATE(2009,6,15)' },
  @{ 名 = '組6(f2 はみ1日 NC3)';  f = 2; 最 = 'DATE(2008,11,30)'; 満 = 'DATE(2010,5,31)'; 決 = 'DATE(2009,1,15)' }
)

foreach ($g in $組) {
  foreach ($b in 0, 1, 2, 3, 4) {
    # ㋐ 利率 0 ⇒ ★DSC が 直に★
    & $足す ($g.名 + ' (ア)DSC 利率0 basis=' + $b) ('=ODDLPRICE(' + $g.決 + ',' + $g.満 + ',' + $g.最 + ',0,0.05,100,' + $g.f + ',' + $b + ')')
    # ㋑ 利回り 0 ⇒ ★DC − A が 直に★
    & $足す ($g.名 + ' (イ)DC-A 利回り0 basis=' + $b) ('=ODDLPRICE(' + $g.決 + ',' + $g.満 + ',' + $g.最 + ',0.06,0,100,' + $g.f + ',' + $b + ')')
    # ㋒ 決済＝最終利払日（A=0）＋ 利回り 0 ⇒ ★DC が 直に★
    & $足す ($g.名 + ' (ウ)DC 決済=最終利払 basis=' + $b) ('=ODDLPRICE(' + $g.最 + ',' + $g.満 + ',' + $g.最 + ',0.06,0,100,' + $g.f + ',' + $b + ')')
  }
}

# ══ ★対照 3本★（★合わなければ そこで 止める★）══
& $足す '対照1(ODDL f4・紙に在る)' '=ODDLPRICE(DATE(2009,3,10),DATE(2009,8,31),DATE(2008,11,30),0.045,0.05,100,4,1)'
& $足す '対照2(ODDL f2・紙に在る)' '=ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,1)'
& $足す '対照3(ODDF・前の枠と同じ)' '=ODDFPRICE(DATE(2009,3,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2010,1,1),0.06,0.05,100,2,1)'

$式の本数 = 93
Write-Host ('★聞く 式 … ' + $式たち.Count + '本★（決め打ち ' + $式の本数 + '本）')
if ($式たち.Count -ne $式の本数) {
  Write-Host ('★★' + $式の本数 + '本の はずが ' + $式たち.Count + '本です★★／★変えたなら この 数も 直し、便りにも 同じ 数を 書いて ください★')
  exit 4
}

$xl = New-Object -ComObject Excel.Application
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★ODDL の DSC の 元を 割る★（2026-09-18・6回目）')
  $行.Add('#')
  $行.Add('# ★未知を 1つずつ 直に 取り出します★（★解いて 当てる を やめました★）')
  $行.Add('#   (ア)利率0   ⇒ DSC ＝ (100÷価格 − 1)×f÷y   （y ＝ 0.05）')
  $行.Add('#   (イ)利回り0 ⇒ DC − A ＝ (価格 − 100)÷券    （券 ＝ 100×0.06÷f）')
  $行.Add('#   (ウ)決済＝最終利払日 ＋ 利回り0 ⇒ DC ＝ (価格 − 100)÷券')
  $行.Add('#')
  $行.Add('# ★判じ方★ DSC(ア) ＝ DC−A(イ) なら ★DC − A★／小さければ ★NC − A★')
  $行.Add('#')
  $行.Add('# ★見込みは 聞く 前に commit 済み★ … kansuu46/oddl-rokuwakume-no-an.md（4520cb0）')
  $行.Add('#')
  $行.Add('# ★日付・金額は 全部 作り物の 債券です★（repo が 公開の 為）')
  $行.Add('#')
  $行.Add('# ★2つ目の 窓★ =(式)=0 … .Value2 は 0 で ない 値に 0 を 返す')
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
      $行.Add($x.訳 + "`t" + $x.式 + "`t" + '(★打てません★)' + "`t" + ('★Excel が 式を 受け付けません★ ' + $_.Exception.Message) + "`t" + '(★打てません★)' + "`t" + '(★打てません★)')
      $r++
    }
    if (-not $打てた) { continue }
    $v = $c.Value2
    $答 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    $型 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
    $字 = [string]$c.Text
    $w = $sh.Range('E' + $r)
    $ゼロか = '(★窓2が 打てません★)'
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
  Write-Host ('★Excel が 消えるまで ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒 ／ 残り ' + @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count + '個★')
}
