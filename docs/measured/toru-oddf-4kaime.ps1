# toru-oddf-4kaime.ps1 — ★ODDFPRICE … 部品を 直に 取り出す★（2026-09-17）
#
#  ★★考え★★
#    今までは ★答え（価格）から 中を 覗こうと して いました★。
#    ⇒価格には DFC・A・DSC・N が ★全部 混ざって います★
#    ⇒★混ざった 物から 逆に 解くのは 当てに なります★（3枠目の 後で 実際に 詰まりました）
#    ⇒★★引数を 0 に すると 混ざりが 解けます★★
#
#  ★㋐利率を 0★ ⇒ 券（利札）＝0 ⇒ 価格 ＝ 償還 ÷ (1+利回り/f)^((N−1)+DSC)
#     ⇒★DSC ＝ log(償還/価格)/log(1+利回り/f) − (N−1)★＝★1つの 式で 出ます★
#  ★㋑利回りを 0★ ⇒ 割引なし ⇒ 価格 ＝ 償還 ＋ (N−1)×券 ＋ 券×(DFC−A)
#     ⇒★DFC−A（＝DSC）が 出ます★＝★㋐と 突き合わせれば 測りの 確かめに なります★
#  ★㋒利率0 ＋ 利回り0★ ⇒ ★100 の はず★（★対照★・違えば 私の 読みが 違う）
#
#  ★見込みは 聞く 前に 書いて あります★
#    … `docs/measured/kansuu46/oddf-yonwakume-no-an.md`
#
#  ★★★この 道具が 通した 門（6つ）★★★
#    ★なぜ 頭に 書くか★ … ★★門は 引き継がれません★★
#      2026-09-17 … 前の 道具（toru-oddf-to-46ko.ps1）に 在った
#              ★本数の 門が この 道具から 消えて いました★
#      ＝★人は「前に 付けた」と 覚えて いる つもりで 次を 書きます★
#    ⇒★★次の 道具を 書く 人は ここを 写して ください★★
#
#    ①★字の 誤り★ ……… ★0件★（PSParser::Tokenize）
#    ②★本数の 門★ ……… ★在り★（下の $式の本数）
#        ★わざと 1本 減らして ★終わり値 1★ を 確かめます★
#        ★★数を 変えた 時にも 壊れます★★（足した 時だけ では ありません）
#    ③★Excel の 門★ …… ★在り★（Get-Process と Win32_Process の ┅2つで★）
#    ④★1本ずつ 受け止め★ … ★在り★（try/catch）
#        ★訳★ 2026-09-16 に CALL の 式 1本で ★307秒の 枛が 丸ごと 消えました★
#    ⑤★★司さんの 実物を 開く 字★★ … ★0件★
#        ★指示役1 の 一言★「★特に これを 毎回 数えて いるのが 大事
#                        ／★1回でも 抜けたら 取り返しが つきません★」
#    ⑥★BOM★ ………… ★在り★（efbbbf）
#
#    ★★門を 足したら その場で わざと 壊して 確かめる★★
#      2026-09-17 … 門を 2行で 書いて ★字の 誤り★に なって いました
#      ⇒★★足しただけでは 門は 効きません★★
#
#  ★★決まり★★
#    ・走らせる ★その時に★ Excel を 2つの 道具で 数える
#    ・★新しい 空の ブックだけ★（★司さんの 実物は 開きません★）
#    ・Visible=$false / DisplayAlerts=$false / finally で 必ず Quit
#    ・★消えるまで 待って 秒数を 出す★／★BOM 必須★／★2つ目の 窓＋型★
#    ・★1本ずつ 受け止める★
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-oddf-4kaime.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-oddf-4kaime-2026-09-17.tsv'

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel … Get-Process ' + $数1 + '個 ／ Win32_Process ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Error '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$式たち = New-Object System.Collections.Generic.List[object]
$足す = { param($訳, $式) $式たち.Add([pscustomobject]@{ 訳 = $訳; 式 = $式 }) }

# ★端数の 長さを 変える 3つの 形★（★合う 形と 合わない 形を 並べます★）
#   端数1期 … 発行 2009-01-01 ／ 初回 2009-07-01  ★NC=1＝今 合って います★
#   端数2期 … 発行 2009-01-01 ／ 初回 2010-01-01  ★合いません★
#   端数3期 … 発行 2008-07-01 ／ 初回 2010-01-01  ★basis 3 だけ 合います★
$形 = @(
  @{ 名 = '端数1期'; 発 = 'DATE(2009,1,1)'; 初 = 'DATE(2009,7,1)' },
  @{ 名 = '端数2期'; 発 = 'DATE(2009,1,1)'; 初 = 'DATE(2010,1,1)' },
  @{ 名 = '端数3期'; 発 = 'DATE(2008,7,1)'; 初 = 'DATE(2010,1,1)' }
)

foreach ($k in $形) {
  foreach ($b in @(0, 1, 2, 3, 4)) {
    # ㋐★利率 0★ ⇒ DSC が 出ます
    & $足す "㋐利率0 $($k.名) (basis=$b)" `
      "=ODDFPRICE(DATE(2009,3,1),DATE(2013,1,1),$($k.発),$($k.初),0,0.05,100,2,$b)"
    # ㋑★利回り 0★ ⇒ DFC−A が 出ます
    & $足す "㋑利回り0 $($k.名) (basis=$b)" `
      "=ODDFPRICE(DATE(2009,3,1),DATE(2013,1,1),$($k.発),$($k.初),0.06,0,100,2,$b)"
    # ㋒★両方 0★ ⇒ 100 の はず（対照）
    & $足す "㋒両方0 $($k.名) (basis=$b)" `
      "=ODDFPRICE(DATE(2009,3,1),DATE(2013,1,1),$($k.発),$($k.初),0,0,100,2,$b)"
  }
}

# ★㋓利率 0 で 決済を 動かす★（★DSC が 日ごとに どう 動くか★）
foreach ($d in @('DATE(2009,2,1)', 'DATE(2009,4,1)', 'DATE(2009,8,1)', 'DATE(2009,11,1)')) {
  foreach ($b in @(1, 2, 3)) {
    & $足す "㋓利率0で決済を動かす $d (basis=$b)" `
      "=ODDFPRICE($d,DATE(2013,1,1),DATE(2009,1,1),DATE(2010,1,1),0,0.05,100,2,$b)"
  }
}

# ══ ★★ODDLPRICE … ★片方ずつ 動かします★★★ ══
#   ★なぜ★ 前の 紙は ★f と 期越えが 同時に 変わって いました★
#         （合う 組 f=2・第1準期間 ／ 落ちる 組 f=4・第2準期間）
#   ⇒★★本数を 増やしても 解けません＝★組み方★の 話です★★
#   ⇒★★片方を 止めて もう 片方だけ 動かします★★
#   ★見込み★ f を 止めて 変わる ⇒ 期越えが 元
#          期越えを 止めて 変わる ⇒ f が 元
#          両方で 変わる ⇒ ★元が 2つ★（★これも 在り得ます★）
#   ★訳に f と 期越えを 全部 書きます★（★紙を 見る 人は コードを 見ません★）
foreach ($b in @(0, 1, 2, 3, 4)) {
  & $足す "ODDL f=4 第1期の中(期越えなし) (basis=$b)" `
    "=ODDLPRICE(DATE(2009,1,15),DATE(2009,8,31),DATE(2008,11,30),0.045,0.05,100,4,$b)"
  & $足す "ODDL f=4 第2期の中(期越えあり) (basis=$b)" `
    "=ODDLPRICE(DATE(2009,3,10),DATE(2009,8,31),DATE(2008,11,30),0.045,0.05,100,4,$b)"
  & $足す "ODDL f=2 第1期の中(期越えなし) (basis=$b)" `
    "=ODDLPRICE(DATE(2009,1,15),DATE(2009,11,30),DATE(2008,11,30),0.045,0.05,100,2,$b)"
  & $足す "ODDL f=2 第2期の中(期越えあり) (basis=$b)" `
    "=ODDLPRICE(DATE(2009,7,10),DATE(2009,11,30),DATE(2008,11,30),0.045,0.05,100,2,$b)"
}
# ★期越えを 止めて f だけ 動かす★（★決済は どれも 第1期の 中★）
foreach ($b in @(1, 2, 3)) {
  & $足す "ODDL 第1期固定 f=1 (basis=$b)" `
    "=ODDLPRICE(DATE(2009,1,15),DATE(2010,11,30),DATE(2008,11,30),0.045,0.05,100,1,$b)"
  & $足す "ODDL 第1期固定 f=2 (basis=$b)" `
    "=ODDLPRICE(DATE(2009,1,15),DATE(2009,11,30),DATE(2008,11,30),0.045,0.05,100,2,$b)"
  & $足す "ODDL 第1期固定 f=4 (basis=$b)" `
    "=ODDLPRICE(DATE(2009,1,15),DATE(2009,8,31),DATE(2008,11,30),0.045,0.05,100,4,$b)"
}
# ★ODDL も 利率 0 ／ 利回り 0★（★DSC と DC−A が 直に 出ます★）
foreach ($b in @(1, 2, 3)) {
  & $足す "ODDL 利率0 第2期の中 f=4 (basis=$b)" `
    "=ODDLPRICE(DATE(2009,3,10),DATE(2009,8,31),DATE(2008,11,30),0,0.05,100,4,$b)"
  & $足す "ODDL 利回り0 第2期の中 f=4 (basis=$b)" `
    "=ODDLPRICE(DATE(2009,3,10),DATE(2009,8,31),DATE(2008,11,30),0.045,0,100,4,$b)"
}
# ★ODDL の 対照★（★紙に 在って ○ の 物★）
& $足す 'ODDL 対照(紙に在る・○のはず)' `
  '=ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0575,0.0625,100,2,0)'

# ★㋔対照（★紙に 在って ○ の 物★・合わなければ そこで 止める）★
& $足す '㋔対照1(紙に在る・○のはず)' `
  '=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,2)'
& $足す '㋔対照2(紙に在る・○のはず)' `
  '=ODDFPRICE(DATE(2009,3,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2010,1,1),0.06,0.05,100,2,1)'
& $足す '㋔対照3(前の枠と同じ・○のはず)' `
  '=ODDFPRICE(DATE(2009,7,1),DATE(2013,1,1),DATE(2008,7,1),DATE(2010,1,1),0.06,0.05,100,2,3)'

# ★★本数の 門★★（2026-09-17 に 足しました）
#   ★なぜ★ 私は 便りに 「51本」と 書きましたが
#         ★コードを 数えると 60本★ でした（★私の 数え 間違い★）。
#   ⇒★経営者1 は その 数を 見て「300秒で 足りるか」を 判じます★
#   ⇒★指示役1 は その 数で 順番を 決めます★
#   ⇒★★数が 変わっても 気づかずに 走るのが 一番 悪い★★
#     ＝★しかも 300秒 使って から 気づきます★
#   ★同じ 門は `toru-oddf-to-46ko.ps1` に も 在ります（46個）
$式の本数 = 96
Write-Host ('★聞く 式 … ' + $式たち.Count + '本★（決め打ち ' + $式の本数 + '本）')
if ($式たち.Count -ne $式の本数) {
  # ★PowerShell は 括弧の 中で 行を 続けるのに ★逆引用符★ が 要ります
  #   ⇒★踏みましたので 1行に まとめます★
  Write-Error ('★★' + $式の本数 + '本の はずが ' + $式たち.Count + '本です★★／★変えたなら この 数も 直し、便りにも 同じ 数を 書いて ください★')
  exit 4
}

$xl = New-Object -ComObject Excel.Application
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★ODDFPRICE … 部品を 直に 取り出す★（2026-09-17・4回目）')
  $行.Add('#')
  $行.Add('# ★考え★ 価格には DFC・A・DSC・N が 全部 混ざって います')
  $行.Add('#   ⇒★引数を 0 に すると 混ざりが 解けます★')
  $行.Add('#   ㋐利率0  … 券=0  ⇒ 価格 = 償還 ÷ (1+利回り/f)^((N-1)+DSC)')
  $行.Add('#   ㋑利回り0 … 割引なし ⇒ 価格 = 償還 + (N-1)×券 + 券×(DFC-A)')
  $行.Add('#   ㋒両方0  … ★100 の はず★（対照）')
  $行.Add('#')
  $行.Add('# ★見込みは 聞く 前に 書いて あります★ … kansuu46/oddf-yonwakume-no-an.md')
  $行.Add('#')
  $行.Add('# ★2つ目の 窓★ `=(式)=0` … `.Value2` は 0 で ない 値に 0 を 返す')
  $行.Add('#')
  $行.Add('# ★どの Excel か★ … 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('#')
  $行.Add('# 訳' + "`t" + '式' + "`t" + '答え' + "`t" + '出る字' + "`t" + '=(式)=0' + "`t" + '型' + "`t" + '㆒の 印')

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
    # ★★㆒（両方 0）が 100 で ない 時は ★その場で 印を 付けます★★
    #   ★なぜ★ ★「実Excel が 100 で ない」のか
    #         「★私の 式の 形の 読みが 違う★」のかを 分ける 為★
    #   ⇒★後から 見て も 分かる ように 紙に 残します★
    $印 = ''
    if ($x.訳 -like '㆒*') {
      if ($型 -eq 'Double' -and [math]::Abs([double]$v - 100) -le 1e-9) { $印 = '★100 です★' }
      else { $印 = '★★100 で は ありません★★（★ここで 止めて 読みを 見直す★）' }
    }
    $行.Add($x.訳 + "`t" + $x.式 + "`t" + $答 + "`t" + $字 + "`t" + $ゼロか + "`t" + $型 + "`t" + $印)
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
