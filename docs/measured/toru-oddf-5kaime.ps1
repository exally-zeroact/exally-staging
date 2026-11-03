# toru-oddf-5kaime.ps1 — ★DFC と A を 別々に 取り出す★（2026-09-17）
#
#  ★★考え★★
#    価格 ＝ R/割^((N−1)+DSC) ＋ 券×DFC/割^DSC ＋ Σ 券/割^((k−1)+DSC) − 券×A
#    ★★DFC の 項は 割^DSC で 割られ、A の 項は 割られません★★
#    ⇒★利回り（＝割）を 変えると ★DFC の 項だけ 動きます★★
#    ⇒★★利回りを 3つ 使えば DFC と A が 分かれ、しかも 揃うかも 見られます★★
#      （★DSC は 4枠目の「利率 0」で 既に 出て います★）
#
#  ★見込みは 聞く 前に 書いて あります★
#    … `docs/measured/kansuu46/oddf-gowakume-no-an.md`
#
#  ★★★この 道具が 通した 門（6つ）★★★
#    ★なぜ 頭に 書くか★ … ★★門は 引き継がれません★★
#      2026-09-17 … 前の 道具に 在った ★本数の 門★が 次の 道具から 消えて いました
#      ＝★人は「前に 付けた」と 覚えて いる つもりで 次を 書きます★
#    ⇒★★次の 道具を 書く 人は ここを 写して ください★★
#
#    ①★字の 誤り★ ……… ★0件★（PSParser::Tokenize）
#    ②★本数の 門★ ……… ★在り★（下の $式の本数）
#        ★わざと 1本 減らして ★終わり値 1★ を 確かめます★
#        ★★数を 変えた 時にも 壊れます★★（足した 時だけ では ありません）
#    ③★Excel の 門★ …… ★在り★（Get-Process と Win32_Process の ★2つで★）
#    ④★1本ずつ 受け止め★ … ★在り★（try/catch）
#        ★訳★ 2026-09-16 に CALL の 式 1本で ★307秒の 枠が 丸ごと 消えました★
#    ⑤★★司さんの 実物を 開く 字★★ … ★0件★
#        ★指示役1 の 一言★「★特に これを 毎回 数えて いるのが 大事
#                        ／★1回でも 抜けたら 取り返しが つきません★★」
#    ⑥★BOM★ ………… ★在り★（efbbbf）
#
#    ★★門を 足したら その場で わざと 壊して 確かめる★★
#      2026-09-17 … 門を 2行で 書いて ★字の 誤り★に なって いました
#              （PowerShell は 括弧の 中で 行を 続けるのに 逆引用符 が 要る）
#      ⇒★★足しただけでは 門は 効きません★★
#
#  ★★決まり★★
#    ・走らせる ★その時に★ Excel を 2つの 道具で 数える
#    ・★新しい 空の ブックだけ★／Visible=$false／finally で 必ず Quit
#    ・★消えるまで 待って 秒数を 出す★／★2つ目の 窓＋型★
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-oddf-5kaime.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-oddf-5kaime-2026-09-17.tsv'

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel … Get-Process ' + $数1 + '個 ／ Win32_Process ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Error '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$式たち = New-Object System.Collections.Generic.List[object]
$足す = { param($訳, $式) $式たち.Add([pscustomobject]@{ 訳 = $訳; 式 = $式 }) }

# ══ ★ODDFPRICE … 利回りを 3つ 変える★ ══
#   ★利率は 0.06（0 で ない）★＝券の 項を 生かす
#   ★DSC は 4枠目で 出て います★
$形 = @(
  @{ 名 = '端数1期'; 発 = 'DATE(2009,1,1)'; 初 = 'DATE(2009,7,1)' },
  @{ 名 = '端数2期'; 発 = 'DATE(2009,1,1)'; 初 = 'DATE(2010,1,1)' },
  @{ 名 = '端数3期'; 発 = 'DATE(2008,7,1)'; 初 = 'DATE(2010,1,1)' }
)
foreach ($k in $形) {
  foreach ($y in @('0.03', '0.05', '0.08')) {
    foreach ($b in @(0, 1, 2, 3, 4)) {
      & $足す "ODDF 利回り$y $($k.名) (basis=$b)" `
        "=ODDFPRICE(DATE(2009,3,1),DATE(2013,1,1),$($k.発),$($k.初),0.06,$y,100,2,$b)"
    }
  }
}

# ══ ★ODDLPRICE … 同じく 利回りを 3つ★ ══
#   ★f=2 と f=4★（★f が 効いて いる と 4枠目で 出ました★）
#   ★決済は どれも 第1準期間の 中★＝★期越えを 止めて います★
foreach ($f in @(2, 4)) {
  $満 = if ($f -eq 2) { 'DATE(2009,11,30)' } else { 'DATE(2009,8,31)' }
  foreach ($y in @('0.03', '0.05', '0.08')) {
    foreach ($b in @(1, 2, 3)) {
      & $足す "ODDL f=$f 利回り$y 第1期の中 (basis=$b)" `
        "=ODDLPRICE(DATE(2009,1,15),$満,DATE(2008,11,30),0.045,$y,100,$f,$b)"
    }
  }
}

# ══ ★対照（★合わなければ そこで 止める★）★ ══
& $足す '対照1(紙に在る・○のはず)' `
  '=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,2)'
& $足す '対照2(紙に在る・○のはず)' `
  '=ODDFPRICE(DATE(2009,3,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2010,1,1),0.06,0.05,100,2,1)'
& $足す '対照3(前の枠と同じ・○のはず)' `
  '=ODDFPRICE(DATE(2009,7,1),DATE(2013,1,1),DATE(2008,7,1),DATE(2010,1,1),0.06,0.05,100,2,3)'

# ★★本数の 門★★（★数を 変えたら ここも 直し、便りにも 同じ 数を 書く★）
$式の本数 = 66
Write-Host ('★聞く 式 … ' + $式たち.Count + '本★（決め打ち ' + $式の本数 + '本）')
if ($式たち.Count -ne $式の本数) {
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
  $行.Add('# ★DFC と A を 別々に 取り出す★（2026-09-17・5回目）')
  $行.Add('#')
  $行.Add('# ★考え★ DFC の 項は 割^DSC で 割られ、A の 項は 割られません')
  $行.Add('#   ⇒★利回りを 3つ 変えれば DFC と A が 分かれ、揃うかも 見られます★')
  $行.Add('#   ★DSC は 4枠目の「利率 0」で 既に 出て います★')
  $行.Add('#')
  $行.Add('# ★見込みは 聞く 前に 書いて あります★ … kansuu46/oddf-gowakume-no-an.md')
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
