# toru-oddf-to-46ko.ps1 — ★ODDFPRICE の 切り分け ＋ 台に 無い 46個★（2026-09-16）
#
#  ★★1回の 枠で 3つ 聞きます★★（実Excel は Quit まで 120〜300秒＝枠が 重い）
#    ①ODDFPRICE の 切り分け（★見込みは `kansuu46/oddf-kiku-koto.md` に 先に 書いて あります★）
#    ②発行日 ＝ 決済日（★紙に 在りません★）
#    ③台に 無い 46個（★外へ 出る 6個は 入れて いません★）
#
#  ★★外へ 出る 6個は 打ちません★★（経営者1 の 決め・2026-09-16）
#    WEBSERVICE / STOCKHISTORY / TRANSLATE / DETECTLANGUAGE / IMAGE / RTD
#    ★訳★ ㋐司さんの パソコンから 外へ 本当に 出る（お金・秘密・相手の 迷惑）
#          ㋑答えが 日に よって 変わる＝★紙に しても 次の 日 合わない★
#
#  ★★決まり★★
#    ・走らせる 前に ★Excel を 2つの 道具で 数える★（1個でも 居たら 走らせない）
#    ・★新しい 空の ブックだけ★（★司さんの 実物は 開きません★）
#    ・Visible=$false / DisplayAlerts=$false / finally で 必ず Quit
#    ・★消えるまで 待って 秒数を 出す★
#    ・★.ps1 は BOM 必須★
#    ・★2つ目の 窓（=(式)=0）と 型★も 取る（`.Value2` は 0 で ない 値に 0 を 返す）
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-oddf-to-46ko.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-oddf-to-46ko-2026-09-16.tsv'

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★前に 居た Excel … Get-Process ' + $数1 + '個 ／ Win32_Process ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Error '★★Excel が 動いて います＝走らせません★★'; exit 3 }

# ══ ①ODDFPRICE の 切り分け ══
$式たち = New-Object System.Collections.Generic.List[object]
$足す = { param($訳, $式) $式たち.Add([pscustomobject]@{ 訳 = $訳; 式 = $式 }) }

# ㋐★決済を 準利払日の 上に 置く★（★まるごとの 期間を 1と 数えるか だけを 聞く★）
foreach ($b in 0..4) {
  & $足す "㋐決済が準利払日の上(basis=$b)" `
    "=ODDFPRICE(DATE(2009,7,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2010,1,1),0.06,0.05,100,2,$b)"
}
# ㋐′★同じ 形で 端数が 3期★
foreach ($b in 0..4) {
  & $足す "㋐′端数3期・決済が準利払日の上(basis=$b)" `
    "=ODDFPRICE(DATE(2009,7,1),DATE(2013,1,1),DATE(2008,7,1),DATE(2010,1,1),0.06,0.05,100,2,$b)"
}
# ㋑★A（発行→決済）だけを 動かす★（basis 2 と 3）
foreach ($d in @('DATE(2009,2,1)', 'DATE(2009,3,1)', 'DATE(2009,4,1)', 'DATE(2009,7,1)', 'DATE(2009,10,1)')) {
  foreach ($b in @(2, 3)) {
    & $足す "㋑決済を動かす $d (basis=$b)" `
      "=ODDFPRICE($d,DATE(2013,1,1),DATE(2009,1,1),DATE(2010,1,1),0.06,0.05,100,2,$b)"
  }
}
# ㋒★発行日 ＝ 決済日★（★紙に 在りません★）
foreach ($b in 0..4) {
  & $足す "㋒発行日=決済日(basis=$b)" `
    "=ODDFPRICE(DATE(2009,1,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2010,1,1),0.06,0.05,100,2,$b)"
}
# ㋓★対照★（★紙に 在って ○ の 物を もう一度 打つ＝測り台が 変わって いないかの 確かめ★）
& $足す '㋓対照(紙に在る・○のはず)' `
  '=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,2)'
& $足す '㋓対照(紙に在る・○のはず)' `
  '=ODDFPRICE(DATE(2009,3,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2010,1,1),0.06,0.05,100,2,1)'

# ㆖★★ODDLPRICE … 決済を 準利払日の 上を またいで 動かす★★
#   ★なぜ★ ODDLPRICE は 6/12。★落ちて いるのは 1組（6本）だけ★で
#         その 組は ★決済が 最終利払から 1期を 越えて いる★ 形。
#   ★準利払日★ 2008-11-30 ／ 2009-02-28 ／ 2009-05-31 ／ 2009-08-31
#   ★見る 事★ … ★準利払日の 上で 答えに ★段★が 付くか★
#   ★見込み（★先に 書きます★）★
#     ㆁ段が 付かない（なめらか）⇒★期を またぐ 数え方は うちと 同じ★
#                               ⇒★違いは 別の 所（DC か 分母）★
#     ㆂ段が 付く ⇒★期ごとに 別の 数え方を して いる★
#   ★逆算して 止めました★ … 実Excel の 答えから A を 逆に 出したところ
#     A×92 が 99.0367 など ★整数日に なりません★
#     ⇒★A だけが 違う のでは 無い★＝★ここから 先は 当てに なるので 聞きます★
foreach ($d in @('DATE(2009,2,27)', 'DATE(2009,2,28)', 'DATE(2009,3,1)', 'DATE(2009,3,10)',
                 'DATE(2009,5,30)', 'DATE(2009,5,31)', 'DATE(2009,6,1)')) {
  foreach ($b in @(0, 1)) {
    & $足す "㆖ODDLPRICE 決済 $d (basis=$b)" `
      "=ODDLPRICE($d,DATE(2009,8,31),DATE(2008,11,30),0.045,0.05,100,4,$b)"
  }
}
# ㆗★ODDLPRICE の 対照（★紙に 在って ○ の 組★）
& $足す '㆗対照 ODDLPRICE(紙に在る・○のはず)' `
  '=ODDLPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0)'

# ══ ②台に 無い 46個 ══（★外へ 出る 6個は 入れて いません★）
#   ★引数は 全部 かんたんな 物★＝★答えが 誤りでも 構いません★
#   ★見たいのは 「その 名前が 実Excel で 通るか」と 「どんな 答えを 返すか」★
$四十六 = @(
  @('AGGREGATE', '=AGGREGATE(1,0,A1:A5)'),
  @('ASC', '=ASC("ABC")'),
  @('BYCOL', '=BYCOL(A1:B2,LAMBDA(c,SUM(c)))'),
  @('BYROW', '=BYROW(A1:B2,LAMBDA(r,SUM(r)))'),
  @('CALL', '=CALL("x")'),
  @('CUBEKPIMEMBER', '=CUBEKPIMEMBER("x","y",1)'),
  @('CUBEMEMBER', '=CUBEMEMBER("x","y")'),
  @('CUBEMEMBERPROPERTY', '=CUBEMEMBERPROPERTY("x","y","z")'),
  @('CUBERANKEDMEMBER', '=CUBERANKEDMEMBER("x","y",1)'),
  @('CUBESET', '=CUBESET("x","y")'),
  @('CUBESETCOUNT', '=CUBESETCOUNT(A1)'),
  @('CUBEVALUE', '=CUBEVALUE("x","y")'),
  @('DBCS', '=DBCS("ABC")'),
  @('DOLLAR', '=DOLLAR(1234.567,2)'),
  @('ENCODEURL', '=ENCODEURL("a b")'),
  @('FIXED', '=FIXED(1234.567,2)'),
  @('FORECAST', '=FORECAST(6,B1:B5,A1:A5)'),
  @('FORECAST.LINEAR', '=FORECAST.LINEAR(6,B1:B5,A1:A5)'),
  @('GETPIVOTDATA', '=GETPIVOTDATA("x",A1)'),
  @('GROUPBY', '=GROUPBY(A1:A5,B1:B5,SUM)'),
  @('INDIRECT', '=INDIRECT("A1")'),
  @('KURT', '=KURT(A1:A5)'),
  @('LEFTB', '=LEFTB("ABC",2)'),
  @('LENB', '=LENB("ABC")'),
  @('LOOKUP', '=LOOKUP(3,A1:A5,B1:B5)'),
  @('MAKEARRAY', '=MAKEARRAY(2,2,LAMBDA(r,c,r*c))'),
  @('MAP', '=MAP(A1:A5,LAMBDA(x,x*2))'),
  @('MDETERM', '=MDETERM(A1:B2)'),
  @('MIDB', '=MIDB("ABC",2,1)'),
  @('NUMBERVALUE', '=NUMBERVALUE("1.5")'),
  @('OFFSET', '=OFFSET(A1,1,0)'),
  @('PERCENTRANK', '=PERCENTRANK(A1:A5,3)'),
  @('PERMUT', '=PERMUT(5,2)'),
  @('PERMUTATIONA', '=PERMUTATIONA(5,2)'),
  @('PHONETIC', '=PHONETIC(A1)'),
  @('PIVOTBY', '=PIVOTBY(A1:A5,B1:B5,B1:B5,SUM)'),
  @('REDUCE', '=REDUCE(0,A1:A5,LAMBDA(a,b,a+b))'),
  @('REGISTER.ID', '=REGISTER.ID("x","y")'),
  @('RIGHTB', '=RIGHTB("ABC",2)'),
  @('SCAN', '=SCAN(0,A1:A5,LAMBDA(a,b,a+b))'),
  @('TEXTAFTER', '=TEXTAFTER("a-b","-")'),
  @('TEXTBEFORE', '=TEXTBEFORE("a-b","-")'),
  @('TRIMMEAN', '=TRIMMEAN(A1:A5,0.2)'),
  @('TRIMRANGE', '=TRIMRANGE(A1:A5)'),
  @('VALUETOTEXT', '=VALUETOTEXT(A1)'),
  @('XMATCH', '=XMATCH(3,A1:A5)')
)
Write-Host ('★聞く 46個 … ' + $四十六.Count + '個★（★外へ 出る 6個は 入れて いません★）')
if ($四十六.Count -ne 46) { Write-Error ('★46個の はずが ' + $四十六.Count + '個です★'); exit 4 }

$xl = New-Object -ComObject Excel.Application
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)

  # ★材料★ A1:A5＝1〜5 ／ B1:B5＝2,4,6,8,10
  1..5 | ForEach-Object { $sh.Range('A' + $_).Value2 = $_ }
  1..5 | ForEach-Object { $sh.Range('B' + $_).Value2 = $_ * 2 }

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★ODDFPRICE の 切り分け ＋ 台に 無い 46個★（2026-09-16）')
  $行.Add('#')
  $行.Add('# ★見込みは 先に 書いて あります★ … docs/measured/kansuu46/oddf-kiku-koto.md')
  $行.Add('#   ⇒★後から 答えに 寄せられません★')
  $行.Add('#')
  $行.Add('# ★外へ 出る 6個は 打って いません★')
  $行.Add('#   WEBSERVICE / STOCKHISTORY / TRANSLATE / DETECTLANGUAGE / IMAGE / RTD')
  $行.Add('#')
  $行.Add('# ★2つ目の 窓★ `=(式)=0` … `.Value2` は 0 で ない 値に 0 を 返す')
  $行.Add('#')
  $行.Add('# ★どの Excel か★ … 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('#')
  $行.Add('# 種' + "`t" + '訳' + "`t" + '式' + "`t" + '答え' + "`t" + '出る字' + "`t" + '=(式)=0' + "`t" + '型')

  $r = 1
  # ★★打てない 式が 1本 在るだけで 全部 止まる 形を 直しました★★（2026-09-16）
  #   ★実際に 起きました★ … `HRESULT からの 例外:0x800A03EC`
  #     ＝★Excel が その 式を 受け付けなかった★（名前が 無い・引数の 形が 違う など）
  #   ⇒★★その 1本で ★307秒 掛けた 枛が 丸ごと 無駄★に なりました★★
  #   ★しかも★ 「打てない」も ★測った 結果★です（★捨てて は いけません★）
  #   ⇒★★、1本ずつ 受け止めて 「打てません」と 控えます★★
  $打つ = {
    param($種, $訳, $式)
    $c = $sh.Range('D' + $r)
    $打てた = $true
    try { $c.Formula = $式 } catch {
      $打てた = $false
      $行.Add($種 + "`t" + $訳 + "`t" + $式 + "`t" + '(★打てません★)' + "`t" +
        ('★Excel が 式を 受け付けません★ ' + $_.Exception.Message) + "`t" + '(★打てません★)' + "`t" + '(★打てません★)')
      $script:r = $r + 1
    }
    if (-not $打てた) { return }
    $v = $c.Value2
    $答 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    $型 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
    $字 = [string]$c.Text
    $w = $sh.Range('E' + $r)
    $ゼロか = '(★窓②が 打てません★)'
    try { $w.Formula = '=(' + $式.Substring(1) + ')=0'; $ゼロか = [string]$w.Value2 } catch { }
    $行.Add($種 + "`t" + $訳 + "`t" + $式 + "`t" + $答 + "`t" + $字 + "`t" + $ゼロか + "`t" + $型)
    $script:r = $r + 1
  }

  foreach ($x in $式たち) { & $打つ 'ODDF' $x.訳 $x.式 }
  Write-Host ('  ODDFPRICE の 切り分け … ' + $式たち.Count + '本 済')
  foreach ($x in $四十六) { & $打つ '46個' $x[0] $x[1] }
  Write-Host ('  46個 … 済')

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★書いた … ' + $出 + '（' + ($式たち.Count + $四十六.Count) + '行）★')
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
