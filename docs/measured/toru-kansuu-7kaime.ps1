# toru-kansuu-7kaime.ps1 — ★書く 前に 聞く★（2026-09-18・7枠目）
#
#  ★★なぜ この 枠か★★
#    順列 2つを 書いた 時、★境目を COMBIN に 合わせた だけ★でした（★測って いません★）。
#    ★これを 56回 繰り返すと 台に「測って いない 決め」が 56個 積まれます★。
#    ⇒★★書く 前に 聞く★★（経営者1 の 決め）
#
#  ★見込みは 聞く 前に commit 済み★ … docs/measured/kansuu46/nanawakume-no-an.md（9e6d5b9）
#
#  ★★★この 道具が 通した 門（12個）★★★
#    ★なぜ 頭に 書くか★ … ★★門は 引き継がれません★★
#    ①★字の 誤り★ ……… ★0件★（PSParser::Tokenize で 数える）
#    ②★本数の 門★ ……… ★在り★（$式の本数 = 91／違えば exit 4）
#    ③★Excel の 門★ …… ★在り★（Get-Process と Win32_Process の ★2つで★）
#    ④★1本ずつ 受け止め★ … ★在り★（try/catch）
#        ★訳★ 2026-09-16 に CALL の 式 1本で ★307秒の 枠が 丸ごと 消えました★
#    ⑤★★司さんの ブックを 開く 字★★ … ★0件★（★新しい 空の ブックだけ★）
#    ⑥★BOM★ ………… ★在り★（efbbbf）
#    ⑦★2つ目の 窓＋型★ … ★在り★（.Value2 は 0 で ない 値に 0 を 返す）
#    ⑧★消えるまで 待って 秒数を 出す★ … ★在り★
#    ⑨★★外へ 出る 6個が 式に 1つも 無い★★ … ★在り★（在れば exit 5）
#        WEBSERVICE ／ STOCKHISTORY ／ TRANSLATE ／ DETECTLANGUAGE ／ IMAGE ／ RTD
#        ★訳★ ★司さんの パソコンから 本当に 外へ 出る★／★日で 答えが 変わる★（09-16 の 決め）
#    ⑩★★押した 時間を 別に 測る★★ … ★在り★
#        ★訳★ 6枠目 … 全体 307秒 の うち ★300.4秒は Excel が 消えるのを 待った 時間★
#              ＝★★93本を 押したのは 約 6.6秒★★
#        ⇒★「300秒 待った」と「300秒で 消えた」は ★別★★ ⇒ 両方 書きます
#    ⑪★★わざと 壊す 前に「未commit 0本」を 数える★★
#        ★訳★ 2026-09-18 … commit する 前に `git checkout --` を 打って
#              ★自分の 直しを 消しました★
#    ⑫★★門を 壊して 確かめる 時は 管（`| tail`）を 通さない★★
#        ★訳★ 2026-09-18 … `cmd | tail` の 後の `$?` は ★tail の 値★
#              ⇒★「終わり値 0＝嘘の 緑」と 一度 書いた★（★私の 測り誤り★）
#
#  ★★字は UNICHAR で 作ります★★
#    ★訳★ 全角・半角・濁点を `.ps1` と `.tsv` に 直に 書くと ★化ける★
#    ⇒★★式の 中は ASCII だけ★★／★答えも 字でなく コードと 文字数で 受けます★
#
#  ★★repo が 公開です（2026-09-18）★★
#    ⇒★紙に 司さんの 商売の 中身を 書きません★（この 紙は 教科書の 数だけ）
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-kansuu-7kaime.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-kansuu-7kaime-2026-09-18.tsv'

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel … Get-Process ' + $数1 + '個 ／ Win32_Process ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$式たち = New-Object System.Collections.Generic.List[object]
$足す = { param($訳, $式) $式たち.Add([pscustomobject]@{ 訳 = $訳; 式 = $式 }) }

# ══ ★対照 3本★（★合わなければ そこで 止める★）══
& $足す '対照1(PERMUT・紙に在る)' '=PERMUT(5,2)'
& $足す '対照2(ODDL f4・紙に在る)' '=ODDLPRICE(DATE(2009,3,10),DATE(2009,8,31),DATE(2008,11,30),0.045,0.05,100,4,1)'
& $足す '対照3(ODDF・前の枠と同じ)' '=ODDFPRICE(DATE(2009,3,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2010,1,1),0.06,0.05,100,2,1)'

# ══ ⑴★順列の 境目 12本★ ══（★うちは COMBIN に 合わせた だけ＝測って いません★）
foreach ($n in '0,0', '5,0', '0,1', '3,4', '-1,1', '5,-1') {
  & $足す ('(1)順列 PERMUT(' + $n + ') ★境目★') ('=PERMUT(' + $n + ')')
}
foreach ($n in '0,0', '5,0', '0,1', '3,4', '-1,1', '5,-1') {
  & $足す ('(1)順列 PERMUTATIONA(' + $n + ') ★境目★') ('=PERMUTATIONA(' + $n + ')')
}

# ══ ⑵★AGGREGATE 12本★ ══（★19機能×8選択は 聞きません＝迷う 所だけ★）
#   ★材料★ A1:A5 ＝ 1,2,3,4,5 ／ B1:B5 ＝ 1,2,(#DIV/0!),4,5
$AG = @(
  @{ 訳 = '選択0＝何も無視しない'; 式 = '=AGGREGATE(9,0,B1:B5)' },
  @{ 訳 = '選択2＝誤りを無視'; 式 = '=AGGREGATE(9,2,B1:B5)' },
  @{ 訳 = '選択6＝誤りを無視'; 式 = '=AGGREGATE(9,6,B1:B5)' },
  @{ 訳 = '選択4＝何も無視しない'; 式 = '=AGGREGATE(9,4,B1:B5)' },
  @{ 訳 = '選択7＝隠した行だけ無視'; 式 = '=AGGREGATE(9,7,B1:B5)' },
  @{ 訳 = '機能1(平均)×選択6'; 式 = '=AGGREGATE(1,6,B1:B5)' },
  @{ 訳 = '機能14(LARGE)＋k'; 式 = '=AGGREGATE(14,6,B1:B5,2)' },
  @{ 訳 = '機能15(SMALL)＋k'; 式 = '=AGGREGATE(15,6,B1:B5,2)' },
  @{ 訳 = '機能19(QUARTILE)＋k'; 式 = '=AGGREGATE(19,6,B1:B5,1)' },
  @{ 訳 = '★余分な引数★(9に k)'; 式 = '=AGGREGATE(9,6,B1:B5,2)' },
  @{ 訳 = '★k が要るのに無い★(14)'; 式 = '=AGGREGATE(14,6,B1:B5)' },
  @{ 訳 = '★範囲2つ★'; 式 = '=AGGREGATE(9,0,A1:A5,B1:B5)' }
)
foreach ($x in $AG) { & $足す ('(2)AGGREGATE ' + $x.訳) $x.式 }

# ══ ⑶★XMATCH 12本★ ══（★材料★ C1:C5 ＝ 1,3,5,7,9）
$XM = @(
  @{ 訳 = 'ぴたり在り(既定)'; 式 = '=XMATCH(5,C1:C5)' },
  @{ 訳 = 'ぴたり無し(既定)'; 式 = '=XMATCH(4,C1:C5)' },
  @{ 訳 = '一致-1＝次に小さい'; 式 = '=XMATCH(4,C1:C5,-1)' },
  @{ 訳 = '一致1＝次に大きい'; 式 = '=XMATCH(4,C1:C5,1)' },
  @{ 訳 = '一致0＝ぴたりだけ'; 式 = '=XMATCH(4,C1:C5,0)' },
  @{ 訳 = '下に無い(一致-1)'; 式 = '=XMATCH(0,C1:C5,-1)' },
  @{ 訳 = '上に無い(一致1)'; 式 = '=XMATCH(10,C1:C5,1)' },
  @{ 訳 = '検索-1＝後ろから'; 式 = '=XMATCH(9,C1:C5,0,-1)' },
  @{ 訳 = '検索2＝二分(昇順)'; 式 = '=XMATCH(9,C1:C5,0,2)' },
  @{ 訳 = '★検索-2＝二分(降順)★'; 式 = '=XMATCH(9,C1:C5,0,-2)' },
  @{ 訳 = '一致2＝型紙だが数'; 式 = '=XMATCH(3,C1:C5,2)' },
  @{ 訳 = '★無い一致の型(3)★'; 式 = '=XMATCH(4,C1:C5,3)' }
)
foreach ($x in $XM) { & $足す ('(3)XMATCH ' + $x.訳) $x.式 }

# ══ ⑷★ASC ／ DBCS 22本★ ══（★式の 中は ASCII だけ★）
#   ★コードで 字を 作り、コードと 文字数で 受けます★
$ASC字 = @(
  @{ 码 = 65313; 訳 = '全角A' }, @{ 码 = 65297; 訳 = '全角1' },
  @{ 码 = 12450; 訳 = '全角ア' }, @{ 码 = 12460; 訳 = '★全角ガ(濁点)★' },
  @{ 码 = 12288; 訳 = '★全角空白★' }, @{ 码 = 12289; 訳 = '全角読点' }
)
foreach ($x in $ASC字) {
  & $足す ('(4)ASC ' + $x.訳 + ' ⇒出た字のコード') ('=UNICODE(ASC(UNICHAR(' + $x.码 + ')))')
  & $足す ('(4)ASC ' + $x.訳 + ' ⇒何文字か') ('=LEN(ASC(UNICHAR(' + $x.码 + ')))')
}
$DBCS字 = @(
  @{ 码 = 65; 訳 = '半角A' }, @{ 码 = 49; 訳 = '半角1' },
  @{ 码 = 65393; 訳 = '半角ｱ' }, @{ 码 = 65438; 訳 = '★半角濁点★' },
  @{ 码 = 32; 訳 = '半角空白' }
)
foreach ($x in $DBCS字) {
  & $足す ('(4)DBCS ' + $x.訳 + ' ⇒出た字のコード') ('=UNICODE(DBCS(UNICHAR(' + $x.码 + ')))')
  & $足す ('(4)DBCS ' + $x.訳 + ' ⇒何文字か') ('=LEN(DBCS(UNICHAR(' + $x.码 + ')))')
}

# ══ ⑸★TEXTAFTER ／ TEXTBEFORE 12本★ ══
$TX = @(
  @{ 訳 = '1つ目の区切り'; 式 = '=TEXTAFTER("a-b-c","-")' },
  @{ 訳 = '2つ目の区切り'; 式 = '=TEXTAFTER("a-b-c","-",2)' },
  @{ 訳 = '後ろから1つ目'; 式 = '=TEXTAFTER("a-b-c","-",-1)' },
  @{ 訳 = '区切りが無い'; 式 = '=TEXTAFTER("a-b-c","x")' },
  @{ 訳 = '★大小を見るか(既定)★'; 式 = '=TEXTAFTER("a-B-c","b")' },
  @{ 訳 = '大小を見る(1)'; 式 = '=TEXTAFTER("a-B-c","b",1,1)' },
  @{ 訳 = 'BEFORE 1つ目'; 式 = '=TEXTBEFORE("a-b-c","-")' },
  @{ 訳 = 'BEFORE 2つ目'; 式 = '=TEXTBEFORE("a-b-c","-",2)' },
  @{ 訳 = 'BEFORE 後ろから'; 式 = '=TEXTBEFORE("a-b-c","-",-1)' },
  @{ 訳 = 'BEFORE 区切り無い'; 式 = '=TEXTBEFORE("a-b-c","x")' },
  @{ 訳 = '★行き過ぎた番号(5)★'; 式 = '=TEXTAFTER("a-b-c","-",5)' },
  @{ 訳 = '★空の区切り★'; 式 = '=TEXTBEFORE("abc","")' }
)
foreach ($x in $TX) { & $足す ('(5)' + $x.訳) $x.式 }

# ══ ⑹★LENB ／ LEFTB ／ RIGHTB ／ MIDB 12本★ ══（★この国の Excel の 数え方★）
$LB = @(
  @{ 訳 = 'LENB 半角3文字'; 式 = '=LENB("ABC")' },
  @{ 訳 = 'LENB 全角ア'; 式 = '=LENB(UNICHAR(12354))' },
  @{ 訳 = 'LENB 半角ｱ'; 式 = '=LENB(UNICHAR(65393))' },
  @{ 訳 = 'LENB 全角ア＋A'; 式 = '=LENB(UNICHAR(12354)&"A")' },
  @{ 訳 = '★LEFTB 半端な1バイト(文字数)★'; 式 = '=LEN(LEFTB(UNICHAR(12354),1))' },
  @{ 訳 = '★LEFTB 半端な1バイト(コード)★'; 式 = '=UNICODE(LEFTB(UNICHAR(12354),1))' },
  @{ 訳 = 'LEFTB 2バイト'; 式 = '=LEN(LEFTB(UNICHAR(12354),2))' },
  @{ 訳 = 'LEFTB 3バイト(ア＋A)'; 式 = '=LEN(LEFTB(UNICHAR(12354)&"A",3))' },
  @{ 訳 = 'RIGHTB 1バイト'; 式 = '=LEN(RIGHTB(UNICHAR(12354)&"A",1))' },
  @{ 訳 = 'MIDB 1から2'; 式 = '=LEN(MIDB(UNICHAR(12354)&"A",1,2))' },
  @{ 訳 = '★MIDB 2から2(半端)★'; 式 = '=LEN(MIDB(UNICHAR(12354)&"A",2,2))' },
  @{ 訳 = 'LENB 空'; 式 = '=LENB("")' }
)
foreach ($x in $LB) { & $足す ('(6)' + $x.訳) $x.式 }

# ══ ⑺★PERCENTRANK 6本★ ══（★材料★ C1:C5 ＝ 1,3,5,7,9）
$PR = @(
  @{ 訳 = '真ん中'; 式 = '=PERCENTRANK(C1:C5,5)' },
  @{ 訳 = '一番下'; 式 = '=PERCENTRANK(C1:C5,1)' },
  @{ 訳 = '一番上'; 式 = '=PERCENTRANK(C1:C5,9)' },
  @{ 訳 = '★間の値(有効桁)★'; 式 = '=PERCENTRANK(C1:C5,4)' },
  @{ 訳 = '★有効桁5★'; 式 = '=PERCENTRANK(C1:C5,4,5)' },
  @{ 訳 = '範囲の外'; 式 = '=PERCENTRANK(C1:C5,10)' }
)
foreach ($x in $PR) { & $足す ('(7)PERCENTRANK ' + $x.訳) $x.式 }

# ══ ★⑨外へ 出る 6個が 1つも 無いか★ ══（★走らせる 前に 機械が 数える★）
$外へ出る = 'WEBSERVICE', 'STOCKHISTORY', 'TRANSLATE', 'DETECTLANGUAGE', 'IMAGE', 'RTD'
$見つけた = New-Object System.Collections.Generic.List[string]
foreach ($x in $式たち) {
  foreach ($n in $外へ出る) {
    if ($x.式 -match ('(^|[^A-Z])' + $n + '\(')) { $見つけた.Add($n + ' … ' + $x.訳) }
  }
}
Write-Host ('★外へ 出る 6個 … ' + $見つけた.Count + '件★（0で あるべき）')
if ($見つけた.Count -ne 0) {
  Write-Host ('★★外へ 出る 関数が 入って います★★ ' + ($見つけた -join ' / '))
  exit 5
}

# ══ ★②本数の 門★ ══
$式の本数 = 91
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

  # ★材料★（★A 1〜5 ／ B は 3行目が #DIV/0! ／ C 1,3,5,7,9★）
  for ($i = 1; $i -le 5; $i++) { $sh.Range('A' + $i).Value2 = $i }
  $sh.Range('B1').Value2 = 1
  $sh.Range('B2').Value2 = 2
  $sh.Range('B3').Formula = '=1/0'
  $sh.Range('B4').Value2 = 4
  $sh.Range('B5').Value2 = 5
  $C = 1, 3, 5, 7, 9
  for ($i = 1; $i -le 5; $i++) { $sh.Range('C' + $i).Value2 = $C[$i - 1] }

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★書く 前に 聞く★（2026-09-18・7回目）')
  $行.Add('#')
  $行.Add('# ★見込みは 聞く 前に commit 済み★ … kansuu46/nanawakume-no-an.md（9e6d5b9）')
  $行.Add('#')
  $行.Add('# ★字は UNICHAR で 作り、答えは コードと 文字数で 受けます★')
  $行.Add('#   ⇒★式の 中は ASCII だけ＝化ける 所が 1つも ありません★')
  $行.Add('#')
  $行.Add('# ★材料★ A1:A5 = 1,2,3,4,5 ／ B1:B5 = 1,2,(=1/0),4,5 ／ C1:C5 = 1,3,5,7,9')
  $行.Add('#')
  $行.Add('# ★2つ目の 窓★ =(式)=0 … .Value2 は 0 で ない 値に 0 を 返す')
  $行.Add('#')
  $行.Add('# ★外へ 出る 6個★ … ★0件★（門で 数えました）')
  $行.Add('#')
  $行.Add('# ★どの Excel か★ … 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('#')
  $行.Add('# 訳' + "`t" + '式' + "`t" + '答え' + "`t" + '出る字' + "`t" + '=(式)=0' + "`t" + '型')

  # ★⑩押した 時間を 別に 測る★
  $押し時計 = [Diagnostics.Stopwatch]::StartNew()
  $r = 1
  foreach ($x in $式たち) {
    $c = $sh.Range('E' + $r)
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
    $w = $sh.Range('F' + $r)
    $ゼロか = '(★窓2が 打てません★)'
    try { $w.Formula = '=(' + $x.式.Substring(1) + ')=0'; $ゼロか = [string]$w.Value2 } catch { }
    $行.Add($x.訳 + "`t" + $x.式 + "`t" + $答 + "`t" + $字 + "`t" + $ゼロか + "`t" + $型)
    $r++
  }
  $押し時計.Stop()
  $押し秒 = [math]::Round($押し時計.Elapsed.TotalSeconds, 2)
  $行.Add('# ★★押すのに かかった 秒 … ' + $押し秒 + '★★（★Excel が 消えるのを 待つ 時間とは 別★）')

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★★押すのに かかった 秒 … ' + $押し秒 + '秒★★（' + $式たち.Count + '本）')
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
  $残り = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
  # ★★「300秒 待った」と「300秒で 消えた」は 別★★（経営者1 の 注文）
  if ($残り -eq 0) {
    Write-Host ('★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒で ★消えました★★')
  } else {
    Write-Host ('★★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒 待っても ★消えませんでした★★ ／ 残り ' + $残り + '個')
    Write-Host ('  ⇒★この 道具が 何かを 掴んだ まま です★（★次の 枠の 後に 直します★）')
  }
  # ★★指示役1 の 型（2026-09-18）★★
  Write-Host ''
  Write-Host '★★秒が 毎回 同じなら ★上限に 当てて いる★ か ★空振り★ を 疑う★★'
  Write-Host '  ★仕事の 秒なら バラつきます★（総なめ 48／53／576／168 … バラバラ）'
  Write-Host '  ★この 枠は 毎回 300秒台 … ★上限に 当たって います★'
  if ($false) {
  }
}
