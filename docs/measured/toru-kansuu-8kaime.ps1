# toru-kansuu-8kaime.ps1 - ★溜まって いる 問いを 1枠に★（2026-09-18・8枠目）
#
#  ★★中身★★（★7つの 紙から 機械で 集めました／手で 書き写して いません★）
#    集める 道具 … docs/measured/kansuu46/toi-atsumeru.py
#    集めた 紙 … docs/measured/kansuu46/toi-atsumeta.txt（★94本★）
#    ＋ODDL の 割れ目 18本 ＋ 対照 3本 ＝★★115本★★
#
#  ★★★新しい 型で 書いて います（2026-09-18 に 測って 決着）★★★
#    ★終わりは `$sh = $null; $bk = $null; $xl = $null` だけ★
#    ★★`ReleaseComObject` を 掴んだ 物ごと／`GC::Collect` は ★足しません★★★
#    ★訳（交互に 2回ずつ 測りました）★
#      null だけ ………………… ★2.71秒 ／ 13.03秒★
#      ＋Release＋GC ………… ★60.22秒 ／ 60.10秒★
#    ⇒★★世の中で 一番 よく 言われる 直しが ★逆効果★でした★★
#    ⇒★★次の 人は 必ず 足そうと します＝足さないで ください★★
#    ⇒紙 docs/measured/kansuu46/excel-kienai-kekka.md
#
#  ★★★この 道具が 通した 門（11個）★★★
#    ★なぜ 頭に 書くか★ … ★★門は 引き継がれません★★
#    ①★字の 誤り★ ……… ★0件★（PSParser::Tokenize）
#    ②★本数の 門★ ……… ★$式の本数 = 115／違えば exit 4★
#    ③★Excel の 門★ …… ★Get-Process と Win32_Process の 2つで／1個でも 居たら exit 3★
#    ④★1本ずつ 受け止め★ … try/catch
#        ★訳★ 2026-09-16 に CALL の 式 1本で ★307秒の 枠が 丸ごと 消えました★
#    ⑤★★司さんの ブックを 開く 字★★ … ★0件★（★新しい 空の ブックだけ★）
#    ⑥★BOM★（efbbbf）
#    ⑦★2つ目の 窓＋型★（.Value2 は 0 で ない 値に 0 を 返す）
#    ⑧★押した秒と 消え秒を ★別々に★★／★「待った」と「消えた」を 書き分ける★
#    ⑨★★外へ 出る 6個が 0件★★（在れば exit 5）
#        WEBSERVICE ／ STOCKHISTORY ／ TRANSLATE ／ DETECTLANGUAGE ／ IMAGE ／ RTD
#    ⑩★★式に ASCII の 外の 字が 0件★★（★字は `UNICHAR(番号)` で 作る★／在れば exit 6）
#        ★訳★ 化けた 時に ★答えが 違うのか 字が 化けたのかが 見分けられません★
#        ★この 門は 足した その日に ★3本 見つけました★★
#    ⑪★★新しい 型（null だけ・Release と GC は 足さない）★★
#    ⑭★★貝殻の 版★★ … ★`powershell.exe`（5.1）でないと 止まる（exit 8）★
#        ★訳★ 5.1 と 7.x は ★同じ 数を 違う 字で 書きます★（今日 黙って 変わった）
#        ★紙の 頭にも 版と 文字コードを 書きます★
#    ⑬★★集め漏れ★★ … ★㋐集めた 時／㋑走った 時／㋒紙に 書かれた の 3つを 並べる★
#        ★紙ごとにも 並べる★（★1紙 落ちても 合計が 合う 事が 在る★）／違えば exit 7
#        ★指示役1 の 注文（2026-09-18）★ … 機械で 集めた から こそ 1紙 落ちても 気づけない
#
#  ★★見込みは 聞く 前に commit 済み★★ … docs/measured/kansuu46/hachiwakume-no-an.md
#
#  ★★repo は 公開です★★ ⇒★紙に 司さんの 商売の 中身を 書きません★
#    ⇒★この 紙の 日付・金額は 全部 作り物です★
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-kansuu-8kaime.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-kansuu-8kaime-2026-09-18.tsv'

# ══ ★⑭走らせる 貝殻の 版の 門★ ══（2026-09-18・★決め＝5.1 に 揃える★）
#   ★★今日 黙って 変わりました★★
#     7回目 … powershell.exe 5.1 ⇒ 99.776476358001872
#     8回目 … pwsh 7.6           ⇒ 99.77647635800187
#     ★double は bit まで 同じ★（4636721562276654862）
#     ★訳★ `'R'` の 意味が 違う（5.1＝17桁 ／ 7.x＝最短 往復）
#   ★決め★ … ★★34本の 古い 紙が 5.1 で 取って 在る★★ ⇒★揃って いる 方に 合わせる★
#     ＝★選ぶ 訳は「正しさ」では なく「揃って いるか」だけ★（数は どちらも 同じ）
#   ⇒★★門が 無いと また 黙って 変わります★★
$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 … PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) {
  Write-Host '★★この 道具は `powershell.exe`（5.1）で 走らせて ください★★'
  Write-Host ('  ★今 … ' + $版.ToString() + '★')
  Write-Host '  ★訳★ … 5.1 と 7.x は ★同じ 数を 違う 字で 書きます★'
  Write-Host "         5.1 … 99.776476358001872 ／ 7.x … 99.77647635800187"
  Write-Host '  ⇒★34本の 古い 紙は 5.1 で 取って 在ります★＝★揃って いる 方に 合わせます★'
  Write-Host '  ★走らせ方★ powershell.exe -NoProfile -File docs/measured/toru-kansuu-8kaime.ps1'
  exit 8
}

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel … Get-Process ' + $数1 + '個 ／ Win32_Process ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

# ══ ★紙から 集めた 94本★ ══（★docs/measured/kansuu46/toi-atsumeta.txt と 同じ★）
$紙の式 = @(
  @{ 訳 = '(紙)junretsu'; 式 = '=PERMUT(0,0)' },
  @{ 訳 = '(紙)junretsu'; 式 = '=PERMUTATIONA(0,0)' },
  @{ 訳 = '(紙)junretsu'; 式 = '=PERMUT(5,0)' },
  @{ 訳 = '(紙)junretsu'; 式 = '=PERMUTATIONA(5,0)' },
  @{ 訳 = '(紙)junretsu'; 式 = '=PERMUT(0,1)' },
  @{ 訳 = '(紙)junretsu'; 式 = '=PERMUTATIONA(0,1)' },
  @{ 訳 = '(紙)junretsu'; 式 = '=PERMUT(3,4)' },
  @{ 訳 = '(紙)junretsu'; 式 = '=PERMUTATIONA(3,4)' },
  @{ 訳 = '(紙)junretsu'; 式 = '=PERMUT(-1,1)' },
  @{ 訳 = '(紙)junretsu'; 式 = '=PERMUTATIONA(-1,1)' },
  @{ 訳 = '(紙)junretsu'; 式 = '=PERMUT(5,-1)' },
  @{ 訳 = '(紙)junretsu'; 式 = '=PERMUTATIONA(5,-1)' },
  @{ 訳 = '(紙)xmatch'; 式 = '=XMATCH("*b*",{"ab";"bc";"cd"},2)' },
  @{ 訳 = '(紙)xmatch'; 式 = '=XMATCH("B",{"a";"b";"c"},0)' },
  @{ 訳 = '(紙)xmatch'; 式 = '=XMATCH("b",{"a";"b";"c"},2)' },
  @{ 訳 = '(紙)xmatch'; 式 = '=XMATCH(3,A1:B5)' },
  @{ 訳 = '(紙)xmatch'; 式 = '=XMATCH(4,C1:C5,0,7)' },
  @{ 訳 = '(紙)xmatch'; 式 = '=XMATCH(4,C1:C5,-1,-1)' },
  @{ 訳 = '(紙)xmatch'; 式 = '=XMATCH(4,C1:C5,-1,2)' },
  @{ 訳 = '(紙)xmatch'; 式 = '=XMATCH(4,C1:C5,1,-2)' },
  @{ 訳 = '(紙)xmatch'; 式 = '=XMATCH(4,D1:D5,-1)' },
  @{ 訳 = '(紙)xmatch'; 式 = '=XMATCH(4,B1:B5,-1)' },
  @{ 訳 = '(紙)xmatch'; 式 = '=XMATCH(TRUE,{TRUE;FALSE},0)' },
  @{ 訳 = '(紙)xmatch'; 式 = '=XMATCH(1,{})' },
  @{ 訳 = '(紙)percentrank'; 式 = '=PERCENTRANK({1;2;3;4},2.6,1)' },
  @{ 訳 = '(紙)percentrank'; 式 = '=PERCENTRANK({1;2;3;4},2.9,1)' },
  @{ 訳 = '(紙)percentrank'; 式 = '=PERCENTRANK({1;2;3;4},3.9,1)' },
  @{ 訳 = '(紙)percentrank'; 式 = '=PERCENTRANK({1;2;3;4},3.8,2)' },
  @{ 訳 = '(紙)percentrank'; 式 = '=PERCENTRANK({1;2;3;4},3.9,2)' },
  @{ 訳 = '(紙)percentrank'; 式 = '=PERCENTRANK(C1:C5,4,0)' },
  @{ 訳 = '(紙)percentrank'; 式 = '=PERCENTRANK(C1:C5,4,-1)' },
  @{ 訳 = '(紙)percentrank'; 式 = '=PERCENTRANK(C1:C5,0)' },
  @{ 訳 = '(紙)percentrank'; 式 = '=PERCENTRANK({1;3;3;5},3)' },
  @{ 訳 = '(紙)percentrank'; 式 = '=PERCENTRANK({5;3;1},3)' },
  @{ 訳 = '(紙)percentrank'; 式 = '=PERCENTRANK({1;"a";3},2)' },
  @{ 訳 = '(紙)percentrank'; 式 = '=PERCENTRANK({},1)' },
  @{ 訳 = '(紙)asc-dbcs'; 式 = '=UNICODE(ASC(UNICHAR(12539)))' },
  @{ 訳 = '(紙)asc-dbcs'; 式 = '=UNICODE(ASC(UNICHAR(12443)))' },
  @{ 訳 = '(紙)asc-dbcs'; 式 = '=LEN(ASC(UNICHAR(12443)))' },
  @{ 訳 = '(紙)asc-dbcs'; 式 = '=UNICODE(MID(ASC(UNICHAR(12460)),2,1))' },
  @{ 訳 = '(紙)asc-dbcs'; 式 = '=UNICODE(ASC(UNICHAR(12497)))' },
  @{ 訳 = '(紙)asc-dbcs'; 式 = '=LEN(ASC(UNICHAR(12497)))' },
  @{ 訳 = '(紙)asc-dbcs'; 式 = '=UNICODE(ASC(UNICHAR(12540)))' },
  @{ 訳 = '(紙)asc-dbcs'; 式 = '=UNICODE(ASC(UNICHAR(12532)))' },
  @{ 訳 = '(紙)asc-dbcs'; 式 = '=LEN(ASC(UNICHAR(12532)))' },
  @{ 訳 = '(紙)asc-dbcs'; 式 = '=UNICODE(ASC(UNICHAR(12354)))' },
  @{ 訳 = '(紙)asc-dbcs'; 式 = '=UNICODE(ASC(UNICHAR(65281)))' },
  @{ 訳 = '(紙)asc-dbcs'; 式 = '=UNICODE(ASC(UNICHAR(65509)))' },
  @{ 訳 = '(紙)asc-dbcs'; 式 = '=UNICODE(DBCS(UNICHAR(92)))' },
  @{ 訳 = '(紙)asc-dbcs'; 式 = '=UNICODE(DBCS(UNICHAR(126)))' },
  @{ 訳 = '(紙)lenb'; 式 = '=LENB(UNICHAR(26085))' },
  @{ 訳 = '(紙)lenb'; 式 = '=LENB(UNICHAR(12354))' },
  @{ 訳 = '(紙)lenb'; 式 = '=LENB(UNICHAR(65281))' },
  @{ 訳 = '(紙)lenb'; 式 = '=LENB(UNICHAR(8212))' },
  @{ 訳 = '(紙)lenb'; 式 = '=LENB(UNICHAR(9731))' },
  @{ 訳 = '(紙)lenb'; 式 = '=UNICODE(MIDB(UNICHAR(12354)&"A",2,2))' },
  @{ 訳 = '(紙)lenb'; 式 = '=UNICODE(MIDB(UNICHAR(12354)&"A",2,1))' },
  @{ 訳 = '(紙)lenb'; 式 = '=LEN(MIDB(UNICHAR(12354)&"A",2,1))' },
  @{ 訳 = '(紙)lenb'; 式 = '=LEFTB("ABC")' },
  @{ 訳 = '(紙)lenb'; 式 = '=LEFTB("ABC",-1)' },
  @{ 訳 = '(紙)lenb'; 式 = '=MIDB("ABC",9,2)' },
  @{ 訳 = '(紙)lenb'; 式 = '=LENB(TRUE)' },
  @{ 訳 = '(紙)lenb'; 式 = '=LENB(123)' },
  @{ 訳 = '(紙)lenb'; 式 = '=RIGHTB(UNICHAR(12354),1)' },
  @{ 訳 = '(紙)textafter'; 式 = '=TEXTAFTER("abc","x",1,0,1)' },
  @{ 訳 = '(紙)textafter'; 式 = '=TEXTBEFORE("abc","x",1,0,1)' },
  @{ 訳 = '(紙)textafter'; 式 = '=TEXTAFTER("abc","x",1,0,0,UNICHAR(12394)&UNICHAR(12375))' },
  @{ 訳 = '(紙)textafter'; 式 = '=TEXTBEFORE("abc","x",1,0,0,UNICHAR(12394)&UNICHAR(12375))' },
  @{ 訳 = '(紙)textafter'; 式 = '=TEXTAFTER("a-b","-",9,0,0,UNICHAR(12394)&UNICHAR(12375))' },
  @{ 訳 = '(紙)textafter'; 式 = '=TEXTAFTER("abc","")' },
  @{ 訳 = '(紙)textafter'; 式 = '=TEXTAFTER("a-b","-",0)' },
  @{ 訳 = '(紙)textafter'; 式 = '=TEXTAFTER("aaa","aa",2)' },
  @{ 訳 = '(紙)textafter'; 式 = '=TEXTBEFORE("aaa","aa",2)' },
  @{ 訳 = '(紙)textafter'; 式 = '=TEXTAFTER("","x")' },
  @{ 訳 = '(紙)textafter'; 式 = '=TEXTAFTER(123,"2")' },
  @{ 訳 = '(紙)textafter'; 式 = '=TEXTAFTER(TRUE,"R")' },
  @{ 訳 = '(紙)textafter'; 式 = '=TEXTAFTER("a-b-c",{"-";"b"})' },
  @{ 訳 = '(紙)textafter'; 式 = '=TEXTAFTER("a-B-c","B")' },
  @{ 訳 = '(紙)aggregate'; 式 = '=AGGREGATE(9,1,B1:B5)' },
  @{ 訳 = '(紙)aggregate'; 式 = '=AGGREGATE(9,3,B1:B5)' },
  @{ 訳 = '(紙)aggregate'; 式 = '=AGGREGATE(9,5,B1:B5)' },
  @{ 訳 = '(紙)aggregate'; 式 = '=AGGREGATE(9,1,A1:A5)' },
  @{ 訳 = '(紙)aggregate'; 式 = '=AGGREGATE(9,4,A1:A5)' },
  @{ 訳 = '(紙)aggregate'; 式 = '=AGGREGATE(9,5,A1:A5)' },
  @{ 訳 = '(紙)aggregate'; 式 = '=AGGREGATE(9,7,A1:A5)' },
  @{ 訳 = '(紙)aggregate'; 式 = '=AGGREGATE(9,0,A1:A5,AGGREGATE(9,0,A1:A5))' },
  @{ 訳 = '(紙)aggregate'; 式 = '=AGGREGATE(9,0,A1:A2)' },
  @{ 訳 = '(紙)aggregate'; 式 = '=AGGREGATE(9,0,A1:C5)' },
  @{ 訳 = '(紙)aggregate'; 式 = '=AGGREGATE(0,6,A1:A5)' },
  @{ 訳 = '(紙)aggregate'; 式 = '=AGGREGATE(20,6,A1:A5)' },
  @{ 訳 = '(紙)aggregate'; 式 = '=AGGREGATE(9,8,A1:A5)' },
  @{ 訳 = '(紙)aggregate'; 式 = '=AGGREGATE(14,6,B1:B5,99)' },
  @{ 訳 = '(紙)aggregate'; 式 = '=AGGREGATE(9,6,A1:A5,B1:B5,A1:A5)' },
  @{ 訳 = '(紙)aggregate'; 式 = '=AGGREGATE(3,6,B1:B5)' }
)

# ══ ★ODDL の 割れ目 18本★ ══（`oddl-rokuwakume-no-an.md`）
#   ★取り出し方★ 利率0 ⇒ DSC が 直に ／ 利回り0 ⇒ DC−A が 直に
#   ★組3★ f=2 満 2009-05-31 ／ ★組4★ f=4 満 2009-07-15 ／ ★組6★ f=2 満 2010-05-31
$ODDL組 = @(
  @{ 名 = '組3(f2 はみ1日 NC1)'; f = 2; 最 = 'DATE(2008,11,30)'; 満 = 'DATE(2009,5,31)'; 決 = 'DATE(2009,1,15)' },
  @{ 名 = '組4(f4 はみ48日 NC2)'; f = 4; 最 = 'DATE(2008,11,30)'; 満 = 'DATE(2009,7,15)'; 決 = 'DATE(2009,1,15)' },
  @{ 名 = '組6(f2 はみ1日 NC3)'; f = 2; 最 = 'DATE(2008,11,30)'; 満 = 'DATE(2010,5,31)'; 決 = 'DATE(2009,1,15)' }
)

$式たち = New-Object System.Collections.Generic.List[object]
$足す = { param($訳, $式) $式たち.Add([pscustomobject]@{ 訳 = $訳; 式 = $式 }) }

foreach ($x in $紙の式) { & $足す $x.訳 $x.式 }

foreach ($g in $ODDL組) {
  foreach ($b in 0, 1, 4) {
    & $足す ($g.名 + ' (ア)DSC 利率0 basis=' + $b) ('=ODDLPRICE(' + $g.決 + ',' + $g.満 + ',' + $g.最 + ',0,0.05,100,' + $g.f + ',' + $b + ')')
    & $足す ($g.名 + ' (イ)DC-A 利回り0 basis=' + $b) ('=ODDLPRICE(' + $g.決 + ',' + $g.満 + ',' + $g.最 + ',0.06,0,100,' + $g.f + ',' + $b + ')')
  }
}

# ══ ★対照 3本★（★合わなければ そこで 止める★）══
& $足す '対照1(PERMUT・紙に在る)' '=PERMUT(5,2)'
& $足す '対照2(ODDL f4・紙に在る)' '=ODDLPRICE(DATE(2009,3,10),DATE(2009,8,31),DATE(2008,11,30),0.045,0.05,100,4,1)'
& $足す '対照3(ODDF・前の枠と同じ)' '=ODDFPRICE(DATE(2009,3,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2010,1,1),0.06,0.05,100,2,1)'

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

# ══ ★⑩式に ASCII の 外の 字が 無いか★ ══
#   ★なぜ★ … 日本語の 字は `.ps1` → PowerShell → COM → 実 Excel の 間で ★化ける 所が 在ります★
#   ⇒★化けた 時、答えが 違うのか 字が 化けたのかが ★見分けられません★★
#   ⇒★だから 字は `UNICHAR(番号)` で 作ります★
#   ★2026-09-18 に この 門が ★3本 見つけました★★（textafter の `"なし"`）
$外の字 = New-Object System.Collections.Generic.List[string]
foreach ($x in $式たち) {
  foreach ($c in $x.式.ToCharArray()) {
    if ([int][char]$c -gt 126) { $外の字.Add($x.訳 + ' … ' + $x.式); break }
  }
}
Write-Host ('★式に ASCII の 外の 字 … ' + $外の字.Count + '件★（0で あるべき）')
if ($外の字.Count -ne 0) {
  Write-Host '★★式に ASCII の 外の 字が 在ります＝`UNICHAR(番号)` に 直して ください★★'
  foreach ($s in $外の字) { Write-Host ('    ' + $s) }
  exit 6
}

# ══ ★⑬集め漏れの 門★ ══（2026-09-18・★指示役1 の 注文★）
#   ★訳★ … ★機械で 集めた から こそ、途中で 1紙 落ちても 気づけません★
#          ＝「手で 写して いない」は 良い／★集め漏れは 別の 穴★
#   ⇒★★3つの 数を 並べます★★
#       ㋐集めた 時 … `toi-atsumeta.txt` の 行数（★今 その場で 読みます★）
#       ㋑走った 時 … この 道具が 持って いる 紙の 式の 本数
#       ㋒紙に 書かれた … `hachiwakume-no-an.md` の 決め打ち（下の $紙の本数）
$紙の本数 = 94
$紙ごとの本数 = @{ 'junretsu' = 12; 'xmatch' = 12; 'percentrank' = 12; 'asc-dbcs' = 14; 'lenb' = 14; 'textafter' = 14; 'aggregate' = 16 }

$集めた紙 = Join-Path $ここ 'kansuu46/toi-atsumeta.txt'
$集めた時 = -1
if (Test-Path $集めた紙) {
  $集めた時 = @(Get-Content -LiteralPath $集めた紙 -Encoding UTF8 | Where-Object { $_.Trim() -ne '' -and -not $_.StartsWith('#') }).Count
} else {
  Write-Host ('★★集めた 紙が 在りません … ' + $集めた紙 + '★★（★数を 並べられません★）')
}
$走った時 = $紙の式.Count

Write-Host ''
Write-Host '★★集め漏れを 数える（3つを 並べる）★★'
Write-Host ('  ㋐集めた 時（toi-atsumeta.txt の 行数） … ' + $(if ($集めた時 -lt 0) { '★読めません★' } else { [string]$集めた時 + '本' }))
Write-Host ('  ㋑走った 時（この 道具が 持って いる） … ' + $走った時 + '本')
Write-Host ('  ㋒紙に 書かれた（見込みの 紙の 決め打ち） … ' + $紙の本数 + '本')

# ★紙ごとにも 並べる★（★1紙 落ちても 合計が 合う 事が 在ります★）
$紙ごと = @{}
foreach ($x in $紙の式) {
  $k = $x.訳 -replace '^\(紙\)', ''
  if ($紙ごと.ContainsKey($k)) { $紙ごと[$k] = $紙ごと[$k] + 1 } else { $紙ごと[$k] = 1 }
}
$紙の違い = New-Object System.Collections.Generic.List[string]
foreach ($k in ($紙ごとの本数.Keys | Sort-Object)) {
  $いま = 0
  if ($紙ごと.ContainsKey($k)) { $いま = $紙ごと[$k] }
  $しるし = '○'
  if ($いま -ne $紙ごとの本数[$k]) { $しるし = '★違う★'; $紙の違い.Add($k + ' … ' + $いま + '本（はず ' + $紙ごとの本数[$k] + '本）') }
  Write-Host ('    ' + $k.PadRight(14) + ' ' + ([string]$いま).PadLeft(3) + '本（はず ' + $紙ごとの本数[$k] + '） ' + $しるし)
}
foreach ($k in ($紙ごと.Keys | Sort-Object)) {
  if (-not $紙ごとの本数.ContainsKey($k)) { $紙の違い.Add('★知らない 紙★ ' + $k + ' … ' + $紙ごと[$k] + '本') }
}
# ★★読めないのも 赤★★ … ★測る 道具が 返した「読めません」を 見逃すと 2つしか 並びません★
if ($集めた時 -lt 0) { $紙の違い.Add('★㋐が 読めません★ … ' + $集めた紙) }
elseif ($集めた時 -ne $走った時) { $紙の違い.Add('㋐' + $集めた時 + '本 と ㋑' + $走った時 + '本 が 違う') }
if ($走った時 -ne $紙の本数) { $紙の違い.Add('㋑' + $走った時 + '本 と ㋒' + $紙の本数 + '本 が 違う') }
if ($紙の違い.Count -ne 0) {
  Write-Host '★★集め漏れが 在ります★★'
  foreach ($s in $紙の違い) { Write-Host ('    ' + $s) }
  exit 7
}
Write-Host '  ★3つとも 同じ＝集め漏れ 0★'
Write-Host ''

# ══ ★②本数の 門★ ══
$式の本数 = 115
Write-Host ('★聞く 式 … ' + $式たち.Count + '本★（決め打ち ' + $式の本数 + '本）')
if ($式たち.Count -ne $式の本数) {
  Write-Host ('★★' + $式の本数 + '本の はずが ' + $式たち.Count + '本です★★／★変えたなら この 数も 直し、便りにも 同じ 数を 書いて ください★')
  exit 4
}

$xl = New-Object -ComObject Excel.Application
$bk = $null
$sh = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)

  # ★材料★（★紙の 材料と 同じ★）
  #   A1:A5 = 1,2,3,4,5 ／ B1:B5 = 1,2,(#DIV/0!),4,5 ／ C1:C5 = 1,3,5,7,9 ／ D1:D5 = 9,7,5,3,1
  for ($i = 1; $i -le 5; $i++) { $sh.Range('A' + $i).Value2 = $i }
  $sh.Range('B1').Value2 = 1
  $sh.Range('B2').Value2 = 2
  $sh.Range('B3').Formula = '=1/0'
  $sh.Range('B4').Value2 = 4
  $sh.Range('B5').Value2 = 5
  $C = 1, 3, 5, 7, 9
  $D = 9, 7, 5, 3, 1
  for ($i = 1; $i -le 5; $i++) {
    $sh.Range('C' + $i).Value2 = $C[$i - 1]
    $sh.Range('D' + $i).Value2 = $D[$i - 1]
  }

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★溜まって いる 問いを 1枠に★（2026-09-18・8回目）')
  $行.Add('#')
  $行.Add('# ★94本は 7つの 紙から ★機械で 集めました★★（toi-atsumeru.py）')
  $行.Add('# ＋ODDL の 割れ目 18本 ＋ 対照 3本 ＝ ★115本★')
  $行.Add('#')
  $行.Add('# ★材料★ A1:A5 = 1,2,3,4,5 ／ B1:B5 = 1,2,(=1/0),4,5')
  $行.Add('#        C1:C5 = 1,3,5,7,9（昇順）／ D1:D5 = 9,7,5,3,1（降順）')
  $行.Add('#')
  $行.Add('# ★字は UNICHAR で 作って あります★（式の 中は ASCII だけ）')
  $行.Add('#')
  $行.Add('# ★2つ目の 窓★ =(式)=0 … .Value2 は 0 で ない 値に 0 を 返す')
  $行.Add('#')
  $行.Add('# ★外へ 出る 6個★ … ★0件★（門で 数えました）')
  $行.Add('#')
  $行.Add('# ★どの Excel か★ … 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('#')
  # ★★測る 道具の 版も 紙に 書く★★（2026-09-18）
  #   ★訳★ … ★版が 変わると ★同じ 数が 違う 字に なります★★（今日 実際に 起きた）
  $行.Add('# ★どの 貝殻か★ … PowerShell ' + $版.ToString() + '（★5.1 に 揃える 決め★）')
  $行.Add('#')
  # ★★LENB は この パソコンの 文字コードに 縛られます★★（2026-09-18 実測）
  #   LENB(UNICHAR(8212)) ＝ 1 ／ LENB(UNICHAR(9731)) ＝ 1
  #   ＝★CP932 に 無い 字は 1バイト★ ⇒★別の パソコンでは 変わり得ます★
  $行.Add('# ★どの 文字コードか★ … ANSI ' + [System.Text.Encoding]::Default.CodePage + ' ／ OEM ' + [System.Globalization.CultureInfo]::CurrentCulture.Name + '（★LENB/LEFTB/MIDB/RIGHTB は これに 縛られます★）')
  $行.Add('#')
  $行.Add('# 訳' + "`t" + '式' + "`t" + '答え' + "`t" + '出る字' + "`t" + '=(式)=0' + "`t" + '型')

  # ★⑧押した 時間を 別に 測る★
  $押し時計 = [Diagnostics.Stopwatch]::StartNew()
  $r = 1
  foreach ($x in $式たち) {
    $c = $sh.Range('H' + $r)
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
    $w = $sh.Range('I' + $r)
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
  # ══ ★★新しい 型（2026-09-18 に 測って 決着）★★ ══
  #   ★★2026-09-18 に 1つ 足しました … `$c` と `$w` も $null★★
  #   ★訳（2 × 2 で 測りました・1回ずつ）★
  #        貝殻       掴んだ まま   離す（$c/$w も null）
  #        5.1        60.16秒 残    ★★6.55秒で 消えた★★
  #        7.6        60.22秒 残    60.21秒 残
  #   ⇒★★5.1 かつ 離す の 時だけ 消えます＝★両方 要ります★★★
  #   ⇒★8枠目が 120秒 消えなかったのは ★この 2つが どちらも 欠けて いた★から★
  $c = $null
  $w = $null
  #   ★効くのは これだけ★ … $sh / $bk / $xl を $null に する
  #   ★★ReleaseComObject を 掴んだ 物ごと／GC::Collect は ★足さない★★★
  #     ＝足すと ★2.71秒 → 60秒 超★ に なります（交互 2回ずつで 決着）
  $sh = $null
  $bk = $null
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $xl = $null
  $t = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 120)) {
    Start-Sleep -Milliseconds 250
  }
  $t.Stop()
  $残り = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
  # ★★「待った」と「消えた」は 別★★
  if ($残り -eq 0) {
    Write-Host ('★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒で ★消えました★★')
  } else {
    Write-Host ('★★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒 待っても ★消えませんでした★★ ／ 残り ' + $残り + '個')
  }
  Write-Host ''
  Write-Host '★★秒が 毎回 同じなら ★上限に 当てて いる★ か ★空振り★ を 疑う★★'
}
