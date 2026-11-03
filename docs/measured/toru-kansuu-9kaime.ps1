# toru-kansuu-9kaime.ps1 - ★9枠目＝溜まった 問いを 1枠に★（2026-09-18）
#
#  ★★中身は 紙から 機械で 拾いました★★（★手で 書き写して いません★）
#    紙 ... docs/measured/kansuu46/kyuwakume-no-kiku-koto.md
#    ⇒★★見込みも 一緒に 持って います★★＝★出しに 並べます★
#    ⇒★後から 答えに 寄せられません★
#
#  ★★㋕（BYROW／BYCOL を ★分かれる 材料★で）を 一番 上に 置きました★★
#    ★訳★ ... 8枠目までの 材料（A1:B2 ＝ 1,2／2,4）では
#            ★行ごとも 列ごとも 3★＝★見分けられません★
#    ⇒★★F1:G2 ＝ 1,2／10,20 を 足しました★★
#        行の 和 ... 3 と 30 ／ 列の 和 ... 11 と 22
#        ⇒★和は どちらも 33★＝★★和では 分かれません★★
#        ⇒★★左上で 分かれます★★（行なら 3 ／ 列なら 11）
#
#  ★★新しい 型（2026-09-18 に 2×2 で 決着）★★
#    ★終わりは `$c/$w/$sh/$bk/$xl` を 全部 $null★
#    ★`ReleaseComObject` を 物ごと／`GC::Collect` は ★足さない★★
#    ★実測★ 5.1 × 離す ... ★6.55秒★／5.1 × 掴んだまま ... 60秒超
#            7.6 × 離す ... 60秒超    ／7.6 × 掴んだまま ... 60秒超
#    ⇒★★2つ 揃った 時だけ 消えます★★
#
#  ★★門（12個）★★（★門は 引き継がれません＝頭に 名指しで 書きます★）
#    ①字の 誤り 0件 ②本数 決め打ち ③Excel を 2つの 道具で（exit 3）
#    ④1本ずつ 受け止め ⑤司さんの ブックを 開く 字 0件
#    ⑥BOM ⑦2つ目の 窓＋型 ⑧押した秒と 消え秒を 別々
#    ⑨外へ 出る 6個が 0件（exit 5）
#    ⑩式に ASCII の 外の 字が 0件（exit 6）
#    ⑬集め漏れ（紙の 表と 数が 合うか・exit 7）
#    ⑭★貝殻が `powershell.exe`（5.1）★（exit 8）
#
#  使い方: powershell.exe -NoProfile -File docs/measured/toru-kansuu-9kaime.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-kansuu-9kaime-2026-09-18.tsv'

# ══ ★⑭貝殻の 版★ ══
$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) {
  Write-Host '★★この 道具は `powershell.exe`（5.1）で 走らせて ください★★'
  Write-Host '  ★訳★ ... 5.1 と 7.x は ★同じ 数を 違う 字で 書きます★'
  Write-Host '         5.1 ... 99.776476358001872 ／ 7.x ... 99.77647635800187'
  Write-Host '  ⇒★34本の 古い 紙は 5.1 で 取って 在ります★'
  exit 8
}

# ══ ★③Excel を 2つの 道具で★ ══
$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... Get-Process ' + $数1 + '個 ／ Win32_Process ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

# ══ ★紙から 拾った 問い★ ══
$紙の式 = @(
  @{ 番 = '49'; 訳 = '★★行か 列か（F1:G2 ＝ 1,2／10,20）★★'; 式 = '=BYROW(F1:G2,LAMBDA(r,SUM(r)))'; 見込み = '★3★' },
  @{ 番 = '50'; 訳 = '★行なら 3+30★'; 式 = '=SUM(BYROW(F1:G2,LAMBDA(r,SUM(r))))'; 見込み = '★33★' },
  @{ 番 = '51'; 訳 = '★列なら 11+22＝33★ ← ★★和では 分かれない★★'; 式 = '=SUM(BYCOL(F1:G2,LAMBDA(c,SUM(c))))'; 見込み = '★33★' },
  @{ 番 = '52'; 訳 = '★★左上で 分かれます★★（行なら 3）'; 式 = '=BYCOL(F1:G2,LAMBDA(c,SUM(c)))'; 見込み = '★11★' },
  @{ 番 = '1'; 訳 = '正方で ない'; 式 = '=MDETERM(A1:B5)'; 見込み = '#VALUE!' },
  @{ 番 = '2'; 訳 = '1×1 は その 値'; 式 = '=MDETERM(A1:A1)'; 見込み = '1' },
  @{ 番 = '3'; 訳 = 'B3 では ない ... B1:C2 に 誤りは 無い ⇒ ★-1 の はず★'; 式 = '=MDETERM(B1:C2)'; 見込み = '#DIV/0!' },
  @{ 番 = '4'; 訳 = 'n が 4未満'; 式 = '=KURT(A1:A3)'; 見込み = '#DIV/0!' },
  @{ 番 = '5'; 訳 = 'ばらつき 0'; 式 = '=KURT({2;2;2;2})'; 見込み = '#DIV/0!' },
  @{ 番 = '6'; 訳 = '1つも 落ちない'; 式 = '=TRIMMEAN(A1:A5,0)'; 見込み = '3' },
  @{ 番 = '7'; 訳 = '1 は 入らない'; 式 = '=TRIMMEAN(A1:A5,1)'; 見込み = '#NUM!' },
  @{ 番 = '8'; 訳 = '負'; 式 = '=TRIMMEAN(A1:A5,-0.1)'; 見込み = '#NUM!' },
  @{ 番 = '9'; 訳 = '片側 1つずつ 落ちて 2,3,4'; 式 = '=TRIMMEAN(A1:A5,0.5)'; 見込み = '3' },
  @{ 番 = '10'; 訳 = '長さ違い'; 式 = '=FORECAST(6,A1:A5,C1:C3)'; 見込み = '#N/A' },
  @{ 番 = '11'; 訳 = 'x の ばらつき 0'; 式 = '=FORECAST(6,A1:A5,{1;1;1;1;1})'; 見込み = '#DIV/0!' },
  @{ 番 = '12'; 訳 = '小数の 印が 2回'; 式 = '=NUMBERVALUE("1.2.3")'; 見込み = '#VALUE!' },
  @{ 番 = '13'; 訳 = '印が 同じ 字'; 式 = '=NUMBERVALUE("1.5",".",".")'; 見込み = '#VALUE!' },
  @{ 番 = '14'; 訳 = '％を 重ねる'; 式 = '=NUMBERVALUE("50%%")'; 見込み = '0.005' },
  @{ 番 = '15'; 訳 = '前後の 空白'; 式 = '=NUMBERVALUE(" 1.5 ")'; 見込み = '1.5' },
  @{ 番 = '16'; 訳 = 'UTF-8 なら %E3'; 式 = '=UNICODE(MID(ENCODEURL(UNICHAR(12354)),2,1))'; 見込み = '69（"E"）' },
  @{ 番 = '17'; 訳 = '＋も 逃がす'; 式 = '=ENCODEURL("a+b")'; 見込み = '`a%2Bb`' },
  @{ 番 = '18'; 訳 = '負'; 式 = '=FIXED(-1234.567,2)'; 見込み = '`-1,234.57`' },
  @{ 番 = '19'; 訳 = '左へ 丸める'; 式 = '=FIXED(1234.567,-2)'; 見込み = '`1,200`' },
  @{ 番 = '20'; 訳 = '区切らない'; 式 = '=FIXED(1234.567,2,TRUE)'; 見込み = '`1234.57`' },
  @{ 番 = '21'; 訳 = '★★一番 大事★★ ... 丸括弧かも しれない'; 式 = '=DOLLAR(-1234.567,2)'; 見込み = '★`-¥1,234.57`★' },
  @{ 番 = '22'; 訳 = '左へ 丸める'; 式 = '=DOLLAR(1234.567,-2)'; 見込み = '`¥1,200`' },
  @{ 番 = '23'; 訳 = 'きっちり＝引用符'; 式 = '=VALUETOTEXT("a",1)'; 見込み = '`"a"`' },
  @{ 番 = '24'; 訳 = '真偽'; 式 = '=VALUETOTEXT(TRUE)'; 見込み = '`TRUE`' },
  @{ 番 = '25'; 訳 = '★誤りを 字に する★'; 式 = '=VALUETOTEXT(1/0)'; 見込み = '`#DIV/0!`' },
  @{ 番 = '26'; 訳 = '一番 小さい 物より 下'; 式 = '=LOOKUP(0,C1:C5,D1:D5)'; 見込み = '#N/A' },
  @{ 番 = '27'; 訳 = '一番 大きい 物'; 式 = '=LOOKUP(99,C1:C5,D1:D5)'; 見込み = '1' },
  @{ 番 = '28'; 訳 = '★四角の 形（1列）★'; 式 = '=LOOKUP(4,C1:C5)'; 見込み = '3' },
  @{ 番 = '29'; 訳 = '★昇順で ない＝決まって いない★'; 式 = '=LOOKUP(4,D1:D5)'; 見込み = '★分かりません★' },
  @{ 番 = '30'; 訳 = '★R1C1 の 書き方★'; 式 = '=INDIRECT("R1C1",FALSE)'; 見込み = '1' },
  @{ 番 = '31'; 訳 = '四角の 左上'; 式 = '=INDIRECT("A1:B2")'; 見込み = '1' },
  @{ 番 = '32'; 訳 = '高さ 0'; 式 = '=OFFSET(A1,0,0,0,1)'; 見込み = '#REF!' },
  @{ 番 = '33'; 訳 = '板の 外'; 式 = '=OFFSET(A1,-1,0)'; 見込み = '#REF!' },
  @{ 番 = '34'; 訳 = '負の 高さは 上へ'; 式 = '=SUM(OFFSET(A1,1,0,-2,1))'; 見込み = '3' },
  @{ 番 = '35'; 訳 = '★そのまま 打つと 何に なるか★'; 式 = '=LAMBDA(x,x*2)'; 見込み = '#CALC!' },
  @{ 番 = '36'; 訳 = '★`式(引数)` の 形★'; 式 = '=LAMBDA(x,x*2)(3)'; 見込み = '6' },
  @{ 番 = '37'; 訳 = '★名の台★'; 式 = '=LET(x,1,x+1)'; 見込み = '2' },
  @{ 番 = '38'; 訳 = '★★名の台と マスの どちらが 勝つか★★'; 式 = '=LET(A1,5,A1)'; 見込み = '★5★' },
  @{ 番 = '39'; 訳 = '★名を 2つ★'; 式 = '=LET(x,1,y,2,x+y)'; 見込み = '3' },
  @{ 番 = '40'; 訳 = '★最後は 式で ないと いけない★'; 式 = '=LET(x,1)'; 見込み = '#VALUE!' },
  @{ 番 = '41'; 訳 = '溢れの 左上'; 式 = '=MAP(A1:A5,LAMBDA(x,x*2))'; 見込み = '2' },
  @{ 番 = '42'; 訳 = '★溢れの 形★'; 式 = '=ROWS(MAP(A1:A5,LAMBDA(x,x*2)))'; 見込み = '5' },
  @{ 番 = '43'; 訳 = '★溢れの 形★'; 式 = '=COLUMNS(MAP(A1:A5,LAMBDA(x,x*2)))'; 見込み = '1' },
  @{ 番 = '44'; 訳 = 'まるごと'; 式 = '=SUM(MAP(A1:A5,LAMBDA(x,x*2)))'; 見込み = '30' },
  @{ 番 = '45'; 訳 = '★引数の 順★'; 式 = '=REDUCE(0,A1:A5,LAMBDA(a,b,a+b))'; 見込み = '15' },
  @{ 番 = '46'; 訳 = '★a が 溜め・b が 次★の 裏取り'; 式 = '=REDUCE(100,A1:A5,LAMBDA(a,b,a-b))'; 見込み = '85' },
  @{ 番 = '47'; 訳 = '★0 を 含むか＝5 か 6 か★'; 式 = '=ROWS(SCAN(0,A1:A5,LAMBDA(a,b,a+b)))'; 見込み = '★5★' },
  @{ 番 = '48'; 訳 = '1+3+6+10+15'; 式 = '=SUM(SCAN(0,A1:A5,LAMBDA(a,b,a+b)))'; 見込み = '35' },
  @{ 番 = '53'; 訳 = '左上'; 式 = '=MAKEARRAY(2,3,LAMBDA(r,c,r*c))'; 見込み = '1' },
  @{ 番 = '54'; 訳 = '(1+2)×(1+2+3)'; 式 = '=SUM(MAKEARRAY(2,3,LAMBDA(r,c,r*c)))'; 見込み = '18' },
  @{ 番 = '55'; 訳 = '★引数の 数違い★'; 式 = '=MAP(A1:A5,LAMBDA(x,y,x))'; 見込み = '#VALUE!' },
  @{ 番 = '56'; 訳 = '★引数の 数違い★'; 式 = '=REDUCE(0,A1:A5,LAMBDA(a,a))'; 見込み = '#VALUE!' },
  @{ 番 = '57'; 訳 = '★ラムダを 名に 入れて 呼ぶ★'; 式 = '=LET(f,LAMBDA(x,x*2),f(3))'; 見込み = '6' },
  @{ 番 = '58'; 訳 = '★★閉じ込め★★'; 式 = '=LET(a,2,SUM(MAP(A1:A5,LAMBDA(x,x*a))))'; 見込み = '30' }
)

$式たち = New-Object System.Collections.Generic.List[object]
foreach ($x in $紙の式) { $式たち.Add($x) }

# ══ ★対照 3本★ ══
$式たち.Add([pscustomobject]@{ 番 = '対1'; 訳 = '対照(PERMUT)'; 式 = '=PERMUT(5,2)'; 見込み = '20' })
$式たち.Add([pscustomobject]@{ 番 = '対2'; 訳 = '対照(ODDL f4)';
  式 = '=ODDLPRICE(DATE(2009,3,10),DATE(2009,8,31),DATE(2008,11,30),0.045,0.05,100,4,1)';
  見込み = '99.776476358001872' })
$式たち.Add([pscustomobject]@{ 番 = '対3'; 訳 = '対照(ODDF)';
  式 = '=ODDFPRICE(DATE(2009,3,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2010,1,1),0.06,0.05,100,2,1)';
  見込み = '103.37232293583249' })

# ══ ★⑨外へ 出る 6個★ ══
$外へ出る = 'WEBSERVICE', 'STOCKHISTORY', 'TRANSLATE', 'DETECTLANGUAGE', 'IMAGE', 'RTD'
$見つけた = New-Object System.Collections.Generic.List[string]
foreach ($x in $式たち) {
  foreach ($n in $外へ出る) {
    if ($x.式 -match ('(^|[^A-Z])' + $n + '\(')) { $見つけた.Add($n + ' ... ' + $x.訳) }
  }
}
Write-Host ('★外へ 出る 6個 ... ' + $見つけた.Count + '件★（0で あるべき）')
if ($見つけた.Count -ne 0) { foreach ($s in $見つけた) { Write-Host ('    ' + $s) }; exit 5 }

# ══ ★⑩式に ASCII の 外の 字★ ══
$外の字 = New-Object System.Collections.Generic.List[string]
foreach ($x in $式たち) {
  foreach ($c in $x.式.ToCharArray()) {
    if ([int][char]$c -gt 126) { $外の字.Add($x.訳 + ' ... ' + $x.式); break }
  }
}
Write-Host ('★式に ASCII の 外の 字 ... ' + $外の字.Count + '件★（0で あるべき）')
if ($外の字.Count -ne 0) { foreach ($s in $外の字) { Write-Host ('    ' + $s) }; exit 6 }

# ══ ★⑬集め漏れ（★紙と 突き合わせる★）★ ══
$紙の本数 = 58
Write-Host ''
Write-Host '★★集め漏れを 数える（3つを 並べる）★★'
$集めた紙 = Join-Path $ここ 'kansuu46/kyuwakume-no-kiku-koto.md'
$集めた時 = -1
if (Test-Path $集めた紙) {
  $集めた時 = 0
  foreach ($l in (Get-Content -LiteralPath $集めた紙 -Encoding UTF8)) {
    $s = $l.Trim()
    if (-not $s.StartsWith('|')) { continue }
    $c = $s.Trim('|').Split('|')
    if ($c.Count -lt 4) { continue }
    $ban = $c[0].Trim().Trim([char]0x2605)
    $shiki = $c[1].Trim().Trim([char]0x2605).Trim('`').Trim([char]0x2605).Trim('`')
    if ($ban -match '^[0-9]+$' -and $shiki.StartsWith('=')) { $集めた時 = $集めた時 + 1 }
  }
}
Write-Host ('  ㋐集めた 時（紙の 表の 行数） ... ' + $(if ($集めた時 -lt 0) { '★読めません★' } else { [string]$集めた時 + '本' }))
Write-Host ('  ㋑走った 時（この 道具が 持つ） ... ' + $紙の式.Count + '本')
Write-Host ('  ㋒紙に 書かれた（決め打ち） ...... ' + $紙の本数 + '本')
$違い = New-Object System.Collections.Generic.List[string]
if ($集めた時 -lt 0) { $違い.Add('★㋐が 読めません★ ... ' + $集めた紙) }
elseif ($集めた時 -ne $紙の式.Count) { $違い.Add('㋐' + $集めた時 + ' と ㋑' + $紙の式.Count + ' が 違う') }
if ($紙の式.Count -ne $紙の本数) { $違い.Add('㋑' + $紙の式.Count + ' と ㋒' + $紙の本数 + ' が 違う') }
if ($違い.Count -ne 0) {
  Write-Host '★★集め漏れが 在ります★★'
  foreach ($s in $違い) { Write-Host ('    ' + $s) }
  exit 7
}
Write-Host '  ★3つとも 同じ＝集め漏れ 0★'

# ══ ★②本数の 門★ ══
$式の本数 = 61
Write-Host ''
Write-Host ('★聞く 式 ... ' + $式たち.Count + '本★（決め打ち ' + $式の本数 + '本）')
if ($式たち.Count -ne $式の本数) {
  Write-Host ('★★' + $式の本数 + '本の はずが ' + $式たち.Count + '本です★★')
  exit 4
}

$xl = New-Object -ComObject Excel.Application
$bk = $null
$sh = $null
$c = $null
$w = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)

  # ★材料★（★8枠目と 同じ ＋ F1:G2★）
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
  # ★★行と 列が 分かれる 材料★★（2026-09-18 に 足した）
  $sh.Range('F1').Value2 = 1
  $sh.Range('G1').Value2 = 2
  $sh.Range('F2').Value2 = 10
  $sh.Range('G2').Value2 = 20

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★9枠目＝溜まった 問いを 1枠に★（2026-09-18）')
  $行.Add('#')
  $行.Add('# ★問いは 紙から 機械で 拾いました★ ... kansuu46/kyuwakume-no-kiku-koto.md')
  $行.Add('# ★見込みも 一緒に 並べて あります★（★聞く 前に 書いた 物★）')
  $行.Add('#')
  $行.Add('# ★材料★ A1:A5 = 1,2,3,4,5 ／ B1:B5 = 1,2,(=1/0),4,5')
  $行.Add('#        C1:C5 = 1,3,5,7,9（昇順）／ D1:D5 = 9,7,5,3,1（降順）')
  $行.Add('#        ★F1:G2 = 1,2／10,20★（★行ごと 3・30 ／ 列ごと 11・22＝分かれる★）')
  $行.Add('#')
  $行.Add('# ★2つ目の 窓★ =(式)=0 ... .Value2 は 0 で ない 値に 0 を 返す')
  $行.Add('# ★外へ 出る 6個★ ... ★0件★（門で 数えました）')
  $行.Add('#')
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★どの 貝殻か★ ... PowerShell ' + $版.ToString() + '（★5.1 に 揃える 決め★）')
  $コード = [string][System.Text.Encoding]::Default.CodePage
  $土地 = [System.Globalization.CultureInfo]::CurrentCulture.Name
  $行.Add('# ★どの 文字コードか★ ... ANSI ' + $コード + ' ／ ' + $土地)
  $行.Add('#')
  $行.Add('# 番' + "`t" + '訳' + "`t" + '式' + "`t" + '見込み' + "`t" + '答え' + "`t" + '出る字' + "`t" + '=(式)=0' + "`t" + '型' + "`t" + '合ったか')

  $押し時計 = [Diagnostics.Stopwatch]::StartNew()
  $r = 1
  $当たり = 0
  $外れ = 0
  foreach ($x in $式たち) {
    $c = $sh.Range('J' + $r)
    $打てた = $true
    try { $c.Formula = $x.式 } catch {
      $打てた = $false
      $行.Add($x.番 + "`t" + $x.訳 + "`t" + $x.式 + "`t" + $x.見込み + "`t" + '(★打てません★)' + "`t" + ('★Excel が 式を 受け付けません★ ' + $_.Exception.Message) + "`t" + '(★打てません★)' + "`t" + '(★打てません★)' + "`t" + '(★打てません★)')
      $r++
    }
    if (-not $打てた) { continue }
    $v = $c.Value2
    $答 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    $型 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
    $字 = [string]$c.Text
    $w = $sh.Range('K' + $r)
    $ゼロか = '(★窓2が 打てません★)'
    try { $w.Formula = '=(' + $x.式.Substring(1) + ')=0'; $ゼロか = [string]$w.Value2 } catch { }
    # ★見込みと 見比べる★（★字で 見比べます＝出る字 か 答え の どちらかと 合えば ○★）
    # ★★見込みと 見比べる★★
    #   ★答え（生の 値）か 出る字の どちらかと 合えば ○★
    #   ★これは ★目安★です★＝#DIV/0! の ような 誤りは 出る字と 見比べます
    $合 = '×'
    if (($x.見込み -eq $答) -or ($x.見込み -eq $字)) { $合 = '○' }
    $行.Add($x.番 + "`t" + $x.訳 + "`t" + $x.式 + "`t" + $x.見込み + "`t" + $答 + "`t" + $字 + "`t" + $ゼロか + "`t" + $型 + "`t" + $合)
    if ($合 -eq '○') { $当たり = $当たり + 1 } else { $外れ = $外れ + 1 }
    $r++
  }
  $押し時計.Stop()
  $押し秒 = [math]::Round($押し時計.Elapsed.TotalSeconds, 2)
  $行.Add('# ★★見込みが 当たった ' + $当たり + ' ／ 外れた ' + $外れ + '★★（★字で 見比べた 数＝目安★）')
  $行.Add('# ★★押すのに かかった 秒 ... ' + $押し秒 + '★★')

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★★見込みが 当たった ' + $当たり + ' ／ 外れた ' + $外れ + '★★（★目安★＝字で 見比べただけ）')
  Write-Host ('★★押すのに かかった 秒 ... ' + $押し秒 + '秒★★（' + $式たち.Count + '本）')
  Write-Host ('★書いた ... ' + $出 + '（' + $式たち.Count + '行）★')
  $bk.Close($false)
} finally {
  # ══ ★★新しい 型（2×2 で 決着）★★ ══
  $c = $null
  $w = $null
  $sh = $null
  $bk = $null
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $xl = $null
  # ★物ごとの Release と GC::Collect は 足さない★
  $t = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 120)) {
    Start-Sleep -Milliseconds 250
  }
  $t.Stop()
  $残り = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
  if ($残り -eq 0) {
    Write-Host ('★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒で ★消えました★★')
  } else {
    Write-Host ('★★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒 待っても ★消えませんでした★★ ／ 残り ' + $残り + '個')
  }
}
