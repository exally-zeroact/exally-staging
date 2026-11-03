# tsukuru-jitsu-excel-meibo-ni-nai-11.ps1
#   -- ★うちの 名簿（XLFN）に 無い 11個を 実Excel に 書かせて 印を 見る★（52）（2026-09-20）
#
#  ★★なぜ★★
#    `lib/xlsx-io.js` の `XLFN`（130本）に ★無い 物が 11個★ 在りました:
#      WRAPROWS / WRAPCOLS / MAKEARRAY / BYROW / BYCOL / SCAN / MAP / REDUCE
#      GROUPBY / PIVOTBY / TRANSPOSE
#    ★名簿に 無い＝書き出す 時 `_xlfn.` を 付けない★
#    ⇒★実Excel が `_xlfn.` を 要る と するなら お客さんの ファイルは `#NAME?` に なります★
#
#  ★★11個の うち 8個は うちが 計算できます★★（実装を 探して 当たりが 在った）
#    WRAPROWS 10件／WRAPCOLS 9件／MAKEARRAY 16件／BYROW 14件／BYCOL 14件
#    SCAN 16件／MAP 18件／REDUCE 17件／GROUPBY 1件／PIVOTBY 1件／TRANSPOSE 2件
#    ⇒★お客さんが 打てる＝書き出しも 通る＝そこで 壊れる かも しれません★
#
#  ★★但し 先に 分かって いる 事★★
#    ・`TRANSPOSE` は ★実Excel も 裸で 書きます★（51 で 実測）⇒★名簿に 無いのが 正しい★
#    ・`MAKEARRAY` `BYROW` `BYCOL` `SCAN` `MAP` `REDUCE` は ★LAMBDA が 要ります★
#      ⇒`lib/xlsx-io.js` の `NEEDS_XLPM` が ★書き出しを 止めます★（壊れた物を 作らない）
#      ⇒★止まるなら 実害は 「出せない」で あって 「壊れる」では ない★
#    ・★`WRAPROWS` `WRAPCOLS` `GROUPBY` `PIVOTBY` は LAMBDA が 要りません★
#      ⇒★止まらずに 裸で 書かれる＝ここが 一番 危ない★
#
#  ★★この 道具が する 事★★
#    実Excel に 11個を 打たせて ★保存させる★。印（`_xlfn.` 等）は ★別の 道具で 生の 字から 読みます★。
#    ★この 版の Excel に 無い 式は `#NAME?` に なります★
#      ⇒★その 行は 「印が 分からない」として 分けます★（★無い 物を 有る ように 書かない★）
#
#  ★★置き場★★ `%TEMP%\exally-jitsu-excel-meibo11.xlsx`（★1本の 名★）
#    ・★司さんの 実物の 名は 1文字も 在りません★
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③台本の 数 決め打ち（exit 4）／④外へ 出る 6個と ASCII の 外の 字が 0件（exit 5）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$出す先 = Join-Path $env:TEMP 'exally-jitsu-excel-meibo11.xlsx'
$許す名 = 'exally-jitsu-excel-meibo11.xlsx'
if ((Split-Path $出す先 -Leaf) -ne $許す名) { Write-Host '★★書いて よい 名は 1本だけです★★'; exit 7 }

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

# ══ ★台本★（★1つずつ 別の 行に 置く＝ぶつからない★）══
#   材料 A1:A6 = 1..6（2行3列に 巻ける 数）
$台本 = @(
  @{ 頭 = 'C1';  札 = 'WRAPROWS';  式 = '=WRAPROWS(A1:A6,3)' },
  @{ 頭 = 'C5';  札 = 'WRAPCOLS';  式 = '=WRAPCOLS(A1:A6,3)' },
  @{ 頭 = 'C9';  札 = 'TRANSPOSE'; 式 = '=TRANSPOSE(A1:A3)' },
  @{ 頭 = 'C12'; 札 = 'MAKEARRAY'; 式 = '=MAKEARRAY(2,2,LAMBDA(r,c,r*c))' },
  @{ 頭 = 'C16'; 札 = 'BYROW';     式 = '=BYROW(A1:A3,LAMBDA(x,SUM(x)))' },
  @{ 頭 = 'C20'; 札 = 'BYCOL';     式 = '=BYCOL(A1:A3,LAMBDA(x,SUM(x)))' },
  @{ 頭 = 'C24'; 札 = 'SCAN';      式 = '=SCAN(0,A1:A3,LAMBDA(a,b,a+b))' },
  @{ 頭 = 'C28'; 札 = 'MAP';       式 = '=MAP(A1:A3,LAMBDA(x,x*2))' },
  @{ 頭 = 'C32'; 札 = 'REDUCE';    式 = '=REDUCE(0,A1:A3,LAMBDA(a,b,a+b))' },
  @{ 頭 = 'C36'; 札 = 'GROUPBY';   式 = '=GROUPBY(A1:A3,A1:A3,SUM)' },
  @{ 頭 = 'C40'; 札 = 'PIVOTBY';   式 = '=PIVOTBY(A1:A3,A1:A3,A1:A3,SUM)' }
)
$台本の数 = 11
Write-Host ('★台本 ... ' + $台本.Count + '本★（決め打ち ' + $台本の数 + '本）')
if ($台本.Count -ne $台本の数) { exit 4 }

$外へ出る = 'WEBSERVICE', 'STOCKHISTORY', 'TRANSLATE', 'DETECTLANGUAGE', 'IMAGE', 'RTD'
$見つけた = 0
foreach ($x in $台本) {
  foreach ($n in $外へ出る) { if ($x.式.ToUpper().Contains($n)) { $見つけた++ } }
  foreach ($ch in $x.式.ToCharArray()) { if ([int]$ch -gt 127) { $見つけた++ } }
}
Write-Host ('★外へ 出る 6個 ＋ ASCII の 外の 字 ... ' + $見つけた + '件★')
if ($見つけた -ne 0) { exit 5 }

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null; $c = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)
  for ($r = 1; $r -le 6; $r++) { $sh.Cells.Item($r, 1).Value2 = $r }

  Write-Host ''
  Write-Host '★★打った 結果★★（★#NAME? なら この版に 無い＝印は 分かりません★）'
  foreach ($x in $台本) {
    $投げた = ''
    try { $sh.Range($x.頭).Formula2 = $x.式 } catch { $投げた = '★投げました★ ' + $_.Exception.Message }
    $c = $sh.Range($x.頭)
    $v = $c.Value2
    $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    $出し = '  ' + $x.札.PadRight(10) + ' ' + $x.頭.PadRight(4) + ' 値 ' + $値.PadRight(14) + ' 字 ' + ([string]$c.Text).PadRight(10) + ' ' + $投げた
    Write-Host $出し
  }

  if (Test-Path $出す先) { Remove-Item $出す先 -Force }
  $bk.SaveAs($出す先, 51)
  $bk.Close($false)
  $bk = $null
} finally {
  # ★★掴んだ物 全部 $null★★（★これを 1つ 忘れると 120秒 消えません／今日 2回 出ました★）
  $c = $null
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
if (Test-Path $出す先) {
  $x2 = Get-Item $出す先
  $h = (Get-FileHash $出す先 -Algorithm SHA256).Hash.ToLower()
  Write-Host ''
  Write-Host ('★★作りました★★ ... ' + $出す先)
  Write-Host ('  ★大きさ★ ' + $x2.Length + ' バイト ／ ★sha256★ ' + $h)
  Write-Host '  ⇒★印（_xlfn. 等）は 生の 字から 読みます★'
}
