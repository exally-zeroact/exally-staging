# tsukuru-jitsu-excel-nokori-21.ps1
#   -- ★分母から 漏れて いた 45個の 残り 21個を 実Excel に 書かせる★（55）（2026-09-20）
#
#  ★★なぜ★★
#    Exally1 が 数えました:
#      09-16 の 分母（`docs/measured/xlfn-namae.txt` 396個）に ★入って いなかった 45個★
#      ⇒その うち ★まだ 実Excel に 聞いて いない 27個★
#      ⇒その うち 6個（LAMBDA の 一族）は ★書き出しが 投げて 止まる＝実害 0★（実測）
#      ⇒★残り 21個★ ... ★21個 とも 台が 持って いる＝お客さんが 打てます★
#    ⇒★`WRAPROWS` と 同じ 顔が 他に 無いかを 割ります★
#
#  ★★私の 見立て（★測る 前に 書いて おきます★）★★
#    ・★古い 物が 多い＝`_xlfn.` は 要らない はず★（TRANSPOSE と 同じ）
#    ・★気に なるのは 2つ★（Exally1 と 同じ）
#        `MODE.MULT`   ... ★点が 入る／溢れる★
#                          （`MODE.SNGL` `RANK.EQ` `PERCENTILE.INC` は ★既に 名簿に 在る★）
#                          ⇒★`WRAPROWS` と 同じ 顔に 見える★
#        `PERCENTRANK` ... `PERCENTILE.INC` が 名簿に 在る のに こちらは 無い
#                          （★但し 点が 無い＝古い 形★なので 要らない かも）
#    ⇒★先に 書いて おく のは 「出た 数を 見てから そんなものかと 思う」のを 防ぐ 為★
#
#  ★★作る 物★★ `%TEMP%\exally-jitsu-excel-nokori21.xlsx`（★1本の 名★）
#    ★実Excel に 打たせて 保存させます★（★うちの 書き出しは 1回も 通しません★）
#
#  ★★材料と 置き場（★押す 前に 隙間を 計算しました★）★★
#    A1:A6 = 1 / 2 / 2 / 3 / 3 / 4   ★同じ 数を わざと 入れる★（`MODE.MULT` が 2つ 返る 為）
#    B1:B3 = 10 / 20 / 30            （`FORECAST` の もう 片方）
#    E1:F2 = 1 / 2 / 3 / 4           （`MDETERM` の 2x2）
#    ★式は H 列に 3行 おき★（H1 H4 H7 ...）＝★溢れても ぶつかりません★
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③台本の 数 決め打ち（exit 4）／④外へ 出る 6個と ASCII の 外の 字が 0件（exit 5）
#
#  ★★この 道具は 「作って 打った 結果を 見せる」だけ★★
#    ＝★印（`_xlfn.` 等）は `yomu-jitsu-excel-no-shirushi.mjs` で 生の 字から 読みます★
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$出す先 = Join-Path $env:TEMP 'exally-jitsu-excel-nokori21.xlsx'
$許す名 = 'exally-jitsu-excel-nokori21.xlsx'
if ((Split-Path $出す先 -Leaf) -ne $許す名) { Write-Host '★★書いて よい 名は 1本だけです★★'; exit 7 }

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$台本 = @(
  @{ 札 = 'ASC';         式 = '=ASC("ABC")' },
  @{ 札 = 'AVERAGEIFS';  式 = '=AVERAGEIFS(A1:A6,A1:A6,">2")' },
  @{ 札 = 'BAHTTEXT';    式 = '=BAHTTEXT(15)' },
  @{ 札 = 'DBCS';        式 = '=DBCS("ABC")' },
  @{ 札 = 'DOLLAR';      式 = '=DOLLAR(1234.567,2)' },
  @{ 札 = 'FIXED';       式 = '=FIXED(1234.567,1)' },
  @{ 札 = 'FORECAST';    式 = '=FORECAST(4,B1:B3,A1:A3)' },
  @{ 札 = 'INDIRECT';    式 = '=INDIRECT("A1")' },
  @{ 札 = 'KURT';        式 = '=KURT(A1:A6)' },
  @{ 札 = 'LEFTB';       式 = '=LEFTB("ABC",2)' },
  @{ 札 = 'LENB';        式 = '=LENB("ABC")' },
  @{ 札 = 'LOOKUP';      式 = '=LOOKUP(2,A1:A6)' },
  @{ 札 = 'MDETERM';     式 = '=MDETERM(E1:F2)' },
  @{ 札 = 'MIDB';        式 = '=MIDB("ABC",1,2)' },
  @{ 札 = 'MODE.MULT';   式 = '=MODE.MULT(A1:A6)' },
  @{ 札 = 'OFFSET';      式 = '=OFFSET(A1,1,0)' },
  @{ 札 = 'PERCENTRANK'; 式 = '=PERCENTRANK(A1:A6,3)' },
  @{ 札 = 'PERMUT';      式 = '=PERMUT(5,2)' },
  @{ 札 = 'PHONETIC';    式 = '=PHONETIC(A1)' },
  @{ 札 = 'RIGHTB';      式 = '=RIGHTB("ABC",2)' },
  @{ 札 = 'TRIMMEAN';    式 = '=TRIMMEAN(A1:A6,0.2)' }
)
$台本の数 = 21
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
  # ★材料★
  $材 = 1, 2, 2, 3, 3, 4
  for ($r = 1; $r -le 6; $r++) { $sh.Cells.Item($r, 1).Value2 = $材[$r - 1] }
  $sh.Range('B1').Value2 = 10
  $sh.Range('B2').Value2 = 20
  $sh.Range('B3').Value2 = 30
  $sh.Range('E1').Value2 = 1
  $sh.Range('F1').Value2 = 2
  $sh.Range('E2').Value2 = 3
  $sh.Range('F2').Value2 = 4

  Write-Host ''
  Write-Host '★★打った 結果★★（★#NAME? なら この版に 無い＝印は 分かりません★）'
  $行数 = 1
  foreach ($x in $台本) {
    $ma = 'H' + $行数
    $投げた = ''
    try { $sh.Range($ma).Formula2 = $x.式 } catch { $投げた = '★投げました★ ' + $_.Exception.Message }
    $c = $sh.Range($ma)
    $v = $c.Value2
    $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    $出し = '  ' + $x.札.PadRight(12) + ' ' + $ma.PadRight(4) + ' 値 ' + $値.PadRight(20) + ' 字 ' + ([string]$c.Text).PadRight(14) + ' ' + $投げた
    Write-Host $出し
    $行数 = $行数 + 3
  }

  if (Test-Path $出す先) { Remove-Item $出す先 -Force }
  $bk.SaveAs($出す先, 51)
  $bk.Close($false)
  $bk = $null
} finally {
  # ★★掴んだ物 全部 $null★★
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
  Write-Host '  ⇒★印は yomu-jitsu-excel-no-shirushi.mjs で 生の 字から 読みます★'
}
