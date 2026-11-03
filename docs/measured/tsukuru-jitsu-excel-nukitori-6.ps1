# tsukuru-jitsu-excel-nukitori-6.ps1
#   -- ★抜き取り 6個で 「まだ 聞いて いない 324個」を 割る★（59）（2026-09-20）
#
#  ★★なぜ 全部 聞かないのか★★
#    Exally1 が 数え直しました:
#      ★名簿に 在る のに 聞いて いない 110本★
#        ├ 台が 持たない 62本 ⇒★打てない＝書き出しに 出ない＝実害 0★
#        └ ★台が 持つ 47本★ ⇒★付け過ぎなら 壊れる★
#      ★台が 持つ のに 名簿にも 紙にも 無い 277個★ ⇒★付け忘れなら 表が 丸ごと 消える★
#      ⇒★聞く べきは 324個★
#    ⇒★全部 聞くのが 正しい★／★でも 先に 「どこを 見れば 良いか」を 決めます★
#
#  ★★抜き取り 6個（★両端を 押さえます★）★★
#    ㋐★FORMULATEXT★ ... B（台が 持つ／名簿に 無い）・★2013年★
#        ⇒★B 277個の 中で ★唯一 2010年 以降に 見える 物★★
#        ⇒★`_xlfn.` なら 4個目の 穴／裸なら 見立てが 割れる★＝★どちらでも 値打ちが 在る★
#    ㋑SUMPRODUCT ... B・古い ⇒★裸なら 見立て通り★
#    ㋒CONCATENATE . B・古い ⇒同上（★`CONCAT` は 名簿に 在る のに こちらは 無い★
#                              ＝★新旧が 分かれた 形＝`MODE.SNGL`/`MODE.MULT` と 同じ 顔★）
#    ㋓XLOOKUP ..... A（名簿に 在る／台が 持つ）・2019年 ⇒★付けて 正しい はず★
#    ㋔XOR ......... A・2013年 ⇒★A の 中で 一番 古そう★
#    ㋕SHEET ....... A・2013年 ⇒同上
#
#  ★★この 測りで 言える 事／言えない 事★★
#    ★6個とも 見立て通り★ ⇒★「324個を 今すぐ 全部 聞かなくて よい」と ★言える かも★★
#      ＝★但し 「穴が 無い」では ありません★＝★聞いて いない 318個は 聞いて いない★
#    ★1個でも 外れたら★ ⇒★324個 全部 聞きます★
#    ⇒★これは 「測る 順番」の 話です★
#
#  ★★作る 物★★ `%TEMP%\exally-jitsu-excel-nukitori6.xlsx`（★1本の 名★）
#
#  ★★材料★★
#    A1:A6 = 1 / 2 / 2 / 3 / 3 / 4
#    B1 = `=SUM(A1:A3)`（★`FORMULATEXT` に 読ませる 為★）
#    ★式は H 列に 6行 おき★
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③台本の 数 決め打ち（exit 4）／④外へ 出る 6個と ASCII の 外の 字が 0件（exit 5）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$出す先 = Join-Path $env:TEMP 'exally-jitsu-excel-nukitori6.xlsx'
$許す名 = 'exally-jitsu-excel-nukitori6.xlsx'
if ((Split-Path $出す先 -Leaf) -ne $許す名) { Write-Host '★★書いて よい 名は 1本だけです★★'; exit 7 }

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$台本 = @(
  @{ 札 = 'FORMULATEXT'; 式 = '=FORMULATEXT(B1)' },
  @{ 札 = 'SUMPRODUCT';  式 = '=SUMPRODUCT(A1:A3,A1:A3)' },
  @{ 札 = 'CONCATENATE'; 式 = '=CONCATENATE("a","b")' },
  @{ 札 = 'XLOOKUP';     式 = '=XLOOKUP(2,A1:A6,A1:A6)' },
  @{ 札 = 'XOR';         式 = '=XOR(TRUE,FALSE)' },
  @{ 札 = 'SHEET';       式 = '=SHEET()' }
)
$台本の数 = 6
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
  # ★`FORMULATEXT` は ★式を 持つ マス★を 読みます＝先に 1つ 置きます★
  $sh.Range('B1').Formula2 = '=SUM(A1:A3)'
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
    $行数 = $行数 + 6
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
