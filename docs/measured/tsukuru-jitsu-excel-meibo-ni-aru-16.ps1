# tsukuru-jitsu-excel-meibo-ni-aru-16.ps1
#   -- ★名簿に 「在る」16個を 実Excel に 書かせる（★付け過ぎ★を 探す）★（58）（2026-09-20）
#
#  ★★なぜ★★
#    今まで 探して いたのは ★付け忘れ★でした（WRAPROWS / WRAPCOLS / MODE.MULT）。
#    ★逆も 壊れます★＝★`_xlfn.` を 付け過ぎると 実Excel が 読めません★
#    ⇒`TRANSPOSE` で 実測済み（★実Excel は 裸で 書く★）
#    ⇒★うちの 名簿（133本）に 在る 物が 本当に 要るかは ★聞いた 事が 在りません★★
#    ⇒Exally1 の 数え ... 09-16 の 紙は ★分母が 45個 漏れて いた★
#      ⇒★中身も 漏れて いる かも しれません★
#
#  ★★この 16個を 選んだ 訳★★
#    Exally1 が 「分母から 漏れて いた 45個の うち ★名簿には 在った 18個★」を 挙げました。
#    その うち `FORECAST` は 55 で ★裸★ と 出ており ★名簿にも 在りません★（確かめました）。
#    残り 16個 ＋ `MODE.SNGL` は 既に 通って いる ので 除く ⇒★16個★
#
#  ★★割りたい 見立て★★
#    「★2007年 までの 物は 裸／2010年 以降は `_xlfn.`★」
#    ⇒★両端で 割ります★
#        `AGGREGATE`（2010年）... ★裸なら 見立てが 崩れる★
#        `XMATCH`（2019年）..... ★裸なら 見立てが 崩れる★
#    ⇒★どちらかが 外れたら 見立てを 捨てます★（記憶「★外れた 見込みは 捨てない★」＝紙に 残す）
#
#  ★★作る 物★★ `%TEMP%\exally-jitsu-excel-meibo16.xlsx`（★1本の 名★）
#    ★実Excel に 打たせて 保存させます★（★うちの 書き出しは 1回も 通しません★）
#
#  ★★材料と 置き場（★押す 前に 隙間を 計算しました★）★★
#    A1:A6 = 1 / 2 / 2 / 3 / 3 / 4
#    E1:F2 = 1 / 2 / 3 / 4
#    ★式は H 列に 6行 おき★（H1 H7 H13 ...）＝★縦に 5マス 溢れても ぶつかりません★
#    ★横に 溢れる 物★（`TOROW`）は I J K へ 伸びます＝★H 列だけ 使う ので 空いて います★
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③台本の 数 決め打ち（exit 4）／④外へ 出る 6個と ASCII の 外の 字が 0件（exit 5）
#
#  ★★この 道具は 「作って 打った 結果を 見せる」だけ★★
#    ＝★印は `yomu-jitsu-excel-no-shirushi.mjs` で 生の 字から 読みます★
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$出す先 = Join-Path $env:TEMP 'exally-jitsu-excel-meibo16.xlsx'
$許す名 = 'exally-jitsu-excel-meibo16.xlsx'
if ((Split-Path $出す先 -Leaf) -ne $許す名) { Write-Host '★★書いて よい 名は 1本だけです★★'; exit 7 }

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$台本 = @(
  @{ 札 = 'ARRAYTOTEXT';  式 = '=ARRAYTOTEXT(A1:A3)' },
  @{ 札 = 'CHOOSECOLS';   式 = '=CHOOSECOLS(E1:F2,1)' },
  @{ 札 = 'CHOOSEROWS';   式 = '=CHOOSEROWS(E1:F2,1)' },
  @{ 札 = 'DROP';         式 = '=DROP(A1:A6,1)' },
  @{ 札 = 'EXPAND';       式 = '=EXPAND(A1:A2,3,1,0)' },
  @{ 札 = 'TAKE';         式 = '=TAKE(A1:A6,2)' },
  @{ 札 = 'TOCOL';        式 = '=TOCOL(E1:F2)' },
  @{ 札 = 'TOROW';        式 = '=TOROW(E1:F2)' },
  @{ 札 = 'AGGREGATE';    式 = '=AGGREGATE(1,0,A1:A6)' },
  @{ 札 = 'ENCODEURL';    式 = '=ENCODEURL("a b")' },
  @{ 札 = 'NUMBERVALUE';  式 = '=NUMBERVALUE("1.5")' },
  @{ 札 = 'PERMUTATIONA'; 式 = '=PERMUTATIONA(5,2)' },
  @{ 札 = 'TEXTAFTER';    式 = '=TEXTAFTER("a-b","-")' },
  @{ 札 = 'TEXTBEFORE';   式 = '=TEXTBEFORE("a-b","-")' },
  @{ 札 = 'VALUETOTEXT';  式 = '=VALUETOTEXT(A1)' },
  @{ 札 = 'XMATCH';       式 = '=XMATCH(2,A1:A6)' }
)
$台本の数 = 16
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
