# tsukuru-jitsu-excel-kobore-6shurui.ps1
#   -- ★実Excel 自身が 溢れる 式を どう 書くか★（51）（2026-09-20）
#
#  ★★なぜ 先に これを 取るか★★
#    Exally1 の 問い ... 「うちが `cm="1"` を 付けるのは 正しいか」
#    ⇒★物差しは 実Excel★（記憶）＝★実Excel 自身が 付ける 式／付けない 式★を 先に 数える
#    ⇒★もし 実Excel が 付けない 式が 在れば、うちが 付けるのは 間違い★
#
#  ★★もう 1つ の 訳★★
#    今まで 測った 溢れは ★`SEQUENCE` だけ★でした。
#    ★`SEQUENCE` は 台が 得意な 形★＝★通ったのは その せい かも しれません★
#    ⇒`SORT` `UNIQUE` `FILTER` は ★借り物（HyperFormula）も 答えられる 溢れ★
#    ⇒★特に★
#        `FILTER`    ... ★縮む★（09-19 に 34本中 1本 違った 所）
#        `TRANSPOSE` ... ★縦が 横に なる★（範囲の 組み方が 別）
#        `MAKEARRAY` ... ★LAMBDA を 使う★
#
#  ★★作る 物★★ `%TEMP%\exally-jitsu-excel-kobore.xlsx`（★1本の 名★）
#    ★実Excel に 打たせて 保存させます★（★うちの 書き出しは 1回も 通しません★）
#    ・★司さんの 実物の 名は 1文字も 在りません★
#    ・★既に 在れば 消してから 作ります★（★消すのは この 1本だけ★）
#
#  ★★置く 所（★押す 前に 隙間を 計算しました★）★★
#      A1:A3 = 3 / 1 / 2          ★材料★
#      C1  `=SEQUENCE(3)`                      ⇒ C1:C3   ★対照＝`cm` が 付くと 分かって いる 形★
#      E1  `=SORT(A1:A3)`                      ⇒ E1:E3
#      G1  `=UNIQUE(A1:A3)`                    ⇒ G1:G3
#      I1  `=FILTER(A1:A3,A1:A3>1)`            ⇒ I1:I2   ★縮む★（3 と 2 の 2つ）
#      K1  `=TRANSPOSE(A1:A3)`                 ⇒ K1:M1   ★横に なる★
#      A10 `=MAKEARRAY(2,3,LAMBDA(r,c,r*c))`   ⇒ A10:C11 ★2次元★
#      A15 `=SUM(A1:A3)`                       ⇒ 6       ★対照＝溢れない 式★
#    ★どれも 重なりません★（列 3/5/7/9/11-13 ／ 行 10-11 ／ 行 15）
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③台本の 数 決め打ち（exit 4）
#    ④★外へ 出る 6個（WEBSERVICE 等）と ASCII の 外の 字が 式に 0件★（exit 5）
#    ⑤★書いた 後に 読み返して ★溢れた 形★が 待つ 通りか 数える★（1つでも 違えば exit 6）
#      ＝★「打てた」は 「溢れた」では ありません★
#
#  ★★この 道具は 「作って 読み返す」だけ★★
#    ＝★包みの 中の 字（`cm=` `ref=`）は 別の 道具で 読みます★（1つの 道具に 2つの 仕事を させない）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$出す先 = Join-Path $env:TEMP 'exally-jitsu-excel-kobore.xlsx'
$許す名 = 'exally-jitsu-excel-kobore.xlsx'
if ((Split-Path $出す先 -Leaf) -ne $許す名) { Write-Host '★★書いて よい 名は 1本だけです★★'; exit 7 }

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

# ══ ★台本★（★待つ 形も 字で 書いて おきます★）══
$台本 = @(
  @{ 頭 = 'C1';  式 = '=SEQUENCE(3)';                    待つ = '1/2/3';       広さ = 'C1:C3';   札 = 'SEQUENCE(taishou)' },
  @{ 頭 = 'E1';  式 = '=SORT(A1:A3)';                    待つ = '1/2/3';       広さ = 'E1:E3';   札 = 'SORT' },
  @{ 頭 = 'G1';  式 = '=UNIQUE(A1:A3)';                  待つ = '3/1/2';       広さ = 'G1:G3';   札 = 'UNIQUE' },
  @{ 頭 = 'I1';  式 = '=FILTER(A1:A3,A1:A3>1)';          待つ = '3/2';         広さ = 'I1:I2';   札 = 'FILTER(chijimu)' },
  @{ 頭 = 'K1';  式 = '=TRANSPOSE(A1:A3)';               待つ = '3/1/2';       広さ = 'K1:M1';   札 = 'TRANSPOSE(yoko)' },
  @{ 頭 = 'A10'; 式 = '=MAKEARRAY(2,3,LAMBDA(r,c,r*c))'; 待つ = '1/2/3/2/4/6'; 広さ = 'A10:C11'; 札 = 'MAKEARRAY(2jigen)' },
  @{ 頭 = 'A15'; 式 = '=SUM(A1:A3)';                     待つ = '6';           広さ = 'A15';     札 = 'SUM(koborenai)' }
)
$台本の数 = 7
Write-Host ('★台本 ... ' + $台本.Count + '本★（決め打ち ' + $台本の数 + '本）')
if ($台本.Count -ne $台本の数) { exit 4 }

# ══ ★門④＝外へ 出る 式と ASCII の 外の 字を 数える★ ══
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
$bk = $null; $sh = $null
$欠け = 0
$出来た = $false
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)
  $sh.Range('A1').Value2 = 3
  $sh.Range('A2').Value2 = 1
  $sh.Range('A3').Value2 = 2
  # ★`Formula` では 溢れません＝`Formula2` です★（2026-09-19 実測）
  foreach ($x in $台本) {
    try { $sh.Range($x.頭).Formula2 = $x.式 }
    catch { Write-Host ('★★打てません ... ' + $x.頭 + ' ' + $x.式 + ' ⇒ ' + $_.Exception.Message + '★★'); $欠け++ }
  }

  if (Test-Path $出す先) { Remove-Item $出す先 -Force }
  $bk.SaveAs($出す先, 51)
  $bk.Close($false)
  $bk = $null

  # ══ ★★読み返して 「溢れた 形」を 数える★★（★「打てた」は 「溢れた」では ない★）══
  $bk = $xl.Workbooks.Open($出す先, 0, $true)
  $sh = $bk.Sheets.Item(1)
  Write-Host ''
  Write-Host '★★読み返し★★'
  foreach ($x in $台本) {
    $r = $sh.Range($x.広さ)
    $並 = New-Object System.Collections.Generic.List[string]
    foreach ($セル in $r) {
      $v = $セル.Value2
      if ($null -eq $v) { $並.Add('(kara)') }
      elseif ($v -is [double]) { $並.Add($v.ToString('R', [Globalization.CultureInfo]::InvariantCulture)) }
      else { $並.Add([string]$v) }
    }
    $出た = ($並 -join '/')
    $判 = 'ok'
    if ($出た -ne $x.待つ) { $判 = '★違う★'; $欠け++ }
    $一部か = '(?)'
    try { $一部か = [string]$sh.Range($x.頭).HasArray } catch { }
    $出し = '  ' + $x.札.PadRight(22) + ' ' + $x.広さ.PadRight(9) + ' 待つ ' + $x.待つ.PadRight(13) + ' 出た ' + $出た.PadRight(13) + ' HasArray=' + $一部か + ' ' + $判
    Write-Host $出し
  }
  $bk.Close($false)
  $bk = $null
  Write-Host ('★★欠け ... ' + $欠け + '個★★')
  if ($欠け -ne 0) { exit 6 }
  $出来た = $true
} finally {
  $sh = $null
  if ($null -ne $bk) { $bk.Close($false) }
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
  if ($残り -eq 0) { Write-Host ('★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒で ★消えました★★') }
  else { Write-Host ('★★Excel は 消えませんでした★★ ／ 残り ' + $残り + '個') }
}
if ($出来た) {
  $x2 = Get-Item $出す先
  $h = (Get-FileHash $出す先 -Algorithm SHA256).Hash.ToLower()
  Write-Host ''
  Write-Host ('★★作りました★★ ... ' + $出す先)
  Write-Host ('  ★大きさ★ ' + $x2.Length + ' バイト')
  Write-Host ('  ★sha256★ ' + $h)
  Write-Host '  ⇒★包みの 中の 字（cm= / ref=）は 別の 道具で 読みます★'
}
