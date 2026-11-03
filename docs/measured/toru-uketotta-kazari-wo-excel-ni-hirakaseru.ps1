# toru-uketotta-kazari-wo-excel-ni-hirakaseru.ps1
#   -- ★飾り入りの ファイルを 往復させた 物を 実Excel に 開かせる★（㊽）（2026-09-20）
#
#  ★★なぜ 要るか★★
#    ㊼で 測った 材料には ★飾りが 1つも 入って いませんでした★（★私の 落ち度★）
#    ⇒飾り 8種を 入れた 材料を 作り直し（`tsukuru-tameshi-hiraku2-kazari-ari.ps1`）
#    ⇒それを Exally1 が ★お客さんの 道★で 開き ★A1 を 3 ⇒ 42 に した だけ★ で 保存
#    ⇒★包みの 中の 字では `styles.xml`（4,402B）も `drawing1.xml`（1,495B）も
#      ★1バイトも 変わって いません★★（あちらの 実測）
#    ⇒★但し 「実Excel が 飾りとして 見せるか」は ★開くまで 言えません★★
#      ＝記憶「★数字が 全部 緑でも 絵を 開いて 見るまで OKを 出すな★」
#
#  ★★開く 物★★ `%TEMP%\exally-uketotta-kazari.xlsx`（★1本の 名★）
#    元 `exally-tameshi-hiraku2.xlsx` 11,619B
#       sha256 669da7406f2ffbc88bd4a24db643114170f44e4169e21fc02300e04198be2fbf
#    新 9,740B
#       sha256 cdd060636fd2d624a91fd91c0e3a64fbe5a2266f16bcc8d81c8e5fe73e5403e1
#    ・★司さんの 実物の 名は 1文字も 在りません★
#    ・★読むだけ★（`Close($false)`／`SaveAs` は 1文字も 在りません）
#
#  ★★8種を 1つずつ 数えます（★待つ 数を 字で 書いて おきます★）★★
#    ㋐罫線（A1 左） LineStyle ... 待つ 1
#    ㋐罫線（B2 下） Weight .... 待つ 4
#    ㋑塗り（B1） Color ........ 待つ 65535
#    ㋒字（A1） Bold / Color ... 待つ True / 255
#    ㋓表示の 形（B1 / A3） .... 待つ `#` を 含む / `%` を 含む
#    ㋔繋げた マス（A5） ....... 待つ True
#    ㋕図形の 数 ............... 待つ 1（名 `hanko`）
#    ㋖A列の 幅 ................ 待つ 30
#    ㋗溢れ（C1:C3） ........... 待つ 1/2/3 ／ C1 の 式 `=SEQUENCE(3)`
#    ㋘式（B1） ................ 待つ `=SUM(A1:A3)` ／ 値 43.25（42 + 1 + 0.25）
#    ★A1★ .................... 待つ 42（★お客さんが 直した 所★）
#
#  ★★待つ 数を 先に 字で 書く 訳★★
#    ★出た 数を 見て から 「そんな ものか」と 思うのを 防ぐ★
#    ＝記憶「★探す 字を 決めた 時点で 答えが 決まる★」の 裏＝★先に 決めて おく★
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②名が 1文字でも 違えば 走らない（exit 7）
#    ③ファイルが 無ければ 走らない（exit 6）／④走らせる 前の Excel が 0個（exit 3）
#    ⑤★見る 数が 決め打ちと 同じ★（exit 4）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

# ★★`-わざと` ＝ ★待つ 数を 1つ 外して 赤に なるかを 見る★★
#   ★なぜ★ ... 「欠け 0個」は ★赤に なれる 道具の 0★ でないと 値打ちが 在りません
#              （記憶「★わざと 壊したら 赤に なるか★」「★壊したのに 赤に ならない★」）
#   ⇒塗り(B1) の 待つ 数を 65535 ⇒ 1 に 変えます（★出る 数は 触りません★）
#   ⇒★紙は 別の 名に 書きます★（本物の 紙を 汚さない）
param([switch]$わざと)

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = if ($わざと) { Join-Path $ここ 'golden-uketotta-kazari-excel-wazato-2026-09-20.tsv' }
      else { Join-Path $ここ 'golden-uketotta-kazari-excel-2026-09-20.tsv' }

$許す名 = 'exally-uketotta-kazari.xlsx'
$開く = Join-Path $env:TEMP $許す名

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }
if ((Split-Path $開く -Leaf) -ne $許す名) { Write-Host '★★開いて よい ファイルは 1本だけです★★'; exit 7 }
if (-not (Test-Path $開く)) { Write-Host ('★★在りません ... ' + $開く + '★★'); exit 6 }
Write-Host ('★開く 物 ... ' + $開く + '（' + (Get-Item $開く).Length + ' バイト）★')

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$見るマス = @('A1', 'A2', 'A3', 'B1', 'C1', 'C2', 'C3', 'C4', 'A5', 'D1', 'E1')
$見る数 = 11
if ($見るマス.Count -ne $見る数) { exit 4 }

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null; $c = $null; $w = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Open($開く, 0, $true)
  $sh = $bk.Sheets.Item(1)

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★飾り入りの ファイルを 往復させた 物を 実Excel に 開かせた★（㊽）（2026-09-20）')
  $行.Add('# ★開いた 物★ ... ' + $開く + '（' + (Get-Item $開く).Length + ' バイト）')
  $行.Add('# ★読むだけ★（保存して いません）')
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★どの 貝殻か★ ... PowerShell ' + $版.ToString())
  $行.Add('# ★お客さんが した 事★ ... A1 を 3 ⇒ 42 に した ★だけ★（式は 1つも 触って いない）')
  $行.Add('# ★包みの 中の 字★ ... styles.xml 4402B ／ drawing1.xml 1495B ★とも 元と 同じ★（Exally1 の 実測）')

  # ═══ ★★①飾り 8種（★待つ 数を 先に 書いて 突き合わせます★）★★ ═══
  $行.Add('#')
  $行.Add('# ★★①飾り★★（★待つ 数は 材料を 作った 時に 読み返した 数★）')
  $行.Add('# 見た物' + "`t" + '待つ' + "`t" + '出た' + "`t" + '判じ')
  $欠け = 0
  function 突き合わせ {
    param($名, $待, $出)
    $判 = 'ok'
    if ([string]$出 -ne [string]$待) { $判 = '★違う★' }
    return @($名, [string]$待, [string]$出, $判)
  }
  $表 = New-Object System.Collections.Generic.List[object]
  $表.Add((突き合わせ '罫線(A1 左) LineStyle' 1 ([string]$sh.Range('A1').Borders.Item(7).LineStyle)))
  $表.Add((突き合わせ '罫線(B2 下) Weight' 4 ([string]$sh.Range('B2').Borders.Item(9).Weight)))
  # ★★`-わざと` の 時だけ 待つ 数を 外します★★（★出る 数は 触りません★）
  $塗りの待つ = 65535
  if ($わざと) { $塗りの待つ = 1 }
  $表.Add((突き合わせ '塗り(B1) Color' $塗りの待つ ([string]$sh.Range('B1').Interior.Color)))
  $表.Add((突き合わせ '字(A1) Bold' 'True' ([string]$sh.Range('A1').Font.Bold)))
  $表.Add((突き合わせ '字(A1) Color' 255 ([string]$sh.Range('A1').Font.Color)))
  $表.Add((突き合わせ '繋げたマス(A5)' 'True' ([string]$sh.Range('A5').MergeCells)))
  $表.Add((突き合わせ '図形の数' 1 ([string]$sh.Shapes.Count)))
  $表.Add((突き合わせ 'A列の幅' 30 ([string]$sh.Columns.Item(1).ColumnWidth)))
  $表.Add((突き合わせ 'C1 の式' '=SEQUENCE(3)' ([string]$sh.Range('C1').Formula)))
  $表.Add((突き合わせ 'B1 の式' '=SUM(A1:A3)' ([string]$sh.Range('B1').Formula)))
  $表.Add((突き合わせ 'A1 の値' 43.25 ([string]$sh.Range('B1').Value2)))
  $表.Add((突き合わせ 'A1(客が直した)' 42 ([string]$sh.Range('A1').Value2)))
  foreach ($t2 in $表) {
    if ($t2[3] -ne 'ok') { $欠け++ }
    $行.Add($t2[0] + "`t" + $t2[1] + "`t" + $t2[2] + "`t" + $t2[3])
  }
  # ★表示の 形は 「含むか」で 見ます★（国によって 字が 変わる ので 決め打ちに しません）
  $形1 = [string]$sh.Range('B1').NumberFormatLocal
  $形2 = [string]$sh.Range('A3').NumberFormatLocal
  $判1 = 'ok'; if ($形1 -notmatch '#') { $判1 = '★違う★'; $欠け++ }
  $判2 = 'ok'; if ($形2 -notmatch '%') { $判2 = '★違う★'; $欠け++ }
  $行.Add('表示の形(B1)' + "`t" + '# を 含む' + "`t" + $形1 + "`t" + $判1)
  $行.Add('表示の形(A3)' + "`t" + '% を 含む' + "`t" + $形2 + "`t" + $判2)
  # ★図形の 名も 見ます★（★数だけでは 「別の 図形」と 見分けが 付きません★）
  $図名 = '(無し)'
  try { if ($sh.Shapes.Count -ge 1) { $図名 = [string]$sh.Shapes.Item(1).Name } } catch { }
  $判3 = 'ok'; if ($図名 -ne 'hanko') { $判3 = '★違う★'; $欠け++ }
  $行.Add('図形の名' + "`t" + 'hanko' + "`t" + $図名 + "`t" + $判3)
  $行.Add('# ★★欠け ' + $欠け + '個★★')

  # ═══ ★★②マスの 中身★★ ═══
  $行.Add('#')
  $行.Add('# ★★②マスの 中身★★')
  $行.Add('# マス' + "`t" + '値' + "`t" + '出る字' + "`t" + '式' + "`t" + '型' + "`t" + '溢れの一部か')
  foreach ($ma in $見るマス) {
    $c = $sh.Range($ma)
    $v = $c.Value2
    $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    $型 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
    $一部か = '(?)'
    try { $一部か = [string]$c.HasArray } catch { }
    $行.Add($ma + "`t" + $値 + "`t" + [string]$c.Text + "`t" + [string]$c.Formula + "`t" + $型 + "`t" + $一部か)
  }

  # ═══ ★★2つ目の 窓★★（★`=(マス)=0` の 真偽と 型を 一緒に 取る★）═══
  #   ★なぜ★ ... `.Value2` の 「0」は ★本物の 0★ とも ★空★ とも ★誤りの 番号★ とも
  #              区別が 付きません（記憶「意味の 無い 数は 一番 見つけにくい」）
  #   ★上の 読みが 済んだ 後に 打ちます★
  $行.Add('#')
  $行.Add('# ★★2つ目の 窓★★（★上の 読みの 後に 打って います★）')
  $行.Add('# マス' + "`t" + '値' + "`t" + '型' + "`t" + '=(マス)=0')
  $窓行 = 40
  foreach ($ma2 in $見るマス) {
    $c = $sh.Range($ma2)
    $v = $c.Value2
    $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    $型 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
    $ゼロか = '(★窓2が 打てません★)'
    try { $w = $sh.Range('P' + $窓行); $w.Formula2 = '=(' + $ma2 + ')=0'; $ゼロか = [string]$w.Value2 } catch { }
    $窓行 = $窓行 + 1
    $行.Add($ma2 + "`t" + $値 + "`t" + $型 + "`t" + $ゼロか)
  }

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ('★書いた ... ' + $出 + '★')
  Write-Host ('★★飾りの 欠け ... ' + $欠け + '個★★')
  if ($わざと) {
    if ($欠け -eq 0) { Write-Host '★★わざと 外したのに 欠け 0＝門が 効いて いません★★' }
    else { Write-Host ('★わざと 外したら 欠け ' + $欠け + '個＝門は 効いて います★') }
  }
  # ★★保存しません★★
  $bk.Close($false)
  $bk = $null
} finally {
  $c = $null; $w = $null; $sh = $null
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
