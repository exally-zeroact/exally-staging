# toru-moto-no-katachi-wo-excel-ni-hirakaseru.ps1
#   -- ★`.xlsb` の 飾りを 実Excel に 聞く★（81）（2026-09-21）
#
#  ★★なぜ★★
#    76 で ★往復した 物★の 図形を 測りました ⇒ ★Left 449.375★
#    ⇒でも 作る 時の 道具は ★Left=320★ で 置いて います
#    ⇒★往復で 動いた のか／最初から 449.375 だった のか★が 分かりません
#    ⇒★これは 見立ての ままでは 出せません★（イ＝実Excel と 同じ 表示に 直に 効く）
#
#  ★★開く 物★★ `%TEMP%\exally-tameshi-xlsb.xlsb`（★1本の 名★）
#    ＝★司さんの 実物と 同じ 形（.xlsb）／但し 中身は 作り物★
#    ＝★実Excel が 作った 元の ファイル★（往復して いません）
#    ・★司さんの 実物の 名は 1文字も 在りません★／★読むだけ★
#
#  ★★見る 物★★
#    図形 ... Left / Top / Width / Height / Rotation / AutoShapeType / Placement
#             塗り / 線の 色 / 線の 太さ ／ 名
#    ★対照★ ... A列の 幅（★30 の はず★）／図形の 数（1個）
#
#  ★★分かれ道★★
#    ・元も ★449.375★ なら ⇒★動いて いない★（作る 時に 既に そこへ 置かれて いた）
#    ・元が ★320★ なら ..... ⇒★往復で 動いた★＝★別の 穴★
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②名が 1文字でも 違えば 走らない（exit 7）
#    ③ファイルが 無ければ 走らない（exit 6）／④走らせる 前の Excel が 0個（exit 3）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = if ($わざと) { Join-Path $ここ 'golden-xlsb-excel-wazato-2026-09-21.tsv' }
      else { Join-Path $ここ 'golden-xlsb-excel-2026-09-21.tsv' }

$許す名 = 'exally-tameshi-xlsb.xlsb'
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

$見るマス = @('A1', 'A2', 'A3', 'B1', 'A5',
              'C1', 'C2', 'C3', 'C4',
              'E1', 'F1', 'G1', 'H1',
              'A10', 'B10', 'C10', 'A11', 'B11', 'C11', 'D10', 'A12',
              'D1', 'D5')

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null; $c = $null; $w = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Open($開く, 0, $true)
  $sh = $bk.Sheets.Item(1)

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★`.xlsb` の 飾りを 実Excel に 聞いた★（81）（2026-09-21）')
  $行.Add('# ★開いた 物★ ... ' + $開く + '（' + (Get-Item $開く).Length + ' バイト）')
  $行.Add('# ★読むだけ★（保存して いません）')
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★どの 貝殻か★ ... PowerShell ' + $版.ToString())
  $行.Add('# ★お客さんが した 事★ ... A1 を 3 ⇒ 42 に した だけ（式は 1つも 触って いない）')
  $行.Add('# ★包みの 中の 字★ ... styles.xml 4402 ⇒ 4402／drawing1.xml 1494 ⇒ 1494（Exally1 の 実測）')

  $欠け = 0
  $表 = New-Object System.Collections.Generic.List[object]
  function 突き合わせ {
    param($名, $待, $出)
    $判 = 'ok'
    if ([string]$出 -ne [string]$待) { $判 = '★違う★' }
    return @($名, [string]$待, [string]$出, $判)
  }
  # ── ★飾り★
  $塗りの待つ = 65535
  if ($わざと) { $塗りの待つ = 1 }
  $表.Add((突き合わせ '罫線(A1 左) LineStyle' 1            ([string]$sh.Range('A1').Borders.Item(7).LineStyle)))
  $表.Add((突き合わせ '罫線(B2 下) Weight'    4            ([string]$sh.Range('B2').Borders.Item(9).Weight)))
  $表.Add((突き合わせ '塗り(B1) Color'        $塗りの待つ  ([string]$sh.Range('B1').Interior.Color)))
  $表.Add((突き合わせ '字(A1) Bold'           'True'       ([string]$sh.Range('A1').Font.Bold)))
  $表.Add((突き合わせ '字(A1) Color'          255          ([string]$sh.Range('A1').Font.Color)))
  $表.Add((突き合わせ '繋げたマス(A5)'         'True'       ([string]$sh.Range('A5').MergeCells)))
  $表.Add((突き合わせ '図形の数'              1            ([string]$sh.Shapes.Count)))
  $表.Add((突き合わせ 'A列の幅'               30           ([string]$sh.Columns.Item(1).ColumnWidth)))
  # ── ★溢れ 3種★
  $縦 = ([string]$sh.Range('C1').Value2) + '/' + ([string]$sh.Range('C2').Value2) + '/' + ([string]$sh.Range('C3').Value2)
  $横 = ([string]$sh.Range('E1').Value2) + '/' + ([string]$sh.Range('F1').Value2) + '/' + ([string]$sh.Range('G1').Value2)
  $二 = ([string]$sh.Range('A10').Value2) + '/' + ([string]$sh.Range('B10').Value2) + '/' + ([string]$sh.Range('C10').Value2) + '/' + ([string]$sh.Range('A11').Value2) + '/' + ([string]$sh.Range('B11').Value2) + '/' + ([string]$sh.Range('C11').Value2)
  $表.Add((突き合わせ '溢れ縦 C1:C3'      '1/2/3'          $縦))
  $表.Add((突き合わせ '溢れ横 E1:G1'      '1/2/3'          $横))
  $表.Add((突き合わせ '溢れ2次元 A10:C11' '1/2/3/4/5/6'    $二))
  $表.Add((突き合わせ 'C1 の式'  '=SEQUENCE(3)'   ([string]$sh.Range('C1').Formula)))
  $表.Add((突き合わせ 'E1 の式'  '=SEQUENCE(1,3)' ([string]$sh.Range('E1').Formula)))
  $表.Add((突き合わせ 'A10 の式' '=SEQUENCE(2,3)' ([string]$sh.Range('A10').Formula)))
  # ── ★対照（溢れない 式）★
  $表.Add((突き合わせ '対照 B1 の式' '=SUM(A1:A3)' ([string]$sh.Range('B1').Formula)))
  $表.Add((突き合わせ '対照 B1 の値' 4.25          ([string]$sh.Range('B1').Value2)))
  $表.Add((突き合わせ '対照 D5 の値' 'abc!'        ([string]$sh.Range('D5').Value2)))
  # ── ★お客さんの 直し★
  $表.Add((突き合わせ 'A1(元の 値)' 3 ([string]$sh.Range('A1').Value2)))
  # ── ★はみ出し★（★溢れが 広がり過ぎて いないか★）
  foreach ($ma in @('C4', 'H1', 'D10', 'A12')) {
    $v = $sh.Range($ma).Value2
    $ある = 'kara'
    if ($null -ne $v) { $ある = [string]$v }
    $表.Add((突き合わせ ('はみ出し ' + $ma) 'kara' $ある))
  }

  $行.Add('#')
  $行.Add('# ★★①飾り ＋ 溢れ ＋ 対照 を 1つずつ 突き合わせ★★')
  $行.Add('# 見た物' + "`t" + '待つ' + "`t" + '出た' + "`t" + '判じ')
  foreach ($t2 in $表) {
    if ($t2[3] -ne 'ok') { $欠け++ }
    $行.Add($t2[0] + "`t" + $t2[1] + "`t" + $t2[2] + "`t" + $t2[3])
  }
  # ★表示の 形は 「含むか」で 見ます★（国で 字が 変わる ので 決め打ちに しません）
  $形1 = [string]$sh.Range('B1').NumberFormatLocal
  $形2 = [string]$sh.Range('A3').NumberFormatLocal
  $判1 = 'ok'; if ($形1 -notmatch '#') { $判1 = '★違う★'; $欠け++ }
  $判2 = 'ok'; if ($形2 -notmatch '%') { $判2 = '★違う★'; $欠け++ }
  $行.Add('表示の形(B1)' + "`t" + '# を 含む' + "`t" + $形1 + "`t" + $判1)
  $行.Add('表示の形(A3)' + "`t" + '% を 含む' + "`t" + $形2 + "`t" + $判2)
  # ★図形の 名★（★数だけでは 「別の 図形」と 見分けが 付きません★）
  $図名 = '(無し)'
  try { if ($sh.Shapes.Count -ge 1) { $図名 = [string]$sh.Shapes.Item(1).Name } } catch { }
  $判3 = 'ok'; if ($図名 -ne 'hanko') { $判3 = '★違う★'; $欠け++ }
  $行.Add('図形の名' + "`t" + 'hanko' + "`t" + $図名 + "`t" + $判3)
  # ══ ★★2026-09-21 足し ── ★図形の 位置と 大きさ★★ ══
  #   ＝Exally1 が 「イ（実Excel と 全く 同じ 表示）」で ★判子を 画面に 出す★ 為に
  #     ★これだけ 先に 要る★と 言って きた 物（Left / Top / Width / Height）
  #   ★単位は ポイント★（1ポイント＝1/72インチ）＝★実Excel が そう 返します★
  #   ★待つ 数は 書きません★＝★作った 時の 数を そのまま 写す 物★だから
  #     （`tsukuru-tameshi-hiraku3-kazari-to-kobore3.ps1` は 320,20,60,60 で 置いて います）
  $行.Add('#')
  $行.Add('# ★★図形の 位置と 大きさ★★（単位＝ポイント）')
  $行.Add('# 見た物' + "`t" + '数')
  try {
    $か = $sh.Shapes.Item(1)
    $行.Add('図形 Left' + "`t" + [string]$か.Left)
    $行.Add('図形 Top' + "`t" + [string]$か.Top)
    $行.Add('図形 Width' + "`t" + [string]$か.Width)
    $行.Add('図形 Height' + "`t" + [string]$か.Height)
    $行.Add('図形 Rotation' + "`t" + [string]$か.Rotation)
    $行.Add('図形 AutoShapeType' + "`t" + [string]$か.AutoShapeType)
    $行.Add('図形 Placement' + "`t" + [string]$か.Placement)
    $行.Add('図形 塗り Color' + "`t" + [string]$か.Fill.ForeColor.RGB)
    $行.Add('図形 線 Color' + "`t" + [string]$か.Line.ForeColor.RGB)
    $行.Add('図形 線 Weight' + "`t" + [string]$か.Line.Weight)
    $か = $null
  } catch { $行.Add('# ★★図形を 読めません★★ ' + $_.Exception.Message) }
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

  # ═══ ★★2つ目の 窓★★（★`=(マス)=0` の 真偽と 型を 一緒に★）═══
  #   ★なぜ★ ... `.Value2` の 「0」は ★本物の 0★ とも ★空★ とも ★誤りの 番号★ とも
  #              区別が 付きません（記憶「意味の 無い 数は 一番 見つけにくい」）
  $行.Add('#')
  $行.Add('# ★★2つ目の 窓★★（★上の 読みの 後に 打って います★）')
  $行.Add('# マス' + "`t" + '値' + "`t" + '型' + "`t" + '=(マス)=0')
  $窓行 = 60
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
  Write-Host ('★★欠け ... ' + $欠け + '個★★')
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
