# toru-nijigen-kakidashi-wo-excel-ni-hirakaseru.ps1
#   -- ★2次元で 溢れた 物を 書き出して 実Excel に 開かせる★（70）（2026-09-21）
#
#  ★★なぜ★★
#    Exally1 が `見分け` を ★2次元も 溢れる★形に 直しました（1行）。
#    ⇒★画面では 10マスに 溢れました★（あちらの 実測）
#    ⇒★書き出した 生の 字も `ref="D1:E5"`（5行 x 2列）★（あちらの 実測）
#    ⇒★でも 「実Excel が どう 見せるか」は 開くまで 言えません★
#      ＝記憶「★数字が 全部 緑でも 絵を 開いて 見るまで OKを 出すな★」
#
#  ★★開く 物★★ `%TEMP%\exally-kakidashi-nijigen.xlsx`（★1本の 名★）
#    17,312B ／ sha256 072e86783546a697b71adc30470fa0989cfce59b8ebbf5526824d6320e2a6cae
#    ★お客さんの 道★で 作った 物（画面で 打つ ⇒ 本番の 書き出しの 2行）
#    ★盤面★ A1:A5 = 1..5 ／ B1:B5 = 2,4,6,8,10
#    ★生の 字★
#      D1 `<f t="array" ref="D1:E5">ISNUMBER(A1:B5)</f>` cm 在り
#      G1 `<f t="array" ref="G1:H5">ISERR(A1:B5)</f>`    cm 在り
#      J1 `<f t="array" ref="J1:K5">ISERROR(A1:B5)</f>`  cm 在り
#      M1 `<f t="array" ref="M1:M5">ISNUMBER(A1:A5)</f>` cm 在り ★対照（1次元）★
#      P1 `<f>SUM(A1:A5)</f>`                            cm 無し ★対照（溢れない）★
#    ・★司さんの 実物の 名は 1文字も 在りません★／★読むだけ★（保存しません）
#
#  ★★5問★★
#    ⑴D1:E5 が ★10マス 全部 TRUE★ か
#    ⑵G1:H5 と J1:K5 が ★10マス 全部 FALSE★ か
#    ⑶対照 M1:M5 が ★5マス TRUE★ か（★1次元が 下がって いないか★）
#    ⑷対照 P1 が ★15★ か（★溢れない 式★）
#    ⑸★はみ出し★ ... F1 ／ I1 ／ L1 ／ N1 が 空か
#
#  ★★置き場★★ ... ★この 紙は 「開いて 読む」だけ＝式は 打ちません★
#  ★★読む 窓★★ ... 頭の マスから ★6行 x 3列★
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②名が 1文字でも 違えば 走らない（exit 7）
#    ③ファイルが 無ければ 走らない（exit 6）／④走らせる 前の Excel が 0個（exit 3）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-nijigen-kakidashi-excel-2026-09-21.tsv'

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$台本 = @(
  @{ 札 = 'ISNUMBER-2jigen'; 頭 = 'D1'; 式 = '=ISNUMBER(A1:B5)'; 組 = '2jigen' },
  @{ 札 = 'ISERR-2jigen';    頭 = 'G1'; 式 = '=ISERR(A1:B5)';    組 = '2jigen' },
  @{ 札 = 'ISERROR-2jigen';  頭 = 'J1'; 式 = '=ISERROR(A1:B5)';  組 = '2jigen' },
  @{ 札 = 'ISNUMBER-1jigen'; 頭 = 'M1'; 式 = '=ISNUMBER(A1:A5)'; 組 = 'taishou-1jigen' },
  @{ 札 = 'SEQUENCE';        頭 = 'P1'; 式 = '=SEQUENCE(2)';     組 = 'taishou' }
)
$台本の数 = 5
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

$許す名 = 'exally-kakidashi-nijigen.xlsx'
$開く = Join-Path $env:TEMP $許す名
if ((Split-Path $開く -Leaf) -ne $許す名) { Write-Host '★★開いて よい ファイルは 1本だけです★★'; exit 7 }
if (-not (Test-Path $開く)) { Write-Host ('★★在りません ... ' + $開く + '★★'); exit 6 }
Write-Host ('★開く 物 ... ' + $開く + '（' + (Get-Item $開く).Length + ' バイト）★')

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null; $c = $null
$対照ok = 0
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Open($開く, 0, $true)
  $sh = $bk.Sheets.Item(1)

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★範囲を 渡した 時 実Excel は 何を 返すか★（64）（2026-09-20）')
  $行.Add('# ★なぜ★ ... 2次元で 溢れた 物を 書き出して 実Excel が どう 見せるか')
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★どの 貝殻か★ ... PowerShell ' + $版.ToString())
  $行.Add('# ★開いた 物★ ... ' + $開く + '（' + (Get-Item $開く).Length + ' バイト）')
  $行.Add('# ★読むだけ★（保存して いません／式は 1本も 打って いません）')
  $行.Add('# ★盤面★ A1:A5=1,2,3,4,5 ／ B1:B5=2,4,6,8,10')
  $行.Add('# ★置き場★ ... D1 G1 J1 M1 P1（★書き出した 物の まま★）')
  $行.Add('# ★読む 窓★ ... 頭の マスから 6行 x 3列')
  $行.Add('# 組' + "`t" + '名前' + "`t" + 'マス' + "`t" + '式' + "`t" + '型' + "`t" + '=(マス)=0' + "`t" + '埋まった数' + "`t" + '中身（|で マス／ / で 行）')

  $窓行 = 40
  $w2 = $null
  foreach ($x in $台本) {
    $ma = $x.頭
    # ★★式は 打ちません＝開いた 物を 読むだけ★★
    $投げた = ''
    $頭列 = $sh.Range($x.頭).Column
    $並 = New-Object System.Collections.Generic.List[string]
    $埋 = 0
    for ($r = 0; $r -lt 6; $r++) {
      $一行 = New-Object System.Collections.Generic.List[string]
      for ($k = 0; $k -lt 3; $k++) {
        $c = $sh.Cells.Item(1 + $r, $頭列 + $k)
        $v = $c.Value2
        if ($null -eq $v) { $一行.Add('') }
        else {
          $埋++
          $字 = [string]$c.Text
          $一行.Add($字)
        }
      }
      $並.Add($一行 -join '|')
    }
    $出し = ($並 -join ' / ')
    # ══ ★★2つ目の 窓★★（★`=(マス)=0` の 真偽と ★型★ を 一緒に 取る★）══
    #   ★なぜ★ ... `.Value2` の 「0」は ★本物の 0★ とも ★空★ とも ★誤りの 番号★ とも
    #              区別が 付きません（記憶「意味の 無い 数は 一番 見つけにくい」）
    #   ★置き場★ ... T列（★溢れは P まで＝ぶつかりません★）
    $v0 = $sh.Range($x.頭).Value2
    $型 = if ($null -eq $v0) { '(kara)' } elseif ($v0 -is [double]) { 'Double' } elseif ($v0 -is [string]) { 'String' } elseif ($v0 -is [bool]) { 'Boolean' } else { 'Other' }
    $ゼロか = '(★窓2が 打てません★)'
    try { $w2 = $sh.Range('T' + $窓行); $w2.Formula2 = '=(' + $x.頭 + ')=0'; $ゼロか = [string]$w2.Value2 } catch { }
    $窓行 = $窓行 + 1
    $行.Add($x.組 + "`t" + $x.札 + "`t" + $x.頭 + "`t" + $x.式 + "`t" + $型 + "`t" + $ゼロか + "`t" + $埋 + "`t" + $出し + $投げた)
    Write-Host ('  ' + $x.組.PadRight(17) + ' ' + $x.札.PadRight(11) + ' 埋まった ' + ([string]$埋).PadRight(3) + ' ' + $出し)
  }

  # ══ ★門⑤＝対照が 待つ 通りか★ ══
  $行.Add('#')
  $行.Add('# ★★対照★★（★盤面が 狂って いない 事を 見ます★）')
  # SEQUENCE は H121 から 3マス（17本目）／SUM は H129（18本目）
  # ★★対照の 読みを 忘れて いました（★1回目は 0/2 に なりました★）★★
  #   ＝★$t1 $t2 に 何も 入れずに 比べて いました★
  #   ⇒★「対照が 合わない」は 出た＝門は 効いて いました★（★黙って 緑に ならなかった★）
  $t1 = [string]$sh.Range('M5').Value2
  $t2 = [string]$sh.Range('P1').Value2
  if ($t1 -eq 'True') { $対照ok++ }
  if ($t2 -eq '15') { $対照ok++ }
  $行.Add('# 対照 M5（ISNUMBER(A1:A5) の 5マス目） ... ' + $t1 + '（待つ True）')
  $行.Add('# 対照 P1（SUM(A1:A5)） ............... ' + $t2 + '（待つ 15）')
  $行.Add('# ★対照 ok ... ' + $対照ok + ' / 2★')

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★対照 ok ... ' + $対照ok + ' / 2★')
  Write-Host ('★書いた ... ' + $出 + '★')
  $bk.Close($false)
  $bk = $null
} finally {
  # ★★掴んだ物 全部 $null★★
  $w2 = $null
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
if ($対照ok -ne 2) { Write-Host '★★対照が 合いません＝盤面が 狂って います★★'; exit 6 }
