# tsukuru-jitsu-excel-hanni-18.ps1
#   -- ★範囲を 渡した 時 実Excel は 何を 返すか★（64）（2026-09-20）
#
#  ★★なぜ★★
#    Exally1 が 数えました:
#      ★実Excel が 溢れる 105個の うち ★82個★ で うちの 答えが 足りない★
#      ★その うち 11個は 誤りに ならない＝★お客さんが 気付けません★★
#        UNIQUE ISBLANK ISERR ISERROR ISFORMULA ISLOGICAL ISNA ISNONTEXT
#        ISNUMBER ISTEXT ROW
#    ★但し あちらが 見たのは ★私の 紙の `ref` の 字★だけ★です。
#    ⇒★「実Excel が 本当に 何を 返すか」は ★誰も 見て いません★★
#    ⇒★ここで 見ます★（★中身まで★）
#
#  ★★これを 測らないと 何が 起きるか★★
#    ・`ref="H5:H7"` は ★「3マス 使う」と 書いて ある だけ★
#    ・★中身が FALSE/TRUE/FALSE なのか FALSE/FALSE/FALSE なのかは 別の 話★
#    ・★`=ABS(A1:A3)` は 実Excel でも `#VALUE!` かも しれません★
#      （あちらの ⑥⑵＝★自分の 数を 自分で 疑った★所）
#    ⇒★うちが 「足りない」のか 「実Excel と 同じ」のかが 決まりません★
#
#  ★★盤面（★Exally1 の 盤面に 揃えます★）★★
#    A1:A6 = 1 / 2 / 2 / 3 / 3 / 4      B1:B3 = 10 / 20 / 30
#    D1 = abc                            E1 F1 E2 F2 = 1 / 2 / 3 / 4
#    ★足した 所★ A20 = 5（数）／A21 = `=1+1`（★式★）／A22 = abc（字）
#      ⇒★`ISFORMULA` を 割る 為★＝★A1:A6 は 全部 数なので TRUE が 出ません★
#
#  ★★台本（18本）★★
#    ★㋐気付けない 11個★（Exally1 の ③㋑）
#    ★㋑誤りに なる 側から 抜き取り 5個★ ABS LEN UPPER SQRT DAY
#      ⇒★実Excel でも 誤りなら 「うちが 足りない」では 在りません★
#    ★㋒対照 2個★ SEQUENCE（溢れる）／SUM（溢れない）
#
#  ★★読む 物★★
#    頭の マスから ★3行 × 3列★ を 読みます（★中身まで★）
#    ⇒★何マス 埋まったか★と ★何が 入って いるか★の 両方
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③台本の 数 決め打ち（exit 4）／④外へ 出る 6個と ASCII の 外の 字が 0件（exit 5）
#    ⑤★対照 2個が 待つ 通りで なければ 落とす（exit 6）★＝★盤面が 狂って いない 事★
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-jitsu-excel-hanni-2026-09-20.tsv'

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$台本 = @(
  @{ 札 = 'UNIQUE';     式 = '=UNIQUE(A1:A3)';     組 = 'kizukenai' },
  @{ 札 = 'ISBLANK';    式 = '=ISBLANK(A1:A3)';    組 = 'kizukenai' },
  @{ 札 = 'ISERR';      式 = '=ISERR(A1:A3)';      組 = 'kizukenai' },
  @{ 札 = 'ISERROR';    式 = '=ISERROR(A1:A3)';    組 = 'kizukenai' },
  @{ 札 = 'ISFORMULA';  式 = '=ISFORMULA(A20:A22)'; 組 = 'kizukenai' },
  @{ 札 = 'ISLOGICAL';  式 = '=ISLOGICAL(A1:A3)';  組 = 'kizukenai' },
  @{ 札 = 'ISNA';       式 = '=ISNA(A1:A3)';       組 = 'kizukenai' },
  @{ 札 = 'ISNONTEXT';  式 = '=ISNONTEXT(A1:A3)';  組 = 'kizukenai' },
  @{ 札 = 'ISNUMBER';   式 = '=ISNUMBER(A1:A3)';   組 = 'kizukenai' },
  @{ 札 = 'ISTEXT';     式 = '=ISTEXT(A1:A3)';     組 = 'kizukenai' },
  @{ 札 = 'ROW';        式 = '=ROW(A1:A3)';        組 = 'kizukenai' },
  @{ 札 = 'ABS';        式 = '=ABS(A1:A3)';        組 = 'ayamari-nukitori' },
  @{ 札 = 'LEN';        式 = '=LEN(A1:A3)';        組 = 'ayamari-nukitori' },
  @{ 札 = 'UPPER';      式 = '=UPPER(A1:A3)';      組 = 'ayamari-nukitori' },
  @{ 札 = 'SQRT';       式 = '=SQRT(A1:A3)';       組 = 'ayamari-nukitori' },
  @{ 札 = 'DAY';        式 = '=DAY(A1:A3)';        組 = 'ayamari-nukitori' },
  @{ 札 = 'SEQUENCE';   式 = '=SEQUENCE(3)';       組 = 'taishou' },
  @{ 札 = 'SUM';        式 = '=SUM(A1:A3)';        組 = 'taishou' }
)
$台本の数 = 18
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
$対照ok = 0
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)
  $材 = 1, 2, 2, 3, 3, 4
  for ($r = 1; $r -le 6; $r++) { $sh.Cells.Item($r, 1).Value2 = $材[$r - 1] }
  $sh.Range('B1').Value2 = 10
  $sh.Range('B2').Value2 = 20
  $sh.Range('B3').Value2 = 30
  $sh.Range('D1').Value2 = 'abc'
  $sh.Range('E1').Value2 = 1
  $sh.Range('F1').Value2 = 2
  $sh.Range('E2').Value2 = 3
  $sh.Range('F2').Value2 = 4
  # ★`ISFORMULA` を 割る 為の 3マス★
  $sh.Range('A20').Value2 = 5
  $sh.Range('A21').Formula2 = '=1+1'
  $sh.Range('A22').Value2 = 'abc'

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★範囲を 渡した 時 実Excel は 何を 返すか★（64）（2026-09-20）')
  $行.Add('# ★なぜ★ ... Exally1 が 「105個 中 82個 足りない／11個は 気付けない」と 数えた')
  $行.Add('#           ⇒★但し 見たのは 私の 紙の `ref` の 字だけ＝中身は 未測定★')
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★どの 貝殻か★ ... PowerShell ' + $版.ToString())
  $行.Add('# ★盤面★ A1:A6=1,2,2,3,3,4 ／ B1:B3=10,20,30 ／ D1=abc ／ E1 F1 E2 F2=1,2,3,4')
  $行.Add('#         ★足した★ A20=5（数）／A21==1+1（★式★）／A22=abc（字）')
  $行.Add('# ★読む★ ... 頭の マスから 3行 x 3列（★中身まで★）')
  $行.Add('# 組' + "`t" + '名前' + "`t" + '式' + "`t" + '埋まった数' + "`t" + '中身（|で 区切る／/ で 行を 分ける）')

  $行数 = 1
  foreach ($x in $台本) {
    $ma = 'H' + $行数
    $投げた = ''
    try { $sh.Range($ma).Formula2 = $x.式 } catch { $投げた = '★投げました★' }
    $並 = New-Object System.Collections.Generic.List[string]
    $埋 = 0
    for ($r = 0; $r -lt 3; $r++) {
      $一行 = New-Object System.Collections.Generic.List[string]
      for ($k = 0; $k -lt 3; $k++) {
        $c = $sh.Cells.Item($行数 + $r, 8 + $k)
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
    $行.Add($x.組 + "`t" + $x.札 + "`t" + $x.式 + "`t" + $埋 + "`t" + $出し + $投げた)
    Write-Host ('  ' + $x.組.PadRight(17) + ' ' + $x.札.PadRight(11) + ' 埋まった ' + ([string]$埋).PadRight(3) + ' ' + $出し)
    $行数 = $行数 + 8
  }

  # ══ ★門⑤＝対照が 待つ 通りか★ ══
  $行.Add('#')
  $行.Add('# ★★対照★★（★盤面が 狂って いない 事を 見ます★）')
  # SEQUENCE は H121 から 3マス（17本目）／SUM は H129（18本目）
  $s1 = [string]$sh.Range('H129').Value2   # SEQUENCE の 頭（17本目＝(17-1)*8+1 = 129）
  $s2 = [string]$sh.Range('H137').Value2   # SUM の 頭（18本目）
  $行.Add('# SEQUENCE の 頭 ... ' + $s1 + '（待つ 1）')
  $行.Add('# SUM の 答え ...... ' + $s2 + '（待つ 5）')
  if ($s1 -eq '1') { $対照ok++ }
  if ($s2 -eq '5') { $対照ok++ }
  $行.Add('# ★対照 ok ... ' + $対照ok + ' / 2★')

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★対照 ok ... ' + $対照ok + ' / 2★')
  Write-Host ('★書いた ... ' + $出 + '★')
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
if ($対照ok -ne 2) { Write-Host '★★対照が 合いません＝盤面が 狂って います★★'; exit 6 }
