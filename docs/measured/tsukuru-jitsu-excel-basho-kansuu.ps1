# tsukuru-jitsu-excel-basho-kansuu.ps1
#   -- ★場所を 見る 関数と COLUMN の 向きを 実Excel に 聞く★（75）（2026-09-21）
#
#  ★★なぜ★★
#    Exally1 が 「範囲の 溢れ」を 直し 29本 ⇒ 2本 に しました。
#    ⇒★その 中で 2つ 「私では 出せない」と 言って きました★
#      ⑴`=COLUMN(E1:F2)`（★横に 2列 以上★を 渡した 時）
#         ＝今の 紙は ★1列の 範囲しか 無い★＝★片方の 向きしか 押せて いない★
#      ⑵`=FORMULATEXT(A1:A3)` と `=CELL("row",A1:A3)`
#         ＝★場所を 見る 関数★＝`ISFORMULA` と 同じ 穴が 在るかも しれない
#
#  ★★盤面★★
#    A1:A3 = 1 / 2 / 2      （★A2 は 式（=1+1）に します★＝FORMULATEXT を 割る 為）
#    E1 F1 = 1 / 2  ／ E2 F2 = 3 / 4
#
#  ★★置き場★★ ... ★A1:A3 と 交わる 行（行1）★／列は 4つ おき（H L P T）
#  ★★読む 窓★★ ... 頭の マスから 6行 x 4列
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③台本の 数 決め打ち（exit 4）／④外へ 出る 6個と ASCII の 外の 字が 0件（exit 5）
#    ⑤★対照 2個が 待つ 通りで なければ 落とす（exit 6）★
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-jitsu-excel-basho-kansuu-2026-09-21.tsv'

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$台本 = @(
  @{ 札 = 'COLUMN-yoko2';   頭 = 'H1'; 式 = '=COLUMN(E1:F2)';        組 = 'shirabe' },
  @{ 札 = 'COLUMN-tate';    頭 = 'L1'; 式 = '=COLUMN(A1:A3)';        組 = 'shirabe' },
  @{ 札 = 'ROW-yoko2';      頭 = 'P1'; 式 = '=ROW(E1:F2)';           組 = 'shirabe' },
  @{ 札 = 'FORMULATEXT-h';  頭 = 'T1'; 式 = '=FORMULATEXT(A1:A3)';   組 = 'basho' },
  @{ 札 = 'CELL-row-h';     頭 = 'X1'; 式 = '=CELL("row",A1:A3)';    組 = 'basho' },
  @{ 札 = 'ISFORMULA-h';    頭 = 'AB1'; 式 = '=ISFORMULA(A1:A3)';    組 = 'basho' },
  @{ 札 = 'SEQUENCE';       頭 = 'AF1'; 式 = '=SEQUENCE(2)';         組 = 'taishou' },
  @{ 札 = 'SUM';            頭 = 'AJ1'; 式 = '=SUM(A1:A3)';          組 = 'taishou' }
)
$台本の数 = 8
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
  # ★盤面★（★A2 は 式★＝FORMULATEXT / ISFORMULA を 割る 為）
  $sh.Range('A1').Value2 = 1
  $sh.Range('A2').Formula2 = '=1+1'
  $sh.Range('A3').Value2 = 2
  $sh.Range('E1').Value2 = 1
  $sh.Range('F1').Value2 = 2
  $sh.Range('E2').Value2 = 3
  $sh.Range('F2').Value2 = 4

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★範囲を 渡した 時 実Excel は 何を 返すか★（64）（2026-09-20）')
  $行.Add('# ★なぜ★ ... Exally1 が 「105個 中 82個 足りない／11個は 気付けない」と 数えた')
  $行.Add('#           ⇒★但し 見たのは 私の 紙の `ref` の 字だけ＝中身は 未測定★')
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★どの 貝殻か★ ... PowerShell ' + $版.ToString())
  $行.Add('# ★盤面★ A1=1 ／ ★A2==1+1（式）★ ／ A3=2 ／ E1,F1,E2,F2=1,2,3,4')
  $行.Add('# ★置き場★ ... 行1（A1:A3 と 交わる）／列は 4つ おき（H L P T X AB AF AJ）')
  $行.Add('# ★読む★ ... 頭の マスから 6行 x 3列（★中身まで★）')

# ══ ★★`#材料` を 書きます（★これが 無いと 試験は この 紙を 1行も 押しません★）★★ ══
#   ＝2026-09-21 に 気付きました。上の `# ★盤面★` は ★人が 読む 字★で
#     ★機械は 読めません★（`tests/shiki-kansuu-kami.test.mjs` が 見るのは `#材料` の 行だけ）
#   ⇒★盤面を 人の 字と 機械の 字で ★二重に★ 書かない★＝下の 1つの 名簿から 両方 出します
#   ★A2 は ★値では なく 式★★ ... `=` で 始まる 物は 試験の 側も 式と して 打ちます
$行.Add('#材料' + "`t" + 'A1' + "`t" + '1')
$行.Add('#材料' + "`t" + 'A2' + "`t" + '=1+1')
$行.Add('#材料' + "`t" + 'A3' + "`t" + '2')
$行.Add('#材料' + "`t" + 'E1' + "`t" + '1')
$行.Add('#材料' + "`t" + 'F1' + "`t" + '2')
$行.Add('#材料' + "`t" + 'E2' + "`t" + '3')
$行.Add('#材料' + "`t" + 'F2' + "`t" + '4')
  $行.Add('# 組' + "`t" + '名前' + "`t" + 'マス' + "`t" + '式' + "`t" + '埋まった数' + "`t" + '中身（|で マス／ / で 行）')

  foreach ($x in $台本) {
    $ma = $x.頭
    $投げた = ''
    try { $sh.Range($ma).Formula2 = $x.式 } catch { $投げた = '★投げました★' }
    $頭列 = $sh.Range($x.頭).Column
    $並 = New-Object System.Collections.Generic.List[string]
    $埋 = 0
    for ($r = 0; $r -lt 6; $r++) {
      $一行 = New-Object System.Collections.Generic.List[string]
      for ($k = 0; $k -lt 4; $k++) {
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
    $行.Add($x.組 + "`t" + $x.札 + "`t" + $x.頭 + "`t" + $x.式 + "`t" + $埋 + "`t" + $出し + $投げた)
    Write-Host ('  ' + $x.組.PadRight(17) + ' ' + $x.札.PadRight(11) + ' 埋まった ' + ([string]$埋).PadRight(3) + ' ' + $出し)
  }

  # ══ ★門⑤＝対照が 待つ 通りか★ ══
  $行.Add('#')
  $行.Add('# ★★対照★★（★盤面が 狂って いない 事を 見ます★）')
  # SEQUENCE は H121 から 3マス（17本目）／SUM は H129（18本目）
  $s1 = [string]$sh.Range('P2').Value2     # SEQUENCE(2) の 2マス目
  $s2 = [string]$sh.Range('M5').Value2     # ISNUMBER(A1:A5) の 5マス目
  $行.Add('# 対照 SEQUENCE(2) の 2マス目 ... ' + $t1 + '（待つ 2）')
  $行.Add('# 対照 SUM(A1:A3) ............... ' + $t2 + '（待つ 5）')
  # ★★2026-09-21 ── ★対照の 読みを また 入れ忘れました（今日 2回目）★★
  #   ⇒★門が 0/2 で 止めた＝黙って 緑に ならなかった★
  $t1 = [string]$sh.Range('AF2').Value2
  $t2 = [string]$sh.Range('AJ1').Value2
  if ($t1 -eq '2') { $対照ok++ }
  if ($t2 -eq '5') { $対照ok++ }
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
