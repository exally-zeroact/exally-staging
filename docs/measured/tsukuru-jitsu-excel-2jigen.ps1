# tsukuru-jitsu-excel-2jigen.ps1
#   -- ★2次元の 範囲を 渡した 時 実Excel は 何を 返すか★（66）（2026-09-21）
#
#  ★★なぜ★★
#    Exally1 が `lib/shiki-kansuu.js` の `見分け` を ★1次元だけ 溢れる★形に 直しました（PR #92）。
#    ⇒★2次元は 1文字も 触って いません★
#    ⇒★訳★ ... `tests/shiki-kansuu-kami.test.mjs` が 3本 赤に なった
#              `=ISERR(A1:B5)` ⇒ True ／ `=ISERROR(A1:B5)` ⇒ True ／ `=ISNUMBER(A1:B5)` ⇒ False
#              紙 `docs/measured/kansuu46/golden-346-2026-09-08.tsv`
#    ⇒★★但し それは 2026-09-08 の 紙です★★
#      ＝★2次元を 「動く並び」で 打った 時は まだ 誰も 測って いません★
#    ⇒★もし 2次元も 溢れるなら 直しは 足りて いません★
#    ⇒★もし 誤りなら 今の ままで 正しい★
#
#  ★★盤面（★2026-09-08 の 紙と 同じ★）★★
#    A1:A5 = 1 / 2 / 3 / 4 / 5
#    B1:B5 = 2 / 4 / 6 / 8 / 10
#
#  ★★式を どこに 置くか（★ここが 肝★）★★
#    ★暗黙の 交わりは ★式を 置いた 行／列★で 決まります★
#    ⇒★A1:B5 と 交わる 行（行1）に 置きます★
#    ⇒★昨日 Exally1 が 行100 に 置いて 数を 逆に 出した 所です★
#    ⇒★2次元が 溢れると 5行 x 2列 使う ので 列を 3つ おきに します★（D / G / J / M / P）
#
#  ★★台本★★
#    ★2次元★ D1 `=ISNUMBER(A1:B5)` ／ G1 `=ISERR(A1:B5)` ／ J1 `=ISERROR(A1:B5)`
#    ★対照・1次元★ M1 `=ISNUMBER(A1:A5)`（★PR #92 で 溢れる ように した 形★）
#    ★対照・溢れる★ P1 `=SEQUENCE(2)`
#
#  ★★読む 物★★ 頭の マスから ★6行 x 3列★（★埋まった 数★と ★中身★）
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③台本の 数 決め打ち（exit 4）／④外へ 出る 6個と ASCII の 外の 字が 0件（exit 5）
#    ⑤★対照 2個が 待つ 通りで なければ 落とす（exit 6）★
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-jitsu-excel-2jigen-2026-09-21.tsv'

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

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null; $c = $null
$対照ok = 0
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)
  # ★盤面＝2026-09-08 の 紙と 同じ★
  for ($r = 1; $r -le 5; $r++) {
    $sh.Cells.Item($r, 1).Value2 = $r
    $sh.Cells.Item($r, 2).Value2 = ($r * 2)
  }

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★範囲を 渡した 時 実Excel は 何を 返すか★（64）（2026-09-20）')
  $行.Add('# ★なぜ★ ... Exally1 が 「105個 中 82個 足りない／11個は 気付けない」と 数えた')
  $行.Add('#           ⇒★但し 見たのは 私の 紙の `ref` の 字だけ＝中身は 未測定★')
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★どの 貝殻か★ ... PowerShell ' + $版.ToString())
  $行.Add('# ★盤面★ A1:A5=1,2,3,4,5 ／ B1:B5=2,4,6,8,10（★2026-09-08 の 紙と 同じ★）')
  $行.Add('# ★置き場★ ... ★A1:B5 と 交わる 行（行1）★／列は 3つ おき（D G J M P）')
  $行.Add('# ★読む★ ... 頭の マスから 6行 x 3列（★中身まで★）')
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
    $行.Add($x.組 + "`t" + $x.札 + "`t" + $x.頭 + "`t" + $x.式 + "`t" + $埋 + "`t" + $出し + $投げた)
    Write-Host ('  ' + $x.組.PadRight(17) + ' ' + $x.札.PadRight(11) + ' 埋まった ' + ([string]$埋).PadRight(3) + ' ' + $出し)
  }

  # ══ ★門⑤＝対照が 待つ 通りか★ ══
  $行.Add('#')
  $行.Add('# ★★対照★★（★盤面が 狂って いない 事を 見ます★）')
  # SEQUENCE は H121 から 3マス（17本目）／SUM は H129（18本目）
  $s1 = [string]$sh.Range('P2').Value2     # SEQUENCE(2) の 2マス目
  $s2 = [string]$sh.Range('M5').Value2     # ISNUMBER(A1:A5) の 5マス目
  $行.Add('# SEQUENCE(2) の 2マス目 ... ' + $s1 + '（待つ 2）')
  $行.Add('# ISNUMBER(A1:A5) の 5マス目 ... ' + $s2 + '（待つ True）')
  if ($s1 -eq '2') { $対照ok++ }
  if ($s2 -eq 'True') { $対照ok++ }
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
