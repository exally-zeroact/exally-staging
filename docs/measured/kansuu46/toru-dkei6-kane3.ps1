# toru-dkei6-kane3.ps1 — ★Excel の 枠 1回で 3つ 聞く★（2026-09-15）
#
#  ★★何を 聞くか★★
#    ㋐★棚㉝ の 6か所★ … 製品（`_jsDbFunc`）と 台（`D系の値たち`）が 違う 所
#         ②多列（かつ） ③多行（または） ④`>3` `<>x` ⑤field 超過
#         ⑥DPRODUCT 0件 ⑦DGET 0件
#       ⇒★どちらも 実Excel で 測って いません★＝★当て推量で 直さない★
#    ㋑★お金を「マスで 渡した 時」★ … ㉘の ★3通りの 3つ目★
#         数で 渡す（測り済・合う）／DATE() で 渡す（測り済・#VALUE!）／★マス（未測定）★
#    ㋒★YEARFRAC ほか★ … 他の 日付を 取る 関数も 同じか（棚㉘の「見て いない 事」）
#
#  ★★司さんの 実物ブックは 開きません★★＝★新しい 空の ブックだけ★
#  ★走らせる 前に Excel が 動いて いないか 2つの 道具で 数えます★（1個でも 居たら 走らせない）
#
#  ★材料★（`kansuu46` の 紙と 同じ／★式は 材料の 外 J1★）
#    A1:A5 = 1,2,3,4,5 ／ B1:B5 = 2,4,6,8,10
#    D1 = 45292（=DATE(2024,1,1)）／D2 = 46023（=DATE(2026,1,1)）
#    F1:G3 = ★条件の 置き場★（1件ごとに 消してから 書く）
#    H1:H3 = 39508 / 39691 / 39569（＝DATE(2008,3,1) / (2008,8,31) / (2008,5,1)）
#
#  ★出す 物★ docs/measured/kansuu46/golden-dkei6-kane3-2026-09-15.tsv
#    ★「うちの 答え」の 列は 入れません★（混ぜると 前後の 突き合わせが 出来なく なる）
#
#  使い方: powershell -NoProfile -ExecutionPolicy Bypass -File docs\measured\kansuu46\toru-dkei6-kane3.ps1
#          （取り直す 時だけ -上書き）
param([switch]$上書き, [string]$出)

$ErrorActionPreference = 'Stop'

# ══ ★門①★ Excel が 動いて いないか ★2つの 道具で★ 数える ══
#   ★1個でも 居たら 走らせない★＝★司さんが 開いて いるかも しれない★
$道具1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$道具2 = @(tasklist /FI "IMAGENAME eq EXCEL.EXE" /NH 2>$null | Where-Object { $_ -match 'EXCEL' }).Count
Write-Host ("★Excel の 数★ Get-Process {0}個 ／ tasklist {1}個" -f $道具1, $道具2)
if ($道具1 -ne 0 -or $道具2 -ne 0) {
  throw "★★走らせません★★＝Excel が 動いて います（$道具1 / $道具2）。★司さんが 開いて いるかも しれません★"
}

# ══ ★物差しの 2つ目の 窓★ ══
#   ★`.Value2` は ★0 で ない 値に 0 を 返す★事が 在ります★
#     `=0.1+0.2-0.3` … `.Value2` ★0★ ／ `=(式)=0` ★False★
#   ⇒★「0」が 出た 時だけ ★`=(式)=0` の 真偽★も 取る★
#   ⇒ 見張り `tests/monosashi-mado.test.mjs` が これの 無い 道具を 赤に します
function 窓２_本当にゼロか($sh, $xl2, [string]$式) {
  $中 = $式 -replace '^=\s*', ''
  try {
    $sh.Range('BZ1').Clear() | Out-Null
    $sh.Range('BZ1').Formula = ('=(' + $中 + ')=0')
    $xl2.CalculateFull()
    $z = $sh.Range('BZ1').Value2
    $sh.Range('BZ1').Clear() | Out-Null
    if ($z -is [bool]) { return $(if ($z) { 'TRUE' } else { 'FALSE' }) }
    return '★判じられない★'
  } catch { return '★判じられない★' }
}

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $出) { $出 = Join-Path $ここ 'golden-dkei6-kane3-2026-09-15.tsv' }
elseif (-not [System.IO.Path]::IsPathRooted($出)) { $出 = Join-Path $ここ $出 }
if ((Test-Path $出) -and -not $上書き) {
  throw "★もう 在ります★: $出 ／ 取り直すなら -上書き を 付けて ください"
}

$誤りの番号 = @{
  -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!';
  -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!';
  -2146826246 = '#N/A';
  -2146826243 = '#SPILL!';   -2146826242 = '#CONNECT!'; -2146826241 = '#BLOCKED!';
  -2146826240 = '#UNKNOWN!'; -2146826239 = '#FIELD!';   -2146826238 = '#CALC!'; -2146826237 = '#BUSY!'
}

# ══ ★聞く 事（★人が 決めるのは ここだけ★）══
#   置き … 条件の 置き場に 何を 書くか（F1:G3 は 毎回 消してから 書く）
$問い = @(
  @{ 札 = '02多列(かつ)・両方 合う'; 置き = [ordered]@{ 'F1' = 1; 'G1' = 2; 'F2' = 2; 'G2' = 4 }; 式 = '=DSUM(A1:B5,1,F1:G2)' },
  @{ 札 = '02多列(かつ)・片方 外れ'; 置き = [ordered]@{ 'F1' = 1; 'G1' = 2; 'F2' = 2; 'G2' = 99 }; 式 = '=DSUM(A1:B5,1,F1:G2)' },
  @{ 札 = '02多列・2列目の 見出しが 台帳に 無い'; 置き = [ordered]@{ 'F1' = 1; 'G1' = 99; 'F2' = 2; 'G2' = 4 }; 式 = '=DSUM(A1:B5,1,F1:G2)' },
  @{ 札 = '03多行(または)'; 置き = [ordered]@{ 'F1' = 1; 'F2' = 2; 'F3' = 3 }; 式 = '=DSUM(A1:B5,1,F1:F3)' },
  @{ 札 = '04記号 大なり3'; 置き = [ordered]@{ 'F1' = 1; 'F2' = '>3' }; 式 = '=DSUM(A1:B5,1,F1:F2)' },
  @{ 札 = '04記号 等しくない2'; 置き = [ordered]@{ 'F1' = 1; 'F2' = '<>2' }; 式 = '=DSUM(A1:B5,1,F1:F2)' },
  @{ 札 = '04記号 以上3'; 置き = [ordered]@{ 'F1' = 1; 'F2' = '>=3' }; 式 = '=DSUM(A1:B5,1,F1:F2)' },
  @{ 札 = '05field 超過(5)'; 置き = [ordered]@{ 'F1' = 1; 'F2' = 2 }; 式 = '=DSUM(A1:B5,5,F1:F2)' },
  @{ 札 = '05field 0'; 置き = [ordered]@{ 'F1' = 1; 'F2' = 2 }; 式 = '=DSUM(A1:B5,0,F1:F2)' },
  @{ 札 = '05field マイナス1'; 置き = [ordered]@{ 'F1' = 1; 'F2' = 2 }; 式 = '=DSUM(A1:B5,-1,F1:F2)' },
  @{ 札 = '06DPRODUCT 0件'; 置き = [ordered]@{ 'F1' = 1; 'F2' = 99 }; 式 = '=DPRODUCT(A1:B5,1,F1:F2)' },
  @{ 札 = '06DSUM 0件'; 置き = [ordered]@{ 'F1' = 1; 'F2' = 99 }; 式 = '=DSUM(A1:B5,1,F1:F2)' },
  @{ 札 = '06DCOUNT 0件'; 置き = [ordered]@{ 'F1' = 1; 'F2' = 99 }; 式 = '=DCOUNT(A1:B5,1,F1:F2)' },
  @{ 札 = '07DGET 0件'; 置き = [ordered]@{ 'F1' = 1; 'F2' = 99 }; 式 = '=DGET(A1:B5,1,F1:F2)' },
  @{ 札 = '07DGET 2件以上'; 置き = [ordered]@{ 'F1' = 1; 'F2' = '>0' }; 式 = '=DGET(A1:B5,1,F1:F2)' },
  @{ 札 = '条件の 見出しが 空'; 置き = [ordered]@{ 'F2' = 2 }; 式 = '=DSUM(A1:B5,1,F1:F2)' },
  @{ 札 = '条件の マスが 空'; 置き = [ordered]@{ 'F1' = 1 }; 式 = '=DSUM(A1:B5,1,F1:F2)' },
  @{ 札 = 'kane マスで 渡す(ACCRINT)'; 置き = [ordered]@{}; 式 = '=ACCRINT(H1,H2,H3,0.1,1000,2,0)' },
  @{ 札 = 'kane 数で 渡す(ACCRINT)'; 置き = [ordered]@{}; 式 = '=ACCRINT(39508,39691,39569,0.1,1000,2,0)' },
  @{ 札 = 'kane DATEで 渡す(ACCRINT)'; 置き = [ordered]@{}; 式 = '=ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0)' },
  @{ 札 = 'kane マスで 渡す(COUPNUM)'; 置き = [ordered]@{}; 式 = '=COUPNUM(H1,H2,2,0)' },
  @{ 札 = 'hidzuke YEARFRAC DATE'; 置き = [ordered]@{}; 式 = '=YEARFRAC(DATE(2008,3,1),DATE(2008,8,31),0)' },
  @{ 札 = 'hidzuke YEARFRAC マス'; 置き = [ordered]@{}; 式 = '=YEARFRAC(H1,H2,0)' },
  @{ 札 = 'hidzuke YEARFRAC 数'; 置き = [ordered]@{}; 式 = '=YEARFRAC(39508,39691,0)' },
  @{ 札 = 'hidzuke DATEDIF マス'; 置き = [ordered]@{}; 式 = '=DATEDIF(H1,H2,"D")' },
  @{ 札 = 'hidzuke EDATE マス'; 置き = [ordered]@{}; 式 = '=EDATE(H1,1)' }
)

$xl = $null; $wb = $null; $ws = $null
$結果 = New-Object System.Collections.ArrayList
$材料の行 = New-Object System.Collections.ArrayList
$版 = ''; $ビルド = ''
$タブ = [string][char]9
try {
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $版 = $xl.Version
  $ビルド = $xl.Build
  $wb = $xl.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)

  # ★材料★（★書式は 触りません★＝NumberFormat='@' は 式の マスに 当てない）
  $ws.Range('A1').Value2 = 1;  $ws.Range('A2').Value2 = 2
  $ws.Range('A3').Value2 = 3;  $ws.Range('A4').Value2 = 4;  $ws.Range('A5').Value2 = 5
  $ws.Range('B1').Value2 = 2;  $ws.Range('B2').Value2 = 4
  $ws.Range('B3').Value2 = 6;  $ws.Range('B4').Value2 = 8;  $ws.Range('B5').Value2 = 10
  $ws.Range('D1').Value2 = 45292; $ws.Range('D2').Value2 = 46023
  $ws.Range('H1').Value2 = 39508; $ws.Range('H2').Value2 = 39691; $ws.Range('H3').Value2 = 39569

  $i = 0
  foreach ($q in $問い) {
    $i++
    Write-Host ("  … {0}/{1} {2}" -f $i, $問い.Count, $q.札)
    # ★前の 跡を 消す★（条件の 置き場と 式の マス）
    $ws.Range('F1:G5').Clear() | Out-Null
    $ws.Range('J1:Z50').Clear() | Out-Null
    # ★添字で 引かない★＝`[ordered]@{}` の 添字は ★番号とも 鍵とも 取れる★
    #   ⇒ 2026-09-15 実測 … 値が `'>3'` の 時だけ ★InvalidCastException★
    #   （`.Value2 = '>3'` を 直に 書けば 通ります＝★Excel の せいでは ない★）
    # ★★型を 書かずに 渡すと 字だけ 落ちます★★（2026-09-15 実測・3回 転んだ）
    #   `$e.Value` は ★`Object` の 顔★で 出て 来ます
    #   ⇒ COM が 相手の 形を 決められず ★InvalidCastException★
    #   ★実測★ `'>3'` … `= $e.Value` ★NG★ ／ `= [string]$e.Value` ★OK★
    #          `1`    … どちらも OK  ⇒★★数では 出ない＝字の 時だけ 転ぶ★★
    #   （`.Value2 = '>3'` を 直に 書けば 通ります＝★Excel の せいでは ない★）
    foreach ($e in $q.置き.GetEnumerator()) {
      $マス = $ws.Range([string]$e.Key)
      if ($e.Value -is [string]) { $マス.Value2 = [string]$e.Value }
      else { $マス.Value2 = [double]$e.Value }
    }
    $ws.Range('J1').Formula = $q.式
    $xl.CalculateFull()
    $v = $ws.Range('J1').Value2
    if ($null -eq $v) {
      [void]$結果.Add(($q.札 + $タブ + $q.式 + $タブ + '' + $タブ + '空'))
      continue
    }
    if (($v -is [int] -or $v -is [double]) -and $誤りの番号.ContainsKey([int]$v)) {
      [void]$結果.Add(($q.札 + $タブ + $q.式 + $タブ + $誤りの番号[[int]$v] + $タブ + 'error値'))
      continue
    }
    $型 = $v.GetType().Name
    $答 = if ($v -is [double]) { $v.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    # ★「0」は 1つの 窓では 決めない★（2つ目の 窓＝`=(式)=0`）
    $窓２ = ''
    if ($答 -eq '0') { $窓２ = 窓２_本当にゼロか $ws $xl $q.式 }
    [void]$結果.Add(($q.札 + $タブ + $q.式 + $タブ + $答 + $タブ + $型 + $タブ + $窓２))
  }

  # ★材料を 機械が 読める 形で 控える★（★Excel から 読み返す＝私が 計算しない★）
  $ws.Range('F1:G5').Clear() | Out-Null
  $ws.Range('J1:Z50').Clear() | Out-Null
  foreach ($マス in @('A1', 'A2', 'A3', 'A4', 'A5', 'B1', 'B2', 'B3', 'B4', 'B5', 'D1', 'D2', 'H1', 'H2', 'H3')) {
    $mv = $ws.Range($マス).Value2
    if ($null -eq $mv) { [void]$材料の行.Add(('#材料' + $タブ + $マス + $タブ + '' + $タブ + '空')); continue }
    if ($mv -is [string]) { [void]$材料の行.Add(('#材料' + $タブ + $マス + $タブ + $mv + $タブ + '字')); continue }
    $ms = if ($mv -is [double]) { $mv.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture) } else { [string]$mv }
    [void]$材料の行.Add(('#材料' + $タブ + $マス + $タブ + $ms + $タブ + '数'))
  }
}
finally {
  # ★必ず 閉じる★
  if ($wb) { try { $wb.Close($false) } catch {} }
  if ($xl) { try { $xl.Quit() } catch {} }
  foreach ($o in @($ws, $wb, $xl)) {
    if ($o) { try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null } catch {} }
  }
  [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
  # ★★本当に 消えたか を 待って 確かめる★★（1点で 見ない・0.5秒おき）
  #   ★★10秒では 足りませんでした★★（2026-09-15 実測）
  #     26件 押した 後 … ★10秒で まだ 1個★／その後 数えたら ★0個★
  #     ⇒★「残って います」と 出たのに 残って いなかった＝★嘘の 赤★★
  #   ⇒★60秒まで 待つ★／★それでも 残って いたら 本当に 言う★
  $残り = @(); $t = 0
  for ($t = 0; $t -lt 120; $t++) {
    Start-Sleep -Milliseconds 500
    $残り = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
    if ($残り.Count -eq 0) { break }
  }
  if ($残り.Count -eq 0) { Write-Host ('★Excel は 残って いません（0個・' + (($t + 1) * 0.5) + '秒で 消えた）★') }
  else { Write-Host ('★★Excel が ' + $残り.Count + '個 残って います＝10秒 待っても 消えません＝手で 止めて ください★★') }
}

$頭 = @(
  '# ★棚㉝ の 6か所＋お金の 3通り＋日付の 関数 を 実Excel に 聞いた★',
  '# 取った 日 … 2026-09-15',
  ('# ★どの Excel で 打ったか★ … 版 ' + $版 + ' ／ build ' + $ビルド),
  '# ★どの 道で 読んだか★ … .Value2（誤りは error値 として 番号から 名前に した）',
  '#   数は R（丸めない 書き方）で 出して います',
  '# ★★「うちの 答え」の 列は 入れて いません★★（混ぜると 前後の 突き合わせが 出来なく なる）',
  '# ★取った 道★ … docs/measured/kansuu46/toru-dkei6-kane3.ps1',
  '# ★2つ目の 窓★ … 答えが 0 の 行だけ `=(式)=0` の 真偽を 6列目に 入れて います',
  ('# 札' + $タブ + '式' + $タブ + '実Excel の 答え' + $タブ + '型' + $タブ + '2つ目の窓')
)
($頭 + $材料の行 + $結果) -join [System.Environment]::NewLine | Set-Content -Path $出 -Encoding UTF8
Write-Host ('★書きました★ ' + $出 + ' … ' + $結果.Count + '行（材料 ' + $材料の行.Count + '行）')
