# toru-yosoku-3kansuu.ps1 — ★LOGEST / TREND / GROWTH を 実Excel に 聞く★（2026-09-09）
#
#  ★★なぜ★★
#    LINEST で ★x が 2本以上（重回帰）の 時に 静かに 違う 答え★を 出して いた。
#      =LINEST(G1:G6,H1:I6) … 実Excel 0.7708333333333329 ／ うち 7.390243902439025
#    この 3つも ★同じ 土台（最小二乗）★を 使う。
#    ⇒★同じ 穴が 在る 見込みが 高い★
#    ⇒★★見込みで 直さない＝先に 測る★★
#
#  ★測る 事★
#    ①x が 1本の 時 … 実Excel は 何を 返すか
#    ②★x が 2本（重回帰）の 時★ … 実Excel は 何を 返すか
#    ③表の 大きさ（ROWS / COLUMNS）… ★見た目で 数えない★
#    ④TREND / GROWTH の ★新しい x★ を 渡した 時
#
#  ★物差しの 決まり★
#    ・2つ目の 窓（0 が 出たら `=(式)=0` と 型で 疑う）
#    ・材料は 紙に 書き出す（使う 側が 手で 写さない）
#    ・表は =INDEX(式,行,列) で 1つずつ（溢れさせない）
#    ・書き戻しは LF
#
#  ★答えは 実Excel の 実測だけが 正★
#  ★司さんの 実物には 触りません★＝新しい ブック・保存せず
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-yosoku-3kansuu.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-yosoku-3kansuu-2026-09-09.tsv'

$誤りの番号 = @{
  -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!';
  -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!';
  -2146826246 = '#N/A';
  -2146826243 = '#SPILL!';   -2146826242 = '#CONNECT!'; -2146826241 = '#BLOCKED!';
  -2146826240 = '#UNKNOWN!'; -2146826239 = '#FIELD!';   -2146826238 = '#CALC!'; -2146826237 = '#BUSY!'
}

function 窓２_型($v) {
  if ($null -eq $v) { return 'Empty' }
  if ($v -is [string]) { return 'String' }
  if ($v -is [bool]) { return 'Boolean' }
  if ($v -is [double] -or $v -is [int] -or $v -is [long]) { return 'Number' }
  return 'Other'
}
function 窓２_本当にゼロか($sh, [string]$式) {
  $中 = $式 -replace '^=\s*', ''
  try {
    $sh.Range('BZ1').Clear() | Out-Null
    $sh.Range('BZ1').Formula = ('=(' + $中 + ')=0')
    $z = $sh.Range('BZ1').Value2
    $sh.Range('BZ1').Clear() | Out-Null
    if ($z -is [bool]) { return $(if ($z) { 'TRUE' } else { 'FALSE' }) }
    return '★判じられない★'
  } catch { return '★判じられない★' }
}

# ══ ★材料★ ══════════════════════════════════════════════
#  ★入れ子の 配列は PowerShell が 平らに する★ので ★列ごとの 表★に する
#  A/B … x 1本（掛け算で 伸びる＝LOGEST/GROWTH でも 意味の 在る y）
#  D/E,F … ★x 2本（重回帰）★
#  H … TREND / GROWTH に 渡す ★新しい x★
$材料 = [ordered]@{
  'A' = @(2, 4, 8, 16, 32, 64)          # y（2倍ずつ）
  'B' = @(1, 2, 3, 4, 5, 6)             # x 1本
  'D' = @(3, 7, 12, 20, 33, 54)         # y（重回帰の 為）
  'E' = @(1, 2, 3, 4, 5, 6)             # x 1本目
  'F' = @(2, 1, 4, 3, 6, 5)             # x 2本目
  'H' = @(7, 8)                          # 新しい x
}
foreach ($k in $材料.Keys) {
  if ($材料[$k].Count -lt 2) { Write-Error ('★列 ' + $k + ' が 短い★'); exit 2 }
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Worksheets.Item(1)
  foreach ($k in $材料.Keys) {
    for ($i = 0; $i -lt $材料[$k].Count; $i++) {
      $sh.Range($k + ($i + 1)).Value2 = [double]$材料[$k][$i]
    }
  }

  function 押して字に([string]$式) {
    $sh.Range('T1:AZ40').Clear() | Out-Null
    try {
      $sh.Range('T1').Formula = $式
      $v = $sh.Range('T1').Value2
      if ($null -eq $v) { return @{ 値 = '(空)'; 型 = 'Empty'; 窓2 = '—' } }
      if (($v -is [int] -or $v -is [long]) -and $誤りの番号.ContainsKey([int]$v)) {
        return @{ 値 = $誤りの番号[[int]$v]; 型 = 'Error'; 窓2 = '—' }
      }
      if ($v -is [bool]) { return @{ 値 = $(if ($v) { 'TRUE' } else { 'FALSE' }); 型 = 'Boolean'; 窓2 = '—' } }
      if (($v -is [double] -and $v -eq 0) -or ($v -is [string] -and $v -eq '0')) {
        return @{ 値 = '0'; 型 = (窓２_型 $v); 窓2 = (窓２_本当にゼロか $sh $式) }
      }
      if ($v -is [double]) { return @{ 値 = $v.ToString('R'); 型 = 'Double'; 窓2 = '—' } }
      if ($v -is [int] -or $v -is [long]) { return @{ 値 = [string]$v; 型 = 'Int32'; 窓2 = '—' } }
      return @{ 値 = [string]$v; 型 = 'String'; 窓2 = '—' }
    } catch { return @{ 値 = '★受け付けない★'; 型 = 'Rejected'; 窓2 = '—' } }
  }

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★LOGEST / TREND / GROWTH を 実Excel に 聞いた★（2026-09-09）')
  $行.Add('#')
  $行.Add('# ★なぜ★ LINEST が ★x 2本で 静かに 違う 答え★を 出して いた。')
  $行.Add('#   この 3つも 同じ 土台（最小二乗）＝★同じ 穴の 見込みが 高い★')
  $行.Add('#   ⇒★見込みで 直さない＝先に 測る★')
  $行.Add('#')
  $行.Add('# ★どの Excel で 打ったか★ … 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('#')
  $行.Add('# ★★材料（機械が 読む）★★ … `#材料<タブ>マス<タブ>値` の 形')
  foreach ($k in $材料.Keys) {
    for ($i = 0; $i -lt $材料[$k].Count; $i++) {
      $行.Add('#材料' + "`t" + ($k + ($i + 1)) + "`t" + ([double]$材料[$k][$i]).ToString('R'))
    }
  }
  $行.Add('#')
  $行.Add('# 関数' + "`t" + '式' + "`t" + '実Excel の 答え' + "`t" + '型' + "`t" + '窓２(=(式)=0)' + "`t" + '何を 見て いるか')

  $本数 = 0
  function 足す([string]$名, [string]$式, [string]$札) {
    $r = 押して字に $式
    $script:行.Add($名 + "`t" + $式 + "`t" + $r.値 + "`t" + $r.型 + "`t" + $r.窓2 + "`t" + $札)
    $script:本数++
    Write-Host ('  ' + $名.PadRight(7) + ' ' + $式.PadRight(42) + ' → ' + $r.値)
  }

  # ── ★LOGEST★ ──────────────────────────────────────
  足す 'LOGEST' '=LOGEST(A1:A6,B1:B6)'                     '★x 1本★・裸'
  足す 'LOGEST' '=ROWS(LOGEST(A1:A6,B1:B6))'               '表の 大きさ（ROWS）'
  足す 'LOGEST' '=COLUMNS(LOGEST(A1:A6,B1:B6))'            '表の 大きさ（COLUMNS）'
  足す 'LOGEST' '=INDEX(LOGEST(A1:A6,B1:B6),1,1)'          'x1本・1行1列'
  足す 'LOGEST' '=INDEX(LOGEST(A1:A6,B1:B6),1,2)'          'x1本・1行2列'
  足す 'LOGEST' '=LOGEST(D1:D6,E1:F6)'                     '★★x 2本（重回帰）★★・裸'
  足す 'LOGEST' '=ROWS(LOGEST(D1:D6,E1:F6))'               '★x2本★ ROWS'
  足す 'LOGEST' '=COLUMNS(LOGEST(D1:D6,E1:F6))'            '★x2本★ COLUMNS'
  足す 'LOGEST' '=INDEX(LOGEST(D1:D6,E1:F6),1,1)'          '★x2本★ 1行1列'
  足す 'LOGEST' '=INDEX(LOGEST(D1:D6,E1:F6),1,2)'          '★x2本★ 1行2列'
  足す 'LOGEST' '=INDEX(LOGEST(D1:D6,E1:F6),1,3)'          '★x2本★ 1行3列'

  # ── ★TREND★ ───────────────────────────────────────
  足す 'TREND' '=TREND(A1:A6,B1:B6)'                       '★x 1本★・裸'
  足す 'TREND' '=ROWS(TREND(A1:A6,B1:B6))'                 '表の 大きさ（ROWS）'
  足す 'TREND' '=COLUMNS(TREND(A1:A6,B1:B6))'              '表の 大きさ（COLUMNS）'
  足す 'TREND' '=INDEX(TREND(A1:A6,B1:B6),1,1)'            'x1本・1つ目'
  足す 'TREND' '=INDEX(TREND(A1:A6,B1:B6),6,1)'            'x1本・6つ目'
  足す 'TREND' '=TREND(A1:A6,B1:B6,H1:H2)'                 '★新しい x を 渡した★'
  足す 'TREND' '=INDEX(TREND(A1:A6,B1:B6,H1:H2),1,1)'      '★新しい x★ 1つ目'
  足す 'TREND' '=INDEX(TREND(A1:A6,B1:B6,H1:H2),2,1)'      '★新しい x★ 2つ目'
  足す 'TREND' '=TREND(D1:D6,E1:F6)'                       '★★x 2本（重回帰）★★・裸'
  足す 'TREND' '=INDEX(TREND(D1:D6,E1:F6),1,1)'            '★x2本★ 1つ目'
  足す 'TREND' '=INDEX(TREND(D1:D6,E1:F6),6,1)'            '★x2本★ 6つ目'

  # ── ★GROWTH★ ──────────────────────────────────────
  足す 'GROWTH' '=GROWTH(A1:A6,B1:B6)'                     '★x 1本★・裸'
  足す 'GROWTH' '=ROWS(GROWTH(A1:A6,B1:B6))'               '表の 大きさ（ROWS）'
  足す 'GROWTH' '=COLUMNS(GROWTH(A1:A6,B1:B6))'            '表の 大きさ（COLUMNS）'
  足す 'GROWTH' '=INDEX(GROWTH(A1:A6,B1:B6),1,1)'          'x1本・1つ目'
  足す 'GROWTH' '=INDEX(GROWTH(A1:A6,B1:B6),6,1)'          'x1本・6つ目'
  足す 'GROWTH' '=GROWTH(A1:A6,B1:B6,H1:H2)'               '★新しい x を 渡した★'
  足す 'GROWTH' '=INDEX(GROWTH(A1:A6,B1:B6,H1:H2),1,1)'    '★新しい x★ 1つ目'
  足す 'GROWTH' '=INDEX(GROWTH(A1:A6,B1:B6,H1:H2),2,1)'    '★新しい x★ 2つ目'
  足す 'GROWTH' '=GROWTH(D1:D6,E1:F6)'                     '★★x 2本（重回帰）★★・裸'
  足す 'GROWTH' '=INDEX(GROWTH(D1:D6,E1:F6),1,1)'          '★x2本★ 1つ目'
  足す 'GROWTH' '=INDEX(GROWTH(D1:D6,E1:F6),6,1)'          '★x2本★ 6つ目'

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★★書けた … ' + $本数 + '本★★')
  Write-Host ('★書いた … ' + $出 + '★')

  $bk.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
