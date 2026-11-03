# toru-linest-hyou.ps1 — ★LINEST が ★何を 返すか★を 実Excel に 全部 聞く★（2026-09-09）
#
#  ★★なぜ★★（2件目）
#    うちは `=LINEST(範囲,範囲)` の ★裸の 形だけ★を JS層で 拾い、★傾き 1つ★を 返して います。
#    ⇒ `=INDEX(LINEST(…),1,1)` は ★#NAME?★（エンジンに 積んで いない）
#    ⇒★★「在るように 見えて 実は 無い」★★
#    ⇒ 直す 前に ★実Excel が 本当は 何を 返すか★を 全部 測る。
#
#  ★★LINEST は 表を 返す★★（1つの 数では ない）
#    第3引数 定数（既定 TRUE）／第4引数 補正（既定 FALSE）
#    補正を TRUE に すると ★5行★ 返る と 言われて いますが、
#    ★言われて いる を 写さず、★実Excel に 打って 数える★★
#
#  ★物差しの 決まり★
#    ・★2つ目の 窓★（0 が 出たら `=(式)=0` と 型で 疑う）
#    ・★材料は 紙に 書き出す★（使う 側が 手で 写さない）
#    ・★表を 返す 式は =INDEX(式,行,列) で 1つずつ★（★溢れさせない★）
#    ・★行数・列数は ROWS/COLUMNS で 別に 聞く★（見た目で 数えない）
#    ・★書き戻しは LF★
#
#  ★答えは 実Excel の 実測だけが 正★（AI に 出させない）
#  ★司さんの 実物には 触りません★＝新しい ブック・保存せず
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-linest-hyou.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-linest-hyou-2026-09-09.tsv'

$誤りの番号 = @{
  -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!';
  -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!';
  -2146826246 = '#N/A';
  -2146826243 = '#SPILL!';   -2146826242 = '#CONNECT!'; -2146826241 = '#BLOCKED!';
  -2146826240 = '#UNKNOWN!'; -2146826239 = '#FIELD!';   -2146826238 = '#CALC!'; -2146826237 = '#BUSY!'
}

# ══ ★2つ目の 窓★ ══════════════════════════════════════
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

# ══ ★材料（★1つの 組で 決めない＝形の 違う 4組★）★ ══════════════
#   1) x が 1本・きれいに 乗る
#   2) x が 1本・端数が 出る
#   3) ★x が 2本（重回帰）★＝★列の 数が 変わる★
#   4) ★点が 2つだけ★（一番 短い＝境目）
#  ★★PowerShell は 入れ子の 配列を ★平らに する★★★（2026-09-09 に 踏んだ）
#    `x = @(@(1,2,3,4,5))` と 書くと `x` は `@(1,2,3,4,5)` に なり、
#    `x[0]` は ★数の 1★（配列では ない）／`1[$i]` は i=0 で 1・それ以外は 空
#    ⇒★B1=1・B2〜B5=0 が 敷かれ、傾きが 20 では なく −50 に なって いた★
#    ⇒★★紙に 書き出した `#材料` が それを 捕まえた★★（そのために 足した 物）
#    ⇒★入れ子を 使わず ★列ごとの 表（ハッシュ）★に する＝平らに ならない★
$組 = @(
  @{ 札 = 'x1本・きれい';   y列 = 'A'; y = @(100, 120, 140, 160, 180);
     x列 = @('B'); xの中 = @{ 'B' = @(1, 2, 3, 4, 5) } },
  @{ 札 = 'x1本・端数';     y列 = 'D'; y = @(2.7, 5.1, 6.9, 9.4, 11.2);
     x列 = @('E'); xの中 = @{ 'E' = @(1, 2, 3, 4, 5) } },
  @{ 札 = '★x2本（重回帰）★'; y列 = 'G'; y = @(10, 14, 21, 25, 33, 38);
     x列 = @('H', 'I'); xの中 = @{ 'H' = @(1, 2, 3, 4, 5, 6); 'I' = @(2, 1, 4, 3, 6, 5) } },
  @{ 札 = '★点が 2つだけ★';  y列 = 'K'; y = @(3, 7);
     x列 = @('L'); xの中 = @{ 'L' = @(1, 2) } }
)
# ★★横1行の 組（★縦だけを 見て 決めない★）★★（2026-09-09 に 足した）
#   ★x が 1本かどうかは 縦1列 だけで なく ★横1行★も 在る★
#   ⇒ 上の 表は 縦向きなので、横向きは 別に 敷いて 別に 測る
#   N1:S1 ＝ y（横）／ N2:S2 ＝ x（横）
$横 = @{ y = @(100, 120, 140, 160, 180, 200); x = @(1, 2, 3, 4, 5, 6); 列 = @('N', 'O', 'P', 'Q', 'R', 'S') }
if ($横.y.Count -ne $横.列.Count -or $横.x.Count -ne $横.列.Count) {
  Write-Error '★横の 組の 数が 合わない★'; exit 2
}
# ★★敷く 前に 数を 確かめる（★平らに なって いたら ここで 止まる★）★★
foreach ($c in $組) {
  foreach ($col in $c.x列) {
    $v = $c.xの中[$col]
    if ($null -eq $v) { Write-Error ('★列 ' + $col + ' の 中身が 無い★'); exit 2 }
    if ($v.Count -ne $c.y.Count) {
      Write-Error ('★' + $c.札 + ' の 列 ' + $col + ' … x が ' + $v.Count + '個 ／ y が ' + $c.y.Count + '個＝★数が 合わない★')
      exit 2
    }
  }
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Worksheets.Item(1)

  # ★材料を 敷く★
  foreach ($c in $組) {
    for ($i = 0; $i -lt $c.y.Count; $i++) {
      $sh.Range($c.y列 + ($i + 1)).Value2 = [double]$c.y[$i]
      foreach ($col in $c.x列) {
        $sh.Range($col + ($i + 1)).Value2 = [double]$c.xの中[$col][$i]
      }
    }
  }
  # ★横1行の 組を 敷く★
  for ($i = 0; $i -lt $横.列.Count; $i++) {
    $sh.Range($横.列[$i] + '1').Value2 = [double]$横.y[$i]
    $sh.Range($横.列[$i] + '2').Value2 = [double]$横.x[$i]
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
      # ★『0』が 出たら 2つ目の 窓★
      if (($v -is [double] -and $v -eq 0) -or ($v -is [string] -and $v -eq '0')) {
        return @{ 値 = '0'; 型 = (窓２_型 $v); 窓2 = (窓２_本当にゼロか $sh $式) }
      }
      if ($v -is [double]) { return @{ 値 = $v.ToString('R'); 型 = 'Double'; 窓2 = '—' } }
      if ($v -is [int] -or $v -is [long]) { return @{ 値 = [string]$v; 型 = 'Int32'; 窓2 = '—' } }
      return @{ 値 = [string]$v; 型 = 'String'; 窓2 = '—' }
    } catch { return @{ 値 = '★受け付けない★'; 型 = 'Rejected'; 窓2 = '—' } }
  }

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★LINEST が 何を 返すか 実Excel に 全部 聞いた★（2026-09-09・2件目）')
  $行.Add('#')
  $行.Add('# ★なぜ★ うちは 裸の =LINEST(範囲,範囲) だけ 拾い ★傾き 1つ★しか 返して いない')
  $行.Add('#   ⇒ =INDEX(LINEST(…),1,1) は ★#NAME?★＝★在るように 見えて 実は 無い★')
  $行.Add('#   ⇒★直す 前に 実Excel が 本当は 何を 返すかを 測る★')
  $行.Add('#')
  $行.Add('# ★どの Excel で 打ったか★ … 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★物差し★ 表は =INDEX(式,行,列) で 1つずつ（★溢れさせない★）')
  $行.Add('#           行数・列数は ROWS/COLUMNS で ★別に 聞く★（見た目で 数えない）')
  $行.Add('#           『0』が 出たら ★=(式)=0 と 型★で 疑う')
  $行.Add('#')
  $行.Add('# ★★材料（機械が 読む）★★ … `#材料<タブ>マス<タブ>値` の 形')
  foreach ($c in $組) {
    for ($i = 0; $i -lt $c.y.Count; $i++) {
      $行.Add('#材料' + "`t" + ($c.y列 + ($i + 1)) + "`t" + ([double]$c.y[$i]).ToString('R'))
      foreach ($col in $c.x列) {
        $行.Add('#材料' + "`t" + ($col + ($i + 1)) + "`t" + ([double]$c.xの中[$col][$i]).ToString('R'))
      }
    }
  }
  for ($i = 0; $i -lt $横.列.Count; $i++) {
    $行.Add('#材料' + "`t" + ($横.列[$i] + '1') + "`t" + ([double]$横.y[$i]).ToString('R'))
    $行.Add('#材料' + "`t" + ($横.列[$i] + '2') + "`t" + ([double]$横.x[$i]).ToString('R'))
  }
  $行.Add('#')
  $行.Add('# 組' + "`t" + '式' + "`t" + '実Excel の 答え' + "`t" + '型' + "`t" + '窓２(=(式)=0)' + "`t" + '何の 場所')

  $本数 = 0
  foreach ($c in $組) {
    $n = $c.y.Count
    $y = $c.y列 + '1:' + $c.y列 + $n
    $x = $c.x列[0] + '1:' + $c.x列[-1] + $n     # 2本の 時は H1:I6 の ように つながる

    # ★① 裸の 形（うちが 今 拾って いる 形）★
    $r = 押して字に ('=LINEST(' + $y + ',' + $x + ')')
    $行.Add($c.札 + "`t" + '=LINEST(' + $y + ',' + $x + ')' + "`t" + $r.値 + "`t" + $r.型 + "`t" + $r.窓2 + "`t" + '裸（★うちが 拾う 形★）')
    $本数++

    # ★② 表の 大きさ（★見た目で 数えない★）★
    foreach ($sw in @('', ',TRUE,TRUE', ',TRUE,FALSE', ',FALSE,TRUE', ',FALSE,FALSE')) {
      $中 = 'LINEST(' + $y + ',' + $x + $sw + ')'
      foreach ($f in @('ROWS', 'COLUMNS')) {
        $r = 押して字に ('=' + $f + '(' + $中 + ')')
        $行.Add($c.札 + "`t" + '=' + $f + '(' + $中 + ')' + "`t" + $r.値 + "`t" + $r.型 + "`t" + $r.窓2 + "`t" + '表の 大きさ（' + $f + '）')
        $本数++
      }
    }

    # ★③ 表の 中身を 1マスずつ★（補正 TRUE ＝ 一番 大きい 形）
    $中 = 'LINEST(' + $y + ',' + $x + ',TRUE,TRUE)'
    $列数 = $c.x列.Count + 1
    for ($rr = 1; $rr -le 5; $rr++) {
      for ($cc = 1; $cc -le $列数; $cc++) {
        $式 = '=INDEX(' + $中 + ',' + $rr + ',' + $cc + ')'
        $r = 押して字に $式
        $行.Add($c.札 + "`t" + $式 + "`t" + $r.値 + "`t" + $r.型 + "`t" + $r.窓2 + "`t" + ('補正TRUE の ' + $rr + '行' + $cc + '列'))
        $本数++
      }
    }

    # ★④ 補正なし（既定）の 中身★
    $中 = 'LINEST(' + $y + ',' + $x + ')'
    for ($cc = 1; $cc -le $列数; $cc++) {
      $式 = '=INDEX(' + $中 + ',1,' + $cc + ')'
      $r = 押して字に $式
      $行.Add($c.札 + "`t" + $式 + "`t" + $r.値 + "`t" + $r.型 + "`t" + $r.窓2 + "`t" + ('既定の 1行' + $cc + '列'))
      $本数++
    }

    # ★⑤ 範囲の 外を 指したら 何が 返るか（★境目★）★
    foreach ($p in @('6,1', '1,9', '0,1', '1,0')) {
      $式 = '=INDEX(LINEST(' + $y + ',' + $x + ',TRUE,TRUE),' + $p + ')'
      $r = 押して字に $式
      $行.Add($c.札 + "`t" + $式 + "`t" + $r.値 + "`t" + $r.型 + "`t" + $r.窓2 + "`t" + '★表の 外を 指した（境目）★')
      $本数++
    }

    Write-Host ('  ' + $c.札 + ' … 済')
  }

  # ══ ★横1行の 組を 押す（★縦だけを 見て 決めない★）★ ══
  foreach ($f in @('=LINEST(N1:S1,N2:S2)',
                   '=ROWS(LINEST(N1:S1,N2:S2))',
                   '=COLUMNS(LINEST(N1:S1,N2:S2))',
                   '=INDEX(LINEST(N1:S1,N2:S2),1,1)',
                   '=INDEX(LINEST(N1:S1,N2:S2),1,2)')) {
    $r = 押して字に $f
    $行.Add('★横1行★' + "`t" + $f + "`t" + $r.値 + "`t" + $r.型 + "`t" + $r.窓2 + "`t" + '★横1行（x も 1本）★')
    $本数++
  }
  Write-Host '  ★横1行★ … 済'

  # ★★書き戻しは LF★★（WriteAllLines は Windows で CRLF に なる）
  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★★書けた … ' + $本数 + '本★★')
  Write-Host ('★書いた … ' + $出 + '★')

  $bk.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
