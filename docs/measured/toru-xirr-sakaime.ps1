# toru-xirr-sakaime.ps1 — ★XIRR が どこで #NUM! に なるかを 実Excel に 聞く★（2026-09-09）
#
#  ★★なぜ★★
#    実Excel が ★#NUM!★ を 返す 組で、うちは ★数を 返して います★
#      =XIRR(全部プラス)   … 実Excel #NUM! ／ うち ★375686054239746.3★
#      =XIRR(日付が 逆順)  … 実Excel #NUM! ／ うち ★-0.5002981524324954★
#    ⇒★誤りに ならず 数が 出る＝お客さんは 気づけない★（★お金の 利回り★）
#    ⇒ 375686054239746.3 は ★桁が おかしいのに「変」と 思えない 顔★を して います
#
#  ★★どこから 断るかは ★私が 決めません＝実Excel に 聞きます★★★
#    「符号が 両方 要る」「日付は 昇順」等は ★聞いた 話★です。
#    ⇒★書いて ある事を 写さず、★境目の 前後を 両方★ 打って 確かめる★
#
#  ★測る 境目（★1つの 組で 決めない★）★
#    ①全部 プラス ／ ②全部 マイナス ／ ③プラスと マイナスが 1つずつ（一番 短い）
#    ④0 を 含む ／ ⑤全部 0 ／ ⑥日付が 逆順 ／ ⑦日付が バラバラ（昇順で ない）
#    ⑧同じ 日付が 2つ ／ ⑨1件だけ ／ ⑩値と 日付の 数が 違う
#    ⑪空の マスが 混ざる ／ ⑫字が 混ざる ／ ⑬日付が 1900年より 前（負）
#    ⑭見当（第3引数）が −1 以下 ／ ⑮見当が 大きすぎる
#
#  ★物差しの 決まり★
#    ・2つ目の 窓（0 が 出たら `=(式)=0` と 型で 疑う）
#    ・材料は 紙に 書き出す（使う 側が 手で 写さない）
#    ・書き戻しは LF
#
#  ★答えは 実Excel の 実測だけが 正★（AI に 出させない）
#  ★司さんの 実物には 触りません★＝新しい ブック・保存せず
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-xirr-sakaime.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-xirr-sakaime-2026-09-09.tsv'

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

# ══ ★組（★入れ子の 配列は 平らに なる★ので 列ごとの 表に する）★ ══
#   お金の 列 ／ 日付の 列 ／ 札
$組 = @(
  @{ 札 = '★境目★ 全部 プラス';        v = 'A'; d = 'B'; 値 = @(100, 200, 300);        日 = @(45292, 45383, 45474) },
  @{ 札 = '★境目★ 全部 マイナス';      v = 'C'; d = 'D'; 値 = @(-100, -200, -300);     日 = @(45292, 45383, 45474) },
  @{ 札 = '普通（プラスと マイナス）';   v = 'E'; d = 'F'; 値 = @(-1000, 600, 700);      日 = @(45292, 45383, 45474) },
  @{ 札 = '★一番 短い（2件）★';        v = 'G'; d = 'H'; 値 = @(-1000, 1100);          日 = @(45292, 45657) },
  @{ 札 = '★0 を 含む★';               v = 'I'; d = 'J'; 値 = @(-1000, 0, 1100);       日 = @(45292, 45383, 45474) },
  @{ 札 = '★全部 0★';                  v = 'K'; d = 'L'; 値 = @(0, 0, 0);              日 = @(45292, 45383, 45474) },
  @{ 札 = '★日付が 逆順★';             v = 'M'; d = 'N'; 値 = @(-1000, 600, 700);      日 = @(45474, 45383, 45292) },
  @{ 札 = '★日付が バラバラ★';          v = 'O'; d = 'P'; 値 = @(-1000, 600, 700);      日 = @(45383, 45292, 45474) },
  @{ 札 = '★同じ 日付が 2つ★';          v = 'Q'; d = 'R'; 値 = @(-1000, 600, 700);      日 = @(45292, 45292, 45474) },
  @{ 札 = '★1件だけ★';                 v = 'S'; d = 'T'; 値 = @(-1000);                日 = @(45292) },
  @{ 札 = '★日付が 1900年より 前（負）★'; v = 'U'; d = 'V'; 値 = @(-1000, 600, 700);      日 = @(-10, 45383, 45474) }
)
foreach ($c in $組) {
  if ($c.値.Count -ne $c.日.Count) { Write-Error ('★' + $c.札 + ' の 数が 合わない★'); exit 2 }
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Worksheets.Item(1)
  foreach ($c in $組) {
    for ($i = 0; $i -lt $c.値.Count; $i++) {
      $sh.Range($c.v + ($i + 1)).Value2 = [double]$c.値[$i]
      $sh.Range($c.d + ($i + 1)).Value2 = [double]$c.日[$i]
    }
  }
  # ★数が 違う 組★（値 3つ・日付 2つ）… 別に 敷く
  $sh.Range('X1').Value2 = [double](-1000); $sh.Range('X2').Value2 = [double]600; $sh.Range('X3').Value2 = [double]700
  $sh.Range('Y1').Value2 = [double]45292;   $sh.Range('Y2').Value2 = [double]45383
  # ★空が 混ざる 組★（Z2 を 空の まま）
  $sh.Range('AA1').Value2 = [double](-1000); $sh.Range('AA3').Value2 = [double]1100
  $sh.Range('AB1').Value2 = [double]45292;   $sh.Range('AB2').Value2 = [double]45383; $sh.Range('AB3').Value2 = [double]45474
  # ★字が 混ざる 組★
  $sh.Range('AD1').Value2 = [double](-1000); $sh.Range('AD2').Value2 = 'あ'; $sh.Range('AD3').Value2 = [double]1100
  $sh.Range('AE1').Value2 = [double]45292;   $sh.Range('AE2').Value2 = [double]45383; $sh.Range('AE3').Value2 = [double]45474

  function 押して字に([string]$式) {
    $sh.Range('BA1:BY5').Clear() | Out-Null
    try {
      $sh.Range('BA1').Formula = $式
      $v = $sh.Range('BA1').Value2
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
  $行.Add('# ★XIRR が どこで #NUM! に なるかを 実Excel に 聞いた★（2026-09-09）')
  $行.Add('#')
  $行.Add('# ★なぜ★ 実Excel が #NUM! を 返す 組で うちは ★数を 返して いる★')
  $行.Add('#   =XIRR(全部プラス) … 実Excel #NUM! ／ うち 375686054239746.3')
  $行.Add('#   =XIRR(日付が逆順) … 実Excel #NUM! ／ うち -0.5002981524324954')
  $行.Add('#   ⇒★誤りに ならず 数が 出る＝お客さんは 気づけない★（★お金の 利回り★）')
  $行.Add('#')
  $行.Add('# ★どこから 断るかは 私が 決めない＝実Excel の 答えの 通り★')
  $行.Add('# ★どの Excel で 打ったか★ … 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('#')
  $行.Add('# ★★材料（機械が 読む）★★ … `#材料<タブ>マス<タブ>値` の 形')
  foreach ($c in $組) {
    for ($i = 0; $i -lt $c.値.Count; $i++) {
      $行.Add('#材料' + "`t" + ($c.v + ($i + 1)) + "`t" + ([double]$c.値[$i]).ToString('R'))
      $行.Add('#材料' + "`t" + ($c.d + ($i + 1)) + "`t" + ([double]$c.日[$i]).ToString('R'))
    }
  }
  foreach ($m in @(@('X1', -1000), @('X2', 600), @('X3', 700), @('Y1', 45292), @('Y2', 45383),
                   @('AA1', -1000), @('AA3', 1100), @('AB1', 45292), @('AB2', 45383), @('AB3', 45474),
                   @('AD1', -1000), @('AD3', 1100), @('AE1', 45292), @('AE2', 45383), @('AE3', 45474))) {
    $行.Add('#材料' + "`t" + $m[0] + "`t" + ([double]$m[1]).ToString('R'))
  }
  $行.Add('#材料' + "`t" + 'AD2' + "`t" + '★字「あ」★')
  $行.Add('#材料' + "`t" + 'AA2' + "`t" + '★空★')
  $行.Add('#')
  $行.Add('# 組' + "`t" + '式' + "`t" + '実Excel の 答え' + "`t" + '型' + "`t" + '窓２(=(式)=0)' + "`t" + '何を 見て いるか')

  $本数 = 0
  function 足す([string]$札, [string]$式, [string]$見る) {
    $r = 押して字に $式
    $script:行.Add($札 + "`t" + $式 + "`t" + $r.値 + "`t" + $r.型 + "`t" + $r.窓2 + "`t" + $見る)
    $script:本数++
    Write-Host ('  ' + $札.PadRight(26) + ' ' + $式.PadRight(34) + ' → ' + $r.値)
  }

  foreach ($c in $組) {
    $n = $c.値.Count
    $y = $c.v + '1:' + $c.v + $n
    $x = $c.d + '1:' + $c.d + $n
    足す $c.札 ('=XIRR(' + $y + ',' + $x + ')') '見当なし'
    # ★見当を 変えても 同じ 誤りか★（★誤りが 見当で 変わるなら 条件では ない★）
    foreach ($g in @('0.1', '0', '-0.9', '10')) {
      足す $c.札 ('=XIRR(' + $y + ',' + $x + ',' + $g + ')') ('見当 ' + $g)
    }
  }
  足す '★値 3・日付 2（数が 違う）★' '=XIRR(X1:X3,Y1:Y2)' '数が 違う'
  足す '★空が 混ざる★'               '=XIRR(AA1:AA3,AB1:AB3)' '空の マス'
  足す '★字が 混ざる★'               '=XIRR(AD1:AD3,AE1:AE3)' '字の マス'
  足す '★見当が −1★'                 '=XIRR(E1:E3,F1:F3,-1)' '見当の 境目'
  足す '★見当が −2★'                 '=XIRR(E1:E3,F1:F3,-2)' '見当の 境目'
  足す '★見当が 1e6★'                '=XIRR(E1:E3,F1:F3,1000000)' '見当の 境目'
  # ★★見当の 境目を 細かく 測る（2026-09-09 に 足した）★★
  #   1回目は −1／−2／−0.9 の 3つしか 打たず、私は「−1 以下で #NUM!」と 決めました。
  #   ⇒★実Excel は −0.9 でも #NUM! を 返して いた★＝★私の 決めが 間違い★
  #   ⇒★境目は 私が 決めず 刻んで 打つ★
  foreach ($g in @('-0.99','-0.95','-0.9','-0.8','-0.7','-0.5','-0.3','-0.1','-0.01','0.01','0.5','1','2','5','100','10000')) {
    足す '★見当の 刻み★' ('=XIRR(E1:E3,F1:F3,' + $g + ')') ('見当 ' + $g)
  }

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★★書けた … ' + $本数 + '本★★')
  Write-Host ('★書いた … ' + $出 + '★')

  $bk.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
