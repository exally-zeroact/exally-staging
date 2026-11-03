# toru-linest-kuwashiku.ps1 — ★LINEST を ★丸めが 見える データ★で 測り直す★（2026-09-08）
#
#  ★★なぜ 取り直すか★★
#    1回目（`golden-marume-A-2026-09-08.tsv`）の LINEST は
#    傾きが ★20★ と ★2.13★ の 2組だけ だった。
#    ⇒★どちらも 4桁で ぴったり 表せる 数★
#    ⇒★4桁に 丸めても 答えが 変わらない★
#    ⇒ 押したら 相対差 ★3.55e-16★＝★丸めが 見えない★
#    ⇒★★測ったのに 物差しが 空洞★★（今日の 決まり）
#    ⇒ だから ★傾き・切片に 桁が 出る データ★で 取り直す
#
#  ★★もう1つ 分かった 事★★
#    `=INDEX(LINEST(N1:N5,O1:O5),1,1)` … うちは ★#NAME?★
#    ⇒★エンジンに LINEST が 積まれて いない★
#    ⇒ 動くのは ★JS層が 拾う 裸の `=LINEST(範囲,範囲)` だけ★
#    ⇒★入れ子に すると 動かない＝別の 不具合★（この 紙で 数を 出す）
#
#  ★別の 関数で 裏取り しない★
#    実Excel の 中で LINEST と SLOPE が 違う 答えを 出す（1回目の 実測）
#      =INDEX(LINEST(N1:N5,O1:O5),1,1) … 19.999999999999993
#      =SLOPE(N1:N5,O1:O5)             … 20
#    ⇒★LINEST は LINEST の 数と だけ 比べる★
#
#  ★司さんの 実物には 触りません★＝新しい ブック・保存せず
#
#  使い方: powershell -File docs/measured/toru-linest-kuwashiku.ps1

$ErrorActionPreference = 'Stop'

# ══ ★★2つ目の 窓（2026-09-08 に 足した）★★ ══════════════════
#  ★物差しの 欠陥★ .Value2 は ★0 で ない 値に 0 を 返す★
#    =0.1+0.2-0.3    … .Value2 ★0★ ／ =(式)=0 ★False★ ／ (式)*1e17 5.55
#    =11.1+22.2-33.3 … .Value2 0   ／ =(式)=0 ★True★  ／ (式)*1e17 0
#    ⇒★.Value2 では この 2つが どちらも 0 に 見える★
#    正体 …★最後の 演算が ＋か− の 時だけ 実Excel が ★見せる 時に★ 0 に する★
#  ★もう1つ★ =DEC2BIN(0.5) は ★文字列の "0"★＝数の 0 では ない
#    ⇒ ="0"=0 は FALSE ⇒★見せかけの 0 と 同じ 顔★⇒★型を 見ないと 分けられない★
#  ⇒★見張り tests/monosashi-mado.test.mjs が これを 入れて いない 道具を 赤に する★
function 窓２_型($v) {
  if ($null -eq $v) { return 'Empty' }
  if ($v -is [string]) { return 'String' }
  if ($v -is [bool]) { return 'Boolean' }
  if ($v -is [double] -or $v -is [int] -or $v -is [long]) { return 'Number' }
  return 'Other'
}
function 窓２_本当にゼロか($sh, [string]$式) {
  # ★『0』が 出た 時だけ 呼ぶ★ … =(式)=0 の 真偽を 返す
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

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-linest-2026-09-08.tsv'

$誤りの番号 = @{
  -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!';
  -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!';
  -2146826246 = '#N/A';
  -2146826243 = '#SPILL!';   -2146826242 = '#CONNECT!'; -2146826241 = '#BLOCKED!';
  -2146826240 = '#UNKNOWN!'; -2146826239 = '#FIELD!';   -2146826238 = '#CALC!'; -2146826237 = '#BUSY!'
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Worksheets.Item(1)

  # ══ ★組（★桁が 出る 物を わざと 入れる★）★ ══════════════════
  #   1) きれいに 乗る … 傾き 20・切片 80（★1回目と 同じ＝比べる 為に 残す★）
  #   2) 端数 2桁 …… 傾き 2.13（★1回目と 同じ★）
  #   3) ★7で 割る★ … 傾きが 割り切れない
  #   4) ★3で 割る★ … 同上
  #   5) ★ばらつく★ … 最小二乗が 端数だらけに なる
  #   6) ★小さい 数★ … 桁が 下に 伸びる
  #   7) ★大きい 数★ … 桁が 上に 伸びる
  $組 = @(
    @{ 列y = 'A'; 列x = 'B'; y = @(100, 120, 140, 160, 180);            x = @(1, 2, 3, 4, 5); 札 = 'きれいに乗る（傾き20）' },
    @{ 列y = 'C'; 列x = 'D'; y = @(2.7, 5.1, 6.9, 9.4, 11.2);           x = @(1, 2, 3, 4, 5); 札 = '端数2桁（傾き2.13）' },
    @{ 列y = 'E'; 列x = 'F'; y = @(1, 2, 3, 4, 5);                      x = @(7, 14, 21, 28, 35); 札 = '★7で割る★' },
    @{ 列y = 'G'; 列x = 'H'; y = @(1, 2, 3, 4, 5);                      x = @(3, 6, 9, 12, 15);   札 = '★3で割る★' },
    @{ 列y = 'I'; 列x = 'J'; y = @(1, 5, 2, 8, 3, 9, 4);                x = @(1, 2, 3, 4, 5, 6, 7); 札 = '★ばらつく（7点）★' },
    @{ 列y = 'K'; 列x = 'L'; y = @(0.001, 0.0023, 0.0031, 0.0047, 0.0052); x = @(1, 2, 3, 4, 5); 札 = '★小さい数★' },
    @{ 列y = 'M'; 列x = 'N'; y = @(1234567, 2345678, 3456789, 4567890, 5678901); x = @(1, 2, 3, 4, 5); 札 = '★大きい数★' }
  )
  foreach ($c in $組) {
    for ($i = 0; $i -lt $c.y.Count; $i++) {
      $sh.Range($c.列y + ($i + 1)).Value2 = [double]$c.y[$i]
      $sh.Range($c.列x + ($i + 1)).Value2 = [double]$c.x[$i]
    }
  }

  function 押して字に([string]$式) {
    $sh.Range('T1:Z40').Clear() | Out-Null
    try {
      $sh.Range('T1').Formula = $式
      $v = $sh.Range('T1').Value2
      $t = [string]$sh.Range('T1').Text
      if ($null -eq $v) { return @{ 値 = '(空)'; 型 = 'Empty'; 字 = $t } }
      if (($v -is [int] -or $v -is [long]) -and $誤りの番号.ContainsKey([int]$v)) {
        return @{ 値 = $誤りの番号[[int]$v]; 型 = 'Error'; 字 = $t }
      }
      if ($v -is [double]) { return @{ 値 = $v.ToString('R'); 型 = 'Double'; 字 = $t } }
      if ($v -is [int] -or $v -is [long]) { return @{ 値 = [string]$v; 型 = 'Int32'; 字 = $t } }
      return @{ 値 = [string]$v; 型 = 'String'; 字 = $t }
    } catch { return @{ 値 = '★受け付けない★'; 型 = 'Rejected'; 字 = '' } }
  }

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★LINEST を ★丸めが 見える データ★で 測り直した★（2026-09-08）')
  $行.Add('# ★なぜ 取り直したか★')
  $行.Add('#   1回目は 傾きが 20 と 2.13 の 2組だけ＝★どちらも 4桁で ぴったり 表せる★')
  $行.Add('#   ⇒ 4桁に 丸めても 答えが 変わらず、押したら 相対差 ★3.55e-16★')
  $行.Add('#   ⇒★★測ったのに 物差しが 空洞★★＝★桁が 出る データで 取り直した★')
  $行.Add('# ★組（★桁が 出る 物を わざと 入れた★）★')
  $行.Add('#   A/B きれいに乗る（傾き20・1回目と同じ）／ C/D 端数2桁（傾き2.13・1回目と同じ）')
  $行.Add('#   E/F ★7で割る★／ G/H ★3で割る★／ I/J ★ばらつく（7点）★')
  $行.Add('#   K/L ★小さい数★／ M/N ★大きい数★')
  $行.Add('# ★別の 関数で 裏取り しない★＝LINEST は LINEST の 数と だけ 比べる')
  $行.Add('#   （実Excel の 中で LINEST 19.999999999999993 ／ SLOPE 20＝違う 答え）')
  $行.Add('# ★うちは エンジンに LINEST が 積まれて いない★')
  $行.Add('#   =INDEX(LINEST(…),1,1) … うち ★#NAME?★＝★入れ子に すると 動かない★')
  $行.Add('#   動くのは JS層が 拾う ★裸の =LINEST(範囲,範囲) だけ★')
  $行.Add('#')
  $行.Add('# ★★押した 結果（2026-09-08）★★')
  $行.Add('#   ★裸の =LINEST(範囲,範囲)（JS層・4桁の 丸めが 在る 道）★')
  $行.Add('#     =LINEST(K1:K5,L1:L5) 実Excel 0.0010799999999999998 ／ うち ★0.0011★ 相対 ★1.85e-2★ ←一番 悪い')
  $行.Add('#     =LINEST(E1:E5,F1:F5) 実Excel 0.14285714285714282   ／ うち 0.1429   相対 3.00e-4')
  $行.Add('#     =LINEST(G1:G5,H1:H5) 実Excel 0.33333333333333337   ／ うち 0.3333   相対 1.00e-4')
  $行.Add('#     =LINEST(I1:I7,J1:J7) 実Excel 0.6428571428571429    ／ うち 0.6429   相対 6.67e-5')
  $行.Add('#     =LINEST(A1:A5,B1:B5) 実Excel 19.999999999999993    ／ うち 20       相対 3.55e-16 ←★1回目の 組★')
  $行.Add('#     =LINEST(C1:C5,D1:D5) 実Excel 2.1299999999999994    ／ うち 2.13     相対 2.08e-16 ←★1回目の 組★')
  $行.Add('#     =LINEST(M1:M5,N1:N5) 実Excel 1111088               ／ うち 1111088  相対 0')
  $行.Add('#   ⇒★一番 悪くて 相対 1.85e-2＝約 1.9% ずれる（小さい 数ほど ひどい）★')
  $行.Add('#   ⇒★XIRR の 2.4e-5 より ★770倍 悪い★')
  $行.Add('#   ⇒★1回目の 2組は 3.55e-16／2.08e-16＝★丸めが 見えなかった★（見本が 甘かった 証拠）')
  $行.Add('#')
  $行.Add('#   ★★入れ子の =INDEX(LINEST(…),1,1) … うちは ★7組 14本 全部 #NAME?★★★')
  $行.Add('#     ⇒★エンジンに LINEST が 積まれて いない★')
  $行.Add('#     ⇒ しかも うちは ★傾き 1つしか 返して いない★（LINEST は 本来 表を 返す）')
  $行.Add('#     ⇒★★「在るように 見えて 実は 無い」＝別件（2件目）★★')
  $行.Add('#       ★丸めの 件とは 原因も 直し方も 違う★')
  $行.Add('#         丸め … 出す 時の 事 ／ 入れ子 … 積んで いない 事')
  $行.Add('#         丸めを 外す ／ エンジンに 積む')
  $行.Add('#')
  $行.Add('# ★★一致の 幅 … ★仮に 相対 1e-12★★★')
  $行.Add('#   ★★これは ★仮★です＝丸めを 外した 後に 押し直して 決め直します★★')
  $行.Add('#   理由 …★今は 4桁の 丸めが 効いて いて ★本当の 差が 見えない★')
  $行.Add('#          （丸めの せいの 1.85e-2 が 出るだけで、うちの 計算の 良し悪しが 分からない）')
  $行.Add('#   ⇒★決め直したら ★前後 両方★ 取り直す★')
  $行.Add('#   ⇒★★仮の まま PR に 載せない★★（PR の 幅は 測って 決めた 物だけ）')
  $行.Add('#')
  # ══ ★★材料を ★機械が 読める 形★で 書き出す（2026-09-08 に 足した）★★ ══
  #   ★理由★ 押し比べの 道具が 材料を ★人（私）が 書き写して★ 使うと 写し間違える。
  #           2026-09-08 に ★2回 踏んだ★（保留27行の 土台／押し比べの XIRR の お金）
  #   ⇒★紙に 材料を 載せ、使う 側は ★紙から 読む★＝もう 写さない★
  $行.Add('# ★★材料（機械が 読む）★★ … `#材料<タブ>マス<タブ>値` の 形')
  foreach ($c in $組) {
    for ($i = 0; $i -lt $c.y.Count; $i++) {
      $行.Add('#材料' + "`t" + ($c.列y + ($i + 1)) + "`t" + ([double]$c.y[$i]).ToString('R'))
      $行.Add('#材料' + "`t" + ($c.列x + ($i + 1)) + "`t" + ([double]$c.x[$i]).ToString('R'))
    }
  }
  $行.Add('#')
  $行.Add('# 関数' + "`t" + '式' + "`t" + '実Excel の 答え' + "`t" + '型' + "`t" + '実Excel が 出す 字' + "`t" + '土台')

  $本数 = 0
  foreach ($c in $組) {
    $n = $c.y.Count
    $y = $c.列y + '1:' + $c.列y + $n
    $x = $c.列x + '1:' + $c.列x + $n
    foreach ($f in @(
        @{ 式 = '=INDEX(LINEST(' + $y + ',' + $x + '),1,1)'; 札 = '傾き' },
        @{ 式 = '=INDEX(LINEST(' + $y + ',' + $x + '),1,2)'; 札 = '切片' },
        @{ 式 = '=LINEST(' + $y + ',' + $x + ')';            札 = '裸（★JS層が 拾う 形★）' })) {
      $r = 押して字に $f.式
      $行.Add('LINEST' + "`t" + $f.式 + "`t" + $r.値 + "`t" + $r.型 + "`t" + $r.字 + "`t" + $c.札 + '／' + $f.札)
      $本数++
      Write-Host ('  ' + $f.式.PadRight(44) + ' → ' + $r.値.PadRight(24) + ' ' + $c.札 + '／' + $f.札)
    }
  }

  [System.IO.File]::WriteAllLines($出, $行, (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★★書けた … ' + $本数 + '本★★')
  Write-Host ('★書いた … ' + $出 + '★')

  $bk.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
