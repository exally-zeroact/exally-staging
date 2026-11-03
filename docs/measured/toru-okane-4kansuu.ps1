# toru-okane-4kansuu.ps1 — ★IRR / MIRR / NPV / XNPV の 境目を 実Excel に 聞く★（2026-09-09）
#
#  ★★なぜ★★
#    XIRR が ★実Excel の #NUM! を 数で 返して いました★
#      =XIRR(全部プラス) … 実Excel #NUM! ／ うち ★375686054239746.3★
#    この 4つも ★同じ 家族（お金の 収支の 列を 読む 関数）★です。
#    ⇒★同じ 作りの 見込みが 在る★
#    ⇒★★見込みで 直さない＝先に 測る★★
#      （LOGEST/TREND/GROWTH の 時は ★見込みが 外れました★＝測って 助かった）
#
#  ★★境目は 刻んで 詰める★★（2026-09-09 の 学び）
#    XIRR の 見当を ★3点（−1／−2／−0.9）で 決めて「−1 以下」と 書きました★
#    ⇒ 刻んだら ★−0.01 でも #NUM!★＝★境目は 0★でした
#    ⇒★★3点で 決めた 線は 線では ない★★
#
#  ★物差しの 決まり★
#    ・2つ目の 窓（0 が 出たら `=(式)=0` と 型で 疑う）
#    ・材料は 紙に 書き出す（使う 側が 手で 写さない）
#    ・書き戻しは LF
#
#  ★答えは 実Excel の 実測だけが 正★
#  ★司さんの 実物には 触りません★＝新しい ブック・保存せず
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-okane-4kansuu.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-okane-4kansuu-2026-09-09.tsv'

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
#   A … 普通（−と＋が 在る）／ C … 全部 プラス ／ E … 全部 マイナス
#   G … 全部 0 ／ I … 1件だけ ／ K … 空が 混ざる ／ M … 字が 混ざる
#   O/P … XNPV の 為（お金と 日付）／ R/S … 日付が 逆順
$列 = [ordered]@{
  'A' = @(-1000, 600, 700)
  'C' = @(100, 200, 300)
  'E' = @(-100, -200, -300)
  'G' = @(0, 0, 0)
  'I' = @(-1000)
  'K' = @(-1000, $null, 1100)
  'O' = @(-1000, 600, 700)          # XNPV の お金
  'P' = @(45292, 45383, 45474)      # XNPV の 日付（昇順）
  'R' = @(-1000, 600, 700)
  'S' = @(45474, 45383, 45292)      # ★逆順★
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Worksheets.Item(1)
  foreach ($k in $列.Keys) {
    for ($i = 0; $i -lt $列[$k].Count; $i++) {
      if ($null -eq $列[$k][$i]) { continue }     # ★空の まま★
      $sh.Range($k + ($i + 1)).Value2 = [double]$列[$k][$i]
    }
  }
  # ★字が 混ざる 組★
  $sh.Range('M1').Value2 = [double](-1000); $sh.Range('M2').Value2 = 'あ'; $sh.Range('M3').Value2 = [double]1100

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
  $行.Add('# ★IRR / MIRR / NPV / XNPV の 境目を 実Excel に 聞いた★（2026-09-09）')
  $行.Add('#')
  $行.Add('# ★なぜ★ XIRR が ★実Excel の #NUM! を 数で 返して いた★（375686054239746.3）')
  $行.Add('#   この 4つも ★同じ 家族（お金の 収支の 列を 読む）★＝★同じ 作りの 見込み★')
  $行.Add('#   ⇒★見込みで 直さない＝先に 測る★')
  $行.Add('#')
  $行.Add('# ★境目は 刻んで 詰める★（XIRR で 3点で 決めて 間違えた）')
  $行.Add('# ★どの Excel で 打ったか★ … 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('#')
  $行.Add('# ★★材料（機械が 読む）★★ … `#材料<タブ>マス<タブ>値` の 形')
  foreach ($k in $列.Keys) {
    for ($i = 0; $i -lt $列[$k].Count; $i++) {
      if ($null -eq $列[$k][$i]) { $行.Add('#材料' + "`t" + ($k + ($i + 1)) + "`t" + '★空★'); continue }
      $行.Add('#材料' + "`t" + ($k + ($i + 1)) + "`t" + ([double]$列[$k][$i]).ToString('R'))
    }
  }
  $行.Add('#材料' + "`t" + 'M1' + "`t" + '-1000')
  $行.Add('#材料' + "`t" + 'M2' + "`t" + '★字「あ」★')
  $行.Add('#材料' + "`t" + 'M3' + "`t" + '1100')
  $行.Add('#')
  $行.Add('# 関数' + "`t" + '式' + "`t" + '実Excel の 答え' + "`t" + '型' + "`t" + '窓２(=(式)=0)' + "`t" + '何を 見て いるか')

  $本数 = 0
  function 足す([string]$名, [string]$式, [string]$見る) {
    $r = 押して字に $式
    $script:行.Add($名 + "`t" + $式 + "`t" + $r.値 + "`t" + $r.型 + "`t" + $r.窓2 + "`t" + $見る)
    $script:本数++
    Write-Host ('  ' + $名.PadRight(6) + ' ' + $式.PadRight(36) + ' → ' + $r.値)
  }

  # ── ★IRR★ ──（★符号が 両方 要るか／1件だけ／0／空／字★）
  足す 'IRR' '=IRR(A1:A3)'        '普通（−と＋）'
  足す 'IRR' '=IRR(C1:C3)'        '★全部 プラス★'
  足す 'IRR' '=IRR(E1:E3)'        '★全部 マイナス★'
  足す 'IRR' '=IRR(G1:G3)'        '★全部 0★'
  足す 'IRR' '=IRR(I1:I1)'        '★1件だけ★'
  足す 'IRR' '=IRR(K1:K3)'        '★空が 混ざる★'
  足す 'IRR' '=IRR(M1:M3)'        '★字が 混ざる★'
  # ★見当の 境目を 刻む★（XIRR で 間違えた 所）
  foreach ($g in @('-0.99', '-0.9', '-0.5', '-0.1', '-0.01', '0', '0.01', '0.1', '1', '10', '100')) {
    足す 'IRR' ('=IRR(A1:A3,' + $g + ')') ('見当 ' + $g)
  }

  # ── ★MIRR★ ──（★2つの 利率／符号／0★）
  足す 'MIRR' '=MIRR(A1:A3,0.1,0.12)'   '普通'
  足す 'MIRR' '=MIRR(C1:C3,0.1,0.12)'   '★全部 プラス★'
  足す 'MIRR' '=MIRR(E1:E3,0.1,0.12)'   '★全部 マイナス★'
  足す 'MIRR' '=MIRR(G1:G3,0.1,0.12)'   '★全部 0★'
  足す 'MIRR' '=MIRR(I1:I1,0.1,0.12)'   '★1件だけ★'
  足す 'MIRR' '=MIRR(K1:K3,0.1,0.12)'   '★空が 混ざる★'
  足す 'MIRR' '=MIRR(M1:M3,0.1,0.12)'   '★字が 混ざる★'
  foreach ($r2 in @('-1', '-0.5', '-0.01', '0', '0.01')) {
    足す 'MIRR' ('=MIRR(A1:A3,' + $r2 + ',0.12)') ('借りる 利率 ' + $r2)
    足す 'MIRR' ('=MIRR(A1:A3,0.1,' + $r2 + ')') ('回す 利率 ' + $r2)
  }

  # ── ★NPV★ ──（★利率の 境目★）
  足す 'NPV' '=NPV(0.1,A1:A3)'      '普通'
  足す 'NPV' '=NPV(0.1,C1:C3)'      '★全部 プラス（断らない はず）★'
  足す 'NPV' '=NPV(0.1,G1:G3)'      '★全部 0★'
  足す 'NPV' '=NPV(0.1,K1:K3)'      '★空が 混ざる★'
  足す 'NPV' '=NPV(0.1,M1:M3)'      '★字が 混ざる★'
  foreach ($r3 in @('-2', '-1.5', '-1.01', '-1', '-0.99', '-0.5', '0', '0.5', '10')) {
    足す 'NPV' ('=NPV(' + $r3 + ',A1:A3)') ('利率 ' + $r3)
  }

  # ── ★XNPV★ ──（★日付の 順／利率の 境目★）
  足す 'XNPV' '=XNPV(0.1,O1:O3,P1:P3)'  '普通（日付 昇順）'
  足す 'XNPV' '=XNPV(0.1,R1:R3,S1:S3)'  '★日付が 逆順★'
  足す 'XNPV' '=XNPV(0.1,C1:C3,P1:P3)'  '★全部 プラス★'
  足す 'XNPV' '=XNPV(0.1,O1:O3,P1:P2)'  '★範囲の 大きさが 違う★'
  足す 'XNPV' '=XNPV(0.1,K1:K3,P1:P3)'  '★空が 混ざる★'
  足す 'XNPV' '=XNPV(0.1,M1:M3,P1:P3)'  '★字が 混ざる★'
  # ★★境目を 刻んで 詰める（XIRR で 3点で 決めて 間違えた）★★
  foreach ($r4 in @('-2', '-1.01', '-1', '-0.99', '-0.5', '-0.1', '-0.01', '-0.001', '0',
                    '0.0001', '0.001', '0.01', '0.1', '10')) {
    足す 'XNPV' ('=XNPV(' + $r4 + ',O1:O3,P1:P3)') ('利率 ' + $r4)
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
