# toru-zero-torinaosu.ps1 — ★取った 紙の「0」を ★2つ目の 窓★で 取り直す★（2026-09-08）
#
#  ★★物差しの 欠陥★★
#    `.Value2` は ★0 で ない 値に 0 を 返す★ 事が 在る。
#    正体 …★最後の 演算が ＋ か − の 時だけ、実Excel が ★見せる 時に★ 0 に する★
#      =0.1+0.2-0.3    … .Value2 ★0★ ／ =(式)=0 ★False★ ／ (式)*1e17 5.55
#      =11.1+22.2-33.3 … .Value2 0   ／ =(式)=0 ★True★  ／ (式)*1e17 0
#    ⇒★.Value2 では この 2つが どちらも 0 に 見える★
#    ⇒★私は ①の 窓だけで 今日 10,775行 取って きた★
#
#  ★★数えた 結果（`zero-no-gyou-kazoeru.mjs`）★★
#    紙 28枚 ／ 全 10,775行
#      ①答えが 0 の 行 …………………… ★697行★
#      ②足し引きで 終わる 式 …………… ★34行★
#      ③両方に 当たる（一番 危ない）…… ★22行★（★全部 今日の hikizan-zero の 紙★）
#
#  ★★この 紙で する 事★★
#    ★697行の「0」を ★2つの 窓を 並べて★ 取り直す★
#      窓① `.Value2`
#      窓② `=(式)=0` の 真偽   ←★本当に 0 か★
#      窓③ `(式)*1e17`         ←★残りの 大きさ★
#    ⇒★後の 人が「どちらの 窓の 話か」を 引き直せる 形に する★（監査の 注文B）
#
#  ★司さんの 実物には 触りません★＝新しい ブック・保存せず
#
#  使い方: powershell -File docs/measured/toru-zero-torinaosu.ps1

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
$出 = Join-Path $ここ 'golden-zero-torinaoshi-2026-09-08.tsv'

$誤りの番号 = @{
  -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!';
  -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!';
  -2146826246 = '#N/A';   -2146826243 = '#SPILL!'; -2146826238 = '#CALC!'; -2146826237 = '#BUSY!'
}

# ★「0」の 行を 紙から 集める★
$紙ら = @()
foreach ($d in @($ここ, (Join-Path $ここ 'kansuu46'))) {
  if (-not (Test-Path $d)) { continue }
  # ★★自分が 出す 紙を 読み込まない★★（2026-09-08 に 踏んだ）
  #   出す 先も golden で 始まる 紙なので、2回目に 走らせると ★自分の 出力を 読んで★
  #   697行 → 1,367行 に 膨れた。★道具が 自分を 数えて いた★
  $紙ら += Get-ChildItem -Path $d -Filter 'golden-*.tsv' -File |
    Where-Object { $_.Name -ne 'golden-zero-torinaoshi-2026-09-08.tsv' }
}
$対象 = @()
foreach ($f in $紙ら) {
  foreach ($l in (Get-Content $f.FullName -Encoding UTF8)) {
    if ($l.StartsWith('#')) { continue }
    $p = $l -split "`t"
    if ($p.Count -lt 3) { continue }
    $答 = $p[2].Trim()
    if ($答 -eq '0' -or $答 -eq '-0' -or $答 -eq '0.0') {
      $対象 += @{ 紙 = $f.Name; 式 = $p[1] }
    }
  }
}
Write-Host ('★「0」の 行 … ' + $対象.Count + '行を 取り直します★')

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Worksheets.Item(1)
  # ★土台は 紙ごとに 違うので ★一番 広い 形★を 置く★
  #   （★これで 合わない 行は「土台が 違う」として 印を 付ける★）
  for ($r = 1; $r -le 9; $r++) {
    for ($c = 1; $c -le 9; $c++) { $sh.Cells.Item($r, $c).Value2 = [double]($r * 10 + $c) }
  }
  $sh.Range('A1').Value2 = [double]1; $sh.Range('A2').Value2 = [double]2
  $sh.Range('A3').Value2 = [double]3; $sh.Range('A4').Value2 = [double]4; $sh.Range('A5').Value2 = [double]5
  $sh.Range('B1').Value2 = [double]2; $sh.Range('B2').Value2 = [double]4
  $sh.Range('B3').Value2 = [double]6; $sh.Range('B4').Value2 = [double]8; $sh.Range('B5').Value2 = [double]10

  function 生([string]$式) {
    $sh.Range('T1:Z10').Clear() | Out-Null
    try { $sh.Range('T1').Formula = $式; return $sh.Range('T1').Value2 } catch { return '★受け付けない★' }
  }
  function 字に($v) {
    if ($null -eq $v) { return '(空)' }
    if (($v -is [int] -or $v -is [long]) -and $誤りの番号.ContainsKey([int]$v)) { return $誤りの番号[[int]$v] }
    if ($v -is [bool]) { return $(if ($v) { 'TRUE' } else { 'FALSE' }) }
    if ($v -is [double]) { return $v.ToString('R') }
    return [string]$v
  }

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★取った 紙の「0」を ★2つ目の 窓★で 取り直した★（2026-09-08）')
  $行.Add('#')
  $行.Add('# ★★物差しの 欠陥★★')
  $行.Add('#   `.Value2` は ★0 で ない 値に 0 を 返す★')
  $行.Add('#   正体 …★最後の 演算が ＋か− の 時だけ、実Excel が ★見せる 時に★ 0 に する★')
  $行.Add('#     =0.1+0.2-0.3    … .Value2 ★0★ ／ =(式)=0 ★False★ ／ (式)*1e17 5.55')
  $行.Add('#     =11.1+22.2-33.3 … .Value2 0   ／ =(式)=0 ★True★  ／ (式)*1e17 0')
  $行.Add('#   ⇒★.Value2 では この 2つが どちらも 0 に 見える★')
  $行.Add('#')
  $行.Add('# ★数えた 結果★ 紙 28枚／全 10,775行')
  $行.Add('#   答えが 0 … ★697行★ ／ 足し引きで 終わる 式 … ★34行★ ／ 両方 … ★22行★')
  $行.Add('#')
  $行.Add('# ★★3つの 窓を 同じ 行に 並べた（★後の 人が どちらの 窓か 引き直せる ように★）★★')
  $行.Add('#   窓① .Value2 ／ 窓② =(式)=0 の 真偽（★本当に 0 か★）／ 窓③ (式)*1e17（★残りの 大きさ★）')
  $行.Add('#')
  $行.Add('# ★★土台は 一番 広い 形を 1つ 置いた★★')
  $行.Add('#   A1:I9 = 行*10+列 ／ A1:A5=1..5 ／ B1:B5=2,4,6,8,10 で 上書き')
  $行.Add('#   ⇒★元の 紙は 土台が 違う 物が 混ざる★＝★窓①が 元の 紙と 違う 行には 印を 付けた★')
  $行.Add('#   ⇒★その 行は「土台が 違う」だけかも しれない＝★判じを 保留★★')
  $行.Add('# 元の紙' + "`t" + '式' + "`t" + '窓①.Value2' + "`t" + '★型★' + "`t" + '窓②=0の真偽' + "`t" + '窓③*1e17' + "`t" + '判じ')

  $本数 = 0; $本当にゼロ = 0; $見せかけ = 0; $土台違い = 0; $字のゼロ = 0
  foreach ($x in $対象) {
    $中 = $x.式 -replace '^=\s*', ''
    $生v = 生 $x.式
    $v1 = 字に $生v
    $v2 = 字に (生 ('=(' + $中 + ')=0'))
    $v3 = 字に (生 ('=(' + $中 + ')*1e17'))
    # ★★型も 見る★★（=DEC2BIN(0.5) は ★文字列の "0"★＝数の 0 では ない）
    #   ⇒ 字の "0" は ="0"=0 が FALSE に なる＝★見せかけの 0 と 同じ 顔★
    #   ⇒★型を 見ないと 分けられない★（1回目は 分けられて いなかった）
    $型 = if ($null -eq $生v) { 'Empty' } elseif ($生v -is [string]) { 'String' } elseif ($生v -is [bool]) { 'Boolean' } elseif ($生v -is [double] -or $生v -is [int] -or $生v -is [long]) { 'Number' } else { 'Other' }
    $判 = ''
    if ($v1 -ne '0') { $判 = '★土台が 違う（保留）★'; $土台違い++ }
    elseif ($型 -eq 'String') { $判 = '★字の "0"（数では ない）★'; $字のゼロ++ }
    elseif ($v2 -eq 'TRUE') { $判 = '本当に 0'; $本当にゼロ++ }
    elseif ($v2 -eq 'FALSE') { $判 = '★★見せかけの 0★★'; $見せかけ++ }
    else { $判 = '★判じられない（' + $v2 + '）★' }
    $行.Add($x.紙 + "`t" + $x.式 + "`t" + $v1 + "`t" + $型 + "`t" + $v2 + "`t" + $v3 + "`t" + $判)
    $本数++
    if ($判 -like '*見せかけ*') { Write-Host ('  ★見せかけの 0★ ' + $x.紙 + '  ' + $x.式 + '  → 残り ' + $v3 + 'e-17') }
  }

  [System.IO.File]::WriteAllLines($出, $行, (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★★取り直した … ' + $本数 + '行★★')
  Write-Host ('  本当に 0 …………… ' + $本当にゼロ + '行')
Write-Host ('  ★字の "0"（数では ない）… ' + $字のゼロ + '行★')
  Write-Host ('  ★★見せかけの 0 … ' + $見せかけ + '行★★')
  Write-Host ('  ★土台が 違う（保留）… ' + $土台違い + '行★')
  Write-Host ('★書いた … ' + $出 + '★')

  $bk.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
