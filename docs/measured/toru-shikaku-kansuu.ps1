# toru-shikaku-kansuu.ps1 — ★四角（A1:A3）と 関数の 扱いを 実Excel に 聞く★（2026-09-13）
#
#  ★★なぜ 要るか★★
#    土台の ④後半＝★関数と 四角を 計算する★ を 作ります。
#    ここは ★当て推量が 必ず 外れる★ 所です。いちばん 大きいのは これ：
#      `=SUM(A1:A5)` … ★四角の 中の 字と 真偽は 無視される★
#      `=SUM(A1,"2",TRUE)` … ★直に 書いた 字と 真偽は 数に なる★
#    ★同じ SUM なのに 中身の 扱いが 違う★＝知らずに 作ると 全部 外します。
#
#  ★物差しの 決まり★
#    ・★見たい物は 1つだけ★＝ここは「四角と 関数の 型の 扱い」だけ
#    ・★3つ とも 見る★ ①出る字(.Text) ②型 ③★誤りか★(=ISERROR)
#    ・★溢れたか★も 見る（四角を そのまま 打つと 下に 広がる）
#    ・★司さんの 実物には 触りません★＝新しい ブックを 作って 保存せず 閉じる
#    ・書き戻しは LF
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-shikaku-kansuu.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-shikaku-kansuu-2026-09-13.tsv'

# ★材料★（A列）
#   A1 = 1（数）／ A2 = "2"（★字★）／ A3 = TRUE（真偽）／ A4 = ★空★／ A5 = 2（数）
#   A6 = =1/0（★誤り★）／ B1:B3 = ★全部 空★
$試 = @(
  @{ 式='=SUM(A1:A5)';          何='★四角★の 中の 字と 真偽は 無視されるか' },
  @{ 式='=SUM(A1,"2",TRUE)';    何='★直に 書いた★ 字と 真偽は 数に なるか' },
  @{ 式='=SUM(A1:A6)';          何='四角の 中の 誤りは 伝わるか' },
  @{ 式='=SUM(A2)';             何='1マスだけ 指した 字は どうなるか' },
  @{ 式='=SUM(A1:A3,A5)';       何='引数が 2つ（四角＋1マス）' },
  @{ 式='=SUM(B1:B3)';          何='全部 空の 四角' },
  @{ 式='=COUNT(A1:A5)';        何='四角で 数だけ 数えるか' },
  @{ 式='=COUNT("1",TRUE)';     何='★直に 書いた★ 字と 真偽を 数えるか' },
  @{ 式='=COUNT(B1:B3)';        何='全部 空を 数える' },
  @{ 式='=COUNTA(A1:A5)';       何='空 以外を 数えるか' },
  @{ 式='=AVERAGE(A1:A5)';      何='分母は 何か（字と 真偽を 入れるか）' },
  @{ 式='=AVERAGE(A1,A3)';      何='1マス指しの 真偽は 分母に 入るか' },
  @{ 式='=AVERAGE(B1:B3)';      何='全部 空の 平均' },
  @{ 式='=MAX(A1:A5)';          何='四角の 中の 字と 真偽を 見るか' },
  @{ 式='=MAX("3",1)';          何='★直に 書いた★ 字は 通るか' },
  @{ 式='=MAX(B1:B3)';          何='全部 空の 一番大きい物' },
  @{ 式='=MIN(A1:A5)';          何='四角の 一番小さい物' },
  @{ 式='=MIN(A2,1)';           何='1マス指しの 字は 無視か 誤りか' },
  @{ 式='=PRODUCT(A1:A5)';      何='掛け算は 空を どう 見るか' },
  @{ 式='=A1:A3';               何='★四角を そのまま 打つ★（溢れるか）' },
  @{ 式='=A1:A3*2';             何='四角に 掛ける（溢れるか）' },
  @{ 式='=SUMPRODUCT(A1:A5,A1:A5)'; 何='四角どうしを 掛けて 足す' },
  @{ 式='=MIN(B1:B3)';          何='全部 空の 一番小さい物（MAXが0なら MINも0か）' },
  @{ 式='=PRODUCT(B1:B3)';      何='全部 空の 掛け算' },
  @{ 式='=COUNTA(B1:B3)';       何='全部 空を 数える' },
  @{ 式='=MAX(A1:A6)';          何='四角に 誤りが 在る時の 一番大きい物' },
  @{ 式='=COUNT(A1:A6)';        何='誤りを 数えるか' },
  @{ 式='=COUNTA(A1:A6)';       何='誤りを 空以外として 数えるか' },
  @{ 式='=AVERAGE(A1:A6)';      何='誤り入りの 平均' },
  @{ 式='=SUM("x")';            何='★直に 書いた★ 数に ならない 字' },
  @{ 式='=SUM(A1:A5,"x")';      何='四角 ＋ 数に ならない 字' },
  @{ 式='=AVERAGE(A2)';         何='字だけを 指した 平均（分母0か）' },
  @{ 式='=MAX(A2)';             何='字だけを 指した 一番大きい物' },
  @{ 式='=SUM(0.1,0.2,-0.3)';   何='★SUMの 中でも 消え残りを 0に するか★' },
  @{ 式='=SUM(C1:C3)';          何='★四角でも★ 消え残りを 0に するか' },
  @{ 式='=MAX(C5:C6)';          何='★負の数 だけの 四角（0が 混ざらないか）★' },
  @{ 式='=MIN(C5:C6)';          何='負の数 だけの 一番小さい物' },
  @{ 式='=MAX(C5)';             何='負の数 1マス' },
  @{ 式='=PRODUCT(2,3)';        何='普通の 掛け算' },
  @{ 式='=AVERAGE(1,2)';        何='普通の 平均' }
)

function 窓２_型($v) {
  if ($null -eq $v) { return 'Empty' }
  if ($v -is [string]) { return 'String' }
  if ($v -is [bool]) { return 'Boolean' }
  if ($v -is [double] -or $v -is [int] -or $v -is [long]) { return 'Number' }
  return 'Other'
}

# ★誤りかを 別の 窓で 聞く★
#   ★.Value2 は 誤りを ★大きな 負の数★ で 返す★（#NUM! は -2146826252）
#   ⇒ 数に 見えて しまう ので ★=ISERROR() を 打って 聞く★
# ★本当に 0 かを 別の 窓で 聞く★
#   ★出る字が「0」でも 5.55E-17 の ことが 在る★（消え残り）
#   ⇒ ★=(E1)=0 を 打って 聞く★（字だけでは 見分けが 付かない）
function 窓２_本当にゼロか($sh, [string]$マス) {
  try {
    $sh.Range('BY1').Clear() | Out-Null
    $sh.Range('BY1').Formula = ('=(' + $マス + ')=0')
    $z = $sh.Range('BY1').Value2
    $sh.Range('BY1').Clear() | Out-Null
    if ($z -is [bool]) { return $(if ($z) { 'TRUE' } else { 'FALSE' }) }
    return '★判じられない★'
  } catch { return '★判じられない★' }
}
function 窓２_誤りか($sh, [string]$マス) {
  try {
    $sh.Range('BZ1').Clear() | Out-Null
    $sh.Range('BZ1').Formula = ('=ISERROR(' + $マス + ')')
    $z = $sh.Range('BZ1').Value2
    $sh.Range('BZ1').Clear() | Out-Null
    if ($z -is [bool]) { return $(if ($z) { 'TRUE' } else { 'FALSE' }) }
    return '★判じられない★'
  } catch { return '★判じられない★' }
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Worksheets.Item(1)

  $sh.Range('A1').Value2 = 1
  $sh.Range('A2').NumberFormat = '@'
  $sh.Range('A2').Value2 = '2'
  $sh.Range('A3').Value2 = $true
  # A4 は ★空のまま★
  $sh.Range('A5').Value2 = 2
  $sh.Range('A6').Formula = '=1/0'
  $sh.Range('C1').Value2 = 0.1
  $sh.Range('C2').Value2 = 0.2
  $sh.Range('C3').Value2 = -0.3
  $sh.Range('C5').Value2 = -1
  $sh.Range('C6').Value2 = -2

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★四角（A1:A3）と 関数の 扱いを 実Excel に 聞いた★（2026-09-13）')
  $行.Add('#')
  $行.Add('# ★一番 大きい 分かれ目★')
  $行.Add('#   ★四角★の 中の 字と 真偽 … 無視される')
  $行.Add('#   ★直に 書いた★ 字と 真偽 … 数に なる')
  $行.Add('#   ⇒ 同じ SUM でも 中身の 扱いが 違う')
  $行.Add('#')
  $行.Add('# ★材料★ A1=1(数) A2="2"(★字★) A3=TRUE(真偽) A4=★空★ A5=2(数) A6==1/0(誤り) B1:B3=全部空')
  $行.Add('#')
  $行.Add('# ★どの Excel か★ … 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('#')
  $行.Add('# ★材料 その2★ C1=0.1 C2=0.2 C3=-0.3（消え残り）／ C5=-1 C6=-2（負だけ）')
  $行.Add('#')
  $行.Add('# 式' + "`t" + '出る字' + "`t" + '型' + "`t" + '誤りか' + "`t" + '本当にゼロか' + "`t" + '溢れた(E2)' + "`t" + '何を 見て いるか')

  for ($i = 0; $i -lt $試.Count; $i++) {
    $sh.Range('E1:G30').Clear() | Out-Null
    $sh.Range('E1').Formula2 = $試[$i].式
    $字 = [string]$sh.Range('E1').Text
    $v = $sh.Range('E1').Value2
    $型 = 窓２_型 $v
    $誤 = 窓２_誤りか $sh 'E1'
    $下 = $sh.Range('E2').Value2
    $溢 = if ($null -eq $下) { '—' } else { '★溢れた★ E2=' + [string]$sh.Range('E2').Text }
    $ゼ = if ($字 -match '^-?0(\.0+)?$') { 窓２_本当にゼロか $sh 'E1' } else { '—' }
    $行.Add($試[$i].式 + "`t" + $字 + "`t" + $型 + "`t" + $誤 + "`t" + $ゼ + "`t" + $溢 + "`t" + $試[$i].何)
    Write-Host ('  ' + $試[$i].式.PadRight(28) + ' 出る字=' + $字.PadRight(12) + ' 誤=' + $誤.PadRight(6) + ' ゼロ=' + $ゼ.PadRight(6) + ' ' + $溢)
  }

  $sh.Range('E1:G30').Clear() | Out-Null
  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★書いた … ' + $出 + '★')
  $bk.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
