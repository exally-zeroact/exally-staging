# toru-afure.ps1 — ★1つの 式が 何マスにも 広がる（溢れ）を 実Excel に 聞く★（2026-09-13）
#
#  ★★なぜ 要るか★★
#    借り物(HyperFormula・GPLv3)を ★自作で 外す★（司さん 2026-09-13）。
#    土台の ⑤枚目＝★溢れ★。`=A1:A3` を 1マスに 打つと ★下に 広がる★（2026-09-13 実測済み）。
#    ⇒★どこまで 広がるか／邪魔が 在ると どうなるか／@ は 何をするか★を 1つずつ 聞きます。
#
#  ★物差しの 決まり★
#    ・★広がった 先を 全部 読む★（E1 だけでなく E1:G5 を なめる）
#    ・★誰が 書いた 物かも 見る★（.HasArray / .CurrentArray で 溢れの 印を 取る）
#    ・★司さんの 実物には 触りません★＝新しい ブックを 作って 保存せず 閉じる
#    ・1つ 測るごとに ★板を 作り直す★
#    ・書き戻しは LF
#
#  ★材料★ A1=1 A2=2 A3=3 / B1=10 B2=20 B3=30
#  ★ジャマ★ … まっさら=何も置かない / ジャマ=E2に 9 / ジャマ字=E2に "あ" /
#               ジャマ空字=E2に ="" / 下が式=E2に =1+1
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-afure.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-afure-2026-09-13.tsv'

$試 = @(
  @{ 式='=A1:A3'; ジャマ='まっさら'; 何='★四角を そのまま 打つ（下に 広がるか）★' },
  @{ 式='=A1:A3*2'; ジャマ='まっさら'; 何='★四角に 掛ける★' },
  @{ 式='=A1:B2'; ジャマ='まっさら'; 何='★2列×2行（横にも 広がるか）★' },
  @{ 式='=A1:A3+B1:B3'; ジャマ='まっさら'; 何='★四角どうしを 足す★' },
  @{ 式='=A1:A3+B1:B2'; ジャマ='まっさら'; 何='★形が 違う 四角どうし★' },
  @{ 式='=A1:A3'; ジャマ='ジャマ'; 何='★溢れる先に 物が 在る（#SPILL! か）★' },
  @{ 式='=A1:A3'; ジャマ='ジャマ字'; 何='★溢れる先に 字が 在る★' },
  @{ 式='=A1:A3'; ジャマ='ジャマ空字'; 何='★溢れる先に 空の字("") が 在る★' },
  @{ 式='=SUM(A1:A3)'; ジャマ='まっさら'; 何='関数は 溢れないはず' },
  @{ 式='=@A1:A3'; ジャマ='まっさら'; 何='★@ を 付けると 1マスに なるか（暗黙の交差）★' },
  @{ 式='=A1:A3&"x"'; ジャマ='まっさら'; 何='★四角に 字を つなぐ★' },
  @{ 式='=-A1:A3'; ジャマ='まっさら'; 何='★四角に 前置きの マイナス★' },
  @{ 式='=A1:A3=1'; ジャマ='まっさら'; 何='★四角の 大小くらべ★' },
  @{ 式='=IF(A1:A3>1,1,0)'; ジャマ='まっさら'; 何='★関数の 引数に 四角（溢れるか）★' },
  @{ 式='=ROW(A1:A3)'; ジャマ='まっさら'; 何='★ROW に 四角★' },
  @{ 式='=A1:A3'; ジャマ='下が式'; 何='★溢れる先が 別の 式★' }
)

# ══ ★★2つ目の 窓（★「0」を 1つの 窓だけで 取らない★）★★ ══
#   ★出る字が「0」でも 5.55E-17 の ことが 在る★（消え残り）
#   ⇒★=(マス)=0 を 打って 聞く★／★型も 一緒に 取る★（String と Number を 取り違えない）
function 窓２_型($v) {
  if ($null -eq $v) { return 'Empty' }
  if ($v -is [string]) { return 'String' }
  if ($v -is [bool]) { return 'Boolean' }
  if ($v -is [double] -or $v -is [int] -or $v -is [long]) { return 'Number' }
  if ($v -is [datetime]) { return 'DateTime' }
  return 'Other'
}
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

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $bk = $xl.Workbooks.Add()

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★1つの 式が 何マスにも 広がる（溢れ）を 実Excel に 聞いた★（2026-09-13）')
  $行.Add('#')
  $行.Add('# ★借り物を 自作で 外す 為＝土台の ⑤枚目★')
  $行.Add('#')
  $行.Add('# ★材料★ A1=1 A2=2 A3=3 / B1=10 B2=20 B3=30')
  $行.Add('# ★式は E1 に 置く★／広がった 先は E1:G5 を なめて 読む')
  $行.Add('#')
  $行.Add('# ★どの Excel か★ … 版 ' + $xl.Version + ' / build ' + $xl.Build)
  $行.Add('#')
  $行.Add('# ★型★と★本当にゼロか★も 取って います（「0」を 1つの 窓だけで 取らない）')
  $行.Add('# 式' + "`t" + 'ジャマ' + "`t" + 'E1の字' + "`t" + '型' + "`t" + '本当にゼロか' + "`t" + '広がった形' + "`t" + '広がった先の字' + "`t" + '何を 見て いるか')

  for ($i = 0; $i -lt $試.Count; $i++) {
    $sh = $bk.Worksheets.Add()
    try {
      $sh.Range('A1').Value2 = 1
      $sh.Range('A2').Value2 = 2
      $sh.Range('A3').Value2 = 3
      $sh.Range('B1').Value2 = 10
      $sh.Range('B2').Value2 = 20
      $sh.Range('B3').Value2 = 30

      $j = $試[$i].ジャマ
      if ($j -eq 'ジャマ')       { $sh.Range('E2').Value2 = 9 }
      elseif ($j -eq 'ジャマ字') { $sh.Range('E2').Value2 = 'あ' }
      elseif ($j -eq 'ジャマ空字') { $sh.Range('E2').Formula = '=""' }
      elseif ($j -eq '下が式')   { $sh.Range('E2').Formula = '=1+1' }

      $ok = $true
      try { $sh.Range('E1').Formula2 = $試[$i].式 } catch { $ok = $false }

      if (-not $ok) {
        $行.Add($試[$i].式 + "`t" + $j + "`t" + '★打てない★' + "`t" + '—' + "`t" + '—' + "`t" + $試[$i].何)
        Write-Host ('  ' + $試[$i].式.PadRight(18) + ' ' + $j.PadRight(10) + ' -> ★打てない★')
        continue
      }

      $字 = [string]$sh.Range('E1').Text

      # ★広がった 形★ … HasArray が 真なら CurrentArray の 大きさ
      $形 = '1x1'
      try {
        if ($sh.Range('E1').HasArray) {
          $ca = $sh.Range('E1').CurrentArray
          $形 = [string]$ca.Rows.Count + 'x' + [string]$ca.Columns.Count
        }
      } catch { $形 = '★読めない★' }

      # ★広がった 先の 字★ … E1:G5 を なめて 空でない 物を 並べる
      $先 = @()
      foreach ($r in 1..5) {
        foreach ($c in 5..7) {
          if ($r -eq 1 -and $c -eq 5) { continue }
          $cell = $sh.Cells.Item($r, $c)
          $t = [string]$cell.Text
          if ($t -ne '') { $先 += ($cell.Address($false,$false) + '=' + $t) }
        }
      }
      $先字 = if ($先.Count -eq 0) { '—' } else { $先 -join ' ' }

      $型 = 窓２_型 $sh.Range('E1').Value2
      $ゼ = if ($字 -match '^-?0(\.0+)?$') { 窓２_本当にゼロか $sh 'E1' } else { '—' }
      $行.Add($試[$i].式 + "`t" + $j + "`t" + $字 + "`t" + $型 + "`t" + $ゼ + "`t" + $形 + "`t" + $先字 + "`t" + $試[$i].何)
      Write-Host ('  ' + $試[$i].式.PadRight(18) + ' ' + $j.PadRight(10) + ' -> E1=' + $字.PadRight(10) + ' 形=' + $形.PadRight(6) + ' ' + $先字)
    } finally {
      $sh.Delete() | Out-Null
    }
  }

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★書いた … ' + $出 + '★')
  $bk.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
