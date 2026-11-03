# toru-sashikomi.ps1 — ★行/列を 入れたり 消したりした時、式が どう 追従するか★（2026-09-13）
#
#  ★★なぜ 要るか★★
#    借り物(HyperFormula・GPLv3)を ★自作で 外す★（司さん 2026-09-13）。
#    借り物の 中で ★一番 手が 込んでいる★のが ここ＝
#    ★行を 1本 入れただけで 表じゅうの 式の 参照が ずれる★。
#    （借り物では LazilyTransformingAstService が やって いる所）
#    ⇒★当て推量で 作ると 全部 外す★ので 実Excel に 1つずつ 聞きます。
#
#  ★物差しの 決まり★
#    ・★式そのもの（.Formula）を 見る★＝答えの 値では なく ★参照が どう 書き換わったか★
#    ・★式を 置く マスも 動く★ ので ★置いた後の マスを 追って 読む★
#    ・★司さんの 実物には 触りません★＝新しい ブックを 作って 保存せず 閉じる
#    ・1つ 測るごとに ★板を 作り直す★（前の 測りが 残らない様に）
#    ・書き戻しは LF
#
#  ★材料★ A1=1 A2=2 A3=3 A4=4 ／ B1=10 B2=20 ／ C1=100
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-sashikomi.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-sashikomi-2026-09-13.tsv'

$試 = @(
  @{ 式='=A2+1'; わざ='Rows(2).Insert'; 何='★指す先の 上に 行を 入れる（A2→A3 に なるか）★' },
  @{ 式='=A2+1'; わざ='Rows(3).Insert'; 何='指す先の 下に 入れる（変わらないはず）' },
  @{ 式='=A2+1'; わざ='Rows(2).Delete'; 何='★指す先を 消す（#REF! に なるか）★' },
  @{ 式='=A3+1'; わざ='Rows(2).Delete'; 何='指す先の 上を 消す（A3→A2 に なるか）' },
  @{ 式='=$A$2+1'; わざ='Rows(2).Insert'; 何='★絶対参照も 追従するか★' },
  @{ 式='=SUM(A1:A3)'; わざ='Rows(2).Insert'; 何='★四角の 途中に 入れる（A1:A4 に 広がるか）★' },
  @{ 式='=SUM(A2:A4)'; わざ='Rows(2).Insert'; 何='★四角の 先頭に 入れる（A3:A5 か A2:A5 か）★' },
  @{ 式='=SUM(A1:A3)'; わざ='Rows(4).Insert'; 何='四角の 下に 入れる（変わらないはず）' },
  @{ 式='=SUM(A1:A3)'; わざ='Rows(2).Delete'; 何='★四角の 中を 消す（A1:A2 に 縮むか）★' },
  @{ 式='=SUM(A2:A2)'; わざ='Rows(2).Delete'; 何='★四角が 丸ごと 消える（#REF! か）★' },
  @{ 式='=A2+1'; わざ='Columns(1).Insert'; 何='★左に 列を 入れる（A2→B2 か）★' },
  @{ 式='=B2+1'; わざ='Columns(1).Insert'; 何='列を 入れる（B2→C2 か）' },
  @{ 式='=B2+1'; わざ='Columns(2).Delete'; 何='★指す列を 消す（#REF! か）★' },
  @{ 式='=SUM(A1:C1)'; わざ='Columns(2).Insert'; 何='★横の 四角の 途中に 列を 入れる★' },
  @{ 式='=A1+A2'; わざ='Rows(1).Insert'; 何='★式 自身が 下に ずれる時★' },
  @{ 式='=A$2+1'; わざ='Rows(2).Insert'; 何='★行だけ 絶対★' },
  @{ 式='=$A2+1'; わざ='Rows(2).Insert'; 何='★列だけ 絶対★' }
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
  $行.Add('# ★行/列を 入れたり 消したりした時、式が どう 追従するか★（2026-09-13）')
  $行.Add('#')
  $行.Add('# ★借り物を 自作で 外す 為★＝ここが 借り物の 中で 一番 手が 込んでいる所')
  $行.Add('#')
  $行.Add('# ★材料★ A1=1 A2=2 A3=3 A4=4 / B1=10 B2=20 / C1=100')
  $行.Add('# ★式は E5 に 置く★（挿入で E5 自身も 動く事が 在るので 追って 読む）')
  $行.Add('#')
  $行.Add('# ★どの Excel か★ … 版 ' + $xl.Version + ' / build ' + $xl.Build)
  $行.Add('#')
  $行.Add('# ★型★と★本当にゼロか★も 取って います（「0」を 1つの 窓だけで 取らない）')
  $行.Add('# 置いた式' + "`t" + 'わざ' + "`t" + '後の式' + "`t" + '後のマス' + "`t" + '出る字' + "`t" + '型' + "`t" + '本当にゼロか' + "`t" + '何を 見て いるか')

  for ($i = 0; $i -lt $試.Count; $i++) {
    $sh = $bk.Worksheets.Add()
    try {
      $sh.Range('A1').Value2 = 1
      $sh.Range('A2').Value2 = 2
      $sh.Range('A3').Value2 = 3
      $sh.Range('A4').Value2 = 4
      $sh.Range('B1').Value2 = 10
      $sh.Range('B2').Value2 = 20
      $sh.Range('C1').Value2 = 100

      $置く = $sh.Range('E5')
      $置く.Formula = $試[$i].わざ -eq '' ? $試[$i].式 : $試[$i].式

      # ★式を 置いた マスを 覚えて おく★（挿入で 動くので 名前で 追えない）
      $前名 = $置く.Address($false, $false)

      # ★わざを かける★
      $w = $試[$i].わざ
      if ($w -like 'Rows(*).Insert')    { $n = [int]($w -replace '[^0-9]',''); $sh.Rows.Item($n).Insert() | Out-Null }
      elseif ($w -like 'Rows(*).Delete') { $n = [int]($w -replace '[^0-9]',''); $sh.Rows.Item($n).Delete() | Out-Null }
      elseif ($w -like 'Columns(*).Insert') { $n = [int]($w -replace '[^0-9]',''); $sh.Columns.Item($n).Insert() | Out-Null }
      elseif ($w -like 'Columns(*).Delete') { $n = [int]($w -replace '[^0-9]',''); $sh.Columns.Item($n).Delete() | Out-Null }

      # ★式が どこへ 動いたかを 探す★（E列の 前後・5行目の 前後を なめる）
      $後マス = ''
      $後式 = ''
      $後字 = ''
      foreach ($r in 4..7) {
        foreach ($c in 4..7) {
          $cell = $sh.Cells.Item($r, $c)
          $f = [string]$cell.Formula
          if ($f.StartsWith('=')) {
            $後マス = $cell.Address($false, $false)
            $後式 = $f
            $後字 = [string]$cell.Text
            break
          }
        }
        if ($後式 -ne '') { break }
      }
      if ($後式 -eq '') { $後式 = '★式が 消えた★'; $後マス = '—'; $後字 = '—' }

      $型 = if ($後マス -eq '—') { '—' } else { 窓２_型 $sh.Range($後マス).Value2 }
      $ゼ = if ($後字 -match '^-?0(\.0+)?$') { 窓２_本当にゼロか $sh $後マス } else { '—' }
      $行.Add($試[$i].式 + "`t" + $w + "`t" + $後式 + "`t" + $後マス + "`t" + $後字 + "`t" + $型 + "`t" + $ゼ + "`t" + $試[$i].何)
      Write-Host ('  ' + $試[$i].式.PadRight(16) + ' ' + $w.PadRight(20) + ' -> ' + $後式.PadRight(20) + ' (' + $後マス + ') ' + $後字)
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
