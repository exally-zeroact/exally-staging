# toru-jitsubutsu-deruji.ps1 — ★司さんの実物の「出る字」を 実Excel に 聞く★（2026-09-11）
#
#  ★★なぜ★★
#    今まで 実物で 合わせたのは ★答え★だけ（2026-08-29 … 19,323/19,323 一致）。
#    ★出る字（画面に 出る 字）は 一度も 突き合わせて いません★。
#    2026-09-11 に 小さい ファイル 19本で 測ったら ★出る字だけで 本当の 穴が 4つ★ 出ました
#      （列の 幅を 読んで いない／1,234,567.89 が "1.00"／TEXT の 3.00 が "3"／溢れ）。
#    ⇒★実物 丸ごとで 出る字を 測ります★
#
#  ★司さんの 実物には 触りません★
#    ★写しを 開きます★（scratchpad の jitsubutsu.xlsb）／保存せず 閉じる
#
#  ★測る 物★ … 使って いる 範囲の ★空で ない マス 全部★
#    ①出る字(.Text) ②型 ③列の 幅(点) ④書式
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-jitsubutsu-deruji.ps1 [板の名前]

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$写し = 'C:\Users\zeroa\AppData\Local\Temp\claude\C--WINDOWS-System32-WindowsPowerShell-v1-0\5b4e50e6-20a1-4af5-8ffb-8b6d6ca3f52b\scratchpad\jitsubutsu.xlsb'
$出 = Join-Path $ここ 'golden-jitsubutsu-deruji-2026-09-11.tsv'
$板名 = if ($args.Count -ge 1) { [string]$args[0] } else { '' }

if (-not (Test-Path $写し)) { Write-Error ('★写しが 無い … ' + $写し + '★'); exit 2 }

function 型を見る($v) {
  if ($null -eq $v) { return 'Empty' }
  if ($v -is [string]) { return 'String' }
  if ($v -is [bool]) { return 'Boolean' }
  if ($v -is [double] -or $v -is [int] -or $v -is [long]) { return 'Number' }
  return 'Other'
}

# ★★この 下の 2つは docs/measured/toru-shisuu-mitame.ps1 から そのまま 写しました★★
#   ★1文字も 変えて いません★（見張り … tests/monosashi-mado.test.mjs）
# ══ ★★2つ目の 窓（★「0」を 1つの 窓だけで 取らない★）★★ ══
#   ★この 道具は 画面に 出る 字（.Text）を 見ます★
#   ⇒★狭い 列では ★0で ない 小さい 数が「0」に 見える★★
#     （例 … 幅が 足りないと 実Excel は 丸めて 出す）
#   ⇒★字が「0」に 見えた 時は 実Excel に ★=(A1)=0★ を 打って
#     ★本当に 0 か★を 聞く★（.Value2 だけ／字だけ では 見分けが 付かない）
#   ⇒★型も 一緒に 取る★（String と Number を 取り違えない）
function 窓２_型($v) {
  if ($null -eq $v) { return 'Empty' }
  if ($v -is [string]) { return 'String' }
  if ($v -is [bool]) { return 'Boolean' }
  if ($v -is [double] -or $v -is [int] -or $v -is [long]) { return 'Number' }
  return 'Other'
}
function 窓２_本当にゼロか($sh, [string]$マス) {
  try {
    $sh.Range('BZ1').Clear() | Out-Null
    $sh.Range('BZ1').Formula = ('=(' + $マス + ')=0')
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
  $bk = $xl.Workbooks.Open($写し, 0, $true)   # ReadOnly
  # ★板は 番号でも 指せます★＝名前は 貝殻で 文字化けして 通らない事が 在る
  $sh = if ($板名 -match '^[0-9]+$') { $bk.Worksheets.Item([int]$板名) }
        elseif ($板名) { $bk.Worksheets.Item($板名) }
        else { $bk.Worksheets.Item(1) }
  Write-Host ('★板★ ' + $sh.Name)

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★司さんの実物の「出る字」を 実Excel に 聞いた★（2026-09-11）')
  $行.Add('#')
  $行.Add('# ★なぜ★ 今まで 実物で 合わせたのは ★答え★だけ（08-29 … 19,323/19,323）。')
  $行.Add('#   ★出る字は 一度も 突き合わせて いません★。')
  $行.Add('#   09-11 に 小さい ファイル 19本で 測ったら ★出る字だけで 穴が 4つ★ 出た。')
  $行.Add('#')
  $行.Add('# ★写しを 読むだけ★（司さんの 実物には 触って いません）')
  $行.Add('# ★どの Excel か★ … 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★板★ … ' + $sh.Name)
  $行.Add('#')
  $行.Add('# マス' + "`t" + '出る字' + "`t" + '型' + "`t" + '列の点' + "`t" + '書式' + "`t" + '窓２(=(A1)=0)')

  $範囲 = $sh.UsedRange
  $r0 = $範囲.Row; $c0 = $範囲.Column
  $rn = $範囲.Rows.Count; $cn = $範囲.Columns.Count
  Write-Host ('★使って いる 範囲★ ' + $rn + '行 × ' + $cn + '列')

  # ★一気に 読む★（1マスずつ COM で 聞くと 遅い）
  $字ら = $範囲.Text
  $値ら = $範囲.Value2
  $数 = 0
  for ($i = 1; $i -le $rn; $i++) {
    for ($j = 1; $j -le $cn; $j++) {
      $v = if ($rn -eq 1 -and $cn -eq 1) { $値ら } else { $値ら[$i, $j] }
      if ($null -eq $v) { continue }
      $c = $sh.Cells.Item($r0 + $i - 1, $c0 + $j - 1)
      $字 = [string]$c.Text
      if ($字 -eq '') { continue }
      $型 = 窓２_型 $v
      $幅 = [math]::Round($c.EntireColumn.Width * 96 / 72, 2)
      $書 = [string]$c.NumberFormat
      $マス = $c.Address(0, 0)
      # ★「0」に 見えた 時だけ 実Excel に 聞く★（狭い 列では 0で ない 数が 0 に 見える）
      $窓2 = if ($字 -match '^-?0(\.0+)?$') { 窓２_本当にゼロか $sh $マス } else { '—' }
      $行.Add($マス + "`t" + $字 + "`t" + $型 + "`t" + $幅 + "`t" + $書 + "`t" + $窓2)
      $数++
    }
  }

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★空で ない マス … ' + $数 + '個★')
  Write-Host ('★書いた … ' + $出 + '★')
  $bk.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
