# toru-hashira-haba.ps1 — ★列の 幅を 実Excel に 聞く★（2026-09-11）
#  ★なぜ★ SheetJS の `wpx` は ★実Excel と 違います★
#    実測 … 列A（5字）を 実Excel は ★44.5px★／SheetJS は ★78px★
#    ⇒ そのまま 使うと ★#### に ならない★（幅が 足りて しまう）
#  ★やる事★ … 幅を 何通りも 作って ★実Excel の 点★と ★ファイルの 中の 数★を 並べる
#  ★司さんの 実物には 触りません★＝新しい ブックを 作って 保存するだけ
$ErrorActionPreference='Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$ファイル = Join-Path $ここ 'hashira-haba-2026-09-11.xlsx'
$出 = Join-Path $ここ 'golden-hashira-haba-2026-09-11.tsv'
$幅ら = @(2, 3, 5, 8.43, 10, 12, 15, 20, 30)

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

$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
try{
  $bk=$xl.Workbooks.Add(); $sh=$bk.Worksheets.Item(1)
  for($i=0; $i -lt $幅ら.Count; $i++){
    $sh.Columns.Item([int]($i+1)).ColumnWidth = [double]$幅ら[$i]
    $sh.Cells.Item(1, [int]($i+1)).Value2 = 123456789
  }
  $bk.SaveAs($ファイル, 51)

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★列の 幅を 実Excel に 聞いた★（2026-09-11）')
  $行.Add('#')
  $行.Add('# ★なぜ★ SheetJS の wpx は 実Excel と 違う（列A 5字 … 実Excel 44.5px ／ SheetJS 78px）')
  $行.Add('#   ⇒ そのまま 使うと ★#### に ならない★')
  $行.Add('#')
  $行.Add('# ★字体★ … ' + $bk.Styles.Item('Normal').Font.Name + ' / ' + $bk.Styles.Item('Normal').Font.Size + 'pt')
  $行.Add('# ★標準の 幅★ … ' + $sh.StandardWidth + ' 字')
  $行.Add('# ★どの Excel か★ … 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('#')
  $行.Add('# 列' + "`t" + '打った 字数' + "`t" + 'ColumnWidth(字)' + "`t" + 'Width(ポイント)' + "`t" + '実Excel の 点(px)' + "`t" + '出る字' + "`t" + '型' + "`t" + '窓２(=(A1)=0)')
  for($i=0; $i -lt $幅ら.Count; $i++){
    $col=$sh.Columns.Item([int]($i+1))
    $px=[math]::Round($col.Width*96/72,2)
    $t=[string]$sh.Cells.Item(1,[int]($i+1)).Text
    # ★「0」に 見えた 時だけ 実Excel に 聞く★（狭い 列では 0で ない 数が 0 に 見える）
    $型 = 窓２_型 $sh.Cells.Item(1,[int]($i+1)).Value2
    $窓2 = if ($t -match '^-?0(\.0+)?$') { 窓２_本当にゼロか $sh ('R1C' + ($i+1)) } else { '—' }
    $行.Add(($i+1).ToString() + "`t" + $幅ら[$i] + "`t" + $col.ColumnWidth + "`t" + $col.Width + "`t" + $px + "`t" + $t + "`t" + $型 + "`t" + $窓2)
    Write-Host ('  列' + ($i+1) + ' 打った=' + $幅ら[$i] + ' → ' + $col.ColumnWidth + '字 / ' + $px + 'px / 出る字=' + $t)
  }
  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★書いた … ' + $出 + '★')
  $bk.Close($false)
} finally { $xl.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }

# ★作者の 欄を「見本」に（10-10 司さんの 決め＝実在の 名前は 架空に）★
#   Excel は サインインの アカウント名を 入れるので、Excel を 閉じた 後に scripts/sakusha-mihon.mjs で 替える
& node (Join-Path $PSScriptRoot '../../scripts/sakusha-mihon.mjs') $ファイル
if ($LASTEXITCODE -ne 0) { throw '★作者の 欄を 見本に 替えられない★' }
