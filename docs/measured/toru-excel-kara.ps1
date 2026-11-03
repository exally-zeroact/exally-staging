# toru-excel-kara.ps1 — ★実Excel に 材料を 作らせて 測る★（うち で 開く 前の 正解を 取る）
#
#  ★★なぜ 要るか★★
#    今まで 確かめた 向きは 2つだけ
#      ① 実Excel に 打たせて 答えを 合わせる（Excel → うち・値だけ）
#      ② うちが 書いた 物を Excel で 開き直す（うち → Excel）
#    ★③ Excel が 書いた ファイルを うちで 開く★ … ★一度も 測って いません★
#    ⇒★お客さんが 一番 よく やる 道★（Excel で 作った 表を Exally に 持ち込む）
#
#  ★★測る 事（3つ とも）★★
#    ①★式★(.Formula2) ②★出る字★(.Text) ③★答え★(.Value2)
#
#  ★物差しの 決まり★
#    ・★司さんの 実物には 触りません★＝新しい ブックを 作って 名前を 付けて 保存するだけ
#    ・書き戻しは LF
#
#  使い方: pwsh -NoProfile -File toru-excel-kara.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$ファイル = Join-Path $ここ 'excel-kara-2026-09-11.xlsx'
$出 = Join-Path $ここ 'golden-excel-kara-2026-09-11.tsv'

# ★材料★ … マス ／ 入れる 物 ／ 何を 見て いるか
$材料 = @(
  @{ マス='D1'; 入=  '=XIRR(A1:A3,B1:B3)';     何='お金の 利回り' },
  @{ マス='D2'; 入=  '=MIRR(A1:A3,0.1,0.12)';  何='お金の 利回り' },
  @{ マス='D3'; 入=  '=XNPV(0.1,A1:A3,B1:B3)'; 何='お金の 今の 値' },
  @{ マス='D4'; 入=  '=1/3';                   何='割り切れない 数の 出る字' },
  @{ マス='D5'; 入=  '=206800/1.1';            何='税抜き' },
  @{ マス='D6'; 入=  '=1.64E-14';              何='とても 小さい 数' },
  @{ マス='G1'; 入=  '=SORT(E1:E5)';           何='溢れ（並べ替え）' },
  @{ マス='H1'; 入=  '=UNIQUE(E1:E5)';         何='溢れ（重なりを 取る）' },
  @{ マス='I1'; 入=  '=SEQUENCE(3)';           何='溢れ（数を 並べる）' },
  @{ マス='J1'; 入=  '=TEXT(E1,"#,##0.00")';   何='書式の 字' },
  @{ マス='K1'; 入=  '=TODAY()-TODAY()';       何='日付の 引き算（0 に なる）' }
)

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
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Worksheets.Item(1)

  # ★お金の 流れと 日付★
  $流 = @(-1000, 600, 700); $日 = @(45292, 45383, 45474)
  for ($i = 0; $i -lt 3; $i++) {
    $sh.Range('A' + ($i+1)).Value2 = $流[$i]
    $sh.Range('B' + ($i+1)).Value2 = $日[$i]
    $sh.Range('B' + ($i+1)).NumberFormatLocal = 'yyyy/m/d'
  }
  # ★並べ替えの 種（★重なりを わざと 入れる★）★
  $種 = @(3, 1, 5, 1, 4)
  for ($i = 0; $i -lt 5; $i++) { $sh.Range('E' + ($i+1)).Value2 = $種[$i] }

  foreach ($m in $材料) { $sh.Range($m.マス).Formula2 = $m.入 }

  # ★書式・結合・狭い 列★（うちで 開いた時に 崩れないか）
  $sh.Range('A5').Value2 = 1234567.891
  $sh.Range('A5').NumberFormatLocal = '#,##0.00'
  $sh.Range('A6').Value2 = 0.1234
  $sh.Range('A6').NumberFormatLocal = '0.0%'
  $sh.Range('C8:D8').Merge()
  $sh.Range('C8').Value2 = 'むすんだ マス'
  $sh.Range('A8').Value2 = 100000
  $sh.Columns.Item(1).ColumnWidth = 5   # ★狭くして ##### を 出させる★

  $bk.SaveAs($ファイル, 51)   # 51 = xlOpenXMLWorkbook (.xlsx)

  # ★材料 ＋ 書式の マス を 測る★
  $見る = @()
  foreach ($m in $材料) { $見る += $m }
  $見る += @{ マス='G2'; 入='(溢れた 先)'; 何='溢れの 先' }
  $見る += @{ マス='G3'; 入='(溢れた 先)'; 何='溢れの 先' }
  $見る += @{ マス='H2'; 入='(溢れた 先)'; 何='溢れの 先（重なりを 取る）' }
  $見る += @{ マス='I2'; 入='(溢れた 先)'; 何='溢れの 先（数を 並べる）' }
  $見る += @{ マス='A5'; 入='1234567.891 ＋ #,##0.00'; 何='桁区切りの 書式（★列は 狭い★）' }
  $見る += @{ マス='A6'; 入='0.1234 ＋ 0.0%';          何='割合の 書式（★列は 狭い★）' }
  $見る += @{ マス='A8'; 入='100000';                 何='★狭い 列＝##### に なるか★' }
  $見る += @{ マス='C8'; 入='むすんだ マス';           何='結合した マス' }

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★実Excel が 作った ファイルの「正解」★（2026-09-11）')
  $行.Add('#')
  $行.Add('# ★向き★ … ★Excel が 書いた 物を うちで 開く★（★今日まで 一度も 測って いません★）')
  $行.Add('#   ①実Excel に 打たせて 答えを 合わせる … やって いた')
  $行.Add('#   ②うちが 書いた 物を Excel で 開き直す … 2026-09-10 に 足した')
  $行.Add('#   ③★Excel が 書いた 物を うちで 開く★ … ★これ★')
  $行.Add('#')
  $行.Add('# ★3つ とも 見る★ ①式(.Formula2) ②出る字(.Text) ③答え(.Value2)')
  $行.Add('#')
  $行.Add('# ★どの Excel か★ … 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★作った ファイル★ … ' + (Split-Path -Leaf $ファイル))
  $行.Add('#')
  $行.Add('# マス' + "`t" + '入れた 物' + "`t" + '実Excel の 式' + "`t" + '出る字' + "`t" + '答え' + "`t" + '何を 見て いるか')

  foreach ($m in $見る) {
    $c = $sh.Range($m.マス)
    $式 = [string]$c.Formula2
    $字 = [string]$c.Text
    $v = $c.Value2
    $答 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    $型 = 窓２_型 $v
    # ★「0」に 見えた 時だけ 実Excel に 聞く★（狭い 列では 0で ない 数が 0 に 見える）
    $窓2 = if ($字 -match '^-?0(\.0+)?$') { 窓２_本当にゼロか $sh $m.マス } else { '—' }
    $行.Add($m.マス + "`t" + $m.入 + "`t" + $式 + "`t" + $字 + "`t" + $答 + "`t" + $型 + "`t" + $窓2 + "`t" + $m.何)
    Write-Host ('  ' + $m.マス.PadRight(4) + ' 式=' + $式.PadRight(26) + ' 出る字=' + $字.PadRight(16) + ' 答え=' + $答)
  }

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★書いた … ' + $出 + '★')
  Write-Host ('★材料 … ' + $ファイル + '★')
  $bk.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}

# ★作者の 欄を「見本」に（10-10 司さんの 決め＝実在の 名前は 架空に）★
#   Excel は サインインの アカウント名を 入れるので、Excel を 閉じた 後に scripts/sakusha-mihon.mjs で 替える
& node (Join-Path $PSScriptRoot '../../scripts/sakusha-mihon.mjs') $ファイル
if ($LASTEXITCODE -ne 0) { throw '★作者の 欄を 見本に 替えられない★' }
