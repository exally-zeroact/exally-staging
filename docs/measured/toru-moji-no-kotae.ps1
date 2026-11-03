# toru-moji-no-kotae.ps1 — ★文字を 返す 式の 答えを 実Excel に 聞く★（2026-09-11）
#
#  ★★なぜ 要るか★★
#    `=TEXT(A1,"0.00")` の 答えは 実Excel では ★文字の "3.00"★（LEN=4／ISTEXT=TRUE）。
#    うちは ★答えは 文字で 持って いる★のに、★画面に 出す 時に 数として 読み直して★
#    ★"3" と 出して いました★（2026-09-11 実測）。
#    ⇒★TEXT だけの 話では ありません★
#      `=LEFT("3.00",4)` `=RIGHT("12.50",5)` `=TRIM("  4.50  ")` … ★文字を 返す 式は 全部★
#      （`="007"` だけは ★頭の ゼロ★の 決まりで たまたま 守られて いた）
#
#  ★物差しの 決まり★
#    ・★3つ とも 見る★ ①出る字(.Text) ②型 ③窓２(`=(A1)=0`)
#    ・★司さんの 実物には 触りません★＝新しい ブックを 作って 保存せず 閉じる
#    ・書き戻しは LF
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-moji-no-kotae.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-moji-no-kotae-2026-09-11.tsv'

# ★見る 式★ … 式 ／ 何を 見て いるか
$試 = @(
  @{ 式='=TEXT(A1,"0.00")';            何='書式の 字（うちは 3 に なって いた）' },
  @{ 式='=TEXT(A1,"000")';             何='頭に ゼロを 足す' },
  @{ 式='=TEXT(1234.5,"#,##0.00")';    何='桁区切り つき' },
  @{ 式='=LEFT("3.00",4)';             何='左から 切る' },
  @{ 式='=RIGHT("12.50",5)';           何='右から 切る' },
  @{ 式='=MID("x12.50y",2,5)';         何='真ん中を 切る' },
  @{ 式='=A1&""';                      何='つなげる（数 → 文字）' },
  @{ 式='=CONCATENATE("1",".","50")';  何='つなげる（3つ）' },
  @{ 式='=REPT("1",3)';                何='繰り返す' },
  @{ 式='=TRIM("  4.50  ")';           何='前後の 空白を 取る' },
  @{ 式='=SUBSTITUTE("1,5",",",".")';  何='置き換える' },
  @{ 式='=LEN(TEXT(A1,"0.00"))';       何='★数を 返す★（4 で なければ 文字で 持って いない）' },
  @{ 式='=ISTEXT(TEXT(A1,"0.00"))';    何='★真偽を 返す★（TRUE で なければ 文字で ない）' },
  @{ 式='=A1*2';                       何='★数を 返す★（ここは 数の まま で 正しい）' }
)

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
  $sh.Range('A1').Value2 = 3

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★文字を 返す 式の 答えを 実Excel に 聞いた★（2026-09-11）')
  $行.Add('#')
  $行.Add('# ★なぜ★ うちは 答えを 文字で 持って いるのに、画面に 出す 時に')
  $行.Add('#   ★数として 読み直して★ いた ⇒ "3.00" が ★"3"★ に なる')
  $行.Add('#   ★TEXT だけの 話では ない★＝LEFT／RIGHT／TRIM … 文字を 返す 式は 全部')
  $行.Add('#')
  $行.Add('# ★3つ とも 見る★ ①出る字(.Text) ②型 ③窓２(=(A1)=0)')
  $行.Add('#')
  $行.Add('# ★どの Excel か★ … 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★A1 = 3★')
  $行.Add('#')
  $行.Add('# 式' + "`t" + '出る字' + "`t" + '型' + "`t" + '窓２(=(A1)=0)' + "`t" + '何を 見て いるか')

  for ($i = 0; $i -lt $試.Count; $i++) {
    $c = $sh.Cells.Item([int]($i + 3), 2)
    $c.Formula2 = $試[$i].式
    $字 = [string]$c.Text
    $v = $c.Value2
    $型 = 窓２_型 $v
    $窓2 = if ($字 -match '^-?0(\.0+)?$') { 窓２_本当にゼロか $sh ('R' + ($i + 3) + 'C2') } else { '—' }
    $行.Add($試[$i].式 + "`t" + $字 + "`t" + $型 + "`t" + $窓2 + "`t" + $試[$i].何)
    Write-Host ('  ' + $試[$i].式.PadRight(30) + ' 出る字=' + $字.PadRight(12) + ' 型=' + $型)
  }

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★書いた … ' + $出 + '★')
  $bk.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
