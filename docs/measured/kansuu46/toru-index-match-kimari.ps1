# toru-index-match-kimari.ps1 — ★INDEX と MATCH の 決まりを 実Excel に 聞く★（2026-09-15）
#
#  ★★なぜ 別に 取るか★★
#    総当たりの 紙（`golden-index-match2-2026-09-15.tsv`）には
#    ★字の 表が 1つも 在りません★（材料は 数と 日付だけ）。
#    ⇒★客の 使い方＝「字で 探して 隣を 取る」が 1本も 測れて いません★。
#    ⇒ ここは ★総当たりでは なく「決まりを 名指しで 測る」★ 紙です。
#
#  ★★この 道具が 守る 事★★（2026-09-15 に 踏んだ 穴）
#    ①★前の 式の 跡を 消してから 打つ★
#    ②★CalculateFull してから 読む★
#      ＝消さずに 測る 道具は ★同じ 式に 2回で 違う 答え★を 返しました
#        （`=MATCH(3,A1:A5,2)` … 1回目 #N/A ／ 2回目 3）
#    ③★打った マスを 1本ずつ 紙に 書く★
#      ＝実Excel は ★四角を 1つの 値に 詰める 時、式の 行／列と 交わる 所★を 採る（暗黙の 交わり）
#        ⇒★同じ 式でも 打つ マスで 答えが 変わります★
#
#  使い方: powershell -NoProfile -File docs/measured/kansuu46/toru-index-match-kimari.ps1
$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-index-match-kimari-2026-09-15.tsv'

$誤りの番号 = @{
  -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!';
  -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!';
  -2146826246 = '#N/A';   -2146826243 = '#SPILL!'; -2146826238 = '#CALC!'; -2146826237 = '#BUSY!'
}

# ══ ★★2つ目の 窓★★ ══（2026-09-15・見張り `tests/monosashi-mado.test.mjs` の 決め）
#  ★物差しの 欠陥★ `.Value2` は ★0 で ない 値に 0 を 返します★
#    =0.1+0.2-0.3    … .Value2 ★0★ ／ =(式)=0 ★False★
#    =11.1+22.2-33.3 … .Value2 0   ／ =(式)=0 ★True★
#    ⇒★.Value2 だけでは この 2つが 同じ 顔に なります★
#  ⇒★『0』が 出た 時だけ ★=(式)=0 の 真偽★も 取って 紙に 並べます★
#    （この 紙にも 0 が 出ます … `=INDEX(E1:E5,4)`＝★空マスを 指した 時★）
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

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
$版 = $xl.Version; $ビルド = $xl.Build
$wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)

# ★材料★
#   A … 昇順の 数 ／ B … 隣の 列（取りに 行く 先） ／ C … 降順の 数
#   D … 字（★大文字小文字が 混ざって います★） ／ E … 並んで いない＋★E4 は 空★
#   G … 真偽
$数の材料 = @(
  @('A1',1), @('A2',2), @('A3',3), @('A4',4), @('A5',5),
  @('B1',2), @('B2',4), @('B3',6), @('B4',8), @('B5',10),
  @('C1',5), @('C2',4), @('C3',3), @('C4',2), @('C5',1),
  @('E1',10), @('E2',30), @('E3',20), @('E5',50)
)
foreach ($z in $数の材料) { $ws.Range($z[0]).Value2 = [double]$z[1] }
$ws.Range('D1').Value2 = 'apple'
$ws.Range('D2').Value2 = 'Banana'
$ws.Range('D3').Value2 = 'cherry'
$ws.Range('D4').Value2 = 'date'
$ws.Range('D5').Value2 = 'egg'
$ws.Range('G1').Value2 = $true
$ws.Range('G2').Value2 = $false

# ★測る 式★ … @(打つマス, 式, 何を 測って いるか)
$式たち = @(
  @('J1', '=MATCH("banana",D1:D5,0)',      '字の ぴたり＝大文字小文字を 区別するか'),
  @('J1', '=MATCH("BANANA",D1:D5,0)',      '同上（全部 大文字）'),
  @('J1', '=MATCH("cherry",D1:D5,0)',      '字の ぴたり（そのまま）'),
  @('J1', '=MATCH("b*",D1:D5,0)',          'ワイルドカード ＊'),
  @('J1', '=MATCH("?anana",D1:D5,0)',      'ワイルドカード ？'),
  @('J1', '=MATCH("z*",D1:D5,0)',          'ワイルドカードで 見つからない'),
  @('J1', '=MATCH(3,C1:C5,-1)',            '降順＋型 -1（決まり通り）'),
  @('J1', '=MATCH(3,C1:C5,1)',             '降順に 型 1（決まりを 破ると どうなる）'),
  @('J1', '=MATCH(20,E1:E5,0)',            '並んで いない＋ぴたり'),
  @('J1', '=MATCH(20,E1:E5,1)',            '並んで いない＋型 1'),
  @('J1', '=MATCH(3,A1:A5,2)',             '型 2（-1/0/1 の 外）'),
  @('J1', '=MATCH(3,A1:A5,-2)',            '型 -2'),
  @('J1', '=MATCH(3,A1:A5,0.5)',           '型 0.5（小数）'),
  @('J1', '=MATCH(3,A1:A5,1.9)',           '型 1.9'),
  @('J1', '=MATCH(3,A1:A5,-0.5)',          '型 -0.5'),
  @('J1', '=MATCH(TRUE,G1:G2,0)',          '真偽を 探す'),
  @('J1', '=MATCH("3",A1:A5,0)',           '字で 数を 探す'),
  @('J1', '=MATCH(3,"3",0)',               '数で 字を 探す'),
  @('J1', '=MATCH(3,{1,2,3,4,5},0)',       '直に 書いた 表'),
  @('J1', '=MATCH(4,E1:E5,0)',             '空マスを 跨いで 無い'),
  @('J1', '=MATCH("",D1:D5,0)',            '空の 字'),
  @('J1', '=MATCH(1,A1:A5,-1)',            '昇順に 型 -1（先頭が 当たる）'),
  @('J1', '=MATCH(5,A1:A5,1)',             '一番 後ろ'),
  @('J1', '=MATCH(30,E1:E5,0)',            '空マスの 手前'),
  @('J1', '=MATCH(50,E1:E5,0)',            '空マスの 後ろ'),
  @('J1', '=INDEX(D1:D5,2)',               '字を 返す'),
  @('J1', '=INDEX(E1:E5,4)',               '★空マスを 指す★（0 か 空か）'),
  @('J1', '=INDEX(E1:E5,4)&"|"',           '同上＝字に した時の 見え方'),
  @('J1', '=ISBLANK(INDEX(E1:E5,4))',      '同上＝空の マスとして 見えるか（★参照か★）'),
  @('J1', '=INDEX(A1:A5,1.9)',             '小数の 行番号'),
  @('J1', '=INDEX(A1:A5,"2")',             '字の 行番号'),
  @('J1', '=INDEX(A1:A5,TRUE)',            '真偽の 行番号'),
  @('J1', '=INDEX(A1:B5,3)',               '2次元に 番号 1つ'),
  @('J1', '=INDEX(A1:B5,2,2)',             '行と 列'),
  @('J1', '=INDEX(A1:A5)',                 '引数 1つ'),
  @('J1', '=INDEX(A1:B5,0,0)',             '行も 列も 0'),
  @('J1', '=INDEX(A1:B5,1,1,1)',           '4つ目＝区画 1'),
  @('J1', '=INDEX(A1:B5,1,1,2)',           '区画が 外'),
  @('J1', '=SUM(INDEX(A1:B5,0,2))',        '★参照を 返して いるか★（縦 まるごとを 足す）'),
  @('J1', '=SUM(INDEX(A1:B5,2,0))',        '★横 まるごとを 足す★'),
  @('J1', '=COUNT(INDEX(A1:B5,0,1))',      'まるごとの 個数'),
  @('J1', '=ROWS(INDEX(A1:B5,0,2))',       'まるごとの 行数'),
  @('J1', '=COLUMNS(INDEX(A1:B5,2,0))',    'まるごとの 列数'),
  @('J1', '=CELL("address",INDEX(A1:B5,2,2))', '★1マスでも 参照か★（決め手）'),
  @('J1', '=SUM(INDEX(A1:A5,2):A5)',       '★INDEX の 答えに : が 使えるか★'),
  @('J1', '=ROW(INDEX(A1:B5,3,2))',        '★参照の 行番号★'),
  @('J1', '=INDEX(A1:B5,0,2)',             '暗黙の 交わり … 行1 に 打つ'),
  @('J3', '=INDEX(A1:B5,0,2)',             '同じ 式を 行3 に 打つ'),
  @('J9', '=INDEX(A1:B5,0,2)',             '同じ 式を 行9（材料の 外）に 打つ'),
  @('A8', '=INDEX(A1:B5,2,0)',             '横 まるごとを 列A に 打つ'),
  @('B8', '=INDEX(A1:B5,2,0)',             '同じ 式を 列B に 打つ'),
  @('J1', '=INDEX(B1:B5,MATCH(3,A1:A5,0))', '★客の 使い方★ INDEX＋MATCH（数）'),
  @('J1', '=INDEX(B1:B5,MATCH("cherry",D1:D5,0))', '★客の 使い方★ INDEX＋MATCH（字）'),
  @('J1', '=IFERROR(INDEX(B1:B5,MATCH(99,A1:A5,0)),"ない")', '★客の 使い方★ 見つからない時'),
  @('J1', '=INDEX(B1:B5,MATCH(99,A1:A5,0))', '同上＝IFERROR で 包まない と')
)

$結果 = New-Object System.Collections.ArrayList
$i = 0
foreach ($s in $式たち) {
  $i++
  $マス = $s[0]; $式 = $s[1]; $訳 = $s[2]
  Write-Host ("  … {0}/{1}  {2}" -f $i, $式たち.Count, $マス)
  # ★前の 式の 跡（溢れ）を 消す★／★打つ 場所も 材料の 外だけ★
  $ws.Range('I1:Z30').Clear() | Out-Null
  $ws.Range('A7:B12').Clear() | Out-Null
  $答 = ''; $型 = ''
  try {
    $ws.Range($マス).Formula = $式
    $xl.CalculateFull()
    $v = $ws.Range($マス).Value2
    if ($null -eq $v) { $答 = '(空)'; $型 = 'Empty' }
    elseif (($v -is [int] -or $v -is [long]) -and $誤りの番号.ContainsKey([int]$v)) { $答 = $誤りの番号[[int]$v]; $型 = 'error値' }
    elseif ($v -is [bool]) { $答 = $(if ($v) { 'TRUE' } else { 'FALSE' }); $型 = 'Boolean' }
    elseif ($v -is [double]) { $答 = $v.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture); $型 = 'Double' }
    elseif ($v -is [object[,]]) { $答 = '(こぼれ)'; $型 = 'Object[,]' }
    else { $答 = [string]$v; $型 = $v.GetType().Name }
  } catch { $答 = '★受け付けない★'; $型 = '-' }
  # ★『0』が 出た 時だけ 2つ目の 窓★（★見せかけの 0 と 本当の 0 を 分ける★）
  $窓２ = '-'
  if ($答 -eq '0' -and $型 -eq 'Double') { $窓２ = 窓２_本当にゼロか $ws $式 }
  [void]$結果.Add(("{0}`t{1}`t{2}`t{3}`t{4}`t{5}" -f $マス, $式, $答, $型, $窓２, $訳))
}

# ★材料は Excel から 読み返す★（私が 計算しない）
$材料の行 = New-Object System.Collections.ArrayList
foreach ($マス in @('A1','A2','A3','A4','A5','B1','B2','B3','B4','B5','C1','C2','C3','C4','C5','D1','D2','D3','D4','D5','E1','E2','E3','E4','E5','G1','G2')) {
  $mv = $ws.Range($マス).Value2
  if ($null -eq $mv) { [void]$材料の行.Add("#材料`t$マス`t`t空"); continue }
  if ($mv -is [bool]) { [void]$材料の行.Add(("#材料`t{0}`t{1}`t真偽" -f $マス, $(if ($mv) { 'TRUE' } else { 'FALSE' }))); continue }
  if ($mv -is [string]) { [void]$材料の行.Add(("#材料`t{0}`t{1}`t字" -f $マス, $mv)); continue }
  [void]$材料の行.Add(("#材料`t{0}`t{1}`t数" -f $マス, $mv.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture)))
}

$wb.Close($false); $xl.Quit()
foreach ($o in @($ws, $wb, $xl)) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null }
[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
# ★1点で 判じない★＝0.5秒ごとに 10秒まで 見る（0.8秒では まだ 消えて いない事が 在る）
$残り = @()
for ($t = 0; $t -lt 20; $t++) {
  Start-Sleep -Milliseconds 500
  $残り = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
  if ($残り.Count -eq 0) { break }
}
if ($残り.Count -eq 0) { Write-Host ('★Excel は 残って いません（0個・' + (($t + 1) * 0.5) + '秒で 消えた）★') }
else { Write-Host ('★★Excel が ' + $残り.Count + '個 残って います＝10秒 待っても 消えません＝手で 止めて ください★★') }

$頭 = @(
  '# ★INDEX と MATCH の 決まりを 実Excel に 聞いた★（2026-09-15）',
  '#   ★総当たりの 紙（golden-index-match2-2026-09-15.tsv）には 字の 表が 1つも 在りません★',
  '#   ＝★客の 使い方（字で 探して 隣を 取る）が 1本も 測れて いない★ ので 別に 取った。',
  "# ★どの Excel で 打ったか★ … 版 $版 ／ build $ビルド",
  '# ★★打った マスを 1本ずつ 書いて います★★（1列目）',
  '#   ＝実Excel は ★四角を 1つの 値に 詰める 時、式の 行／列と 交わる 所★を 採ります（暗黙の 交わり）',
  '#   ⇒★★突き合わせる 側も 同じ マスに 打って ください★★（H1 の 紙を A10 で 押すと 答えが 変わります）',
  '# ★前の 式の 跡を 消してから 打ち、CalculateFull してから 読んで います★',
  '#   ＝2026-09-15、消さずに 測る 道具が ★同じ 式に 2回で 違う 答え★を 返しました',
  '# ★★材料（機械が 読む）★★ … #材料<タブ>マス<タブ>値<タブ>型（★Excel から 読み返した 値★）',
  '# ★★2つ目の 窓★★ … 答えが 0 の 行だけ `=(式)=0` の 真偽も 取って 在ります（5列目）',
  '#   ＝`.Value2` は ★0 で ない 値に 0 を 返す★事が 在る（足し引きで 終わる 式）',
  '#   ★この 紙の 0 … `=INDEX(E1:E5,4)`＝空マスを 指した 時★'
)
$見出し = @("# 打ったマス`t式`t実Excel の 答え`t型`t窓②=0の真偽`t何を 測って いるか")
[System.IO.File]::WriteAllText($出, ((($頭 + $材料の行 + $見出し + $結果) -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding $false))
Write-Host "★書いた … $出★（式 $($式たち.Count) 本）"
