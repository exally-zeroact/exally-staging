# toru-tana-12-13.ps1 — ★棚 ⑫⑬ を 埋める★（2026-09-15）
#
#  ★★なぜ★★
#    2026-09-15 に INDEX/MATCH を 書きました。★その 時に 測って いない 所が 残って います★。
#    ★書いた 後でも 測る★＝★「たぶん こう」で 置いた 所を 紙に 変える★。
#
#  ★埋める 物★
#    ⑫ MATCH の 型 -1 … 今は ★実測 3本に 合わせた 模型★（実Excel の 中身では ない）
#         ⇒★ちゃんと 降順の 列で 何通りか★／★崩れた 列で 何通りか★ 測って 確かめる
#    ⑬ 5つ … ①2次元の 並びを MATCH に ②並びの 後ろが 空 ③ワイルドカードの `~`
#             ④番地の 無い 表を INDEX で まるごと ⑤横1本の 表に 番号 1つ
#
#  ★守って いる 事★ 跡を 消す／CalculateFull／打ったマス／2つ目の 窓／Excel の 残りを 数える
$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-tana-12-13-2026-09-15.tsv'

$誤りの番号 = @{
  -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!';
  -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!';
  -2146826246 = '#N/A';   -2146826243 = '#SPILL!'; -2146826238 = '#CALC!'; -2146826237 = '#BUSY!'
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

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
$版 = $xl.Version; $ビルド = $xl.Build
$wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)

# ★材料★
#   A1:A5 … 1..5（昇順）        B1:B5 … 2,4,6,8,10
#   C1:C5 … 10,8,6,4,2（★ちゃんと 降順★）
#   E1:E3 … 1,3,5 ／ E4,E5 空（★後ろが 空の 列★）
#   F1:F5 … 5,1,4,2,3（★崩れた 並び★）
#   D1:D3 … a*b ／ p?q ／ plain（★型紙の 字そのもの★）
#   H1:J1 … 7,8,9（★横1本の 表★）
$数の材料 = @(
  @('A1',1), @('A2',2), @('A3',3), @('A4',4), @('A5',5),
  @('B1',2), @('B2',4), @('B3',6), @('B4',8), @('B5',10),
  @('C1',10), @('C2',8), @('C3',6), @('C4',4), @('C5',2),
  @('E1',1), @('E2',3), @('E3',5),
  @('F1',5), @('F2',1), @('F3',4), @('F4',2), @('F5',3),
  @('H1',7), @('I1',8), @('J1',9)
)
foreach ($z in $数の材料) { $ws.Range($z[0]).Value2 = [double]$z[1] }
$ws.Range('D1').Value2 = 'a*b'
$ws.Range('D2').Value2 = 'p?q'
$ws.Range('D3').Value2 = 'plain'

$式たち = @(
  # ── ⑫ 型 -1（ちゃんと 降順 C＝10,8,6,4,2）────────────
  @('N1', '=MATCH(6,C1:C5,-1)',   '⑫ 降順・ぴたり 在り'),
  @('N1', '=MATCH(7,C1:C5,-1)',   '⑫ 降順・間（以上で 一番 小さいのは 8＝2番目）'),
  @('N1', '=MATCH(10,C1:C5,-1)',  '⑫ 降順・先頭'),
  @('N1', '=MATCH(2,C1:C5,-1)',   '⑫ 降順・最後'),
  @('N1', '=MATCH(1,C1:C5,-1)',   '⑫ 降順・全部より 小さい'),
  @('N1', '=MATCH(11,C1:C5,-1)',  '⑫ 降順・全部より 大きい'),
  # ── ⑫ 型 -1（崩れた 並び F＝5,1,4,2,3）────────────
  @('N1', '=MATCH(4,F1:F5,-1)',   '⑫ 崩れた 並び・4'),
  @('N1', '=MATCH(3,F1:F5,-1)',   '⑫ 崩れた 並び・3'),
  @('N1', '=MATCH(5,F1:F5,-1)',   '⑫ 崩れた 並び・5'),
  @('N1', '=MATCH(1,F1:F5,-1)',   '⑫ 崩れた 並び・1'),
  @('N1', '=MATCH(2,F1:F5,1)',    '⑫ 崩れた 並び・型 1（くらべる 相手）'),
  @('N1', '=MATCH(4,F1:F5,1)',    '⑫ 崩れた 並び・型 1'),
  # ── ⑬① 2次元の 並び ───────────────────────
  @('N1', '=MATCH(3,A1:B5,0)',    '⑬① 2次元の 並び・ぴたり'),
  @('N1', '=MATCH(3,A1:B5,1)',    '⑬① 2次元の 並び・型 1'),
  @('N1', '=MATCH(99,A1:B5,0)',   '⑬① 2次元の 並び・無い'),
  # ── ⑬② 並びの 後ろが 空（E＝1,3,5,空,空）──────────
  @('N1', '=MATCH(9,E1:E5,1)',    '⑬② 後ろが 空・全部より 大きい'),
  @('N1', '=MATCH(3,E1:E5,1)',    '⑬② 後ろが 空・ぴたり'),
  @('N1', '=MATCH(4,E1:E5,1)',    '⑬② 後ろが 空・間'),
  @('N1', '=MATCH(0,E1:E5,1)',    '⑬② 後ろが 空・全部より 小さい'),
  @('N1', '=MATCH(5,E1:E5,1)',    '⑬② 後ろが 空・最後の 数'),
  # ── ⑬③ ワイルドカードの ~ ───────────────────
  @('N1', '=MATCH("a*b",D1:D3,0)',   '⑬③ ＊を そのまま 探す（型紙として 見ると 全部 当たる）'),
  @('N1', '=MATCH("a~*b",D1:D3,0)',  '⑬③ ~＊＝★の 字そのもの'),
  @('N1', '=MATCH("p~?q",D1:D3,0)',  '⑬③ ~？'),
  @('N1', '=MATCH("p?q",D1:D3,0)',   '⑬③ ？を 型紙として'),
  @('N1', '=MATCH("~*",D1:D3,0)',    '⑬③ ~＊ だけ'),
  # ── ⑬④ 番地の 無い 表を INDEX で まるごと ──────────
  @('N1', '=INDEX({1;2;3},0)',       '⑬④ 直に 書いた 表・行 0'),
  @('N1', '=SUM(INDEX({1;2;3},0))',  '⑬④ 同上を 足す'),
  @('N1', '=INDEX({1;2;3},2)',       '⑬④ 直に 書いた 表・2番目'),
  @('N1', '=INDEX({1,2;3,4},0,2)',   '⑬④ 直に 書いた 2次元・列 まるごと'),
  @('N1', '=SUM(INDEX({1,2;3,4},0,2))', '⑬④ 同上を 足す'),
  # ── ⑬⑤ 横1本の 表に 番号 1つ（H1:J1＝7,8,9）─────────
  @('N1', '=INDEX(H1:J1,2)',      '⑬⑤ 横1本・番号 1つ'),
  @('N1', '=INDEX(H1:J1,1,2)',    '⑬⑤ 横1本・行と 列'),
  @('N1', '=INDEX(H1:J1,2,1)',    '⑬⑤ 横1本・行を 2 に すると'),
  @('N1', '=INDEX(H1:J1,0,2)',    '⑬⑤ 横1本・行 0'),
  @('N1', '=MATCH(8,H1:J1,0)',    '⑬⑤ 横1本を MATCH で')
)

$結果 = New-Object System.Collections.ArrayList
$i = 0
foreach ($s in $式たち) {
  $i++
  $マス = $s[0]; $式 = $s[1]; $訳 = $s[2]
  Write-Host ("  … {0}/{1}" -f $i, $式たち.Count)
  $ws.Range('M1:Z30').Clear() | Out-Null
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
  $窓２ = '-'
  if ($答 -eq '0' -and $型 -eq 'Double') { $窓２ = 窓２_本当にゼロか $ws $式 }
  [void]$結果.Add(("{0}`t{1}`t{2}`t{3}`t{4}`t{5}" -f $マス, $式, $答, $型, $窓２, $訳))
}

$材料の行 = New-Object System.Collections.ArrayList
foreach ($マス in @('A1','A2','A3','A4','A5','B1','B2','B3','B4','B5','C1','C2','C3','C4','C5','D1','D2','D3','E1','E2','E3','E4','E5','F1','F2','F3','F4','F5','H1','I1','J1')) {
  $mv = $ws.Range($マス).Value2
  if ($null -eq $mv) { [void]$材料の行.Add("#材料`t$マス`t`t空"); continue }
  if ($mv -is [bool]) { [void]$材料の行.Add(("#材料`t{0}`t{1}`t真偽" -f $マス, $(if ($mv) { 'TRUE' } else { 'FALSE' }))); continue }
  if ($mv -is [string]) { [void]$材料の行.Add(("#材料`t{0}`t{1}`t字" -f $マス, $mv)); continue }
  [void]$材料の行.Add(("#材料`t{0}`t{1}`t数" -f $マス, $mv.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture)))
}

$wb.Close($false); $xl.Quit()
foreach ($o in @($ws, $wb, $xl)) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null }
[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
$残り = @()
for ($t = 0; $t -lt 20; $t++) {
  Start-Sleep -Milliseconds 500
  $残り = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
  if ($残り.Count -eq 0) { break }
}
if ($残り.Count -eq 0) { Write-Host ('★Excel は 残って いません（0個・' + (($t + 1) * 0.5) + '秒で 消えた）★') }
else { Write-Host ('★★Excel が ' + $残り.Count + '個 残って います＝手で 止めて ください★★') }

$頭 = @(
  '# ★棚 ⑫⑬ を 埋めた★（2026-09-15）＝★書いた 後でも 測る★',
  '#   ⑫ MATCH の 型 -1（今は 実測 3本に 合わせた 模型）',
  '#   ⑬ ①2次元の 並び ②後ろが 空 ③ワイルドカードの ~ ④番地の 無い 表 ⑤横1本の 表',
  "# ★どの Excel で 打ったか★ … 版 $版 ／ build $ビルド",
  '# ★★打った マスを 1本ずつ 書いて います★★（1列目・全部 N1）',
  '# ★前の 式の 跡を 消してから 打ち、CalculateFull してから 読んで います★',
  '# ★★材料（機械が 読む）★★ … #材料<タブ>マス<タブ>値<タブ>型',
  '#   A＝昇順 1..5 ／ B＝2,4,6,8,10 ／ ★C＝10,8,6,4,2（ちゃんと 降順）★',
  '#   ★E＝1,3,5,空,空（後ろが 空）★ ／ ★F＝5,1,4,2,3（崩れた 並び）★',
  '#   D＝a*b／p?q／plain（★型紙の 字そのもの★） ／ ★H1:J1＝7,8,9（横1本）★',
  '# ★★2つ目の 窓★★ … 答えが 0 の 行だけ `=(式)=0` の 真偽も 取って 在ります（5列目）'
)
$見出し = @("# 打ったマス`t式`t実Excel の 答え`t型`t窓②=0の真偽`t何を 測って いるか")
[System.IO.File]::WriteAllText($出, ((($頭 + $材料の行 + $見出し + $結果) -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding $false))
Write-Host "★書いた … $出★（式 $($式たち.Count) 本）"
