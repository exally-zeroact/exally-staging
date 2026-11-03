# toru-nokori4-kimari.ps1 — ★IF／IFERROR／SUBTOTAL／TEXT の 決まりを 実Excel に 聞く★（2026-09-15）
#
#  ★★なぜ この 4つか★★
#    司さんの 実物 1冊に 出てくる 関数は ★8個だけ★。
#      INDEX 12,383 ／ MATCH 12,383 ／ ★IFERROR 4,329★ ／ ★IF 2,884★ ／
#      MAX 2,865 ／ ★TEXT 740★ ／ ★SUBTOTAL 496★ ／ SUM 238
#    ⇒ INDEX/MATCH/MAX/SUM は 済み。★残りは この 4つ★。
#
#  ★★総当たりでは なく「決まりを 名指し」★★
#    ・IF        … ★短絡するか★（偽の 枝の 誤りを 見るか）／字・空・0 の 扱い／引数 2つ
#    ・IFERROR   … 何を 誤りと 見るか（空の 字は？）／2つ目が 空の 時
#    ・SUBTOTAL  … 11個の 番号 ＋ 101番台 ／ ★入れ子は 数えないか★（これが 本体）
#    ・TEXT      … 丸め・桁区切り・日付・和暦・曜日・％・負の 書き分け
#
#  ★★守って いる 事★★（今日 踏んだ 穴）
#    ①★前の 式の 跡を 消してから 打つ★／②★CalculateFull してから 読む★
#    ③★打った マスを 紙に 書く★（暗黙の 交わり）／④★『0』は 2つ目の 窓でも 見る★
#    ⑤★終わりに Excel が 残って いないか 数える★
#
#  使い方: powershell -NoProfile -File docs/measured/kansuu46/toru-nokori4-kimari.ps1
$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-nokori4-kimari-2026-09-15.tsv'

$誤りの番号 = @{
  -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!';
  -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!';
  -2146826246 = '#N/A';   -2146826243 = '#SPILL!'; -2146826238 = '#CALC!'; -2146826237 = '#BUSY!'
}

# ══ ★2つ目の 窓★ ══（`.Value2` は 0 で ない 値に 0 を 返す 事が 在る）
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
#   A1:A5 … 1..5      B1:B5 … 2,4,6,8,10
#   ★A6 は 式★＝`=SUBTOTAL(9,A1:A5)`（★入れ子を 数えないか★を 見る 為）
#   E     … 並んで いない＋★E4 は 空★    G1/G2 … TRUE/FALSE
#   D1:D5 … 字
$数の材料 = @(
  @('A1',1), @('A2',2), @('A3',3), @('A4',4), @('A5',5),
  @('B1',2), @('B2',4), @('B3',6), @('B4',8), @('B5',10),
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
$ws.Range('A6').Formula = '=SUBTOTAL(9,A1:A5)'

$式たち = @(
  # ── IF ──────────────────────────────
  @('J1', '=IF(TRUE,1,2)',            'IF … 真'),
  @('J1', '=IF(FALSE,1,2)',           'IF … 偽'),
  @('J1', '=IF(1,1,2)',               '数の 1 は 真か'),
  @('J1', '=IF(0,1,2)',               '数の 0 は 偽か'),
  @('J1', '=IF(-1,1,2)',              '負も 真か'),
  @('J1', '=IF(0.5,1,2)',             '小数も 真か'),
  @('J1', '=IF("TRUE",1,2)',          '★字の TRUE★ は 真か 誤りか'),
  @('J1', '=IF("",1,2)',              '空の 字'),
  @('J1', '=IF("あ",1,2)',            '字'),
  @('J1', '=IF(A1,1,2)',              'マス（A1=1）'),
  @('J1', '=IF(A9,1,2)',              '★空の マス★'),
  @('J1', '=IF(G1,1,2)',              'マスの TRUE'),
  @('J1', '=IF(TRUE,1)',              '★引数 2つ★（真）'),
  @('J1', '=IF(FALSE,1)',             '★引数 2つ★（偽）＝何を 返すか'),
  @('J1', '=IF(TRUE,,2)',             '★真の 枝を 省く★'),
  @('J1', '=IF(FALSE,1,)',            '★偽の 枝を 省く★'),
  @('J1', '=IF(1/0,1,2)',             '条件が 誤り'),
  @('J1', '=IF(TRUE,1/0,2)',          '選ぶ 枝が 誤り'),
  @('J1', '=IF(FALSE,1/0,2)',         '★選ばない 枝の 誤りを 見るか★（短絡するか）'),
  @('J1', '=IF(TRUE,2,1/0)',          '★同上（逆）★'),
  @('J1', '=IF(TRUE,"あ","い")',      '字を 返す'),
  @('J1', '=IF(A1>2,1,2)',            'くらべた 答えを 条件に'),
  @('J1', '=IF(A1:A5>2,1,2)',         '★四角を 条件に★（暗黙の 交わり／溢れ）'),
  @('J1', '=IF(TRUE,A1:A5,2)',        '★四角を 枝に★'),
  # ── IFERROR ─────────────────────────
  @('J1', '=IFERROR(1,"x")',          'IFERROR … 誤りで ない'),
  @('J1', '=IFERROR(1/0,"x")',        '#DIV/0! を 受ける'),
  @('J1', '=IFERROR(NA(),"x")',       '#N/A を 受ける'),
  @('J1', '=IFERROR(MATCH(99,A1:A5,0),"なし")', '★客の 使い方★'),
  @('J1', '=IFERROR("","x")',         '★空の 字は 誤りか★'),
  @('J1', '=IFERROR(A9,"x")',         '★空の マス★'),
  @('J1', '=IFERROR(1/0,)',           '★2つ目を 省く★'),
  @('J1', '=IFERROR(1/0,1/0)',        '★受け手も 誤り★'),
  @('J1', '=IFERROR(1/0,"")',         '2つ目が 空の 字'),
  @('J1', '=IFERROR(A1:A5,"x")',      '★四角★（暗黙の 交わり／溢れ）'),
  @('J1', '=IFERROR(1)',              '★引数 1つ★'),
  @('J1', '=IFERROR(TRUE,"x")',       '真偽'),
  # ── SUBTOTAL ────────────────────────
  @('J1', '=SUBTOTAL(1,A1:A5)',       'SUBTOTAL 1 … AVERAGE'),
  @('J1', '=SUBTOTAL(2,A1:A5)',       '2 … COUNT'),
  @('J1', '=SUBTOTAL(3,E1:E5)',       '3 … COUNTA（E4 は 空）'),
  @('J1', '=SUBTOTAL(4,A1:A5)',       '4 … MAX'),
  @('J1', '=SUBTOTAL(5,A1:A5)',       '5 … MIN'),
  @('J1', '=SUBTOTAL(6,A1:A5)',       '6 … PRODUCT'),
  @('J1', '=SUBTOTAL(7,A1:A5)',       '7 … STDEV'),
  @('J1', '=SUBTOTAL(8,A1:A5)',       '8 … STDEVP'),
  @('J1', '=SUBTOTAL(9,A1:A5)',       '9 … SUM'),
  @('J1', '=SUBTOTAL(10,A1:A5)',      '10 … VAR'),
  @('J1', '=SUBTOTAL(11,A1:A5)',      '11 … VARP'),
  @('J1', '=SUBTOTAL(101,A1:A5)',     '101 … AVERAGE（隠した 行を 抜く）'),
  @('J1', '=SUBTOTAL(103,E1:E5)',     '103 … COUNTA'),
  @('J1', '=SUBTOTAL(109,A1:A5)',     '109 … SUM'),
  @('J1', '=SUBTOTAL(111,A1:A5)',     '111 … VARP'),
  @('J1', '=SUBTOTAL(9,A1:A6)',       '★★入れ子を 数えないか★★（A6 は =SUBTOTAL(9,A1:A5)）'),
  @('J1', '=SUBTOTAL(2,A1:A6)',       '★同上（個数）★'),
  @('J1', '=SUM(A1:A6)',              '★SUM は 入れ子を 数えるか★（くらべる 相手）'),
  @('J1', '=SUBTOTAL(9,A1:A5,B1:B5)', '範囲 2つ'),
  @('J1', '=SUBTOTAL(9,E1:E5)',       '空マスを 跨ぐ'),
  @('J1', '=SUBTOTAL(9,D1:D5)',       '★字ばかりの 列★'),
  @('J1', '=SUBTOTAL(0,A1:A5)',       '★番号 0★'),
  @('J1', '=SUBTOTAL(12,A1:A5)',      '★番号 12（外）★'),
  @('J1', '=SUBTOTAL(9.5,A1:A5)',     '★番号が 小数★'),
  @('J1', '=SUBTOTAL(109.5,A1:A5)',   '★同上（101番台）★'),
  @('J1', '=SUBTOTAL("9",A1:A5)',     '★番号が 字★'),
  @('J1', '=SUBTOTAL(9,5)',           '★範囲で なく 数★'),
  @('J1', '=SUBTOTAL(9)',             '★引数 1つ★'),
  # ── TEXT ────────────────────────────
  @('J1', '=TEXT(1234.5,"0")',        'TEXT … 丸め（0）'),
  @('J1', '=TEXT(1234.5,"0.00")',     '小数 2桁'),
  @('J1', '=TEXT(1234.5,"#,##0")',    '桁区切り'),
  @('J1', '=TEXT(1234.5,"#,##0.0")',  '桁区切り＋小数'),
  @('J1', '=TEXT(1234.5,"#,##0円")',  '★字を 混ぜる★'),
  @('J1', '=TEXT(1234,"0000000")',    '0 で 埋める'),
  @('J1', '=TEXT(1234,"#")',          '# だけ'),
  @('J1', '=TEXT(0,"#")',             '★0 を # で★'),
  @('J1', '=TEXT(0,"0")',             '0 を 0 で'),
  @('J1', '=TEXT(0.5,"0%")',          'パーセント'),
  @('J1', '=TEXT(0.1234,"0.0%")',     'パーセント＋小数'),
  @('J1', '=TEXT(-1234,"#,##0")',     '負'),
  @('J1', '=TEXT(-1234,"#,##0;(#,##0)")', '★負の 書き分け★'),
  @('J1', '=TEXT(0,"#,##0;(#,##0);-")',   '★0 の 書き分け★'),
  @('J1', '=TEXT(1234,"0;;")',        '★0 を 消す 形★'),
  @('J1', '=TEXT(45292,"yyyy/m/d")',  '日付'),
  @('J1', '=TEXT(45292,"yyyy-mm-dd")', '日付（0埋め）'),
  @('J1', '=TEXT(45292,"m/d")',       '月日'),
  @('J1', '=TEXT(45292,"yyyy年m月d日")', '日付（字を 混ぜる）'),
  @('J1', '=TEXT(45292,"aaa")',       '★曜日（短）★'),
  @('J1', '=TEXT(45292,"aaaa")',      '★曜日（長）★'),
  @('J1', '=TEXT(45292,"ggge年m月d日")', '★和暦★'),
  @('J1', '=TEXT(45292,"ge.m.d")',    '★和暦（短）★'),
  @('J1', '=TEXT(0.520833333333333,"h:mm")', '時刻'),
  @('J1', '=TEXT(0.520833333333333,"h:mm:ss")', '時刻（秒）'),
  @('J1', '=TEXT(1.5,"[h]:mm")',      '★24時間を 超える★'),
  @('J1', '=TEXT("あ","0")',          '★字を 渡す★'),
  @('J1', '=TEXT("123","0")',         '★数に 見える 字★'),
  @('J1', '=TEXT(TRUE,"0")',          '真偽'),
  @('J1', '=TEXT(1,"@")',             '★@（字として）★'),
  @('J1', '=TEXT(1234.5,"")',         '★書式が 空★'),
  @('J1', '=TEXT(A9,"0")',            '★空の マス★'),
  @('J1', '=TEXT(1/0,"0")',           '誤りを 渡す'),
  @('J1', '=TEXT(1234.5,"General")',  '★General★'),
  @('J1', '=TEXT(1.5,"# ?/?")',       '★分数★')
)

$結果 = New-Object System.Collections.ArrayList
$i = 0
foreach ($s in $式たち) {
  $i++
  $マス = $s[0]; $式 = $s[1]; $訳 = $s[2]
  Write-Host ("  … {0}/{1}  {2}" -f $i, $式たち.Count, $マス)
  $ws.Range('I1:Z30').Clear() | Out-Null
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

# ★材料は Excel から 読み返す★（式の マスは ★式のまま★ 書く）
$材料の行 = New-Object System.Collections.ArrayList
foreach ($マス in @('A1','A2','A3','A4','A5','A6','B1','B2','B3','B4','B5','D1','D2','D3','D4','D5','E1','E2','E3','E4','E5','G1','G2')) {
  $f = [string]$ws.Range($マス).Formula
  if ($f.StartsWith('=')) { [void]$材料の行.Add(("#材料`t{0}`t{1}`t式" -f $マス, $f)); continue }
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
else { Write-Host ('★★Excel が ' + $残り.Count + '個 残って います＝10秒 待っても 消えません＝手で 止めて ください★★') }

$頭 = @(
  '# ★IF／IFERROR／SUBTOTAL／TEXT の 決まりを 実Excel に 聞いた★（2026-09-15）',
  '#   ★司さんの 実物 1冊に 出てくる 8個の うち 残り 4個★',
  '#   （IFERROR 4,329回 ／ IF 2,884回 ／ TEXT 740回 ／ SUBTOTAL 496回）',
  "# ★どの Excel で 打ったか★ … 版 $版 ／ build $ビルド",
  '# ★★打った マスを 1本ずつ 書いて います★★（1列目）',
  '#   ＝実Excel は ★四角を 1つの 値に 詰める 時、式の 行／列と 交わる 所★を 採ります',
  '#   ⇒★★突き合わせる 側も 同じ マスに 打って ください★★',
  '# ★前の 式の 跡を 消してから 打ち、CalculateFull してから 読んで います★',
  '# ★★材料（機械が 読む）★★ … #材料<タブ>マス<タブ>値<タブ>型（型＝数／字／真偽／式／空）',
  '#   ★A6 は 式★＝`=SUBTOTAL(9,A1:A5)`（★入れ子を 数えないか★を 見る 為に 置いて 在る）',
  '# ★★2つ目の 窓★★ … 答えが 0 の 行だけ `=(式)=0` の 真偽も 取って 在ります（5列目）'
)
$見出し = @("# 打ったマス`t式`t実Excel の 答え`t型`t窓②=0の真偽`t何を 測って いるか")
[System.IO.File]::WriteAllText($出, ((($頭 + $材料の行 + $見出し + $結果) -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding $false))
Write-Host "★書いた … $出★（式 $($式たち.Count) 本）"
