# toru-jitsu-excel-no-sel-shoshiki-wareki-keika.ps1
#   -- ★和暦（g/e）と 経過時間（[h]）を ★マスの 書式★ に した 時、実Excel が 画面に 出す 字（.Text）を 取る★（132）（2026-10-02）
#
#  ★★なぜ★★
#    Exally1 の 依頼（10-02）＝TEXT_era_wareki ／ TEXT_elapsed_hours を 作る 前の 物差し。
#    ★TEXT 関数の 答え★は tools/golden-excel.ps1 の 真値（cases/61-text-wareki-keika.json）で 取った。
#    ★画面の 道（マスの 書式）は TEXT と 同じとは 限らない★ ⇒ ここで 別に 取る（㋒）。
#
#  ★★司さんの 本は 使いません★★＝★新しい 空の 本★に 日付と 数を 置くだけ（★お客さんの 数字は 入りません★）
#    ⇒★出来た tsv は repo に 置いて よい★
#
#  ★★取り方★★
#    ・値は ★式（=DATE(...) ／ =0.5）★で 置く（★化けない 為★）
#    ・書式は ★.NumberFormatLocal★ に 渡す（★画面で 人が 打つ 字と 同じ★）
#      ＝★.NumberFormat（US）も 読み戻して 並べる★（★Excel が どう 受け取ったか★）
#    ・列の 幅は 60＝★`####` で 字が 隠れない 様に★（★`####` が 出たら 数える★）
#
#  ★★門★★
#    ①5.1（exit 8）／②走らせる 前の Excel が 0個（exit 3）★他の 席を 閉じません★
#    ③★頼んだ 数 ≠ 書いた 数 なら 赤（exit 6）★
#    ④★書式が 付かなかった 行が 在れば 赤（exit 9）★（2026-10-04）
#
#  ★★2026-10-04 追記 ── -幅 と -区切りも★★（Exally1 の 依頼・画面の「入らない ⇒ ####」と 区切り2つ以上の 書式）
#    -幅 <数> ... 列の 幅（既定 60＝★今までと 同じ★）／-幅 標準 ... ★列の 幅を 触らない★（その 本の 標準の 幅）
#    -区切りも ... ★区切りが 2つ以上 在る 書式★ と ★指数★ の 組を 足す（★渡さなければ 今までと 1行も 変わらない★）
#    ⇒★出しの 頭に ★実際の 列の 幅★ を 書く★（★渡した 数で なく Excel が 読み戻した 数★）
#    -指数も ... ★String に すると 指数の 形に なる 値★ × 和暦・経過の 書式（組 E・Exally1 の 依頼 10-04）
#    -字も   ... ★字の "1e-7"★（'1e-7 と ="1e-7"）× 数・和暦・経過の 書式（組 J）
#              ★字の マスでは 「本当に0か」の 窓は 意味を 持たない★（=(1e-7)=0 を 見て いる だけ）＝型の 窓（String）を 見る
#    -境目も ... ★日付の 上限 9999/12/31（通し 2958465）の 前後★ × 和暦・経過・yyyy/m/d（組 B）
#    -境目の他も ... 同じ 値 × ★経過で ない 時刻・曜日・月名・日だけ 等★（組 C・Exally1 の 案A の 前に 取る）
#
#  使い方:
#    powershell.exe -NoProfile -ExecutionPolicy Bypass -File docs/measured/toru-jitsu-excel-no-sel-shoshiki-wareki-keika.ps1 -出す先 <tsv> [-幅 60|標準|<数>] [-区切りも]

param([string]$出す先 = '', [string]$幅 = '60', [switch]$区切りも, [switch]$指数も, [switch]$字も, [switch]$境目も, [switch]$境目の他も)

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }
if (-not $出す先) { Write-Host '★★-出す先 を 渡して ください★★'; exit 4 }

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

# ★Exally1 の 依頼 そのまま★
$日付 = @('1900,1,1', '1912,7,29', '1912,7,30', '1926,12,24', '1926,12,25', '1989,1,7', '1989,1,8',
  '2019,4,30', '2019,5,1', '2020,1,1', '2026,1,31')
$和暦の書式 = @('g', 'gg', 'ggg', 'e', 'ee', 'ge.m.d', 'ggge年m月d日', 'gggee年', '[$-ja-JP]ggge年')
$値 = @('0', '0.5', '1.5', '46234', '1/24/60*59.5/60', '0.999999', '-1.5')
$経過の書式 = @('[h]:mm', '[hh]:mm', '[h]:mm:ss', '[m]:ss', '[mm]:ss', '[s]', '[h]', '[h]:mm:ss.00')

$頼 = New-Object System.Collections.ArrayList
foreach ($d in $日付) { foreach ($f in $和暦の書式) { [void]$頼.Add(@('W', ('=DATE(' + $d + ')'), $f)) } }
foreach ($v in $値) { foreach ($f in $経過の書式) { [void]$頼.Add(@('K', ('=' + $v), $f)) } }
if ($指数も) {
  # ★String に すると 指数に なる 数★（1e-7・-5.55e-17・1e+21・引き算の 端数 5.55e-17）
  foreach ($v in @('=1E-7', '=-5.55E-17', '=1E+21', '=0.1+0.2-0.3')) {
    foreach ($f in @('[h]:mm:ss', '[h]:mm:ss.00', '[s]', 'ge.m.d')) { [void]$頼.Add(@('E', $v, $f)) }
  }
}
if ($字も) {
  foreach ($v in @("'1e-7", '="1e-7"')) {
    foreach ($f in @('[s]', '[h]:mm:ss', 'ge.m.d', '0.00', 'General')) { [void]$頼.Add(@('J', $v, $f)) }
  }
}
if ($境目も) {
  foreach ($v in @('=2958465', '=2958465.99999', '=2958466', '=2958466.5')) {
    foreach ($f in @('ge.m.d', 'ggge年m月d日', '[h]:mm:ss', '[h]', '[s]', 'yyyy/m/d')) { [void]$頼.Add(@('B', $v, $f)) }
  }
}
if ($境目の他も) {
  foreach ($v in @('=2958465', '=2958465.99999', '=2958466', '=2958466.5')) {
    foreach ($f in @('h:mm', 'hh:mm:ss', 'h:mm AM/PM', 'm/d', 'd', 'yyyy', 'mmm', 'mmmm', 'aaa', 'aaaa', 'ddd', 'mm:ss', '[m]:ss', 'yyyy/m/d h:mm')) { [void]$頼.Add(@('C', $v, $f)) }
  }
}
if ($区切りも) {
  # ★区切りが 2つ以上★ ... 正・負・0・日付・字 の 5つの 値に 当てる
  $区切りの値 = @('=1.5', '=-1.5', '=0', '=DATE(2026,1,31)', '="abc"')
  $区切りの書式 = @('ge.m.d;@', '[h]:mm;[Red]-[h]:mm')
  foreach ($v in $区切りの値) { foreach ($f in $区切りの書式) { [void]$頼.Add(@('S', $v, $f)) } }
  # ★指数★（判じの 取り違えの 形）
  foreach ($v in @('=1234.5', '=0', '=-0.000123')) { [void]$頼.Add(@('X', $v, '0.00E+00')) }
}
Write-Host ('★頼んだ 数★ ' + $頼.Count)

$xl = New-Object -ComObject Excel.Application
$wb = $null; $ws = $null; $c = $null
$行 = New-Object System.Collections.ArrayList
$版xl = ''
$隠れ = 0
$付かず = 0
$誤ら = New-Object System.Collections.ArrayList
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $版xl = [string]$xl.Version + ' build ' + [string]$xl.Build
  $wb = $xl.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  # ★★列の 幅は ★全部 入れた 後★ に 決める★★（2026-10-04）
  #   ★なぜ★ ... 標準の 幅の 列に 日付・時刻の 書式の 値を 入れると ★Excel が 列を 勝手に 広げる★（8.44 → 15.19・実測）
  #            ⇒★先に 幅を 読むと 「標準 8.44」と 書いて 広がった 幅の 字を 取る★（10-04 に 1回 やった）
  $r = 1
  foreach ($q in $頼) {
    $c = $ws.Cells.Item($r, 1)
    $c.Formula = [string]$q[1]
    $誤 = ''
    try { $c.NumberFormatLocal = [string]$q[2] } catch { $誤 = 'NumberFormatLocal が 投げた: ' + $_.Exception.Message }
    # ★★2026-10-04 ── Local が 投げたら US の .NumberFormat で 付け直す★★
    #   ★なぜ★ ... 日本語の Excel は Local に `[Red]` を 受けない（Local の 名は `[赤]`）
    #            ⇒★投げた まま 書式が 付かず G/標準 の 字を 取って いた★（区切りの 5行・10-04 に 見つけた）
    #   ★付け直しても 付かなければ 「付かなかった」に 数える★＝★その 行は 物差しに 使えない★
    if ($誤 -ne '') {
      try { $c.NumberFormat = [string]$q[2]; $誤 = $誤 + ' ⇒ ★US の NumberFormat で 付け直した★' }
      catch {
        # ★この 環境では .NumberFormat も 日本語の 名で 受ける★（読み戻しが `G/標準`）⇒★色の 名を 日本語に 換えて Local で 付ける★
        $換 = ([string]$q[2]).Replace('[Red]', '[赤]')
        if ([string]$q[2] -eq 'General') { $換 = 'G/標準' }   # ★General の 日本語の 名★（10-04 に 付かなかった）
        if ($換 -ne [string]$q[2]) {
          try { $c.NumberFormatLocal = $換; $誤 = $誤 + ' ⇒ US でも 投げた ⇒ ★日本語の 名（' + $換 + '）に 換えて Local で 付けた★' }
          catch { $誤 = $誤 + ' ⇒ ★[赤] に 換えても 付かない★'; $付かず++ }
        } else { $誤 = $誤 + ' ⇒ ★US でも 付かない★'; $付かず++ }
      }
    }
    [void]$誤ら.Add($誤)
    $c = $null
    $r++
  }
  # ★★全部 入れた 後で 幅を 決める★★
  if ($幅 -eq '標準') { $ws.Columns.Item(1).ColumnWidth = [double]$ws.StandardWidth } else { $ws.Columns.Item(1).ColumnWidth = [double]$幅 }
  $実の幅 = [string]$ws.Columns.Item(1).ColumnWidth
  Write-Host ('★列の 幅（読む 直前に 読み戻した）★ ' + $実の幅)
  $r = 1
  foreach ($q in $頼) {
    $c = $ws.Cells.Item($r, 1)
    $誤 = [string]$誤ら[$r - 1]
    $字 = [string]$c.Text
    if ($字 -match '^#+$') { $隠れ++ }
    $us = [string]$c.NumberFormat
    $lo = [string]$c.NumberFormatLocal
    # ★★0 を 2つの 窓で 見る★★（2026-10-04・tests/monosashi-mado.test.mjs）
    #   ★2つ目の 窓★ ... 隣の マスに `=(式)=0` を 打ち ★本当に 0 か★ の 真偽を 取る
    #   ★3つ目の 窓★ ... .Value2 の 型（String / Double / 他）
    $v = $c.Value2
    $型 = '他'
    if ($v -is [string]) { $型 = 'String' } elseif ($v -is [double]) { $型 = 'Double' }
    $c2 = $ws.Cells.Item($r, 2)
    $式の中 = ([string]$q[1]).Substring(1)
    if (([string]$q[1]).StartsWith("'")) { $式の中 = '"' + $式の中 + '"' }
    $c2.Formula = ('=(' + $式の中 + ')=0')
    $真 = [string]$c2.Value2
    $c2 = $null
    [void]$行.Add(($q[0] + "`t" + $q[1] + "`t" + $q[2] + "`t" + $字 + "`t" + $us + "`t" + $lo + "`t" + $誤 + "`t" + $型 + "`t" + $真))
    $c = $null
    $r++
  }
  $wb.Close($false)
} finally {
  $c = $null; $ws = $null; $wb = $null
  $xl.Quit()
  $xl = $null
}

$書 = New-Object System.IO.StreamWriter($出す先, $false, (New-Object System.Text.UTF8Encoding($false)))
$書.NewLine = "`n"
$書.WriteLine('# 132 実Excel のマスの書式としての .Text（和暦・経過時間）／Excel ' + $版xl + '／列の幅 ' + $実の幅 + '（渡した -幅 ' + $幅 + '）／頼んだ ' + $頼.Count + '／書いた ' + $行.Count)
$書.WriteLine("組`t値の式`t渡した書式(Local)`t画面の字(.Text)`t読み戻し(.NumberFormat US)`t読み戻し(.NumberFormatLocal)`t誤り`t値の型(.Value2)`t本当に0か(=(式)=0)")
foreach ($l in $行) { $書.WriteLine($l) }
$書.Close()

Write-Host ('★Excel★ ' + $版xl)
Write-Host ('★頼んだ ' + $頼.Count + ' ／ 書いた ' + $行.Count + ' ／ `####` で 隠れた ' + $隠れ + '★')
Write-Host ('★書式が 付かなかった 行★ ' + $付かず + '（★0 で ないなら その 行は 物差しに 使えません★）')
if ($頼.Count -ne $行.Count) { Write-Host '★★頼んだ 数と 書いた 数が 違います★★'; exit 6 }
if ($付かず -ne 0) { exit 9 }
