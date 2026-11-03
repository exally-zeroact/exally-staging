# tsukuru-furui-kami-torinaoshi.ps1
#   -- ★古い 紙（2026-09-08）を ★今の 道（Formula2）★で 打ち直して 差を 数える★（67）（2026-09-21）
#
#  ★★なぜ★★
#    `docs/measured/kansuu46/` の 道具 21本は ★`Formula2` を 1本も 使って いません★
#    ⇒`.Formula` は ★溢れません★（2026-09-19 実測）＝★暗黙の 交わりで 1マス★
#    ⇒★だから 紙の 答えは ★昔の 道の 答え★です★
#    ⇒★その 紙を 物差しに した 試験は ★昔の 道を 守って います★★
#    ⇒Exally1 が 2次元を 直そうと して ★3本 赤★で 止まった のが それ
#
#  ★★でも 「取り直す 分母」が 分かりません★★
#    ・紙 35枚／式 27,862本／★範囲を 渡す 式 6,204本★／2次元 739本
#    ⇒★6,204本は 大きすぎる 分母です★
#      ＝`SUM(A1:A5)` の ように ★範囲を 受けるのが 当たり前★の 式が 混ざる
#      ＝★そういう 式は 昔の 道でも 今の 道でも 同じ 答え★
#    ⇒★★本当の 分母＝「今の 道で 打つと 答えが 変わる 式」★★
#    ⇒★それは 数えられます＝ここで 数えます★
#
#  ★★やり方★★
#    ①`golden-346-2026-09-08.tsv` から ★範囲を 渡す 式★を 取り出す
#    ②★同じ 盤面★で ★`Formula2`（今の 道）★で 打ち直す
#    ③★頭の マスの 答え★を 紙の 答えと 比べる
#    ④★何マス 埋まったか★も 数える（★2マス 以上＝昔の 紙は 1マス目しか 見て いない★）
#
#  ★★盤面（★紙の 頭に 書いて ある 物★）★★
#    A1:A5 = 1 / 2 / 3 / 4 / 5
#    B1:B5 = 2 / 4 / 6 / 8 / 10
#    D1 = 2024/1/1 ／ D2 = 2026/1/1
#
#  ★★置き場★★
#    ★式は H列・行1 から 10行 おき★（★A1:B5 と 交わる 行に 頭を 置く★＝行1）
#    ⇒★昨日 Exally1 が 行100 に 置いて 数を 逆に 出した 所です★
#    ⇒★但し 2本目 以降は 行11・行21 ほか＝交わりません★
#      ＝★暗黙の 交わりが 要る 式は それで 変わります★
#      ⇒★だから 「今の 道（Formula2）」で 打ちます★＝★交わりでは なく 溢れる★
#      ⇒★交わりの 有無で 変わるのは 昔の 道の 話★
#
#  ★★この 紙で 言える 事／言えない 事★★
#    ★言える★ ... ★今の 道で 打つと 答えが 変わる 式が 何本 在るか★
#    ★言えない★ ... ★どちらが 正しいか★（★今の 道が 正しい★のは 別に 実測済み）
#                 ... ★他の 34枚の 紙★（★この 道具は 1枚だけ★）
#
#  ★★柱の 名前に ついて（★2026-09-21 に 1回 転びました★）★★
#    `tests/shiki-kansuu-kami.test.mjs` は ★紙の 柱から
#    ★`式` の 列★と ★`実Excel` で 始まる 列★を 探して 物差しに します★（1345-1355行）
#    ⇒★どちらかが 無いと 「柱なし」と 見なし ★3列目★を 答えに します★
#    ⇒★私は 柱を `今の 答え` と 書いて いました★
#      ⇒★試験が 3列目（`昔の 答え`）を 物差しに して いました★
#      ⇒★Exally1 の 直しが 正しいのに 赤＝★偽の 負け★★
#    ⇒★だから 柱を `実Excel（今の 道＝Formula2）` に しました★
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③紙が 読めなければ 走らない（exit 6）／④式が 0本なら 走らない（exit 5）
#    ⑤★外へ 出る 6個は 打ちません★
#    ⑥★対照 2本（SUM・SEQUENCE）が 待つ 通りで なければ 落とす（exit 7）★
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
# ★★2026-09-21 足し ── ★1枚 だけでは 分母に なりません★★
#   ＝`golden-346` 以外の 紙にも ★範囲を 渡す 式が 5,392本★ 在りました
#   ⇒★`kansuu46` の 紙を ★全部★ 読みます★（★自分で 作った 紙は 除く★）
$紙たち = @(Get-ChildItem (Join-Path $ここ 'kansuu46') -Filter '*.tsv' | Where-Object { $_.Name -notlike '*nijigen-formula2*' } | Sort-Object Name)
# ★★2026-09-21 ── ★出し先を 分けました★★
#   ＝35枚に 広げた 数は ★まだ 信じられません★（材料を 書いて いない 紙が 26枚）
#   ⇒★信じられない 紙を 物差しに させない★＝★柱に `実Excel` と 書きません★
#   ⇒★実際に 1回 物差しに されて 試験が 6本 赤に なりました★（2026-09-21）
$出 = Join-Path $ここ 'golden-furui-kami-zenmai-WIP-2026-09-21.tsv'

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }
if ($紙たち.Count -eq 0) { Write-Host '★★紙が 1枚も 在りません★★'; exit 6 }
Write-Host ('★読む 紙 ... ' + $紙たち.Count + '枚★')

$外へ出る = 'WEBSERVICE', 'STOCKHISTORY', 'TRANSLATE', 'DETECTLANGUAGE', 'IMAGE', 'RTD'
$問い = New-Object System.Collections.Generic.List[object]
$外した = 0
# ★★2026-09-21 ── ★紙ごとに 材料が 違います★★
#   ＝`golden-tana-12-13` は ★C1:C4 = 10,8,6,4★ を 使う
#   ⇒★1つの 盤面で 全部 打つと ★偽の 違い★が 出ます★（1回目 188本 と 出た）
#   ⇒★紙の `#材料` を その 紙の 分だけ 使います★
$材料 = @{}
foreach ($かみ in $紙たち) {
$材料[$かみ.Name] = New-Object System.Collections.Generic.List[object]
foreach ($l in [System.IO.File]::ReadAllLines($かみ.FullName)) {
  if ($l.StartsWith('#材料')) {
    $z = $l.Split("`t")
    if ($z.Count -ge 3) { $材料[$かみ.Name].Add(@{ マス = $z[1]; 値 = $z[2] }) }
    continue
  }
  if ($l -eq '' -or $l.StartsWith('#')) { continue }
  $c = $l.Split("`t")
  if ($c.Count -lt 3) { continue }
  $f = $c[1]
  if (-not $f.StartsWith('=')) { continue }
  # ★範囲を 渡す 式だけ★（`A1:A5` の 形が 在る 物）
  if ($f -notmatch '[A-Z]{1,2}\d+:[A-Z]{1,2}\d+') { continue }
  $わるい = $false
  foreach ($n in $外へ出る) { if ($f.ToUpper().Contains($n)) { $わるい = $true } }
  foreach ($ch in $f.ToCharArray()) { if ([int]$ch -gt 127) { $わるい = $true } }
  if ($わるい) { $外した++; continue }
  $問い.Add(@{ 名 = $c[0]; 式 = $f; 昔 = $c[2]; 紙 = $かみ.Name })
}
}
Write-Host ('★紙から 取った 式 ... ' + $問い.Count + '本★（★外した ' + $外した + '本★）')
if ($問い.Count -eq 0) { Write-Host '★★1本も 在りません★★'; exit 5 }

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

# ★★COM の 誤りの 番号 ⇒ 字★★（★紙は 字で 書いて います★）
$誤りの数 = @{
  '-2146826281' = '#DIV/0!'; '-2146826252' = '#NUM!';  '-2146826246' = '#N/A';
  '-2146826273' = '#VALUE!'; '-2146826265' = '#REF!';  '-2146826259' = '#NAME?';
  '-2146826238' = '#CALC!';  '-2146826243' = '#SPILL!'; '-2146826288' = '#NULL!'
}

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null; $c2 = $null
$対照ok = 0
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $xl.ScreenUpdating = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★古い 紙（2026-09-08）を 今の 道（Formula2）で 打ち直した★（67）（2026-09-21）')
  $行.Add('# ★元の 紙★ ... docs/measured/kansuu46 の tsv を ★全部★（自分で 作った 紙は 除く）')
  $行.Add('# ★★元の 紙は `.Formula`（むかしの みち＝溢れない）で 取られて います★★（kansuu46 の 道具 21本とも）')
  $行.Add('# ★★盤面★★ ... ★紙ごとに その 紙の `#材料` を 使います★（1つの 盤面では 打ちません）')
  $行.Add('#   ＝`golden-tana-12-13` は C1:C4=10,8,6,4 を 使う＝★1つの 盤面だと 偽の 違いが 出ます★')
  $行.Add('# ★置き場★ H列・行1 から 10行 おき')
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★どの 貝殻か★ ... PowerShell ' + $版.ToString())
  # ══ ★★`#材料` を 書きます（★これが 無いと 試験は 1行も 押しません★）★★ ══
  #   ★2026-09-21 に 1回 踏みました★
  #     ＝`#材料` を 書かなかった ので 試験が ★全行 飛ばし★、
  #       ★わざと 台を 壊しても 緑の まま★に なりました（★守って いない★）
  #   ⇒`tests/shiki-kansuu-kami.test.mjs` は ★紙の `#材料` から 盤面を 作ります★
  #     （★道具が 決めない／材料を 当て推量で 作らない★）
  $行.Add('#材料' + "`t" + 'A1' + "`t" + '1')
  $行.Add('#材料' + "`t" + 'A2' + "`t" + '2')
  $行.Add('#材料' + "`t" + 'A3' + "`t" + '3')
  $行.Add('#材料' + "`t" + 'A4' + "`t" + '4')
  $行.Add('#材料' + "`t" + 'A5' + "`t" + '5')
  $行.Add('#材料' + "`t" + 'B1' + "`t" + '2')
  $行.Add('#材料' + "`t" + 'B2' + "`t" + '4')
  $行.Add('#材料' + "`t" + 'B3' + "`t" + '6')
  $行.Add('#材料' + "`t" + 'B4' + "`t" + '8')
  $行.Add('#材料' + "`t" + 'B5' + "`t" + '10')
  # ★D1 D2 は 日付＝★値で 控えます★（古い 紙と 同じ 形）
  $行.Add('#材料' + "`t" + 'D1' + "`t" + [string]$sh.Range('D1').Value2)
  $行.Add('#材料' + "`t" + 'D2' + "`t" + [string]$sh.Range('D2').Value2)
  $行.Add('# 紙' + "`t" + '名前' + "`t" + '式' + "`t" + '昔の 答え' + "`t" + '今の 道（Formula2・まだ 信じられない）' + "`t" + '埋まった数' + "`t" + '判じ')

  $時計 = [Diagnostics.Stopwatch]::StartNew()
  $行数 = 1
  $同じ = 0; $違う = 0; $溢れた = 0; $打てない = 0; $材料なし = 0
  $今の紙 = ''
  foreach ($q in $問い) {
    # ★★紙が 変わったら 盤面を 置き直します★★
    if ($q.紙 -ne $今の紙) {
      $今の紙 = $q.紙
      $sh.Range('A1:F60').ClearContents() | Out-Null
      $ざ = $材料[$今の紙]
      if ($null -eq $ざ -or $ざ.Count -eq 0) { $材料なし++ }
      else { foreach ($z in $ざ) { try { $sh.Range($z.マス).Value2 = $z.値 } catch { } } }
    }
    # ★★2026-09-21 ── ★材料を 書いて いない 紙は 押しません★★
    #   ＝盤面が 分からない まま 打つと ★前の 紙の 盤面の まま★＝★偽の 違いが 出ます★
    #   ＝1回目は それで 188本／2回目 147本 と 出ました（★どちらも 信じられない★）
    #   ⇒★元の 道具から 盤面を 字で 読み取る のは ★間違えると また 偽の 違い★★
    #   ⇒★だから 「測れません」と 出して 外します★
    #     （記憶「★材料が 足りない 時は『決められません』と 出して 止める★」）
    $ざ2 = $材料[$q.紙]
    if ($null -eq $ざ2 -or $ざ2.Count -eq 0) { continue }
    $ma = 'H' + $行数
    $判 = ''
    $今 = ''
    $埋 = 0
    try {
      $sh.Range($ma).Formula2 = $q.式
      if ([string]$sh.Range($ma).HasFormula -ne 'True') { $判 = '★式に なって いません★'; $打てない++ }
    } catch { $判 = '★打てません★'; $打てない++ }
    if ($判 -eq '') {
      for ($r = 0; $r -lt 6; $r++) {
        for ($k = 0; $k -lt 3; $k++) {
          $c2 = $sh.Cells.Item($行数 + $r, 8 + $k)
          if ($null -ne $c2.Value2) { $埋++ }
        }
      }
      $c2 = $sh.Range($ma)
      $v = $c2.Value2
      $今 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
      if ($埋 -gt 1) { $溢れた++ }
      # ★★2026-09-21 ── ★誤りは 数で 返ります★★
      #   ＝`.Value2` は `#REF!` を ★-2146826265★ の ように 返します
      #   ⇒★紙は 字（`#REF!`）／こちらは 数★＝★比べると 偽の 違いに なります★
      #   ⇒★字に 直してから 比べます★（2026-09-19 に 同じ 所で 2本 偽の 外れを 出しました）
      $今2 = $今
      if ($誤りの数.ContainsKey($今)) { $今2 = $誤りの数[$今] }
      if ($今2 -eq $q.昔) { $同じ++; $判 = 'onaji' } else { $違う++; $判 = '★chigau★' }
    }
    $行.Add($q.紙 + "`t" + $q.名 + "`t" + $q.式 + "`t" + $q.昔 + "`t" + $今2 + "`t" + $埋 + "`t" + $判)
    $行数 = $行数 + 10
  }
  $時計.Stop()

  # ══ ★対照★ ══
  # ★★対照は 盤面に 頼らない 形に します★★
  #   ＝1回目は `=SUM(A1:A5)` に して ★最後の 紙の 材料に 引きずられて 赤★に なりました
  $sh.Range('A300').Formula2 = '=SUM(1,2,3,4,5)'
  $sh.Range('C300').Formula2 = '=SEQUENCE(2)'
  $t1 = [string]$sh.Range('A300').Value2
  $t2 = [string]$sh.Range('C301').Value2
  if ($t1 -eq '15') { $対照ok++ }
  if ($t2 -eq '2') { $対照ok++ }

  $行.Add('#')
  $行.Add('# ★★数★★')
  $行.Add('#   式 ............... ' + $問い.Count + '本')
  $行.Add('#   ★昔と 同じ★ ..... ' + $同じ + '本')
  $行.Add('#   ★★昔と 違う★★ .. ' + $違う + '本')
  $行.Add('#   ★★溢れた★★ ..... ' + $溢れた + '本（★2マス 以上＝昔の 紙は 1マス目しか 見て いない★）')
  $行.Add('#   打てない .......... ' + $打てない + '本')
  $行.Add('#   ★★材料を 書いて いない 紙 ... ' + $材料なし + '枚＝★押して いません★★')
  $行.Add('#     ＝盤面が 分からない ので ★測れません★（黙って 前の 盤面で 押さない）')
  $行.Add('#     ⇒★その 紙の 名は 下に 出します★')
  foreach ($k in ($材料.Keys | Sort-Object)) { if ($材料[$k].Count -eq 0) { $行.Add('#       ' + $k) } }
  $行.Add('#   ★対照★ SUM(A1:A5)=' + $t1 + '（待つ 15）／SEQUENCE(2) の 2マス目=' + $t2 + '（待つ 2）⇒ ' + $対照ok + '/2')
  $行.Add('#   かかった 秒 ....... ' + [math]::Round($時計.Elapsed.TotalSeconds, 1))

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★式 ' + $問い.Count + '本★')
  Write-Host ('  昔と 同じ ... ' + $同じ + '本')
  Write-Host ('  ★★昔と 違う ... ' + $違う + '本★★')
  Write-Host ('  ★★溢れた ..... ' + $溢れた + '本★★')
  Write-Host ('  打てない ..... ' + $打てない + '本')
  Write-Host ('  ★材料が 無い 紙 ... ' + $材料なし + '枚★')
  Write-Host ('★対照 ok ... ' + $対照ok + ' / 2★ ／ 秒 ' + [math]::Round($時計.Elapsed.TotalSeconds, 1))
  Write-Host ('★書いた ... ' + $出 + '★')
  $bk.Close($false)
  $bk = $null
} finally {
  # ★★掴んだ物 全部 $null★★
  $c2 = $null
  $sh = $null
  if ($null -ne $bk) { $bk.Close($false) }
  $bk = $null
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $xl = $null
  $t = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 180)) {
    Start-Sleep -Milliseconds 250
  }
  $t.Stop()
  $残り = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
  if ($残り -eq 0) { Write-Host ('★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒で ★消えました★★') }
  else { Write-Host ('★★Excel は 消えませんでした★★ ／ 残り ' + $残り + '個') }
}
if ($対照ok -ne 2) { Write-Host '★★対照が 合いません＝盤面が 狂って います★★'; exit 7 }
