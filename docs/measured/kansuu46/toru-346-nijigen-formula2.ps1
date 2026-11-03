# toru-346-nijigen-formula2.ps1
#   -- ★2次元の 範囲を 渡す 式を ★今の 道（Formula2）★で 取り直す★（68）（2026-09-21）
#
#  ★★なぜ★★
#    `golden-346-2026-09-08.tsv` は ★`.Formula`（昔の 道）★で 取って います。
#    ⇒`.Formula` は ★溢れません★＝★暗黙の 交わりで 1マス★
#    ⇒★範囲を 渡した 式の 答えは 1マス目だけ★（★間違いでは なく 足りない★）
#    ⇒★実測（67・2026-09-21）★ 範囲を 渡す 812本を `Formula2` で 打ち直すと
#       ★昔と 同じ 809本／違う 3本／溢れた 416本★
#       違う 3本 ＝ `ISERR(A1:B5)` `ISERROR(A1:B5)` `ISNUMBER(A1:B5)`
#    ⇒★その 3本は `tests/shiki-kansuu-kami.test.mjs` が 赤に なる 3本と 同じ★
#
#  ★★この 道具が する 事★★
#    ★2次元の 範囲（`A1:B5` の 形）を 渡す 式だけ★を `Formula2` で 取り直し、
#    ★何マス 埋まったか★と ★中身★を 書きます。
#    ⇒★古い 紙は 1文字も 触りません★（★昔の 道の 答えも 残す＝後で 比べられる★）
#
#  ★★盤面（★古い 紙と 同じ★）★★
#    A1:A5 = 1 / 2 / 3 / 4 / 5
#    B1:B5 = 2 / 4 / 6 / 8 / 10
#    D1 = 2024/1/1 ／ D2 = 2026/1/1
#
#  ★★置き場（★これを 書かないと 次の 人が 転びます★）★★
#    ★式は H列・行1 から 12行 おき★（H1 H13 H25 ほか）
#    ⇒★2026-09-21 に Exally1 が 行100 に 置いて 数を 逆に 出しました★
#      ＝★暗黙の 交わりは 式を 置いた 行／列で 決まる★
#      ⇒★今の 道（Formula2）は 交わりでは なく 溢れる★ので 置き場で 変わりませんが、
#        ★昔の 道と 比べる 時に 効きます★＝★だから 書きます★
#
#  ★★読む 窓★★ 頭の マスから ★8行 x 8列★
#    ⇒★窓が 狭いと 「溢れが 小さい」と 嘘を 吐きます★
#      ＝2026-09-21 に Exally1 が 4列の 窓で `TRANSPOSE(A1:B5)` を 8マスと 出しました
#        （★正しくは 10マス★）
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
#    ⑥★対照 2本が 待つ 通りで なければ 落とす（exit 7）★
#    ⑦★`HasFormula` を 読み戻す★＝★「投げなかった」を 「式に なった」に しない★
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$紙 = Join-Path $ここ 'golden-346-2026-09-08.tsv'
$出 = Join-Path $ここ 'golden-346-nijigen-formula2-2026-09-21.tsv'

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }
if (-not (Test-Path $紙)) { Write-Host ('★★紙が 在りません ... ' + $紙 + '★★'); exit 6 }

$外へ出る = 'WEBSERVICE', 'STOCKHISTORY', 'TRANSLATE', 'DETECTLANGUAGE', 'IMAGE', 'RTD'
$問い = New-Object System.Collections.Generic.List[object]
$外した = 0
foreach ($l in [System.IO.File]::ReadAllLines($紙)) {
  if ($l -eq '' -or $l.StartsWith('#')) { continue }
  $c = $l.Split("`t")
  if ($c.Count -lt 3) { continue }
  $f = $c[1]
  if (-not $f.StartsWith('=')) { continue }
  # ★範囲を 渡す 式だけ★（`A1:A5` の 形が 在る 物）
  # ★★2次元の 範囲だけ★★（★列も 行も 違う 物★＝`A1:B5`）
  $にじ = $false
  foreach ($m in [regex]::Matches($f, '([A-Z]{1,2})(\d+):([A-Z]{1,2})(\d+)')) {
    if ($m.Groups[1].Value -ne $m.Groups[3].Value -and $m.Groups[2].Value -ne $m.Groups[4].Value) { $にじ = $true }
  }
  if (-not $にじ) { continue }
  $わるい = $false
  foreach ($n in $外へ出る) { if ($f.ToUpper().Contains($n)) { $わるい = $true } }
  foreach ($ch in $f.ToCharArray()) { if ([int]$ch -gt 127) { $わるい = $true } }
  if ($わるい) { $外した++; continue }
  $問い.Add(@{ 名 = $c[0]; 式 = $f; 昔 = $c[2] })
}
Write-Host ('★紙から 取った 式 ... ' + $問い.Count + '本★（★外した ' + $外した + '本★）')
if ($問い.Count -eq 0) { Write-Host '★★1本も 在りません★★'; exit 5 }

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null; $c2 = $null; $w2 = $null
$対照ok = 0
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $xl.ScreenUpdating = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)
  for ($r = 1; $r -le 5; $r++) {
    $sh.Cells.Item($r, 1).Value2 = $r
    $sh.Cells.Item($r, 2).Value2 = ($r * 2)
  }
  $sh.Range('D1').Value2 = [datetime]'2024-01-01'
  $sh.Range('D2').Value2 = [datetime]'2026-01-01'

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★2次元の 範囲を 渡す 式を 今の 道（Formula2）で 取り直した★（68）（2026-09-21）')
  $行.Add('# ★元の 紙★ ... docs/measured/kansuu46/golden-346-2026-09-08.tsv')
  $行.Add('# ★★元の 紙は `.Formula`（むかしの みち＝溢れない）で 取られて います★★（kansuu46 の 道具 21本とも）')
  $行.Add('# ★盤面★ A1:A5=1,2,3,4,5 ／ B1:B5=2,4,6,8,10 ／ D1=2024/1/1 ／ D2=2026/1/1')
  $行.Add('# ★★置き場★★ ... ★式は H列・行1 から 12行 おき★（H1 H13 H25 ほか）')
  $行.Add('# ★読む 窓★ ... 頭の マスから 8行 x 8列（★窓が 狭いと「溢れが 小さい」と 嘘を 吐きます★）')
  $行.Add('# ★取り直した 訳★ ... 古い 紙は `.Formula`（むかしの みち＝溢れない）で 取って いる')
  $行.Add('# ★古い 紙は 1文字も 触って いません★（むかしの みちの 答えも 残す）')
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
  $行.Add('# 名前' + "`t" + '式' + "`t" + '昔の 答え' + "`t" + '実Excel（今の 道＝Formula2）' + "`t" + '型' + "`t" + '=(マス)=0' + "`t" + '埋まった数' + "`t" + '判じ' + "`t" + '中身（|で マス／ / で 行）')

  $時計 = [Diagnostics.Stopwatch]::StartNew()
  $行数 = 1
  $同じ = 0; $違う = 0; $溢れた = 0; $打てない = 0
  foreach ($q in $問い) {
    $ma = 'H' + $行数
    $判 = ''
    $今 = ''
    $埋 = 0
    $中身 = ''
    $型 = ''
    $ゼロか = ''
    try {
      $sh.Range($ma).Formula2 = $q.式
      if ([string]$sh.Range($ma).HasFormula -ne 'True') { $判 = '★式に なって いません★'; $打てない++ }
    } catch { $判 = '★打てません★'; $打てない++ }
    if ($判 -eq '') {
      $並 = New-Object System.Collections.Generic.List[string]
      for ($r = 0; $r -lt 8; $r++) {
        $一行 = New-Object System.Collections.Generic.List[string]
        $からの行 = $true
        for ($k = 0; $k -lt 8; $k++) {
          $c2 = $sh.Cells.Item($行数 + $r, 8 + $k)
          if ($null -ne $c2.Value2) { $埋++; $からの行 = $false; $一行.Add([string]$c2.Text) }
          else { $一行.Add('') }
        }
        if (-not $からの行) { $並.Add(($一行 -join '|').TrimEnd('|')) }
      }
      $中身 = ($並 -join ' / ')
      $c2 = $sh.Range($ma)
      $v = $c2.Value2
      $今 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
      # ══ ★★2つ目の 窓★★（★`=(マス)=0` の 真偽と ★型★ を 一緒に 取る★）══
      #   ★なぜ★ ... `.Value2` の 「0」は ★本物の 0★ とも ★空★ とも ★誤りの 番号★ とも
      #              区別が 付きません（記憶「意味の 無い 数は 一番 見つけにくい」）
      #   ★置き場★ ... Z列（★溢れは H から 8列＝O まで＝ぶつかりません★）
      $型 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
      $ゼロか = '(★窓2が 打てません★)'
      try { $w2 = $sh.Range('Z' + $行数); $w2.Formula2 = '=(' + $ma + ')=0'; $ゼロか = [string]$w2.Value2 } catch { }
      if ($埋 -gt 1) { $溢れた++ }
      if ($今 -eq $q.昔) { $同じ++; $判 = 'onaji' } else { $違う++; $判 = '★chigau★' }
    }
    $行.Add($q.名 + "`t" + $q.式 + "`t" + $q.昔 + "`t" + $今 + "`t" + $型 + "`t" + $ゼロか + "`t" + $埋 + "`t" + $判 + "`t" + $中身)
    $行数 = $行数 + 12
  }
  $時計.Stop()

  # ══ ★対照★ ══
  $sh.Range('A300').Formula2 = '=SUM(A1:A5)'
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
  $行.Add('#   ★対照★ SUM(A1:A5)=' + $t1 + '（待つ 15）／SEQUENCE(2) の 2マス目=' + $t2 + '（待つ 2）⇒ ' + $対照ok + '/2')
  $行.Add('#   かかった 秒 ....... ' + [math]::Round($時計.Elapsed.TotalSeconds, 1))

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★式 ' + $問い.Count + '本★')
  Write-Host ('  昔と 同じ ... ' + $同じ + '本')
  Write-Host ('  ★★昔と 違う ... ' + $違う + '本★★')
  Write-Host ('  ★★溢れた ..... ' + $溢れた + '本★★')
  Write-Host ('  打てない ..... ' + $打てない + '本')
  Write-Host ('★対照 ok ... ' + $対照ok + ' / 2★ ／ 秒 ' + [math]::Round($時計.Elapsed.TotalSeconds, 1))
  Write-Host ('★書いた ... ' + $出 + '★')
  $bk.Close($false)
  $bk = $null
} finally {
  # ★★掴んだ物 全部 $null★★
  $w2 = $null
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
