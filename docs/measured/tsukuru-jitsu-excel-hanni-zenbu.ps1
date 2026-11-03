# tsukuru-jitsu-excel-hanni-zenbu.ps1
#   -- ★溢れる 式 全部を ★1つの 盤面★で 取り直す（★中身まで★）★（65）（2026-09-20）
#
#  ★★なぜ 取り直すか（★私の 落ち度★）★★
#    紙 `golden-jitsu-excel-no-shirushi-2026-09-20.tsv` は ★6回分が 1枚に なって います★
#      51 ... A1:A3 = 3,1,2 ／ 52 ... A1:A6 = 1..6 ／ 55・58・59・61 ... A1:A6 = 1,2,2,3,3,4
#    ⇒★`ref`（何マス）は ★その 回の 盤面★で 決まります★
#    ⇒Exally1 は 61 の 盤面で 打って 私の `ref` と 比べました
#    ⇒★`UNIQUE` が 「うち 2／実Excel 3」に 見えました＝★偽の 外れ★★
#       （★どちらの 盤面でも 実Excel と 同じ★でした）
#    ⇒★★私が 紙に 「どの 盤面で 測ったか」を 書いて いなかった★★
#    ⇒記憶「★何を 何で 数えたか★」「★どこを 数えたか★」の ★次★
#       ＝★★その 数は どの 盤面で 取った 物か★★
#
#  ★★もう 1つ の 訳★★
#    ★`UNIQUE` は 差が 1 だったので 気付けました★
#    ⇒★「3 → 3 で 中身だけ 違う」物は ★マス数だけでは 気付けません★★
#    ⇒★だから 中身（値）も 出します★
#
#  ★★盤面（★この 紙は 全部 この 1つ★）★★
#    A1:A6 = 1 / 2 / 2 / 3 / 3 / 4
#    B1:B3 = 10 / 20 / 30
#    D1    = abc
#    E1 F1 = 1 / 2     E2 F2 = 3 / 4
#    ★A20 = 5（数）／A21 = `=1+1`（★式★）／A22 = abc（字）★
#      ⇒`ISFORMULA` `FORMULATEXT` を 割る 為
#    ⇒★61 の 盤面 ＋ A20:A22★（★Exally1 が 打った 盤面に 揃えます★）
#
#  ★★何を 打つか★★
#    ★紙から `cm` が 在る 行の 式を 取り出します★
#    ★印を 外します★ `_xlfn._xlws.` ⇒ `` ／ `_xlfn.` ⇒ `` ／ `_xlpm.` ⇒ `` ／ `_xleta.` ⇒ ``
#      ＝★COM に 打つ 時は 印を 付けません★（実Excel が 自分で 付けます）
#    ★同じ 式は 1回だけ★（紙に 2回 出る 物が 在ります）
#
#  ★★読む 物★★
#    頭の マスから ★6行 × 6列★（★大きい 物も 収める★）
#    ・★埋まった マスの 数★
#    ・★中身（値）★（★「3 → 3 で 中身だけ 違う」を 見る 為★）
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③紙が 読めなければ 走らない（exit 6）／④式が 0本なら 走らない（exit 5）
#    ⑤★外へ 出る 6個は 打ちません★（数を 出します）
#    ⑥★対照 2個（SEQUENCE・SUM）が 待つ 通りで なければ 落とす（exit 7）★
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$紙 = Join-Path $ここ 'golden-jitsu-excel-no-shirushi-2026-09-20.tsv'
$出 = Join-Path $ここ 'golden-jitsu-excel-hanni-zenbu-2026-09-20.tsv'

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }
if (-not (Test-Path $紙)) { Write-Host ('★★紙が 在りません ... ' + $紙 + '★★'); exit 6 }

# ══ ★紙から 「cm が 在る」式を 取り出す★ ══
$外へ出る = 'WEBSERVICE', 'STOCKHISTORY', 'TRANSLATE', 'DETECTLANGUAGE', 'IMAGE', 'RTD'
$式たち = New-Object System.Collections.Generic.List[string]
$見た = New-Object 'System.Collections.Generic.HashSet[string]'
$外した = 0
foreach ($l in [System.IO.File]::ReadAllLines($紙)) {
  if ($l -eq '' -or $l.StartsWith('#')) { continue }
  $c = $l.Split("`t")
  if ($c.Count -lt 4) { continue }
  if ($c[1].Trim() -ne '在り') { continue }
  $f = $c[3]
  $f = $f.Replace('_xlfn._xlws.', '').Replace('_xlfn.', '').Replace('_xlws.', '').Replace('_xlpm.', '').Replace('_xleta.', '')
  $わるい = $false
  foreach ($n in $外へ出る) { if ($f.ToUpper().Contains($n)) { $わるい = $true } }
  foreach ($ch in $f.ToCharArray()) { if ([int]$ch -gt 127) { $わるい = $true } }
  if ($わるい) { $外した++; continue }
  # ★★2026-09-20 ── ★紙の 式には `=` が 付いて いません★★
  #   ＝xlsx は `<f>SEQUENCE(3)</f>` の ように ★`=` 無しで★ 持ちます
  #   ⇒★そのまま 打つと ★字★として 入ります★（★投げません／もっともらしい★）
  #   ⇒1回目は これで ★112本 とも 「1マスに 式の 字」★に なりました
  if (-not $f.StartsWith('=')) { $f = '=' + $f }
  if ($見た.Add($f)) { $式たち.Add($f) }
}
Write-Host ('★紙から 取った 式 ... ' + $式たち.Count + '本★（★外した ' + $外した + '本★）')
if ($式たち.Count -eq 0) { Write-Host '★★1本も 在りません★★'; exit 5 }

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null; $c2 = $null
$対照ok = 0
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)
  $材 = 1, 2, 2, 3, 3, 4
  for ($r = 1; $r -le 6; $r++) { $sh.Cells.Item($r, 1).Value2 = $材[$r - 1] }
  $sh.Range('B1').Value2 = 10
  $sh.Range('B2').Value2 = 20
  $sh.Range('B3').Value2 = 30
  $sh.Range('D1').Value2 = 'abc'
  $sh.Range('E1').Value2 = 1
  $sh.Range('F1').Value2 = 2
  $sh.Range('E2').Value2 = 3
  $sh.Range('F2').Value2 = 4
  $sh.Range('A20').Value2 = 5
  $sh.Range('A21').Formula2 = '=1+1'
  $sh.Range('A22').Value2 = 'abc'

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★溢れる 式 全部を ★1つの 盤面★で 取り直した（★中身まで★）★（65）（2026-09-20）')
  $行.Add('# ★★盤面（★この 紙は 全部 この 1つ★）★★')
  $行.Add('#   A1:A6 = 1,2,2,3,3,4 ／ B1:B3 = 10,20,30 ／ D1 = abc ／ E1,F1,E2,F2 = 1,2,3,4')
  $行.Add('#   A20 = 5（数）／A21 = =1+1（★式★）／A22 = abc（字）')
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★どの 貝殻か★ ... PowerShell ' + $版.ToString())
  $行.Add('# ★印（_xlfn. 等）は 外して 打って います★＝実Excel が 自分で 付けます')
  $行.Add('# ★読む★ ... 頭の マスから 6行 x 6列')
  # ══ ★★`#材料` を 書きます★★（★これが 無いと 試験は 1行も 押しません★）
  $行.Add('#材料' + "`t" + 'A1' + "`t" + '1')
  $行.Add('#材料' + "`t" + 'A2' + "`t" + '2')
  $行.Add('#材料' + "`t" + 'A3' + "`t" + '2')
  $行.Add('#材料' + "`t" + 'A4' + "`t" + '3')
  $行.Add('#材料' + "`t" + 'A5' + "`t" + '3')
  $行.Add('#材料' + "`t" + 'A6' + "`t" + '4')
  $行.Add('#材料' + "`t" + 'B1' + "`t" + '10')
  $行.Add('#材料' + "`t" + 'B2' + "`t" + '20')
  $行.Add('#材料' + "`t" + 'B3' + "`t" + '30')
  $行.Add('#材料' + "`t" + 'D1' + "`t" + 'abc')
  $行.Add('#材料' + "`t" + 'E1' + "`t" + '1')
  $行.Add('#材料' + "`t" + 'F1' + "`t" + '2')
  $行.Add('#材料' + "`t" + 'E2' + "`t" + '3')
  $行.Add('#材料' + "`t" + 'F2' + "`t" + '4')
  $行.Add('#材料' + "`t" + 'A20' + "`t" + '5')
  $行.Add('#材料' + "`t" + 'A22' + "`t" + 'abc')
  $行.Add('# ★A21 は 式（=1+1）なので 材料に 書いて いません★＝ISFORMULA の 行は 押されない')
  $行.Add('# 式' + "`t" + '実Excel（今の 道＝Formula2）' + "`t" + '埋まった数' + "`t" + '中身（|で マス／ / で 行）')

  $時計 = [Diagnostics.Stopwatch]::StartNew()
  $行数 = 1
  $打てた = 0
  foreach ($f in $式たち) {
    $ma = 'H' + $行数
    $投げた = ''
    # ★★門＝「投げなかった」を 「式に なった」に しない★★
    #   ＝`HasFormula` が True か 読み戻して 判じます（★今日 何度も 踏んだ 形★）
    try {
      $sh.Range($ma).Formula2 = $f
      if ([string]$sh.Range($ma).HasFormula -eq 'True') { $打てた++ }
      else { $投げた = '(★式に なって いません＝字で 入りました★)' }
    } catch { $投げた = '(★打てません★)' }
    $並 = New-Object System.Collections.Generic.List[string]
    $埋 = 0
    for ($r = 0; $r -lt 6; $r++) {
      $一行 = New-Object System.Collections.Generic.List[string]
      $からの行 = $true
      for ($k = 0; $k -lt 6; $k++) {
        $c2 = $sh.Cells.Item($行数 + $r, 8 + $k)
        $v = $c2.Value2
        if ($null -eq $v) { $一行.Add('') }
        else {
          $埋++; $からの行 = $false
          # ★★2026-09-21 ── ★`.Text` は 桁を 丸めます★★
          #   ＝`-0.41615`（5桁）／`1.313035`（7桁）＝★桁が 一定で 在りません★
          #   ⇒Exally1 が `COS` `COT` `COTH` `RADIANS` の 4本を
          #     ★「揃えても 違う」★として 分けられませんでした
          #   ⇒★数は `R`（全桁）／数で ない 物は 出る字★
          if ($v -is [double]) { $一行.Add($v.ToString('R', [Globalization.CultureInfo]::InvariantCulture)) }
          else { $一行.Add([string]$c2.Text) }
        }
      }
      if (-not $からの行) { $並.Add(($一行 -join '|').TrimEnd('|')) }
    }
    # ★★物差しの 列＝頭の マスの 答え★★（★全桁★）
    $あたま = $sh.Range($ma).Value2
    $答 = if ($null -eq $あたま) { '(kara)' } elseif ($あたま -is [double]) { $あたま.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$sh.Range($ma).Text }
    $行.Add($f + "`t" + $答 + "`t" + $埋 + "`t" + ($並 -join ' / ') + $投げた)
    $行数 = $行数 + 10
  }
  $時計.Stop()

  # ══ ★門⑥＝対照★ ══
  $sh.Range('A200').Formula2 = '=SEQUENCE(3)'
  $sh.Range('C200').Formula2 = '=SUM(A1:A3)'
  $t1 = [string]$sh.Range('A202').Value2
  $t2 = [string]$sh.Range('C200').Value2
  if ($t1 -eq '3') { $対照ok++ }
  if ($t2 -eq '5') { $対照ok++ }
  $行.Add('#')
  $行.Add('# ★★対照★★ SEQUENCE(3) の 3マス目 ... ' + $t1 + '（待つ 3）／ SUM(A1:A3) ... ' + $t2 + '（待つ 5）')
  $行.Add('# ★対照 ok ... ' + $対照ok + ' / 2★')
  $行.Add('# ★打てた ' + $打てた + ' / ' + $式たち.Count + '本★ ／ かかった 秒 ... ' + [math]::Round($時計.Elapsed.TotalSeconds, 1))

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★★打てた ' + $打てた + ' / ' + $式たち.Count + '本★★ ／ 秒 ' + [math]::Round($時計.Elapsed.TotalSeconds, 1))
  Write-Host ('★対照 ok ... ' + $対照ok + ' / 2★')
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
