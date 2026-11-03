# toru-kobore-fusagi-yoko-2jigen-wo-excel-ni-hirakaseru.ps1
#   -- ★横・2次元が 塞がれた 溢れを 実Excel に 開かせる★（㊾）（2026-09-20）
#
#  ★★なぜ これが 一番 大事か★★
#    PR #90（塞がれた 溢れも 動く並びとして 書き出す）の 門は
#      ★書き出した 字（`ref="A1:A1"`）を 見て いるだけ★です。
#    ⇒★実Excel が ★横・2次元でも★ 同じに 振る舞うかは 誰も 測って いません★
#    ⇒★★これが 「#90 を 直す かも しれない」唯一の 測り★★
#    ⇒★もし 縦と 違う 顔が 出たら 緑でも 出しません★
#
#  ★★開く 物★★ `%TEMP%\exally-kakidashi-spill2.xlsx`（★1本の 名★）
#    16,385B ／ sha256 da5519868bc9894557dd990d97792d4c42bf5e5aca3e9436651b180737e8535a
#    ★お客さんの 道★で 作った 物（画面で 打つ ⇒ 本番の 書き出しの 2行）
#    ・★司さんの 実物の 名は 1文字も 在りません★
#    ・★読むだけ★（`Close($false)`／`SaveAs` は 1文字も 在りません）
#
#  ★★中身（生の 字で 確かめ済み）★★
#    ★塞がれた もの★
#      A2  `<c r="A2" t="str" cm="1"><f t="array" ref="A2:A2">_xlfn.SEQUENCE(1,3)</f><v>#SPILL!</v></c>`
#      A5  `<c r="A5" t="str" cm="1"><f t="array" ref="A5:A5">_xlfn.SEQUENCE(2,3)</f><v>#SPILL!</v></c>`
#    ★邪魔★   B2 `jama`（本来 A2:C2 の 行き先）／B6 `jama2`（本来 A5:C6 の 中）
#    ★対照★   E1 `ref="E1:E3"`（縦・塞がれて いない）／A10 `ref="A10:C10"`（横・同）
#    `cm="1"` 4個 ／ `t="array"` 4個
#
#  ★★4問（★縦の 時と 同じ★）★★
#    ⑴★開いた 瞬間★の A2 / A5 ... ★狙い＝#SPILL!（-2146826243）★
#        ★もし 1 なら 縦と 違う＝#90 を 直す 必要が 在ります★
#    ⑵`HasArray` が A2 / A5 で False か
#    ⑶★邪魔を 消したら 溢れ直すか★
#        B2 を 消す ⇒ A2 B2 C2 = 1/2/3 に なるか
#        B6 を 消す ⇒ A5:C6 = 1..6 に なるか（★2次元は ここが 一番 見たい 所★）
#    ⑷★対照（E1・A10）が 下がって いないか★
#
#  ★★測る 順番★★
#    ①★開いた 瞬間★（★1字も 打ちません★）
#    ②`ISERROR` `ISTEXT` `ISNUMBER`（★式を 打つと 計算し直しが 起きる ので ①の 後★）
#    ③★邪魔を 消す★（B2 ⇒ 読む ⇒ B6 ⇒ 読む）
#    ④★2つ目の 窓★（`=(マス)=0` の 真偽と 型）
#
#  ★★「変わったか」の 判じ方★★
#    ・★投げなかった＝書けた、では ありません★（COM は 黙って 何も しない）
#    ・★1マスだけ 見ては いけません★
#    ⇒★見る 範囲 全部 の 字と 値を 繋げて 比べます★
#    ・★配列の 要素は 1つずつ 括弧で 囲みます★
#      （`,` が `+` より 強く 結び付き ★3本が 1本に 繋がる★／今日 実測で 踏みました）
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②名が 1文字でも 違えば 走らない（exit 7）
#    ③ファイルが 無ければ 走らない（exit 6）／④走らせる 前の Excel が 0個（exit 3）
#    ⑤見る マスの 数が 決め打ちと 同じ（exit 4）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-kobore-fusagi-yoko-2jigen-excel-2026-09-20.tsv'

$許す名 = 'exally-kakidashi-spill2.xlsx'
$開く = Join-Path $env:TEMP $許す名

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }
if ((Split-Path $開く -Leaf) -ne $許す名) { Write-Host '★★開いて よい ファイルは 1本だけです★★'; exit 7 }
if (-not (Test-Path $開く)) { Write-Host ('★★在りません ... ' + $開く + '★★'); exit 6 }
Write-Host ('★開く 物 ... ' + $開く + '（' + (Get-Item $開く).Length + ' バイト）★')

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

# ★見る マス★（★溢れ先より 1マス 広く★＝はみ出しを 見る 為）
$見るマス = @('A2', 'B2', 'C2', 'D2',
              'A5', 'B5', 'C5', 'A6', 'B6', 'C6', 'A7',
              'E1', 'E2', 'E3', 'E4',
              'A10', 'B10', 'C10', 'D10')
$見る数 = 19
if ($見るマス.Count -ne $見る数) { Write-Host '★★見る マスの 数が 合いません★★'; exit 4 }

function 読み出す {
  param($sh, $マス)
  $c = $sh.Range($マス)
  $v = $c.Value2
  $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
  $型 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
  $一部か = '(?)'
  try { $一部か = [string]$c.HasArray } catch { }
  return @($値, [string]$c.Text, [string]$c.Formula, $型, $一部か)
}

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null; $w = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Open($開く, 0, $true)
  $sh = $bk.Sheets.Item(1)

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★横・2次元が 塞がれた 溢れを 実Excel に 開かせた★（㊾）（2026-09-20）')
  $行.Add('# ★開いた 物★ ... ' + $開く + '（' + (Get-Item $開く).Length + ' バイト）')
  $行.Add('# ★読むだけ★（保存して いません）')
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★どの 貝殻か★ ... PowerShell ' + $版.ToString())
  $行.Add('# ★生の 字★ A2 `ref="A2:A2"` =SEQUENCE(1,3)／A5 `ref="A5:A5"` =SEQUENCE(2,3)')
  $行.Add('# ★邪魔★   B2 jama（本来 A2:C2）／B6 jama2（本来 A5:C6）')
  $行.Add('# ★対照★   E1 `ref="E1:E3"`（縦）／A10 `ref="A10:C10"`（横）')
  $行.Add('# ★縦で 測った 時★ ... #SPILL! に なり 塞ぎを 消すと 溢れ直した（golden-kobore-fusagi-excel-ato-）')

  # ═══ ★★①開いた 瞬間★★ ═══
  $行.Add('#')
  $行.Add('# ★★①開いた 瞬間★★（★まだ 何も 打って いません★）')
  $行.Add('# マス' + "`t" + '値' + "`t" + '出る字' + "`t" + '式' + "`t" + '型' + "`t" + '溢れの一部か')
  $控え = New-Object System.Collections.Generic.List[string]
  foreach ($ma in $見るマス) {
    $r = 読み出す $sh $ma
    $控え.Add($ma + '=' + $r[0] + ':' + $r[2])
    $行.Add($ma + "`t" + $r[0] + "`t" + $r[1] + "`t" + $r[2] + "`t" + $r[3] + "`t" + $r[4])
  }

  # ═══ ★★②文字列か 誤りか 数かを 直に 訊く★★ ═══
  $行.Add('#')
  $行.Add('# ★★②ISERROR / ISTEXT / ISNUMBER★★（★ここから 計算し直しが 起きます★）')
  $行.Add('# マス' + "`t" + 'ISERROR' + "`t" + 'ISTEXT' + "`t" + 'ISNUMBER' + "`t" + '窓に入っている式(3つ)')
  $訊くマス = @('A2', 'A5', 'E1', 'A10')
  # ── ★1周目＝打つだけ★（窓は ★H/I/J の 40行目から★＝表と 離します）
  $窓行 = 40
  foreach ($ma in $訊くマス) {
    # ★★1つずつ 括弧で 囲みます★★（`,` が `+` より 強い＝今日 実測で 踏んだ 所）
    $式たち = @(('=ISERROR(' + $ma + ')'), ('=ISTEXT(' + $ma + ')'), ('=ISNUMBER(' + $ma + ')'))
    if ($式たち.Count -ne 3) { Write-Host '★★式が 3本 在りません＝繋がって います★★'; exit 9 }
    for ($k = 0; $k -lt 3; $k++) {
      try { $w = $sh.Cells.Item($窓行, 8 + $k); $w.Formula2 = $式たち[$k] } catch { }
    }
    $窓行 = $窓行 + 1
  }
  try { $xl.CalculateFull() } catch { }
  # ── ★2周目＝読むだけ★
  $窓行 = 40
  foreach ($ma in $訊くマス) {
    $答 = New-Object System.Collections.Generic.List[string]
    $窓式 = New-Object System.Collections.Generic.List[string]
    for ($k = 0; $k -lt 3; $k++) {
      $w = $sh.Cells.Item($窓行, 8 + $k)
      $v2 = $w.Value2
      $x2 = '(kara)'
      if ($null -ne $v2) { $x2 = [string]$v2 }
      $答.Add($x2)
      $窓式.Add([string]$w.Formula)
    }
    $窓行 = $窓行 + 1
    $行.Add($ma + "`t" + $答[0] + "`t" + $答[1] + "`t" + $答[2] + "`t" + ($窓式 -join ' / '))
  }

  # ═══ ★★③邪魔を 消したら 溢れ直すか★★ ═══
  #   ★1つずつ 消します★（★まとめて 消すと どちらが 効いたか 分かりません★）
  $消す = @(('B2'), ('B6'))
  foreach ($ke in $消す) {
    $行.Add('#')
    $行.Add('# ★★③邪魔 ' + $ke + ' を 消した 後★★')
    $判じ = '(★消せません★)'
    try {
      $sh.Range($ke).ClearContents() | Out-Null
      $束 = New-Object System.Collections.Generic.List[string]
      foreach ($ma in $見るマス) {
        $r = 読み出す $sh $ma
        $束.Add($ma + '=' + $r[0] + ':' + $r[2])
      }
      # ★★見る 範囲 全部 を 繋げて 比べます★★
      if (($束 -join '|') -eq ($控え -join '|')) { $判じ = '★1マスも 変わりません★' }
      else { $判じ = 'ok（★変わりました★）' }
      $控え = $束
    } catch { $判じ = '★投げました★ ' + $_.Exception.Message }
    $行.Add('# ★消した 結果★ ... ' + $判じ)
    $行.Add('# マス' + "`t" + '値' + "`t" + '出る字' + "`t" + '式' + "`t" + '型' + "`t" + '溢れの一部か')
    foreach ($ma in $見るマス) {
      $r = 読み出す $sh $ma
      $行.Add($ma + "`t" + $r[0] + "`t" + $r[1] + "`t" + $r[2] + "`t" + $r[3] + "`t" + $r[4])
    }
  }

  # ═══ ★★2つ目の 窓★★（★`=(マス)=0` の 真偽と 型を 一緒に★）═══
  #   ★なぜ★ ... `.Value2` の 「0」は ★本物の 0★ とも ★空★ とも ★誤りの 番号★ とも
  #              区別が 付きません（記憶「意味の 無い 数は 一番 見つけにくい」）
  $行.Add('#')
  $行.Add('# ★★2つ目の 窓★★（★上の 読みの 後に 打って います★）')
  $行.Add('# マス' + "`t" + '値' + "`t" + '型' + "`t" + '=(マス)=0')
  $窓行 = 60
  foreach ($ma2 in $見るマス) {
    $r = 読み出す $sh $ma2
    $ゼロか = '(★窓2が 打てません★)'
    try { $w = $sh.Range('N' + $窓行); $w.Formula2 = '=(' + $ma2 + ')=0'; $ゼロか = [string]$w.Value2 } catch { }
    $窓行 = $窓行 + 1
    $行.Add($ma2 + "`t" + $r[0] + "`t" + $r[3] + "`t" + $ゼロか)
  }

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ('★書いた ... ' + $出 + '★')
  # ★★保存しません★★
  $bk.Close($false)
  $bk = $null
} finally {
  $w = $null; $sh = $null
  if ($null -ne $bk) { $bk.Close($false) }
  $bk = $null
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $xl = $null
  $t = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 120)) {
    Start-Sleep -Milliseconds 250
  }
  $t.Stop()
  $残り = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
  if ($残り -eq 0) { Write-Host ('★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒で ★消えました★★') }
  else { Write-Host ('★★Excel は 消えませんでした★★ ／ 残り ' + $残り + '個') }
}
