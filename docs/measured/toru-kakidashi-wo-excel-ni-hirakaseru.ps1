# toru-kakidashi-wo-excel-ni-hirakaseru.ps1
#   -- ★うちが 書き出した .xlsx を 実Excel に 開かせる★（㊶）（2026-09-20）
#
#  ★★なぜ★★
#    司さん（2026-09-16）「Exally と Excel に 引き渡しても ちゃんと 動くか
#      ★実際に 動作確認しながら★ やれよ」
#    ⇒★溢れ（スピル）を 書き出した ファイルを ★実Excel が どう 読むか★ は 未測定★
#
#  ★★先に 分かって いる 差（XML を 直に 読んだ）★★
#    ★実Excel★ `<c r="D1" cm="1"><f t="array" ref="D1:D3">_xlfn.SEQUENCE(3)</f><v>1</v></c>`
#    ★うち★   `<c r="D1"><f>_xlfn.SEQUENCE(3)</f><v>1</v></c>`
#    ⇒★`cm=` `t="array"` `ref=` が ★3つとも 在りません★★
#    ⇒★でも D2 D3 には ★値だけ★ 書いて ある★
#    ⇒★★だから 開いた 時に ★溢れ先が 塞がって #SPILL!★ に なる 恐れが 在ります★★
#
#  ★★開く 門（★ここが 大事★）★★
#    ★この 道具は ★1本の ファイルしか 開けません★★
#      `%TEMP%\exally-kakidashi.xlsx`（★うちの 画面が 書き出した 物★）
#    ・★名前が 1文字でも 違えば 走りません（exit 7）★
#    ・★`.xlsb` や 司さんの 実物の 名は 1文字も 在りません★
#    ・★読むだけ★（★保存しません／`$bk.Close($false)`★）
#
#  ★他の 門★
#    ①走らせる 前の Excel が 0個（exit 3）／②貝殻が powershell.exe（5.1）（exit 8）
#    ③開く ファイルが 無ければ 走らない（exit 6）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

# ★★2026-09-20 足し ── ★`-書ける` ＝ ReadOnly を 外して 開く★★
#   ★なぜ★ ... ㊷の ⑶（頭を 直したら 下が 追随するか）が 「書き換わらない」で 終わった
#              ⇒★その訳が ①CSE（並びの 一部は 直せない） なのか
#                          ②ReadOnly なのか ★切り分けられません★
#   ⇒★★開く 名前は 1本の まま★★（`%TEMP%\exally-kakidashi.xlsx`）
#   ⇒★★保存は どちらでも しません★★（`Close($false)`／`SaveAs` は 1文字も 在りません）
param([switch]$書ける)

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = if ($書ける) { Join-Path $ここ 'golden-kakidashi-excel-kakeru-2026-09-20.tsv' }
      else { Join-Path $ここ 'golden-kakidashi-excel-2026-09-20.tsv' }

# ══ ★開いて よい ただ 1本★ ══
$許す名 = 'exally-kakidashi.xlsx'
$開く = Join-Path $env:TEMP $許す名

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

# ★★開く 相手を 1本に 縛る★★（★名前が 違えば 走りません★）
if ((Split-Path $開く -Leaf) -ne $許す名) {
  Write-Host '★★開いて よい ファイルは 1本だけです★★'
  exit 7
}
if (-not (Test-Path $開く)) {
  Write-Host ('★★書き出した ファイルが 在りません ... ' + $開く + '★★')
  Write-Host '  ⇒★先に 画面から 書き出して ください★'
  exit 6
}
Write-Host ('★開く 物 ... ' + $開く + '（' + (Get-Item $開く).Length + ' バイト）★')

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$xl = New-Object -ComObject Excel.Application
$bk = $null
$sh = $null
$c = $null
$w = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  # ★`ReadOnly` を 立てる／外す★（★どちらでも 保存は しません★）
  $読むだけ = -not $書ける
  Write-Host ('★開き方 ... ReadOnly=' + $読むだけ + '★（★保存は しません★）')
  $bk = $xl.Workbooks.Open($開く, 0, $読むだけ)
  $sh = $bk.Sheets.Item(1)

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★うちが 書き出した .xlsx を 実Excel に 開かせた★（2026-09-20）')
  $行.Add('# ★開いた 物★ ... ' + $開く)
  $行.Add('# ★開き方★ ... ReadOnly=' + $読むだけ + '（★保存して いません★）')
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★どの 貝殻か★ ... PowerShell ' + $版.ToString())
  $行.Add('# マス' + "`t" + '値' + "`t" + '出る字' + "`t" + '式' + "`t" + '型' + "`t" + '=(マス)=0')

  for ($r = 1; $r -le 5; $r++) {
    $c = $sh.Cells.Item($r, 4)
    $v = $c.Value2
    $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    $型 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
    $字 = [string]$c.Text
    $式 = [string]$c.Formula
    # ★2つ目の 窓★（★他の 道具と 同じ 書き方★）
    $見る = 'D' + $r
    $ゼロか = '(★窓2が 打てません★)'
    try { $w = $sh.Range('J' + $r); $w.Formula2 = '=(' + $見る + ')=0'; $ゼロか = [string]$w.Value2 } catch { }
    $行.Add(('D' + $r) + "`t" + $値 + "`t" + $字 + "`t" + $式 + "`t" + $型 + "`t" + $ゼロか)
  }

  # ══════════════════════════════════════════════════════════════════
  # ★★㊷ ── ★生きて いるか★ の 決め手★★（2026-09-20 足し）
  #   Exally1 の 問い ⑶「★頭を 直したら 下が 追随するか★」
  #     ＝D1 を `=SEQUENCE(2)` に 書き換えたら ★D3 は 消えるか★
  #   ・★死んだ 値★なら D3 は ★3 の まま 残ります★
  #   ・★生きた 溢れ★なら D3 は ★空に なります★（㊵の 台本5 と 同じ 形）
  #   ★保存は しません★（`ReadOnly` の まま 覚えの 中だけで 書き換え／`Close($false)`）
  # ══════════════════════════════════════════════════════════════════
  $行.Add('#')
  $行.Add('# ★★㊷ ── 頭（D1）を =SEQUENCE(2) に 書き換えた 後★★（★保存して いません★）')
  $行.Add('# マス' + "`t" + '値' + "`t" + '出る字' + "`t" + '式' + "`t" + '型' + "`t" + '溢れの一部か')
  # ★★2026-09-20 ── ★「投げなかった」を 「書けた」に するな★★
  #   ＝実測で ★`.Formula2` は 投げずに 何も しません★（並びの 一部だから）
  #   ⇒★★読み戻して 字が 変わったかで 判じます★★
  #     （記憶「見張りは 印では なく ★絵が 変わったか★」／「壊したのに 赤に ならない」）
  $書けたか = '(★書き換えられません★)'
  $前の字 = [string]$sh.Range('D1').Formula
  try {
    $sh.Range('D1').Formula2 = '=SEQUENCE(2)'
    $後の字 = [string]$sh.Range('D1').Formula
    if ($後の字 -eq $前の字) {
      $書けたか = '★投げないが 変わりません★（前 ' + $前の字 + ' ／ 後 ' + $後の字 + '）'
    } else {
      $書けたか = 'ok（' + $前の字 + ' ⇒ ' + $後の字 + '）'
    }
  } catch { $書けたか = '★投げました★ ' + $_.Exception.Message }
  $行.Add('# ★頭を 書き換えられたか★ ... ' + $書けたか)
  if ($書けたか -like 'ok*') {
    for ($r = 1; $r -le 5; $r++) {
      $c = $sh.Cells.Item($r, 4)
      $v = $c.Value2
      $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
      $型 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
      $字 = [string]$c.Text
      $式 = [string]$c.Formula
      $一部か = '(?)'
      try { $一部か = [string]$c.HasArray } catch { }
      $行.Add(('D' + $r) + "`t" + $値 + "`t" + $字 + "`t" + $式 + "`t" + $型 + "`t" + $一部か)
    }
  }

  # ══════════════════════════════════════════════════════════════════
  # ★★㊸ ── ★並び 全体★を 書き換える★★（★切り分けの 2本目★）
  #   ・昔の CSE は ★一部だけ★は 直せないが ★範囲ごと★なら 直せる
  #   ⇒★ここが 通れば 「⑶が 動かなかった 訳は ★並びの 一部だから★」★
  #   ⇒★ここも 通らなければ 「訳は ★ReadOnly★」★（`-書ける` で もう 一度 測る）
  # ══════════════════════════════════════════════════════════════════
  $行.Add('#')
  $行.Add('# ★★㊸ ── 並び 全体（D1:D3）を =SEQUENCE(2) に 書き換えた 後★★')
  # ★★ここも 読み戻して 判じます★★（★同じ 穴を 2つ 空けない★）
  $全体 = '(★書き換えられません★)'
  # ★★2026-09-20 ── ★D1 の 字だけ 見ては いけません★★
  #   ＝㊷で 既に D1 が `=SEQUENCE(2)` に なって いると
  #     ★同じ 字を 書いた 事に なり 「変わりません」と 嘘を 吐きます★
  #   ⇒★★D1:D3 の 3つ 全部の 字を 繋げて 比べます★★
  $字を繋ぐ = { (1..3 | ForEach-Object { [string]$sh.Cells.Item($_, 4).Formula }) -join '|' }
  $前の字2 = & $字を繋ぐ
  try {
    $sh.Range('D1:D3').Formula2 = '=SEQUENCE(2)'
    $後の字2 = & $字を繋ぐ
    if ($後の字2 -eq $前の字2) {
      $全体 = '★投げないが 変わりません★（前 ' + $前の字2 + ' ／ 後 ' + $後の字2 + '）'
    } else {
      $全体 = 'ok（' + $前の字2 + ' ⇒ ' + $後の字2 + '）'
    }
  } catch { $全体 = '★投げました★ ' + $_.Exception.Message }
  $行.Add('# ★並び 全体を 書き換えられたか★ ... ' + $全体)
  $行.Add('# マス' + "`t" + '値' + "`t" + '出る字' + "`t" + '式' + "`t" + '型' + "`t" + '溢れの一部か')
  for ($r = 1; $r -le 5; $r++) {
    $c = $sh.Cells.Item($r, 4)
    $v = $c.Value2
    $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    $型 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
    $字 = [string]$c.Text
    $式 = [string]$c.Formula
    $一部か = '(?)'
    try { $一部か = [string]$c.HasArray } catch { }
    $行.Add(('D' + $r) + "`t" + $値 + "`t" + $字 + "`t" + $式 + "`t" + $型 + "`t" + $一部か)
  }

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ('★書いた ... ' + $出 + '★')
  # ★★保存しません★★（`$false`＝変更を 捨てる）
  $bk.Close($false)
} finally {
  $c = $null
  $w = $null
  $sh = $null
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
  if ($残り -eq 0) {
    Write-Host ('★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒で ★消えました★★')
  } else {
    Write-Host ('★★Excel は 消えませんでした★★ ／ 残り ' + $残り + '個')
  }
}
