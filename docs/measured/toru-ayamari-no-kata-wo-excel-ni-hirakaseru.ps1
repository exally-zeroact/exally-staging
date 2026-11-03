# toru-ayamari-no-kata-wo-excel-ni-hirakaseru.ps1
#   -- ★誤りの 型（`t="str"` と `t="e"`）を 実Excel に 開かせて 比べる★（㊻）（2026-09-20）
#
#  ★★なぜ 要るか★★
#    うちの 書き出しは 誤りを ★`t="str"`（文字列）★で 書いて います（★正しくは `t="e"`★）。
#    Exally1 が 「古い 8種は `t:"e"` に 直せる」と 数えました。
#    ⇒★でも 「直したら 良く なるか」は ★実Excel で 開くまで 言えません★★
#    ⇒★甲（今の まま）と 乙（直した 物）を ★同じ 道具で 続けて 開いて★ 比べます★
#
#  ★★開く 物（★2本とも 名前を 字で 書いて います★）★★
#    甲 `%TEMP%\exally-ayamari-mae.xlsx`  ... `t="str"` 5個／`t="e"` 0個
#    乙 `%TEMP%\exally-ayamari-ato.xlsx`  ... `t="e"` 5個／`t="str"` 0個
#    ★引数は 取りません★＝★この 2本しか 開けません★（★門を 緩めない 為に わざと★）
#    ・★司さんの 実物の 名は 1文字も 在りません★
#    ・★読むだけ★（`Close($false)`／`SaveAs` は 1文字も 在りません）
#
#  ★★中身（両方 同じ）★★
#    A1 `=1/0`      ⇒ #DIV/0!
#    A2 `=NA()`     ⇒ #N/A
#    A3 `=ZZZ()`    ⇒ #NAME?
#    A4 `=1+"a"`    ⇒ #VALUE!
#    A5 `=SQRT(-1)` ⇒ #NUM!
#    A6 `=SUM(1,2)` ⇒ 3 ★対照＝誤りで ない 式★
#
#  ★★測る 順番（★ここが 大事★）★★
#    ①★開いた 瞬間★（★まだ 1字も 打って いません★）
#    ②★F9（`CalculateFull`）の 後★
#       ＝★開いた 瞬間だけの 見た目か／残る 物かを 分ける★
#    ③★`ISERROR` `ISTEXT` `ISNUMBER` で 直に 訊く★
#       ＝★式を 打つと そこで 計算し直しが 起きる ので ①②の 後★
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②2本とも 無ければ 走らない（exit 6）
#    ③走らせる 前の Excel が 0個（exit 3）／④台本の 数 決め打ち（exit 4）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-ayamari-no-kata-excel-2026-09-20.tsv'

# == ★開いて よい 2本（★字で 書いて います★）== #
$二本 = @(
  @{ 札 = 'kou(t=str・今のまま)'; 名 = 'exally-ayamari-mae.xlsx' },
  @{ 札 = 'otsu(t=e・直した物)';  名 = 'exally-ayamari-ato.xlsx' }
)
if ($二本.Count -ne 2) { exit 4 }

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

foreach ($x in $二本) {
  $p = Join-Path $env:TEMP $x.名
  if (-not (Test-Path $p)) { Write-Host ('★★在りません ... ' + $p + '★★'); exit 6 }
  Write-Host ('★開く 物 ... ' + $p + '（' + (Get-Item $p).Length + ' バイト）★')
}

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null; $c = $null; $w = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★誤りの 型（t="str" と t="e"）を 実Excel に 開かせて 比べた★（㊻）（2026-09-20）')
  $行.Add('# ★読むだけ★（保存して いません）／★開けるのは 字で 書いた 2本だけ★')
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★どの 貝殻か★ ... PowerShell ' + $版.ToString())

  foreach ($x in $二本) {
    $p = Join-Path $env:TEMP $x.名
    $bk = $xl.Workbooks.Open($p, 0, $true)
    $sh = $bk.Sheets.Item(1)

    $行.Add('#')
    $行.Add('# ═══ ★★' + $x.札 + '★★ ... ' + $x.名 + '（' + (Get-Item $p).Length + ' バイト）═══')

    # ── ★★①開いた 瞬間★★（★まだ 1字も 打って いません★）
    $行.Add('# ★★①開いた 瞬間★★（★まだ 何も 打って いません★）')
    $行.Add('# 札' + "`t" + 'マス' + "`t" + '値' + "`t" + '出る字' + "`t" + '式' + "`t" + '型')
    $控え = New-Object System.Collections.Generic.List[string]
    for ($r = 1; $r -le 6; $r++) {
      $c = $sh.Cells.Item($r, 1)
      $v = $c.Value2
      $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
      $型 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
      $控え.Add('A' + $r + '=' + $値 + ':' + $型)
      $行.Add($x.札 + "`t" + ('A' + $r) + "`t" + $値 + "`t" + [string]$c.Text + "`t" + [string]$c.Formula + "`t" + $型)
    }

    # ── ★★②F9（計算し直し）の 後★★
    $行.Add('# ★★②F9（CalculateFull）の 後★★')
    $直った = '(★呼べません★)'
    try { $xl.CalculateFull(); $直った = 'ok' } catch { $直った = '★投げました★ ' + $_.Exception.Message }
    $後 = New-Object System.Collections.Generic.List[string]
    for ($r = 1; $r -le 6; $r++) {
      $c = $sh.Cells.Item($r, 1)
      $v = $c.Value2
      $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
      $型 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
      $後.Add('A' + $r + '=' + $値 + ':' + $型)
      $行.Add($x.札 + "`t" + ('A' + $r) + "`t" + $値 + "`t" + [string]$c.Text + "`t" + [string]$c.Formula + "`t" + $型)
    }
    # ★★「変わったか」は ★6マス 全部★ を 繋げて 比べます★★
    $判じ = if (($後 -join '|') -eq ($控え -join '|')) { '★①と ②は 同じ★（＝開いた 瞬間の 見た目では ない）' }
            else { '★①と ②は 違う★（＝F9 で 変わる）' }
    $行.Add('# ★CalculateFull★ ... ' + $直った + ' ／ ' + $判じ)

    # ── ★★③直に 訊く★★
    $行.Add('# ★★③ISERROR / ISTEXT / ISNUMBER で 直に 訊く★★')
    # ★★2026-09-20 ── ★打った 直後に 同じ 窓を 読む 作りは 嘘を 吐きました★★
    #   ＝1回目は A6（＝3）まで ISERROR が True／ISTEXT と ISNUMBER が ★空★ に なった
    #   ⇒★打つ のと 読む のを 2周に 分け、★窓に 何が 入って いるか（式そのもの）も 書きます★★
    #   ⇒★これで 「打てて いない」のか 「答えが そう なのか」が 見分けられます★
    $行.Add('# 札' + "`t" + 'マス' + "`t" + 'ISERROR' + "`t" + 'ISTEXT' + "`t" + 'ISNUMBER' + "`t" + '窓に入っている式(3つ)')
    # ── ★1周目＝打つだけ★（窓は ★C/D/E 列の 20行目から 下★＝A列と 離します）
    for ($r = 1; $r -le 6; $r++) {
      # ★★2026-09-20 ── ★PowerShell は `,` が `+` より 強く 結び付きます★★
      #   ＝括弧を 付けないと ★3本が 1本に 繋がって 個数 1★ に なります（実測）
      #     `@('=ISERROR(A' + $r + ')', '=ISTEXT(A' + $r + ')', ほか)`
      #       ⇒ 個数 1 ／ 中身 `=ISERROR(A1) =ISTEXT(A1) =ISNUMBER(A1)`
      #   ⇒★しかも その 式は Excel で ★比べ算の 連なり★に なり ★6行 全部 True★ を 返します★
      #     ＝★投げない／空に ならない／もっともらしい 嘘★＝一番 見つけにくい 形
      #   ⇒★1つずつ 括弧で 囲みます★
      $式たち = @(('=ISERROR(A' + $r + ')'), ('=ISTEXT(A' + $r + ')'), ('=ISNUMBER(A' + $r + ')'))
      if ($式たち.Count -ne 3) { Write-Host '★★式が 3本 在りません＝繋がって います★★'; exit 9 }
      for ($k = 0; $k -lt 3; $k++) {
        try { $w = $sh.Cells.Item(19 + $r, 3 + $k); $w.Formula2 = $式たち[$k] } catch { }
      }
    }
    try { $xl.CalculateFull() } catch { }
    # ── ★2周目＝読むだけ★
    for ($r = 1; $r -le 6; $r++) {
      $答 = New-Object System.Collections.Generic.List[string]
      $窓式 = New-Object System.Collections.Generic.List[string]
      for ($k = 0; $k -lt 3; $k++) {
        $w = $sh.Cells.Item(19 + $r, 3 + $k)
        $v2 = $w.Value2
        # ★5.1 は `if` を 式として 渡せません★（$答.Add((if ～)) は 落ちます）
        $x2 = '(kara)'
        if ($null -ne $v2) { $x2 = [string]$v2 }
        $答.Add($x2)
        $窓式.Add([string]$w.Formula)
      }
      $行.Add($x.札 + "`t" + ('A' + $r) + "`t" + $答[0] + "`t" + $答[1] + "`t" + $答[2] + "`t" + ($窓式 -join ' / '))
    }

    $窓行 = 40
    # ═══ ★★2つ目の 窓★★（★`=(マス)=0` の 真偽と 型を 一緒に 取る★）═══
    #   ★なぜ★ ... `.Value2` の 「0」は ★本物の 0★ とも ★空★ とも ★誤りの 番号★ とも
    #              区別が 付きません（記憶「意味の 無い 数は 一番 見つけにくい」）
    #   ⇒★別の 口（式）で もう 一度 0 かを 訊いて 型と 並べます★
    #   ★この 窓は ★上の 読みが 済んだ 後★に 打ちます★
    #     ＝式を 打つと 計算し直しが 起きる ので ★「開いた 瞬間」を 汚さない★
    $行.Add('#')
    $行.Add('# ★★2つ目の 窓★★（★上の 読みの 後に 打って います★）')
    $行.Add('# マス' + "`t" + '値' + "`t" + '型' + "`t" + '=(マス)=0')
    foreach ($ma2 in @('A1','A2','A3','A4','A5','A6')) {
      $c = $sh.Range($ma2)
      $v = $c.Value2
      $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
      $型 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
      $ゼロか = '(★窓2が 打てません★)'
      # ★他の 道具と ★同じ 書き方★（`'=(' + マス + ')=0'`）★
      try { $w = $sh.Range('H' + $窓行); $w.Formula2 = '=(' + $ma2 + ')=0'; $ゼロか = [string]$w.Value2 } catch { }
      $窓行 = $窓行 + 1
      $行.Add($ma2 + "`t" + $値 + "`t" + $型 + "`t" + $ゼロか)
    }

    $bk.Close($false)
    $bk = $null
    $sh = $null
  }

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ('★書いた ... ' + $出 + '★')
} finally {
  $c = $null; $w = $null; $sh = $null
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
