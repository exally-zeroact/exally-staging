# toru-kobore-fusagi-wo-excel-ni-hirakaseru.ps1
#   -- ★塞がれた 溢れ（#SPILL!）を 書き出した 物を 実Excel に 開かせる★（㊺）（2026-09-20）
#
#  ★★なぜ 要るか★★
#    ㊹で 測れたのは ★ちゃんと 溢れた 物★だけでした。
#    ⇒★画面で 既に `#SPILL!` に なって いる 物★は ★1本も 測って いません★
#    ⇒生の 字を 読むと ★誤りが `t="e"`（誤り）では なく `t="str"`（文字列）★で 書かれて います:
#        `<c r="D1" t="str"><f>_xlfn.SEQUENCE(3)</f><v>#SPILL!</v></c>`
#    ⇒★「型が 違う」は 事実／「客が 困る」は まだ 見立て★
#      ＝★実Excel が 開いた 時に 何に なるか で 決まります★
#
#  ★★開く 物★★ `%TEMP%\exally-kakidashi-spill.xlsx`（★1本の 名★）
#    D1 ... 塞がれた 溢れ（D2 に 9 が 先に 在る）
#    F1 ... ★対照★＝同じ 式で 塞がれて いない（F1:F3）
#
#  ★★4つ 見ます★★
#    ⑴★開いた 瞬間★ D1 は 文字列か 誤りか（★型を 見ます／まだ 何も 書きません★）
#    ⑵`HasArray`（塞がれた 物は どう 読まれるか）
#    ⑶F1 が ちゃんと 動く並びか（★対照★）
#    ⑷★D2 の 9 を 消したら D1 は 溢れ直すか★
#
#  ★★⑴を 先に 読む 訳★★
#    ★式を 1つ 打つと そこで 計算し直しが 起きます★
#    ⇒`ISERROR` を 打つ 前に ★開いた ままの 姿★を 取って おかないと
#      ★「開いた 瞬間」が 二度と 測れません★
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②開く 名が 1文字でも 違えば 走らない（exit 7）
#    ③ファイルが 無ければ 走らない（exit 6）／④走らせる 前の Excel が 0個（exit 3）
#    ★保存しません★（`$bk.Close($false)`／`SaveAs` は 1文字も 在りません）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
# ★★2026-09-20 ── ★同じ 名前の ファイルの 「直す 前」と 「直した 後」を 別の 紙に します★★
#   ＝`%TEMP%` の 名は 1本の まま（門を 緩めない）／★紙が 上書きされると 前が 消えます★
#   ⇒`-mae-`（`cm` 無し）／`-ato-`（`cm="1" ref="D1:D1"` 付き）
$出 = Join-Path $ここ 'golden-kobore-fusagi-excel-ato-2026-09-20.tsv'

$許す名 = 'exally-kakidashi-spill.xlsx'
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

$見るマス = @('D1', 'D2', 'D3', 'D4', 'E1', 'F1', 'F2', 'F3', 'F4', 'G1')

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null; $c = $null; $w = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Open($開く, 0, $true)
  $sh = $bk.Sheets.Item(1)

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★塞がれた 溢れ（#SPILL!）を 書き出した 物を 実Excel に 開かせた★（㊺）（2026-09-20）')
  $行.Add('# ★開いた 物★ ... ' + $開く + '（' + (Get-Item $開く).Length + ' バイト）')
  $行.Add('# ★読むだけ★（保存して いません）')
  $行.Add('# ★どの Excel か★ ... 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★どの 貝殻か★ ... PowerShell ' + $版.ToString())
  $行.Add('# ★生の 字★ D1 `<c r="D1" t="str" cm="1"><f t="array" ref="D1:D1">_xlfn.SEQUENCE(3)</f><v>#SPILL!</v></c>`')
  $行.Add('# ★前（-mae- の 紙）★ D1 は `cm` も `ref` も 無く ⇒ ★値 1（Double）★ でした')
  $行.Add('# ★生の 字★ F1 `<c r="F1" cm="1"><f t="array" ref="F1:F3">_xlfn.SEQUENCE(3)</f><v>1</v></c>`')

  # ═══ ★★①開いた 瞬間（★まだ 1字も 書いて いません★）★★ ═══
  $行.Add('#')
  $行.Add('# ★★①開いた 瞬間★★（★まだ 何も 打って いません★）')
  $行.Add('# マス' + "`t" + '値' + "`t" + '出る字' + "`t" + '式' + "`t" + '型' + "`t" + '溢れの一部か')
  $控え = New-Object System.Collections.Generic.List[string]
  foreach ($ma in $見るマス) {
    $c = $sh.Range($ma)
    $v = $c.Value2
    $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    $型 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
    $出字 = [string]$c.Text
    $式2 = [string]$c.Formula
    $一部か = '(?)'
    try { $一部か = [string]$c.HasArray } catch { }
    $控え.Add($ma + '=' + $値 + ':' + $式2)
    $行.Add($ma + "`t" + $値 + "`t" + $出字 + "`t" + $式2 + "`t" + $型 + "`t" + $一部か)
  }

  # ═══ ★★②文字列か 誤りかを 直に 訊く★★ ═══
  #   ★★ここから 先は 計算し直しが 起きます★★（①を 先に 取って ある 訳）
  $行.Add('#')
  $行.Add('# ★★②文字列か 誤りかを 直に 訊く★★（★ここから 計算し直しが 起きます★）')
  $問い = @(
    @{ 名 = 'J1'; 式 = '=ISERROR(D1)' },
    @{ 名 = 'J2'; 式 = '=ISTEXT(D1)' },
    @{ 名 = 'J3'; 式 = '=ISNUMBER(D1)' },
    @{ 名 = 'J4'; 式 = '=ISERROR(F1)' },
    @{ 名 = 'J5'; 式 = '=ISTEXT(F1)' }
  )
  foreach ($q in $問い) {
    $答 = '(★打てません★)'
    try { $w = $sh.Range($q.名); $w.Formula2 = $q.式; $答 = [string]$w.Value2 } catch { $答 = '★投げました★ ' + $_.Exception.Message }
    $行.Add('# ' + $q.式 + "`t" + $答)
  }
  $行.Add('# ★D1 を もう 一度★' + "`t" + [string]$sh.Range('D1').Value2 + "`t" + [string]$sh.Range('D1').Text)

  # ═══ ★★③邪魔（D2）を 消したら 溢れ直すか★★ ═══
  $行.Add('#')
  $行.Add('# ★★③邪魔（D2）を 消したら D1 は 溢れ直すか★★')
  $判じ = '(★消せません★)'
  try {
    $sh.Range('D2').ClearContents() | Out-Null
    $束 = New-Object System.Collections.Generic.List[string]
    foreach ($ma in $見るマス) {
      $c = $sh.Range($ma)
      $v = $c.Value2
      $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
      $束.Add($ma + '=' + $値 + ':' + [string]$c.Formula)
    }
    $判じ = if (($束 -join '|') -eq ($控え -join '|')) { '★1マスも 変わりません★' } else { 'ok（★変わりました★）' }
  } catch { $判じ = '★投げました★ ' + $_.Exception.Message }
  $行.Add('# ★消した 結果★ ... ' + $判じ)
  $行.Add('# マス' + "`t" + '値' + "`t" + '出る字' + "`t" + '式' + "`t" + '型' + "`t" + '溢れの一部か')
  foreach ($ma in $見るマス) {
    $c = $sh.Range($ma)
    $v = $c.Value2
    $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    $型 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
    $出字 = [string]$c.Text
    $式2 = [string]$c.Formula
    $一部か = '(?)'
    try { $一部か = [string]$c.HasArray } catch { }
    $行.Add($ma + "`t" + $値 + "`t" + $出字 + "`t" + $式2 + "`t" + $型 + "`t" + $一部か)
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
    foreach ($ma2 in $見るマス) {
      $c = $sh.Range($ma2)
      $v = $c.Value2
      $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
      $型 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } elseif ($v -is [bool]) { 'Boolean' } else { 'Other' }
      $ゼロか = '(★窓2が 打てません★)'
      # ★他の 道具と ★同じ 書き方★（`'=(' + マス + ')=0'`）★
      try { $w = $sh.Range('N' + $窓行); $w.Formula2 = '=(' + $ma2 + ')=0'; $ゼロか = [string]$w.Value2 } catch { }
      $窓行 = $窓行 + 1
      $行.Add($ma2 + "`t" + $値 + "`t" + $型 + "`t" + $ゼロか)
    }

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ('★書いた ... ' + $出 + '★')
  # ★★保存しません★★
  $bk.Close($false)
} finally {
  $c = $null; $w = $null; $sh = $null; $bk = $null
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
