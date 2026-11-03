# toru-tsurarete-kawaru-kazu.ps1
#   -- ★1マス 打つと 実Excel では 何か所 つられて 変わるか★（103）（2026-09-22）
#
#  ★★なぜ★★
#    Exally1 が ア②の 1文を 出しました
#      「★4月!C4 を 1 から 99 に しました。つられて 4か所が 変わります★」
#    ⇒★この 数は 台（うちの 計算）の 数★です。
#    ⇒★実Excel でも 同じかを 割ります★（★お客さんに 見せる 字★なので）
#
#  ★★数え方（★先に 書きます★）★★
#    ①本の ★全部の 板の 使って いる マス★の 値を 控える
#    ②`4月!C4` に 99 を 打つ ⇒ ★計算させる★
#    ③もう一度 控える ⇒ ★変わった マスを 数える★（★打った マス 自身は 除く★）
#    ★保存しません★（読むだけで 開いて、閉じる 時は 捨てる）
#
#  ★★先に 断って おく★★
#    ・★「4か所」の 数え方が 相手と 同じとは 限りません★
#      ＝相手は ★台の 数★／私は ★実Excel の 数★
#    ・★式の 在る マスだけ 数える★か ★値も 含める★かで 変わります
#      ⇒★両方 出します★（★どちらか 1つに しない★）
#    ・★同じ 板／別の 板★も 分けて 出します
#
#  ★門★
#    ①貝殻が 5.1（exit 8）／②走らせる 前の Excel が 0個（exit 3）／③材料が 在る（exit 4）
#    ④★打つ 前の 値が 待った 物か★（exit 5）
#
#  使い方:
#    powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>
#      [-道 <xlsb>] [-板 4月] [-マス C4] [-打つ 99] [-前の値 1] [-印 <sha256>]

param(
  [string]$道 = '',
  [string]$板 = '4月',
  [string]$マス = 'C4',
  [string]$打つ = '99',
  [string]$前の値 = '1',
  [string]$印 = ''
)

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$repo = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
if (-not $道) { $道 = Join-Path $repo ('tests' + [string][char]92 + 'fixtures' + [string][char]92 + 'cross-sheet-sample.xlsb') }
if (-not (Test-Path -LiteralPath $道)) { Write-Host ('★★在りません★★ ' + $道); exit 4 }
$h = (Get-FileHash $道 -Algorithm SHA256).Hash.ToLower()
Write-Host ('★見る★ ' + $道 + ' ／ ' + (Get-Item $道).Length + ' バイト ／ sha256 ' + $h)
if ($印) {
  if ($h -ne $印.ToLower()) { Write-Host ('★★材料が 違います★★ 待ち ' + $印); exit 2 }
  Write-Host '★印は 渡された 物と 合って います★'
} else { Write-Host '★★印を 渡されて いません＝すり替わりを 見て いません★★' }

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$xl = New-Object -ComObject Excel.Application
$bk = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Open($道, $false, $true)

  # ── ★①控える★（★式の 在る マスか どうかも 一緒に★）
  $控え = @{}
  $式か = @{}
  $零 = @{}
  $型表 = @{}
  $マスの数 = 0
  for ($i = 1; $i -le [int]$bk.Sheets.Count; $i++) {
    $sh = $bk.Sheets.Item($i)
    $な = [string]$sh.Name
    $used = $sh.UsedRange
    $行数 = [int]$used.Rows.Count
    $列数 = [int]$used.Columns.Count
    $r0 = [int]$used.Row
    $c0 = [int]$used.Column
    for ($r = 0; $r -lt $行数; $r++) {
      for ($c = 0; $c -lt $列数; $c++) {
        $cell = $sh.Cells.Item($r0 + $r, $c0 + $c)
        $v = $cell.Value2
        if ($null -eq $v) { $cell = $null; continue }
        $k = $な + '!' + [string]$cell.Address($false, $false)
        $控え[$k] = [string]$v
        $式か[$k] = ([string]$cell.HasFormula -eq 'True')
        # ══ ★★「0」を 1つの 窓だけで 取らない★★ ══（`tests/monosashi-mado.test.mjs`）
        #   `.Value2` は ★0で ない 値に 0 を 返します★（実Excel の 見せ方の 側）
        #   ⇒★2つ目の 窓＝`=(そのマス)=0` の 真偽★ ／ ★3つ目＝型★
        #   ★本は 1マスも 触りません★＝`Application.Evaluate` は ★書き込みません★
        #     （マスに 打つと ★使って いる 範囲が 広がって★ 前後の 比べが 狂います）
        $式 = '=(' + "'" + $な + "'!" + [string]$cell.Address($false, $false) + ')=0'
        $窓 = '(hakarenai)'
        try { $窓 = [string]$xl.Evaluate($式.Substring(1)) } catch { $窓 = '(utenai)' }
        $零[$k] = $窓
        $型 = '(kara)'
        if ($v -is [double]) { $型 = 'Double' }
        elseif ($v -is [string]) { $型 = 'String' }
        elseif ($v -is [bool]) { $型 = 'Bool' }
        $型表[$k] = $型
        $マスの数++
        $cell = $null
      }
    }
    $used = $null
    $sh = $null
  }
  Write-Host ('★控えた マス★ ' + $マスの数 + '個（板 ' + [string]$bk.Sheets.Count + '枚）')

  # ── ★②打つ 前の 値を 確かめる★
  $的 = $板 + '!' + $マス
  if (-not $控え.ContainsKey($的)) { Write-Host ('★★' + $的 + ' が 控えに 在りません★★'); exit 5 }
  Write-Host ('★' + $的 + ' の 前の 値★ ' + $控え[$的] + '（待ち ' + $前の値 + '）')
  if ($控え[$的] -ne $前の値) { Write-Host '★★前の 値が 待った 物と 違います★★'; exit 5 }

  # ── ★③打つ★
  $sh2 = $bk.Sheets.Item($板)
  $sh2.Range($マス).Value2 = [double]$打つ
  $xl.CalculateFullRebuild()
  $sh2 = $null

  # ── ★④数える★
  $変わった = New-Object System.Collections.Generic.List[string]
  for ($i = 1; $i -le [int]$bk.Sheets.Count; $i++) {
    $sh = $bk.Sheets.Item($i)
    $な = [string]$sh.Name
    $used = $sh.UsedRange
    $行数 = [int]$used.Rows.Count
    $列数 = [int]$used.Columns.Count
    $r0 = [int]$used.Row
    $c0 = [int]$used.Column
    for ($r = 0; $r -lt $行数; $r++) {
      for ($c = 0; $c -lt $列数; $c++) {
        $cell = $sh.Cells.Item($r0 + $r, $c0 + $c)
        $v = $cell.Value2
        $k = $な + '!' + [string]$cell.Address($false, $false)
        # ★★2026-09-22 ── ★1回目と 2回目で 数える 物を 揃える★★
        #   1回目は ★空の マスを 飛ばして います★（`$null` で `continue`）
        #   ⇒2回目で 空の マスも 見ると ★零が (kara) ⇒ True に 見えて★
        #     ★空の マス 71個が 「変わった」に なりました★（★嘘の 76個★）
        #   ⇒★元も 今も 空なら 飛ばします★（★窓を 足した 時ほど 揃いが 崩れる★）
        if ($null -eq $v -and -not $控え.ContainsKey($k)) { $cell = $null; continue }
        $今 = if ($null -eq $v) { '(kara)' } else { [string]$v }
        $元 = if ($控え.ContainsKey($k)) { $控え[$k] } else { '(kara)' }
        # ★2つ目の 窓と 型も 見ます★（★値が 同じでも 零や 型が 変われば 変わった★）
        $式2 = '=(' + "'" + $な + "'!" + [string]$cell.Address($false, $false) + ')=0'
        $窓2 = '(hakarenai)'
        try { $窓2 = [string]$xl.Evaluate($式2.Substring(1)) } catch { $窓2 = '(utenai)' }
        $型2 = '(kara)'
        if ($v -is [double]) { $型2 = 'Double' }
        elseif ($v -is [string]) { $型2 = 'String' }
        elseif ($v -is [bool]) { $型2 = 'Bool' }
        $零元 = if ($零.ContainsKey($k)) { $零[$k] } else { '(kara)' }
        $型元 = if ($型表.ContainsKey($k)) { $型表[$k] } else { '(kara)' }
        if ($今 -ne $元 -or $窓2 -ne $零元 -or $型2 -ne $型元) {
          # ★★`(if ...)` を 式の 中に 書かない★★（5.1 では ★走らせた 時に★ 落ちます
          #   ＝`ParseFile` は 通します＝★読めた は 動く では ない★）
          $しるし = '  （値）'
          if ($式か.ContainsKey($k) -and $式か[$k]) { $しるし = '  （式）' }
          $変わった.Add($k + '  ' + $元 + ' ⇒ ' + $今 + $しるし + '  ／ 零 ' + $零元 + ' ⇒ ' + $窓2 + ' ／ 型 ' + $型元 + ' ⇒ ' + $型2)
        }
        $cell = $null
      }
    }
    $used = $null
    $sh = $null
  }

  Write-Host ''
  Write-Host ('★★変わった マス★★ ' + $変わった.Count + '個（★打った ' + $的 + ' を 含む★）')
  foreach ($x in $変わった) { Write-Host ('  ' + $x) }
  $つられ = $変わった.Count - 1
  Write-Host ''
  Write-Host ('★★つられて 変わった マス ＝ ' + $つられ + '個★★（打った マスを 除く）')
  $別板 = 0
  foreach ($x in $変わった) { if (-not $x.StartsWith($板 + '!')) { $別板++ } }
  Write-Host ('  ★同じ 板★ ' + ($変わった.Count - $別板 - 1) + '個 ／ ★別の 板★ ' + $別板 + '個')

  $bk.Close($false); $bk = $null
} finally {
  if ($null -ne $bk) { $bk.Close($false) }
  $bk = $null
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $xl = $null
  $t = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 300)) { Start-Sleep -Milliseconds 250 }
  $t.Stop()
  Write-Host ('★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒で 消えました★')
}
