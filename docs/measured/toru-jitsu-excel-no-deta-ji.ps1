# toru-jitsu-excel-no-deta-ji.ps1
#   -- ★実Excel が ★実際に 出して いる 字★を 全マスぶん 取る★（120）（2026-09-25）
#
#  ★★なぜ★★
#    司さんの 決め イ「★全く同じように表示できるようにしろや★」
#    ⇒★突き合わせるには 「実Excel が 出して いる 字」が 要ります★
#    ⇒★私は 09-25 に 借り物（SheetJS）の `w` を 物差しに して ★3,403個 違う★と 出しました★
#      ＝★但し 借り物の `w` は 曜日 `aaa` を 出せて いませんでした★
#      ＝★Exally の 方が 合って いる 所が 在った★
#    ⇒★物差しを 実Excel に 取り替えます★
#
#  ★★112（`toru-jitsu-excel-no-idageta.ps1`）と 何が 違うか★★
#    112 ... ★`####` の マスだけ★ 数える
#    120 ... ★全マスの `.Text` を 書き出す★（突き合わせの 物差しに する為）
#    ⇒★COM の 作りは 112 と 同じ★（★5.1／掴んだ物 全部 $null／Close($false)★）
#
#  ★★出し方（★ここが 大事★）★★
#    ・★マスの 字は 画面に 出しません★＝★`-出す先` の ファイルにだけ 書きます★
#    ・★画面に 出すのは 数だけ★（板・マスの 数・空の 数・`####` の 数）
#    ・★`-出す先` は scratchpad を 指して ください★＝★repo に 入れては いけません★
#      （★お客さんの 数字が そのまま 入って います★）
#    ・形 ... `板の名<TAB>行<TAB>列<TAB>出た字`（行・列は ★0から★＝Exally と 同じ）
#      ＝★板の 名は 要ります★（突き合わせの 鍵）＝★だから 報告には 貼りません★
#
#  ★★読むだけ★★（`Workbooks.Open(道, $false, $true)` ＝ 読み取り専用／`Close($false)`）
#    ★1バイトも 書き戻しません★
#
#  ★★門★★
#    ①5.1（exit 8）／②材料が 在る（exit 4）／③印が 合う（exit 2）
#    ④走らせる 前の Excel が 0個（exit 3）★他の 席を 閉じません★
#    ⑤★出す先が 空で ない（exit 5）★／⑥★1行も 書けなかったら 赤（exit 6）★
#
#  使い方:
#    <この道具> -道 <xlsb> -出す先 <tsv> [-板数 0] [-印 <sha256>]
#      -板数 0 ＝★全部の 板★（既定）

param(
  [string]$道 = '',
  [string]$出す先 = '',
  [int]$板数 = 0,
  [string]$印 = ''
)

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }
if (-not $道 -or -not (Test-Path -LiteralPath $道)) { Write-Host '★★材料が 在りません★★'; exit 4 }
if (-not $出す先) { Write-Host '★★-出す先 を 渡して ください★★'; exit 5 }

$h = (Get-FileHash -LiteralPath $道 -Algorithm SHA256).Hash.ToLower()
Write-Host ('★見る★ ' + (Split-Path $道 -Leaf) + ' ／ sha256 ' + $h)
if ($印) {
  if ($h -ne $印.ToLower()) { Write-Host '★★材料が 違います★★'; exit 2 }
  Write-Host '★印は 合って います★'
}
Write-Host ('★出す先★ ' + $出す先)
Write-Host '  ★★この ファイルには お客さんの 数字が そのまま 入ります＝repo に 入れないで ください★★'

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$書 = $null
$xl = New-Object -ComObject Excel.Application
$bk = $null
try {
  $書 = New-Object System.IO.StreamWriter($出す先, $false, (New-Object System.Text.UTF8Encoding($false)))
  $書.NewLine = "`n"                       # ★書き戻しは LF★（記憶「書き戻しは必ずLF」）
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $xl.AutomationSecurity = 3
  $bk = $xl.Workbooks.Open($道, $false, $true)
  $枚 = [int]$bk.Sheets.Count
  # ★★`$x = if (...) {...}` は PowerShell 5.1 で ★式では ありません★★★
  #    ＝`ParseFile` は 通す 事が 在り、★走らせた 時に 落ちます★
  #    ＝記憶「★`(if ...)` は 式では ない★」を 私が 自分で 踏みました（2026-09-25）
  $みる = $枚
  if ($板数 -gt 0) { $みる = [math]::Min($枚, $板数) }
  Write-Host ('★板★ ' + $枚 + '枚（★' + $みる + '枚を 見ます★）')
  Write-Host ''
  $全マス = 0; $全空 = 0; $全井桁 = 0; $全行 = 0
  $食い違い = 0; $食い違いを出した = 0; $井桁を出した = 0
  $型の数 = @{}
  for ($i = 1; $i -le $みる; $i++) {
    $sh = $bk.Sheets.Item($i)
    $な = [string]$sh.Name
    $used = $sh.UsedRange
    $r0 = [int]$used.Row; $c0 = [int]$used.Column
    $rn = [int]$used.Rows.Count; $cn = [int]$used.Columns.Count
    $マス = 0; $空 = 0; $井桁 = 0
    for ($r = 0; $r -lt $rn; $r++) {
      for ($c = 0; $c -lt $cn; $c++) {
        $cell = $sh.Cells.Item($r0 + $r, $c0 + $c)
        $v = $cell.Value2
        $t = [string]$cell.Text
        # ══ ★★2つ目の 窓★★ ══（2026-09-25・見張り `tests/monosashi-mado.test.mjs` が 赤に した）
        #   ★何を 疑うか★ ... ★`.Value2` が null の マスを 「空」として 捨てて います★
        #                     ⇒★もし `.Text` が 空で ないのに null が 返るなら 字を 落とします★
        #   ★1つ目の 窓★ ... `.Value2` が null か
        #   ★2つ目の 窓★ ... ★`=(そのマス)=0` の 真偽★（`Evaluate`・★本は 1マスも 触りません★）
        #                     ＋ ★`.Value2` の 型★ ＋ ★`.Text` が 空か★
        #   ⇒★3つが 食い違った マスを 数えます★（★0件なら 「空」の 数は 信じられます★）
        if ($null -eq $v) {
          if ($t -ne '') {
            # ★★字が 出て いるのに 値が null★★＝★捨てては いけない マス★
            $食い違い++
            if ($食い違いを出した -lt 10) {
              # ★112 と 同じ 書き方に 揃えます★（見張りが 探す 形＝`')=0'`）
              $しき = '=(' + "'" + $な.Replace("'", "''") + "'!" + [string]$cell.Address($false, $false) + ')=0'
              $零 = '(hakarenai)'
              try { $零 = [string]$xl.Evaluate($しき.Substring(1)) } catch { $零 = '(utenai)' }
              Write-Host ('  ★★字が 出て いるのに Value2 が null★★ 板' + $i + ' ' + [string]$cell.Address($false, $false) + ' ／ =0 の 真偽 ' + $零 + ' ／ 字数 ' + $t.Length)
              $食い違いを出した++
            }
          }
          $空++; $cell = $null; continue
        }
        $マス++
        # ★★型も 一緒に 数えます★★（★「0」を 型 抜きで 信じない★）
        $かた = 'other'
        if ($v -is [double]) { $かた = 'Double' }
        elseif ($v -is [string]) { $かた = 'String' }
        elseif ($v -is [bool]) { $かた = 'Bool' }
        if (-not $型の数.ContainsKey($かた)) { $型の数[$かた] = 0 }
        $型の数[$かた] = $型の数[$かた] + 1
        if ($t -match '^#+$') {
          $井桁++
          # ★★`####` も 2つ目の 窓で 見ます★★（★112 と 同じ 手★）
          if ($井桁を出した -lt 10) {
            $しき2 = '=(' + "'" + $な.Replace("'", "''") + "'!" + [string]$cell.Address($false, $false) + ')=0'
            $零2 = '(hakarenai)'
            try { $零2 = [string]$xl.Evaluate($しき2.Substring(1)) } catch { $零2 = '(utenai)' }
            Write-Host ('  ★#### の マス★ 板' + $i + ' ' + [string]$cell.Address($false, $false) + ' ／ =0 の 真偽 ' + $零2 + ' ／ 型 ' + $かた + ' ／ 井桁 ' + $t.Length)
            $井桁を出した++
          }
        }
        # ★★行・列は 0から★★（Exally の `data['行,列']` と 同じ 数え方）
        $行0 = $r0 + $r - 1
        $列0 = $c0 + $c - 1
        # ★★タブと 改行は 逃がします★★（1行 1マスを 壊さない為）
        $字 = $t.Replace("`t", '\t').Replace("`r", '\r').Replace("`n", '\n')
        $名 = $な.Replace("`t", '\t').Replace("`r", '\r').Replace("`n", '\n')
        $書.WriteLine($名 + "`t" + $行0 + "`t" + $列0 + "`t" + $字)
        $全行++
        $cell = $null
      }
    }
    $いち = '  板 ' + $i.ToString().PadLeft(2) + ' ... 字の 在る マス ' + $マス.ToString().PadLeft(6)
    $いち = $いち + ' / 空 ' + $空.ToString().PadLeft(6) + ' / #### ' + $井桁.ToString().PadLeft(4)
    Write-Host $いち
    $全マス += $マス; $全空 += $空; $全井桁 += $井桁
    $used = $null; $sh = $null
  }
  Write-Host ''
  Write-Host ('★★合計★★ 字の 在る マス ' + $全マス + '個 ／ 空 ' + $全空 + '個 ／ ★`####` ' + $全井桁 + '個★')
  Write-Host ''
  Write-Host '★★2つ目の 窓★★'
  Write-Host ('  ★字が 出て いるのに Value2 が null ... ' + $食い違い + '個★')
  if ($食い違い -eq 0) { Write-Host '    ✓ ★「空」の 数は 信じられます★' }
  else { Write-Host '    ★★「空」として 捨てた 中に 字が 在ります＝数が 足りません★★' }
  $型の行 = ''
  foreach ($k in $型の数.Keys) { $型の行 = $型の行 + $k + ' ' + $型の数[$k] + '個 ／ ' }
  Write-Host ('  ★Value2 の 型★ ' + $型の行)
  Write-Host ('  ★`####` の 数★ ' + $全井桁 + '個（★上で 1つずつ =0 の 真偽と 型を 出して います★）')
  Write-Host ('★書いた 行★ ' + $全行 + '行')
  $bk.Close($false); $bk = $null
} finally {
  if ($null -ne $書) { $書.Flush(); $書.Close(); $書.Dispose(); $書 = $null }
  if ($null -ne $bk) { $bk.Close($false) }
  $bk = $null
  if ($null -ne $xl) {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    $xl = $null
  }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  $t2 = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t2.Elapsed.TotalSeconds -lt 300)) {
    Start-Sleep -Milliseconds 250
  }
  $t2.Stop()
  $残り = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
  if ($残り -eq 0) { Write-Host ('★Excel は ' + [math]::Round($t2.Elapsed.TotalSeconds, 1) + '秒で ★消えました★★') }
  else { Write-Host ('★★Excel は 消えませんでした★★ ／ 残り ' + $残り + '個') }
}

if (-not (Test-Path -LiteralPath $出す先)) { Write-Host '★★出せませんでした★★'; exit 6 }
$x = Get-Item -LiteralPath $出す先
# ★★`Get-Content | Measure-Object -Line` は 嘘を 返しました★★（2026-09-25 実測）
#    書いた 21,227行 に 対して ★20,190★ と 出ました（★1,037行 少ない★）
#    ＝BOM 無しの UTF-8 を 読み違えて 行を 繋げて いると 思われます（★因は 未測定★）
#    ⇒★改行の 数を バイトで 数えます★
$中身 = [System.IO.File]::ReadAllBytes($出す先)
$行数 = 0
foreach ($b in $中身) { if ($b -eq 10) { $行数++ } }
$中身 = $null
Write-Host ''
Write-Host ('★★出来た 物★★ ' + $x.Length.ToString('N0') + ' バイト ／ ' + $行数 + ' 行（★改行を バイトで 数えました★）')
if ($食い違い -ne 0) {
  Write-Host '★★字が 在る マスを 「空」として 捨てて います＝赤★★'
  exit 7
}
if ($行数 -ne $全行) {
  Write-Host ('★★書いた 行（' + $全行 + '）と 数えた 行（' + $行数 + '）が 合いません★★')
  exit 6
}
Write-Host '  ✓ ★書いた 行と 数えた 行が 同じ★'
if ($行数 -le 0) { Write-Host '★★1行も 書けて いません★★'; exit 6 }
Write-Host '  ⇒★次は これを 物差しに して Exally と 突き合わせます★'
Write-Host '     node docs/measured/kazoeru-hiraita-ato-keisan-ga-iruka.mjs "<xlsb>" --実Excel "<この tsv>"'
