# toru-tashita-ita-wo-excel-ni-hirakaseru.ps1
#   -- ★板を 足して 保存した 物を 実Excel に 開かせる★（97）（2026-09-21）
#
#  ★★なぜ★★
#    Exally1 が ア①（足した 板も 書き出す）を 書きました。★まだ 押して いません★。
#    ⇒★作る 人と 数える 人を 分ける★＝私が 実Excel に 開かせて 数えます。
#
#  ★★数える 物（Exally1 の 注文 ＋ 私から）★★
#    ①`Sheet2` が 在るか（A1=123 ／ B1=tashita）
#    ②★元の 判子・罫線・塗り・図形が 1つも 減って いないか★
#    ③元の 板（Sheet1）の 値が 1つも 変わって いないか
#    ④★`Open` が 投げないか★（93 の 測り）
#    ⑤★`calcChain` を 残した ままで 計算が 壊れて いないか★
#       ＝Sheet1 の `=SUM(A1:A3)` が ★4.25★ の まま か
#    ★私から 足す★
#    ⑥★元も 同じ やり方で 開いて 並べる★（★後だけ 見ても「減った」は 分からない★）
#    ⑦★板の 並び★（足した 板が 後ろに 入ったか）
#
#  ★★読むだけ★★（★1バイトも 書きません★・`SaveAs` を 1回も 呼びません）
#
#  ★門★
#    ①貝殻が 5.1（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③★2本 とも 在るか★（exit 4）／④★2本 とも 開けたか★（exit 5）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

param(
  [string]$前 = '',
  [string]$後 = '',
  [string]$前の印 = '',
  [string]$後の印 = '',
  [string]$見るマス = '',      # ★材料ごとに 違う★（カンマ区切り）
  [string]$繋げた = 'A5',      # ★繋げた マスの 左上★
  [string]$塗り = 'B1',        # ★塗った マス★
  [string]$線 = 'B2',          # ★線を 引いた マス★
  # ★★2枚目に 何を 書いたかは ★材料に よって 違います★★（2026-09-22 に 踏みました）
  #   `.xlsx` は A1=123 / B1='tashita'（字）／`.xlsb` は A1=123 / B1=2.5（数だけ）
  #   ⇒★決め打ちに すると 「相手の 直しの 赤」に 見えます★
  #   ⇒★待つ 値を 渡します★
  [string]$板2A1 = '123',
  [string]$板2B1 = 'tashita'
)

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

# ★★材料は 選べます★★（★手元の 決め打ちを 焼き込まない★＝記憶の 決まり）
#   既定は 1組目（飾り 8種＋溢れ 3種）。グラフ入りは `-前`/`-後` で 渡します。
$repo = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
if (-not $前) { $前 = Join-Path $repo ('tests' + [string][char]92 + 'fixtures' + [string][char]92 + 'kazari-hiraku3.xlsx') }
if (-not $後) { $後 = Join-Path $env:TEMP 'exally-tashita-ita.xlsx' }
foreach ($f in $前, $後) {
  if (-not (Test-Path -LiteralPath $f)) { Write-Host ('★★在りません★★ ' + $f); exit 4 }
  Write-Host ('★見る★ ' + $f + ' ／ ' + (Get-Item $f).Length + ' バイト ／ sha256 ' + (Get-FileHash $f -Algorithm SHA256).Hash.ToLower())
}

# ══ ★★材料の 印を 先に 突き合わせます★★ ══（2026-09-22 に 踏みました）
#   `%TEMP%` の 名は ★上書きされます★。実際に
#   `exally-tashita-ita.xlsx` の 中身が ★グラフの 物に 入れ替わって いました★
#   ⇒★気づかずに 走らせると 別の 材料どうしを 比べて ★嘘の 赤★が 出ます★
#   ⇒★渡された 印と 合わない なら 走りません★（`-前の印` / `-後の印`）
if ($前の印 -or $後の印) {
  $h1 = (Get-FileHash $前 -Algorithm SHA256).Hash.ToLower()
  $h2 = (Get-FileHash $後 -Algorithm SHA256).Hash.ToLower()
  if ($前の印 -and $h1 -ne $前の印.ToLower()) { Write-Host ('★★前の 材料が 違います★★ 待ち ' + $前の印 + ' ／ 実物 ' + $h1); exit 2 }
  if ($後の印 -and $h2 -ne $後の印.ToLower()) { Write-Host ('★★後の 材料が 違います★★ 待ち ' + $後の印 + ' ／ 実物 ' + $h2); exit 2 }
  Write-Host '★印は 渡された 物と 合って います★'
} else {
  Write-Host '★★印を 渡されて いません＝すり替わりを 見て いません★★'
}

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

# ★見る マス★（★元の 材料の 中身は `tsukuru-tameshi-hiraku3-kazari-to-kobore3.ps1` に 書いて あります★）
# ★★材料ごとに 見る 所が 違います★★（2026-09-22 に 踏みました）
#   グラフ入りの 材料を ★飾り入りの 材料の マス★で 測って いました。
#   ⇒`繋げた A5`＝False／`塗り B1`＝白／`線 B2`＝2 ... ★どれも 飾りの 無い マス★
#   ⇒★前後 とも 同じ なので 緑に なります★＝★何も 守って いない 緑★
#   ⇒記憶「★掃く 窓は 押す前に 隙間を 計算して から 置く★」
$既定の見るマス = 'A1,A2,A3,B1,C1,C2,C3,D1,D5,E1,A5,A10,B10,C10'
# ★★2026-09-22 ── ★`param` で `[string]` と 書いた 変数は ★後から 配列を 入れても 文字に 戻ります★★★
#   `$見るマス = @(...)` と 書いたのに ★中身は 1本の 長い 文字★に なり、
#   `Range('A1,B1,...')` を 引いて ★COM が 投げました★（0x800A03EC）。
#   ⇒★型の 縛りは その 変数に ずっと 付いて 回ります★
#   ⇒★別の 名前（縛りの 無い 物）に 入れます★
$マス指定 = $見るマス
if (-not $マス指定) { $マス指定 = $既定の見るマス }
$マス達 = @($マス指定.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
if ($マス達.Count -lt 2) { Write-Host ('★★見る マスが ' + $マス達.Count + '個＝割れて いません★★'); exit 6 }
Write-Host ('★見る マス★ ' + ($マス達 -join ' ') + '（' + $マス達.Count + '個）')
Write-Host ('★飾りを 見る 所★ 繋げた ' + $繋げた + ' ／ 塗り ' + $塗り + ' ／ 線 ' + $線)

$結果 = @{}
$開けた = 0
foreach ($組 in @(@('mae', $前), @('ato', $後))) {
  $札 = $組[0]; $道 = $組[1]
  Write-Host ''
  Write-Host ('★★' + $札 + '★★ ... ' + (Split-Path $道 -Leaf))
  $xl = New-Object -ComObject Excel.Application
  $bk = $null; $sh = $null
  $投げた = ''
  $出 = @{ 投げた = ''; 板 = @(); 図形 = 0; 値 = @{}; 零 = @{}; 型 = @{}; 繋げた = ''; 塗り = 0; 線 = 0 }
  try {
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    try { $bk = $xl.Workbooks.Open($道, $false, $true) }
    catch { $投げた = $_.Exception.Message }
    $出.投げた = $投げた
    if ($null -ne $bk) {
      $開けた++
      for ($i = 1; $i -le [int]$bk.Sheets.Count; $i++) { $出.板 += [string]$bk.Sheets.Item($i).Name }
      $sh = $bk.Sheets.Item(1)
      $出.図形 = [int]$sh.Shapes.Count
      # ★★グラフも Shapes に 入ります★★＝★分けずに 「図形 N個」と 書くと 嘘★
      $出.グラフ = 0
      try { $出.グラフ = [int]$sh.ChartObjects().Count } catch { $出.グラフ = -1 }
      $出.判子 = $出.図形 - $出.グラフ
      # ══ ★★1つの 窓だけで 「0」を 取らない★★ ══（`tests/monosashi-mado.test.mjs`）
      #   `.Value2` は ★0で ない 値に 0 を 返します★（実Excel の 見せ方の 側）
      #     `=0.1+0.2-0.3` ... `.Value2` ★0★ ／ `=(式)=0` ★False★
      #   さらに ★字の "0"★ も 同じ 顔を します（`=DEC2BIN(0.5)`）
      #   ⇒★3つの 窓で 取ります★
      #       ①`.Value2`
      #       ②★`=(そのマス)=0` の 真偽★（★本当に 0 か★）
      #       ③★型★（`-is [double]` ／ `-is [string]` ／ 空）
      #   ★打つ 先は 材料の 外★（BZ1000）＝★材料の マスには 1文字も 書きません★
      #   ★保存しません★（`Close($false)`）ので ★ファイルは 1バイトも 変わりません★
      foreach ($m in $マス達) {
        $c = $sh.Range($m)
        $v = $c.Value2
        $出.値[$m] = if ($null -eq $v) { '(kara)' } else { [string]$v }
        # ★②2つ目の 窓★
        $窓 = '(hakarenai)'
        try {
          $sh.Range('BZ1000').Formula2 = '=(' + $m + ')=0'
          $窓 = [string]$sh.Range('BZ1000').Value2
        } catch { $窓 = '(utenai)' }
        $出.零[$m] = $窓
        # ★③型★（★字の "0" と 数の 0 を 分ける★）
        $型 = '(kara)'
        if ($v -is [double]) { $型 = 'Double' }
        elseif ($v -is [string]) { $型 = 'String' }
        elseif ($v -is [bool]) { $型 = 'Bool' }
        elseif ($null -ne $v) { $型 = $v.GetType().Name }
        $出.型[$m] = $型
      }
      # ★打った 所を 片づけます★（★材料の 外だが 残さない★）
      $sh.Range('BZ1000').ClearContents() | Out-Null
      $出.繋げた = [string]$sh.Range($繋げた).MergeCells
      $出.塗り = [int]$sh.Range($塗り).Interior.Color
      $出.線 = [int]$sh.Range($線).Borders.Item(9).Weight
      # ★足した 板の 中身★
      if ([int]$bk.Sheets.Count -ge 2) {
        $s2 = $bk.Sheets.Item(2)
        $出.板2A1 = [string]$s2.Range('A1').Value2
        $出.板2B1 = [string]$s2.Range('B1').Value2
        $s2 = $null
      }
      $bk.Close($false); $bk = $null
    }
  } finally {
    $sh = $null
    if ($null -ne $bk) { $bk.Close($false) }
    $bk = $null
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    $xl = $null
    $t = [Diagnostics.Stopwatch]::StartNew()
    while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 300)) { Start-Sleep -Milliseconds 250 }
  }
  $結果[$札] = $出
  if ($投げた) { Write-Host ('  ★★投げました★★ ' + $投げた) }
  else {
    Write-Host ('  板 ....... ' + ($出.板 -join ' / '))
    Write-Host ('  図形 ..... ' + $出.図形 + '個（グラフ ' + $出.グラフ + ' ／ 判子 ' + $出.判子 + '）')
    Write-Host ('  繋げた ' + $繋げた + ' ' + $出.繋げた + ' ／ 塗り ' + $塗り + ' ' + $出.塗り + ' ／ 線 ' + $線 + ' ' + $出.線)
    if ($出.ContainsKey('板2A1')) { Write-Host ('  2枚目 ... A1=' + $出.板2A1 + ' ／ B1=' + $出.板2B1) }
  }
}
if ($開けた -ne 2) { Write-Host ''; Write-Host ('★★開けたのは ' + $開けた + ' / 2 本★★'); exit 5 }

# ══ ★★並べて 数えます★★ ══
Write-Host ''
Write-Host '★★★前と 後を 並べる★★★'
$赤 = 0
$見 = {
  param($札, $前値, $後値, $待ち)
  $合 = if ($待ち -eq 'onaji') { $前値 -eq $後値 } else { $後値 -eq $待ち }
  $印 = if ($合) { 'ok  ' } else { '★赤★' }
  if (-not $合) { $script:赤++ }
  Write-Host ('  ' + $印 + ' ' + $札.PadRight(22) + ' 前 ' + ([string]$前値).PadRight(14) + ' 後 ' + [string]$後値)
}
& $見 '①2枚目が 在る' $結果['mae'].板.Count $結果['ato'].板.Count 2
& $見 ('①2枚目 A1（待ち ' + $板2A1 + '）') '-' $結果['ato'].板2A1 $板2A1
& $見 ('①2枚目 B1（待ち ' + $板2B1 + '）') '-' $結果['ato'].板2B1 $板2B1
& $見 '②図形の 数（全部）' $結果['mae'].図形 $結果['ato'].図形 'onaji'
& $見 '②うち グラフ' $結果['mae'].グラフ $結果['ato'].グラフ 'onaji'
& $見 '②うち 判子' $結果['mae'].判子 $結果['ato'].判子 'onaji'
& $見 ('②繋げた マス ' + $繋げた) $結果['mae'].繋げた $結果['ato'].繋げた 'onaji'
& $見 ('②塗り ' + $塗り) $結果['mae'].塗り $結果['ato'].塗り 'onaji'
& $見 ('②線 ' + $線) $結果['mae'].線 $結果['ato'].線 'onaji'
foreach ($m in $マス達) {
  & $見 ('③値 ' + $m) $結果['mae'].値[$m] $結果['ato'].値[$m] 'onaji'
  # ★★値だけ 揃っても 「同じ」では ありません★★
  #   ＝★前後 とも 空なら 0 対 0 で 揃って 見えます★（★偽の 緑★）
  #   ⇒★2つ目の 窓と 型も 並べます★
  & $見 ('③零 ' + $m) $結果['mae'].零[$m] $結果['ato'].零[$m] 'onaji'
  & $見 ('③型 ' + $m) $結果['mae'].型[$m] $結果['ato'].型[$m] 'onaji'
}
& $見 '④投げなかった' '(なし)' $結果['ato'].投げた ''
& $見 '⑦1枚目の 名' $結果['mae'].板[0] $結果['ato'].板[0] 'onaji'

Write-Host ''
Write-Host ('★★赤 ... ' + $赤 + '件★★')
if ($赤 -ne 0) { exit 1 }
