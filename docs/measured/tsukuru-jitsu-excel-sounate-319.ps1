# tsukuru-jitsu-excel-sounate-319.ps1
#   -- ★台が 持つ 名前の うち まだ 聞いて いない 319個を 実Excel に 書かせる★（61）（2026-09-20）
#
#  ★★なぜ★★
#    ・穴は これまで ★4個★（WRAPROWS / WRAPCOLS / MODE.MULT / FORMULATEXT）
#    ・★4個とも 「見立てでは そこに 無い はず」の 所から 出ました★
#    ⇒★見立てで 絞るのを やめます＝全部 聞きます★
#    ⇒記憶「★探す 字を 決めた 時点で 答えが 決まる★」
#
#  ★★分母★★ `docs/measured/dai-namae-zenbu-2026-09-20.txt`（Exally1 が 数えた 376個）
#    ★「聞いた」欄が 「まだ」の 物だけ★ ⇒★319個★
#    ★断り★ ... この 名簿は `lib/shiki-kansuu.js` と `lib/shiki-tsunagi.js` の 2本だけ。
#               ★皮（`lib/formula-extra-plug.js`）は 数えて いません★（Exally1 の 断り）
#
#  ★★式の 形を どう 決めるか（★ここが 肝★）★★
#    ★関数ごとに 正しい 引数が 違います★＝★1つの 形では 打てません★
#    ⇒★梯子で 上から 試して ★最初に 通った 形★を 使います★
#    ⇒★通らなければ 「打てません」と 書きます★（★黙って 埋めない★）
#    ★★「通った」の 意味に 気を 付ける★★
#      ・★実Excel が 知らない 名前でも 「通り」ます★（＝決めた 名前の 参照と 見なされる）
#        ⇒その 時 値は ★#NAME?★／式は ★裸★で 入ります
#      ・★実Excel が 知って いる 名前は 印（`_xlfn.` 等）が 付いて 入ります★
#      ⇒★だから 値が #NAME? か どうかも 一緒に 書きます★
#      ⇒★#NAME? の 物は 「★うちの 台に 在って 実Excel に 無い★」かも しれません★
#        ＝★お客さんが 打てて 相手が 開けない＝今まで 数えて いません★（Exally1 の ⑤）
#
#  ★★作る 物★★ `%TEMP%\exally-jitsu-excel-sounate.xlsx`（★1本の 名★）
#    ＋ `docs/measured/golden-sounate-uchikata-2026-09-20.tsv`（★どのマスに 何を 打ったか★）
#    ⇒★印（`_xlfn.` 等）は 生の 字から 別の 道具で 読み、この 紙と 突き合わせます★
#
#  ★★材料と 置き場★★
#    A1:A6 = 1 / 2 / 2 / 3 / 3 / 4      B1:B3 = 10 / 20 / 30
#    E1:F2 = 1 / 2 / 3 / 4              D1 = "abc"
#    ★式は H 列に 8行 おき★（H1 H9 H17 ...）＝★縦に 7マス 溢れても ぶつかりません★
#    ★横に 溢れる 物★は I J K ... へ 伸びます＝★H 列だけ 使う ので 空いて います★
#
#  ★門★
#    ①貝殻が powershell.exe（5.1）（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③名簿が 読めなければ 走らない（exit 6）
#    ④★外へ 出る 6個は 名簿から 外す★（数を 出します）
#    ⑤★打てた 数が 0なら 落とす（exit 5）★＝★空振りを 緑に しない★
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出す先 = Join-Path $env:TEMP 'exally-jitsu-excel-sounate.xlsx'
$許す名 = 'exally-jitsu-excel-sounate.xlsx'
$打ち方の紙 = Join-Path $ここ 'golden-sounate-uchikata-2026-09-20.tsv'
$名簿 = Join-Path $ここ 'dai-namae-zenbu-2026-09-20.txt'

if ((Split-Path $出す先 -Leaf) -ne $許す名) { Write-Host '★★書いて よい 名は 1本だけです★★'; exit 7 }

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

if (-not (Test-Path $名簿)) { Write-Host ('★★名簿が 在りません ... ' + $名簿 + '★★'); exit 6 }

# ══ ★名簿を 読む★（★「まだ」の 物だけ★） ══
$まだ = New-Object System.Collections.Generic.List[string]
$外へ出る = 'WEBSERVICE', 'STOCKHISTORY', 'TRANSLATE', 'DETECTLANGUAGE', 'IMAGE', 'RTD'
$外した = 0
foreach ($l in [System.IO.File]::ReadAllLines($名簿)) {
  if ($l -eq '' -or $l.StartsWith('#')) { continue }
  $c = $l.Split("`t")
  if ($c.Count -lt 3) { continue }
  if ($c[0] -eq '名前') { continue }
  if ($c[2].Trim() -ne 'まだ') { continue }
  if ($外へ出る -contains $c[0].Trim()) { $外した++; continue }
  $まだ.Add($c[0].Trim())
}
Write-Host ('★名簿から 読んだ ... ' + $まだ.Count + '個★（★外へ 出る 6個の うち ' + $外した + '個を 外しました★）')
if ($まだ.Count -eq 0) { Write-Host '★★1個も 在りません★★'; exit 6 }

# ══ ★式の 形の 梯子★（★上から 試して 最初に 通った 物を 使う★） ══
#    ★よく 当たる 形を 上に 置きます★（★下まで 行くほど 遅い★）
$梯子 = @(
  ('(A1:A3)'), ('(A1)'), ('(1)'), ('()'), ('(A1,A2)'), ('(A1:A3,1)'),
  ('(1,1)'), ('("abc")'), ('(A1:A3,A1:A3)'), ('(1,A1:A3)'), ('(1,1,1)'),
  ('("abc",1)'), ('("abc","b")'), ('(TRUE)'), ('(1,0,A1:A6)'), ('(A1:A3,1,1)'),
  ('(D1)'), ('(E1:F2)'), ('(1,1,1,1)'), ('(A1:A6,A1:A6)'),
  # ★★2026-09-20 足し ── ★1回目に 5個 打てませんでした★★
  #   `CUMIPMT` `CUMPRINC` `MAXIFS` `MINIFS` `SUMIFS`
  #   ⇒★実Excel の 話では なく ★梯子に その 形が 無かった★＝私の 道具の 穴★
  ('(A1:A6,A1:A6,">2")'),        # SUMIFS / MAXIFS / MINIFS（★3つ組★）
  ('(0.01,12,1000,1,1,0)'),      # CUMIPMT / CUMPRINC（★6つ★）
  ('(0.01,12,1000,1,1)'),
  ('(A1:A3,A1:A3,A1:A3)'),
  ('(1,1,1,1,1)'), ('(1,1,1,1,1,1)')
)
Write-Host ('★式の 形の 梯子 ... ' + $梯子.Count + '通り★')

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null; $c = $null
$打てた = 0
$打てない = 0
$なまえ無し = 0
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $xl.ScreenUpdating = $false
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

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★319個を 実Excel に 打たせた＝どのマスに 何を 打ったか★（61）（2026-09-20）')
  $行.Add('# ★分母★ ... docs/measured/dai-namae-zenbu-2026-09-20.txt の 「まだ」')
  $行.Add('# ★式の 形★ ... 梯子を 上から 試して ★最初に 通った 物★')
  $行.Add('# ★#NAME? ＝ 実Excel が その 名前を 知らない★（うちの 台には 在る）')
  $行.Add('# マス' + "`t" + '名前' + "`t" + '打った式' + "`t" + '値' + "`t" + '出る字')

  $時計 = [Diagnostics.Stopwatch]::StartNew()
  $行数 = 1
  foreach ($な in $まだ) {
    $ma = 'H' + $行数
    $入った = $false
    $使った = ''
    foreach ($か in $梯子) {
      $f = '=' + $な + $か
      try {
        $sh.Range($ma).Formula2 = $f
        $入った = $true
        $使った = $f
        break
      } catch { }
    }
    if ($入った) {
      $c = $sh.Range($ma)
      $v = $c.Value2
      $値 = if ($null -eq $v) { '(kara)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
      $字 = [string]$c.Text
      if ($字 -eq '#NAME?') { $なまえ無し++ }
      $打てた++
      $行.Add($ma + "`t" + $な + "`t" + $使った + "`t" + $値 + "`t" + $字)
    } else {
      $打てない++
      $行.Add($ma + "`t" + $な + "`t" + '(★打てません★)' + "`t" + '' + "`t" + '')
    }
    $行数 = $行数 + 8
  }
  $時計.Stop()
  $行.Add('# ★打てた ' + $打てた + '個 ／ 打てない ' + $打てない + '個 ／ その うち #NAME? ' + $なまえ無し + '個★')
  $行.Add('# ★かかった 秒 ... ' + [math]::Round($時計.Elapsed.TotalSeconds, 1) + '★')

  Write-Host ''
  Write-Host ('★★打てた ' + $打てた + '個 ／ 打てない ' + $打てない + '個★★')
  Write-Host ('★その うち #NAME?（★実Excel が 知らない★） ... ' + $なまえ無し + '個★')
  Write-Host ('★かかった 秒 ... ' + [math]::Round($時計.Elapsed.TotalSeconds, 1) + '★')

  [System.IO.File]::WriteAllText($打ち方の紙, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ('★書いた ... ' + $打ち方の紙 + '★')

  if (Test-Path $出す先) { Remove-Item $出す先 -Force }
  $bk.SaveAs($出す先, 51)
  $bk.Close($false)
  $bk = $null
  if ($打てた -eq 0) { Write-Host '★★1個も 打てて いません★★'; exit 5 }
} finally {
  # ★★掴んだ物 全部 $null★★
  $c = $null
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
if (Test-Path $出す先) {
  $x2 = Get-Item $出す先
  $h = (Get-FileHash $出す先 -Algorithm SHA256).Hash.ToLower()
  Write-Host ''
  Write-Host ('★★作りました★★ ... ' + $出す先)
  Write-Host ('  ★大きさ★ ' + $x2.Length + ' バイト ／ ★sha256★ ' + $h)
}
