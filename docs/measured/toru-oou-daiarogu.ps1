# toru-oou-daiarogu.ps1 — ★実Excel は 窓を 開けた 時 シートを 暗くするか★（2026-09-10）
#
#  ★★なぜ 測るか★★
#    うちは ★窓を 1つ 開けるたびに 画面 全体を 45〜55% 黒く★します（45か所）。
#      book.html … rgba(0,0,0,0.45) が 42か所 ／ rgba(0,0,0,0.55) が 3か所
#    司さん（2026-09-10）「★前みたいな 重たい感じに なってる所 あったら 先に 直せ★」
#    ⇒★これが 今 一番 大きい「重たい」所★
#
#  ★★けれど「実Excel は 暗くしない」を ★言葉で 断言しません★★★
#    2026-09-09 に 私は「画面は きれい」と 報告して ★間違えました★
#    （実Excel の ★中の 数★と 比べて いて ★画面に 出る 字★と 比べて いなかった）
#    ⇒★今度は 撮って 測る★
#
#  ★★何を 測るか★★
#    ①実Excel を ★見える形★で 立てる
#    ②シートに ★色を 塗った マス★を 置く（暗くなれば その 色が 沈む）
#    ③★窓を 開ける 前★の 絵を 撮る
#    ④★窓を 開けた 後★の 絵を 撮る（`Application.Dialogs` は 止まるので 使わない）
#      ⇒★止まらない 出し方★＝`CommandBars.ExecuteMso` で リボンの 窓を 開ける
#    ⑤2枚の ★同じ 場所の 点★を 数えて ★明るさが 落ちたか★を 見る
#
#  ★★止まる 出し方は 使いません★★
#    `Application.Dialogs(xlDialogFormatNumber).Show` は ★閉じるまで 返って きません★
#    ⇒ COM が 固まる ⇒★お金を 使う 物が 二重に 走る★の 型
#    ⇒ `ExecuteMso` は ★窓を 出して すぐ 返る★ので 絵が 撮れる
#
#  ★物差しの 決まり★
#    ・★絵を 撮って 点を 数える★（当て推量で 書かない）
#    ・★同じ 場所を 同じ 大きさで 切る★（比べる 物は 同じ 枠で 切る）
#    ・書き戻しは LF
#
#  ★司さんの 実物には 触りません★＝新しい ブック・保存せず
#
#  ★★2026-09-10 追記：この 機械では 走りません／司さんの 決定は「今のまま」★★
#    ①★この セッションからは Windows の 画面が 撮れません★
#      `CopyFromScreen` が「ハンドルが 無効です」で 止まります
#      ⇒★実Excel の 絵は 1枚も 出して いません★
#      ⇒★だから「実Excel は 暗くしない」を 根拠に して いません★
#    ②★代わりに うちの 画面を 4通り 実測しました★
#      窓の 外の 同じ 場所 350×240点の 明るさ（窓なし 251.5）
#        ★今（0.55/0.45）… 112.2＝55.4% 暗くなる★
#        0.25 … 187.6＝25.4% ／ 0.12 … 221.1＝12.1% ／ 無し … 251.5＝0%
#      絵 … docs/measured/e-oou-4tsu-2026-09-10.png（4通りを 1枚に）
#           docs/measured/e-oou-mae-ato-2026-09-10.png
#      ★どれも 実際に book.html を 書き換えて 撮った 本物の 画面★
#    ③★司さんが 4通りの 絵を 見て →「今でええよ」★（2026-09-10）
#      ⇒★覆い 46か所は そのまま★＝★もう 直しに 行かない★
#      ⇒ 変えるなら ★司さんの 新しい 一言★が 要る
#    ④★別の 機械なら この 道具は 走ります★
#      走らせて 実Excel が 撮れたら ★結果を ここに 足して ください★
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-oou-daiarogu.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-oou-daiarogu-2026-09-10.tsv'
$絵1 = Join-Path $ここ 'e-excel-mado-mae-2026-09-10.png'
$絵2 = Join-Path $ここ 'e-excel-mado-ato-2026-09-10.png'

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

function 窓の絵を撮る([string]$道) {
  # ★画面 全体を 撮る（Excel の 窓が 前に 出て いる 前提）★
  $w = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
  $bmp = New-Object System.Drawing.Bitmap($w.Width, $w.Height)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.CopyFromScreen($w.X, $w.Y, 0, 0, $bmp.Size)
  $g.Dispose()
  $bmp.Save($道, [System.Drawing.Imaging.ImageFormat]::Png)
  $bmp.Dispose()
  return @{ 幅 = $w.Width; 高 = $w.Height }
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $false
try {
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Worksheets.Item(1)

  # ★色を 塗った マスを 置く★＝暗くなれば この 色が 沈む
  #   ★2つ 置く★（明るい 黄 と 白）＝1つだと「たまたま」が 消せない
  $sh.Range('B2:E8').Interior.Color = 65535        # 黄（BGR 0x00FFFF）
  $sh.Range('B10:E16').Interior.Color = 16777215   # 白
  $sh.Range('B2').Value2 = 'ここが 暗くなるか 見る'
  $xl.WindowState = -4137                          # xlMaximized
  Start-Sleep -Milliseconds 1200

  Write-Host '★窓を 開ける 前の 絵を 撮ります★'
  $大きさ = 窓の絵を撮る $絵1
  Start-Sleep -Milliseconds 300

  # ★窓を 開ける（★止まらない 出し方★）★
  #   FormatCellsDialog … 「セルの書式設定」＝うちの `.fmtmd-ov` に あたる 物
  $開けた = $false
  foreach ($名 in @('FormatCellsDialog', 'FormatCells', 'FunctionWizard')) {
    try { $xl.CommandBars.ExecuteMso($名); $開けた = $true; Write-Host ('★開けた … ' + $名 + '★'); break }
    catch { }
  }
  if (-not $開けた) { Write-Error '★窓を 開けられなかった★'; exit 2 }
  Start-Sleep -Milliseconds 1500

  Write-Host '★窓を 開けた 後の 絵を 撮ります★'
  窓の絵を撮る $絵2 | Out-Null

  # ★窓を 閉じる★
  [System.Windows.Forms.SendKeys]::SendWait('{ESC}')
  Start-Sleep -Milliseconds 600

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★実Excel は 窓を 開けた 時 シートを 暗くするか★（2026-09-10）')
  $行.Add('#')
  $行.Add('# ★なぜ★ うちは 窓を 1つ 開けるたびに ★画面 全体を 45〜55% 黒く★する（45か所）')
  $行.Add('#   司さん「前みたいな 重たい感じに なってる所 あったら 先に 直せ」')
  $行.Add('#')
  $行.Add('# ★測り方★ 色を 塗った マスを 置き、窓を 開ける 前と 後の 絵を 撮って')
  $行.Add('#   ★同じ 場所の 点の 明るさ★を 比べる（★当て推量で 書かない★）')
  $行.Add('# ★止まる 出し方（Application.Dialogs…Show）は 使って いません★')
  $行.Add('#   ＝閉じるまで 返らない＝COM が 固まる')
  $行.Add('#   ⇒ CommandBars.ExecuteMso（出して すぐ 返る）')
  $行.Add('#')
  $行.Add('# ★どの Excel で 打ったか★ … 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★画面の 大きさ★ … ' + $大きさ.幅 + ' x ' + $大きさ.高)
  $行.Add('# ★絵★ 前 ' + (Split-Path -Leaf $絵1) + ' ／ 後 ' + (Split-Path -Leaf $絵2))
  $行.Add('#')
  $行.Add('# ★点を 数えるのは 別の 道具★ … node docs/measured/osu-oou-daiarogu.mjs')
  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))

  Write-Host ''
  Write-Host ('★絵を 撮った … ' + $絵1)
  Write-Host ('★絵を 撮った … ' + $絵2)
  Write-Host ('★書いた … ' + $出)

  $bk.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
