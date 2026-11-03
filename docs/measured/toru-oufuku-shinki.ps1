# toru-oufuku-shinki.ps1 — ★うちが 書き出した 新しい 関数を 実Excel で 開き直す★（2026-09-16）
#
#  ★★司さん（2026-09-16）★★
#    「Exally と Excel に 引き渡しても ちゃんと 動くか ★実際に 動作確認しながら★ やれよ」
#
#  ★★測る 事（3つとも 見る）★★
#    ①★式★  … `.Formula`   ＝★落ちて いないか★
#    ②★出る字★ … `.Text`
#    ③★答え★ … `.Value2`   ＝★うちの 答えと 合うか★
#    ⇒★答えだけ 見ない★（★式が 落ちて 値だけ 残る★ 事が 在る）
#
#  ★★実Excel を 叩く 決まり（memory の 通り）★★
#    ・★走らせる 前に Excel が 動いて いないか 2つの 道具で 数える★（1個でも 居たら 走らせない）
#    ・★Visible=$false / DisplayAlerts=$false★
#    ・★finally で 必ず Quit★／★消えるまで 待って 秒数を 出す★（実測 59秒）
#    ・★この .ps1 は BOM 付きで 保存する★
#    ・★司さんの 実物は 開きません★＝うちが 書き出した ファイルだけ
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-oufuku-shinki.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$ファイル = Join-Path $ここ 'oufuku-shinki.xlsx'
$出 = Join-Path $ここ 'golden-oufuku-shinki-2026-09-16.tsv'

if (-not (Test-Path $ファイル)) {
  Write-Error ('★先に `node docs/measured/osu-oufuku-shinki.mjs` を 走らせて ください★')
  exit 2
}

# ★★走らせる 前に Excel が 動いて いないか 2つの 道具で 数える★★
$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★前に 居た Excel … Get-Process ' + $数1 + '個 ／ Win32_Process ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) {
  Write-Error '★★Excel が 動いて います＝走らせません★★（司さんが 開いて いるかも しれません）'
  exit 3
}

# ★見る マス★（★osu-oufuku-shinki.mjs の 一覧と 同じ★）
# ★★見る マスは ★手で 書きません★★（2026-09-16 に 直しました）
#   ★なぜ★ 前は `@('C1',…,'C23')` と 手で 並べて いました。
#     ⇒★式を 3本 足したのに 名簿を 直さず、★26本 押して 23本しか 突き合わせて いませんでした★
#     ⇒★しかも 「全部 合って います」と 緑が 出ました★＝★分母を 出さない 緑は 嘘★
#   ★同じ 型が 今日 3回目★（PERMUTATIONA ／ RANK.AVG ／ `_xlfn.` 91個）
#   ⇒★★`osu-oufuku-shinki.mjs` が 書いた 紙から 読みます★★＝★手で 直す 所が 無い★
$うちの紙 = Join-Path $ここ 'oufuku-shinki-uchi.tsv'
if (-not (Test-Path $うちの紙)) {
  Write-Error '★先に `node docs/measured/osu-oufuku-shinki.mjs` を 走らせて ください★'; exit 4
}
$見る = @(Get-Content -LiteralPath $うちの紙 -Encoding UTF8 |
  Where-Object { $_ -and $_ -notmatch '^#' } |
  ForEach-Object { ($_ -split "`t")[0] })
if ($見る.Count -lt 1) { Write-Error '★紙から 1本も 読めません★'; exit 4 }
Write-Host ('★紙から 読んだ マス … ' + $見る.Count + '個★（★手で 並べて いません★）')

$xl = New-Object -ComObject Excel.Application
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Open($ファイル)
  $sh = $bk.Sheets.Item(1)

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★うちが 書き出した 新しい 関数を 実Excel で 開き直した★（2026-09-16）')
  $行.Add('#')
  $行.Add('# ★司さん（2026-09-16）★ 「Exally と Excel に 引き渡しても ちゃんと 動くか')
  $行.Add('#   ★実際に 動作確認しながら★ やれよ」')
  $行.Add('#')
  $行.Add('# ★3つとも 見る★ … ①式（.Formula） ②出る字（.Text） ③答え（.Value2）')
  $行.Add('#   ⇒★答えだけ 見ない★＝★式が 落ちて 値だけ 残る★ 事が 在る')
  $行.Add('#')
  $行.Add('# ★どの Excel で 開いたか★ … 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★開いた ファイル★ … ' + (Split-Path -Leaf $ファイル) + '（★うちが 書いた 物★）')
  $行.Add('#')
  $行.Add('# ★2つ目の 窓★ … `.Value2` は ★0 で ない 値に 0 を 返す★＝`=(式)=0` で 本当に 0 かを 聞く')
  $行.Add('#   ★型も 控える★（String / Double）')
  $行.Add('#')
  $行.Add('# マス' + "`t" + '実Excel の 式' + "`t" + '出る字' + "`t" + '答え' + "`t" + '=(式)=0' + "`t" + '型')

  foreach ($m in $見る) {
    $c = $sh.Range($m)
    $式 = [string]$c.Formula
    $字 = [string]$c.Text
    $v = $c.Value2
    $答 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    # ★★2つ目の 窓★★ … `.Value2` は ★0 で ない 値に 0 を 返す★事が 在る
    #   ⇒ `=(式)=0` を ★別の マスに 打って★ 本当に 0 かを 聞く
    #   ★型も 一緒に 控える★（String か Number か で 見分けが 変わる）
    $窓 = 'Z' + ($見る.IndexOf($m) + 1)
    $sh.Range($窓).Formula = '=(' + $式.Substring(1) + ')=0'
    $ゼロか = [string]$sh.Range($窓).Value2
    # ★型は 決まりの 形で 見る★（`-is [string]` / `-is [double]`）
    #   ＝★見張り（tests/monosashi-mado.test.mjs）が 探す 形★
    $型 = if ($null -eq $v) { '(空)' }
          elseif ($v -is [double]) { 'Double' }
          elseif ($v -is [string]) { 'String' }
          else { 'Other' }
    $行.Add($m + "`t" + $式 + "`t" + $字 + "`t" + $答 + "`t" + $ゼロか + "`t" + $型)
    Write-Host ('  ' + $m.PadRight(4) + ' 式=' + $式.PadRight(30) + ' 答え=' + $答)
  }

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★書いた … ' + $出 + '★')

  $bk.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  # ★★消えるまで 待って 秒数を 出す★★（★実測 59秒★／10秒の 見張りは 嘘の 赤）
  $t = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 120)) {
    Start-Sleep -Milliseconds 500
  }
  $t.Stop()
  $残 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
  Write-Host ('★Excel が 消えるまで ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒 ／ 残り ' + $残 + '個★')
  if ($残 -gt 0) { Write-Host '★★まだ 残って います＝次に 走らせる 前に 見て ください★★' }
}
