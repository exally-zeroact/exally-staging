# toru-golden.ps1 — ★実Excel に 打たせて 答えを 取る★（2026-09-07）
#
#  ★★私が 計算しません★★
#    お金の 関数は ★日数の 数え方が 5通り★ある。
#    ⇒ 私が 式を 立てると ★私の 間違いが そのまま 正解に なります★。
#    ⇒★★答えは 実Excel から 取る★★（作るのは その後）
#
#  ★取り方★
#    ①`cases-kane.txt` の 式を そのまま `.Formula` に 入れる（★英語の 書き方★）
#    ②`.Value2` で ★丸めていない 数★を 読む（`.Text` は 桁が 落ちる）
#    ③★どの Excel で 打ったか★も 一緒に 書く（版・build・UI の 言語）
#
#  使い方: powershell -File docs/measured/kansuu46/toru-golden.ps1

$ErrorActionPreference = 'Stop'

# ══ ★★2つ目の 窓（2026-09-08 に 足した）★★ ══════════════════
#  ★物差しの 欠陥★ .Value2 は ★0 で ない 値に 0 を 返す★
#    =0.1+0.2-0.3    … .Value2 ★0★ ／ =(式)=0 ★False★ ／ (式)*1e17 5.55
#    =11.1+22.2-33.3 … .Value2 0   ／ =(式)=0 ★True★  ／ (式)*1e17 0
#    ⇒★.Value2 では この 2つが どちらも 0 に 見える★
#    正体 …★最後の 演算が ＋か− の 時だけ 実Excel が ★見せる 時に★ 0 に する★
#  ★もう1つ★ =DEC2BIN(0.5) は ★文字列の "0"★＝数の 0 では ない
#    ⇒ ="0"=0 は FALSE ⇒★見せかけの 0 と 同じ 顔★⇒★型を 見ないと 分けられない★
#  ⇒★見張り tests/monosashi-mado.test.mjs が これを 入れて いない 道具を 赤に する★
function 窓２_型($v) {
  if ($null -eq $v) { return 'Empty' }
  if ($v -is [string]) { return 'String' }
  if ($v -is [bool]) { return 'Boolean' }
  if ($v -is [double] -or $v -is [int] -or $v -is [long]) { return 'Number' }
  return 'Other'
}
function 窓２_本当にゼロか($sh, [string]$式) {
  # ★『0』が 出た 時だけ 呼ぶ★ … =(式)=0 の 真偽を 返す
  $中 = $式 -replace '^=\s*', ''
  try {
    $sh.Range('BZ1').Clear() | Out-Null
    $sh.Range('BZ1').Formula = ('=(' + $中 + ')=0')
    $z = $sh.Range('BZ1').Value2
    $sh.Range('BZ1').Clear() | Out-Null
    if ($z -is [bool]) { return $(if ($z) { 'TRUE' } else { 'FALSE' }) }
    return '★判じられない★'
  } catch { return '★判じられない★' }
}

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$入名 = 'cases-kane.txt'
if ($args.Count -ge 1) { $入名 = $args[0] }
$入 = Join-Path $ここ $入名
$出名 = 'golden-kane-2026-09-07.tsv'
if ($args.Count -ge 2) { $出名 = $args[1] }
$出 = Join-Path $ここ $出名

$行 = Get-Content -LiteralPath $入 -Encoding UTF8 | Where-Object { $_.Trim() -ne '' }
Write-Host "式 $($行.Count) 本"

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$版 = $xl.Version
$ビルド = $xl.Build
$言語 = $xl.LanguageSettings.LanguageID(2)   # 2 = msoLanguageIDUI

$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)

# ★Excel の 赤い 値 ⇒ 見て 分かる 字に する★（.Value2 は 負の 整数で 返す）
$誤りの番号 = @{
  -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!';
  -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!';
  -2146826246 = '#N/A';   -2146826245 = '#GETTING_DATA'; -2146826243 = '#SPILL!'
}

$結果 = New-Object System.Collections.ArrayList
$i = 0
foreach ($l in $行) {
  $i++
  $部 = $l -split "`t"
  $名 = $部[0]
  $式 = $部[1]
  $答 = ''
  $型 = ''
  try {
    $ws.Cells.Item(1, 1).Formula = $式
    $v = $ws.Cells.Item(1, 1).Value2
    if ($null -eq $v) { $答 = '(空)' ; $型 = 'null' }
    else {
      $型 = $v.GetType().Name
      if ($v -is [double]) { $答 = $v.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture) }
      elseif ($v -is [int] -and $誤りの番号.ContainsKey([int]$v)) {
        # ★Excel の 赤い 値は .Value2 では ★負の 整数★で 返る★
        #   ⇒ そのまま 書くと「-2146826252」という ★数字に 見える 答え★に なる
        $答 = $誤りの番号[[int]$v]; $型 = 'error値'
      }
      else { $答 = [string]$v }
    }
  } catch {
    # ★引数の 数が 違う＝名前は 在る／#NAME? ＝名前が 無い★
    $答 = '★打てない★ ' + ($_.Exception.Message -replace "`r?`n", ' ')
    $型 = 'error'
  }
  [void]$結果.Add(("{0}`t{1}`t{2}`t{3}" -f $名, $式, $答, $型))
  if ($i % 50 -eq 0) { Write-Host "  $i / $($行.Count)" }
}

$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

$頭 = @(
  "# ★実Excel に 打たせた 答え★（入 $入）",
  "# 取った 日 … 2026-09-07",
  "# ★どの Excel で 打ったか★ … 版 $版 ／ build $ビルド ／ UI の 言語 $言語",
  "# ★読み方★ … 数は .Value2 を 'R'（丸めない 書き方）で 出している",
  "# 関数`t式`t実Excel の 答え`t型"
)
$中 = $頭 + $結果
# ★改行は LF★（CRLF だと 読む 側の 最後の 列に 復帰の 字が 付いて ★型の 判定が 黙って 外れる★
#   ＝2026-09-07 に 実際に 踏んだ。101本が「違う」に 見えて 中身は 1e-15 の ずれだった）
[System.IO.File]::WriteAllText($出, (($中 -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding $false))
Write-Host "★書いた … $出（$($結果.Count) 本）★"
Write-Host "★Excel … 版 $版 ／ build $ビルド ／ UI $言語★"
