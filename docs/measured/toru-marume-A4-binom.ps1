# toru-marume-A4-binom.ps1 — ★A群の 4か所目★ BINOM.DIST.RANGE の 4桁丸めを 実Excel と 突き合わせる（2026-09-08）
#
#  ★★なぜ この 紙が 要るか★★
#    `exally-formula.js:1301`
#      String(Math.round(_jsBinomDistRange(…)*10000)/10000)
#    ⇒★お客さんの 計算の 答えを 黙って 4桁に して いる★＝A群
#    ⇒★A群は 3か所だと 思って いたが 4か所★でした
#      （見張りの 探す 形が ★入れ子 2段★を 読めず、★コードに 在るのに 見えて いなかった★）
#
#  ★★丸めの 悪さは 小さい 確率で 一番 大きい★★
#    確率が 4桁より 小さいと ★0 に なる★＝★起きない事に なる★
#    ⇒ その 形を わざと 入れて 測る
#
#  ★答えは 実Excel の 実測だけが 正★（AI に 出させない）
#  ★司さんの 実物には 触りません★＝新しい ブック・保存せず
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-marume-A4-binom.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-marume-A4-binom-2026-09-08.tsv'

$誤りの番号 = @{
  -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!';
  -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!';
  -2146826246 = '#N/A';   -2146826243 = '#SPILL!'; -2146826238 = '#CALC!'; -2146826237 = '#BUSY!'
}

# ══ ★2つ目の 窓（★.Value2 は 0で ない 値に 0 を 返す事が 在る★）★ ══
function 窓２_型($v) {
  if ($null -eq $v) { return 'Empty' }
  if ($v -is [string]) { return 'String' }
  if ($v -is [bool]) { return 'Boolean' }
  if ($v -is [double] -or $v -is [int] -or $v -is [long]) { return 'Number' }
  return 'Other'
}
function 窓２_本当にゼロか($sh, [string]$式) {
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

# ══ ★測る 式（★1つの 事だけを 見る／どちらでも 合う 組を 避ける★）★ ══
#   ★小さい 確率★ … 4桁で 丸めると 0 に なる か
#   ★普通の 確率★ … 4桁で 丸めると どれだけ ずれる か
$式ら = @(
  '=BINOM.DIST.RANGE(10,0.5,3)',
  '=BINOM.DIST.RANGE(20,0.3,4,8)',
  '=BINOM.DIST.RANGE(60,0.75,48)',
  '=BINOM.DIST.RANGE(100,0.5,60)',
  '=BINOM.DIST.RANGE(100,0.5,45,55)',
  '=BINOM.DIST.RANGE(100,0.02,5)',
  '=BINOM.DIST.RANGE(500,0.01,20)',
  '=BINOM.DIST.RANGE(1000,0.5,550)',
  '=BINOM.DIST.RANGE(1000,0.001,10)',
  '=BINOM.DIST.RANGE(50,0.9,49,50)',
  '=BINOM.DIST.RANGE(200,0.05,25)',
  '=BINOM.DIST.RANGE(30,0.5,0,30)'
)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Worksheets.Item(1)

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★A群の 4か所目★ BINOM.DIST.RANGE の 4桁丸めを 実Excel と 突き合わせた（2026-09-08）')
  $行.Add('#')
  $行.Add('# ★直す 所★ exally-formula.js:1301')
  $行.Add('#   String(Math.round(_jsBinomDistRange(…)★*10000)/10000★)')
  $行.Add('#   ⇒★お客さんの 計算の 答えを 黙って 4桁に して いる★')
  $行.Add('#')
  $行.Add('# ★★A群は 3か所だと 思って いたが 4か所でした★★')
  $行.Add('#   見張りの 探す 形が ★入れ子 2段★を 読めず、★コードに 在るのに 見えて いなかった★')
  $行.Add('#   ⇒ 見張りを ★括弧を 数えて 読む★形に 直してから 見つけた')
  $行.Add('#')
  $行.Add('# ★どの Excel で 打ったか★ … 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('#')
  $行.Add('# ★列★ 式／実Excel の 答え／型／4桁に 丸めた 数／相対の ずれ／★0に なるか★')
  $行.Add('# 式' + "`t" + '実Excelの答え' + "`t" + '型' + "`t" + '4桁に丸めた数' + "`t" + '相対のずれ' + "`t" + '判じ')

  $ゼロ化 = 0; $ずれ最大 = 0.0; $本数 = 0; $誤り = 0
  foreach ($f in $式ら) {
    $sh.Range('BZ1:CF5').Clear() | Out-Null
    $sh.Range('A1').Formula = $f
    $v = $sh.Range('A1').Value2
    $型 = 窓２_型 $v
    $本数++

    if (($v -is [int] -or $v -is [long]) -and $誤りの番号.ContainsKey([int]$v)) {
      $行.Add($f + "`t" + $誤りの番号[[int]$v] + "`t" + $型 + "`t" + '—' + "`t" + '—' + "`t" + '★実Excel が 誤りを 返した＝この 式は 使わない★')
      $誤り++
      Write-Host ('  ' + $f.PadRight(36) + ' ★' + $誤りの番号[[int]$v] + '★')
      continue
    }
    if ($型 -ne 'Number') {
      $行.Add($f + "`t" + [string]$v + "`t" + $型 + "`t" + '—' + "`t" + '—' + "`t" + '★数で 返って いない＝別に 見る★')
      Write-Host ('  ' + $f.PadRight(36) + ' ★型が ' + $型 + '★')
      continue
    }

    $本 = [double]$v
    $丸 = [Math]::Round($本 * 10000) / 10000
    $ずれ = 0.0
    if ($本 -ne 0) { $ずれ = [Math]::Abs(($丸 - $本) / $本) }
    if ($ずれ -gt $ずれ最大) { $ずれ最大 = $ずれ }

    $判 = ''
    if ($丸 -eq 0 -and $本 -ne 0) {
      # ★.Value2 の 0 に 騙されない＝2つ目の 窓で 実Excel 自身に 聞く★
      $ほんとに0 = 窓２_本当にゼロか $sh $f
      $判 = '★★4桁に すると 0 に なる（実Excel の 答えは 0 では ない … =(式)=0 は ' + $ほんとに0 + '）★★'
      $ゼロ化++
    } elseif ($ずれ -eq 0) {
      $判 = '4桁でも 同じ（この 式では 差が 出ない）'
    } else {
      $判 = '★4桁に すると ずれる（相対 ' + $ずれ.ToString('E3') + '）★'
    }

    $行.Add($f + "`t" + $本.ToString('R') + "`t" + $型 + "`t" + $丸.ToString('R') + "`t" + $ずれ.ToString('E3') + "`t" + $判)
    Write-Host ('  ' + $f.PadRight(36) + ' 実Excel=' + $本.ToString('R').PadRight(24) + ' 4桁=' + $丸.ToString('R').PadRight(10) + ' ' + $判)
  }

  $行.Add('#')
  $行.Add('# ★★締め★★ 測った ' + $本数 + '本 ／ ★4桁で 0 に なる ' + $ゼロ化 + '本★ ／ 相対の ずれ 最大 ' + $ずれ最大.ToString('E3'))
  if ($ゼロ化 -gt 0) {
    $行.Add('#   ⇒★★『起きる事』が『起きない事』に 変わる＝一番 悪い 形★★')
  }

  [System.IO.File]::WriteAllLines($出, $行, (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★★締め★★ 測った ' + $本数 + '本 ／ 実Excel が 誤り ' + $誤り + '本')
  Write-Host ('  ★4桁に すると 0 に なる … ' + $ゼロ化 + '本★')
  Write-Host ('  相対の ずれ 最大 … ' + $ずれ最大.ToString('E3'))
  Write-Host ('★書いた … ' + $出 + '★')

  $bk.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
