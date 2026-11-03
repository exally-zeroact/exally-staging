# toru-xlfn-zenbu.ps1 — ★台が 知る 関数 全部★を 実Excel に 打たせて `_xlfn.` が 要るかを 測る（2026-09-16）
#
#  ★★なぜ 要るか★★
#    `lib/xlsx-io.js` の 一覧に 漏れが 有ると ★その 式だけ #NAME?★ に なる。
#    ★同じ 型で 既に 2回 踏んで います★（PERMUTATIONA ／ RANK.AVG）
#    ⇒★手で 足すと また 漏れる★＝★全部 測る★
#
#  ★★測り方★★
#    ① 関数名を `=名前(1)` `=名前(1,1)` … と ★引数の 数を 増やしながら★ 打つ
#       （★Excel が 受け付けた 所で 止める★＝答えが 誤りでも 式は 保存される）
#    ② `.xlsx` で 保存
#    ③ `node docs/measured/osu-xlfn-zenbu.mjs` で ★中の xml★を 読み
#       `_xlfn.` が 付いた 名前を 拾う
#
#  ★★決まり★★
#    ・★2つの 道具で 数えてから★／★新しい 空の ブックだけ★（司さんの 実物は 開かない）
#    ・★Visible=$false★／★finally で Quit★／★消えるまで 待つ★／★BOM 付き★
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-xlfn-zenbu.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$名簿 = Join-Path $ここ 'xlfn-namae.txt'
$出ファイル = Join-Path $ここ 'xlfn-zenbu.xlsx'

if (-not (Test-Path $名簿)) {
  Write-Error '★先に `node docs/measured/osu-xlfn-zenbu.mjs --namae` を 走らせて ください★'
  exit 2
}

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★前に 居た Excel … Get-Process ' + $数1 + '個 ／ Win32_Process ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Error '★★Excel が 動いて います＝走らせません★★'; exit 3 }

$名前たち = @(Get-Content -Path $名簿 -Encoding UTF8 | Where-Object { $_ -and -not $_.StartsWith('#') })
Write-Host ('★測る 関数 … ' + $名前たち.Count + '個★')

$xl = New-Object -ComObject Excel.Application
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)
  1..5 | ForEach-Object { $sh.Range('A' + $_).Value2 = $_ }

  $行 = 1
  $打てた = 0
  $打てない = New-Object System.Collections.Generic.List[string]
  foreach ($n in $名前たち) {
    $ok = $false
    foreach ($k in 0..6) {
      $引 = if ($k -eq 0) { '' } else { (@('1') * $k) -join ',' }
      $式 = '=' + $n + '(' + $引 + ')'
      try {
        $sh.Range('C' + $行).Formula = $式
        $ok = $true
        break
      } catch { }
    }
    if ($ok) { $打てた++; $行++ } else { [void]$打てない.Add($n) }
  }

  Write-Host ('★実Excel が 打てた … ' + $打てた + '個 ／ 打てない … ' + $打てない.Count + '個★')
  if ($打てない.Count -gt 0) {
    Write-Host ('  ★打てない … ' + ($打てない -join ' ') + '★')
  }

  if (Test-Path $出ファイル) { Remove-Item $出ファイル -Force }
  $bk.SaveAs($出ファイル, 51)
  $bk.Close($false)
  Write-Host ''
  Write-Host ('★実Excel が 保存した … ' + $出ファイル + '★')
  Write-Host '★次に … node docs/measured/osu-xlfn-zenbu.mjs★'
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $t = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 180)) {
    Start-Sleep -Milliseconds 500
  }
  $t.Stop()
  Write-Host ('★Excel が 消えるまで ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒 ／ 残り ' +
    @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count + '個★')
}

# ★作者の 欄を「見本」に（10-10 司さんの 決め＝実在の 名前は 架空に）★
#   Excel は サインインの アカウント名を 入れるので、Excel を 閉じた 後に scripts/sakusha-mihon.mjs で 替える
& node (Join-Path $PSScriptRoot '../../scripts/sakusha-mihon.mjs') $出ファイル
if ($LASTEXITCODE -ne 0) { throw '★作者の 欄を 見本に 替えられない★' }
