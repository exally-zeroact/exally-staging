# toru-xlfn.ps1 — ★実Excel 自身に 打たせて ★ファイルの 中で どう 書くか★ を 読む★（2026-09-16）
#
#  ★★なぜ 要るか★★
#    うちが 書き出した `=NORM.DIST(...)` を 実Excel が ★#NAME?★ に する。
#    でも `=BINOM.DIST(...)` は ★動く★。★どちらも 2010年の 関数★。
#    ⇒★どこが 違うのかを 当てては いけません★
#      （2026-09-16 に ★当てて 2回 失敗して います★＝BESSEL の 係数／ODDFPRICE の 長さ）
#    ⇒★★実Excel 自身に 打たせて、保存した ファイルの 中を 読む★★
#
#  ★★やる 事★★
#    ①★新しい 空の ブック★に 実Excel が 式を 打つ（★司さんの 実物は 開きません★）
#    ②`.xlsx` で 保存して 閉じる
#    ③ その ファイルの 中の xml を `node docs/measured/osu-xlfn.mjs` で 読む
#
#  ★★実Excel を 叩く 決まり★★
#    ・★2つの 道具で 数えてから★／★Visible=$false★／★finally で Quit★
#    ・★消えるまで 待って 秒数を 出す★
#    ・★この .ps1 は BOM 付き★
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-xlfn.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出ファイル = Join-Path $ここ 'xlfn-shirabe.xlsx'

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★前に 居た Excel … Get-Process ' + $数1 + '個 ／ Win32_Process ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) {
  Write-Error '★★Excel が 動いて います＝走らせません★★'
  exit 3
}

# ★実Excel に 打たせる 式★（★壊れた 4本 ＋ 動いた 物（比べる 為）★）
$打つ = @(
  @{ マス = 'C1';  式 = '=NORM.DIST(5,3,2,TRUE)';        何 = '★壊れた★' },
  @{ マス = 'C2';  式 = '=BINOM.DIST(6,10,0.5,FALSE)';   何 = '動いた' },
  @{ マス = 'C3';  式 = '=CHISQ.DIST(2,3,TRUE)';         何 = '★壊れた★' },
  @{ マス = 'C4';  式 = '=CONFIDENCE.NORM(0.05,2.5,50)'; 何 = '★壊れた★' },
  @{ マス = 'C5';  式 = '=ISOMITTED(2)';                 何 = '★壊れた★' },
  @{ マス = 'C6';  式 = '=RANK.EQ(3,A1:A5)';             何 = '動いた' },
  @{ マス = 'C7';  式 = '=NORM.S.DIST(1,TRUE)';          何 = '比べる' },
  @{ マス = 'C8';  式 = '=T.DIST(2,3,TRUE)';             何 = '比べる' },
  @{ マス = 'C9';  式 = '=F.DIST(2,3,4,TRUE)';           何 = '比べる' },
  @{ マス = 'C10'; 式 = '=AREAS((A1,A2))';               何 = '束ね（動いた）' }
)

$xl = New-Object -ComObject Excel.Application
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)

  # ★材料★
  1..5 | ForEach-Object { $sh.Range('A' + $_).Value2 = $_ }

  foreach ($m in $打つ) {
    try {
      $sh.Range($m.マス).Formula = $m.式
      $v = $sh.Range($m.マス).Value2
      $答 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
      Write-Host ('  ' + $m.マス.PadRight(4) + $m.式.PadRight(34) + ' 答え=' + $答 + '  ' + $m.何)
    } catch {
      Write-Host ('  ' + $m.マス.PadRight(4) + $m.式.PadRight(34) + ' ★実Excel が 打てない★  ' + $m.何)
    }
  }

  if (Test-Path $出ファイル) { Remove-Item $出ファイル -Force }
  # 51 = xlOpenXMLWorkbook (.xlsx)
  $bk.SaveAs($出ファイル, 51)
  $bk.Close($false)
  Write-Host ''
  Write-Host ('★実Excel が 保存した … ' + $出ファイル + '★')
  Write-Host '★次に … node docs/measured/osu-xlfn.mjs★'
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $t = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 150)) {
    Start-Sleep -Milliseconds 500
  }
  $t.Stop()
  $残 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
  Write-Host ('★Excel が 消えるまで ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒 ／ 残り ' + $残 + '個★')
}

# ★作者の 欄を「見本」に（10-10 司さんの 決め＝実在の 名前は 架空に）★
#   Excel は サインインの アカウント名を 入れるので、Excel を 閉じた 後に scripts/sakusha-mihon.mjs で 替える
& node (Join-Path $PSScriptRoot '../../scripts/sakusha-mihon.mjs') $出ファイル
if ($LASTEXITCODE -ne 0) { throw '★作者の 欄を 見本に 替えられない★' }
