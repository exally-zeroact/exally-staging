﻿# toru-jitsu-excel-ga-shuufuku-shita-ka.ps1
#   -- ★「実Excel が 修復しました」を 機械で 拾えるか★（93）（2026-09-21）
#
#  ★★なぜ★★
#    Exally1「足した 板を 包みに 入れる」を 書く ⇒ ★私が 実Excel に 開かせて 数える★
#    その 数え方の 中で ★一番 怖いのが「修復しました」★。
#    ⇒★これは 窓で 出る ので COM では 拾えないかも しれません★（Exally1 の 断り）
#    ⇒★見張りに する 前に ★わざと 壊して 拾えるか★ を 測ります★
#      （記憶「壊したのに 赤に ならない は まず 壊れて いるかを 見る」）
#
#  ★★測る 物★★ ... ★4通り★
#    ①きれいな 包み                 ⇒ 拾って は いけない（★嘘の 赤★を 出さない）
#    ②`[Content_Types].xml` を 消す ⇒ ★開けない はず★
#    ③`xl/worksheets/sheet1.xml` の 字を 壊す ⇒ ★修復に なる はず★
#    ④`xl/workbook.xml` に 在りもしない 板を 1行 足す ⇒ ★修復に なる はず★
#      （★これが Exally1 の 書く 物に 一番 近い 壊れ方★）
#
#  ★★拾い方の 候補（★どれが 効くか 分からないので 全部 見ます★）★★
#    ㋐`Workbooks.Open` が ★投げるか★
#    ㋑`Application.DisplayAlerts = $false` の まま 開けるか
#    ㋒`$bk.Name` が ★「修復済み」等に 変わるか★
#    ㋓★Excel が 書く 記録（log）★が `%TEMP%` 等に 出来るか
#    ㋔`$xl.RecentFiles` / `$bk.Saved` など 目に 見える 違い
#    ㋕★開いた 後の 板の 数★が 減るか
#
#  ★門★
#    ①貝殻が 5.1（exit 8）／②走らせる 前の Excel が 0個（exit 3）
#    ③★4通り 全部 走ったか★（exit 4）
#    ★★この 道具は 「拾えた／拾えない」を そのまま 出します★★
#      ＝★拾えなければ 「拾えません」と 書くのが 仕事★（当て推量で 見張りに しない）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }

$置場 = Join-Path $env:TEMP 'exally-shuufuku'
if (Test-Path $置場) { Remove-Item $置場 -Recurse -Force }
New-Item -ItemType Directory -Path $置場 | Out-Null

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個★')
if ($数1 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

Add-Type -AssemblyName System.IO.Compression.FileSystem

# ── ★まず きれいな 元を 1つ 作ります★（実Excel に 作らせる）
$元 = Join-Path $置場 'moto.xlsx'
$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)
  $sh.Range('A1').Value2 = [double]1
  $sh.Range('A2').Value2 = [string]'abc'
  $sh.Range('B1').Interior.Color = [int]65535
  $sh.Shapes.AddShape(1, 100, 20, 50, 50).Name = 'hanko'
  $bk.SaveAs($元, 51)
  $bk.Close($false); $sh = $null; $bk = $null
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
if (-not (Test-Path $元)) { Write-Host '★★元が 作れて いません★★'; exit 5 }
Write-Host ('★元★ ' + (Get-Item $元).Length + ' バイト')

# ── ★4通りの 包みを 作ります★（★zip を 組み直す＝Excel は 使いません★）
function 写して細工($名, $する) {
  $先 = Join-Path $置場 $名
  Copy-Item $元 $先 -Force
  $tmp = Join-Path $置場 ($名 + '.d')
  if (Test-Path $tmp) { Remove-Item $tmp -Recurse -Force }
  [IO.Compression.ZipFile]::ExtractToDirectory($先, $tmp)
  & $する $tmp
  Remove-Item $先 -Force
  # ★★2026-09-21 ── `CreateFromDirectory` は ★名前に 逆斜線を 使います★★（.NET Framework）
  #   ⇒`xl\styles.xml` に なり ★zip の 決まり（斜線）と 違います★
  #   ⇒前後を 突き合わせると ★全部の 部品が 「増えた」に 見えました★
  #   ⇒さらに ★実Excel が 投げた 因が 「壊れ」か 「名前」か 分けられなく なります★
  #   ⇒★自分で 1本ずつ 入れ、名前は 斜線に 直します★
  $z = [IO.Compression.ZipFile]::Open($先, 'Create')
  try {
    foreach ($f in Get-ChildItem -LiteralPath $tmp -Recurse -File) {
      # ★逆斜線を 1文字も 書きません★（heredoc や 便りで ★落ちます★＝記憶の 決まり）
      #   ⇒2026-09-21 に ここで 落ちて `.Replace('', '/')` に なり、
      #     ★空の 包み（22バイト）★が 4本 出来ました。
      $逆 = [string][char]92
      $な = $f.FullName.Substring($tmp.Length + 1).Replace($逆, '/')
      $e = $z.CreateEntry($な)
      $w = $e.Open()
      $bytes = [IO.File]::ReadAllBytes($f.FullName)
      $w.Write($bytes, 0, $bytes.Length)
      $w.Close()
    }
  } finally { $z.Dispose() }
  return $先
}

# ★★★⓪が 一番 大事です★★★
#   ★壊さずに ★詰め直しただけ★ の 物★
#   ＝これが 投げるなら ★投げた 因は 「壊れ」では なく 「詰め直し」★＝★測りが 嘘に なります★
#   ⇒2026-09-21 に ★これを 置かずに 測り、危うく 嘘の 答えを 出す 所でした★
$品 = @()
$品 += @{ 札 = '①きれい（Excel が 書いた まま）'; 道 = $元 }
$品 += @{ 札 = '⓪詰め直しただけ（壊さない）';     道 = (写して細工 'tsumenaoshi.xlsx' { param($d) }) }
# ★★2026-09-21 ── ★`[Content_Types].xml` の 角括弧は PowerShell では ワイルドカード★★
#   `Remove-Item (Join-Path $d '[Content_Types].xml')` は ★文字の 集まり★と 読まれ、
#   ★何にも 当たらず 黙って 何も しません★（投げません）。
#   ⇒その まま 開かせて ★「壊したのに 開けた」＝嘘の 緑★を 1回 出しました。
#   ⇒★`-LiteralPath` を 使います★。★消えたかを その場で 数えます★。
$品 += @{ 札 = '②型の 名簿を 消す';     道 = (写して細工 'kowasu2.xlsx' { param($d)
    $f = Join-Path $d '[Content_Types].xml'
    Remove-Item -LiteralPath $f -Force
    if (Test-Path -LiteralPath $f) { Write-Host '★★消えて いません★★'; exit 6 }
  }) }
$品 += @{ 札 = '③板の 字を 壊す';       道 = (写して細工 'kowasu3.xlsx' { param($d)
    $f = Join-Path $d 'xl\worksheets\sheet1.xml'
    $s = [IO.File]::ReadAllText($f)
    [IO.File]::WriteAllText($f, $s.Replace('</sheetData>', '<row r="99"><c r="A99"><v>x'))
  }) }
$品 += @{ 札 = '④在りもしない 板を 足す'; 道 = (写して細工 'kowasu4.xlsx' { param($d)
    $f = Join-Path $d 'xl\workbook.xml'
    $s = [IO.File]::ReadAllText($f)
    [IO.File]::WriteAllText($f, $s.Replace('</sheets>', '<sheet name="Nai" sheetId="9" r:id="rId99"/></sheets>'))
  }) }
Write-Host ('★作った 包み★ ' + $品.Count + '本（決め打ち 5本）')
if ($品.Count -ne 5) { exit 4 }

# ── ★1本ずつ 開かせて 見える 物を 全部 書き出します★
$走った = 0
foreach ($x in $品) {
  Write-Host ''
  Write-Host ('★★' + $x.札 + '★★ ... ' + (Split-Path $x.道 -Leaf) + ' / ' + (Get-Item $x.道).Length + ' バイト')
  $前記録 = @(Get-ChildItem $env:TEMP -Filter '*.log' -ErrorAction SilentlyContinue).Count
  $xl = New-Object -ComObject Excel.Application
  $bk = $null
  $投げた = ''
  try {
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    try { $bk = $xl.Workbooks.Open($x.道, $false, $true) }
    catch { $投げた = $_.Exception.Message }
    if ($null -ne $bk) {
      # ★★PowerShell 5.1 に `if` を ★式★ として 書けません★★（7の 書き方）
      #   ＝`(if (...) {a} else {b})` は 「`if` という コマンドが 無い」で 落ちます
      $投札 = 'いいえ'
      if ($投げた) { $投札 = '★投げました★ ' + $投げた }
      Write-Host ('  ㋐投げたか ... ' + $投札)
      Write-Host ('  ㋒名 ......... ' + [string]$bk.Name)
      Write-Host ('  ㋕板の 数 .... ' + [string]$bk.Sheets.Count)
      Write-Host ('  ㋔Saved ...... ' + [string]$bk.Saved + ' ／ ReadOnly ' + [string]$bk.ReadOnly)
      $図 = 0
      try { $図 = [int]$bk.Sheets.Item(1).Shapes.Count } catch { $図 = -1 }
      Write-Host ('  　図形 ....... ' + $図)
      $bk.Close($false); $bk = $null
    } else {
      Write-Host ('  ㋐投げたか ... ★投げました★ ' + $投げた)
    }
  } finally {
    if ($null -ne $bk) { $bk.Close($false) }
    $bk = $null
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    $xl = $null
    $t = [Diagnostics.Stopwatch]::StartNew()
    while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 300)) { Start-Sleep -Milliseconds 250 }
  }
  $後記録 = @(Get-ChildItem $env:TEMP -Filter '*.log' -ErrorAction SilentlyContinue).Count
  Write-Host ('  ㋓TEMP の log ... ' + $前記録 + ' ⇒ ' + $後記録)
  $走った++
}
Write-Host ''
Write-Host ('★走った 数★ ' + $走った + ' / 5')
if ($走った -ne 5) { exit 4 }
