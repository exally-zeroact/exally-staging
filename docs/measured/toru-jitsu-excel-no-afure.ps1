# toru-jitsu-excel-no-afure.ps1
#   -- ★新しい 本で 範囲を 返す 式（溢れ）を 打った 時、実Excel が 各マスに 出す 字★を 取る（137）（2026-10-05・経営者）
#
#  ★★なぜ★★
#    Exally1 の 見つけ（10-05）＝★新しい 本で E1 `=A1:A3` が 画面で #VALUE!（溢れない）★／★溢れを 見る 式を 先に 打つと 0 の まま★
#    ⇒ 直す 前に ★実Excel の 答え★ が 要る（司さんの 本は 溢れ 0 個＝受け入れでは 1度も 見えない 形）
#
#  ★★取り方★★
#    ・★新しい 空の 本★（司さんの 本は 使わない）。A1:A3 に 1・2・3 を 置く
#    ・組ごとに ★新しい 板★ を 使う（組どうしが 混ざらない）
#    ・式は .Formula2（★人が 画面で 打った 時と 同じ 動的配列の 解釈★）
#    ・★打つ 順★ を 組ごとに 名指し（「溢れを 見る 式を 先に」を 再現する）
#    ・打った 後に 読む マス ＝ 組ごとに 名指し（.Text・.Value2 の 型・=(マス)=0 の 真偽・.HasSpill・.SpillingToRange の 番地）
#
#  ★★0 を 2つの 窓で 見る★★（tests/monosashi-mado.test.mjs）
#    ★2つ目の 窓★ 別の マスに `=(番地)=0` を 打つ（('=(' + 番地 + ')=0')）／★3つ目の 窓★ .Value2 の 型（-is [string] / -is [double]）
#
#  ★★門★★
#    ①5.1（exit 8）／②走らせる 前の Excel が 0個（exit 3）／③★頼んだ 数 ≠ 書いた 数 なら 赤（exit 6）★
#
#  使い方:
#    powershell.exe -NoProfile -ExecutionPolicy Bypass -File docs/measured/toru-jitsu-excel-no-afure.ps1 -出す先 <tsv>

param([string]$出す先 = '')

$版 = $PSVersionTable.PSVersion
Write-Host ('★走らせて いる 貝殻 ... PowerShell ' + $版.ToString() + '★')
if ($版.Major -ne 5) { Write-Host '★★powershell.exe（5.1）で 走らせて ください★★'; exit 8 }
if (-not $出す先) { Write-Host '★★-出す先 を 渡して ください★★'; exit 4 }
$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★走らせる 直前の Excel ... ' + $数1 + '個 ／ ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Host '★★Excel が 動いて います＝走らせません★★'; exit 3 }

# ★組★ ... 名・打つ 順（番地＝式 を 順に）・読む マス
# ★1組だけの 打つ は @(,@(...)) に する★＝@(@('E1','=A1:A3')) は PowerShell が 平らに して $p[0] が 'E' 1字に なる（10-05 に 止まった）
$組ら = @(
  @{ 名 = 'A 範囲そのもの'; 打つ = @(,@('E1', '=A1:A3')); 読む = @('E1', 'E2', 'E3', 'E4') },
  @{ 名 = 'B SEQUENCE'; 打つ = @(,@('E1', '=SEQUENCE(3)')); 読む = @('E1', 'E2', 'E3', 'E4') },
  @{ 名 = 'C FILTER'; 打つ = @(,@('E1', '=FILTER(A1:A3,A1:A3>1)')); 読む = @('E1', 'E2', 'E3') },
  @{ 名 = 'D SORT 逆'; 打つ = @(,@('E1', '=SORT(A1:A3,1,-1)')); 読む = @('E1', 'E2', 'E3') },
  @{ 名 = 'E 溢れが先・見る式が後'; 打つ = @(@('E1', '=A1:A3'), @('F1', '=E2+E3')); 読む = @('E1', 'E2', 'E3', 'F1') },
  @{ 名 = 'F 見る式が先・溢れが後'; 打つ = @(@('F1', '=E2+E3'), @('E1', '=A1:A3')); 読む = @('E1', 'E2', 'E3', 'F1') },
  @{ 名 = 'G 溢れを丸ごと見る（#）'; 打つ = @(@('E1', '=A1:A3'), @('F1', '=SUM(E1#)')); 読む = @('E1', 'F1') },
  @{ 名 = 'H 溢れ先が埋まっている'; 打つ = @(@('E2', '9'), @('E1', '=A1:A3')); 読む = @('E1', 'E2', 'E3') },
  @{ 名 = 'I 横に溢れる'; 打つ = @(,@('E1', '=TRANSPOSE(A1:A3)')); 読む = @('E1', 'F1', 'G1', 'H1') },
  @{ 名 = 'J 塞いだ 9 を 消す'; 打つ = @(@('E2', '9'), @('E1', '=A1:A3'), @('E2', '')); 読む = @('E1', 'E2', 'E3') },
  @{ 名 = 'K 塞いだ 9 を 字 x に 打ち直す'; 打つ = @(@('E2', '9'), @('E1', '=A1:A3'), @('E2', 'x')); 読む = @('E1', 'E2', 'E3') },
  @{ 名 = 'L E3 だけ 塞ぐ'; 打つ = @(@('E3', '9'), @('E1', '=A1:A3')); 読む = @('E1', 'E2', 'E3') },
  @{ 名 = 'M 溢れた 後に E2 に 打つ'; 打つ = @(@('E1', '=A1:A3'), @('E2', '9')); 読む = @('E1', 'E2', 'E3') },
  @{ 名 = 'N 溢れた 後に E2 に 打って 消す'; 打つ = @(@('E1', '=A1:A3'), @('E2', '9'), @('E2', '')); 読む = @('E1', 'E2', 'E3') },
  @{ 名 = 'O 塞いだ 時に E1# を 見る'; 打つ = @(@('E2', '9'), @('E1', '=A1:A3'), @('F1', '=SUM(E1#)')); 読む = @('E1', 'F1') },
  @{ 名 = 'P 誤り #CALC!（空の FILTER）'; 打つ = @(,@('E1', '=FILTER(A1:A3,A1:A3>99)')); 読む = @(,'E1') },
  @{ 名 = 'P 誤り #NULL!（交わらない）'; 打つ = @(,@('E1', '=SUM(A1 B1)')); 読む = @(,'E1') },
  @{ 名 = 'P 誤り #DIV/0!'; 打つ = @(,@('E1', '=1/0')); 読む = @(,'E1') },
  @{ 名 = 'P 誤り #VALUE!'; 打つ = @(,@('E1', '="a"+1')); 読む = @(,'E1') },
  @{ 名 = 'P 誤り #REF!'; 打つ = @(,@('E1', '=INDEX(A1:A3,9)')); 読む = @(,'E1') },
  @{ 名 = 'P 誤り #NAME?'; 打つ = @(,@('E1', '=NAIMONO(1)')); 読む = @(,'E1') },
  @{ 名 = 'P 誤り #NUM!'; 打つ = @(,@('E1', '=SQRT(-1)')); 読む = @(,'E1') },
  @{ 名 = 'P 誤り #N/A'; 打つ = @(,@('E1', '=NA()')); 読む = @(,'E1') },
  @{ 名 = 'P 誤り #CALC!（LAMBDA 呼ばず）'; 打つ = @(,@('E1', '=LAMBDA(x,x+1)')); 読む = @(,'E1') },
  @{ 名 = 'Q 交わり 2マス を SUM'; 打つ = @(@('B1', '10'), @('B2', '20'), @('B3', '30'), @('C1', '100'), @('C2', '200'), @('C3', '300'), @('E1', '=SUM(A1:B2 B1:C2)')); 読む = @(,'E1') },
  @{ 名 = 'Q 交わり 1マス を SUM'; 打つ = @(@('B1', '10'), @('B2', '20'), @('B3', '30'), @('C1', '100'), @('C2', '200'), @('C3', '300'), @('E1', '=SUM(A1:C1 B1:B3)')); 読む = @(,'E1') },
  @{ 名 = 'Q 交わり 無し を SUM'; 打つ = @(@('B1', '10'), @('B2', '20'), @('B3', '30'), @('C1', '100'), @('C2', '200'), @('C3', '300'), @('E1', '=SUM(A1:A2 C1:C2)')); 読む = @(,'E1') },
  @{ 名 = 'Q 交わり 1マス そのまま'; 打つ = @(@('B1', '10'), @('B2', '20'), @('B3', '30'), @('C1', '100'), @('C2', '200'), @('C3', '300'), @('E1', '=A2:C2 B1:B3')); 読む = @('E1', 'E2') },
  @{ 名 = 'Q 交わり 2マス そのまま（溢れる）'; 打つ = @(@('B1', '10'), @('B2', '20'), @('B3', '30'), @('C1', '100'), @('C2', '200'), @('C3', '300'), @('E1', '=A1:B2 B1:C2')); 読む = @('E1', 'E2', 'F1') },
  @{ 名 = 'Q 交わり 無し そのまま'; 打つ = @(@('B1', '10'), @('B2', '20'), @('B3', '30'), @('C1', '100'), @('C2', '200'), @('C3', '300'), @('E1', '=A1 C3')); 読む = @(,'E1') },
  @{ 名 = 'Q 交わり 3つ'; 打つ = @(@('B1', '10'), @('B2', '20'), @('B3', '30'), @('C1', '100'), @('C2', '200'), @('C3', '300'), @('E1', '=SUM(A1:C3 B1:C2 B2:B3)')); 読む = @(,'E1') },
  @{ 名 = 'R 空白 =A1 + B1'; 打つ = @(@('B1', '10'), @('B2', '20'), @('B3', '30'), @('E1', '=A1 + B1')); 読む = @(,'E1') },
  @{ 名 = 'R 空白 =A1 +B1'; 打つ = @(@('B1', '10'), @('B2', '20'), @('B3', '30'), @('E1', '=A1 +B1')); 読む = @(,'E1') },
  @{ 名 = 'R 空白 =A1 - B1'; 打つ = @(@('B1', '10'), @('B2', '20'), @('B3', '30'), @('E1', '=A1 - B1')); 読む = @(,'E1') },
  @{ 名 = 'R 空白 =A1 -B1'; 打つ = @(@('B1', '10'), @('B2', '20'), @('B3', '30'), @('E1', '=A1 -B1')); 読む = @(,'E1') },
  @{ 名 = 'R 空白 =A1 * 2 + B1'; 打つ = @(@('B1', '10'), @('B2', '20'), @('B3', '30'), @('E1', '=A1 * 2 + B1')); 読む = @(,'E1') },
  @{ 名 = 'R 空白 =A1 & B1'; 打つ = @(@('B1', '10'), @('B2', '20'), @('B3', '30'), @('E1', '=A1 & B1')); 読む = @(,'E1') },
  @{ 名 = 'R 空白 =(A1) (B1)'; 打つ = @(@('B1', '10'), @('B2', '20'), @('B3', '30'), @('E1', '=(A1) (B1)')); 読む = @(,'E1') },
  @{ 名 = 'R 空白 =(A1) + (B1)'; 打つ = @(@('B1', '10'), @('B2', '20'), @('B3', '30'), @('E1', '=(A1) + (B1)')); 読む = @(,'E1') },
  @{ 名 = 'R 空白 2つの 交わり'; 打つ = @(@('B1', '10'), @('B2', '20'), @('B3', '30'), @('E1', '=SUM(A1:A3  B1:B3)')); 読む = @(,'E1') },
  @{ 名 = 'R タブ 1つの 交わり'; 打つ = @(@('B1', '10'), @('B2', '20'), @('B3', '30'), @('E1', "=SUM(A1:A3`tB1:B3)")); 読む = @(,'E1') },
  @{ 名 = 'R 改行 入りの 交わり'; 打つ = @(@('B1', '10'), @('B2', '20'), @('B3', '30'), @('E1', "=SUM(A1:A3`nB1:B3)")); 読む = @(,'E1') },
  @{ 名 = 'R 改行 入りの 引数'; 打つ = @(@('B1', '10'), @('B2', '20'), @('B3', '30'), @('E1', "=SUM(A1,`nB1)")); 読む = @(,'E1') }
)
$頼む数 = 0; foreach ($g in $組ら) { $頼む数 += $g.読む.Count }
Write-Host ('★頼んだ 数★ ' + $頼む数 + '（' + $組ら.Count + '組）')

$ErrorActionPreference = 'Stop'   # ★誤りを 黙って 飛ばさない★（1回目は 書いた 0 で 訳が 出なかった）
$誤り = ''
$xl = New-Object -ComObject Excel.Application
$wb = $null; $ws = $null; $c = $null; $c2 = $null; $sp = $null
$行 = New-Object System.Collections.ArrayList
$版xl = ''
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $版xl = [string]$xl.Version + ' build ' + [string]$xl.Build
  $wb = $xl.Workbooks.Add()
  $i = 0
  foreach ($g in $組ら) {
    $i++
    if ($i -eq 1) { $ws = $wb.Worksheets.Item(1) } else { $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    $ws.Range('A1').Value2 = [double]1
    $ws.Range('A2').Value2 = [double]2
    $ws.Range('A3').Value2 = [double]3
    $受けない = ''
    foreach ($p in $g.打つ) {
      $c = $ws.Range([string]$p[0])
      # ★2026-10-05 ── 消す（''）と 字（数で ない）を 足した★（Exally1 の 依頼＝塞いだ 物を 消す／打ち直す）
      $d = 0.0
      if (([string]$p[1]).StartsWith('=')) {
        # ★2026-10-05 ── Excel が 式を 受けない 時（タブ入り 等）は 止めずに 「受けない」と 記す★（1回目は 止まって 3組 取れなかった）
        try { $c.Formula2 = [string]$p[1] } catch { try { $c.Formula = [string]$p[1] } catch { $受けない = '★Excel が 式を 受けない★ ' + $_.Exception.Message } }
      }
      elseif ([string]$p[1] -eq '') { [void]$c.ClearContents() }
      elseif ([double]::TryParse([string]$p[1], [ref]$d)) { $c.Value2 = $d }
      else { $c.Value2 = [string]$p[1] }
      $xl.Calculate()
      $c = $null
    }
    $xl.Calculate()
    $j = 0
    foreach ($a in $g.読む) {
      $j++
      $c = $ws.Range([string]$a)
      $字 = [string]$c.Text
      if ($受けない -ne '') { $字 = $受けない }
      $v = $c.Value2
      $型 = '空'
      if ($v -is [string]) { $型 = 'String' } elseif ($v -is [double]) { $型 = 'Double' } elseif ($null -ne $v) { $型 = $v.GetType().Name }
      $溢 = ''
      try { $溢 = [string]$c.HasSpill } catch { $溢 = '?' }
      $先 = ''
      try { if ($c.HasSpill) { $sp = $c.SpillingToRange; $先 = [string]$sp.Address($false, $false); $sp = $null } } catch { $先 = '' }
      $c2 = $ws.Cells.Item(20 + $j, 26)
      $c2.Formula = ('=(' + [string]$a + ')=0')
      $真 = [string]$c2.Value2
      $c2 = $null
      $式 = [string]$c.Formula2
      # ★紙の 行が 割れない 様に 改行・タブを 印に 換える★（2026-10-05 に 改行入りの 式で 紙が 割れた）
      $式 = $式.Replace("`r", '').Replace("`n", '⏎').Replace("`t", '→')
      $字 = $字.Replace("`r", '').Replace("`n", '⏎').Replace("`t", '→')
      [void]$行.Add(($g.名 + "`t" + $a + "`t" + $式 + "`t" + $字 + "`t" + $型 + "`t" + $真 + "`t" + $溢 + "`t" + $先))
      $c = $null
    }
    $ws = $null
  }
  $wb.Close($false)
} catch {
  $誤り = $_.Exception.Message + ' ／ 場所 ' + $_.InvocationInfo.ScriptLineNumber + '行'
  Write-Host ('★★止まった★★ ' + $誤り)
  try { if ($null -ne $wb) { $wb.Close($false) } } catch {}
} finally {
  $c = $null; $c2 = $null; $sp = $null; $ws = $null; $wb = $null
  $xl.Quit()
  $xl = $null
}
$書 = New-Object System.IO.StreamWriter($出す先, $false, (New-Object System.Text.UTF8Encoding($false)))
$書.NewLine = "`n"
$書.WriteLine('# この 紙は 式の 答えの 紙では ありません（溢れの 物差し）')
$書.WriteLine('# 137 実Excel の 溢れ（新しい 本・A1:A3＝1,2,3・.Formula2）／Excel ' + $版xl + '／頼んだ ' + $頼む数 + '／書いた ' + $行.Count)
$書.WriteLine("組`t番地`t式(Formula2)`t画面の字(.Text)`t値の型(.Value2)`t本当に0か(=(マス)=0)`t溢れの元か(HasSpill)`t溢れる先")
foreach ($l in $行) { $書.WriteLine($l) }
$書.Close()
Write-Host ('★Excel★ ' + $版xl + ' ／ ★頼んだ ' + $頼む数 + ' ／ 書いた ' + $行.Count + '★')
if ($頼む数 -ne $行.Count) { Write-Host '★★頼んだ 数と 書いた 数が 違います★★'; exit 6 }
