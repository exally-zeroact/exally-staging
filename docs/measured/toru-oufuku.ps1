# toru-oufuku.ps1 — ★うちが 書いた 物を 実Excel で 開き直して 確かめる★（2026-09-10）
#
#  ★★なぜ 要るか（★今日まで 一度も やって いませんでした★）★★
#    今まで やって いた事 … ★実Excel に 打たせて 答えを 合わせる★（Excel → うち）
#    やって いなかった事 … ★うちが 書いた 物を Excel で 開き直す★（うち → Excel）
#    ⇒ 司さん（2026-09-10）「★Exally でも Excel でも 使えるように 確かめながら やってるか？★」
#    ⇒★半分しか やって いませんでした★
#
#  ★★これが 無くて 危なかった 事★★
#    ・お客さんは ★Exally で 作って Excel で 開く★／★Excel で 作って Exally で 開く★
#    ・答えが 合って いても ★式が 落ちる／値が 消える★ かもしれない
#    ・実際 甲（溢れ）では ★前は C2〜C5 が 空で 渡って いました★
#      （監査に 言われて 初めて 測った）
#
#  ★★測る 事（3つ とも 見る）★★
#    ①★式★ … Excel が 持って いる 式（`.Formula`）
#    ②★出る字★ … 画面に 出る 字（`.Text`）
#    ③★答え★ … 中の 数（`.Value2`）
#    ⇒★答えだけ 見ない★（式が 落ちて 値だけ 残る 事が 在る）
#
#  ★★実Excel が わざと 書き換える 物が 在ります★★
#    実Excel は ★指数の 字を 打つと 式そのものを 平らな 数に 直して 持ちます★
#      打った `=1.64E-14` → 実Excel の 式 ★`=0.0000000000000164`★（出る字は 1.64E-14）
#      打った `=1E+20`    → ★`=100000000000000000000`★
#    ⇒★これは 穴では ありません★＝★実Excel と 同じ★
#    ⇒★2026-09-10 に 私は これを「穴だ」と 早とちりしました★
#      （実Excel が どうするかを 先に 測って いなかった）
#    ⇒★だから 下の 一覧に「書き換わって 正しい」と 書いて 在ります★
#
#  ★物差しの 決まり★
#    ・★材料は 別の 道具が 書いた ファイル★（`osu-oufuku.mjs` が 書き出す）
#    ・★司さんの 実物には 触りません★＝書き出した ファイルだけ 開く・保存せず
#    ・書き戻しは LF
#
#  使い方:
#    ① node docs/measured/osu-oufuku.mjs      … うちで 書き出す
#    ② pwsh -NoProfile -File docs/measured/toru-oufuku.ps1  … 実Excel で 開いて 測る
#    ③ node docs/measured/osu-oufuku.mjs --awase … 突き合わせ

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$ファイル = Join-Path $ここ 'oufuku-2026-09-10.xlsx'
$出 = Join-Path $ここ 'golden-oufuku-2026-09-10.tsv'

if (-not (Test-Path $ファイル)) {
  Write-Error ('★先に `node docs/measured/osu-oufuku.mjs` を 走らせて ください★（' + $ファイル + ' が 無い）')
  exit 2
}

# ★見る マス（★osu-oufuku.mjs と 同じ 並び★）★
$見る = @(
  @{ マス = 'D1'; 打った = '=XIRR(A1:A3,B1:B3)';     何 = 'お金の 利回り（09-10 PR #59）' },
  @{ マス = 'D2'; 打った = '=MIRR(A1:A3,0.1,0.12)';  何 = 'お金の 利回り（09-10 PR #60）' },
  @{ マス = 'D3'; 打った = '=XNPV(0.1,A1:A3,B1:B3)'; 何 = 'お金の 今の 値（09-10 PR #59）' },
  @{ マス = 'D4'; 打った = '=1.64E-14';              何 = '指数の 字（09-10 PR #62）★Excel が 式を 書き換えて 正しい★' },
  @{ マス = 'D5'; 打った = '=1/3';                   何 = '画面に 出る 字（09-10 PR #63）' },
  @{ マス = 'D6'; 打った = '=206800/1.1';            何 = '税抜き（09-10 PR #63）' },
  @{ マス = 'G1'; 打った = '=SORT(E1:E5)';           何 = '溢れの 頭（09-10 PR #65）' },
  @{ マス = 'G2'; 打った = '(溢れた 先)';            何 = '溢れの 先（09-10 PR #65）' },
  @{ マス = 'G3'; 打った = '(溢れた 先)';            何 = '溢れの 先（09-10 PR #65）' },
  @{ マス = 'G4'; 打った = '(溢れた 先)';            何 = '溢れの 先（09-10 PR #65）' },
  @{ マス = 'G5'; 打った = '(溢れた 先)';            何 = '溢れの 先（09-10 PR #65）' }
)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $bk = $xl.Workbooks.Open($ファイル)
  $sh = $bk.Worksheets.Item(1)

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★うちが 書いた 物を 実Excel で 開き直して 測った★（2026-09-10）')
  $行.Add('#')
  $行.Add('# ★なぜ★ 今まで ★実Excel に 打たせて 答えを 合わせる★しか して いなかった')
  $行.Add('#   ⇒ 司さん「Exally でも Excel でも 使えるように 確かめながら やってるか？」')
  $行.Add('#   ⇒★うちが 書いた 物を Excel で 開き直す★を 足した')
  $行.Add('#')
  $行.Add('# ★3つ とも 見る★ ①式(.Formula) ②出る字(.Text) ③答え(.Value2)')
  $行.Add('#   ⇒★答えだけ 見ない★＝式が 落ちて 値だけ 残る 事が 在る')
  $行.Add('#')
  $行.Add('# ★実Excel が わざと 書き換える 物★')
  $行.Add('#   指数の 字は ★式そのものを 平らな 数に 直して 持つ★')
  $行.Add('#     =1.64E-14 → =0.0000000000000164 ／ =1E+20 → =100000000000000000000')
  $行.Add('#   ⇒★穴では ありません＝実Excel と 同じ★')
  $行.Add('#')
  $行.Add('# ★どの Excel で 開いたか★ … 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('# ★開いた ファイル★ … ' + (Split-Path -Leaf $ファイル))
  $行.Add('#')
  $行.Add('# マス' + "`t" + '打った 字' + "`t" + '実Excel の 式' + "`t" + '出る字' + "`t" + '答え' + "`t" + '何を 見て いるか')

  foreach ($m in $見る) {
    $c = $sh.Range($m.マス)
    $式 = [string]$c.Formula
    $字 = [string]$c.Text
    $v = $c.Value2
    $答 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    $行.Add($m.マス + "`t" + $m.打った + "`t" + $式 + "`t" + $字 + "`t" + $答 + "`t" + $m.何)
    Write-Host ('  ' + $m.マス.PadRight(4) + ' 式=' + $式.PadRight(30) + ' 出る字=' + $字.PadRight(14) + ' 答え=' + $答)
  }

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★書いた … ' + $出 + '★')

  $bk.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
