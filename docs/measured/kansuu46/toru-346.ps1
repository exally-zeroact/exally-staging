# toru-346.ps1 — ★答えを 1度も 確かめていない 346個を 実Excel に 聞く★（2026-09-08）
#
#  ★★司さん 2026-09-08「答え確かめて」★★
#
#  ★★やり方★★
#    ★346個ぶんの「正しい 呼び方」を 私が 手で 書いては いけません★
#    ⇒ 今日 70個の 名簿を 手で 写して ★10個 落としました★
#    ⇒★★手で 写す 作りに した のが 悪い★★（指示役 2026-09-08）
#    ⇒ だから ★実Excel 自身に 選ばせます★
#      ①引数の 形を 何通りか 用意する（★人が 決めるのは ここだけ★）
#      ②★実Excel が 誤りに しなかった 形★を その 関数の「正しい 呼び方」に する
#      ③その 式と 答えを 紙に 書く
#
#  ★★試し打ち（20個）で 分かった 事★★
#    ・19/20 で 呼び方が 見つかった ⇒ やり方は 効く
#    ・★でも 最初に 当たった 形 1本だけだと 弱い★
#      `=AVEDEV(A1)` → 0 の ような ★当たり前の 答え★に なり、
#      ★間違った 作りでも 通ってしまう★
#    ⇒★★当たった 形を ★全部★ 集める★★（1つの 関数に 何本も 答えを 持たせる）
#    ・ARABIC は 見つからなかった ⇒ ローマ数字の 形を 足した
#
#  ★★これは「答えが 合っているか」を 測る 紙です★★
#    ⇒ この 後 `tests/kansuu-346.test.mjs` が
#      ★同じ 式を Exally に 打たせて 1本ずつ 突き合わせます★
#
#  使い方: powershell -File docs/measured/kansuu46/toru-346.ps1
#          （名簿と 出し先を 変える 時）
#          pwsh -NoProfile -File docs/measured/kansuu46/toru-346.ps1 `
#               -名簿 ../mada-hakatte-inai-2026-09-14.txt -出 golden-mada-2026-09-14.tsv

# ★名簿と 出し先は 引数で 変えられます★（2026-09-14 に 足した＝★道具を 複製しない★）
#   例）pwsh -File toru-346.ps1 -名簿 ../mada-hakatte-inai.txt -出 golden-mada-2026-09-14.tsv
param(
  [string]$名簿 = '',
  [string]$出 = '',
  [switch]$上書き,
  [switch]$誤りも残す
)

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
if (-not $出) { $出 = Join-Path $ここ 'golden-346-2026-09-08.tsv' }
elseif (-not [System.IO.Path]::IsPathRooted($出)) { $出 = Join-Path $ここ $出 }
# ★もう 在る 紙を 黙って 上書きしない★（2026-09-14 に 足した）
#   訳＝この 紙には 後から ★`#材料` の 行★や 直しが 入って いる 事が 在ります。
#       走らせ直すと ★その 手入れが 消え★、★取った 日も 黙って 変わります★。
#   ⇒ わざと 取り直す 時だけ ★-上書き★ を 付ける。
if ((Test-Path $出) -and -not $上書き) {
  throw "★もう 在ります★: $出 `n  取り直すなら -上書き を 付けて ください（★前の 紙は 消えます★）"
}

# ★★2026-09-08 に 足した＝★この 表に 無い 番号は「数」として 紙に 載ってしまう★★
#   ⇒ 実際に 起きた … =SEQUENCE(0.5) の 答えが ★-2146826238★ の まま 6本 載っていた
#   ⇒ ★それを「実Excel は 答えを 返すのに うちは 誤り」の 山に 数えていた★＝★偽の 赤★
#   ⇒ 訳 … 古い 7つしか 書いていなかった（新しい #SPILL! 系は 2021年 以降の 物）
#   ⇒★見つけ方★ 金の 紙を「-20億より 小さい 数」で 洗った（誤りの 番号は 全部 -21億台）
#   ★数え方★ 番号 = 2000 + (2146826288 - |値|)  ⇒ 2050 は #CALC!
$誤りの番号 = @{
  -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!';
  -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!';
  -2146826246 = '#N/A';
  # ★★ここから 下が 2026-09-08 に 足した 分★★
  -2146826243 = '#SPILL!';   -2146826242 = '#CONNECT!'; -2146826241 = '#BLOCKED!';
  -2146826240 = '#UNKNOWN!'; -2146826239 = '#FIELD!';   -2146826238 = '#CALC!'; -2146826237 = '#BUSY!'
}

# ★引数の 形の 候補★（★人が 決めるのは ここだけ★）
$候補 = @(
  '()', '(A1)', '(A2)', '(A1:A5)', '(B1:B5)', '(A1:B5)',
  '(2)', '(2,3)', '(2,3,4)', '(2,3,4,5)', '(1,2,3,4,5)', '(1,2,3,4,5,6)', '(1,2,3,4,5,6,7)',
  '(0.5)', '(0.5,1)', '(0.5,1,2)', '(0.5,1,2,3)', '(0.5,1,2,3,4)',
  '("あ")', '("あ","い")', '("あ",1)', '("あ",1,2)', '("あいう",2)', '("あいう",2,1)',
  '("IV")', '("MMXXIV")', '("2024-01-15")', '("12:30")', '("101",2)',
  '(A1:A5,B1:B5)', '(A1:A5,2)', '(A1:A5,2,3)', '(A1:B5,2,FALSE)', '(A1:A5,A1)',
  '(A1:A5,"<3")', '(A1:A5,"<3",B1:B5)', '(A1:A5,1,B1:B5,1)',
  '(D1)', '(D1,D2)', '(D1,2)', '(D1,D2,1)', '(D1,D2,0)',
  '(D1,D2,0.05,0.1,100,2,0)', '(D1,D2,0.05,0.1,2)', '(D1,D2,0.05,100,2)',
  '(TRUE)', '(FALSE)', '(2,TRUE)', '(2,3,TRUE)',
  '(A1:A5,B1:B5,1)', '(2,3,4,5,6)', '(-2)', '(2.5,1)',
  # ★2026-09-14 に 足した★＝★まだ 答えを 取って いない 45個★に 当たる 形。
  #   （今までの 53通りでは VLOOKUP／MATCH／SUBTOTAL／TEXTJOIN／XLOOKUP 等に 当たりません）
  #   ★どれが 正しいかは 実Excel が 決めます★（人が 決めるのは 候補だけ）
  '(A1,A1:A5,B1:B5)',            # XLOOKUP／LOOKUP
  '(A1,A1:B5,2,FALSE)',          # VLOOKUP（完全一致）
  '(A1,A1:B5,2)',                # HLOOKUP／VLOOKUP（近い一致）
  '(A1,A1:A5,0)',                # MATCH
  '(9,A1:A5)',                   # SUBTOTAL（9＝合計）
  '(",",TRUE,A1:A5)',            # TEXTJOIN
  '(A1:A5,A1:A5>2)',             # FILTER
  '(A1:A5,0.5)',                 # PERCENTILE／QUARTILE.EXC（0〜1 の 位置）
  '(A1:A5,B1:B5,1,1)',           # T.TEST／TTEST
  '(A1:B5,1,D1:D2)',             # D系（DSUM DGET DAVERAGE …）
  '(TRUE,1)', '(TRUE,1,2)',      # IFS／IF
  '(1/0,2)',                     # IFERROR／IFNA
  '(D1,1)',                      # EDATE／EOMONTH
  '(0.05,12,100)',               # RATE／PMT／NPER
  '(0.05,12,100,0,0)',           # RATE（見当つき）
  '(100,0.05,1,12,0,0)',         # CUMIPMT／CUMPRINC（6つ）
  '(A1:A5,1,0)',                 # PERCENTRANK ほか
  # ★2026-09-14 の 1回目で ★呼び方が 見つからなかった 9個★に 当てる 形★
  #   BINOM.INV CRITBINOM CUMIPMT CUMPRINC DATEDIF DGET NEGBINOM.DIST NEGBINOMDIST RATE
  '(10,0.5,0.5)',                # BINOM.INV／CRITBINOM（試行・確率・α）
  '(2,3,0.5)',                   # NEGBINOM.DIST／NEGBINOMDIST（失敗・成功・確率）
  '(2,3,0.5,TRUE)',              # NEGBINOM.DIST（累積つき）
  '(0.05,12,100,1,12,0)',        # CUMIPMT／CUMPRINC（利率・期間・現価・始・終・型）
  '(D1,D2,"D")', '(D1,D2,"M")',  # DATEDIF（日数・月数）
  '(12,-10,100)',                # RATE（期間・支払・現価）
  '(12,-10,100,0,0,0.1)',        # RATE（見当つき）
  '(A1:B5,1,A1:A2)',             # DGET（表・列・条件）
  '(A1:B5,"A",A1:A2)',           # DGET（列を 名前で）
  # ★2026-09-15 に 足した★＝★INDEX と MATCH の 決まりを 測る 為★
  #   （司さんの 実物 1冊で ★のべの 68%★＝★一番 効く 2個★）
  #   ★今まで の 形では MATCH が 4本しか 当たらず、決まりを 測れませんでした★
  '(3,A1:A5,0)', '(3,A1:A5,1)', '(3,A1:A5,-1)',   # MATCH … ぴたり／以下で 一番 大きい／以上で 一番 小さい
  '(0,A1:A5,1)', '(9,A1:A5,0)',                   # 範囲の 外／見つからない
  '("あ",A1:A5,0)', '(2,B1:B5,0)', '(5,B1:B5,1)', # 字で 探す／別の 列
  '(A1:A5,3)', '(A1:B5,2,2)',                     # INDEX … 縦だけ／行と 列
  '(A1:B5,0,2)', '(A1:B5,2,0)',                   # ★0＝まるごと★（列／行）
  '(A1:A5,6)', '(A1:A5,0)',                       # 外れ／0
  # 2026-09-15（2回目）… 「見つからない時 どう なるか」を 測る 為に 足した
  #   1回目は 誤りを 捨てて いたので ★これらは 紙に 1行も 残りませんでした★
  '(0,A1:A5)',                                    # MATCH … 型を 省いて 見つからない
  '(A1:A5,-1)', '(A1:B5,6,1)', '(A1:B5,2,3)',     # INDEX … 負／行が 外／列が 外
  '(3,A1:A5,2)', '(3,A1:A5,"x")'                  # MATCH … 型が 変な 値
)

if (-not $名簿) { $名簿 = Join-Path (Split-Path -Parent $ここ) 'ugoku-tana-mikakunin.txt' }
elseif (-not [System.IO.Path]::IsPathRooted($名簿)) { $名簿 = Join-Path $ここ $名簿 }
if (-not (Test-Path $名簿)) { throw "★名簿が 無い★: $名簿" }
Write-Host "★名簿★ $名簿"
Write-Host "★出し先★ $出"
# ★★名前を 分ける★★（2026-09-14 に 踏んだ）
#   `param([string]$名簿)` に ★並びを 入れると PowerShell が 1本の 字に 潰します★
#   （`[string]$x = @('A','B')` → `"A B"`）
#   ⇒ 2026-09-14、45個の 名簿が ★1個の 関数★に なり ★式 0本★で 終わりました。
#   ⇒★読む 所（道）と 読んだ 物（並び）で 名前を 分ける★
$関数たち = Get-Content $名簿 -Encoding UTF8 |
  Where-Object { $_ -and -not $_.StartsWith('#') -and -not $_.StartsWith('★') } |
  ForEach-Object { $_ -split '\s+' } | Where-Object { $_ -match '^[A-Z][A-Z0-9._]*$' } |
  Sort-Object -Unique

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.ScreenUpdating = $false
$版 = $xl.Version; $ビルド = $xl.Build
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)

# ★材料★
$ws.Range('A1').Value2 = 1;  $ws.Range('A2').Value2 = 2
$ws.Range('A3').Value2 = 3;  $ws.Range('A4').Value2 = 4;  $ws.Range('A5').Value2 = 5
$ws.Range('B1').Value2 = 2;  $ws.Range('B2').Value2 = 4
$ws.Range('B3').Value2 = 6;  $ws.Range('B4').Value2 = 8;  $ws.Range('B5').Value2 = 10
$ws.Range('D1').Formula = '=DATE(2024,1,1)'
$ws.Range('D2').Formula = '=DATE(2026,1,1)'

$結果 = New-Object System.Collections.ArrayList
$呼び方あり = 0; $呼び方なし = 0; $本数 = 0
$見つからない名 = New-Object System.Collections.ArrayList
$i = 0
foreach ($f in $関数たち) {
  $i++
  # ★進みを こまめに 出す★（2026-09-14 に 直した）
  #   前は ★25個ごと★＝名簿が 45個だと ★1回しか 出ません★。
  #   ★殺された 時に「どこまで 測ったか」が 分かる 様に★ 毎回 出す（指示役1 の 注文）。
  #   ★紙は 書き上がってから 置きます★＝★進みは 画面だけ★（出来かけの 紙を 残さない）
  Write-Host ("  … {0}/{1} {2}  （ここまで 式 {3}本）" -f $i, $関数たち.Count, $f, $本数)
  $当たり = 0
  foreach ($a in $候補) {
    $式 = '=' + $f + $a
    try {
      # ★★同じ 道具を 2回 回したら 答えが 1行 違った★★（2026-09-15 実測）
      #   `=MATCH(3,A1:A5,2)` … 1回目 ★#N/A★ ／ 2回目 ★3★
      #   ⇒★測り道具が 再現しない＝どちらが 実Excel の 答えか 言えません★
      #   ⇒ 直し 2つ … ①★前の 式の 跡（溢れ）を 消す★ ②★計算させてから 読む★
      #     （`.Formula` を 入れただけで 読むと ★前の 値を 読む 事が 在ります★）
      $ws.Range('H1:Z50').Clear() | Out-Null
      $ws.Range('H1').Formula = $式
      $xl.CalculateFull()
      $v = $ws.Range('H1').Value2
      if ($null -eq $v) { continue }
      # ★誤りは 既定では 捨てます★＝★この 道具は 引数の 形を 総当たりする 物★なので
      #   ★誤り＝その 形は 呼び方では ない★と 見るのが 元の 狙い。
      # ★★でも「正しい 誤り」も 在ります★★（2026-09-15 に 踏んだ）
      #   `=MATCH(9,A1:A5,0)` → ★#N/A★＝★見つからない＝正しい 答え★
      #   ⇒ `-誤りも残す` を 付けると ★誤りも 紙に 残します★
      #   ★既定は 今の まま★＝★他の 紙の 数を 動かさない★
      if ($v -is [int] -and $誤りの番号.ContainsKey([int]$v)) {
        if (-not $誤りも残す) { continue }
        [void]$結果.Add(("{0}`t{1}`t{2}`t{3}" -f $f, $式, $誤りの番号[[int]$v], 'error値'))
        $当たり++; $本数++
        continue
      }
      $型 = $v.GetType().Name
      if ($型 -eq 'Object[,]') { continue }                                  # こぼれる 物は 別の 話
      $答 = if ($v -is [double]) { $v.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
      [void]$結果.Add(("{0}`t{1}`t{2}`t{3}" -f $f, $式, $答, $型))
      $当たり++; $本数++
    } catch { continue }
  }
  if ($当たり) { $呼び方あり++ } else { $呼び方なし++; [void]$見つからない名.Add($f) }
}

# ★★材料を 機械が 読める 形で 控える★★（2026-09-14 に 足した）
#   前は 頭に ★文章でしか★ 書いて いませんでした。
#   ⇒ 押す 道具が ★空の 表★で 押して ★偽の 負け★が 出ます
#     （2026-09-14 実測＝`=PRODUCT(A1:A5)` が 0 に なり「土台が 壊れている」と 報告される所だった）
#   ★値は Excel から 読み返します★＝★私が 計算しません★（日付の 通し番号も そのまま）
#   形 … `#材料<タブ>マス<タブ>値<タブ>型`（型＝数／字／真偽／式／空）
$材料の行 = New-Object System.Collections.ArrayList
foreach ($マス in @('A1','A2','A3','A4','A5','B1','B2','B3','B4','B5','D1','D2')) {
  $mv = $ws.Range($マス).Value2
  if ($null -eq $mv) { [void]$材料の行.Add("#材料`t$マス`t`t空"); continue }
  if ($mv -is [bool]) { [void]$材料の行.Add(("#材料`t{0}`t{1}`t真偽" -f $マス, $(if ($mv) {'TRUE'} else {'FALSE'}))); continue }
  if ($mv -is [string]) { [void]$材料の行.Add(("#材料`t{0}`t{1}`t字" -f $マス, $mv)); continue }
  $ms = if ($mv -is [double]) { $mv.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture) } else { [string]$mv }
  [void]$材料の行.Add(("#材料`t{0}`t{1}`t数" -f $マス, $ms))
}

$wb.Close($false); $xl.Quit()
foreach ($o in @($ws, $wb, $xl)) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null }

# ★★Excel が 残って いないか 数える★★（2026-09-15・指示役1 の 注文）
#   `Quit()` だけでは ★消えない事が 在ります★（2026-09-15 に 1本 残って いて 手で 止めた）
#   ⇒★止めた 後に 数えて 画面に 出す★＝★次の 回に 持ち越さない★
#   ★0.8秒では まだ 終わって いない 事が 在ります★（2026-09-15 実測＝0.8秒で 1個／その後 0個）
#   ⇒★1回 見て 騒がない★＝★10秒まで 待って、それでも 残って いたら 言う★
[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
$残り = @()
for ($t = 0; $t -lt 20; $t++) {
  Start-Sleep -Milliseconds 500
  $残り = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
  if ($残り.Count -eq 0) { break }
}
if ($残り.Count -eq 0) { Write-Host ('★Excel は 残って いません（0個・' + (($t + 1) * 0.5) + '秒で 消えた）★') }
else { Write-Host ('★★Excel が ' + $残り.Count + '個 残って います＝10秒 待っても 消えません＝手で 止めて ください★★ … ' + (($残り | ForEach-Object { $_.Id }) -join ' ')) }

$頭 = @(
  "# ★答えを 1度も 確かめていない 関数を 実Excel に 聞いた★",
  "# 取った 日 … 2026-09-08 ／ ★司さん「答え確かめて」★",
  "# ★どの Excel で 打ったか★ … 版 $版 ／ build $ビルド",
  "# ★やり方★ 引数の 形を $($候補.Count) 通り 用意し、★実Excel が 誤りに しなかった 形を 全部★ 採る",
  "#   ⇒★人が 決めるのは 形の 候補だけ／どれが 正しいかは 実Excel が 決める★",
  "#   ⇒★1本だけだと『当たり前の 答え』で 通ってしまう★ので ★当たった 形は 全部 残す★",
  "# ★材料★ A1:A5=1..5 ／ B1:B5=2,4,6,8,10 ／ D1=2024/1/1 ／ D2=2026/1/1",
  "# ★★式を 打った マス★★ … H1（★どこに 打ったかで 答えが 変わる 式が 在ります★）",
  "#   ＝実Excel は ★四角を 1つの 値に 詰める 時、式の 行／列と 交わる 所★を 採ります（暗黙の 交わり）",
  "#   例）`=INDEX(A1:B5,0,2)` は B1:B5（縦）→ ★式の 行＝1★ → B1 → 2",
  "#      `=INDEX(A1:B5,2,0)` は A2:B2（横）→ ★式の 列＝H★ → A:B と 交わらない → ★#VALUE!★",
  "#   ⇒★この 紙で 突き合わせる 時は 式を ★1行目・A〜B の 外の 列★に 打って ください★",
  "# ★★材料（機械が 読む）★★ … `#材料<タブ>マス<タブ>値<タブ>型`（★Excel から 読み返した 値★）",
  "# ★誤りも 残したか★ … $(if ($誤りも残す) { '★残した（-誤りも残す）＝「正しい 誤り」も 紙に 在ります★' } else { '残して いません（既定）＝誤りを 返す 形は 捨てました' })",
  "# ★数★ 関数 $($関数たち.Count)個 ／ 呼び方が 見つかった $呼び方あり 個 ／ 見つからない $呼び方なし 個 ／ 式 $本数 本",
  "# ★呼び方が 見つからない★ … " + ($見つからない名 -join ' ')
)

# ★見出しは 必ず 一番 下の # 行★（読む 側が 「最後の # 行」を 見出しと 見る 為）
#   ⇒ ★材料の 行を 見出しの 前に 挟む★
$見出し = @("# 関数`t式`t実Excel の 答え`t型")
[System.IO.File]::WriteAllText($出, ((($頭 + $材料の行 + $見出し + $結果) -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding $false))
Write-Host "★書いた … $出★"
Write-Host "★関数 $($関数たち.Count)個 ／ 呼び方あり $呼び方あり 個 ／ なし $呼び方なし 個 ／ ★式 $本数 本★★"
