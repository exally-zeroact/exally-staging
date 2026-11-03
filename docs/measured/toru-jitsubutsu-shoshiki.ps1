# toru-jitsubutsu-shoshiki.ps1 — ★司さんの実物が 使って いる「書式の 字」を 数える★（2026-09-15）
#
#  ★★なぜ★★
#    残り 1個が ★TEXT（実物で 740回）★。TEXT は ★1本の 関数では ありません★＝
#    ★書式の 字を 読んで 数を 字に する 台★が 要ります。
#    ⇒★全部 作ってから 測るのでは なく、実物が 使う 分から 作ります★。
#    ★指示役1 の 決め（2026-09-15）★＝
#      ★台は 1つ★＝TEXT も 画面（`js/book-open.js:378` の XLSX.SSF）も ★同じ 台を 呼ぶ★。
#
#  ★★条件（指示役1）★★
#    ①★読むだけ＝1バイトも 書かない★（★写しを 開く★／保存しない／Excel は 必ず 閉じる）
#    ②★出してよいのは 書式の 字の「形」と 数だけ★
#       ・★記号だけの 物は そのまま★（"#,##0.0" "yyyy/m/d" "aaa" "[h]:mm" …）
#       ・★他の 字が 混ざる 物は 伏せる★（"様" "円" も 伏せる側）
#         ⇒★伏せた 物も「何種類・何回 在ったか」は 出す★（★0件と 未測定を 分ける為★）
#       ・★会社名・金額・人の 名前は 1文字も 出しません★
#    ③★出来た 紙は repo に 入れない★（`.gitignore` の `golden-jitsubutsu-*`）
#    ④★本体だけ★＝競合コピー／復旧／旧版は ★開かない★
#
#  ★★分けて 出す★★（指示役1 の 注文）
#    ㋐★TEXT の 2つ目の 引数★（式の 中の 書式）… 740回ぶん
#    ㋑★マスに 付いて いる 書式★（画面が 使う）… 785マスぶん
#    ⇒★同じとは 限りません★＝★台が 両方を 賄えるか は ここで 決まる★
#
#  使い方: powershell -NoProfile -File docs/measured/toru-jitsubutsu-shoshiki.ps1
$ErrorActionPreference = 'Stop'

$本体 = 'C:\Users\zeroa\kyukyu-0907\代行計算表2026.xlsb'
$ここ  = Split-Path -Parent $MyInvocation.MyCommand.Path
$写し  = Join-Path $ここ 'jitsubutsu-utsushi-shoshiki.xlsb'
$出    = Join-Path $ここ 'golden-jitsubutsu-shoshiki-2026-09-15.tsv'

if (-not (Test-Path $本体)) { Write-Error ('★本体が 無い … ' + $本体 + '★'); exit 2 }

# ★★出してよい 字か★★（★記号だけの 物★）
#   書式の 記号 … 0 # ? . , % / \ - + ( ) : ; [ ] " 空白 と
#                 y m d h s a g e E "General" の 英字 ／ 色の 名前は 日本語が 混ざるので 伏せる側
#   ⇒★この 表に 無い 字が 1つでも 在れば 伏せます★（★迷ったら 伏せる★）
#  ★★この 国の Excel の「書式の 言葉」も 記号と 同じ 扱い★★（2026-09-15 に 直した）
#    1回目は ★記号だけ★に した ので ★18種の うち 12種・76% が 伏せ字★に なりました。
#    ⇒★`G/標準` `[赤]` `年 月 日 時 分 秒 曜` は ★書式の 決まり字★であって 客の 中身では ない★
#    ⇒★これを 伏せると 台が 何を 作れば よいか 分かりません★（＝測った 値打ちが 消える）
#    ★でも 増やしすぎない★＝★`円` `様` など「客が 書いた 字」は 伏せた まま★（指示役1 の 決め）
$書式の言葉 = '標準赤青緑黄紫水黒白年月日時分秒曜'
function 出してよいか([string]$s) {
  if ([string]::IsNullOrEmpty($s)) { return $true }
  foreach ($c in $s.ToCharArray()) {
    if ([string]$c -match '^[0#?.,%/\\\-+():;\[\]"@* _yYmMdDhHsSaAgGeE0-9]$') { continue }
    if ($書式の言葉.Contains([string]$c)) { continue }
    return $false
  }
  return $true
}
# ★伏せる 時の 見せ方★＝★字の 中身は 出さず 形だけ★
#   例 `#,##0"円"` → `記号(6)＋伏せ字(1)`
function 形にする([string]$s) {
  $記号 = 0; $伏せ = 0
  foreach ($c in $s.ToCharArray()) {
    if ([string]$c -match '^[0#?.,%/\\\-+():;\[\]"@* _yYmMdDhHsSaAgGeE0-9]$' -or $書式の言葉.Contains([string]$c)) { $記号++ } else { $伏せ++ }
  }
  return ('記号' + $記号 + '＋伏せ字' + $伏せ)
}

Copy-Item -LiteralPath $本体 -Destination $写し -Force
Write-Host ('★写しを 作った★ … ' + (Get-Item $写し).Length + ' バイト')

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$TEXTの書式 = @{}     # 書式の 字 → 回数（式の 中）
$マスの書式 = @{}     # 書式の 字 → マス数
$TEXT回数 = 0
$式の数 = 0
$マスの数 = 0
$板の数 = 0
try {
  $bk = $xl.Workbooks.Open($写し, 0, $true)     # ReadOnly
  $板の数 = $bk.Worksheets.Count
  foreach ($ws in $bk.Worksheets) {
    $r = $null
    try { $r = $ws.UsedRange } catch { continue }
    if ($null -eq $r) { continue }
    $マスの数 += $r.Count

    # ══ ㋐ TEXT の 2つ目の 引数 ══（★式だけ 見る★）
    $f = $null
    try { $f = $r.SpecialCells(-4123) } catch { $f = $null }   # xlCellTypeFormulas
    if ($null -ne $f) {
      foreach ($a in $f.Areas) {
        $v = $a.Formula
        if ($null -eq $v) { continue }
        $箱 = @()
        if ($v -is [object[,]]) {
          for ($i = 1; $i -le $a.Rows.Count; $i++) {
            for ($j = 1; $j -le $a.Columns.Count; $j++) {
              $c = $a.Cells.Item($i, $j)
              $親か = $true
              try { if ($c.HasArray) { $親か = ($c.Address(0,0) -eq $c.CurrentArray.Cells.Item(1,1).Address(0,0)) } } catch { $親か = $true }
              if ($親か) { $箱 += [string]$v[$i, $j] }
            }
          }
        } else {
          $親か = $true
          try { if ($a.HasArray) { $親か = ($a.Address(0,0) -eq $a.CurrentArray.Cells.Item(1,1).Address(0,0)) } } catch { $親か = $true }
          if ($親か) { $箱 += [string]$v }
        }
        foreach ($s in $箱) {
          if ([string]::IsNullOrEmpty($s)) { continue }
          if (-not $s.StartsWith('=')) { continue }
          $式の数++
          if ($s -notmatch '(?i)TEXT\s*\(') { continue }
          # ★TEXT( の 2つ目の 引数＝★二重引用符の 中★を 取る★
          #   ★式そのものは 1本も 出しません★＝★取るのは 書式の 字だけ★
          foreach ($m in [regex]::Matches($s, '(?i)TEXT\s*\(\s*(?:[^,()"]|"(?:[^"]|"")*"|\([^()]*\))*,\s*"((?:[^"]|"")*)"\s*\)')) {
            $TEXT回数++
            $書 = $m.Groups[1].Value -replace '""', '"'
            if ($TEXTの書式.ContainsKey($書)) { $TEXTの書式[$書]++ } else { $TEXTの書式[$書] = 1 }
          }
        }
      }
    }

    # ══ ㋑ マスに 付いて いる 書式 ══（★中身は 読みません＝書式の 字だけ★）
    #   ★UsedRange 全部を 1マスずつ 見ると 重い★ので
    #   ★同じ 書式の かたまりごと★に 数えます（`NumberFormatLocal` が 混ざると null）
    foreach ($a in $r.Areas) {
      $nf = $null
      try { $nf = $a.NumberFormatLocal } catch { $nf = $null }
      if ($nf -is [string]) {
        $n = $a.Count
        if ($マスの書式.ContainsKey($nf)) { $マスの書式[$nf] += $n } else { $マスの書式[$nf] = $n }
        continue
      }
      # ★混ざって いる＝1列ずつ 見る★（それでも 混ざれば 1マスずつ）
      foreach ($col in $a.Columns) {
        $nf2 = $null
        try { $nf2 = $col.NumberFormatLocal } catch { $nf2 = $null }
        if ($nf2 -is [string]) {
          $n2 = $col.Count
          if ($マスの書式.ContainsKey($nf2)) { $マスの書式[$nf2] += $n2 } else { $マスの書式[$nf2] = $n2 }
          continue
        }
        foreach ($cell in $col.Cells) {
          $nf3 = $null
          try { $nf3 = [string]$cell.NumberFormatLocal } catch { $nf3 = $null }
          if ($null -eq $nf3) { continue }
          if ($マスの書式.ContainsKey($nf3)) { $マスの書式[$nf3]++ } else { $マスの書式[$nf3] = 1 }
        }
      }
    }
    Write-Host ('  … 板 ' + $ws.Index + '／ここまで 式 ' + $式の数 + '本 ／ TEXT ' + $TEXT回数 + '回 ／ マスの書式 ' + $マスの書式.Count + '種')
  }
  $bk.Close($false)
} finally {
  $xl.Quit()
  foreach ($o in @($xl)) { try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null } catch {} }
  Remove-Item -LiteralPath $写し -Force -ErrorAction SilentlyContinue   # ★写しも 消す★
}
[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
$残り = @()
for ($t = 0; $t -lt 20; $t++) {
  Start-Sleep -Milliseconds 500
  $残り = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
  if ($残り.Count -eq 0) { break }
}
if ($残り.Count -eq 0) { Write-Host ('★Excel は 残って いません（0個・' + (($t + 1) * 0.5) + '秒で 消えた）★') }
else { Write-Host ('★★Excel が ' + $残り.Count + '個 残って います★★') }

function 書き出す($名, $表, $行) {
  $合 = 0
  foreach ($k in $表.Keys) { $合 += $表[$k] }
  $並 = $表.Keys | Sort-Object { -$表[$_] }
  $伏せ種 = 0; $伏せ回 = 0
  foreach ($k in $並) { if (-not (出してよいか $k)) { $伏せ種++; $伏せ回 += $表[$k] } }
  [void]$行.Add('')
  [void]$行.Add('# ══ ' + $名 + ' ══')
  [void]$行.Add('#   ★種類 ' + $表.Count + ' ／ のべ ' + $合 + '★')
  [void]$行.Add('#   ★伏せた 物 … 種類 ' + $伏せ種 + ' ／ のべ ' + $伏せ回 + '★（他の 字が 混ざる＝中身は 出しません）')
  $積 = 0; $i = 0
  foreach ($k in $並) {
    $i++
    $積 += $表[$k]
    $出字 = if (出してよいか $k) { $k } else { '★伏せ★ ' + (形にする $k) }
    [void]$行.Add(($名 + "`t" + $i + "`t" + $表[$k] + "`t" + [Math]::Round($積 / [Math]::Max(1, $合) * 100, 1) + "`t" + $出字))
  }
}

$行 = New-Object System.Collections.ArrayList
[void]$行.Add('# ★司さんの実物が 使って いる「書式の 字」★（2026-09-15）')
[void]$行.Add('#')
[void]$行.Add('#  ★この 紙は repo に 入れません★（.gitignore の golden-jitsubutsu-*）')
[void]$行.Add('#  ★中身は 1文字も 入って いません★＝★書式の 字と 数だけ★')
[void]$行.Add('#    ・記号だけの 物は そのまま／★他の 字が 混ざる 物は 伏せて 形だけ★')
[void]$行.Add('#    ・会社名・金額・人の 名前は ★1文字も★ 出して いません')
[void]$行.Add('#  ★読んだ 物★ … 代行計算表2026.xlsb の ★写し★（読むだけ・保存せず 閉じた・写しも 消した）')
[void]$行.Add('#  ★見た 数★ … 板 ' + $板の数 + '枚 ／ 使って いる マス ' + $マスの数 + ' ／ 式 ' + $式の数 + '本')
[void]$行.Add('#  ★TEXT( … ,"書式") を 取れた 回数★ … ' + $TEXT回数 + ' 回')
[void]$行.Add('#    ★これは「TEXT が 出てくる 回数」とは 違います★＝')
[void]$行.Add('#    ★2つ目が 字で 直に 書いて ある 物だけ★（マスを 指して いる 物は 取れません）')
[void]$行.Add('#  ★見て いない 物★ … VBA ／ 名前の 定義 ／ 条件付き書式 ／ テーブルの 式')
[void]$行.Add('# 種別' + "`t" + '順' + "`t" + '回数' + "`t" + 'ここまで%' + "`t" + '書式の字')
書き出す '㋐TEXTの書式' $TEXTの書式 $行
書き出す '㋑マスの書式' $マスの書式 $行
[System.IO.File]::WriteAllText($出, (($行 -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding $false))
Write-Host ('★書いた … ' + $出 + '★')
Write-Host ('★TEXT の 書式 ' + $TEXTの書式.Count + '種 ／ マスの 書式 ' + $マスの書式.Count + '種★')
