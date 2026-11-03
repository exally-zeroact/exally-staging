# measure-view.ps1 — ★表示タブ（ブックの表示・ウィンドウ）を 実Excel 16.0 で 測る★ 2026-08-30
#   ★読むだけ★＝新しいブックを 作って 測り、保存せずに 閉じる。客の実物には 触らない。
#   走らせ方: powershell -ExecutionPolicy Bypass -File tools\measure-view.ps1

$ErrorActionPreference = 'Stop'
function 出す($名, $値) { Write-Output ("{0,-40} = {1}" -f $名, $値) }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Worksheets.Item(1)
  $win = $xl.ActiveWindow

  Write-Output '=== ブックの 表示（View）==='
  出す 'はじめの View' $win.View          # 1=xlNormalView
  $win.View = 2; 出す 'ページ改行プレビュー(2) に した後' $win.View
  $win.View = 3; 出す 'ページレイアウト(3) に した後' $win.View
  $win.View = 1; 出す '標準(1) に 戻した後' $win.View

  Write-Output ''
  Write-Output '=== ウィンドウ枠の 固定（FreezePanes）==='
  出す 'はじめの FreezePanes' $win.FreezePanes
  出す 'はじめの SplitRow / SplitColumn' ("{0} / {1}" -f $win.SplitRow, $win.SplitColumn)
  $sh.Range('C3').Select() | Out-Null
  $win.FreezePanes = $true
  出す 'C3 で 固定した後 FreezePanes' $win.FreezePanes
  出す 'C3 で 固定した後 SplitRow / SplitColumn' ("{0} / {1}" -f $win.SplitRow, $win.SplitColumn)
  $win.FreezePanes = $false
  出す '外した後 SplitRow / SplitColumn' ("{0} / {1}" -f $win.SplitRow, $win.SplitColumn)

  # 先頭行だけ 固定（実Excelの「先頭行の固定」）
  $sh.Range('A2').Select() | Out-Null
  $win.FreezePanes = $true
  出す '先頭行だけ 固定 SplitRow / SplitColumn' ("{0} / {1}" -f $win.SplitRow, $win.SplitColumn)
  $win.FreezePanes = $false
  $sh.Range('B1').Select() | Out-Null
  $win.FreezePanes = $true
  出す '先頭列だけ 固定 SplitRow / SplitColumn' ("{0} / {1}" -f $win.SplitRow, $win.SplitColumn)
  $win.FreezePanes = $false

  Write-Output ''
  Write-Output '=== 分割（Split）==='
  出す 'はじめの Split' $win.Split
  $sh.Range('C3').Select() | Out-Null
  $win.Split = $true
  出す 'C3 で 分割した後 Split' $win.Split
  出す 'C3 で 分割した後 SplitRow / SplitColumn' ("{0} / {1}" -f $win.SplitRow, $win.SplitColumn)
  出す '窓の 数（Panes.Count）' $win.Panes.Count
  $win.Split = $false
  出す '外した後 Panes.Count' $win.Panes.Count

  Write-Output ''
  Write-Output '=== 表示しない / 再表示（Visible）==='
  出す 'はじめの Visible' $win.Visible

  Write-Output ''
  Write-Output '=== 新しいウィンドウ・整列・並べて比較 ==='
  出す '窓の 数（はじめ）' $xl.Windows.Count
  $win.NewWindow() | Out-Null
  出す '新しいウィンドウ の後 窓の 数' $xl.Windows.Count
  出す '1つ目の 名前' $xl.Windows.Item(1).Caption
  出す '2つ目の 名前' $xl.Windows.Item(2).Caption
  # 整列（xlArrangeStyleTiled=1 / Horizontal=-4128 / Vertical=-4166 / Cascade=7）
  foreach ($m in @(1, -4128, -4166, 7)) {
    try { $xl.Windows.Arrange($m); 出す ("整列 " + $m + " → 出来た") 'はい' }
    catch { 出す ("整列 " + $m) ('★出来ない★ ' + $_.Exception.Message) }
  }
  $xl.Windows.Item(2).Close() | Out-Null
  出す '閉じた後 窓の 数' $xl.Windows.Count

  # ★窓を 閉じると 前の 手が 使えなくなる★＝取り直す（実測 0x800A01A8）
  $win = $xl.ActiveWindow

  Write-Output ''
  Write-Output '=== 同時にスクロール ==='
  try { 出す 'SyncScrollingSideBySide' $xl.SyncScrollingSideBySide }
  catch { 出す 'SyncScrollingSideBySide' ('★読めない★ ' + $_.Exception.Message) }

  Write-Output ''
  Write-Output '=== ズーム（表示タブ）==='
  出す 'はじめの Zoom' $win.Zoom
  $sh.Range('A1:E10').Select() | Out-Null
  try {
    $win.Zoom = $true      # 選択範囲に 合わせる
    出す '選択範囲(A1:E10)に 合わせた後 Zoom' $win.Zoom
  } catch { 出す '選択範囲に 合わせる' ('★出来ない★ ' + $_.Exception.Message) }
  $win.Zoom = 100
  $sh.Range('A1:B3').Select() | Out-Null
  try { $win.Zoom = $true; 出す '選択範囲(A1:B3)に 合わせた後 Zoom' $win.Zoom }
  catch { 出す '選択範囲(A1:B3)' ('★出来ない★ ' + $_.Exception.Message) }
  $win.Zoom = 100

  Write-Output ''
  Write-Output '=== 枠線・見出し（表示タブ / ページレイアウトタブ）==='
  出す '枠線を 見せる（DisplayGridlines）' $win.DisplayGridlines
  出す '見出しを 見せる（DisplayHeadings）' $win.DisplayHeadings
  出す '数式バー（DisplayFormulaBar）' $xl.DisplayFormulaBar
  出す '枠線を 刷る（PageSetup.PrintGridlines）' $sh.PageSetup.PrintGridlines
  出す '見出しを 刷る（PageSetup.PrintHeadings）' $sh.PageSetup.PrintHeadings

  Write-Output ''
  Write-Output '=== テーマ（ページレイアウトタブ）==='
  try {
    $th = $bk.Theme
    出す 'テーマの 見出し色（dk2）' $th.ThemeColorScheme.Colors(3).RGB
    出す 'テーマの 差し色1（accent1）' $th.ThemeColorScheme.Colors(5).RGB
    出す 'テーマの 差し色2（accent2）' $th.ThemeColorScheme.Colors(6).RGB
    出す 'テーマの 見出しフォント' $th.ThemeFontScheme.MajorFont.Item(1).Name
    出す 'テーマの 本文フォント' $th.ThemeFontScheme.MinorFont.Item(1).Name
  } catch { 出す 'テーマ' ('★読めない★ ' + $_.Exception.Message) }

  Write-Output ''
  Write-Output '=== 改ページ（ページレイアウトタブ）==='
  出す 'はじめの 横の 改ページ数' $sh.HPageBreaks.Count
  出す 'はじめの 縦の 改ページ数' $sh.VPageBreaks.Count
  $sh.Range('A1:C50').Value2 = 'あ'
  $sh.Range('A20').Select() | Out-Null
  try {
    $sh.HPageBreaks.Add($sh.Range('A20')) | Out-Null
    出す 'A20 に 足した後 横の 改ページ数' $sh.HPageBreaks.Count
    出す '足した 改ページの 行' $sh.HPageBreaks.Item(1).Location.Row
  } catch { 出す '改ページを 足す' ('★出来ない★ ' + $_.Exception.Message) }

  $bk.Close($false) | Out-Null
} finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
