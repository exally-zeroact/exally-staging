# measure-sheetview.ps1 — ★シート ビュー／ユーザー設定のビュー／窓の名前★ 2026-08-30
#   ★読むだけ★＝新しいブックで 測り、保存せずに 閉じる。
#   走らせ方: pwsh -NoProfile -ExecutionPolicy Bypass -File tools\measure-sheetview.ps1

$ErrorActionPreference = 'Continue'
function 出す($名, $値) { Write-Output ("{0,-46} = {1}" -f $名, $値) }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Worksheets.Item(1)

  Write-Output '=== ユーザー設定のビュー（CustomViews）==='
  出す 'はじめの 数' $bk.CustomViews.Count
  $win = $xl.ActiveWindow
  $sh.Range('A1:C10').Value2 = 'あ'
  $sh.Rows('3:4').Hidden = $true
  $win.Zoom = 75
  try {
    $cv = $bk.CustomViews.Add('うちの見え方', $true, $true)   # 名前, 印刷設定も, 隠した行も
    出す '足した後の 数' $bk.CustomViews.Count
    出す '足した 物の 名前' $bk.CustomViews.Item(1).Name
    出す '印刷設定も 覚えるか（PrintSettings）' $bk.CustomViews.Item(1).PrintSettings
    出す '隠した行も 覚えるか（RowColSettings）' $bk.CustomViews.Item(1).RowColSettings
    # 別の 見え方に して 戻せるか 見る
    $sh.Rows('3:4').Hidden = $false
    $win.Zoom = 100
    出す '変えた後 3行目は 隠れているか' $sh.Rows('3:3').Hidden
    出す '変えた後 Zoom' $win.Zoom
    $bk.CustomViews.Item(1).Show()
    出す '★戻した後 3行目は 隠れているか★' $sh.Rows('3:3').Hidden
    出す '★戻した後 Zoom★' $xl.ActiveWindow.Zoom
  } catch { 出す 'ユーザー設定のビュー' ('★出来ない★ ' + $_.Exception.Message) }

  Write-Output ''
  Write-Output '=== シート ビュー（SheetViews）==='
  # 実Excelの「シート ビュー」は 共同編集（OneDrive/SharePoint）の 機能。
  # COM から 触れるか 試す（触れなければ ★測れない★ と 書き残す）。
  foreach ($n in @('SheetViews', 'NamedSheetViews', 'ActiveSheetView')) {
    try {
      $v = $sh.PSObject.Properties[$n]
      if ($null -eq $v) { 出す ('Worksheet.' + $n) '★COM に 無い★' }
      else { 出す ('Worksheet.' + $n) $v.Value }
    } catch { 出す ('Worksheet.' + $n) ('★読めない★ ' + $_.Exception.Message) }
  }
  try { $x = $sh.SheetViews; 出す 'Worksheet.SheetViews（直に）' $x }
  catch { 出す 'Worksheet.SheetViews（直に）' ('★読めない★ ' + $_.Exception.Message.Trim()) }

  Write-Output ''
  Write-Output '=== 窓の 名前（ウィンドウの切り替え に 並ぶ 字）==='
  出す '1つの 時の 名前' $xl.Windows.Item(1).Caption
  $xl.ActiveWindow.NewWindow() | Out-Null
  出す '2つに した時 1つ目' $xl.Windows.Item(1).Caption
  出す '2つに した時 2つ目' $xl.Windows.Item(2).Caption
  出す 'WindowNumber（1つ目）' $xl.Windows.Item(1).WindowNumber
  出す 'WindowNumber（2つ目）' $xl.Windows.Item(2).WindowNumber
  出す '同じ ブックを 見ているか' ($xl.Windows.Item(1).Parent.Name -eq $xl.Windows.Item(2).Parent.Name)

  Write-Output ''
  Write-Output '=== 表示しない / 再表示（Window.Visible）==='
  $xl.Windows.Item(2).Visible = $false
  出す '2つ目を 表示しない に した後 Windows.Count' $xl.Windows.Count
  出す '  その 窓の Visible' $xl.Windows.Item(2).Visible
  $xl.Windows.Item(2).Visible = $true
  出す '再表示した後 Visible' $xl.Windows.Item(2).Visible

  Write-Output ''
  Write-Output '=== 並べて比較（CompareSideBySideWith）==='
  try {
    $r = $xl.Windows.CompareSideBySideWith($xl.Windows.Item(2).Caption)
    出す '並べて比較 出来たか' $r
    出す '  SyncScrollingSideBySide' $xl.SyncScrollingSideBySide
    出す '  整列（1つ目の 上/左/幅/高）' ("{0}/{1}/{2}/{3}" -f `
      [math]::Round($xl.Windows.Item(1).Top,1), [math]::Round($xl.Windows.Item(1).Left,1), `
      [math]::Round($xl.Windows.Item(1).Width,1), [math]::Round($xl.Windows.Item(1).Height,1))
    出す '  整列（2つ目の 上/左/幅/高）' ("{0}/{1}/{2}/{3}" -f `
      [math]::Round($xl.Windows.Item(2).Top,1), [math]::Round($xl.Windows.Item(2).Left,1), `
      [math]::Round($xl.Windows.Item(2).Width,1), [math]::Round($xl.Windows.Item(2).Height,1))
    $xl.Windows.ResetPositionsSideBySide()
    出す '  位置を 元に戻した後（1つ目 上/左）' ("{0}/{1}" -f `
      [math]::Round($xl.Windows.Item(1).Top,1), [math]::Round($xl.Windows.Item(1).Left,1))
  } catch { 出す '並べて比較' ('★出来ない★ ' + $_.Exception.Message.Trim()) }

  $xl.Windows.Item(2).Close() | Out-Null
  $bk.Close($false) | Out-Null
} finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
