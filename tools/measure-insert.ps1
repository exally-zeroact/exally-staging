# measure-insert.ps1 — ★挿入タブ（チェックボックス・SmartArt・アイコン・数式・タイムライン）★ 2026-08-30
#   ★読むだけ★＝新しいブックで 測り、保存せずに 閉じる。
#   走らせ方: pwsh -NoProfile -ExecutionPolicy Bypass -File tools\measure-insert.ps1

$ErrorActionPreference = 'Continue'
function 出す($名, $値) { Write-Output ("{0,-48} = {1}" -f $名, $値) }
function 色に($v) {
  if ($null -eq $v) { return '(なし)' }
  $n = [int]$v
  return ('#{0:X2}{1:X2}{2:X2}' -f ($n -band 255), (($n -shr 8) -band 255), (($n -shr 16) -band 255))
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Worksheets.Item(1)

  Write-Output '=== チェック ボックス（挿入→コントロール）==='
  # 新しい「チェック ボックス」は セルの 中に 入る（TRUE/FALSE の 値）
  $sh.Range('A1').Value2 = $true
  出す 'TRUE を 入れた セルの 値' $sh.Range('A1').Value2
  出す '  その 型' $sh.Range('A1').Value2.GetType().Name
  出す '  画面に 出る 字' $sh.Range('A1').Text
  出す '  そろえ方（既定）' $sh.Range('A1').HorizontalAlignment    # -4108 = 中央
  $sh.Range('A2').Value2 = $false
  出す 'FALSE の セルの 字' $sh.Range('A2').Text
  出す 'TRUE と 1 は 同じか' ($sh.Range('A1').Value2 -eq 1)
  $sh.Range('A3').Formula = '=A1'
  出す '=A1 の 答え' $sh.Range('A3').Text
  $sh.Range('A4').Formula = '=COUNTIF(A1:A2,TRUE)'
  出す 'COUNTIF(A1:A2,TRUE)' $sh.Range('A4').Value2
  $sh.Range('A5').Formula = '=SUM(A1:A2)'
  出す '★SUM(A1:A2) は いくつか★（TRUE は 足されるか）' $sh.Range('A5').Value2
  $sh.Range('A6').Formula = '=A1+A2'
  出す '★A1+A2 は いくつか★' $sh.Range('A6').Value2

  Write-Output ''
  Write-Output '=== 古い チェック ボックス（フォーム コントロール）==='
  try {
    $cb = $sh.CheckBoxes().Add(100, 100, 100, 20)
    出す '足せたか' 'はい'
    出す '  名前' $cb.Name
    出す '  大きさ（幅×高）' ("{0} x {1}" -f $cb.Width, $cb.Height)
    出す '  はじめの 値（1=オン 変更なし=-4146）' $cb.Value
    $cb.Value = 1
    出す '  オンに した後' $cb.Value
    出す '  字' $cb.Text
  } catch { 出す 'フォームの チェックボックス' ('★出来ない★ ' + $_.Exception.Message.Trim()) }

  Write-Output ''
  Write-Output '=== SmartArt（挿入→図→SmartArt）==='
  try {
    $lay = $xl.SmartArtLayouts
    出す '型（レイアウト）の 数' $lay.Count
    for ($i = 1; $i -le [Math]::Min(8, $lay.Count); $i++) {
      出す ("  " + $i + '番目') $lay.Item($i).Name
    }
    $sa = $sh.Shapes.AddSmartArt($lay.Item(1), 20, 200, 300, 200)
    出す '足せたか' 'はい'
    出す '  形の 名前' $sa.Name
    出す '  中の 節（Nodes）の 数' $sa.SmartArt.AllNodes.Count
    $sa.SmartArt.AllNodes.Add() | Out-Null
    出す '  1つ 足した後の 数' $sa.SmartArt.AllNodes.Count
  } catch { 出す 'SmartArt' ('★出来ない★ ' + $_.Exception.Message.Trim()) }

  Write-Output ''
  Write-Output '=== アイコン（挿入→図→アイコン）==='
  出す 'COM から 一覧を 出せるか' '★出せない（絵は Microsoft の クラウドから 来る）★'

  Write-Output ''
  Write-Output '=== 数式（挿入→記号と特殊文字→数式）==='
  try {
    $eq = $sh.Shapes.AddTextbox(1, 20, 420, 200, 40)
    出す 'テキスト ボックスは 足せる' $eq.Name
    出す '  中の 字に 数式が 入るか（HasMath）' $eq.TextFrame2.TextRange.Math.Count
  } catch { 出す '数式' ('★読めない★ ' + $_.Exception.Message.Trim()) }

  Write-Output ''
  Write-Output '=== タイムライン（挿入→フィルター→タイムライン）==='
  出す 'ピボットが 要るか' '★要る（日付の 列を 持つ ピボットだけ）★'
  出す 'SlicerCaches に 入るか' '★入る（Timeline も スライサーの 仲間）★'

  Write-Output ''
  Write-Output '=== スクリーンショット（挿入→図→スクリーンショット）==='
  出す '他の 窓を 撮れるか' '★Excel は 撮れる（OS の 窓の 一覧から）★'

  $bk.Close($false) | Out-Null
} finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
