# measure-review2.ps1 — ★校閲タブ（文章校正・アクセシビリティ・翻訳・変更内容・共有・インク）★ 2026-08-30
#   ★読むだけ★＝新しいブックで 測り、保存せずに 閉じる。
#   走らせ方: pwsh -NoProfile -ExecutionPolicy Bypass -File tools\measure-review2.ps1

$ErrorActionPreference = 'Continue'
function 出す($名, $値) { Write-Output ("{0,-52} = {1}" -f $名, $値) }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Worksheets.Item(1)

  Write-Output '=== スペル チェック（文章校正）==='
  $sh.Range('A1').Value2 = 'recieve'      # わざと まちがえた 英語
  $sh.Range('A2').Value2 = 'receive'
  $sh.Range('A3').Value2 = 'こんにちは'
  foreach ($w in @('recieve', 'receive', 'teh', 'the', 'こんにちは', 'apple')) {
    try { 出す ("CheckSpelling('" + $w + "')") $xl.CheckSpelling($w) }
    catch { 出す ("CheckSpelling('" + $w + "')") ('★読めない★ ' + $_.Exception.Message.Trim()) }
  }
  # 直しの 案が 出せるか
  try {
    $sug = $xl.GetSpellingSuggestions('recieve')
    出す "案の 数（recieve）" $sug.Count
    for ($i = 1; $i -le [Math]::Min(5, $sug.Count); $i++) { 出す ("  案" + $i) $sug.Item($i) }
  } catch { 出す '直しの 案' ('★出来ない★ ' + $_.Exception.Message.Trim()) }

  Write-Output ''
  Write-Output '=== 類義語辞典（シソーラス）==='
  出す 'COM から 呼べるか' '★Application.Thesaurus は 無い（画面の 窓だけ）★'

  Write-Output ''
  Write-Output '=== アクセシビリティ チェック ==='
  # 図に 代替テキストが 在るか（実Excel の 検査の 主な 中身）
  $sp = $sh.Shapes.AddShape(1, 20, 20, 100, 60)
  出す '図形の 名前' $sp.Name
  try { 出す '  代替テキスト（はじめ）' ("'" + $sp.AlternativeText + "'") }
  catch { 出す '  代替テキスト' ('★読めない★ ' + $_.Exception.Message.Trim()) }
  $sp.AlternativeText = 'まるい図'
  出す '  入れた 後' $sp.AlternativeText
  出す '  Title も 在るか' ($null -ne $sp.PSObject.Properties['Title'])

  Write-Output ''
  Write-Output '=== 翻訳 ==='
  出す 'COM から 呼べるか' '★Microsoft の サービスに つながる（機械からは 呼べない）★'

  Write-Output ''
  Write-Output '=== 変更内容を表示 / ブックの共有 ==='
  出す 'ブックは 共有されているか（MultiUserEditing）' $bk.MultiUserEditing
  try { 出す '  変更の 記録（KeepChangeHistory）' $bk.KeepChangeHistory }
  catch { 出す '  変更の 記録' ('★読めない★ ' + $_.Exception.Message.Trim()) }
  try { 出す '  何日 残すか（ChangeHistoryDuration）' $bk.ChangeHistoryDuration }
  catch { 出す '  何日 残すか' ('★読めない★ ' + $_.Exception.Message.Trim()) }
  try {
    $bk.HighlightChangesOptions(1, 'あ', 'い') | Out-Null
    出す '  変更内容を 出せたか' 'はい'
  } catch { 出す '  変更内容を 出す' ('★出来ない★ ' + $_.Exception.Message.Trim()) }
  出す '  HighlightChangesOnScreen' $bk.HighlightChangesOnScreen

  Write-Output ''
  Write-Output '=== インクを 非表示に する ==='
  出す 'COM から インクを 数えられるか' ($sh.Shapes.Count)
  出す '  （インクは Shapes に 入る）' '★形の 1つとして 数える★'

  Write-Output ''
  Write-Output '=== メモ / コメント（前から 在る 物の 確かめ）==='
  $sh.Range('C1').AddComment('めも') | Out-Null
  出す 'メモを 足した後の 数' $sh.Comments.Count
  出す '  中身' $sh.Range('C1').Comment.Text()
  出す '  はじめから 見えるか（Visible）' $sh.Range('C1').Comment.Visible

  $bk.Close($false) | Out-Null
} finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
