# measure-data3.ps1 — ★データタブ（接続・データの種類・詳細な絞り込み・データモデル）★ 2026-08-30
#   ★読むだけ★＝新しいブックで 測り、保存せずに 閉じる。★外の 世界へは つながない★
#   走らせ方: pwsh -NoProfile -ExecutionPolicy Bypass -File tools\measure-data3.ps1

$ErrorActionPreference = 'Continue'
function 出す($名, $値) { Write-Output ("{0,-50} = {1}" -f $名, $値) }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Worksheets.Item(1)

  Write-Output '=== クエリと接続（Connections）==='
  出す 'はじめの つなぎの 数（Connections）' $bk.Connections.Count
  出す 'はじめの クエリの 数（Queries）' $bk.Queries.Count
  出す 'すべて更新（RefreshAll）が 在るか' ($null -ne $bk.PSObject.Methods['RefreshAll'])
  try { $bk.RefreshAll(); 出す 'つなぎが 0でも RefreshAll は 通るか' 'はい（何も 起きない）' }
  catch { 出す 'RefreshAll' ('★出来ない★ ' + $_.Exception.Message.Trim()) }

  Write-Output ''
  Write-Output '=== ブックのリンク（外のブックへの つなぎ）==='
  $links = $bk.LinkSources(1)   # xlExcelLinks = 1
  出す 'はじめの 外のブックへの つなぎ' ($(if ($null -eq $links) { '★0（何も 無い）★' } else { $links.Count }))
  $sh.Range('A1').Formula = "='C:\ない\[ないブック.xlsx]Sheet1'!A1"
  $links2 = $bk.LinkSources(1)
  出す '外のブックを 指す 式を 書いた後' ($(if ($null -eq $links2) { '0' } else { $links2.Count }))
  if ($null -ne $links2) { 出す '  その 中身' $links2[1] }
  出す '  そのセルの 答え' $sh.Range('A1').Text
  $sh.Range('A1').Clear() | Out-Null

  Write-Output ''
  Write-Output '=== 詳細設定（フィルター オプション）==='
  $sh.Range('A1').Value2 = '名前'
  $sh.Range('A2').Value2 = 'あ'
  $sh.Range('A3').Value2 = 'い'
  $sh.Range('A4').Value2 = 'あ'
  $sh.Range('A5').Value2 = 'う'
  $sh.Range('C1').Value2 = '名前'
  $sh.Range('C2').Value2 = 'あ'
  try {
    # xlFilterCopy = 2 / xlFilterInPlace = 1
    $sh.Range('A1:A5').AdvancedFilter(2, $sh.Range('C1:C2'), $sh.Range('E1'), $false) | Out-Null
    出す '★条件で 別の所へ 出す（コピー）★ E1' $sh.Range('E1').Text
    出す '  E2' $sh.Range('E2').Text
    出す '  E3' $sh.Range('E3').Text
    出す '  出た 行の 数' ($sh.Range('E1').CurrentRegion.Rows.Count)
  } catch { 出す '詳細設定（コピー）' ('★出来ない★ ' + $_.Exception.Message.Trim()) }
  $sh.Range('E1:E9').Clear() | Out-Null
  try {
    $sh.Range('A1:A5').AdvancedFilter(2, $sh.Range('C1:C2'), $sh.Range('G1'), $true) | Out-Null
    出す '★重複を 除く（Unique=True）★ G2' $sh.Range('G2').Text
    出す '  出た 行の 数' ($sh.Range('G1').CurrentRegion.Rows.Count)
  } catch { 出す '詳細設定（重複を除く）' ('★出来ない★ ' + $_.Exception.Message.Trim()) }
  $sh.Range('G1:G9').Clear() | Out-Null
  # 条件を 空に して 重複を 除くだけ
  try {
    $sh.Range('A1:A5').AdvancedFilter(2, [System.Reflection.Missing]::Value, $sh.Range('I1'), $true) | Out-Null
    出す '★条件なしで 重複を 除く★ 出た 行の 数' ($sh.Range('I1').CurrentRegion.Rows.Count)
    出す '  I2 / I3 / I4' ("{0} / {1} / {2}" -f $sh.Range('I2').Text, $sh.Range('I3').Text, $sh.Range('I4').Text)
  } catch { 出す '条件なしの 重複除き' ('★出来ない★ ' + $_.Exception.Message.Trim()) }

  Write-Output ''
  Write-Output '=== データの種類（株式・通貨・地理）==='
  出す 'つながっているか' '★インターネットが 要る（Microsoft の サービス）★'
  try {
    $sh.Range('K1').Value2 = 'Japan'
    $sh.Range('K1:K1').ConvertToLinkedDataType(268435456, 'en-US') | Out-Null
    出す 'ConvertToLinkedDataType を 呼べたか' 'はい'
    出す '  そのセルの 型' $sh.Range('K1').Value2.GetType().Name
  } catch { 出す 'データの種類' ('★出来ない★ ' + $_.Exception.Message.Trim()) }

  Write-Output ''
  Write-Output '=== データ モデル / データ分析 ==='
  try { 出す 'Workbook.Model が 在るか' ($null -ne $bk.Model) }
  catch { 出す 'Workbook.Model' ('★読めない★ ' + $_.Exception.Message.Trim()) }
  try { 出す '  モデルの 表の 数' $bk.Model.ModelTables.Count }
  catch { 出す '  モデルの 表の 数' ('★読めない★ ' + $_.Exception.Message.Trim()) }
  出す 'データ分析（旧 アイデア）' '★Microsoft の クラウドで 動く（機械からは 呼べない）★'

  $bk.Close($false) | Out-Null
} finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
