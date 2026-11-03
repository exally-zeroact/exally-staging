# measure-xml.ps1 — ★開発タブ（XML の 対応付け・入出力／コントロール）★ 2026-08-30
#   ★読むだけ★＝新しいブックで 測り、保存せずに 閉じる。
#   走らせ方: pwsh -NoProfile -ExecutionPolicy Bypass -File tools\measure-xml.ps1

$ErrorActionPreference = 'Continue'
function 出す($名, $値) { Write-Output ("{0,-50} = {1}" -f $名, $値) }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$tmp = Join-Path $env:TEMP ('exally-xml-' + [System.Guid]::NewGuid().ToString('N').Substring(0,8))
New-Item -ItemType Directory -Path $tmp | Out-Null
try {
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Worksheets.Item(1)

  Write-Output '=== XML の 対応付け（XmlMaps）==='
  出す 'はじめの 対応付けの 数' $bk.XmlMaps.Count

  # 小さな XML を 作って 読ませる
  $xml = @'
<?xml version="1.0" encoding="UTF-8"?>
<people>
  <person><name>ta</name><age>20</age></person>
  <person><name>ni</name><age>30</age></person>
</people>
'@
  $xmlPath = Join-Path $tmp 'people.xml'
  [System.IO.File]::WriteAllText($xmlPath, $xml, [System.Text.Encoding]::UTF8)

  try {
    $map = $bk.XmlMaps.Add($xmlPath, 'people')
    出す '★XmlMaps.Add できたか★' 'はい'
    出す '  対応付けの 数' $bk.XmlMaps.Count
    出す '  対応付けの 名前' $map.Name
    出す '  ルートの 名前' $map.RootElementName
    出す '  読み込み時に 検査するか（IsExportable）' $map.IsExportable
    出す '  スキーマの 数' $map.Schemas.Count
  } catch { 出す 'XmlMaps.Add' ('★出来ない★ ' + $_.Exception.Message.Trim()) }

  Write-Output ''
  Write-Output '=== インポート（読み込み）==='
  try {
    $r = $bk.XmlImport($xmlPath, [ref]$null, $true, $sh.Range('A1'))
    出す '★XmlImport の 戻り（1=成功）★' $r
    出す '  A1' $sh.Range('A1').Text
    出す '  B1' $sh.Range('B1').Text
    出す '  A2' $sh.Range('A2').Text
    出す '  B2' $sh.Range('B2').Text
    出す '  A3' $sh.Range('A3').Text
    出す '  B3' $sh.Range('B3').Text
    出す '  ★表に なったか（ListObjects）★' $sh.ListObjects.Count
    if ($sh.ListObjects.Count -gt 0) {
      出す '    表の 名前' $sh.ListObjects.Item(1).Name
      出す '    表の 大きさ' $sh.ListObjects.Item(1).Range.Address($false, $false)
    }
  } catch { 出す 'XmlImport' ('★出来ない★ ' + $_.Exception.Message.Trim()) }

  Write-Output ''
  Write-Output '=== エクスポート（書き出し）==='
  $outPath = Join-Path $tmp 'out.xml'
  try {
    if ($bk.XmlMaps.Count -gt 0) {
      $r2 = $bk.XmlMaps.Item(1).Export($outPath, $true)
      出す '★Export の 戻り（0=成功）★' $r2
      if (Test-Path $outPath) {
        $t = [System.IO.File]::ReadAllText($outPath)
        出す '  出た 字の 長さ' $t.Length
        出す '  はじめの 90文字' ($t.Substring(0, [Math]::Min(90, $t.Length)) -replace "`r`n", ' ')
      }
    }
  } catch { 出す 'Export' ('★出来ない★ ' + $_.Exception.Message.Trim()) }

  Write-Output ''
  Write-Output '=== データの 更新（XmlMap.DataBinding）==='
  try {
    出す 'DataBinding が 在るか' ($null -ne $bk.XmlMaps.Item(1).DataBinding)
    $bk.XmlMaps.Item(1).DataBinding.Refresh() | Out-Null
    出す '  Refresh 出来たか' 'はい'
  } catch { 出す 'DataBinding.Refresh' ('★出来ない★ ' + $_.Exception.Message.Trim()) }

  Write-Output ''
  Write-Output '=== 対応付けの プロパティ ==='
  try {
    $m = $bk.XmlMaps.Item(1)
    出す '  AdjustColumnWidth' $m.AdjustColumnWidth
    出す '  AppendOnImport' $m.AppendOnImport
    出す '  PreserveColumnFilter' $m.PreserveColumnFilter
    出す '  PreserveNumberFormatting' $m.PreserveNumberFormatting
    出す '  SaveDataSourceDefinition' $m.SaveDataSourceDefinition
    出す '  ShowImportExportValidationErrors' $m.ShowImportExportValidationErrors
  } catch { 出す '対応付けの プロパティ' ('★読めない★ ' + $_.Exception.Message.Trim()) }

  Write-Output ''
  Write-Output '=== 拡張パック（SmartDocument）==='
  出す 'COM から 触れるか' '★XML 拡張パックは 古い 仕組み（今の Excel では ほぼ 使わない）★'

  Write-Output ''
  Write-Output '=== コントロール（開発→コントロール）==='
  try {
    $b = $sh.Buttons().Add(300, 20, 90, 24)
    出す 'ボタンを 足せたか' 'はい'
    出す '  名前' $b.Name
    出す '  大きさ（幅×高）' ("{0} x {1}" -f $b.Width, $b.Height)
    出す '  字' $b.Text
    出す '  つないだ マクロ' ("'" + $b.OnAction + "'")
  } catch { 出す 'ボタン' ('★出来ない★ ' + $_.Exception.Message.Trim()) }
  try {
    $d = $sh.DropDowns().Add(300, 60, 90, 24)
    $d.AddItem('あ') ; $d.AddItem('い')
    出す 'ドロップダウンの 名前' $d.Name
    出す '  中の 数' $d.ListCount
    出す '  えらんでいる 番号（はじめ）' $d.ListIndex
  } catch { 出す 'ドロップダウン' ('★出来ない★ ' + $_.Exception.Message.Trim()) }
  出す 'デザイン モード' '★COM からは 切り替えられない（画面の ボタン）★'

  $bk.Close($false) | Out-Null
} finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
  Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
}
