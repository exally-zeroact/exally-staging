# measure-xml2.ps1 — ★XML の 読み込み（XmlImport）を きちんと 測る★ 2026-08-30
#   ★読むだけ★。走らせ方: pwsh -NoProfile -ExecutionPolicy Bypass -File tools\measure-xml2.ps1
$ErrorActionPreference = 'Continue'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$tmp = Join-Path $env:TEMP ('exally-xml2-' + [System.Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $tmp | Out-Null
try {
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Worksheets.Item(1)
  $xml = @'
<?xml version="1.0" encoding="UTF-8"?>
<people>
  <person><name>ta</name><age>20</age></person>
  <person><name>ni</name><age>30</age></person>
</people>
'@
  $p = Join-Path $tmp 'people.xml'
  [System.IO.File]::WriteAllText($p, $xml, [System.Text.Encoding]::UTF8)

  # ★[ref] は COM に 渡らない（実測＝Object reference not set）★
  #   ⇒ 先に 対応付けを 作り、`ImportXml` ではなく `XmlMaps.Add` → `Import` で 読む
  $map = $bk.XmlMaps.Add($p, 'people')
  Write-Output ('先に 作った 対応付け = ' + $map.Name)
  Write-Output ('  対応付ける前の IsExportable = ' + $map.IsExportable)
  # ★実Excel の 言うとおり★＝先に セルへ XPath を 結ぶ（Range.XPath.SetValue）
  $sh.Range('A1').XPath.SetValue($map, '/people/person/name')
  $sh.Range('B1').XPath.SetValue($map, '/people/person/age')
  Write-Output ('  結んだ後の A1 の XPath = ' + $sh.Range('A1').XPath.Value)
  Write-Output ('  結んだ後の IsExportable = ' + $map.IsExportable)
  $r = $map.Import($p, $true)
  Write-Output ("XmlMap.Import の 戻り（0=成功） = " + $r)
  Write-Output ("対応付けの 数 = " + $bk.XmlMaps.Count)
  if ($null -ne $map) {
    Write-Output ("作られた 対応付け = " + $map.Name)
    Write-Output ("  IsExportable = " + $map.IsExportable)
    Write-Output ("  RootElementName = " + $map.RootElementName)
  }
  foreach ($a in @('A1', 'B1', 'A2', 'B2', 'A3', 'B3', 'C1')) {
    Write-Output ("  " + $a + " = '" + $sh.Range($a).Text + "'")
  }
  Write-Output ("表の 数（ListObjects） = " + $sh.ListObjects.Count)
  if ($sh.ListObjects.Count -gt 0) {
    Write-Output ("  表の 名前 = " + $sh.ListObjects.Item(1).Name)
    Write-Output ("  表の 場所 = " + $sh.ListObjects.Item(1).Range.Address($false, $false))
  }

  $out = Join-Path $tmp 'out.xml'
  if ($bk.XmlMaps.Count -gt 0 -and $bk.XmlMaps.Item(1).IsExportable) {
    $e = $bk.XmlMaps.Item(1).Export($out, $true)
    Write-Output ("Export の 戻り（0=成功） = " + $e)
    $t = [System.IO.File]::ReadAllText($out)
    Write-Output ("  出た 字の 長さ = " + $t.Length)
    $one = $t.Replace([char]13, ' ').Replace([char]10, ' ')
    while ($one.Contains('  ')) { $one = $one.Replace('  ', ' ') }
    Write-Output ("  中身 = " + $one)
  } else {
    Write-Output 'Export = ★IsExportable が False なので 出せない★'
  }

  # ★書き変えてから 出す★＝出た 中身が 変わるか
  $sh.Range('A2').Value2 = 'TA2'
  if ($bk.XmlMaps.Count -gt 0 -and $bk.XmlMaps.Item(1).IsExportable) {
    $out2 = Join-Path $tmp 'out2.xml'
    $bk.XmlMaps.Item(1).Export($out2, $true) | Out-Null
    $t2 = [System.IO.File]::ReadAllText($out2)
    $one2 = $t2.Replace([char]13, ' ').Replace([char]10, ' ')
    while ($one2.Contains('  ')) { $one2 = $one2.Replace('  ', ' ') }
    Write-Output ('★セルを 直してから 出すと★ = ' + $one2)
  }
  $bk.Close($false) | Out-Null
} finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
  Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
