# measure-theme.ps1 — ★テーマ（配色・フォント・効果）と 背景・ふりがな★ 2026-08-30
#   ★読むだけ★＝新しいブックで 測り、保存せずに 閉じる。
#   走らせ方: pwsh -NoProfile -ExecutionPolicy Bypass -File tools\measure-theme.ps1

$ErrorActionPreference = 'Continue'
function 出す($名, $値) { Write-Output ("{0,-44} = {1}" -f $名, $値) }
function BGRを色に($v) {
  if ($null -eq $v) { return '(なし)' }
  $n = [int]$v
  $b = ($n -shr 16) -band 255; $g = ($n -shr 8) -band 255; $r = $n -band 255
  return ('#{0:X2}{1:X2}{2:X2}  (BGRの数 {3})' -f $r, $g, $b, $n)
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Worksheets.Item(1)
  $th = $bk.Theme

  Write-Output '=== 既定の テーマの 12色（ThemeColorScheme）==='
  # 1..12 = dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink
  $名 = @('dk1(濃い1)', 'lt1(薄い1)', 'dk2(濃い2)', 'lt2(薄い2)',
    'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
    'hlink(リンク)', 'folHlink(見たリンク)')
  for ($i = 1; $i -le 12; $i++) {
    try { 出す ("  " + $名[$i - 1]) (BGRを色に $th.ThemeColorScheme.Colors($i).RGB) }
    catch { 出す ("  " + $名[$i - 1]) ('★読めない★ ' + $_.Exception.Message.Trim()) }
  }

  Write-Output ''
  Write-Output '=== 既定の テーマの フォント（ThemeFontScheme）==='
  出す '見出し（MajorFont ラテン）' $th.ThemeFontScheme.MajorFont.Item(1).Name
  出す '本文（MinorFont ラテン）' $th.ThemeFontScheme.MinorFont.Item(1).Name
  # 2=東アジア
  try { 出す '見出し（東アジア）' $th.ThemeFontScheme.MajorFont.Item(2).Name } catch { 出す '見出し（東アジア）' '★読めない★' }
  try { 出す '本文（東アジア）' $th.ThemeFontScheme.MinorFont.Item(2).Name } catch { 出す '本文（東アジア）' '★読めない★' }

  Write-Output ''
  Write-Output '=== セルの 既定の フォント（テーマの 本文が 入るか）==='
  出す 'Application.StandardFont' $xl.StandardFont
  出す 'Application.StandardFontSize' $xl.StandardFontSize
  出す 'A1 の Font.Name' $sh.Range('A1').Font.Name
  出す 'A1 の Font.Size' $sh.Range('A1').Font.Size

  Write-Output ''
  Write-Output '=== 塗りつぶしの 色を テーマの 色で 指定する ==='
  # xlThemeColorAccent1 = 5
  $sh.Range('A1').Interior.ThemeColor = 5
  出す 'accent1 で 塗った A1 の 実際の 色' (BGRを色に $sh.Range('A1').Interior.Color)
  $sh.Range('A2').Interior.ThemeColor = 5
  $sh.Range('A2').Interior.TintAndShade = 0.4
  出す '  ＋ TintAndShade 0.4（薄く）' (BGRを色に $sh.Range('A2').Interior.Color)
  $sh.Range('A3').Interior.ThemeColor = 5
  $sh.Range('A3').Interior.TintAndShade = -0.25
  出す '  ＋ TintAndShade -0.25（濃く）' (BGRを色に $sh.Range('A3').Interior.Color)

  Write-Output ''
  Write-Output '=== 背景（ページ レイアウト→背景）==='
  出す 'Worksheet.SetBackgroundPicture が 在るか' ($null -ne $sh.PSObject.Methods['SetBackgroundPicture'])
  出す '★背景は 刷られるか★（実Excel の 決まり）' '★画面だけ。紙には 刷られない★'

  Write-Output ''
  Write-Output '=== ふりがな（ホーム→フォント→ふりがなの表示）==='
  $sh.Range('B1').Value2 = '東京都'
  try {
    出す 'Phonetics.Visible（はじめ）' $sh.Range('B1').Phonetics.Visible
    $sh.Range('B1').Phonetics.Visible = $true
    出す 'Visible=True に した後' $sh.Range('B1').Phonetics.Visible
    出す '  ふりがなの 数' $sh.Range('B1').Phonetics.Count
    if ($sh.Range('B1').Phonetics.Count -gt 0) {
      出す '  1つ目の ふりがな' $sh.Range('B1').Phonetics.Item(1).Text
    }
    出す '  ふりがなの 種類（CharacterType）' $sh.Range('B1').Phonetics.CharacterType
    出す '  ふりがなの 大きさ' $sh.Range('B1').Phonetics.Font.Size
    出す '  ふりがなの 行の 高さは 変わるか' $sh.Range('B1').RowHeight
  } catch { 出す 'ふりがな' ('★読めない★ ' + $_.Exception.Message.Trim()) }

  $sh.Range('B2').Value2 = 'ABC'
  try {
    $sh.Range('B2').Phonetics.Visible = $true
    出す '★英語に ふりがなを 出すと★ 数' $sh.Range('B2').Phonetics.Count
  } catch { 出す '英語の ふりがな' ('★出来ない★ ' + $_.Exception.Message.Trim()) }

  $bk.Close($false) | Out-Null
} finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
