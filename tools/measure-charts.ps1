# measure-charts.ps1 - 実Excel(COM)で ★グラフの種類ごとの 既定★ を 測る
#   pwsh -File tools/measure-charts.ps1
#   ★読むだけ★（開いたブックは 保存せずに 閉じる）
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
# 材料（見出し1列 + 数2列 + 3行）
$ws.Range('A1').Value2 = 'あ'; $ws.Range('B1').Value2 = 1;  $ws.Range('C1').Value2 = 10
$ws.Range('A2').Value2 = 'い'; $ws.Range('B2').Value2 = 4;  $ws.Range('C2').Value2 = 20
$ws.Range('A3').Value2 = 'う'; $ws.Range('B3').Value2 = 9;  $ws.Range('C3').Value2 = 30
$ws.Range('A4').Value2 = 'え'; $ws.Range('B4').Value2 = 16; $ws.Range('C4').Value2 = 40

$種類 = [ordered]@{
  '集合縦棒(既定)'      = 51
  '散布図'              = -4169
  '散布図(線つき)'      = 74
  'バブル'              = 15
  'レーダー'            = -4151
  'レーダー(印つき)'    = 82
  'ヒストグラム'        = 118
  '箱ひげ図'            = 121
  'ツリーマップ'        = 117
  'サンバースト'        = 120
  'ウォーターフォール'  = 119
  'じょうご'            = 123
  '株価(高安終)'        = 88
  '等高線'              = -4109
  'ドーナツ'            = -4120
  '面'                  = 1
}
foreach ($k in $種類.Keys) {
  $t = $種類[$k]
  try {
    $ws.Range('A1:C4').Select() | Out-Null
    $co = $ws.Shapes.AddChart2(-1, $t)
    $ch = $co.Chart
    $幅 = [math]::Round($co.Width,2); $高 = [math]::Round($co.Height,2)
    $凡例 = $ch.HasLegend
    $題 = $ch.HasTitle
    $題の字 = if ($題) { try { $ch.ChartTitle.Text } catch { '(読めない)' } } else { '' }
    $系列 = $ch.SeriesCollection().Count
    $軸 = try { $ch.Axes().Count } catch { -1 }
    '{0,-18} 番号={1,6}  大きさ={2}x{3}  凡例={4}  題={5}「{6}」 系列={7} 軸={8}' -f $k,$t,$幅,$高,$凡例,$題,$題の字,$系列,$軸
    $co.Delete()
  } catch {
    '{0,-18} 番号={1,6}  ★作れない★ {2}' -f $k,$t,$_.Exception.Message
  }
}
$wb.Close($false)
$xl.Quit()
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
