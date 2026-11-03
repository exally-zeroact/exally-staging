# measure-python.ps1 — ★数式タブ Python (プレビュー) の 4個★ 2026-08-30
#   ★読むだけ★＝新しいブックで 測り、保存せずに 閉じる。
#   走らせ方: pwsh -NoProfile -ExecutionPolicy Bypass -File tools\measure-python.ps1

$ErrorActionPreference = 'Continue'
function 出す($名, $値) { Write-Output ("{0,-52} = {1}" -f $名, $値) }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Worksheets.Item(1)

  Write-Output '=== Excel の Python（=PY(...)）==='
  出す 'この Excel の 版' $xl.Version
  出す '  ビルド' $xl.Build

  # PY 関数が 在るか（式を 入れて 答えを 見る）
  foreach ($f in @('=PY("1+1",0)', '=PY(1+1)', '=_xlfn.PY("1+1",0)')) {
    try {
      $sh.Range('A1').Formula = $f
      出す ("式 " + $f) ("'" + $sh.Range('A1').Text + "'")
    } catch { 出す ("式 " + $f) ('★入らない★ ' + $_.Exception.Message.Trim()) }
  }
  $sh.Range('A1').Clear() | Out-Null

  # Excel が 知っている 関数か（日本語版でも 英語名で 聞く）
  foreach ($n in @('PY', 'SUM', 'XLOOKUP')) {
    try {
      $v = $xl.Evaluate($n + '(1,1)')
      出す ("Evaluate " + $n + "(1,1)") $v
    } catch { 出す ("Evaluate " + $n + "(1,1)") ('★出来ない★ ' + $_.Exception.Message.Trim()) }
  }

  Write-Output ''
  Write-Output '=== Python の 決め（COM から 触れるか）==='
  foreach ($n in @('PythonEnabled', 'PythonSettings', 'CalculationState')) {
    try {
      $v = $xl.PSObject.Properties[$n]
      if ($null -eq $v) { 出す ("Application." + $n) '★COM に 無い★' }
      else { 出す ("Application." + $n) $v.Value }
    } catch { 出す ("Application." + $n) ('★読めない★ ' + $_.Exception.Message.Trim()) }
  }

  Write-Output ''
  Write-Output '=== Python の 答えの 型（データ型）==='
  出す 'Python の 答えは 何で 返るか' '★Microsoft の クラウドで 動く（機械からは 動かせない）★'
  出す '  だから 測れる 事' '★式の 形（=PY(...)）と、入るか どうかだけ★'

  Write-Output ''
  Write-Output '=== 参考：うちが 出来る 事の 確かめ ==='
  出す 'ブラウザで Python が 動くか' '★そのままでは 動かない（外の 仕組みが 要る）★'
  出す '  うちが 出来る 事' '★Python の 式を 覚えて 見せる／同じ 事を うちの 関数で する★'

  $bk.Close($false) | Out-Null
} finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
