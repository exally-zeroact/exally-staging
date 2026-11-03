# 「記号と特殊文字」の 特殊文字タブに 何が 並ぶかを 実Excel の 一覧から 測る（読むだけ）
# ★COM から この一覧は 読めない★ので、代わりに ★AutoCorrect の 置き換え表★（実物）を 見る。
$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add()
$ac = $xl.AutoCorrect
"AutoCorrect の 置き換え数 = {0}" -f $ac.Entries.Count
$n=0
foreach($e in $ac.Entries){
  if($n -ge 12){ break }
  "  {0} → {1}" -f $e.Name, $e.Value
  $n++
}
# 特殊文字（Excelの 特殊文字タブに 在る物）の 文字コードを 実際に 入れて 読み返す
$ws=$wb.Worksheets.Item(1)
$表 = @{
 '著作権'=0x00A9; '登録商標'=0x00AE; '商標'=0x2122; 'セクション'=0x00A7; '段落'=0x00B6
 '省略記号'=0x2026; 'emダッシュ'=0x2014; 'enダッシュ'=0x2013; '左二重引用符'=0x201C; '右二重引用符'=0x201D
}
foreach($k in $表.Keys){
  $ws.Range('A1').Value2 = [char]$表[$k]
  "  {0} = U+{1:X4} 読み返し='{2}'" -f $k, $表[$k], $ws.Range('A1').Text
}
$wb.Close($false); $xl.Quit()
