# toru-vba-no-tobira-ga-aite-iru-ka.ps1
#   -- ★VBA の 扉が 開いて いるか（★控え★）★（102）（2026-09-22）
#
#  ★★なぜ★★
#    `tests/fixtures/vba-sample.xlsm` を 実Excel に 開かせたら
#      `Open` 投げず ／ `HasVBProject` ★True★ ／ ★でも VBProject の 部品 0個★
#    ⇒★0個の 訳が 2通り 在ります★
#        ㋐★材料の 側★ ... Excel が この 材料の マクロを 読めて いない
#        ㋑★設定の 側★ ... 「VBA プロジェクト オブジェクト モデルへの アクセスを 信頼する」が 切
#    ⇒★分けずに 「材料が 悪い」と 書いたら 嘘に なります★
#    ⇒★控えを 置きます★＝★実Excel が 作った ばかりの 本★で 部品を 数える
#        1個以上 見えたら ⇒ 扉は 開いて いる ⇒ ★0個は 材料の 側★
#        0個なら       ⇒ ★設定の 側★＝材料の せいでは ない
#
#  ★★読むだけ★★（新しい 本を 作って 閉じるだけ・保存しません）
#
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具>

# ══ ★★2026-09-22 ── ★最初の 控えの 置き方が 間違って いました★★ ══
#   私は「★実Excel が 作った ばかりの 本★で 部品を 数える」と しました。
#   ⇒★新しい 本には そもそも VBA の 入れ物が 在りません★
#     （実測 ... `HasVBProject` ★False★ ／ 部品 0個）
#   ⇒★扉が 開いて いるかは 1つも 分かりません★＝★控えに なって いません★
#   ⇒記憶「★掃く 窓は 押す前に 隙間を 計算して から 置く★」を また 踏みました
#
# ★★代わりに 設定を 直に 読みます★★（★これなら 1回で 分かれます★）
#   `HKCU:\Software\Microsoft\Office\<版>\Excel\Security` の `AccessVBOM`
#     1 なら ★開いて いる★ ／ 0 か 書かれて いない なら ★切★
#   ★実測（2026-09-22）★ ... Office 16.0 ／ ★書かれて いない＝既定＝切★
#   ⇒★`vba-sample.xlsm` の 部品 0個は ★設定の 側★★＝★材料は 無実★
#   ⇒★COM では マクロの 中を 数えられません★（設定を 変えるのは ★守りを 開ける★事）
#     ⇒★変えません★。中を 読むのは ★うちの 台（`lib/vba.js`）の 仕事★に します。
foreach ($v in '16.0', '15.0', '14.0') {
  $k = 'HKCU:' + [string][char]92 + 'Software' + [string][char]92 + 'Microsoft' + [string][char]92 + 'Office' + [string][char]92 + $v + [string][char]92 + 'Excel' + [string][char]92 + 'Security'
  if (Test-Path $k) {
    $a = (Get-ItemProperty -Path $k -ErrorAction SilentlyContinue).AccessVBOM
    $札 = if ($null -eq $a) { '（書かれて いません＝既定＝0＝★切★）' } else { [string]$a }
    Write-Host ('★Office ' + $v + ' の AccessVBOM★ ' + $札)
  }
}
Write-Host '⇒★1 なら 開いて いる／0 か 無し なら 切★'
Write-Host ''

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
if ($数1 -ne 0) { Write-Host '★★Excel が 動いて います★★'; exit 3 }
$xl = New-Object -ComObject Excel.Application
$bk = $null
try {
  $xl.Visible = $false; $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  Write-Host ('★新しい 本★ 板 ' + [string]$bk.Sheets.Count + '枚 ／ HasVBProject ' + [string]$bk.HasVBProject)
  try {
    $n = [int]$bk.VBProject.VBComponents.Count
    Write-Host ('★VBProject の 部品★ ' + $n + '個')
    for ($k = 1; $k -le $n; $k++) {
      $c = $bk.VBProject.VBComponents.Item($k)
      Write-Host ('   ' + [string]$c.Name + ' ／ 種類 ' + [string]$c.Type)
      $c = $null
    }
    if ($n -gt 0) { Write-Host '⇒★「信頼する」は 入って います★＝0個の 材料は ★材料の 側★' }
    else { Write-Host '⇒★新しい 本でも 0個★＝★設定の 側の 話★（材料の せいでは ない）' }
  } catch {
    Write-Host ('★VBProject が 読めません★ ' + $_.Exception.Message)
    Write-Host '⇒★「信頼する」が 切って あります★＝材料の 0個は ★判じられません★'
  }
  $bk.Close($false); $bk = $null
} finally {
  if ($null -ne $bk) { $bk.Close($false) }
  $bk = $null
  $xl.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl); $xl = $null
  $t = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 300)) { Start-Sleep -Milliseconds 250 }
  Write-Host ('★Excel は ' + [math]::Round($t.Elapsed.TotalSeconds,1) + '秒で 消えました★')
}
