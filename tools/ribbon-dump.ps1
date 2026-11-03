# 実Excelのリボンを UI Automation で歩いて、タブ→グループ→ボタンを「並び順のまま」取る。
# 読むだけ。ブックは新規・保存しない。
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
$out = 'C:\Users\zeroa\AppData\Local\Temp\claude\C--WINDOWS-System32-WindowsPowerShell-v1-0\6251529a-d1e7-44f1-9575-b9d6bd6a93e3\scratchpad\ribbon_result.txt'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$wb = $xl.Workbooks.Add()
Start-Sleep -Milliseconds 2500

$AE   = [System.Windows.Automation.AutomationElement]
$TS   = [System.Windows.Automation.TreeScope]
$CT   = [System.Windows.Automation.ControlType]
$root = $AE::FromHandle([IntPtr]$xl.Hwnd)

$lines = @()
$lines += "Excel $($xl.Version) build $($xl.Build) / 日本語UI"
$lines += "測り方: UI Automation でリボンを歩いた（並び順そのまま）"
$lines += ""

# タブ（TabItem）を並び順で取る
$cond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::TabItem)
$tabs = $root.FindAll($TS::Descendants, $cond)
$lines += "★リボンのタブ: $($tabs.Count) 個★"
$tabNames = @()
foreach ($t in $tabs) { $tabNames += $t.Current.Name }
$lines += "  " + ($tabNames -join " ／ ")
$lines += ""

$grpCond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Group)
$totalGroups = 0; $totalCtrls = 0

foreach ($t in $tabs) {
  $name = $t.Current.Name
  try {
    $sip = $t.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
    $sip.Select()
  } catch { }
  Start-Sleep -Milliseconds 700

  $lines += "=" * 58
  $lines += "■ タブ: $name"
  $groups = $root.FindAll($TS::Descendants, $grpCond)
  $gi = 0
  foreach ($g in $groups) {
    $gn = $g.Current.Name
    if ([string]::IsNullOrWhiteSpace($gn)) { continue }
    # リボン以外(シート等)のGroupを除くため、上部にある物だけ
    $r = $g.Current.BoundingRectangle
    if ($r.Top -gt 260 -or $r.Height -le 0) { continue }
    $gi++; $totalGroups++
    $kids = @()
    foreach ($c in $g.FindAll($TS::Descendants, [System.Windows.Automation.Condition]::TrueCondition)) {
      $cn = $c.Current.Name
      $ct = $c.Current.ControlType.ProgrammaticName -replace 'ControlType\.',''
      if (-not [string]::IsNullOrWhiteSpace($cn) -and $ct -ne 'Text') {
        $kids += "$cn"
      }
    }
    $kids = $kids | Select-Object -Unique
    $totalCtrls += $kids.Count
    $lines += ("  [{0}] グループ: {1}  （{2}個）" -f $gi, $gn, $kids.Count)
    if ($kids.Count) { $lines += "       " + ($kids -join " / ") }
  }
  if ($gi -eq 0) { $lines += "  （グループを取れなかった）" }
}

$lines += "=" * 58
$lines += "★合計: タブ $($tabs.Count) 個 / グループ $totalGroups 個 / 部品 $totalCtrls 個★"

$wb.Close($false)
$xl.Quit()
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
$lines | Set-Content -Path $out -Encoding UTF8
"書き出した: $out"
"タブ $($tabs.Count) / グループ $totalGroups / 部品 $totalCtrls"
