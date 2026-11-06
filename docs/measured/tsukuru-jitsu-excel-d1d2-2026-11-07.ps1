# tsukuru-jitsu-excel-d1d2-2026-11-07.ps1  (Exally-値 席)
#  ★なぜ★ 346 の紙（09-08）の (D1,D2) の行で 台と外れた 3関数（BESSELJ SEQUENCE WEEKNUM）が
#    ★本物の外れ★か ★紙の盤面の古さ★かを決める。346 と同じ盤面・同じ道（H1 に .Formula）と
#    今の道（H1 に Formula2）の両方で打ち直す。
#  盤面: A1:A5=1..5 / B1:B5=2,4,6,8,10 / D1==DATE(2024,1,1) / D2==DATE(2026,1,1)
#  既存の Excel には触らない（Hwnd→PID の門・hikae-nashi-48.ps1 と同じ）
param([string]$Out)
if (-not $Out) { exit 6 }
if ($PSVersionTable.PSVersion.Major -ne 5) { exit 8 }
$fx = @('=BESSELJ(D1,D2)', '=SEQUENCE(D1,D2)', '=SEQUENCE(D1,D2,1)', '=SEQUENCE(D1,D2,0)', '=WEEKNUM(D1,D2)', '=ABS(D1)')
Add-Type -Namespace W -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
$before = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" | ForEach-Object { [int]$_.ProcessId })
$xl = New-Object -ComObject Excel.Application
$xpid = [uint32]0; [void][W.U]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$xpid)
Write-Host ('前 ' + ($before -join ',') + ' / 立てた ' + $xpid)
if ($xpid -eq 0 -or $before -contains [int]$xpid) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl); exit 4 }
$errs = @{ -2146826281 = '#DIV/0!'; -2146826252 = '#NUM!'; -2146826246 = '#N/A'; -2146826273 = '#VALUE!';
  -2146826265 = '#REF!'; -2146826259 = '#NAME?'; -2146826288 = '#NULL!'; -2146826238 = '#CALC!'; -2146826245 = '#SPILL!' }
$inv = [System.Globalization.CultureInfo]::InvariantCulture
function V($v) { if ($null -eq $v) { return '' }; if ($v -is [int]) { if ($errs.ContainsKey($v)) { return $errs[$v] }; return $v.ToString($inv) }
  if ($v -is [double]) { return $v.ToString('R', $inv) }; return [string]$v }
$rows = New-Object System.Collections.Generic.List[string]
$rows.Add("# 実Excel $($xl.Version) build $($xl.Build)・$((Get-Date).ToString('yyyy-MM-dd HH:mm'))・346 と同じ盤面・H1")
$rows.Add("道`t式`tValue2`t式の字")
try {
  $xl.Visible = $false; $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add(); $sh = $bk.Sheets.Item(1)
  for ($r = 1; $r -le 5; $r++) { $sh.Cells.Item($r, 1).Value2 = $r; $sh.Cells.Item($r, 2).Value2 = 2 * $r }
  $sh.Range('D1').Formula = '=DATE(2024,1,1)'; $sh.Range('D2').Formula = '=DATE(2026,1,1)'
  foreach ($f in $fx) {
    foreach ($way in 'Formula', 'Formula2') {
      $c = $sh.Range('H1')
      try { if ($way -eq 'Formula') { $c.Formula = $f } else { $c.Formula2 = $f }; $xl.Calculate()
        $rows.Add("$way`t$f`t$(V $c.Value2)`t$($c.Formula2)") } catch { $rows.Add("$way`t$f`t打てません`t") }
      [void]$sh.Range('H1:Z40').Clear()
    }
  }
  $bk.Close($false)
} finally {
  $xl.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl); [GC]::Collect(); [GC]::WaitForPendingFinalizers(); Start-Sleep 3
  if (Get-Process -Id $xpid -ErrorAction SilentlyContinue) { Stop-Process -Id $xpid -Force; Write-Host ('自分の ' + $xpid + ' だけ止めた') }
}
[System.IO.File]::WriteAllLines($Out, $rows, (New-Object System.Text.UTF8Encoding($false)))
Write-Host ('後 ' + (@(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" | ForEach-Object { $_.ProcessId }) -join ','))
