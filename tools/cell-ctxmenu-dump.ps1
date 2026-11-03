# 実Excel の ★セルの 右クリック★を「本物の マウス」で 開いて、
# 出た メニューを UI Automation で ★並び順のまま★ 取る。
#   ・読むだけ。新規ブック・保存しない。終わったら 必ず 閉じる。
#   ・★ShowPopup() は 使わない★（2026-09-02 に Excel が 2分 固まった）
#   ・★.Visible（CommandBars）も 使わない★（当てに ならない事が 実測で 分かっている）
# 出力: 引数 $Out（既定は scratchpad/cell_ctxmenu.txt）
#
# ★★大事：これは「司さんの 画面（デスクトップ）」で 走らせて ください★★
#   ・Claude のセッションから 走らせると ★何も 出ません★。
#     理由＝★そのセッションには 前面の 窓が 無い★（実測 2026-09-03）
#            GetForegroundWindow() が ★0★／「1」を 打っても A1 は 空のまま
#            ＝★作った キーも マウスも Excel に 届かない★
#   ・普通の PowerShell（Win+X → ターミナル）で これを 走らせれば 動きます。
#   ・読むだけ／新規ブック／保存しない／終わったら 閉じる。
param(
  [string]$Out = 'C:\Users\zeroa\AppData\Local\Temp\claude\C--WINDOWS-System32-WindowsPowerShell-v1-0\5b4e50e6-20a1-4af5-8ffb-8b6d6ca3f52b\scratchpad\cell_ctxmenu.txt'
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

# ★本物の マウス★（SendInput）＝JSで イベントを 投げるのとは 違う
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Mouse {
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int X, int Y);
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint dx, uint dy, uint d, IntPtr e);
  public const uint RIGHTDOWN = 0x0008, RIGHTUP = 0x0010, LEFTDOWN = 0x0002, LEFTUP = 0x0004;
  public static void RightClick(int x, int y) {
    SetCursorPos(x, y); System.Threading.Thread.Sleep(200);
    mouse_event(RIGHTDOWN, 0, 0, 0, IntPtr.Zero);
    System.Threading.Thread.Sleep(60);
    mouse_event(RIGHTUP, 0, 0, 0, IntPtr.Zero);
  }
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int n);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, IntPtr pid);
  [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint a, uint b, bool f);
  [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr h);
  [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
  public static void Front(IntPtr h) { ShowWindow(h, 3); SetForegroundWindow(h); }
  // ★Windows は 後ろの プロセスからの 前面化を 断る★
  //   ⇒ 今 前面に 居る スレッドに 自分を くっつけてから 頼むと 通る
  public static void ForceFront(IntPtr h) {
    IntPtr fg = GetForegroundWindow();
    uint me = GetCurrentThreadId();
    uint other = (fg == IntPtr.Zero) ? me : GetWindowThreadProcessId(fg, IntPtr.Zero);
    if (other != me) AttachThreadInput(me, other, true);
    ShowWindow(h, 9); ShowWindow(h, 3);
    BringWindowToTop(h); SetForegroundWindow(h);
    if (other != me) AttachThreadInput(me, other, false);
  }
  // ★キーボードでも 開ける★＝Shift+F10 / メニューキー（お客さんが 使う 道）
  public static void ShiftF10() {
    keybd_event(0x10,0,0,IntPtr.Zero);              // Shift ↓
    keybd_event(0x79,0,0,IntPtr.Zero);              // F10 ↓
    System.Threading.Thread.Sleep(60);
    keybd_event(0x79,0,2,IntPtr.Zero);              // F10 ↑
    keybd_event(0x10,0,2,IntPtr.Zero);              // Shift ↑
  }
  public static void AppsKey() {
    keybd_event(0x5D,0,0,IntPtr.Zero); System.Threading.Thread.Sleep(60); keybd_event(0x5D,0,2,IntPtr.Zero);
  }
  public static void Esc() { keybd_event(0x1B,0,0,IntPtr.Zero); System.Threading.Thread.Sleep(30); keybd_event(0x1B,0,2,IntPtr.Zero); }
  [DllImport("user32.dll")] public static extern void keybd_event(byte k, byte s, uint f, IntPtr e);
  public static void LeftClick(int x, int y) {
    SetCursorPos(x, y); System.Threading.Thread.Sleep(150);
    mouse_event(LEFTDOWN, 0, 0, 0, IntPtr.Zero);
    System.Threading.Thread.Sleep(50);
    mouse_event(LEFTUP, 0, 0, 0, IntPtr.Zero);
  }
}
"@

$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$CT = [System.Windows.Automation.ControlType]

# ★画面の 拡大率に 合わせる★（これが 無いと SetCursorPos が 別の 場所を 押す
#   ＝2026-09-03 実測。後ろの Chrome の メニューが 開いた）
[void][Mouse]::SetProcessDPIAware()

$lines = @()
$xl = $null
try {
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $true
  $xl.DisplayAlerts = $false
  $wb = $xl.Workbooks.Add()
  Start-Sleep -Milliseconds 3000
  $lines += "Excel $($xl.Version) build $($xl.Build)"
  $lines += "測り方: ★本物の マウスで セルを 右クリック★ → UI Automation で メニューを 歩いた"
  $lines += "        （ShowPopup は 使っていない／CommandBars の .Visible も 使っていない）"
  $lines += ""

  # ★C5 を 選んで、その セルの 画面の 場所を 取る★
  # ★Excel を 前面に 出す★（前は 後ろの Chrome を 右クリックしていた＝2026-09-03 実測）
  try { $xl.WindowState = -4137 } catch {}          # xlMaximized
  [Mouse]::ForceFront([IntPtr]$xl.Hwnd)
  Start-Sleep -Milliseconds 1500

  $sh = $wb.Worksheets.Item(1)
  $cell = $sh.Range("C5")
  $cell.Select() | Out-Null
  Start-Sleep -Milliseconds 500

  # ポイント → 画面の 点（Excel は ポイント。1pt = 96/72 px を 画面の DPI で 直す）
  $win = $xl.ActiveWindow
  $ptX = $win.PointsToScreenPixelsX($cell.Left + $cell.Width / 2)
  $ptY = $win.PointsToScreenPixelsY($cell.Top + $cell.Height / 2)
  $lines += "右クリックした 場所: C5 → 画面(${ptX}, ${ptY})"
  # ★その 点に 居るのが 本当に Excel か 先に 確かめる★（違う 窓を 押さない為）
  $under = $AE::FromPoint((New-Object System.Windows.Point([double]$ptX, [double]$ptY)))
  $w = $under
  while ($w -ne $null -and $w.Current.ControlType -ne $CT::Window) {
    $w = [System.Windows.Automation.TreeWalker]::ControlViewWalker.GetParent($w)
  }
  $mado = "不明"
  for ($try = 1; $try -le 5; $try++) {
    $under = $AE::FromPoint((New-Object System.Windows.Point([double]$ptX, [double]$ptY)))
    $w = $under
    while ($w -ne $null -and $w.Current.ControlType -ne $CT::Window) {
      $w = [System.Windows.Automation.TreeWalker]::ControlViewWalker.GetParent($w)
    }
    $mado = if ($w) { $w.Current.Name + " / " + $w.Current.ClassName } else { "不明" }
    if ($mado -match 'XLMAIN') { break }
    $lines += "  ($try 回目) その 点は「$mado」だった ⇒ Excel を 前へ 出し直す"
    [Mouse]::ForceFront([IntPtr]$xl.Hwnd)
    Start-Sleep -Milliseconds 1200
  }
  $lines += "その 点に 居る 窓: $mado"
  if ($mado -notmatch 'XLMAIN') { throw "★Excel を 前面に 出せませんでした（$mado）＝押しません★" }

  # ★他の アプリの 開きっぱなしの メニューを 先に 閉じる★
  #   （2026-09-03 実測＝1回目に Chrome の メニューを 開いてしまい、
  #     2回目に それを「Excel の メニュー」と 読み違えかけた）
  [Mouse]::Esc(); Start-Sleep -Milliseconds 300
  [Mouse]::Front([IntPtr]$xl.Hwnd); Start-Sleep -Milliseconds 500

  # ★Excel の プロセス番号★＝これ以外の メニューは 数に 入れない
  $xlPid = $w.Current.ProcessId
  $lines += "Excel の プロセス番号: $xlPid"

  $root = $AE::RootElement
  $menuCond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Menu)

  # ★押す 前★に 在った メニューを 覚える（後から 出た物だけ 見る）
  $mae = @()
  foreach ($m in $root.FindAll($TS::Descendants, $menuCond)) {
    try { $mae += ($m.Current.ProcessId.ToString() + "|" + $m.Current.Name + "|" + $m.Current.BoundingRectangle.ToString()) } catch {}
  }
  $lines += "押す 前から 在った メニュー窓: $($mae.Count) 個"
  $walker0 = [System.Windows.Automation.TreeWalker]::ControlViewWalker
  $maeMado = @()
  $c0 = $walker0.GetFirstChild($root)
  while ($c0 -ne $null) {
    try { $maeMado += ($c0.Current.ProcessId.ToString() + "|" + $c0.Current.Name + "|" + $c0.Current.BoundingRectangle.ToString()) } catch {}
    $c0 = $walker0.GetNextSibling($c0)
  }
  $lines += "押す 前の トップレベル窓: $($maeMado.Count) 個"

  [Mouse]::RightClick([int]$ptX, [int]$ptY)
  Start-Sleep -Milliseconds 1500
  # ★マウスで 出なかった時は キーボードで 出す★（どちらも お客さんの 道）
  $imaMado = 0
  foreach ($m in $root.FindAll($TS::Descendants, $menuCond)) { $imaMado++ }
  $tsukatta = "マウスの 右クリック"
  $walkA = [System.Windows.Automation.TreeWalker]::ControlViewWalker
  $kazu = 0; $cA = $walkA.GetFirstChild($root)
  while ($cA -ne $null) { $kazu++; $cA = $walkA.GetNextSibling($cA) }
  if ($kazu -le $maeMado.Count) {
    [Mouse]::Front([IntPtr]$xl.Hwnd); Start-Sleep -Milliseconds 400
    [Mouse]::ShiftF10(); Start-Sleep -Milliseconds 1500
    $tsukatta = "Shift+F10（キーボード）"
  }
  $lines += "開け方: $tsukatta"

  # ★Excel の ポップアップは UIA の Menu では ない★（2026-09-03 実測＝0個だった）
  #   ⇒ ★プロセスが Excel の トップレベル窓を 全部 見て、押した後に 増えた物を 取る★
  $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
  function Get-TopWindows($root, $walker) {
    $r = @(); $c = $walker.GetFirstChild($root)
    while ($c -ne $null) { $r += $c; $c = $walker.GetNextSibling($c) }
    return $r
  }
  $atoMado = Get-TopWindows $root $walker
  $lines += "押した 後の トップレベル窓: $($atoMado.Count) 個"
  $lines += ""

  $total = 0; $mi = 0
  foreach ($m in $atoMado) {
    $cc = $null
    try { $cc = $m.Current } catch { continue }
    if ($cc.ProcessId -ne $xlPid) { continue }
    $key = ($cc.ProcessId.ToString() + "|" + $cc.Name + "|" + $cc.BoundingRectangle.ToString())
    if ($maeMado -contains $key) { continue }
    $mi++
    $lines += "== 出た 窓 $mi : 「$($cc.Name)」 class=$($cc.ClassName) type=$($cc.ControlType.ProgrammaticName) =="
    $kids = $m.FindAll($TS::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
    $n = 0
    foreach ($it in $kids) {
      $c2 = $null
      try { $c2 = $it.Current } catch { continue }
      if ([string]::IsNullOrWhiteSpace($c2.Name)) { continue }
      $n++; $total++
      $en = if ($c2.IsEnabled) { "押せる" } else { "押せない" }
      $lines += ("{0,3}. {1}`t{2}`t{3}" -f $n, $c2.Name, $en, $c2.ControlType.ProgrammaticName)
    }
    $lines += ""
  }
  $lines += "★合計 $total 個★"

  # ★閉じる★
  [Mouse]::Esc(); Start-Sleep -Milliseconds 300; [Mouse]::Esc()
} catch {
  $lines += "★落ちた★ $($_.Exception.Message)"
} finally {
  if ($xl) {
    try { $xl.DisplayAlerts = $false } catch {}
    try { foreach ($b in @($xl.Workbooks)) { $b.Close($false) } } catch {}
    try { $xl.Quit() } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch {}
  }
}
$lines | Set-Content -Path $Out -Encoding UTF8
$lines | ForEach-Object { $_ }
