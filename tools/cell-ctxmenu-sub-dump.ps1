# ★実Excel の セルの 右クリックの ★▸ の 中★ を 撮る★（読むだけ・保存しない）
#   ・UI Automation では メニューが 1個も 見えない ⇒ ★絵で 読む★
#   ・開き方は ★アクセスキー★（フィルター(E)／並べ替え(O)／リンク(I)）＝お客さんの 道
#   ・★司さんの 画面（デスクトップ）が 起きている 時だけ 動きます★
param([string]$Dir='C:\Users\zeroa\AppData\Local\Temp\claude\C--WINDOWS-System32-WindowsPowerShell-v1-0\5b4e50e6-20a1-4af5-8ffb-8b6d6ca3f52b\scratchpad\shot')
$ErrorActionPreference='Stop'
Add-Type -AssemblyName UIAutomationClient; Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Drawing; Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System; using System.Runtime.InteropServices;
public class S {
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x,int y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint f,uint a,uint b,uint c,IntPtr d);
  [DllImport("user32.dll")] public static extern void keybd_event(byte k,byte s,uint f,IntPtr e);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, IntPtr p);
  [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint a,uint b,bool f);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h,int n);
  [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
  public static void ForceFront(IntPtr h){
    IntPtr fg=GetForegroundWindow(); uint me=GetCurrentThreadId();
    uint o=(fg==IntPtr.Zero)?me:GetWindowThreadProcessId(fg,IntPtr.Zero);
    if(o!=me) AttachThreadInput(me,o,true);
    ShowWindow(h,9); ShowWindow(h,3); BringWindowToTop(h); SetForegroundWindow(h);
    if(o!=me) AttachThreadInput(me,o,false);
  }
  public static void Key(byte k){ keybd_event(k,0,0,IntPtr.Zero); System.Threading.Thread.Sleep(70); keybd_event(k,0,2,IntPtr.Zero); }
  public static void R(int x,int y){ SetCursorPos(x,y); System.Threading.Thread.Sleep(300);
    mouse_event(0x0008,0,0,0,IntPtr.Zero); System.Threading.Thread.Sleep(90); mouse_event(0x0010,0,0,0,IntPtr.Zero); }
}
"@
[void][S]::SetProcessDPIAware()
$AE=[System.Windows.Automation.AutomationElement]
$CT=[System.Windows.Automation.ControlType]
function Shot($name,$dir){
  $b=[System.Windows.Forms.Screen]::PrimaryScreen.Bounds
  $bmp=New-Object System.Drawing.Bitmap($b.Width,$b.Height)
  ([System.Drawing.Graphics]::FromImage($bmp)).CopyFromScreen(0,0,0,0,$bmp.Size)
  $p=Join-Path $dir $name; $bmp.Save($p,[System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
  return $p
}
$xl=$null
try{
  if(-not (Test-Path $Dir)){ New-Item -ItemType Directory -Force $Dir | Out-Null }
  $xl=New-Object -ComObject Excel.Application; $xl.Visible=$true; $xl.DisplayAlerts=$false
  $wb=$xl.Workbooks.Add(); Start-Sleep -Milliseconds 3000
  try{$xl.WindowState=-4137}catch{}
  [S]::ForceFront([IntPtr]$xl.Hwnd); Start-Sleep -Milliseconds 1500
  $sh=$wb.Worksheets.Item(1)
  # ★中身を 少し 入れる★＝フィルター/並べ替えが 灰に ならないように
  $sh.Range("B2").Value2='くだもの'; $sh.Range("B3").Value2='りんご'; $sh.Range("B4").Value2='みかん'
  $c=$sh.Range("B3"); $c.Select()|Out-Null; Start-Sleep -Milliseconds 400
  $w=$xl.ActiveWindow
  $x=[int]$w.PointsToScreenPixelsX($c.Left+$c.Width/2); $y=[int]$w.PointsToScreenPixelsY($c.Top+$c.Height/2)

  # ★押す前に その 点に 居るのが Excel か 確かめる★
  for($t=1;$t -le 5;$t++){
    $u=$AE::FromPoint((New-Object System.Windows.Point([double]$x,[double]$y)))
    $ww=$u; while($ww -ne $null -and $ww.Current.ControlType -ne $CT::Window){ $ww=[System.Windows.Automation.TreeWalker]::ControlViewWalker.GetParent($ww) }
    $nm = if($ww){ $ww.Current.ClassName } else { '不明' }
    if($nm -eq 'XLMAIN'){ break }
    "($t 回目) その 点は [$nm] ⇒ Excel を 前へ 出し直す"
    [S]::ForceFront([IntPtr]$xl.Hwnd); Start-Sleep -Milliseconds 1200
  }
  if($nm -ne 'XLMAIN'){ throw "★Excel を 前面に 出せませんでした（$nm）★" }

  # ▸ ごとに 撮る（アクセスキーで 開く＝お客さんの 道）
  $kumi=@(
    @{ na='フィルター'; key=0x45; f='excel_sub_filter.png' },   # E
    @{ na='並べ替え';   key=0x4F; f='excel_sub_sort.png'   },   # O
    # ★リンク(I) は 挿入(I) と ★同じ アクセスキー★★＝1回 押すと 挿入の方に 当たる
    #   ⇒ ★2回 押して リンクへ 移り、→ で 開く★（2026-09-03 実測で 分かった）
    @{ na='リンク';     key=0x49; f='excel_sub_link.png'; nikai=$true }
  )
  foreach($k in $kumi){
    # ★前の 回の 名残を 完全に 落とす★
    #   （1回目は Esc 1つ だけで、親メニューが 開いたまま 右クリック ⇒ 閉じただけに なり、
    #     次に 押した 文字が ★セルに 入って しまった★＝2026-09-03 実測）
    [S]::Key(0x1B); Start-Sleep -Milliseconds 250
    [S]::Key(0x1B); Start-Sleep -Milliseconds 250
    [S]::Key(0x1B); Start-Sleep -Milliseconds 400
    $c.Select() | Out-Null                      # ★選び直す（編集中なら 抜ける）★
    Start-Sleep -Milliseconds 400
    [S]::R($x,$y); Start-Sleep -Milliseconds 1500
    [S]::Key([byte]$k.key); Start-Sleep -Milliseconds 900
    if($k.nikai){ [S]::Key([byte]$k.key); Start-Sleep -Milliseconds 900
                  [S]::Key(0x27); Start-Sleep -Milliseconds 900 }   # → で ▸ を 開く
    Start-Sleep -Milliseconds 600
    $p=Shot $k.f $Dir
    "$($k.na) ▸ … $p"
    [S]::Key(0x1B); Start-Sleep -Milliseconds 300
  }
  [S]::Key(0x1B)
}catch{ "★落ちた★ $($_.Exception.Message)" }
finally{ if($xl){ try{foreach($z in @($xl.Workbooks)){$z.Close($false)}}catch{}; try{$xl.Quit()}catch{} } }
