# 実Excel の ★「挿入(I)...」「削除(D)...」の 窓★を 撮る（読むだけ・保存しない）
#   ・新規ブック／読むだけ／保存しない／終わったら 閉じる
#   ・★本物の マウスで 右クリック → アクセスキーで 開く★（お客さんの 道）
#   ・★ShowPopup() は 使わない★／★CommandBars の .Visible も 使わない★
#   ・★司さんの 画面（デスクトップ）が 起きている 時だけ 動きます★
#     （OpenInputDesktop が Default で ない時は 何も せずに 止まる）
param([string]$Dir='C:\Users\zeroa\AppData\Local\Temp\claude\C--WINDOWS-System32-WindowsPowerShell-v1-0\5b4e50e6-20a1-4af5-8ffb-8b6d6ca3f52b\scratchpad\shot')
$ErrorActionPreference='Stop'
Add-Type -AssemblyName UIAutomationClient; Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Drawing; Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System; using System.Runtime.InteropServices;
public class D {
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
  [DllImport("user32.dll")] public static extern IntPtr OpenInputDesktop(int f, bool i, uint a);
  [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern bool GetUserObjectInformation(IntPtr h,int i,System.Text.StringBuilder p,int n,out int len);
  [DllImport("user32.dll")] public static extern bool CloseDesktop(IntPtr h);
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
[void][D]::SetProcessDPIAware()
$AE=[System.Windows.Automation.AutomationElement]; $CT=[System.Windows.Automation.ControlType]

# ★画面が 起きているか 先に 見る★（寝ていたら 何も しない＝2026-09-03 に 何度も 踏んだ）
$hd=[D]::OpenInputDesktop(0,$false,0x0001)
$nb=New-Object System.Text.StringBuilder 256; $nl=0
$名='(開けない)'
if($hd -ne [IntPtr]::Zero){ [void][D]::GetUserObjectInformation($hd,2,$nb,256,[ref]$nl); [void][D]::CloseDesktop($hd); $名=$nb.ToString() }
"今 画面に 出ている デスクトップ = $名"
if($名 -ne 'Default'){ "★画面が 起きていません（$名）＝何も しません★"; exit 0 }

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
  [D]::ForceFront([IntPtr]$xl.Hwnd); Start-Sleep -Milliseconds 1500
  $sh=$wb.Worksheets.Item(1)
  $sh.Range("B2").Value2='あ'; $sh.Range("B3").Value2='い'; $sh.Range("C2").Value2='1'
  $c=$sh.Range("B3"); $c.Select()|Out-Null; Start-Sleep -Milliseconds 400
  $w=$xl.ActiveWindow
  $x=[int]$w.PointsToScreenPixelsX($c.Left+$c.Width/2); $y=[int]$w.PointsToScreenPixelsY($c.Top+$c.Height/2)

  # ★押す前に その 点に 居るのが Excel か 確かめる★（別の 窓を 押さない）
  $nm=''
  for($t=1;$t -le 5;$t++){
    $u=$AE::FromPoint((New-Object System.Windows.Point([double]$x,[double]$y)))
    $ww=$u; while($ww -ne $null -and $ww.Current.ControlType -ne $CT::Window){ $ww=[System.Windows.Automation.TreeWalker]::ControlViewWalker.GetParent($ww) }
    $nm = if($ww){ $ww.Current.ClassName } else { '不明' }
    if($nm -eq 'XLMAIN'){ break }
    "($t 回目) その 点は [$nm] ⇒ Excel を 前へ 出し直す"
    [D]::ForceFront([IntPtr]$xl.Hwnd); Start-Sleep -Milliseconds 1200
  }
  if($nm -ne 'XLMAIN'){ throw "★Excel を 前面に 出せませんでした（$nm）★" }

  # ★挿入(I)... と 削除(D)... の 窓★
  #   ★挿入(I) は リンク(I) と 同じ アクセスキー★（2026-09-03 実測）
  #   ⇒ 挿入は ★I を 1回 押して Enter★（1回目で 挿入に 当たる＝上に 在る）
  # ★アクセスキーが 1つだけの 行は ★押した 瞬間に 実行★★＝Enter を 押しては いけない
  #   （2026-09-03 実測＝削除(D) は 1つだけ／挿入(I) は リンク(I) と かぶるので 選ばれるだけ）
  $kumi=@(
    @{ na='削除'; key=0x44; f='excel_dialog_sakujo.png'; enter=$false },  # D＝1つだけ⇒即 実行
    @{ na='挿入'; key=0x49; f='excel_dialog_sonyu.png';  enter=$true  }   # I＝リンクと かぶる⇒Enter が 要る
  )
  foreach($k in $kumi){
    [D]::Key(0x1B); Start-Sleep -Milliseconds 250
    [D]::Key(0x1B); Start-Sleep -Milliseconds 250
    [D]::Key(0x1B); Start-Sleep -Milliseconds 400
    $c.Select() | Out-Null
    Start-Sleep -Milliseconds 400
    [D]::R($x,$y); Start-Sleep -Milliseconds 1500
    [D]::Key([byte]$k.key); Start-Sleep -Milliseconds 900
    if($k.enter){ [D]::Key(0x0D) }      # ★かぶる 時だけ Enter★（1つだけの 時は もう 実行されている）
    Start-Sleep -Milliseconds 2000
    $p=Shot $k.f $Dir
    "$($k.na) の 窓 … $p"
    [D]::Key(0x1B); Start-Sleep -Milliseconds 500   # ★窓を 閉じる（何も しない）★
    [D]::Key(0x1B); Start-Sleep -Milliseconds 300
  }
  [D]::Key(0x1B)
}catch{ "★落ちた★ $($_.Exception.Message)" }
finally{ if($xl){ try{foreach($z in @($xl.Workbooks)){$z.Close($false)}}catch{}; try{$xl.Quit()}catch{} } }
