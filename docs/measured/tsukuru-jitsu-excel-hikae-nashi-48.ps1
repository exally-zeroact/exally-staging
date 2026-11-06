# tsukuru-jitsu-excel-hikae-nashi-48.ps1  (2026-11-07 Exally-値 席)
#
#  ★なぜ★ kansuu-519-atai-souname.mjs で「実Excel の控えが 1本も 無い」関数が 56個。
#    そのうち 外へ出る 6個（WEBSERVICE STOCKHISTORY TRANSLATE DETECTLANGUAGE IMAGE RTD）と
#    DLL を呼ぶ 2個（CALL REGISTER.ID）を除いた 48個に、手で決めた形を打って 値を取る。
#    あわせて ★空のマスのずれ★（COUNTIF(A1:A7,"") 等）を 同じ盤面で取る。
#
#  ★盤面（kansuu-519 の 盤面C ＋ 字の C1:C3）★
#    A1:A5=1..5 ／ B1:B5=2,4,6,8,10 ／ D1==DATE(2024,1,1) ／ D2==DATE(2026,1,1)
#    C1="apple-123" ／ C2="Banana" ／ C3="あいう" ／ A6:A7・B6:B7・Z1 は空
#    式は BZ1 に Formula2（今の道）で打つ＝1本ずつ打って 読んで 消す。
#
#  ★既存の Excel には 触らない★
#    New-Object で立てた Excel の Hwnd から PID を引き、★走らせる前に居なかった PID★ で
#    なければ Quit せずに抜ける（exit 4）。終わりは 自分の PID だけを Quit→残れば その PID だけ止める。
#
#  出す物: 引数1 の tsv（名 / 式 / Value2 / Text / 溢れ全部）
#  使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <これ> <出す先.tsv>

param([string]$Out)
if (-not $Out) { Write-Host 'NG 出す先が無い'; exit 6 }
if ($PSVersionTable.PSVersion.Major -ne 5) { Write-Host 'NG powershell.exe(5.1) で走らせる'; exit 8 }

$forms = @(
  @('ARRAYTOTEXT', '=ARRAYTOTEXT(A1:B2)', '=ARRAYTOTEXT(A1:B2,1)', '=ARRAYTOTEXT(C1:C3,1)'),
  @('AVERAGEIFS', '=AVERAGEIFS(B1:B5,A1:A5,">2")', '=AVERAGEIFS(B1:B5,A1:A5,">1",A1:A5,"<5")', '=AVERAGEIFS(B1:B7,A1:A7,"<>3")'),
  @('BAHTTEXT', '=BAHTTEXT(1234.5)', '=BAHTTEXT(0)'),
  @('BINOM.DIST.RANGE', '=BINOM.DIST.RANGE(10,0.5,3)', '=BINOM.DIST.RANGE(10,0.5,3,6)'),
  @('CELL', '=CELL("row",B3)', '=CELL("col",B3)', '=CELL("address",B3)', '=CELL("contents",A2)', '=CELL("type",C1)', '=CELL("type",Z9)'),
  @('CHOOSECOLS', '=CHOOSECOLS(A1:B5,2)', '=CHOOSECOLS(A1:B5,-1)'),
  @('CHOOSEROWS', '=CHOOSEROWS(A1:B5,2)', '=CHOOSEROWS(A1:B5,-1)'),
  @('CONCATENATE', '=CONCATENATE("a",A1,"b")', '=CONCATENATE(C1,C2)'),
  @('DROP', '=DROP(A1:B5,2)', '=DROP(A1:B5,-1,1)'),
  @('ERF.PRECISE', '=ERF.PRECISE(0.5)', '=ERF.PRECISE(1)'),
  @('ERFC.PRECISE', '=ERFC.PRECISE(0.5)', '=ERFC.PRECISE(1)'),
  @('ERROR.TYPE', '=ERROR.TYPE(1/0)', '=ERROR.TYPE(NA())', '=ERROR.TYPE(1)'),
  @('EXPAND', '=EXPAND(A1:A2,3,2,0)', '=EXPAND(A1:A2,3)'),
  @('FINDB', '=FINDB("い",C3)', '=FINDB("n",C2)'),
  @('FORECAST.ETS', '=FORECAST.ETS(6,B1:B5,A1:A5)'),
  @('FORECAST.ETS.CONFINT', '=FORECAST.ETS.CONFINT(6,B1:B5,A1:A5)'),
  @('FORECAST.ETS.SEASONALITY', '=FORECAST.ETS.SEASONALITY(B1:B5,A1:A5)'),
  @('FORECAST.ETS.STAT', '=FORECAST.ETS.STAT(B1:B5,A1:A5,1)'),
  @('GROWTH', '=GROWTH(B1:B5,A1:A5,6)', '=GROWTH(B1:B5)'),
  @('INFO', '=INFO("recalc")'),
  @('LOGEST', '=LOGEST(B1:B5,A1:A5)'),
  @('MINVERSE', '=MINVERSE({1,2;3,4})', '=MINVERSE(A1:B2)'),
  @('MODE.MULT', '=MODE.MULT({1,2,2,3,3})', '=MODE.MULT(A1:A5)'),
  @('MUNIT', '=MUNIT(2)', '=MUNIT(1)'),
  @('N', '=N(A1)', '=N(C1)', '=N(TRUE)', '=N(D1)'),
  @('PERCENTOF', '=PERCENTOF(A1,A1:A5)', '=PERCENTOF(B1:B5,B1:B5)'),
  @('PERCENTRANK.EXC', '=PERCENTRANK.EXC(A1:A5,3)', '=PERCENTRANK.EXC(A1:A5,2.5,2)'),
  @('PERCENTRANK.INC', '=PERCENTRANK.INC(A1:A5,3)', '=PERCENTRANK.INC(A1:A5,2.5,2)'),
  @('PHONETIC', '=PHONETIC(C3)', '=PHONETIC(C1)'),
  @('PIVOTBY', '=PIVOTBY(A1:A5,B1:B5,B1:B5,SUM)'),
  @('PROB', '=PROB(A1:A5,{0.1,0.2,0.3,0.2,0.2},2)', '=PROB(A1:A5,{0.1,0.2,0.3,0.2,0.2},2,4)'),
  @('REGEXEXTRACT', '=REGEXEXTRACT(C1,"[0-9]+")'),
  @('REGEXREPLACE', '=REGEXREPLACE(C1,"[0-9]","#")'),
  @('REGEXTEST', '=REGEXTEST(C2,"^B")', '=REGEXTEST(C2,"^b",1)'),
  @('REPLACEB', '=REPLACEB(C3,1,2,"X")', '=REPLACEB("abcdef",2,3,"Z")'),
  @('SEARCHB', '=SEARCHB("う",C3)', '=SEARCHB("an",C2)'),
  @('SORT', '=SORT(B1:B5,1,-1)', '=SORT(A1:B5,2,-1)'),
  @('SORTBY', '=SORTBY(A1:A5,B1:B5,-1)'),
  @('SUMPRODUCT', '=SUMPRODUCT(A1:A5,B1:B5)', '=SUMPRODUCT(A1:A5*B1:B5)', '=SUMPRODUCT((A1:A5>2)*B1:B5)'),
  @('T', '=T(C1)', '=T(A1)'),
  @('TAKE', '=TAKE(A1:B5,2)', '=TAKE(A1:B5,-1,-1)'),
  @('TEXTSPLIT', '=TEXTSPLIT("a,b,c",",")', '=TEXTSPLIT(C1,"-")'),
  @('TOCOL', '=TOCOL(A1:B2)'),
  @('TOROW', '=TOROW(A1:B2)'),
  @('TREND', '=TREND(B1:B5,A1:A5,6)', '=TREND(B1:B5)'),
  @('UNIQUE', '=UNIQUE({1,2,2,3})', '=UNIQUE(A1:A5)'),
  @('WRAPCOLS', '=WRAPCOLS(A1:A5,2)', '=WRAPCOLS(A1:A5,2,0)'),
  @('WRAPROWS', '=WRAPROWS(A1:A5,2)', '=WRAPROWS(A1:A5,2,0)'),
  # ★空のマスのずれ★（名は 空 と付ける＝関数の数には入れない）
  @('空', '=COUNTIF(A1:A7,"")', '=COUNTIF(A1:A7,"<>3")', '=COUNTIF(A1:A7,"<>")', '=COUNTIF(A1:A7,"=")',
        '=COUNTIFS(A1:A7,"")', '=SUMIF(A1:A7,"",B1:B7)', '=SUMIF(A1:A7,"<>3",B1:B7)', '=AVERAGEIF(A1:A7,"<>3")',
        '=COUNTBLANK(A1:A7)', '=XLOOKUP(Z1,A1:A7,B1:B7)', '=XLOOKUP(Z1,A1:A7,B1:B7,"無")', '=MATCH(Z1,A1:A7,0)',
        '=VLOOKUP(Z1,A1:B7,2,FALSE)', '=XMATCH(Z1,A1:A7)', '=COUNTIF(A1:A7,Z1)')
)

Add-Type -Namespace W -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
$before = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" | ForEach-Object { [int]$_.ProcessId })
Write-Host ('走らせる前の Excel PID: ' + ($before -join ','))

$xl = New-Object -ComObject Excel.Application
$xpid = [uint32]0
[void][W.U]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$xpid)
Write-Host ('立てた Excel の PID: ' + $xpid)
if ($xpid -eq 0 -or $before -contains [int]$xpid) {
  Write-Host 'NG 既存の Excel に繋がった＝Quit せずに抜ける'
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl); exit 4
}

$errs = @{ -2146826281 = '#DIV/0!'; -2146826252 = '#NUM!'; -2146826246 = '#N/A'; -2146826273 = '#VALUE!';
  -2146826265 = '#REF!'; -2146826259 = '#NAME?'; -2146826288 = '#NULL!'; -2146826238 = '#CALC!'; -2146826245 = '#SPILL!' }
$inv = [System.Globalization.CultureInfo]::InvariantCulture
function V($v) {
  if ($null -eq $v) { return '' }
  if ($v -is [int]) { if ($errs.ContainsKey($v)) { return $errs[$v] } ; return $v.ToString($inv) }
  if ($v -is [double]) { return $v.ToString('R', $inv) }
  if ($v -is [bool]) { if ($v) { return 'TRUE' } else { return 'FALSE' } }
  return ([string]$v) -replace "[`t`r`n]", ' '
}

$rows = New-Object System.Collections.Generic.List[string]
$rows.Add("# 実Excel の値（$((Get-Date).ToString('yyyy-MM-dd HH:mm'))・Excel $($xl.Version) build $($xl.Build)・BZ1 に Formula2）")
$rows.Add("名`t式`tValue2`tText`t溢れ")
$ok = 0
try {
  $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
  $bk = $xl.Workbooks.Add(); $sh = $bk.Sheets.Item(1)
  for ($r = 1; $r -le 5; $r++) { $sh.Cells.Item($r, 1).Value2 = $r; $sh.Cells.Item($r, 2).Value2 = 2 * $r }
  $sh.Range('D1').Formula = '=DATE(2024,1,1)'; $sh.Range('D2').Formula = '=DATE(2026,1,1)'
  $sh.Range('C1').Value2 = 'apple-123'; $sh.Range('C2').Value2 = 'Banana'; $sh.Range('C3').Value2 = 'あいう'
  $c = $sh.Range('BZ1')
  foreach ($f in $forms) {
    $name = $f[0]
    for ($i = 1; $i -lt $f.Count; $i++) {
      $fx = $f[$i]
      try {
        $c.Formula2 = $fx
        $xl.Calculate()
        $v = V $c.Value2
        $t = ([string]$c.Text) -replace "[`t`r`n]", ' '
        $sp = ''
        try { if ($c.HasSpill) { $a = $c.SpillingToRange.Value2; $parts = @(); foreach ($x in $a) { $parts += (V $x) }; $sp = ($parts -join '|') } } catch {}
        $rows.Add("$name`t$fx`t$v`t$t`t$sp"); $ok++
      } catch {
        $rows.Add("$name`t$fx`t打てません`t`t")
      }
      $sh.Range('BZ1:CZ30').Clear()
    }
  }
  $bk.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Start-Sleep -Seconds 3
  $still = Get-Process -Id $xpid -ErrorAction SilentlyContinue
  if ($still) { Write-Host ('自分の Excel(PID ' + $xpid + ') が残った＝それだけ止める'); Stop-Process -Id $xpid -Force }
}
[System.IO.File]::WriteAllLines($Out, $rows, (New-Object System.Text.UTF8Encoding($false)))
Write-Host ('打てた ' + $ok + '本 → ' + $Out)
if ($ok -eq 0) { exit 5 }
$after = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" | ForEach-Object { [int]$_.ProcessId })
Write-Host ('走らせた後の Excel PID: ' + ($after -join ','))
