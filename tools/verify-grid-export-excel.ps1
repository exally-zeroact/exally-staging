# verify-grid-export-excel.ps1 — グリッドの「Excelに保存」で出るファイルを【実Excelで開いて】確かめる。
#   使い方: pwsh -File tools/verify-grid-export-excel.ps1
#
# なぜ要るか:
#   xlsx の書き出しは「ファイルにする所まで緑」でも、実Excelで開くと違う事が何度も起きた。
#     ・_xlfn. の付け忘れ  → その式だけ #NAME?（RANK.AVG で実際に踏んだ）
#     ・表示形式を運び忘れ → 画面で「54,000」が ファイルでは「54000」（2026-08-02 に実際に踏んだ）
#   CIには Excel が無いので、ここは人が走らせる道具。★リリース前に必ず1回。
#
# やること: グリッドと同じ形のデータ → lib/grid-xlsx.js → lib/xlsx-io.js で書き出し
#           → 実Excelで開いて再計算 → 値・式・表示形式・列幅・セル結合を突き合わせ

$ErrorActionPreference = 'Stop'
$ROOT = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$xlsx = Join-Path $env:TEMP 'exally-grid-export-check.xlsx'

# ── グリッドと同じ形のデータから xlsx を作る ──
$js = @'
const path=require('path');
const ROOT=process.argv[2];
const G=require(path.join(ROOT,'lib','grid-xlsx.js'));
const IO=require(path.join(ROOT,'lib','xlsx-io.js'));
const sheets=[{
  name:'請求',
  data:{
    '0,0':{v:'品名',f:'品名',d:'品名'}, '0,3':{v:'金額',f:'金額',d:'金額'},
    '1,0':{v:'エアコン清掃',f:'エアコン清掃',d:'エアコン清掃'},
    '1,1':{v:'2',f:'2',d:'2'}, '1,2':{v:'15000',f:'15000',d:'15000'},
    '1,3':{f:'=B2*C2',v:'=B2*C2',d:'30000',numFmt:'#,##0'},
    '2,0':{v:'ダクト清掃',f:'ダクト清掃',d:'ダクト清掃'},
    '2,1':{v:'3',f:'3',d:'3'}, '2,2':{v:'8000',f:'8000',d:'8000'},
    '2,3':{f:'=B3*C3',v:'=B3*C3',d:'24000',numFmt:'#,##0'},
    '3,0':{v:'小計',f:'小計',d:'小計',bold:true},
    '3,3':{f:'=SUM(D2:D3)',v:'=SUM(D2:D3)',d:'54000',numFmt:'#,##0'},
    '4,0':{v:'消費税',f:'消費税',d:'消費税'},
    '4,3':{f:'=ROUNDDOWN(D4*0.1,0)',v:'=ROUNDDOWN(D4*0.1,0)',d:'5400',numFmt:'#,##0'},
    '5,0':{v:'合計',f:'合計',d:'合計',bold:true},
    '5,3':{f:'=D4+D5',v:'=D4+D5',d:'59400',numFmt:'#,##0'},
    '6,0':{v:'0007',f:'0007',d:'0007'},
    '6,1':{v:'1,234',f:'1,234',d:'1,234'},
    '7,0':{v:'摘要',f:'摘要',d:'摘要',merged:{r:7,c:0},mergeEnd:{r:7,c:3}},
    '8,0':{f:'=XLOOKUP("ダクト清掃",A2:A3,D2:D3,0)',v:'',d:'24000'}
  },
  colW:{0:140}
}];
require('fs').writeFileSync(process.argv[3], IO.writeBook(G.gridToBook(sheets)));
'@
$tmpjs = Join-Path $env:TEMP 'exally-grid-export-build.js'
Set-Content -Path $tmpjs -Value $js -Encoding UTF8
node $tmpjs $ROOT $xlsx
Write-Host ("書き出した: {0} ({1} バイト)" -f $xlsx, (Get-Item $xlsx).Length)

# ── 実Excelで開いて突き合わせ ──
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$ng = 0
try {
  $wb = $x.Workbooks.Open($xlsx)
  $ws = $wb.Worksheets.Item(1)
  function Check($what, $got, $want) {
    if ("$got" -eq "$want") { Write-Host ("  OK   {0}: {1}" -f $what, $got) }
    else { Write-Host ("  ★NG  {0}: 期待={1} 実際={2}" -f $what, $want, $got); $script:ng++ }
  }
  Check 'シート名'            $ws.Name                      '請求'
  Check 'D2 の式'             $ws.Range('D2').Formula       '=B2*C2'
  Check 'D4 の式'             $ws.Range('D4').Formula       '=SUM(D2:D3)'
  Check 'D5 の式'             $ws.Range('D5').Formula       '=ROUNDDOWN(D4*0.1,0)'
  Check 'D4 の再計算値'       $ws.Range('D4').Value2        54000
  Check 'D6 の再計算値(合計)' $ws.Range('D6').Value2        59400
  Check 'D4 の表示形式'       $ws.Range('D4').NumberFormat  '#,##0'
  Check 'D4 の見え方'         $ws.Range('D4').Text          '54,000'
  Check 'A7(0007)'            $ws.Range('A7').Value2        7
  Check 'B7(1,234)は文字列'   $ws.Range('B7').Value2        '1,234'
  Check 'A8 の結合'           $ws.Range('A8').MergeArea.Address($false,$false) 'A8:D8'
  Check 'A列の幅(140px相当)'  ([math]::Round($ws.Columns.Item(1).ColumnWidth,1)) 22.8
  # ★新しい関数が #NAME? になっていないか（_xlfn. の付け忘れ検知）
  Check 'A9 XLOOKUP が動く'   $ws.Range('A9').Value2        24000
  $wb.Close($false)
} finally {
  $x.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($x) | Out-Null
}
Write-Host ""
if ($ng -gt 0) { Write-Host ("★実Excelで開いた結果 {0} 件ズレています" -f $ng); exit 1 }
Write-Host "実Excelで開いて全部一致（式・再計算値・表示形式・列幅・セル結合・新関数）"
