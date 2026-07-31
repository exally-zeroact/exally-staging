# golden-excel.ps1 - 実Excel(COM)で「真値」を作る。
#   使い方: pwsh -File tools/golden-excel.ps1            ... 既存goldenと比較して pending + DIFF を出す(上書きしない)
#           pwsh -File tools/golden-excel.ps1 -Init      ... goldenが1つも無い時だけ、初回作成する
#
# ★重要な決まり(RECIPE.md と対)
#   ・式は .Formula (US-English構文) だけを使う。.FormulaLocal は使わない(環境で式が変わるため)。
#   ・文字列セルは NumberFormat='@' を先に当ててから書く。でないと '0007'/'1,234'/'2026-07-31' が
#     Excelに勝手に数値・日付へ化けて、比べている物が変わってしまう(実際に踏んだ)。
#   ・既存の golden は絶対に黙って上書きしない。差分を出して人が確認してから差し替える。
#   ・pwsh(PowerShell 7)で実行する。Windows PowerShell 5.1 でも動くが、その場合このファイルはBOM付きで保存が必要。
param([switch]$Init)

$ErrorActionPreference = 'Stop'
$ROOT      = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$CASEDIR   = Join-Path $ROOT 'tests\xlsx-harness\cases'
$GOLDENDIR = Join-Path $ROOT 'tests\xlsx-harness\golden'

function ReadJson($path) {
  [System.IO.File]::ReadAllText($path, [Text.Encoding]::UTF8) | ConvertFrom-Json
}

# ---- ケース読み込み ------------------------------------------------------
$inputs = ReadJson (Join-Path $CASEDIR '_inputs.json')
$caseFiles = Get-ChildItem $CASEDIR -Filter '*.json' | Where-Object { $_.Name -ne '_inputs.json' } | Sort-Object Name
$allCases = @()
foreach ($f in $caseFiles) {
  $j = ReadJson $f.FullName
  foreach ($c in $j.cases) { $allCases += ,@($c.id, $c.f, [bool]$c.volatile, [bool]$c.spill) }
}
Write-Host ("cases: {0} (files: {1})" -f $allCases.Count, $caseFiles.Count)

# ---- エラー値の数値コード -> 表示 --------------------------------------
#   出典: Microsoft Learn "XlCVError enumeration (Excel)"
#         https://learn.microsoft.com/en-us/office/vba/api/excel.xlcverror
#   規則: COM Value2 = -2146828288 + XlCVError値  (8件中7件を実測で確認済)
#   ★キーは文字列で持つ(数値キーだと Int32/Int64 の型違いで ContainsKey が当たらない。実際に踏んだ)
$ERRMAP = @{
  '-2146826288' = '#NULL!'   # xlErrNull  2000 (公式+実測)
  '-2146826281' = '#DIV/0!'  # xlErrDiv0  2007 (公式+実測)
  '-2146826273' = '#VALUE!'  # xlErrValue 2015 (公式+実測)
  '-2146826265' = '#REF!'    # xlErrRef   2023 (公式+実測)
  '-2146826259' = '#NAME?'   # xlErrName  2029 (公式+実測)
  '-2146826252' = '#NUM!'    # xlErrNum   2036 (公式+実測)
  '-2146826246' = '#N/A'     # xlErrNA    2042 (公式+実測)
  '-2146826243' = '#SPILL!'  # xlErrSpill 2045 (公式のみ・★未実測)
  '-2146826238' = '#CALC!'   # 2050 相当   (★実測のみ・公式ページに記載なし)
}

# ---- Excel 起動 ---------------------------------------------------------
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)

function ColRow($addr) {
  $m = [regex]::Match($addr, '^([A-Z]+)(\d+)$')
  $col = 0
  foreach ($ch in $m.Groups[1].Value.ToCharArray()) { $col = $col * 26 + ([int][char]$ch - 64) }
  return @([int]$m.Groups[2].Value, $col)
}

# ---- 入力セル(★型を守って書く) ----------------------------------------
foreach ($p in $inputs.cells.PSObject.Properties) {
  $rc = ColRow $p.Name
  $cell = $ws.Cells.Item($rc[0], $rc[1])
  $spec = $p.Value
  switch ([string]$spec.t) {
    's' { $cell.NumberFormat = [string]'@'; $cell.Value2 = [string]$spec.v }   # ★先に文字列書式
    'n' { $cell.Value2 = [double]$spec.v }
    'b' { $cell.Value2 = [bool]$spec.v }
    default { throw ("unknown input type: " + $spec.t + " at " + $p.Name) }
  }
}

# ---- 式を1本ずつ入れて値を読む -----------------------------------------
$cases = [ordered]@{}
$volatileSkipped = @()
$row = 1
$spillCol = 30                        # ★スピルする式は1本ずつ別の列(AD以降)に置く。
                                      #   T列に並べると隣の行へ溢れて次の式を潰し、#SPILL! だらけになる。
foreach ($c in $allCases) {
  $id = $c[0]; $formula = $c[1]; $isVolatile = $c[2]; $isSpill = $c[3]
  if ($isVolatile) { $volatileSkipped += $id; continue }   # ★揮発性はgoldenに入れない
  if ($isSpill) { $cell = $ws.Cells.Item(1, $spillCol); $spillCol++ }
  else          { $cell = $ws.Cells.Item($row, 20); $row++ }   # T列
  $rec = [ordered]@{ f = $formula }
  try {
    # ★.Formula2 で入れる(動的配列の解釈)。.Formula は旧来の解釈で、
    #   配列を返す式に「暗黙の交差」が働き、置いた行によって答えが変わってしまう(実測):
    #     .Formula  : =SUM(C1:C6*E1:E6) が row1で100 / row2で0 …(行依存のデタラメ)
    #     .Formula2 : 97100(Excelの画面で人が入力した時と同じ)
    #   配列でない式は両者で同じ。古いExcelには Formula2 が無いので、その時は Formula に落とす。
    try   { $cell.Formula2 = [string]$formula }
    catch { $cell.Formula  = [string]$formula }
    $v = $cell.Value2
    if ($null -eq $v) {
      $rec['t'] = 'z'; $rec['v'] = $null
    } elseif ($v -is [string]) {
      $rec['t'] = 's'; $rec['v'] = [string]$v
    } elseif ($v -is [bool]) {
      $rec['t'] = 'b'; $rec['v'] = [bool]$v
    } else {
      $n = [double]$v
      $key = [string]([int64]$n)
      if ($ERRMAP.ContainsKey($key)) { $rec['t'] = 'e'; $rec['v'] = $ERRMAP[$key]; $rec['code'] = [int64]$n }
      else                           { $rec['t'] = 'n'; $rec['v'] = $n }
    }
    $rec['text'] = [string]$cell.Text
    if ($isSpill) {
      # Excelが何セルに溢れたか(スピル範囲)。うちのグリッドは再現できないので記録だけする。
      try { $rec['spillCells'] = [int]$cell.SpillingToRange.Count } catch { $rec['spillCells'] = $null }
    }
  } catch {
    $rec['t'] = 'throw'; $rec['v'] = $_.Exception.Message; $rec['text'] = ''
  }
  $cases[$id] = $rec
}

# ---- 入力プローブ(★標準書式のセルに打ち込んだ時、Excelがどう解釈するか) --
#   .Formula に文字列を渡す = 人がキーで打ち込むのと同じ扱いになる。
#   ここは「文字列のまま残るか」ではなく「Excelがどうしたか」を真値として記録する。
$probes = [ordered]@{}
$prow = 1
foreach ($p in $inputs.inputProbes.probes) {
  $cell = $ws.Cells.Item($prow, 24)   # X列(標準書式のまま)
  $prow++
  $cell.Formula = [string]$p.raw
  $v = $cell.Value2
  $rec = [ordered]@{ raw = [string]$p.raw; label = [string]$p.label }
  if ($null -eq $v)            { $rec['t'] = 'z'; $rec['v'] = $null }
  elseif ($v -is [string])     { $rec['t'] = 's'; $rec['v'] = [string]$v }
  elseif ($v -is [bool])       { $rec['t'] = 'b'; $rec['v'] = [bool]$v }
  else                         { $rec['t'] = 'n'; $rec['v'] = [double]$v }
  $rec['text'] = [string]$cell.Text
  $probes[[string]$p.id] = $rec
}

# ---- 環境の記録(RECIPEに載せる物と同じ) --------------------------------
$c2r = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
$meta = [ordered]@{
  source          = 'real-excel'
  product         = if ($c2r) { [string]$c2r.ProductReleaseIds } else { 'unknown' }
  version         = if ($c2r) { [string]$c2r.VersionToReport } else { [string]$xl.Version }
  platform        = if ($c2r) { [string]$c2r.Platform } else { 'unknown' }
  updateChannel   = if ($c2r) { [string]$c2r.UpdateChannel } else { 'unknown' }
  uiLanguageId    = [int]$xl.LanguageSettings.LanguageID(2)
  installLangId   = [int]$xl.LanguageSettings.LanguageID(1)
  countrySetting  = [int]$xl.International(2)
  decimalSep      = [string]$xl.International(3)
  thousandsSep    = [string]$xl.International(4)
  listSep         = [string]$xl.International(5)
  date1904        = [bool]$wb.Date1904
  formulaProperty = '.Formula (US-English) only; .FormulaLocal is NOT used'
  caseCount       = $cases.Count
  volatileSkipped = $volatileSkipped
}

$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

# ---- 出力(★既存を上書きしない) ----------------------------------------
if (-not (Test-Path $GOLDENDIR)) { New-Item -ItemType Directory -Path $GOLDENDIR | Out-Null }
$tag = ('excel-365-' + ($meta.version -replace '^(\d+\.\d+\.\d+).*$', '$1'))
$goldenPath = Join-Path $GOLDENDIR ($tag + '.json')
$payload = [ordered]@{ meta = $meta; cases = $cases; inputProbes = $probes }
$utf8 = New-Object Text.UTF8Encoding($false)
$jsonOut = ($payload | ConvertTo-Json -Depth 8)

if ((-not (Test-Path $goldenPath)) -and $Init) {
  [System.IO.File]::WriteAllText($goldenPath, $jsonOut, $utf8)
  Write-Host ("[Init] golden 新規作成: {0}  ({1} cases)" -f $goldenPath, $cases.Count)
  exit 0
}
if (-not (Test-Path $goldenPath)) {
  Write-Host ("golden がありません。初回は -Init を付けて実行してください: {0}" -f $goldenPath)
  exit 1
}

# 既存があるとき = pending に出して差分を報告する(勝手に差し替えない)
$stamp = (Get-Date -Format 'yyyy-MM-dd')
$pendingPath = Join-Path $GOLDENDIR ("pending-$stamp.json")
[System.IO.File]::WriteAllText($pendingPath, $jsonOut, $utf8)

$old = ReadJson $goldenPath
$diffLines = @()
foreach ($k in $cases.Keys) {
  $oldRec = $old.cases.PSObject.Properties[$k]
  if (-not $oldRec) { $diffLines += "| $k | (なし) | $($cases[$k].v) | ★新規ケース |"; continue }
  $ov = [string]$oldRec.Value.v; $nv = [string]$cases[$k].v
  $ot = [string]$oldRec.Value.t; $nt = [string]$cases[$k].t
  if ($ov -ne $nv -or $ot -ne $nt) { $diffLines += "| $k | $ov ($ot) | $nv ($nt) | ★値が変わった |" }
}
foreach ($p in $old.cases.PSObject.Properties) {
  if (-not $cases.Contains($p.Name)) { $diffLines += "| $($p.Name) | $($p.Value.v) | (なし) | ケースが消えた |" }
}

$diffPath = Join-Path $GOLDENDIR ("DIFF-$stamp.md")
$head = @(
  "# golden 差分 $stamp",
  "",
  "- 既存: $goldenPath  (version $($old.meta.version))",
  "- 新規: $pendingPath (version $($meta.version))",
  ""
)
if ($diffLines.Count -eq 0) {
  $head += "**差分なし。Excel側に変化なし。** pending は破棄してよい。"
  Write-Host "差分なし(Excel側に変化なし)。pending: $pendingPath"
} else {
  $head += @("| ケース | 旧(既存golden) | 新(今回のExcel) | |", "|---|---|---|---|") + $diffLines
  $head += @("", "★値が変わった = Excel側が変わったという重大な情報。報告してから差し替えること。")
  Write-Host ("★差分 {0} 件。DIFF: {1}" -f $diffLines.Count, $diffPath)
}
[System.IO.File]::WriteAllText($diffPath, ($head -join "`n"), $utf8)
