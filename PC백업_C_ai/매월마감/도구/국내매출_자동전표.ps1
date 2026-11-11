# 국내매출 자동전표(아마란스10 '자동전표처리(일반)' 업로드 엑셀) 만들기 — 매달
# 사용: powershell -ExecutionPolicy Bypass -File 국내매출_자동전표.ps1 -YM 202610
# 근거(2026-11-11, 9월 실제 국내매출 전표 35장 분석 + 사용자 지정):
#  - 기본 틀 = 국내영업 'YYYY년M월 라인별 매출확정.xlsx'의 'M월' 시트(거래처별 사형·자동·금형·가공3 금액)
#  - 세금계산서 1장 = 전표 1장. 세금계산서 번호·금액·날짜 = 'YYYY년M월 매출마감 확정.xlsx'의 'YYMM' 시트(BL 칸 = 세금계산서 날짜 = 신고기준일)
#    라인별 줄과 세금계산서는 거래처 안에서 금액으로 짝지음(라인별의 '(10일)' 같은 표시는 실제 날짜와 다를 때가 있음)
#  - 줄: 국내매출(4130000, 대변, 라인별 프로젝트 사형101·자동102·금형201·가공3 403) → 부가세예수금(2550000, 대변, 세액) → 외상매출금(1080000, 차변, 증빙 1 세금계산서)
#  - 부가세 줄: 부가세사업장(공급자 사업장: 대신금속 1000 / 마산3 4000 …), 신고기준일, 세무구분 11 과세매출, 공급가액, 전자세금계산서 Y, 국세청 승인번호
#  - 신코·코마츠는 해외매출 자동전표에서 처리하므로 뺌(사용자 지정 2026-11-11)
#  - 거래처코드·적요·품의내역·부가세사업장 = 도구\국내전표_거래처.csv (9월 실제 전표에서 뽑음, 사람이 엑셀로 고침)
#  - 세액 = 아마란스 세금계산서 목록(acb2050…xlsx)에 있으면 그 값, 없으면 공급가액×10% 원 미만 버림(확인 표시)
#  - 작성번호 = 날짜마다 2000부터(해외 자동전표 1000번대·손 전표와 안 겹치게)
param([Parameter(Mandatory = $true)][string]$YM, [string]$LineFile, [string]$LineSheet, [string]$CloseFile, [string]$CloseSheet, [string]$TaxFile, [string]$OutDir, [string]$CompareCsv, [int]$StartNo = 2000, [switch]$NoIssNo)
$ErrorActionPreference = 'Stop'
$root = 'C:\ai\매월마감'; $toolDir = Split-Path $MyInvocation.MyCommand.Path
$Y = [int]$YM.Substring(0, 4); $M = [int]$YM.Substring(4, 2); $cur = [DateTime]::new($Y, $M, 1); $ym4 = $cur.ToString('yyMM')
$baseDir = Join-Path $root '국내자동전표'; $monDir = Join-Path $baseDir $YM
function FindFile([string[]]$pats, [string[]]$dirs) { foreach ($d in $dirs) { if (-not (Test-Path -LiteralPath $d)) { continue }; foreach ($p in $pats) { $f = Get-ChildItem -LiteralPath $d -File | Where-Object { $_.Name -like $p -and $_.Name -notlike '~$*' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1; if ($f) { return $f.FullName } } }; $null }
if (-not $LineFile) { $LineFile = FindFile @("*$($Y)년$($M)월*라인별*매출확정*.xlsx", "*$($Y)년 $($M)월*라인별*매출확정*.xlsx", '*라인별*매출확정*.xlsx') @($monDir, $baseDir) }
if (-not $CloseFile) { $CloseFile = FindFile @("*$($Y)년$($M)월*매출마감*확정*.xlsx", "*$($Y)년 $($M)월*매출마감*확정*.xlsx", '*매출마감*확정*.xlsx') @($monDir, $baseDir) }
# 세금계산서 목록 후보: ① 홈택스 '매출전자세금계산서목록' (가장 정확, 수정세금계산서까지 다 있음) → ② 아마란스 세금계산서 목록 acb2050…
#  이번 달 작성분이 들어 있는 첫 파일을 씀 (다른 달 파일이 같은 폴더에 있어도 괜찮게)
$taxCands = @(); $sumCands = @()
if ($TaxFile) { $taxCands = @($TaxFile) } else {
  foreach ($pat in '*매출전자세금계산서목록*.xls*', 'acb2050*.xlsx') { foreach ($d in @($monDir, $baseDir, "C:\ai\기초자료\$YM")) { if (Test-Path -LiteralPath $d) { $taxCands += @(Get-ChildItem -LiteralPath $d -File | Where-Object { $_.Name -like $pat -and $_.Name -notlike '~$*' } | Sort-Object LastWriteTime -Descending | ForEach-Object { $_.FullName }) } } } }
foreach ($d in @($monDir, $baseDir, "C:\ai\기초자료\$YM")) { if (Test-Path -LiteralPath $d) { $sumCands += @(Get-ChildItem -LiteralPath $d -File | Where-Object { $_.Name -like '*매출명세서목록*.xls*' -and $_.Name -notlike '~$*' } | Sort-Object LastWriteTime -Descending | ForEach-Object { $_.FullName }) } }
# 홈택스 공급자 종사업장번호 → 아마란스 부가세사업장 (2026-10 홈택스 목록: 종사업장 1 = 다이헨·맥시스·알멕 물품대·신우테크윈·아트라스 = 아마란스 '대신금속(주)마산3', 빈칸 = 본점)
$JONG = @{ '' = '1000'; '1' = '4000' }
if (-not $LineSheet) { $LineSheet = "$($M)월" }; if (-not $CloseSheet) { $CloseSheet = $ym4 }
if (-not $OutDir) { $OutDir = $monDir }
$template = Join-Path $baseDir '자동전표처리(일반) 양식.xlsx'; if (-not (Test-Path -LiteralPath $template)) { $template = Join-Path $root '해외자동전표\자동전표처리(일반) 양식.xlsx' }
New-Item -ItemType Directory -Force $OutDir | Out-Null
$workDir = Join-Path $OutDir '작업'; New-Item -ItemType Directory -Force $workDir | Out-Null
$logF = Join-Path $workDir '전표_실행기록.txt'; $doneF = Join-Path $workDir '전표_결과.txt'; $U8 = New-Object Text.UTF8Encoding($true)
[IO.File]::WriteAllText($logF, "$($Y)년 $($M)월 국내매출 자동전표 ($(Get-Date -f 'yyyy-MM-dd HH:mm:ss'))`r`n", $U8); if (Test-Path $doneF) { Remove-Item $doneF }
function Say([string]$m, [string]$c = 'Gray') { Write-Host $m -ForegroundColor $c; [IO.File]::AppendAllText($logF, "$m`r`n", $U8) }
function Done([string]$st, [string]$msg) { [IO.File]::WriteAllText($doneF, "$st`r`n$msg", $U8) }
$CHK = New-Object Collections.Generic.List[object]
function Chk($lvl, $who, $text) { $CHK.Add([pscustomobject]@{ 구분 = $lvl; 거래처 = $who; 내용 = $text; 순 = $CHK.Count }) | Out-Null; Say "  [$lvl] $who : $text" $(if ($lvl -eq '에러') { 'Red' } elseif ($lvl -eq '참고') { 'Gray' } else { 'Yellow' }) }
function NumV($v) { if ($null -eq $v -or "$v".Trim() -eq '') { return $null }; if ($v -is [double]) { return $v }; $d = 0.0; if ([double]::TryParse(("$v" -replace '[,₩\s]', ''), [ref]$d)) { $d } else { $null } }
function NormHead($s) { ("$s" -replace '\s', '') }
# 거래처 이름 맞추기: 날짜 표시 '(10일)·(말일)', ㈜·(주)·주식회사, 띄어쓰기·'/' 를 뺌 → '현대로템㈜철차(31일)' = '현대로템㈜ / 철차' = '현대로템철차'
function NormKey($s) { ("$s" -replace '\((\d+일|말일)\)', '' -replace '㈜|\(주\)|주식회사', '' -replace '[\s/]', '').Trim() }
function ToDate($v, $no) {
  if ($v -is [double] -and $v -gt 30000) { return [DateTime]::FromOADate($v).Date }
  $t = "$v".Trim(); $dt = [DateTime]::MinValue
  if ($t -match '^(\d{4})[-./](\d{1,2})[-./](\d{1,2})') { return [DateTime]::new([int]$matches[1], [int]$matches[2], [int]$matches[3]) }
  if ($t -match '^(\d{1,2})/(\d{1,2})$') { return [DateTime]::new($Y, [int]$matches[1], [int]$matches[2]) }
  if ("$no" -match '^(\d{4})(\d{2})(\d{2})') { try { return [DateTime]::new([int]$matches[1], [int]$matches[2], [int]$matches[3]) } catch { } }
  $null }

# ── 설정 ──
$CUST = @{}; $BYMG = @{}; $EXCL = @()
Import-Csv (Join-Path $toolDir '국내전표_거래처.csv') -Encoding UTF8 | ForEach-Object { if ($_.제외 -eq '1') { $EXCL += $_.키; return }; $CUST[$_.키] = $_; $BYMG[$(if ($_.마감키) { $_.마감키 } else { $_.키 })] = $_ }
$SITEN = @{}; Import-Csv (Join-Path $toolDir '해외전표_사업장.csv') -Encoding UTF8 | ForEach-Object { $SITEN[($_.사업장명 -replace '\s', '')] = $_.부가세사업장코드 }
function IsExcl($k) { foreach ($e in $EXCL) { if ($k -like "$e*") { return $true } }; $false }
$code = 1
try {
Say "══ $($Y)년 $($M)월 국내매출 자동전표 ══" 'Cyan'
foreach ($f in $LineFile, $CloseFile, $template) { if (-not $f -or -not (Test-Path -LiteralPath $f)) { throw "파일이 없습니다: $f (국내자동전표 폴더에 라인별 매출확정·매출마감 확정 파일을 넣어 주세요)" } }
Say "  라인별 매출확정: $([IO.Path]::GetFileName($LineFile)) / 시트 $LineSheet"
Say "  매출마감 확정: $([IO.Path]::GetFileName($CloseFile)) / 시트 $CloseSheet"

$xl = New-Object -ComObject Excel.Application; $xl.DisplayAlerts = $false
try {
  function OpenSheet($path, $name) { $bk = $xl.Workbooks.Open($path, 0, $true); $ws = $null; foreach ($s0 in $bk.Worksheets) { if (($s0.Name -replace '\s', '') -eq ($name -replace '\s', '')) { $ws = $s0 } }
    if (-not $ws) { $bk.Close($false); throw "'$([IO.Path]::GetFileName($path))'에 '$name' 시트가 없습니다." }; , @($bk, $ws) }

  # ── ① 라인별 매출확정: 거래처별 사형·자동·금형·가공3 ──
  $bw = OpenSheet $LineFile $LineSheet; $v = $bw[1].Range('A1:R400').Value2; $bw[0].Close($false)
  $hr = 0; for ($r = 1; $r -le 30 -and -not $hr; $r++) { if ((NormHead $v[$r, 1]) -eq '업체명') { $hr = $r } }
  if (-not $hr) { throw "라인별 매출확정 '$LineSheet' 시트에서 '업체명' 머리글을 찾지 못했습니다." }
  $LC = @{}; for ($c = 2; $c -le 18; $c++) { foreach ($rr in $hr, ($hr + 1)) { $h = NormHead $v[$rr, $c]; if (-not $h) { continue }
      $nm = switch -regex ($h) { '^사형$' { '사형' } '^자동$' { '자동' } '^금형$' { '금형' } '^가공3$' { '가공3' } '^기타' { '기타' } '^선급금' { '선급금' } '^금형비$' { '금형비' } '^합계' { '합계' } '^비고$' { '비고' } default { $null } }
      if ($nm -and -not $LC.ContainsKey($nm)) { $LC[$nm] = $c } } }
  foreach ($need in '사형', '자동', '금형', '가공3', '합계') { if (-not $LC.ContainsKey($need)) { throw "라인별 매출확정에서 '$need' 칸을 찾지 못했습니다." } }
  $LROWS = New-Object Collections.Generic.List[object]
  for ($r = $hr + 2; $r -le 400; $r++) { $a = ("$($v[$r,1])" -replace '\s+', ' ').Trim(); if (-not $a) { $any = $false; foreach ($c in 2..10) { if (NumV $v[$r, $c]) { $any = $true } }; if ($any) { break } else { continue } }
    $tot = NumV $v[$r, $LC['합계']]; $parts = [ordered]@{}
    foreach ($p in '사형', '자동', '금형', '가공3') { $x = NumV $v[$r, $LC[$p]]; if ($x -and [Math]::Abs($x) -ge 0.5) { $parts[$p] = [Math]::Round($x, 0) } }
    $extra = @(); foreach ($p in '기타', '선급금', '금형비') { if ($LC.ContainsKey($p)) { $x = NumV $v[$r, $LC[$p]]; if ($x -and [Math]::Abs($x) -ge 0.5) { $extra += "$p $('{0:N0}' -f $x)" } } }
    if ((-not $tot -or [Math]::Abs($tot) -lt 0.5) -and $parts.Count -eq 0) { continue }
    $k = NormKey $a
    if (IsExcl $k) { continue }
    $LROWS.Add([pscustomobject]@{ 행 = $r; 이름 = $a; 키 = $k; 합계 = [Math]::Round([double]$tot, 0); 라인 = $parts; 기타 = ($extra -join ', '); 비고 = "$($v[$r, $(if ($LC['비고']) { $LC['비고'] } else { 11 })])".Trim(); 짝 = $false }) | Out-Null }
  Say "  라인별 매출 줄(신코·코마츠 제외): $($LROWS.Count)"

  # ── ② 매출마감 확정: 세금계산서 번호·금액·날짜 ──
  $bw = OpenSheet $CloseFile $CloseSheet; $v = $bw[1].Range('A1:P400').Value2; $bw[0].Close($false)
  $hr = 0; for ($r = 1; $r -le 30 -and -not $hr; $r++) { if ((NormHead $v[$r, 1]) -eq '거래처' -and (NormHead $v[$r, 2]) -like '세금계산서*') { $hr = $r } }
  if (-not $hr) { throw "매출마감 확정 '$CloseSheet' 시트에서 '거래처 / 세금계산서 NO.' 머리글을 찾지 못했습니다." }
  $cNo = 2; $cKind = 3; $cAmt = 0; $cBL = 0; $cIss = 0
  for ($c = 2; $c -le 16; $c++) { $h0 = NormHead $v[$hr, $c]; $h1 = NormHead $v[($hr + 1), $c]
    if ($h0 -like '종류*') { $cKind = $c }; if ($h0 -eq 'BL') { $cBL = $c }; if ($h0 -like '주요*') { $cIss = $c }; if ($h1 -eq '실적') { $cAmt = $c } }
  if (-not $cAmt -or -not $cBL) { throw "매출마감 확정에서 '실적' 또는 'BL' 칸을 찾지 못했습니다." }
  $INV = [ordered]@{}; $curName = ''; $hint = @{}
  for ($r = $hr + 2; $r -le 400; $r++) { $a = ("$($v[$r,1])" -replace '\s+', ' ').Trim(); if ($a -like '세금계산서*발행*') { break }
    if ($a) { $curName = $a }; if (-not $curName) { continue }
    $k = NormKey $curName; if ("$($v[$r,$cIss])" -match '종사업장') { $hint[$k] = $true }
    $no = "$($v[$r,$cNo])".Trim(); if ($no -eq '계') { continue }
    $amt = NumV $v[$r, $cAmt]; if (-not $amt -or [Math]::Abs($amt) -lt 0.5) { continue }
    if (IsExcl $k) { continue }
    if (-not $no -and -not "$($v[$r,$cKind])".Trim()) { continue }   # 번호·종류 없이 금액만 있는 줄 = 소계('계' 글자가 빠진 줄, 예: 9월 엘에스엠트론 110행)
    $id = if ($no) { "$k|$no" } else { "$k|행$r" }
    if (-not $INV.Contains($id)) { $INV[$id] = [pscustomobject]@{ 키 = $k; 이름 = $curName; 번호 = $no; 금액 = 0.0; 종류 = @(); 날짜 = (ToDate $v[$r, $cBL] $no); BL = $v[$r, $cBL]; 행 = @(); 짝 = $false; tx = $null } }
    $iv = $INV[$id]; $iv.금액 += [Math]::Round($amt, 0); $iv.종류 += "$("$($v[$r,$cKind])".Trim()) $('{0:N0}' -f $amt)"; $iv.행 += $r
    if (-not $iv.날짜) { $iv.날짜 = ToDate $v[$r, $cBL] $no } }
  $INVS = @($INV.Values)
  Say "  매출마감 세금계산서(신코·코마츠 제외): $($INVS.Count)장"

  # ── ③ 아마란스 세금계산서 목록(있으면): 세액·날짜·승인번호·사업장 확인용 ──
  $TAX = @(); $TAXBY = @{}
  if ($TaxFile) {
    $tb = $xl.Workbooks.Open($TaxFile, 0, $true); $tv = $tb.Worksheets.Item(1).UsedRange.Value2; $tb.Close($false)
    $tc = @{}; for ($c = 1; $c -le $tv.GetLength(1); $c++) { $tc[(NormHead $tv[1, $c])] = $c }
    for ($r = 2; $r -le $tv.GetLength(0); $r++) { if ("$($tv[$r,$tc['구분']])" -ne '매출') { continue }
      $d0 = ToDate $tv[$r, $tc['작성일']] ''; if (-not $d0 -or $d0.Year -ne $Y -or $d0.Month -ne $M) { continue }
      $t = [pscustomobject]@{ 작성일 = $d0; 사업장 = "$($tv[$r,$tc['사업장']])"; 거래처코드 = "$($tv[$r,$tc['거래처코드']])"; 거래처명 = "$($tv[$r,$tc['거래처명']])"; 공급가액 = [double](NumV $tv[$r, $tc['공급가액']]); 세액 = [double](NumV $tv[$r, $tc['세액']]); 승인번호 = "$($tv[$r,$tc['승인번호']])".Trim().ToLower(); 전표 = "$($tv[$r,$tc['전표여부']])"; 비고 = "$($tv[$r,$tc['비고']])"; 품목 = "$($tv[$r,$tc['대표품목']])"; 씀 = $false }
      $TAX += $t; if ($t.승인번호) { $TAXBY[$t.승인번호] = $t } }
    Say "  세금계산서 목록(이번 달 작성, 매출): $($TAX.Count)장" }
  function IssNo($no) { $n = ("$no" -replace '-', '').ToLower(); if ($n -match '^\d{8}[0-9a-z]{16}$') { $n } else { '' } }
  $setOfI = @{}; foreach ($iv in $INVS) { $st = $BYMG[$iv.키]; if (-not $st) { $st = $CUST[$iv.키] }; $setOfI[$iv.번호 + '|' + $iv.키] = $st }
  # 짝: ① 승인번호가 같고 공급가액도 같음 → ② 같은 거래처코드·같은 공급가액·같은 날짜 → ③ 같은 거래처코드·같은 공급가액(이번 달 안에서 하나뿐)
  #  9월 예: 엘에스엠트론 BL 칸 9/20이지만 실제 세금계산서 9/30, 브이텍스 매출마감 번호는 수정 전(오류발행) 세금계산서 번호
  if ($TAX.Count) { foreach ($iv in $INVS) { $st = $setOfI[$iv.번호 + '|' + $iv.키]; $iss = IssNo $iv.번호; $tx = $null; $wrongNo = $null
      if ($iss -and $TAXBY.ContainsKey($iss)) { if ([Math]::Abs($TAXBY[$iss].공급가액 - $iv.금액) -lt 0.5) { $tx = $TAXBY[$iss] } else { $wrongNo = $TAXBY[$iss] } }
      if ($tx -and $tx.씀) { $tx = $null }
      if (-not $tx -and $st) { $cand = @($TAX | Where-Object { -not $_.씀 -and $_.거래처코드 -eq $st.거래처코드 -and [Math]::Abs($_.공급가액 - $iv.금액) -lt 0.5 })
        $same = @($cand | Where-Object { $iv.날짜 -and $_.작성일 -eq $iv.날짜 }); if ($same.Count) { $tx = $same[0] } elseif ($cand.Count -eq 1) { $tx = $cand[0] } }
      if ($tx) { $tx.씀 = $true; $iv.tx = $tx
        if ($wrongNo) { Chk '확인' $iv.이름 "매출마감 확정의 세금계산서 번호 $($iv.번호)는 세금계산서 목록에서 공급가액 $('{0:N0}' -f $wrongNo.공급가액)원짜리(수정 전 번호로 보임)입니다. 금액이 같은 세금계산서(승인번호 $($tx.승인번호))로 넣었습니다." }
        if ($iv.날짜 -and $tx.작성일 -ne $iv.날짜) { Chk '확인' $iv.이름 "매출마감 확정 BL 칸 날짜는 $($iv.날짜.ToString('yyyy-MM-dd'))인데 세금계산서 목록 작성일은 $($tx.작성일.ToString('yyyy-MM-dd'))입니다. 세금계산서 작성일로 넣었습니다($('{0:N0}' -f $iv.금액)원)." }
        $iv.날짜 = $tx.작성일 }
      elseif ($wrongNo) { Chk '확인' $iv.이름 "매출마감 확정의 세금계산서 번호 $($iv.번호)는 세금계산서 목록에서 공급가액 $('{0:N0}' -f $wrongNo.공급가액)원짜리인데 매출마감은 $('{0:N0}' -f $iv.금액)원입니다. 번호를 확인해 주세요." } } }
  foreach ($iv in $INVS) { if ($iv.tx) { continue }; if ($iv.번호 -match '^(\d{4})-?(\d{2})-?(\d{2})' -and $iv.날짜) { $nd = $null; try { $nd = [DateTime]::new([int]$matches[1], [int]$matches[2], [int]$matches[3]) } catch { }
      if ($nd -and $nd -ne $iv.날짜) { Chk '확인' $iv.이름 "세금계산서 번호의 날짜($($nd.ToString('yyyy-MM-dd')))와 BL 칸 날짜($($iv.날짜.ToString('yyyy-MM-dd')))가 다릅니다. BL 칸 날짜로 넣었습니다(번호 $($iv.번호))." } }
    if (-not $iv.날짜) { $iv.날짜 = $cur.AddMonths(1).AddDays(-1)
      Chk '확인' $iv.이름 "매출마감 확정 $($iv.행 -join ',')행에 세금계산서 번호·BL 날짜가 없고 세금계산서 목록에도 없어, 말일($($iv.날짜.ToString('yyyy-MM-dd')))로 넣었습니다($('{0:N0}' -f $iv.금액)원). 세금계산서가 발행됐는지, 날짜가 맞는지 확인해 주세요(발행 전이면 이 전표는 빼고 올리세요)." }
    elseif ($iv.날짜.Year -ne $Y -or $iv.날짜.Month -ne $M) { Chk '확인' $iv.이름 "세금계산서 날짜 $($iv.날짜.ToString('yyyy-MM-dd'))가 이번 달이 아닙니다(번호 $($iv.번호))." } }

  # ── ④ 라인별 줄 ↔ 세금계산서 짝짓기 (거래처 안에서 금액으로) ──
  $GROUPS = New-Object Collections.Generic.List[object]
  $setOfL = @{}; foreach ($lr in $LROWS) { $st = $CUST[$lr.키]; if (-not $st) { $st = $BYMG[$lr.키] }; $setOfL[$lr.행] = $st }
  foreach ($lr in $LROWS) { if (-not $setOfL[$lr.행]) { Chk '에러' $lr.이름 "거래처 설정(국내전표_거래처.csv)에 '$($lr.키)'가 없습니다. 라인별 매출확정 $($lr.행)행($('{0:N0}' -f $lr.합계)원)은 넣지 않았습니다. 거래처코드·적요를 설정표에 넣고 다시 만들어 주세요." } }
  foreach ($iv in $INVS) { if (-not $setOfI[$iv.번호 + '|' + $iv.키]) { Chk '에러' $iv.이름 "거래처 설정(국내전표_거래처.csv)에 '$($iv.키)'가 없습니다. 세금계산서 $($iv.번호) ($('{0:N0}' -f $iv.금액)원)는 넣지 않았습니다." } }
  $custKeys = @($LROWS | ForEach-Object { $setOfL[$_.행] } | Where-Object { $_ } | ForEach-Object { $_.키 }) + @($INVS | ForEach-Object { $setOfI[$_.번호 + '|' + $_.키] } | Where-Object { $_ } | ForEach-Object { $_.키 }) | Select-Object -Unique
  foreach ($ck in $custKeys) { $st = $CUST[$ck]
    $Ls = @($LROWS | Where-Object { $setOfL[$_.행] -and $setOfL[$_.행].키 -eq $ck }); $Is = @($INVS | Where-Object { $setOfI[$_.번호 + '|' + $_.키] -and $setOfI[$_.번호 + '|' + $_.키].키 -eq $ck })
    # 1) 금액이 똑같은 것  2) 100원 이내 차이
    foreach ($tol in 0, 100) { foreach ($iv in ($Is | Where-Object { -not $_.짝 })) { $m0 = $Ls | Where-Object { -not $_.짝 -and [Math]::Abs($_.합계 - $iv.금액) -le $tol } | Select-Object -First 1
        if ($m0) { $m0.짝 = $true; $iv.짝 = $true; $GROUPS.Add([pscustomobject]@{ Id = $GROUPS.Count + 1; 설정 = $st; L = @($m0); I = @($iv) }) | Out-Null } } }
    # 3) 라인별 한 줄 = 남은 세금계산서 여러 장의 합 (예: 에스앤티 1,776,000 + 382,460,742) / 세금계산서 한 장 = 남은 라인별 여러 줄의 합
    $rl = @($Ls | Where-Object { -not $_.짝 }); $ri = @($Is | Where-Object { -not $_.짝 })
    if ($rl.Count -ge 1 -and $ri.Count -ge 1) { $sl = ($rl | Measure-Object 합계 -Sum).Sum; $si = ($ri | Measure-Object 금액 -Sum).Sum
      if ([Math]::Abs($sl - $si) -le 100 -and ($rl.Count -eq 1 -or $ri.Count -eq 1)) { foreach ($x in $rl) { $x.짝 = $true }; foreach ($x in $ri) { $x.짝 = $true }; $GROUPS.Add([pscustomobject]@{ Id = $GROUPS.Count + 1; 설정 = $st; L = $rl; I = $ri }) | Out-Null } }
    foreach ($x in ($Ls | Where-Object { -not $_.짝 })) { Chk '에러' $x.이름 "라인별 매출확정 $($x.행)행($('{0:N0}' -f $x.합계)원)과 금액이 맞는 세금계산서가 매출마감 확정에 없습니다. 넣지 않았습니다." }
    foreach ($x in ($Is | Where-Object { -not $_.짝 })) { Chk '에러' $x.이름 "세금계산서 $($x.번호) ($('{0:N0}' -f $x.금액)원)와 금액이 맞는 줄이 라인별 매출확정에 없어 라인을 나눌 수 없습니다. 넣지 않았습니다." } }

  # ── ⑤ 전표 만들기 ──
  $LINES = New-Object Collections.Generic.List[object]; $VCH = New-Object Collections.Generic.List[object]; $seqByDate = @{}
  $projFor = @{ 사형 = '101'; 자동 = '102'; 금형 = '201'; 가공3 = '403' }
  function GFirst($g) { $mn = 99999; foreach ($x in $g.L) { if ($x.행 -lt $mn) { $mn = $x.행 } }; $mn }
  function GDate($g) { $mx = $null; foreach ($x in $g.I) { if ($x.날짜 -and (-not $mx -or $x.날짜 -gt $mx)) { $mx = $x.날짜 } }; $mx }
  $ordered = @($GROUPS | Sort-Object { GFirst $_ })
  # 작성번호: 날짜마다 $StartNo(2000)부터, 같은 날짜 안에서는 라인별 매출확정 순서
  $numOf = @{}; foreach ($g in ($ordered | Sort-Object { $d0 = GDate $_; if ($d0) { $d0.ToString('yyyyMMdd') } else { '99999999' } }, { GFirst $_ })) {
    $d = GDate $g; if (-not $d) { continue }; $d8 = $d.ToString('yyyyMMdd')
    if (-not $seqByDate.ContainsKey($d8)) { $seqByDate[$d8] = $StartNo - 1 }; $seqByDate[$d8]++; $numOf[$g.Id] = @($d8, $seqByDate[$d8]) }
  foreach ($g in $ordered) { $st = $g.설정; $who = $g.L[0].이름
    if (-not $numOf.ContainsKey($g.Id)) { continue }
    $d8 = $numOf[$g.Id][0]; $no = $numOf[$g.Id][1]
    if ($st.근거 -like '추정*') { Chk '확인' $who "거래처 설정이 추정입니다($($st.근거)). 거래처코드 $($st.거래처코드), 부가세사업장 $($st.부가세사업장), 적요를 확인해 주세요." }
    # 라인별 금액
    $parts = [ordered]@{}; foreach ($lr in $g.L) { foreach ($p in $lr.라인.Keys) { $parts[$p] = [double]$parts[$p] + $lr.라인[$p] }; if ($lr.기타) { Chk '확인' $who "라인별 매출확정 $($lr.행)행에 '$($lr.기타)' 금액이 있습니다(라인을 정할 수 없어 넣지 않음). 맞는지 봐 주세요." } }
    $tot = ($g.I | Measure-Object 금액 -Sum).Sum
    if ($parts.Count -eq 0) { $parts['사형'] = $tot; Chk '에러' $who "라인별 금액이 없어 전액을 사형 한 줄로 넣었습니다." }
    $sum = 0.0; foreach ($p in @($parts.Keys)) { $sum += $parts[$p] }
    # 차이는 가공3 줄에서 맞춤(없으면 가장 큰 줄) — 9월 STX 방산 144,810,010 → 144,810,000을 실제 전표에서 가공3으로 맞춘 것과 같게
    if ([Math]::Abs($sum - $tot) -ge 0.5) { $big = if ($parts.Contains('가공3')) { '가공3' } else { @($parts.Keys | Sort-Object { -[Math]::Abs($parts[$_]) })[0] }; $parts[$big] = $parts[$big] + ($tot - $sum)
      Chk '확인' $who "라인별 합계($('{0:N0}' -f $sum))가 세금계산서 금액($('{0:N0}' -f $tot))과 $('{0:N0}' -f ($tot - $sum))원 달라, 차이를 '$big' 줄에서 맞췄습니다(세금계산서 금액 기준, 9월 실제 전표와 같은 방법)." }
    if (@($g.I).Count -gt 1) { Chk '확인' $who "세금계산서 $(@($g.I).Count)장($((($g.I | ForEach-Object { '{0:N0}' -f $_.금액 }) -join ' + ')))이 라인별 매출확정에는 한 줄로 합쳐져 있어, 전표 1장에 부가세·외상 줄을 세금계산서마다 넣었습니다. 세금계산서별로 라인을 알려 주시면 전표를 나누겠습니다." }
    foreach ($lr in $g.L) { if ($lr.비고 -and $lr.비고 -ne '0') { Chk '참고' $who "라인별 매출확정 $($lr.행)행 비고: '$($lr.비고)'" } }
    foreach ($iv in $g.I) { if (@($iv.종류).Count -gt 1) { Chk '확인' $who "세금계산서 $($iv.번호) 한 장에 '$($iv.종류 -join ' / ')'가 같이 있습니다. 라인별 매출확정대로 넣었습니다($((($parts.Keys | ForEach-Object { "$_ $('{0:N0}' -f $parts[$_])" }) -join ', '))). 적요를 나눠야 하면 알려 주세요." } }
    $ln = 0
    foreach ($p in $parts.Keys) { $ln++
      $LINES.Add([pscustomobject]@{ 회계단위 = '1000'; 작성일자 = $d8; 작성번호 = $no; 라인순번 = $ln; 차대구분 = 4; 계정과목 = 4130000; 거래처 = $st.거래처코드; 거래처명 = $st.거래처명; 계정금액 = [double]$parts[$p]; 증빙코드 = ''; 적요 = ($st.라인적요 -replace '\{라인\}', $p); 부가세사업장 = ''; 신고기준일 = ''; 세무구분 = ''; 공급가액 = $null; 전자 = ''; 승인번호 = ''; 프로젝트 = $projFor[$p]; 품의내역 = $st.품의내역; 전표유형 = '3'; 순서 = $g.L[0].행; 묶음 = $g.Id }) | Out-Null }
    $vatSum = 0.0; $ivInfo = @()
    foreach ($iv in ($g.I | Sort-Object 날짜, 번호)) {
      $iss = IssNo $iv.번호; $tx = $iv.tx
      $site = $st.부가세사업장; $taxSrc = ''
      if ($tx) { $vat = $tx.세액; $taxSrc = '세금계산서 목록'
        $tsite = $SITEN[($tx.사업장 -replace '\s', '')]; if ($tsite -and $tsite -ne $site) { Chk '확인' $who "세금계산서 목록의 공급자 사업장은 '$($tx.사업장)'($tsite)인데 설정표는 $site 입니다. 세금계산서 목록대로 $tsite 로 넣었습니다."; $site = $tsite }
        if ($tx.거래처코드 -and $tx.거래처코드 -ne $st.거래처코드) { Chk '확인' $who "세금계산서 목록의 거래처코드($($tx.거래처코드))와 설정표($($st.거래처코드))가 다릅니다." }
        if ($tx.승인번호) { $iss = $tx.승인번호 } }
      else { $vat = [Math]::Floor($iv.금액 / 10 + 1e-9); $taxSrc = '계산(10% 원 미만 버림)'
        if ([Math]::Abs($iv.금액 / 10 - $vat) -gt 1e-6) { Chk '확인' $who "세액 $('{0:N0}' -f $vat)원은 공급가액 $('{0:N0}' -f $iv.금액)원의 10%에서 원 미만을 버린 값입니다(세금계산서 목록에 없음). 실제 세금계산서 세액과 다르면 고쳐 주세요." } }
      if ($hint[$iv.키] -and $site -eq '1000') { Chk '확인' $who "매출마감 확정 '주요 issue' 칸에 '종사업장'이 적혀 있는데 부가세사업장이 1000(대신금속)입니다. 마산3 등 종사업장 발행인지 봐 주세요." }
      if (-not $iss) { if ($iv.번호) { Chk '알림' $who "세금계산서 번호 '$($iv.번호)'가 국세청 승인번호 모양(24자리)이 아니라 승인번호 칸을 비웠습니다." } }
      elseif ($iss.Substring(0, 8) -ne $iv.날짜.ToString('yyyyMMdd')) { Chk '확인' $who "세금계산서 번호 '$($iv.번호)'의 앞 8자리 날짜가 세금계산서 날짜($($iv.날짜.ToString('yyyy-MM-dd')))와 달라 오타로 보입니다. 잘못 연결되지 않게 승인번호 칸은 비웠습니다(아마란스에서 세금계산서와 직접 연결하거나 번호를 고쳐 다시 만들기)."; $iss = '' }
      if ($NoIssNo) { $iss = '' }
      $ivInfo += [pscustomobject]@{ iv = $iv; 세액 = $vat; 사업장 = $site; 승인 = $iss; 출처 = $taxSrc }; $vatSum += $vat }
    foreach ($q in $ivInfo) { $ln++
      $LINES.Add([pscustomobject]@{ 회계단위 = '1000'; 작성일자 = $d8; 작성번호 = $no; 라인순번 = $ln; 차대구분 = 4; 계정과목 = 2550000; 거래처 = $st.거래처코드; 거래처명 = $st.거래처명; 계정금액 = [double]$q.세액; 증빙코드 = ''; 적요 = $st.부가세적요; 부가세사업장 = $q.사업장; 신고기준일 = $q.iv.날짜.ToString('yyyyMMdd'); 세무구분 = '11'; 공급가액 = [double]$q.iv.금액; 전자 = 'Y'; 승인번호 = $q.승인; 프로젝트 = ''; 품의내역 = $st.품의내역; 전표유형 = '3'; 순서 = $g.L[0].행; 묶음 = $g.Id }) | Out-Null }
    foreach ($q in $ivInfo) { $ln++
      $LINES.Add([pscustomobject]@{ 회계단위 = '1000'; 작성일자 = $d8; 작성번호 = $no; 라인순번 = $ln; 차대구분 = 3; 계정과목 = 1080000; 거래처 = $st.거래처코드; 거래처명 = $st.거래처명; 계정금액 = [double]($q.iv.금액 + $q.세액); 증빙코드 = '1'; 적요 = $st.부가세적요; 부가세사업장 = ''; 신고기준일 = ''; 세무구분 = ''; 공급가액 = $null; 전자 = ''; 승인번호 = ''; 프로젝트 = ''; 품의내역 = $st.품의내역; 전표유형 = '3'; 순서 = $g.L[0].행; 묶음 = $g.Id }) | Out-Null }
    $VCH.Add([pscustomobject]@{ 순서 = $g.L[0].행; 작성일자 = $d8; 작성번호 = $no; 거래처 = $who; 거래처코드 = $st.거래처코드; 아마란스거래처 = $st.거래처명
        부가세사업장 = (($ivInfo | ForEach-Object { $_.사업장 } | Select-Object -Unique) -join ','); 세금계산서 = (($ivInfo | ForEach-Object { "$($_.iv.날짜.ToString('MM-dd')) $($_.iv.번호)" }) -join ' / ')
        공급가액 = [double]$tot; 세액 = [double]$vatSum; 합계 = [double]($tot + $vatSum); 세액출처 = (($ivInfo | ForEach-Object { $_.출처 } | Select-Object -Unique) -join ',')
        라인 = (($parts.Keys | ForEach-Object { "$_ $('{0:N0}' -f $parts[$_])" }) -join ' / '); 라인별행 = (($g.L | ForEach-Object { "$($_.행)행 $($_.이름)" }) -join ', '); 마감행 = (($g.I | ForEach-Object { $_.행 -join ',' }) -join ' / '); 품의내역 = $st.품의내역; 묶음 = $g.Id }) | Out-Null }

  # 차대 점검
  foreach ($gg in ($LINES | Group-Object 작성일자, 작성번호)) { $dr = 0.0; $cr = 0.0; foreach ($l in $gg.Group) { if ($l.차대구분 -eq 3) { $dr += $l.계정금액 } else { $cr += $l.계정금액 } }
    if ([Math]::Abs($dr - $cr) -ge 0.5) { Chk '에러' $gg.Group[0].거래처명 "차변 $dr 과 대변 $cr 이 다릅니다." } }
  # 합계 대조
  $lSum = ($LROWS | Measure-Object 합계 -Sum).Sum; $iSum = ($INVS | Measure-Object 금액 -Sum).Sum; $vSum = ($VCH | Measure-Object 공급가액 -Sum).Sum
  Chk '참고' '' "공급가액 합계: 전표 $('{0:N0}' -f $vSum)원 / 라인별 매출확정(신코·코마츠 제외) $('{0:N0}' -f $lSum)원 / 매출마감 확정 세금계산서(신코·코마츠 제외) $('{0:N0}' -f $iSum)원 → 차이 $('{0:N0}' -f ($vSum - $iSum))원"
  # 세금계산서 목록에만 있는 것
  foreach ($t in ($TAX | Where-Object { -not $_.씀 })) { $known = @($CUST.Values | Where-Object { $_.거래처코드 -eq $t.거래처코드 }).Count -gt 0
    if ("$($t.품목) $($t.비고)" -match '외주가공|임대|임차|고철|스크랩|매각|차량') { $known = $false }   # 같은 거래처라도 제품 매출이 아닌 것(예: 알멕 외주가공비 = 마산6)
    $desc = "$($t.작성일.ToString('MM-dd')) $($t.거래처명) 공급가액 $('{0:N0}' -f $t.공급가액) (품목 $($t.품목)$(if ($t.비고) { ", $($t.비고)" }), 아마란스 전표: $($t.전표))"
    if ($known) { Chk '알림' $t.거래처명 "아마란스 세금계산서 목록에는 있는데 매출마감 확정에 없어 전표를 만들지 않았습니다: $desc. 수정세금계산서(발행·취소)이면 9월 브이텍스처럼 따로 넣어 주세요." }
    else { Chk '참고' $t.거래처명 "제품 매출이 아닌 세금계산서(고철·스크랩·아트라스 등, 손으로 넣는 것)라 이 전표에서 뺐습니다: $desc" } }

  $LINES = @($LINES | Sort-Object 순서, 작성일자, 작성번호, 라인순번); $VCH = @($VCH | Sort-Object 순서, 작성일자, 작성번호)
  $LINES | Select-Object * -ExcludeProperty 묶음 | Export-Csv (Join-Path $OutDir "국내매출_자동전표_$($YM)_줄.csv") -NoTypeInformation -Encoding UTF8
  Say "  전표 $($VCH.Count)장, $($LINES.Count)줄, 공급가액 $('{0:N0}' -f $vSum) / 세액 $('{0:N0}' -f (($VCH | Measure-Object 세액 -Sum).Sum)) / 외상매출금 $('{0:N0}' -f (($VCH | Measure-Object 합계 -Sum).Sum))" 'Green'

  # ── ⑥ 업로드 엑셀 (아마란스 '자동전표처리(일반)' 전체 칸 양식, 4행부터, 2행 칸 코드로 넣음) ──
  $upload = Join-Path $OutDir "국내매출_자동전표_$($YM).xlsx"
  $tmpUp = Join-Path $workDir "만드는중_업로드_$(Get-Date -Format 'HHmmss').xlsx"
  if (Test-Path -LiteralPath $tmpUp) { Remove-Item -LiteralPath $tmpUp -Force }
  $wb = $xl.Workbooks.Open($template, 0, $true); $wb.SaveAs($tmpUp, 51)
  if ($wb.ReadOnly) { $wb.Close($false); throw "업로드 임시 파일을 쓰기 가능한 상태로 만들지 못했습니다: $tmpUp" }
  $sh = $wb.Worksheets.Item(1)
  $maxC = $sh.Cells(2, $sh.Columns.Count).End(-4159).Column
  $colOf = @{}; $isText = @{}
  for ($ci2 = 1; $ci2 -le $maxC; $ci2++) { $cd = "$($sh.Cells(2, $ci2).Value2)".Trim(); if ($cd -and -not $colOf.ContainsKey($cd)) { $colOf[$cd] = $ci2; $isText[$ci2] = ("$($sh.Cells(3, $ci2).Value2)" -match '타입\s*:\s*문자') } }
  $codeMap = [ordered]@{ IN_DIV_CD = '회계단위'; MENU_DT = '작성일자'; MENU_SQ = '작성번호'; MENU_LN_SQ = '라인순번'; ISU_DOC = '품의내역'; DOCU_TY = '전표유형'; DRCR_FG = '차대구분'; ACCT_CD = '계정과목'; TR_CD = '거래처'; TR_NM = '거래처명'; ACCT_AM = '계정금액'; ATTR_CD = '증빙코드'; RMK_DC = '적요'
    VAT_DIV_CD = '부가세사업장'; ISS_DT = '신고기준일'; TAX_FG = '세무구분'; SUP_AM = '공급가액'; JEONJA_YN = '전자'; ISS_NO = '승인번호'; PJT_CD = '프로젝트' }
  foreach ($cd in $codeMap.Keys) { if (-not $colOf.ContainsKey($cd)) { Chk $(if ($cd -in 'ISS_DT', 'VAT_DIV_CD', 'TAX_FG', 'SUP_AM') { '에러' } else { '확인' }) '' "업로드 양식에 '$cd' 칸이 없어 '$($codeMap[$cd])' 값을 넣지 못했습니다. 아마란스에서 받은 전체 칸 양식을 넣어 주세요." } }
  $n = $LINES.Count
  if ($n) {
    foreach ($cd in $codeMap.Keys) { if ($colOf.ContainsKey($cd) -and $isText[$colOf[$cd]]) { $sh.Range($sh.Cells(4, $colOf[$cd]), $sh.Cells(3 + $n, $colOf[$cd])).NumberFormat = '@' } }
    $data = New-Object 'object[,]' $n, $maxC
    for ($i = 0; $i -lt $n; $i++) { $l = $LINES[$i]
      foreach ($cd in $codeMap.Keys) { if (-not $colOf.ContainsKey($cd)) { continue }; $val = $l.($codeMap[$cd]); if ($null -eq $val -or "$val" -eq '') { continue }; $cx = $colOf[$cd] - 1
        if ($isText[$colOf[$cd]]) { $data.SetValue([object][string]$val, [int]$i, [int]$cx) } else { $data.SetValue([object][double]$val, [int]$i, [int]$cx) } } }
    $sh.Range($sh.Cells(4, 1), $sh.Cells(3 + $n, $maxC)).Value2 = $data
    foreach ($fm in @(@('ACCT_AM', '#,##0'), @('SUP_AM', '#,##0'))) { if ($colOf.ContainsKey($fm[0])) { $sh.Range($sh.Cells(4, $colOf[$fm[0]]), $sh.Cells(3 + $n, $colOf[$fm[0]])).NumberFormat = $fm[1] } } }
  $wb.Save(); $wb.Close($false)
  $wbC = $xl.Workbooks.Open($tmpUp, 0, $true); $lastC = $wbC.Worksheets.Item(1).Cells($wbC.Worksheets.Item(1).Rows.Count, 1).End(-4162).Row; $wbC.Close($false)
  if ($lastC -ne 3 + $n) { throw "업로드 파일에 데이터가 제대로 저장되지 않았습니다(저장된 마지막 줄 $lastC, 있어야 할 줄 $(3 + $n)): $tmpUp" }
  if (Test-Path -LiteralPath $upload) { try { Remove-Item -LiteralPath $upload -Force -ErrorAction Stop } catch { throw "업로드 파일을 바꿀 수 없습니다. 엑셀에서 열려 있으면 닫고 다시 해 주세요(새 파일은 $tmpUp 에 있음): $upload" } }
  Move-Item -LiteralPath $tmpUp $upload -Force
  Say "  업로드 파일: $upload ($($n)줄 저장 확인)" 'Green'

  # ── ⑦ 실제 전표와 대조(시험용: 참고1에서 뽑은 국내매출 전표 CSV) ──
  $CMP = @()
  if ($CompareCsv) {
    $act = Import-Csv $CompareCsv -Encoding UTF8 | Where-Object { $_.작성일 -like "$($cur.ToString('yyyy-MM'))*" }
    $AV = @($act | Group-Object { "$($_.작성일)|$($_.작성번호)" } | ForEach-Object { $gr = $_.Group; $ar = 0.0; $vt = 0.0; $pj = @{}; $rm = @()
        foreach ($x in $gr) { $dv = [double](NumV $x.차변금액); $cv = [double](NumV $x.대변금액)
          if ($x.계정과목코드 -eq '1080000') { $ar += $dv } elseif ($x.계정과목코드 -eq '2550000') { $vt += $cv } elseif ($x.계정과목코드 -eq '4130000') { $pj[$x.프로젝트코드] = [double]$pj[$x.프로젝트코드] + $cv; $rm += $x.적요 } }
        [pscustomobject]@{ 날짜 = $gr[0].작성일; 번호 = $gr[0].작성번호; 코드 = $gr[0].거래처코드; 이름 = $gr[0].거래처명; 외상 = $ar; 세액 = $vt; 프로젝트 = $pj; 적요 = $rm; 씀 = $false } })
    foreach ($v0 in $VCH) { $a = $AV | Where-Object { -not $_.씀 -and $_.코드 -eq $v0.거래처코드 -and [Math]::Abs($_.외상 - $v0.합계) -lt 1 } | Select-Object -First 1
      if (-not $a) { $a = $AV | Where-Object { -not $_.씀 -and $_.코드 -eq $v0.거래처코드 -and [Math]::Abs($_.외상 - $v0.합계) -lt 1000 } | Select-Object -First 1 }
      if (-not $a) { $CMP += [pscustomobject]@{ 거래처 = $v0.거래처; 만든전표 = "$($v0.작성일자) 합계 $('{0:N0}' -f $v0.합계)"; 실제전표 = ''; 결과 = '실제 전표 없음'; 내용 = '' }; continue }
      $a.씀 = $true; $msg = @()
      if (($a.날짜 -replace '-', '') -ne $v0.작성일자) { $msg += "날짜 실제 $($a.날짜) / 만든 것 $($v0.작성일자)" }
      if ([Math]::Abs($a.외상 - $v0.합계) -ge 1) { $msg += "외상 실제 $('{0:N0}' -f $a.외상) / 만든 것 $('{0:N0}' -f $v0.합계)" }
      if ([Math]::Abs($a.세액 - $v0.세액) -ge 1) { $msg += "세액 실제 $('{0:N0}' -f $a.세액) / 만든 것 $('{0:N0}' -f $v0.세액)" }
      $gp = @{}; foreach ($gl in ($LINES | Where-Object { $_.묶음 -eq $v0.묶음 -and $_.계정과목 -eq 4130000 })) { $gp[$gl.프로젝트] = [double]$gp[$gl.프로젝트] + $gl.계정금액 }
      foreach ($pj in (@($a.프로젝트.Keys) + @($gp.Keys) | Select-Object -Unique)) { if ([Math]::Abs([double]$a.프로젝트[$pj] - [double]$gp[$pj]) -ge 1) { $msg += "프로젝트 $pj 실제 $('{0:N0}' -f [double]$a.프로젝트[$pj]) / 만든 것 $('{0:N0}' -f [double]$gp[$pj])" } }
      $myRm = @($LINES | Where-Object { $_.묶음 -eq $v0.묶음 -and $_.계정과목 -eq 4130000 } | ForEach-Object { $_.적요 })
      $rmDiff = @($myRm | Where-Object { $_ -notin $a.적요 }); $rmNote = if ($rmDiff.Count) { "적요 다름(실제: $($a.적요 -join ' | '))" } else { '' }
      $CMP += [pscustomobject]@{ 거래처 = $v0.거래처; 만든전표 = "$($v0.작성일자) 합계 $('{0:N0}' -f $v0.합계)"; 실제전표 = "$($a.날짜) #$($a.번호)"; 결과 = $(if ($msg) { '다름' } elseif ($rmNote) { '금액 같음(적요만 다름)' } else { '같음' }); 내용 = (@($msg) + @($rmNote) | Where-Object { $_ }) -join ' / ' } }
    foreach ($a in ($AV | Where-Object { -not $_.씀 })) { $CMP += [pscustomobject]@{ 거래처 = $a.이름; 만든전표 = ''; 실제전표 = "$($a.날짜) #$($a.번호) 외상 $('{0:N0}' -f $a.외상)"; 결과 = '만든 전표 없음'; 내용 = ($a.적요 | Select-Object -First 1) } }
    $same = @($CMP | Where-Object { $_.결과 -like '같음' -or $_.결과 -like '금액 같음*' }).Count
    Say "  실제 전표와 대조: 금액 같음 $same / 만든 전표 $($VCH.Count)장 (실제 전표 $($AV.Count)장)" 'Cyan'
    $CMP | Where-Object { $_.결과 -notlike '같음' } | ForEach-Object { Say "    [$($_.결과)] $($_.거래처) $($_.만든전표) $($_.실제전표) $($_.내용)" } }

  # ── ⑧ 검토 엑셀 ──
  $review = Join-Path $OutDir "국내매출_자동전표_$($YM)_검토.xlsx"
  $wb = $xl.Workbooks.Add(); while ($wb.Worksheets.Count -lt 4) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  function FillSheet($sh, $hdrs, $rows, $props) { $nr = @($rows).Count; $dd = New-Object 'object[,]' ($nr + 1), $hdrs.Count
    for ($c = 0; $c -lt $hdrs.Count; $c++) { $dd.SetValue([object][string]$hdrs[$c], 0, $c) }
    $ri = 0; foreach ($rw in $rows) { $ri++; for ($c = 0; $c -lt $props.Count; $c++) { $val = $rw.($props[$c]); if ($null -eq $val) { continue }; $val = $val.psobject.BaseObject; if ($val -is [double] -or $val -is [int] -or $val -is [long] -or $val -is [decimal]) { $dd.SetValue([object][double]$val, $ri, $c) } else { $dd.SetValue([object][string]$val, $ri, $c) } } }
    $sh.Range($sh.Cells(1, 1), $sh.Cells($nr + 1, $hdrs.Count)).Value2 = $dd; $hh = $sh.Range($sh.Cells(1, 1), $sh.Cells(1, $hdrs.Count)); $hh.Font.Bold = $true; $hh.Interior.Color = 0xF2E6DD; [void]$sh.Columns.AutoFit() }
  $s1 = $wb.Worksheets.Item(1); $s1.Name = '전표 목록'
  FillSheet $s1 @('작성일자', '작성번호', '거래처(라인별 매출확정)', '거래처코드', '아마란스 거래처', '부가세 사업장', '세금계산서(날짜 번호)', '공급가액', '세액', '합계(외상매출금)', '세액 출처', '라인별 금액', '라인별 매출확정 행', '매출마감 확정 행', '품의내역') @($VCH) @('작성일자', '작성번호', '거래처', '거래처코드', '아마란스거래처', '부가세사업장', '세금계산서', '공급가액', '세액', '합계', '세액출처', '라인', '라인별행', '마감행', '품의내역')
  $s1.Range('H:J').NumberFormat = '#,##0'
  $s2 = $wb.Worksheets.Item(2); $s2.Name = '확인할 것'
  if ($CHK.Count) { $ord = @{ '에러' = 0; '확인' = 1; '알림' = 2; '참고' = 3 }; FillSheet $s2 @('구분', '거래처', '내용') @($CHK | Sort-Object { $ord[$_.구분] }, 순) @('구분', '거래처', '내용'); $s2.Columns('C').ColumnWidth = 140; $s2.Columns('C').WrapText = $true } else { $s2.Range('A1').Value2 = '확인할 것 없음' }
  $s3 = $wb.Worksheets.Item(3); $s3.Name = '세금계산서 목록 대조'
  if ($TAX.Count) { $trow = $TAX | Sort-Object 작성일 | ForEach-Object { [pscustomobject]@{ 작성일 = $_.작성일.ToString('yyyy-MM-dd'); 사업장 = $_.사업장; 거래처코드 = $_.거래처코드; 거래처명 = $_.거래처명; 공급가액 = $_.공급가액; 세액 = $_.세액; 승인번호 = $_.승인번호; 품목 = $_.품목; 비고 = $_.비고; 아마란스전표 = $_.전표; 이번자동전표 = $(if ($_.씀) { '넣음' } else { '안 넣음' }) } }
    FillSheet $s3 @('작성일', '공급자 사업장', '거래처코드', '거래처명', '공급가액', '세액', '승인번호', '대표품목', '비고', '아마란스 전표', '이번 자동전표') @($trow) @('작성일', '사업장', '거래처코드', '거래처명', '공급가액', '세액', '승인번호', '품목', '비고', '아마란스전표', '이번자동전표'); $s3.Range('E:F').NumberFormat = '#,##0' } else { $s3.Range('A1').Value2 = '세금계산서 목록 파일 없음' }
  $s4 = $wb.Worksheets.Item(4); $s4.Name = '실제 전표 대조'
  if ($CMP.Count) { FillSheet $s4 @('거래처', '만든 전표', '실제 전표', '결과', '내용') @($CMP) @('거래처', '만든전표', '실제전표', '결과', '내용') } else { $s4.Range('A1').Value2 = '대조 안 함' }
  $s1.Activate(); $wb.SaveAs($review, 51); $wb.Close($false)
  Say "  검토 파일: $review" 'Green'
}
finally { if ($xl) { try { foreach ($w0 in @($xl.Workbooks)) { $w0.Close($false) } } catch { }; $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } }
  Done 'OK' ("전표 $($VCH.Count)장, $($LINES.Count)줄, 외상매출금 합계 $('{0:N0}' -f (($VCH | Measure-Object 합계 -Sum).Sum))원`r`n확인할 것: $(@($CHK | Where-Object { $_.구분 -ne '참고' }).Count)건 (검토 파일 '확인할 것' 시트)`r`n업로드 파일: $upload"); $code = 0
}
catch { $m0 = "중간에 멈췄습니다: $($_.Exception.Message) (스크립트 $($_.InvocationInfo.ScriptLineNumber)행)"; Say $m0 'Red'; Done 'ERROR' $m0; $code = 1 }
exit $code
