$ErrorActionPreference = 'Stop'
$sp = Split-Path $MyInvocation.MyCommand.Path
$newF = 'C:\ai\매월마감\202610\2026년 수출 (2026.10)_10월반영.xlsx'
$W = Import-Csv "$sp\b_rows.csv" -Encoding UTF8
$SA = Import-Csv "$sp\c_A.csv" -Encoding UTF8 | ? { $_.BL -match '^\d{4}-' }
$LB = Import-Csv "$sp\c_B.csv" -Encoding UTF8 | ? { $_.인보이스 -and $_.인보이스 -notmatch '매출마감일' }
$P9 = Import-Csv "$sp\c_S9.csv" -Encoding UTF8 | ? 인보이스
$JPD = @(Import-Csv "$sp\c_JP.csv" -Encoding UTF8); $JSM = @{}; $lastJ = $null; Import-Csv "$sp\smbs.csv" -Encoding UTF8 | % { if ($_.JPY) { $lastJ = [double]$_.JPY }; $JSM[$_.날짜] = $lastJ }
$RT = @{}; Import-Csv "$sp\c_FX.csv" -Encoding UTF8 | ? 환율 | % { $RT[$_.날짜] = [double]$_.환율 }
function Put($s, $r, $vals, $c0 = 1) { $a = New-Object 'object[,]' 1, $vals.Count
  for ($i = 0; $i -lt $vals.Count; $i++) { $x = $vals[$i]; if ($null -ne $x) { $x = $x.psobject.BaseObject }; $a[0, $i] = $x }
  try { $s.Range($s.Cells($r, $c0), $s.Cells($r, $c0 + $vals.Count - 1)).Value2 = $a } catch { throw "Put 실패 $r 행: " + (($vals | % { if ($null -eq $_) { "null" } else { $_.GetType().Name + ":" + "$_".Substring(0, [Math]::Min(30, "$_".Length)) } }) -join " | ") } }
function Num($v) { if ($null -eq $v -or "$v" -eq '') { $null } else { [double]$v } }
$HD = 0xF2E6DD; $RED = 0xC7CEFF; $YEL = 0x99FFFF; $GRN = 0xCEEFC6; $GRY = 0xF2F2F2
$kName = @{ '전월' = '9월에서 넘어온 매출(녹색)'; '당월' = '10월 매출'; '툴링' = '툴링비'; '관세' = '관세'; '이월' = '11월로 이월(회색)' }

# ── 줄별 비교 ──
$cmp = foreach ($wr in $W) { $inv = $wr.원인보이스
  $a = $SA | ? 인보이스 -eq $inv | Select -First 1; $b = $LB | ? 인보이스 -eq $inv | Select -First 1
  $blR = if ($wr.BL -match '^\d{4}-') { $RT[$wr.BL] } else { $null }
  $msg = @(); $lvl = '맞음'
  $jp = $JPD | ? 인보이스 -eq $inv | Select -First 1
  if ($jp) { $blR = $JSM[$jp.BL]; $a = $null; $b = $null
    if ([Math]::Abs([double]$jp.환율 - $blR) -gt 0.00001) { $msg += "엔화 환율 $($jp.환율)이 서울외국환중개 $blR 과 다름"; $lvl = '에러' }
    if ([Math]::Abs((Num $wr.계산) - [double]$jp.원화) -ge 1) { $msg += "원화가 매출내역서 $([double]$jp.원화)와 다름"; $lvl = '에러' }
    $msg += "$($jp.파일) '$($jp.시트)' 시트 $($jp.행)행 (엔화)" }
  elseif ($wr.거래처 -like 'JS TECH*' -and $wr.종류 -eq '전월') { $msg += '국내영업(코마츠) 9월 이월분 — 해외 자료에 없어 9월 시트 값 그대로 넣음'; $lvl = '참고' }
  elseif (-not $inv) { $msg += '10월분 직접 입력용 빈 줄'; $lvl = '참고' }
  else {
    if (-not $a -and $wr.종류 -ne '이월') { $msg += '인보이스별 상세에 없음'; $lvl = '확인' }
    if (-not $b) { $msg += '라인별 매출에 없음'; $lvl = '확인' }
    if ($a -and $b) {
      if ($a.BL -ne $b.BL) { $msg += "BL날짜 다름(상세 $($a.BL), 라인별 $($b.BL))"; $lvl = '확인' }
      if ((Num $a.환율) -and (Num $b.환율) -and [Math]::Abs((Num $a.환율) - (Num $b.환율)) -gt 0.001) { $msg += '환율 다름'; $lvl = '확인' }
      if ((Num $a.외화) -and [Math]::Abs((Num $a.외화) - (Num $b.외화)) -gt 0.005) { $msg += '달러 다름'; $lvl = '확인' }
      if ([Math]::Abs((Num $a.원화) - (Num $b.원화)) -ge 1) { $msg += '원화 다름'; $lvl = '확인' } }
    if ($wr.종류 -eq '이월' -and -not $a) { $msg += '다음 달 이월이라 인보이스별 상세에는 없음(정상)' }
    if ($blR -and (Num $wr.환율) -and [Math]::Abs($blR - (Num $wr.환율)) -gt 0.001 -and $wr.거래처 -ne 'BERGEN ENG') { $msg += "기초자료 환율 $((Num $wr.환율).ToString('N2'))원이 BL일($($wr.BL)) 환율 $($blR.ToString('N2'))원과 다름 → 작업 파일은 BL일 환율로 계산"; $lvl = '에러' }
    if ($wr.원화 -and [Math]::Abs((Num $wr.계산) - (Num $wr.원화)) -ge 1) { $msg += "작업 파일 원화가 기초자료보다 $(((Num $wr.계산) - (Num $wr.원화)).ToString('N0'))원"; if ($lvl -eq '맞음') { $lvl = '확인' } }
    if ($wr.인보이스 -like '*오기재*') { $msg += '인보이스별 상세(ALLISON 관세 매출내역서)에만 있고 라인별 매출 당월 목록에는 없음'; $lvl = '에러' } }
  [pscustomobject]@{ 행 = [int]$wr.행; 거래처 = $wr.거래처; 구분 = $kName[$wr.종류]; 인보이스 = $wr.인보이스; BL상세 = $(if ($a) { $a.BL } elseif ($jp) { $jp.BL }); BL라인 = $(if ($b) { $b.BL }); 환율상세 = $(if ($a) { Num $a.환율 } elseif ($jp) { [double]$jp.환율 }); 환율라인 = $(if ($b) { Num $b.환율 }); BL일환율 = $blR
    달러상세 = $(if ($a) { Num $a.외화 } elseif ($jp) { [double]$jp.엔화 }); 달러라인 = $(if ($b) { Num $b.외화 }); 원화상세 = $(if ($a) { Num $a.원화 } elseif ($jp) { [double]$jp.원화 }); 원화라인 = $(if ($b) { Num $b.원화 }); 작업원화 = (Num $wr.계산); 결과 = $lvl; 내용 = ($msg -join ' / ') } }
# 상세에 있는데 작업 파일에 없는 것
$miss = $SA | ? { $inv = $_.인보이스; $inv -notlike '*오기재*' -and -not ($W | ? 원인보이스 -eq $inv) }

$xl = New-Object -ComObject Excel.Application; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
try {
  $wb = $xl.Workbooks.Open($newF); $ws = $wb.Worksheets.Item('26.10(NEW)')
  try { $old = $wb.Worksheets.Item('10월 기초자료 검증'); $old.Delete() } catch { }
  $s = $wb.Worksheets.Add([Type]::Missing, $ws); $s.Name = '10월 기초자료 검증'
  $s.Cells.Font.Name = '맑은 고딕'; $s.Cells.Font.Size = 10
  Put $s 1 @('2026년 10월 해외 기초자료 검증'); $s.Cells(1,1).Font.Size = 14; $s.Cells(1,1).Font.Bold = $true
  Put $s 2 @('비교한 파일: ■2026년 10월 매출마감_라인별_인보이스별 상세■(이하 "인보이스별 상세") ↔ ■26년 10월 수출마감_해외영업팀(라인별 매출)■(이하 "라인별 매출") ↔ 이 파일의 26.10(NEW) 시트(이하 "10월 시트")')
  Put $s 3 @('비교 항목: 인보이스 · BL날짜 · 환율 · 달러 · 원화. 10월 시트의 원화는 =ROUNDDOWN(달러×환율,0), 환율은 BL날짜로 환율 시트에서 찾음(BERGEN 1,350원 고정, ATLAS는 전자세금계산서 원화).')
  # ── 1. 합계 비교 ──
  $r = 5; Put $s $r @('1. 합계 비교'); $s.Cells($r,1).Font.Bold = $true; $s.Cells($r,1).Font.Size = 12; $r++
  Put $s $r @('항목', '인보이스별 상세', '라인별 매출', '10월 시트', '설명'); $s.Range("A$($r):E$r").Interior.Color = $HD; $s.Range("A$($r):E$r").Font.Bold = $true; $r1 = $r; $r++
  $sumOf = { param($rows, $col) ($rows | ? { $null -ne $_.$col } | Measure-Object $col -Sum).Sum }
  $prod = $cmp | ? { $_.구분 -in $kName['전월'], $kName['당월'] -and $_.거래처 -notlike 'JS TECH*' -and $_.거래처 -ne 'SHINKO' }
  $tool = $cmp | ? 구분 -eq $kName['툴링']; $duty = $cmp | ? 구분 -eq $kName['관세']; $carry = $cmp | ? 구분 -eq $kName['이월']; $prev = $cmp | ? { $_.구분 -eq $kName['전월'] -and $_.거래처 -notlike 'JS TECH*' }
  $lines = @(
    @('제품 매출 (9월 이월분 + 10월분, 툴링·관세 제외)', (& $sumOf $prod '원화상세'), (& $sumOf $prod '원화라인'), (& $sumOf $prod '작업원화'), '세 곳 모두 같음'),
    @('  그중 9월에서 넘어온 매출(녹색)', (& $sumOf $prev '원화상세'), (& $sumOf $prev '원화라인'), (& $sumOf $prev '작업원화'), '9월 이월 26건이 모두 10월에 들어옴(아래 4번 표)'),
    @('툴링비 (5건)', (& $sumOf $tool '원화상세'), (& $sumOf $tool '원화라인'), (& $sumOf $tool '작업원화'), '세 곳 모두 같음'),
    @('관세 (ALLISON)', (& $sumOf $duty '원화상세'), (& $sumOf $duty '원화라인'), (& $sumOf $duty '작업원화'), '관세 3건 환율 차이와 8월 오기재 차액분(10월 시트에서 뺌) 때문에 서로 다름'),
    @('11월로 이월 (회색)', $null, (& $sumOf $carry '원화라인'), (& $sumOf $carry '작업원화'), '이월은 라인별 매출의 이월 시트에만 있음. 매출에는 넣지 않고 이월 칸에만 넣음'),
    @('신코·코마츠 (엔화, 국내영업)', (& $sumOf ($cmp | ? { $_.거래처 -in 'SHINKO','JS TECH(코마츠)' }) '원화상세'), $null, (& $sumOf ($cmp | ? { $_.거래처 -in 'SHINKO','JS TECH(코마츠)' }) '작업원화'), '신코_지출품의서 마감 신코202610 시트 1건 + 코마츠_지출품의서 마감 코마츠202610 시트 4건 (왼쪽 칸 = 매출내역서 금액)')
  )
  foreach ($l in $lines) { Put $s $r $l; $r++ }
  Put $s $r @('10월 시트 당월매출 합계(TOTAL)', $null, $null, $ws.Range('J24').Value2, '제품 매출 + 툴링비 + 관세 + 신코·코마츠'); $s.Range("A$($r):E$r").Font.Bold = $true; $r++
  $s.Range("B$($r1):D$r").NumberFormat = '#,##0;-#,##0;'; $s.Range("A$($r1):E$($r-1)").Borders.LineStyle = 1; $s.Range("A$($r1):E$($r-1)").Borders.Color = 0xBFBFBF
  $r++
  # ── 2. 확인할 것 ──
  Put $s $r @('2. 확인할 것'); $s.Cells($r,1).Font.Bold = $true; $s.Cells($r,1).Font.Size = 12; $r++
  Put $s $r @('구분', '대상', '금액 영향', '내용', '할 일'); $s.Range("A$($r):E$r").Interior.Color = $HD; $s.Range("A$($r):E$r").Font.Bold = $true; $r2 = $r; $r++
  $dRate = $cmp | ? { $_.결과 -eq '에러' -and $_.인보이스 -notlike '*오기재*' }
  $dTxt = ($dRate | % { "$($_.인보이스) BL $($_.BL라인): 기초자료 $($_.환율상세.ToString('N2'))원 → BL일 환율 $($_.BL일환율.ToString('N2'))원" }) -join "`n"
  $cor = $cmp | ? { $_.인보이스 -like '*오기재*' }
  $am = $P9 | ? 인보이스 -eq 'DSM2609-AM04'; $am10 = $cmp | ? 인보이스 -eq 'DSM2609-AM04'
  $notes = @(
    @('에러', 'ALLISON 관세 3건 환율', ((& $sumOf $dRate '작업원화') - (& $sumOf $dRate '원화상세')), "두 기초자료 모두 이 3건만 환율이 수식이 아닌 숫자로 적혀 있고, BL날짜의 환율과 다릅니다. 숫자가 9월 시트의 같은 관세 줄 환율(1,355.20 / 1,360.00 / 1,360.00)과 똑같아 지난달 값이 남은 것으로 보입니다.`n$dTxt", '10월 시트는 BL날짜 환율로 계산해 두었습니다. 해외영업팀에 알려 관세 매출내역서를 고치게 해 주세요. 기초자료 금액을 그대로 쓰려면 10월 시트 해당 줄 환율 칸에 숫자를 직접 넣으면 됩니다.'),
    @('알림', '8월 오기재 관세 차액분 (DSM2608-AL05 관세)', -2885762, '인보이스별 상세의 ALLISON 관세 매출내역서에만 -2,885,762원이 있고 라인별 매출 당월 목록에는 없습니다. 9월 시트에서 DSM2608-AL05(관세)를 9월 3일 환율 1,369.40원으로 잡았는데 BL은 9월 13일(1,338.20원)이라 생긴 차액입니다($92,492.41 × 31.20원).', '10월 시트에서는 뺐습니다(요청). 그래서 관세 합계가 인보이스별 상세보다 2,885,762원 많습니다. 9월 시트를 고칠지는 따로 정해 주세요.'),
    @('알림', 'MEPPI (새 거래처)', 34115054, '툴링비 2건(DSM2610-ME01-TI 32,037,005원, DSM2610-ME01 2,078,049원, 마산2). 지난달 시트에 없는 거래처입니다. 라인별 매출 비고: "MEPPI 50% 입금".', '요약표 ALLISON 아래에 MEPPI 줄을 새로 넣고 상세에도 MEPPI 줄 2개를 넣었습니다. 다른 거래처 쪽으로 묶어야 하면 알려 주세요.'),
    @('알림', 'AMAT DSM2609-AM04 (9월 이월분)', ($am10.작업원화 - [double]$am.원화), "9월 시트 이월 때 `$$([double]$am.외화)였는데 10월에 `$$($am10.달러라인)로 올랐습니다(라인별 비고: 단가 변경으로 인해 일부 값 상승). 9월에는 면장번호가 비어 있었고 10월은 20576-26-100077X입니다.", '10월 금액으로 넣었습니다. 단가 변경이 맞는지만 확인해 주세요.'),
    @('알림', '환율 시트 10월 2일 ~ 11월 5일', $null, '이 파일의 환율 시트는 10월 1일 이후가 비어 있었습니다. 라인별 매출 파일의 환율 시트 값을 그대로 넣었습니다(35일).', '서울외국환중개 매매기준율과 맞는지 확인해 주세요. 바꾸면 10월 시트가 자동으로 다시 계산됩니다.'),
    @('알림', '유니패스 BL일자체크 칸', $null, '10월 시트의 유니패스 BL일자체크 칸은 비워 두었습니다(코마츠 9월 이월분은 9월 값 그대로).', '직접 채워 주세요. BL날짜가 달라지면 BL일자 칸을 고치면 환율·원화가 따라 바뀝니다.'),
    @('알림', '신코·코마츠 10월분', $null, '신코_지출품의서 마감·코마츠_지출품의서 마감 파일의 10월 시트(신코202610, 코마츠202610)만 넣었습니다. 엔화 환율 5건 모두 서울외국환중개와 같습니다. 면장번호는 이 파일들에 없어 비워 두었습니다. 코마츠 DSM2609-JS03(BL 10월 1일)은 9월 시트에서 회색 이월(9월 시트 표기 DSM2509-JS03)이던 건이라 녹색으로 넣었습니다.', '면장번호·유니패스 칸만 채워 주세요.'),
    @('참고', 'MERCURY 툴링 3건', 57449218, 'DSM2610-24(면장 41072-26-703075X, 비고: $7.00 무상샘플로 제외), DSM2610-25, DSM2610-27(면장 41072-26-703139X, 머큐리 오일쿨러). 모두 인보이스별 상세의 툴링비+PPAP 시트에 있어 툴링으로 넣었습니다.', "인보이스 뒤에 '(툴링비)'를 붙였고, 면장번호가 있으면 면장번호를, 없으면 '툴링비'를 적었습니다."),
    @('참고', '11월 이월 20건', (& $sumOf $carry '작업원화'), '라인별 매출 이월 시트의 날짜(대부분 11월 1일 = 마감일 기준 BL 미발행)와 그 날짜 환율로 계산했습니다. 이월 시트의 MERCURY CHINA 빈 줄(인보이스 없음, 0원)은 넣지 않았습니다.', '11월 마감 때 실제 BL날짜로 바뀝니다.'),
    @('참고', '라인별 매출 파일의 라인별 시트', $null, "'26.10月라인별' 시트의 출하·이월 표(위쪽 두 표)는 제목이 '3월'이고 숫자도 당월 목록과 맞지 않아 지난 값이 남은 것으로 보입니다. 맨 아래 '10월 매출-비율' 표는 당월 목록과 맞습니다.", '비교에는 쓰지 않았습니다.')
  )
  foreach ($n in $notes) { Put $s $r $n; $s.Range("A$($r):E$r").Interior.Color = $(switch ($n[0]) { '에러' { $RED } '알림' { $YEL } default { $GRY } }); $r++ }
  $s.Range("C$($r2):C$r").NumberFormat = '#,##0;-#,##0;'; $s.Range("D$($r2):E$r").WrapText = $true; $s.Range("A$($r2):E$($r-1)").VerticalAlignment = -4160
  $s.Range("A$($r2):E$($r-1)").Borders.LineStyle = 1; $s.Range("A$($r2):E$($r-1)").Borders.Color = 0xBFBFBF
  $r++
  # ── ALLISON 매출·관세 짝 맞춤 ──
  Put $s $r @('3. ALLISON 매출·관세 짝 맞춤 (인보이스별 상세 기준)'); $s.Cells($r,1).Font.Bold = $true; $s.Cells($r,1).Font.Size = 12; $r++
  Put $s $r @('인보이스', '매출 BL날짜', '매출 환율', '매출 외화', '매출 원화', '관세 BL날짜', '관세 환율 (기초자료)', '관세 외화', '관세 원화 (기초자료)', '관세/매출', '관세 원화 (10월 시트)', '결과'); $s.Range("A$($r):L$r").Interior.Color = $HD; $s.Range("A$($r):L$r").Font.Bold = $true; $s.Range("A$($r):L$r").WrapText = $true; $ra = $r; $r++
  foreach ($x in $SA | ? { $_.시트 -eq 'ALLISON' }) { $y = $SA | ? { $_.시트 -eq 'ALLISON 관세' -and $_.인보이스 -eq "$($x.인보이스)(관세)" } | Select -First 1; $z = $cmp | ? 인보이스 -eq "$($x.인보이스)(관세)" | Select -First 1
    $ok = if (-not $y) { '관세 없음' } elseif ([Math]::Abs((Num $x.환율) - (Num $y.환율)) -gt 0.001) { "관세 환율이 같은 날 매출 환율과 다름 → 10월 시트는 매출과 같은 $((Num $x.환율).ToString('N2'))원" } else { '맞음' }
    Put $s $r @($x.인보이스, $x.BL, (Num $x.환율), (Num $x.외화), (Num $x.원화), $y.BL, (Num $y.환율), (Num $y.외화), (Num $y.원화), $(if ($y) { (Num $y.외화) / (Num $x.외화) }), $z.작업원화, $ok)
    if ($ok -ne '맞음') { $s.Range("A$($r):L$r").Interior.Color = $RED }; $r++ }
  $s.Range("C$($ra):C$r").NumberFormat = '#,##0.00'; $s.Range("G$($ra):G$r").NumberFormat = '#,##0.00'; $s.Range("D$($ra):D$r").NumberFormat = '#,##0.00'; $s.Range("H$($ra):H$r").NumberFormat = '#,##0.00'
  $s.Range("E$($ra):E$r").NumberFormat = '#,##0'; $s.Range("I$($ra):I$r").NumberFormat = '#,##0'; $s.Range("K$($ra):K$r").NumberFormat = '#,##0'; $s.Range("J$($ra):J$r").NumberFormat = '0.00%'
  $s.Range("A$($ra):L$($r-1)").Borders.LineStyle = 1; $s.Range("A$($ra):L$($r-1)").Borders.Color = 0xBFBFBF
  $r++  # ── 3. 9월 이월 → 10월 ──
  Put $s $r @('4. 9월 시트 이월(회색) → 10월 매출(녹색) 대조'); $s.Cells($r,1).Font.Bold = $true; $s.Cells($r,1).Font.Size = 12; $r++
  Put $s $r @('거래처', '인보이스', '9월 면장번호', '10월 면장번호', '9월 외화', '10월 외화', '9월 이월 원화', '10월 원화', '결과'); $s.Range("A$($r):I$r").Interior.Color = $HD; $s.Range("A$($r):I$r").Font.Bold = $true; $r3 = $r; $r++
  foreach ($p in $P9) { $t = $cmp | ? 인보이스 -eq $p.인보이스 | Select -First 1; $b = $LB | ? 인보이스 -eq $p.인보이스 | Select -First 1
    $m10 = if ($b) { $b.면장 } else { $p.면장 }; $u10 = if ($b) { [double]$b.외화 } else { [double]$p.외화 }
    $res = if ($p.인보이스 -eq 'DSM2509-JS03') { '10월 코마츠 매출내역서에 DSM2609-JS03으로 들어옴 (BL 10월 1일, 8.6128) — 녹색' } elseif (-not $t) { '10월에 없음' } elseif ([Math]::Abs($u10 - [double]$p.외화) -gt 0.005) { '외화 금액 바뀜' } elseif ($m10 -ne $p.면장) { '면장번호 바뀜' } else { '맞음 (원화는 10월 BL일 환율로 다시 계산)' }
    Put $s $r @($p.거래처, $p.인보이스, $p.면장, $m10, [double]$p.외화, $u10, [double]$p.원화, $t.작업원화, $res)
    if ($res -notlike '맞음*') { $s.Range("A$($r):I$r").Interior.Color = $YEL }; $r++ }
  $s.Range("E$($r3):F$r").NumberFormat = '#,##0.00'; $s.Range("G$($r3):H$r").NumberFormat = '#,##0'
  $s.Range("A$($r3):I$($r-1)").Borders.LineStyle = 1; $s.Range("A$($r3):I$($r-1)").Borders.Color = 0xBFBFBF
  $r++
  # ── 4. 줄별 비교 ──
  Put $s $r @('5. 줄별 비교 (10월 시트 순서)'); $s.Cells($r,1).Font.Bold = $true; $s.Cells($r,1).Font.Size = 12; $r++
  Put $s $r @('10월 시트 행', '거래처', '구분', '인보이스', 'BL날짜 (상세)', 'BL날짜 (라인별)', '환율 (상세)', '환율 (라인별)', 'BL일 환율 (환율 시트)', '외화 (상세)', '외화 (라인별)', '원화 (상세)', '원화 (라인별)', '원화 (10월 시트)', '결과', '내용')
  $s.Range("A$($r):P$r").Interior.Color = $HD; $s.Range("A$($r):P$r").Font.Bold = $true; $s.Range("A$($r):P$r").WrapText = $true; $r4 = $r; $r++
  foreach ($c in $cmp) { Put $s $r @("$($c.행)행", $c.거래처, $c.구분, $c.인보이스, $c.BL상세, $c.BL라인, $c.환율상세, $c.환율라인, $c.BL일환율, $c.달러상세, $c.달러라인, $c.원화상세, $c.원화라인, $c.작업원화, $c.결과, $c.내용)
    $col = switch ($c.결과) { '에러' { $RED } '확인' { $YEL } '참고' { $GRY } default { $null } }; if ($col) { $s.Range("A$($r):P$r").Interior.Color = $col }; $r++ }
  foreach ($m in $miss) { Put $s $r @('없음', $m.거래처, '', $m.인보이스, $m.BL, '', (Num $m.환율), $null, $null, (Num $m.외화), $null, (Num $m.원화), $null, $null, '확인', '인보이스별 상세에 있는데 10월 시트에 없음'); $s.Range("A$($r):P$r").Interior.Color = $YEL; $r++ }
  $s.Range("G$($r4):I$r").NumberFormat = '#,##0.00'; $s.Range("J$($r4):K$r").NumberFormat = '#,##0.00'; $s.Range("L$($r4):N$r").NumberFormat = '#,##0;-#,##0'
  $s.Range("A$($r4):P$($r-1)").Borders.LineStyle = 1; $s.Range("A$($r4):P$($r-1)").Borders.Color = 0xBFBFBF
  $s.Range("A$($r4):P$($r-1)").AutoFilter() | Out-Null
  # 열 너비
  $wd = @{ A = 44; B = 26; C = 18; D = 60; E = 48; F = 14; G = 14; H = 14; I = 22; J = 13; K = 13; L = 14; M = 14; N = 14; O = 7; P = 70 }
  foreach ($k in $wd.Keys) { $s.Columns($k).ColumnWidth = $wd[$k] }
  $s.Range("P$($r4):P$r").WrapText = $true
  $s.Activate(); $xl.ActiveWindow.Zoom = 90; $s.Range('A1').Select() | Out-Null
  $ws.Activate(); $ws.Range('A1').Select() | Out-Null
  $wb.Save(); $wb.Close($false)
  "검증 시트 저장. 줄별 결과: " + (($cmp | Group-Object 결과 | % { "$($_.Name) $($_.Count)" }) -join ', ') + " / 상세에만 있는 것 $(@($miss).Count)건"
  $cmp | ? { $_.결과 -ne '맞음' } | % { "$($_.행) $($_.인보이스) [$($_.결과)] $($_.내용)" }
} finally { $xl.ScreenUpdating = $true; $xl.Quit() }
