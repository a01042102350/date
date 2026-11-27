# MES매칭 통합 ↔ 내 월마감(1~10월) 대조 결과 엑셀 만들기
$ErrorActionPreference = 'Stop'
$W = 'C:\Users\user\AppData\Local\Temp\claude\d3df15a30ffc05d5\ac480345-177e-4ecb-8db7-e819891326b4\scratchpad\mesmatch'
. (Join-Path $W 'common.ps1')
$OUT = 'C:\ai\매월마감\MES매칭 통합 ↔ 내 월마감 대조(2026-01~10).xlsx'
function F($d) { '{0:N0}' -f [decimal]$d }

# ---------- 거래처×월 판단 ----------
$cmp = @(Import-Csv (Join-Path $W 'cmp.csv') -Encoding UTF8)
$KNOWN = @{
  'SAM|2'       = @('제 자료 고칠 것', 'DSM2601-SM01 57,200,724원이 제 2월 해외매출에 없음(1월에서 2월로 이월만 하고 2월에 안 잡힘) → 2월 회계 파일 26.02(NEW)에 있는지 확인')
  'EMERSON|4'   = @('제 자료 고칠 것', 'DSM2603-EM04 1,004,681원이 제 쪽 4월에 또 잡힘(3월 988,729원과 같은 면장 41072-26-700108X, 중복 의심) + 그쪽 기타매출 4,797,325원')
  '맥시스|3'     = @('제 자료 고칠 것', '3월 출고분인데 실제 매출(계산서)은 4월 69,300,000원·5월 22,000,000원. 제 쪽은 3월에 잡고 4월에 69,300,000원을 또 잡음 → 3월 일별마감 이월 표시 확인')
  '영창이엔지|3'  = @('제 자료 고칠 것', '3월 출고분인데 계산서는 4월(4,810,000원). 제 쪽은 3월에 잡고 4월에도 잡힘 → 3월 일별마감 이월 표시 확인')
  '현대로템|1'   = @('확인 필요', '방산: 일별마감 1월 163,964,790원 중 MES에 없는 140,621,090원을 제 도구가 단가 소급으로 봄 + 반품 -18,502,776원. 1월 세금계산서 6매 449,304,115원과는 같지만 장부 1월은 327,185,801원 → 122,118,313원이 장부 1월에 없음(작년 귀속 등 확인)')
  '현대로템|3'   = @('확인 필요', '방산 반품 -18,502,775원이 세금계산서(수정)에는 있는데 장부엔 없음 → 장부가 18,502,775원 많음')
  '현대로템|8'   = @('확인 필요', '방산 반품 -18,502,775원이 세금계산서(수정)에는 있는데 장부엔 없음(8월 패턴비 238,500,000원은 양쪽 다 있음)')
  'LS엠트론|6'   = @('기준 차이', '단가소급 -2,166,183원: 제 쪽은 일별마감(41,574,519원, 11/26 "일별마감이 맞음"), 장부·그쪽은 계산서(39,408,336원) → 어느 쪽으로 볼지 결정')
  '명화공업|6'   = @('기준 차이', '6월 명화 이월 건 24,967,800원(D1): 제 쪽 일별마감 vs 장부·계산서')
  '벽진|2'      = @('기준 차이', '2월 벽진 계산서 차이 5,369,600원: 제 쪽 일별마감 vs 장부·계산서')
  '피에스텍|1'   = @('기준 차이', '1월 피에스텍 11,579,400원을 2·3월에 나눠 청구: 제 쪽 일별마감 vs 장부·계산서')
  'STX엔진|1'   = @('기준 차이', '장부에만 있는 전년도(25년) 수정 420,000,000원(STX엔진 조회서 회신 반영) — 영업·MES 자료 없음')
  'MEPPI|10'    = @('기준 차이', '그쪽 기타매출 34,115,054원만 있음 — 제 해외매출엔 없음')
  '신코|2'      = @('기준 차이', '금형중자 수정비 3,979,448원(DSM2602-SK01-01, 그쪽 기타매출) — 제 해외매출(회계 면장)에 없음')
  '신코|5'      = @('기준 차이', '선급검사비 6건 17,216,412원(그쪽 기타매출) — 제 해외매출(회계 면장)에 없음')
  '신코|8'      = @('기준 차이', '금형비 7,522,284원 + 선급검사비 10건(그쪽 기타매출) — 제 해외매출(회계 면장)에 없음')
}
$rowsC = foreach ($c in $cmp) {
  $k = "$($c.거래처)|$($c.월)"; $d = [decimal]$c.차이_내금액_원장; $dp = [decimal]$c.차이_내금액_그쪽제품; $oth = [decimal]$c.그쪽기타; $mon = [int]$c.월
  $cls = ''; $why = ''
  if ([math]::Abs($d) -le 10) { $cls = '같음'; $why = '' }
  elseif ($KNOWN.ContainsKey($k)) { $cls = $KNOWN[$k][0]; $why = $KNOWN[$k][1] }
  elseif ($c.거래처 -eq '아트라스콥코' -and $mon -le 5) { $cls = '제 자료 고칠 것'; $why = '제 1~5월 국내매출 표에 아트라스콥코가 빠짐(전자세금계산서 금액 오류에만 잡힘) — 6~10월은 같음. 도구에서 1~5월도 국내매출에 넣게 고쳐야 함' }
  elseif ($c.거래처 -eq '신우텍') { $cls = '기준 차이'; $why = '제품 외 매출(스크랩·AL재 판매) — 제 국내매출 비교 대상이 아님' }
  elseif ($c.거래처 -eq '알멕') { $cls = '기준 차이'; $why = '외주가공비 등 기타매출 — 제 국내매출은 제품 매출만' }
  elseif ($c.부서 -eq '해외영업' -and [math]::Abs($dp) -le 10 -and $oth -ne 0) { $cls = '기준 차이'; $why = "그쪽 기타매출(관세·툴링·PPAP·선급비 등) $(F $oth)원 — 제 해외매출(회계 면장)엔 없음" }
  elseif ($c.부서 -eq '해외영업' -and $dp -gt 0 -and $d -lt 0) { $cls = '기준 차이'; $why = "그쪽 기타매출 중 $(F $dp)원(PPAP·툴링·불량청구)은 제 쪽 해외매출에 들어 있고, 나머지 $(F (-$d))원은 제 쪽에 없음" }
  else { $cls = '확인 필요'; $why = '' }
  [pscustomobject][ordered]@{ 판단 = $cls; 부서 = $c.부서; 거래처 = $c.거래처; 월 = $mon; '내 이름' = $c.내이름; '그쪽 제품매출' = [decimal]$c.그쪽제품; '그쪽 기타매출' = $oth; '그쪽 합계' = [decimal]$c.그쪽합계; '계정별원장' = [decimal]$c.원장
    '내 구분' = $c.내구분; '내 금액' = [decimal]$c.내금액; '내 세금계산서' = $(if ($c.내계산서 -ne '') { [decimal]$c.내계산서 } else { $null }); '차이(내-원장)' = $d; '차이(내-그쪽 제품)' = $dp; '이유' = $why }
}

# ---------- 해외 인보이스 판단 ----------
function IK([string]$s) { ($s -replace '\s|\(.*?\)', '').ToUpper() }
$L = Get-Content (Join-Path $W 's4.tsv') -Encoding UTF8; $h = $L[3] -split "`t"; $ix = @{}; for ($i = 0; $i -lt $h.Count; $i++) { if (-not $ix.ContainsKey($h[$i])) { $ix[$h[$i]] = $i } }
$def = @{}; foreach ($line in $L | Select-Object -Skip 4) { $c = $line -split "`t"; if ($c[$ix['원장 반영']] -eq 'O' -and $c[$ix['매출구분']] -ne '제품매출') { $kk = IK $c[$ix['인보이스']]; if ($kk) { $def[$kk] = [decimal]$def[$kk] + (N $c[$ix['원화 계']]) } } }
$L5 = Get-Content (Join-Path $W 's5.tsv') -Encoding UTF8; $h5 = $L5[3] -split "`t"; $i5 = @{}; for ($i = 0; $i -lt $h5.Count; $i++) { if (-not $i5.ContainsKey($h5[$i])) { $i5[$h5[$i]] = $i } }
$oth5 = @{}; foreach ($line in $L5 | Select-Object -Skip 4) { $c = $line -split "`t"; $kk = IK $c[$i5['인보이스']]; if ($kk -and $c[$i5['원장 반영']] -like 'O*') { $oth5[$kk] = [decimal]$oth5[$kk] + (N $c[$i5['원화']]) } }
$inv = @(Import-Csv (Join-Path $W 'inv_cmp.csv') -Encoding UTF8)
$PAIR = @{ 'DSM2608-AL1' = 'DSM2608-AL01'; 'DSM2608-AL01' = 'DSM2608-AL1'; 'DSM2508-JS01' = 'DSM2608-JS01'; 'DSM2608-JS01' = 'DSM2508-JS01'; 'DSM2508-JS02' = 'DSM2608-JS02'; 'DSM2608-JS02' = 'DSM2508-JS02'; 'DSM2509-JS01' = 'DSM2609-JS01'; 'DSM2609-JS01' = 'DSM2509-JS01'; 'DSM2509-JS02' = 'DSM2609-JS02'; 'DSM2609-JS02' = 'DSM2509-JS02' }
$rowsI = foreach ($r in $inv) {
  $k = $r.인보이스; $tp = [decimal]$r.그쪽제품; $mv = [decimal]$r.내금액; $dd = [decimal]$def[$k]; $oo = [decimal]$oth5[$k]; $cls = ''; $why = ''
  switch ($r.상태) {
    '같음' { $cls = '같음' }
    default {
      if ($PAIR.ContainsKey($k)) { $cls = '기준 차이'; $why = "인보이스 번호 표기만 다름(짝: $($PAIR[$k]), 금액·월 같음). 제 쪽은 회계 파일·MES 번호를 고친 번호" }
      elseif ($k -eq 'DSM2601-SM01') { $cls = '제 자료 고칠 것'; $why = '1월에서 2월로 이월했는데 제 2월 해외매출에 없음 → 2월 회계 파일 확인' }
      elseif ($k -eq 'DSM2603-EM04') { $cls = '제 자료 고칠 것'; $why = '제 쪽 3월(988,729)·4월(1,004,681) 두 번 잡힘, 같은 면장 41072-26-700108X → 4월 회계 파일 확인' }
      elseif ($k -in 'DSM2604-SK02', 'DSM2604-SK03') { $cls = '확인 필요'; $why = '그쪽은 MES 출고Lot SK02(92개, MES 160,814,524)를 영업 인보이스 SK03과, SK03(16개)을 SK02와 짝지음. 회계 면장(700274X=SK02 ¥17,245,240)·MES 기준으로는 제 쪽이 맞음 → 영업 매출내역서 번호 확인' }
      elseif ($k -eq 'DSM2506-AT01') { $cls = '확인 필요'; $why = '작년(25년) 누락분 — 그쪽은 5월 매출로 넣음, 제 쪽은 아트라스를 인보이스로 보지 않음(5월은 아예 빠짐)' }
      elseif ($r.거래처 -eq '아트라스콥코') { $m0 = [int]($r.그쪽월); if ($m0 -le 5) { $cls = '제 자료 고칠 것'; $why = '제 쪽은 아트라스를 국내매출(고객사 합계)로 보는데 1~5월은 국내매출 표에서 빠짐' } else { $cls = '기준 차이'; $why = '제 쪽은 아트라스를 국내매출(고객사 합계)로 봄 — 그 달 고객사 합계는 같음' } }
      elseif ($r.상태 -eq '내 쪽만' -and $mv -eq 0) { $cls = '기준 차이'; $why = '제 쪽 0원 줄' }
      elseif ($r.상태 -eq '내 쪽만' -and $oo -ne 0 -and [math]::Abs($mv - $oo) -le 10) { $cls = '기준 차이'; $why = "그쪽은 기타매출(PPAP·샘플·기타)로 분류 — 금액 같음($(F $oo))" }
      elseif ($r.상태 -eq '내 쪽만' -and $oo -ne 0) { $cls = '기준 차이'; $why = "그쪽은 기타매출(PPAP·샘플 등)로 분류 — 그쪽 기타 $(F $oo)원" }
      elseif ($r.상태 -eq '금액 다름' -and [math]::Abs($mv - ($tp + $dd)) -le 10) { $cls = '기준 차이'; $why = "그쪽은 같은 인보이스 안에서 불량대체-청구 $(F $dd)원을 따로 나눔(합치면 같음)" }
      elseif ($r.상태 -eq '금액 다름' -and [math]::Abs($mv - ($tp + $dd + $oo)) -le 10) { $cls = '기준 차이'; $why = "그쪽은 일부($(F ($dd + $oo))원)를 기타매출(PPAP·샘플 등)로 따로 나눔(합치면 같음)" }
      else { $cls = '확인 필요'; $why = '' }
    }
  }
  [pscustomobject][ordered]@{ 판단 = $cls; 상태 = $r.상태; 인보이스 = $k; 거래처 = $r.거래처; '그쪽 매출월' = $r.그쪽월; '내 매출월' = $r.내월; '그쪽 제품매출' = $tp; '그쪽 불량청구 등' = $dd; '그쪽 기타매출' = $oo; '내 금액(회계 원화)' = $mv; '차이(내-그쪽 제품)' = [decimal]$r.차이; '내 이월 표시' = $r.내이월; 이유 = $why }
}

# ---------- 라인별 ----------
$rowsL = @(Import-Csv (Join-Path $W 'line_cmp.csv') -Encoding UTF8 | ForEach-Object { $o = [ordered]@{}; foreach ($p in $_.PSObject.Properties) { $o[$p.Name] = $(if ($p.Name -match '^(그쪽|내) ' ) { [decimal]$p.Value } elseif ($p.Name -eq '월') { [int]$p.Value } else { $p.Value }) }; [pscustomobject]$o })

# ---------- 이월(10월말) ----------
$carryTh = @{}; foreach ($line in $L | Select-Object -Skip 4) { $c = $line -split "`t"; if ($c[$ix['매출월']] -eq '이월') { $carryTh[$c[$ix['부서']]] = [decimal]$carryTh[$c[$ix['부서']]] + (N $c[$ix['원화 계']]) } }
$mineRows = Import-Csv (Join-Path $W 'mine_rows.csv') -Encoding UTF8
$carryMy = @{}; foreach ($r in $mineRows | Where-Object { $_.출처 -eq '이월매출' -and $_.월 -eq '10' }) { $carryMy[$r.구분] = [decimal]$carryMy[$r.구분] + [decimal]$r.금액 }

# ---------- 숫자 모음 ----------
function Cnt($rows, $cls) { @($rows | Where-Object 판단 -eq $cls).Count }
$nC = $rowsC.Count; $nI = $rowsI.Count; $sameL = @($rowsL | Where-Object 합계같음 -eq 'O'); $sameLok = @($sameL | Where-Object { -not $_.'라인 차이' }).Count
$othTot = ($rowsC | Where-Object { $_.부서 -eq '해외영업' } | Measure-Object '그쪽 기타매출' -Sum).Sum
$allisonOth = ($rowsC | Where-Object 거래처 -eq 'ALLISON' | Measure-Object '그쪽 기타매출' -Sum).Sum
$atlasMiss = ($rowsC | Where-Object { $_.거래처 -eq '아트라스콥코' -and $_.월 -le 5 } | Measure-Object '계정별원장' -Sum).Sum
$shinwoo = ($rowsC | Where-Object 거래처 -eq '신우텍' | Measure-Object '계정별원장' -Sum).Sum
$ledger = ($rowsC | Measure-Object '계정별원장' -Sum).Sum; $mytot = ($rowsC | Measure-Object '내 금액' -Sum).Sum

$SUM = New-Object Collections.ArrayList
function S([string]$a, [string]$b = '', [string]$c = '', [string]$d = '') { [void]$SUM.Add(@($a, $b, $c, $d)) }
S '국내·해외영업 MES매칭 통합(2026-01~10) ↔ 내 월마감 결과(기초자료 1~10월) 대조'
S "만든 때 $(Get-Date -Format 'yyyy-MM-dd HH:mm') · 그쪽 = '국내·해외영업 MES매칭 통합(2026-01~10).xlsx' · 내 것 = C:\ai\기초자료\YYYYMM\ekfmsrj.html(해외매출·국내매출·이월매출)"
S ''
S '1. 무엇을 비교했나' '그쪽 시트' '내 자료' '비교 방법'
S '거래처×월 매출 금액' '거래처별 월별 원장 대사 (제품매출·기타매출·계정별원장)' '해외매출 원화(회계, 면장별) + 국내매출 당월 매출(일별마감)·세금계산서' '거래처 이름 짝을 맞춰 월마다 합계 비교. 기준 = 계정별원장(장부)'
S '해외 인보이스별 매출월·금액' 'MES 매칭(통합) 인보이스·매출월·원화 계 + 대상외(기타매출)' '해외매출 인보이스 번호·월·원화(회계)' '인보이스 번호로 짝지어 매출월·금액 비교'
S '라인별 금액(사형·자동·금형·가공1~3)' '프로젝트별 매출' '해외매출·국내매출 라인 칸' '거래처×월 합계가 같은 칸만 라인별 비교'
S '10월말 이월' "MES 매칭(통합) 매출월 '이월'" '10월 이월매출' '부서(구분)별 합계'
S ''
S '비교할 수 없는 것' '그쪽에만' '내 쪽에만'
S '' '계정별원장(장부) 대사, 기타매출 분류(관세·불량청구·단가소급·선급검사비·금형비·PPAP·스크랩·알멕), 거래처 품목별 매출, 국내 MES 행별 매칭(일별마감 수작업), 대상외(MES 무관) 598줄' '수출면장 번호·BL날짜·환율(smbs) 검증, 라인 파일과 비율 비교, 세금계산서 일련번호, 오류목록(에러·참고), 출하금액변동'
S ''
S '2. 결과 숫자' '건수' '같음' '설명'
S '거래처×월' "$nC 칸" "$(Cnt $rowsC '같음') 칸" "1~10월 장부 합계 $(F $ledger)원 / 내 금액 합계 $(F $mytot)원 (차이 $(F ($mytot - $ledger))원 — 대부분 아래 '기준 차이')"
S '해외 인보이스' "$nI 건" "$(Cnt $rowsI '같음') 건" '나머지 대부분은 번호 표기·분류 차이(금액은 같음)'
S '라인별' "합계 같은 $($sameL.Count) 칸" "$sameLok 칸" '라인 나누는 기준이 다름: 내 것 = MES 금액(내부 고정환율 약 1,350원)을 나눔 / 그쪽 = 확정 원화를 MES 주조비·가공비 비율로 나눔'
S '10월말 이월' "그쪽 국내 $(F $carryTh['국내영업'])원 · 해외 $(F $carryTh['해외영업'])원" "내 국내 $(F $carryMy['국내'])원 · 해외 $(F $carryMy['해외'])원" '해외는 같음(1원). 국내 차이 193,090,760원은 전부 신코 DSM2610-SK02(아래 ⑤). 그쪽 이월 = 10월까지 출하·11월 이후 매출'
S ''
S '3. 판단별 건수' '거래처×월' '해외 인보이스' ''
foreach ($cl in '제 자료 고칠 것', '확인 필요', '기준 차이', '같음') { S $cl "$(Cnt $rowsC $cl)" "$(Cnt $rowsI $cl)" '' }
S ''
S '4. 차이 목록' '금액' '근거' '할 일'
S '[제 자료 고칠 것]'
S '① 아트라스콥코 1~5월이 제 국내매출에서 빠짐' "$(F $atlasMiss)원 (장부 1~5월)" '1~5월은 전자세금계산서 금액 오류에만 잡히고 국내매출 표에 없음. 6~10월은 장부와 같음' '도구에서 아트라스(국·해외) 규칙을 1~5월에도 적용 → 1~5월 다시 만들기'
S '② SAM DSM2601-SM01 2월 누락' '57,200,724원' '1월 이월매출(BL 2/5)로 넘겼는데 2월 해외매출에 없음. 장부 2월 SAM 85,972,806 = 28,772,082 + 57,200,724' '2월 회계 파일 26.02(NEW)에 SM01이 있는지 확인'
S '③ EMERSON DSM2603-EM04 두 번' '1,004,681원(4월)' '3월(BL 3/27, 988,729)과 4월(BL 4/1, 1,004,681) 둘 다 해외매출, 면장 41072-26-700108X 같음. 그쪽은 3월만' '4월 회계 파일 확인 → 중복이면 4월에서 빼기'
S '④ 맥시스·영창이엔지 3월' '91,300,000원 · 4,810,000원' '3월 출고분인데 계산서는 4·5월. 제 3월 매출에 잡고, 맥시스 69,300,000원은 4월에 또 잡힘. 장부 3월 0원' '3월 일별마감에 4월이월 표시가 있는지 확인 → 3월에서 빼기'
S '⑤ 신코 10월말 이월이 이월매출 표에 없음' '193,090,760원(그쪽) / MES 195,842,853원' 'DSM2610-SK02(10/30 출고, 97개)는 제 MES 시트에 BL=이월로 표시되지만, 10월 회계 파일에 아직 없어서 회계 기준인 이월매출 표에서 빠짐. 그쪽은 11월 이월로 잡음(국내 이월 차이 전부)' '도구: MES에만 있는 이월(회계 없음)도 이월매출 표에 넣기'
S '[확인 필요]'
S '⑥ 현대로템 1월' '+122,118,313원' '방산 일별마감 1월 163,964,790원 중 MES에 없는 140,621,090원을 도구가 단가 소급으로 봄(반품 -18,502,776원). 1월 세금계산서 449,304,115원과는 같지만 장부 1월은 327,185,801원' '장부에서 1월 계산서 일부가 작년(25년)으로 들어갔는지 확인'
S '⑦ 현대로템 3·8월 방산 반품' '-18,502,775원 × 2' '세금계산서(수정)와 일별마감에는 반품이 있는데 장부엔 없음 → 장부가 제 것보다 18,502,775원씩 많음' '회계에 반품 수정세금계산서 반영 여부 확인'
S '⑧ 신코 4월 SK02·SK03 짝' '100,823,257원씩 엇갈림' '그쪽은 MES 출고Lot SK02(92개)를 인보이스 SK03과 짝지음. 회계 면장 700274X = SK02 ¥17,245,240 = 159,661,605원 → 제 쪽이 MES·회계와 맞음' '그쪽 파일(영업 매출내역서 번호) 확인'
S '⑨ 아트라스 DSM2506-AT01' '21,526,400원' '작년(25년) 누락분 — 그쪽은 5월 매출' '5월 매출로 볼지 확인'
S '[기준·분류 차이 — 틀린 것은 아님]'
S '⑩ 해외 기타매출이 제 해외매출엔 없음' "$(F $othTot)원 (해외영업 1~10월, 그중 ALLISON $(F $allisonOth)원)" '관세(DDP)·툴링·PPAP·선급비 등은 장부엔 있지만 회계 면장 기준인 제 해외매출엔 없음. 신코 금형비·선급검사비(2·5·8월)도 같음' '그쪽처럼 기타매출까지 볼지 결정'
S '⑪ 분류만 다름(금액 같음)' '' 'Allison TA 인보이스·Mercury PPAP·샘플(DSM2606-T01, 2607-22, 2610-24 등)은 그쪽 = 기타매출, 제 쪽 = 해외매출. 신코 불량대체-청구는 그쪽이 같은 인보이스 안에서 따로 나눔' '분류 기준 통일 여부 결정'
S '⑫ 제품 외 매출' "신우텍 $(F $shinwoo)원 등" '신우텍(스크랩·AL재), 알멕 외주가공비, STX 1월 전년도(25년) 수정 420,000,000원(장부에만) — 제 국내매출 비교 대상이 아님' '그대로 두거나 기타매출로 따로 보기'
S '⑬ 국내 일별마감 vs 계산서' 'LS엠트론 6월 -2,166,183 / 명화 6월 -24,967,800 / 벽진 2월 +5,369,600 / 피에스텍 1월 +11,579,400' '제 국내매출 = 일별마감 기준, 장부·그쪽 = 계산서 기준. 이미 알던 건(D1 등)' '국내 매출 기준을 계산서(장부)로 바꿀지 결정'
S '⑭ 번호 표기' '' 'DSM2608-AL1 ↔ AL01, 코마츠 DSM2508·2509-JS01·02 ↔ DSM2608·2609-JS01·02 (제가 MES·회계 번호를 고친 것)' '없음'
S '⑮ 부서 구분' '' '그쪽: 신코·코마츠 = 국내영업, 아트라스 = 해외영업 / 제 쪽: 신코·코마츠 = 해외매출, 아트라스 = 국내매출' '없음(거래처로 짝지음)'
S '⑯ 라인별 금액 기준' "합계 같은 $($sameL.Count)칸 중 $($sameLok)칸만 같음" '제 라인 칸 = MES 금액(내부 고정환율)을 나눈 값 / 그쪽 = 확정 원화를 MES 주조비·가공비 비율로 나눈 값' '그쪽 기준으로 맞추려면 도구 수정 필요'

# ---------- 엑셀 쓰기 ----------
function To2D($objs, [string[]]$cols) { $n = @($objs).Count; $a = New-Object 'object[,]' ($n + 1), $cols.Count; for ($j = 0; $j -lt $cols.Count; $j++) { $a[0, $j] = $cols[$j] }; $i = 1; foreach ($o in $objs) { for ($j = 0; $j -lt $cols.Count; $j++) { $v = $o.($cols[$j]); if ($v -is [decimal]) { $v = [double]$v }; $a[$i, $j] = $v }; $i++ }; return , $a }
$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false
$mine = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
try {
  $wb = $xl.Workbooks.Add(); while ($wb.Worksheets.Count -lt 4) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  $sheets = @(
    @{ N = '요약'; Rows = $null },
    @{ N = '거래처×월 대조'; Rows = $rowsC; Money = @('그쪽 제품매출', '그쪽 기타매출', '그쪽 합계', '계정별원장', '내 금액', '내 세금계산서', '차이(내-원장)', '차이(내-그쪽 제품)') },
    @{ N = '해외 인보이스 대조'; Rows = ($rowsI | Sort-Object { @{ '제 자료 고칠 것' = 0; '확인 필요' = 1; '기준 차이' = 2; '같음' = 3 }[$_.판단] }, 거래처, 인보이스); Money = @('그쪽 제품매출', '그쪽 불량청구 등', '그쪽 기타매출', '내 금액(회계 원화)', '차이(내-그쪽 제품)') },
    @{ N = '라인별 대조'; Rows = $rowsL; Money = @() }
  )
  for ($s = 0; $s -lt 4; $s++) {
    $ws = $wb.Worksheets.Item($s + 1); $ws.Name = $sheets[$s].N; $ws.Cells.Font.Name = '맑은 고딕'; $ws.Cells.Font.Size = 10
    if ($s -eq 0) {
      $a = New-Object 'object[,]' $SUM.Count, 4; for ($i = 0; $i -lt $SUM.Count; $i++) { for ($j = 0; $j -lt 4; $j++) { $a[$i, $j] = $SUM[$i][$j] } }
      $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($SUM.Count, 4)).Value2 = $a
      $ws.Columns.Item(1).ColumnWidth = 34; $ws.Columns.Item(2).ColumnWidth = 40; $ws.Columns.Item(3).ColumnWidth = 70; $ws.Columns.Item(4).ColumnWidth = 46
      $ws.Range('A:D').WrapText = $true; $ws.Range('A:D').VerticalAlignment = -4160
      $ws.Range('A1').Font.Size = 14; $ws.Range('A1').Font.Bold = $true; $ws.Range('A1:D1').WrapText = $false; $ws.Range('A2:D2').WrapText = $false; $ws.Range('A2').Font.Color = 0x595959
      for ($i = 1; $i -le $SUM.Count; $i++) { $t = "$($SUM[$i - 1][0])"; if ($t -match '^\d\. ' -or $t -eq '비교할 수 없는 것') { $rg = $ws.Range($ws.Cells.Item($i, 1), $ws.Cells.Item($i, 4)); $rg.Font.Bold = $true; $rg.Interior.Color = 0xF2E1D9 }; if ($t -match '^\[') { $ws.Cells.Item($i, 1).Font.Bold = $true; $ws.Cells.Item($i, 1).Font.Color = $(if ($t -match '고칠') { 0x1D23C0 } elseif ($t -match '확인') { 0x0063C6 } else { 0x595959 }) } }
      continue
    }
    $rows = @($sheets[$s].Rows); $cols = @($rows[0].PSObject.Properties | ForEach-Object Name)
    $a = To2D $rows $cols; $rg = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($rows.Count + 1, $cols.Count)); $rg.Value2 = $a
    $hd = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(1, $cols.Count)); $hd.Font.Bold = $true; $hd.Interior.Color = 0x3E3D0F; $hd.Font.Color = 0xFFFFFF; $hd.WrapText = $true; $hd.VerticalAlignment = -4108
    for ($j = 0; $j -lt $cols.Count; $j++) { $cn = $cols[$j]; $col = $ws.Columns.Item($j + 1)
      if ($sheets[$s].N -eq '라인별 대조') { if ($cn -match '^(그쪽|내) ') { $col.NumberFormat = '#,##0'; $col.ColumnWidth = 13 } elseif ($cn -eq '라인 차이') { $col.ColumnWidth = 60 } else { $col.ColumnWidth = 10 } }
      elseif ($sheets[$s].Money -contains $cn) { $col.NumberFormat = '#,##0'; $col.ColumnWidth = 15 }
      elseif ($cn -eq '이유') { $col.ColumnWidth = 80; $col.WrapText = $true }
      elseif ($cn -in '판단', '상태') { $col.ColumnWidth = 14 } elseif ($cn -eq '내 이름') { $col.ColumnWidth = 26 } else { $col.ColumnWidth = 12 } }
    $ws.Rows.Item(1).RowHeight = 30
    # 판단 색
    if ($cols[0] -eq '판단') { for ($i = 0; $i -lt $rows.Count; $i++) { $p = $rows[$i].판단; $clr = switch ($p) { '제 자료 고칠 것' { 0xCEC7FF } '확인 필요' { 0x9CEBFF } '기준 차이' { 0xF2F2F2 } default { $null } }; if ($clr) { $ws.Cells.Item($i + 2, 1).Interior.Color = $clr } } }
    $rg.AutoFilter() | Out-Null; $ws.Activate(); $xl.ActiveWindow.SplitRow = 1; $xl.ActiveWindow.FreezePanes = $true
    $rg.Borders.LineStyle = 1; $rg.Borders.Color = 0xBFBFBF
  }
  $wb.Worksheets.Item(1).Activate()
  if (Test-Path -LiteralPath $OUT) { $bk = 'C:\ai\백업\MES매칭대조'; New-Item -ItemType Directory -Force $bk | Out-Null; Copy-Item -LiteralPath $OUT (Join-Path $bk ("대조_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.xlsx')) }
  $tmp = Join-Path $W 'out.xlsx'; if (Test-Path $tmp) { Remove-Item $tmp }; $wb.SaveAs($tmp, 51); $wb.Close($false)
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl); $xl = $null; [GC]::Collect(); [GC]::WaitForPendingFinalizers(); Start-Sleep 2; foreach ($p in @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $mine -contains $_.Id -and $_.MainWindowHandle -eq 0 })) { Stop-Process -Id $p.Id -Force } }
Copy-Item -LiteralPath $tmp $OUT -Force
"만듦: $OUT ($([math]::Round((Get-Item -LiteralPath $OUT).Length / 1KB)) KB)"
"거래처×월: " + (($rowsC | Group-Object 판단 | ForEach-Object { "$($_.Name) $($_.Count)" }) -join ', ')
"해외 인보이스: " + (($rowsI | Group-Object 판단 | ForEach-Object { "$($_.Name) $($_.Count)" }) -join ', ')
"남은 '확인 필요'(이유 없음): " + (($rowsC + $rowsI | Where-Object { $_.판단 -eq '확인 필요' -and -not $_.이유 } | ForEach-Object { "$($_.거래처) $($_.월)$($_.인보이스)" }) -join ' / ')
"10월말 이월: 그쪽 국내 $(F $carryTh['국내영업']) 해외 $(F $carryTh['해외영업']) / 내 국내 $(F $carryMy['국내']) 해외 $(F $carryMy['해외'])"
"해외 기타 합계 $(F $othTot), ALLISON $(F $allisonOth), 아트라스 1~5월 $(F $atlasMiss), 신우텍 $(F $shinwoo)"
