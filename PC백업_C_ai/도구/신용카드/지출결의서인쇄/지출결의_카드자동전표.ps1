param(
  [string]$Data = "$env:TEMP\카드전표\자료.json",                 # 지출결의_카드자료.ps1 결과
  [string]$Journal = "$env:TEMP\참고1_전표.csv",                   # 1~11월 전표(참고1) CSV — 없거나 참고1.xlsx보다 오래되면 아래 원본에서 자동으로 다시 만듦
  [string]$JournalXlsx = 'C:\ai\매월마감\국내자동전표\참고1.xlsx',   # 참고1 원본(아마란스 전표 받은 엑셀)
  [string]$TradeList = "$env:TEMP\카드전표\거래처목록.json",          # 거래처목록_받기.ps1 결과(아마란스 거래처 전체, 오늘 받은 것만 씀)
  [string]$TradeTemplate = (Join-Path $PSScriptRoot '거래처엑셀업로드_양식.xlsx'),   # 아마란스 거래처엑셀업로드 양식(양식만들기로 받은 것)
  [string]$Template = 'C:\ai\매월마감\해외자동전표\자동전표처리(일반) 양식.xlsx',
  [string]$OutDir = "$env:USERPROFILE\Desktop\지출결의서\결과",
  [string]$VatBook = 'C:\ai\매월마감\국내자동전표\참고4.xlsx',          # 부가세 매입매출장: 결제카드 거래처코드·카드번호
  [string]$BizStatus = "$env:TEMP\카드전표\사업자상태.csv",          # 사업자상태_조회.ps1 결과(국세청 일반/간이)
  [int]$StartNo = 3000,                                             # 작성번호: 날짜마다 3000부터(손 전표·해외 1000·국내 2000과 안 겹치게)
  [string]$DeptFile = (Join-Path $PSScriptRoot '부서코드.csv'),      # 아마란스 부서 코드표(부서목록_받기.ps1) → 사용부서(CT_DEPT)
  [bool]$AddName = $true                                            # 적요에 사람 이름이 없으면 '(사원 또는 작성자)'를 붙임 — 모든 카드 (2026-11-27 사용자)
)
# 지출결의서(법인카드) → 아마란스 자동전표처리(일반) 업로드 엑셀 + 검토 엑셀
# 규칙(2026-11-13 사용자 지정 + 1~11월 실제 카드 전표):
#  · 카드 1건(가맹점) = 전표 1장 (같은 날이어도 거래처마다 따로)
#  · 차변 비용(가맹점 거래처·사업자번호·프로젝트) / 부가세 있으면 차변 1350000 부가세대급금(세무구분 27, 신고기준일=카드 쓴 날, 공급가액, 카드번호) / 대변 2530000 미지급금(카드사)
#  · 미지급금 거래처 / 부가세 결제카드: 둘 다 카드사 코드로 통일 — 기업 99706 · 경남 99709 · 농협 99708 · 현대 99903 · 하나 99905 · 국민 99701 (2026-11-19 사용자 지정)
#  · 적요: 지출내역 내용 + "_#기업6941"(카드사+끝 4자리), 품의내역도 같게
#  · 택시·주유·주차·대리·철도·숙박 = 여비교통비, 해외 항공 = 해외출장비
#  · 이미 1~11월 전표에 같은 날·같은 금액·같은 카드로 있는 건은 넣지 않음(이중 방지)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Web.Extensions
$ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer; $ser.MaxJsonLength = [int]::MaxValue
$docs = $ser.DeserializeObject([IO.File]::ReadAllText($Data, [Text.Encoding]::UTF8))
# 참고1 전표 CSV(임시 폴더)가 없거나 원본 참고1.xlsx가 더 새것이면 다시 만듦 (2026-11-26: 임시 파일이 지워져 여기서 멈췄음)
$needJ = -not (Test-Path -LiteralPath $Journal)
if (-not $needJ -and (Test-Path -LiteralPath $JournalXlsx)) { $needJ = (Get-Item -LiteralPath $Journal).LastWriteTime -lt (Get-Item -LiteralPath $JournalXlsx).LastWriteTime }
if ($needJ) {
  if (-not (Test-Path -LiteralPath $JournalXlsx)) { throw "참고 전표 원본이 없습니다: $JournalXlsx (아마란스에서 1~11월 전표를 엑셀로 받아 이 이름으로 저장해 주세요)" }
  Write-Host "  참고 전표 목록을 원본에서 다시 만드는 중... ($JournalXlsx)"
  . (Join-Path $PSScriptRoot '참고1_CSV만들기.ps1')
  $nJ = ConvertTo-JournalCsv $JournalXlsx $Journal
  Write-Host "  → 참고 전표 $nJ 줄 준비됨"
}
$J = Import-Csv $Journal -Encoding UTF8
$bizKind = @{}; if (Test-Path -LiteralPath $BizStatus) { foreach ($s9 in (Import-Csv $BizStatus -Encoding UTF8)) { $bizKind[$s9.사업자번호] = $s9.구분 } }
# 결제카드: 언제나 카드사 거래처코드. 카드번호별로 따로 등록된 코드(예: 90009)는 쓰지 않는다 (2026-11-19 사용자 지정)
. 'C:\ai\도구\엑셀없이_xlsx읽기.ps1'
$payCardByNo = @{}; foreach ($vr in ((Read-Xlsx $VatBook 1).Rows | Where-Object { $_.R -gt 1 })) { $no = ("$($vr.Cells[42])" -replace '\D', ''); $cc = "$($vr.Cells[40])".Trim(); if ($no.Length -ge 15 -and $cc) { $payCardByNo[$no] = @($cc, "$($vr.Cells[41])".Trim()) } }
function Num($s) { $t = ("$s" -replace '[^\d\-]', ''); if ($t) { [double]$t } else { 0 } }

# ── 참고표 ──
$trByBiz = @{}; $seenTr = @{}
foreach ($l in $J) { $b = ($l.사업자번호 -replace '\D', ''); if ($b.Length -eq 10 -and $l.거래처코드 -and $l.거래처코드 -notmatch '^999' -and -not $seenTr.ContainsKey("$b|$($l.거래처코드)")) { $seenTr["$b|$($l.거래처코드)"] = 1; if (-not $trByBiz.ContainsKey($b)) { $trByBiz[$b] = New-Object Collections.ArrayList }; [void]$trByBiz[$b].Add(@($l.거래처코드, "$($l.거래처명)".Trim())) } }
# 아마란스 거래처 전체(거래처목록_받기.ps1, 오늘 받은 것) → ① 전표에 아직 안 쓴 거래처도 사업자번호로 찾음 ② 새 거래처 코드 = 가장 최근 등록 코드(예: 05723) 다음 번호부터, 있는 번호는 건너뜀 (2026-11-26 사용자: 새 거래처는 업로드 파일로 만들고 직접 올림)
$master = @(); $codeSet = @{}; $nextCode = $null
if ((Test-Path -LiteralPath $TradeList) -and (Get-Item -LiteralPath $TradeList).LastWriteTime.Date -eq (Get-Date).Date) {
  $master = @(([IO.File]::ReadAllText($TradeList, [Text.Encoding]::UTF8) | ConvertFrom-Json).rows)
  foreach ($m in $master) {
    $codeSet["$($m.trCd)"] = 1
    $b = ("$($m.regNb)" -replace '\D', '')
    if ($b.Length -eq 10 -and $m.trCd -and "$($m.trCd)" -notmatch '^999' -and -not $seenTr.ContainsKey("$b|$($m.trCd)")) { $seenTr["$b|$($m.trCd)"] = 1; if (-not $trByBiz.ContainsKey($b)) { $trByBiz[$b] = New-Object Collections.ArrayList }; [void]$trByBiz[$b].Add(@("$($m.trCd)", "$($m.trNm)".Trim())) }
  }
  # 가장 최근 등록(같은 시각에 여러 곳이면 그중 큰 코드). 2026-11-27부터 아마란스가 업로드 거래처에 00001~ 로 직접 매기므로 다음은 00008~ 예상(틀려도 '4. 새 거래처 코드 맞추기'가 고침)
  $lastReg = $master | Where-Object { "$($_.trCd)" -match '^\d{5}$' } | Sort-Object @{ Expression = { "$($_.insertDt)" }; Descending = $true }, @{ Expression = { "$($_.trCd)" }; Descending = $true } | Select-Object -First 1
  if ($lastReg) { $nextCode = [int]$lastReg.trCd + 1; Write-Host "  아마란스 거래처 $($master.Count)곳 / 가장 최근 등록 $($lastReg.trCd) $($lastReg.trNm) ($($lastReg.insertDt)) → 새 코드는 $('{0:D5}' -f $nextCode)부터" }
} else { Write-Host '  아마란스 거래처 목록(오늘 것)이 없어 새 거래처 코드는 비워 둡니다' -ForegroundColor Yellow }
function New-TrCode { while ($codeSet.ContainsKey(('{0:D5}' -f $script:nextCode))) { $script:nextCode++ }; $c0 = '{0:D5}' -f $script:nextCode; $codeSet[$c0] = 1; $script:nextCode++; $c0 }
$TRADD = [ordered]@{}   # 이번에 코드를 새로 매긴 거래처 (키: 사업자번호|가게이름)
# 이름 맞추기: 괄호·(주)·띄어쓰기·기호를 빼고 비교. '결제대행/가맹점' 모양이면 '/' 뒤 가맹점 이름끼리 비교(앞 결제대행 이름이 같아도 다른 가게)
#  같음 = 100, 한쪽이 다른 쪽을 품음(2글자 이상) = 50+글자수, 겹친 글자가 4글자 이상이면서 짧은 이름의 60% 이상 = 겹친 글자수, 그 밖 = 0
function NormNm($s) { ("$s" -replace '주식회사|\(주\)|㈜|\(재\)|\(사\)', '' -replace '[\s\(\)\[\]\-_/\.,·&]', '').ToLower() }
# 겹치는 가장 긴 글자 수(이름 맞추기 보조), 영수증 대표자·전화 정리(영수증마다 칸이 달라 '전화번호' 같은 글자·가려진 이름 '이*희'·'(02)-0000-0000'이 들어옴 → 비움) (2026-11-26)
function LcsLen($a, $b) { $x = NormNm (ShopPart $a); $y = NormNm (ShopPart $b); $best = 0; for ($i = 0; $i -lt $x.Length; $i++) { for ($k = $i + $best + 1; $k -le $x.Length; $k++) { if ($y.Contains($x.Substring($i, $k - $i))) { $best = $k - $i } else { break } } }; $best }
function CleanCeo($s) { $t = "$s".Trim(); if ($t -match '전화|번호|주소|사업자|대표|가맹|\*' -or $t.Length -gt 30) { '' } else { $t } }
function CleanTel($s) { $t = "$s".Trim(); $dg = $t -replace '\D', ''; if ($dg.Length -lt 8 -or $dg -match '0000000') { '' } else { $t } }
function ShopPart($s) { $t = "$s"; if ($t.Contains('/')) { $t.Substring($t.LastIndexOf('/') + 1) } else { $t } }
function NameScore($a, $b) {
  $x = NormNm (ShopPart $a); $y = NormNm (ShopPart $b); if (-not $x -or -not $y) { return 0 }
  if ($x -eq $y) { return 100 }
  if ($x.Contains($y) -or $y.Contains($x)) { $m = [Math]::Min($x.Length, $y.Length); if ($m -ge 2) { return 50 + $m } }
  $best = 0; for ($i = 0; $i -lt $x.Length; $i++) { for ($k = $i + $best + 1; $k -le $x.Length; $k++) { if ($y.Contains($x.Substring($i, $k - $i))) { $best = $k - $i } else { break } } }
  if ($best -ge 4 -and $best -ge 0.6 * [Math]::Min($x.Length, $y.Length)) { return $best } else { return 0 }
}
$PGRX = '/|이니시스|KCP|KICC|나이스|NICE|TOSS|토스(?!트)|카카오페이|네이버페이|인터넷상거래|결제|결재|요금창구|세외수입|스마트로'   # 결제대행을 거친 가맹점 이름 ('토스트'는 아님, 2026-11-26)
# 결제대행 사업자번호 = 그 번호로 등록된 거래처 이름 중 결제대행 모양(예: 'KCP결제/한국표준협회', '쏘카_KICC/쏘카')이 2곳 이상이거나 25% 이상 → 가게마다 따로 거래처
#   (KCP 85%·이니시스 58%·토스 67%·KICC 67%·나이스 50%·스마트로 50% / 코레일유통은 17곳 중 0곳 → 일반 사업자)
$pgBiz = @{}; foreach ($b9 in @($trByBiz.Keys)) { $all9 = @($trByBiz[$b9]); $pg9 = @($all9 | Where-Object { "$($_[1])" -match $PGRX }).Count; if ($pg9 -ge 2 -or ($pg9 -ge 1 -and $pg9 / $all9.Count -ge 0.25)) { $pgBiz[$b9] = 1 } }
# 거래처코드별 전표 줄 수(참고1) — 같은 사업자번호에 여러 곳이면 많이 쓴 곳을 고를 때 씀
$trUse = @{}; foreach ($l in $J) { $c9 = "$($l.거래처코드)"; if ($c9) { if ($trUse.ContainsKey($c9)) { $trUse[$c9]++ } else { $trUse[$c9] = 1 } } }
$pjByName = @{}; foreach ($g in ($J | Where-Object { $_.프로젝트코드 } | Group-Object 프로젝트코드)) { $nm = ($g.Group | Group-Object 프로젝트명 | Sort-Object Count -Descending | Select-Object -First 1).Name; $pjByName[$nm] = $g.Name; $pjByName[($nm -replace '\s*\(.*$', '')] = $g.Name }
$CARD = @{ '기업' = @('99706', '기업비씨카드'); '농협' = @('99708', '농협비씨카드'); '경남' = @('99709', '경남비씨카드(2113)'); '하나' = @('99905', '하나카드'); '국민' = @('99701', '국민카드'); '현대' = @('99903', '현대카드') }
foreach ($k in @($CARD.Keys)) { $hit = $J | Where-Object { $_.거래처코드 -eq $CARD[$k][0] } | Select-Object -First 1; if ($hit) { $CARD[$k] = @($CARD[$k][0], $hit.거래처명.Trim()) } }
# 카드 전표(미지급금이 카드사인 전표)의 비용 줄 → 카드 끝 4자리별 경향
$cardCodes = @($CARD.Values | ForEach-Object { $_[0] } | Where-Object { $_ })
$cardV = @{}; foreach ($l in $J) { if ($cardCodes -contains $l.거래처코드 -and $l.계정과목코드 -eq '2530000') { $cardV["$($l.작성일)|$($l.작성번호)"] = 1 } }
$hist = @($J | Where-Object { $cardV.ContainsKey("$($_.작성일)|$($_.작성번호)") -and $_.차변금액 -and $_.계정과목코드 -match '^[58]' } | ForEach-Object {
  $l4 = if ($_.적요 -match '(\d{4})\s*$') { $Matches[1] } else { '' }
  [pscustomobject]@{ l4 = $l4; acct = $_.계정과목코드; pj = $_.프로젝트코드; d = $_.작성일; tr = $_.거래처명; amt = (Num $_.차변금액); rmk = $_.적요 } })
$payLines = @($J | Where-Object { $_.계정과목코드 -eq '2530000' -and $_.대변금액 -and $cardCodes -contains $_.거래처코드 } | ForEach-Object {
  [pscustomobject]@{ d = [datetime]$_.작성일; amt = (Num $_.대변금액); l4 = $(if ($_.적요 -match '(\d{4})\s*$') { $Matches[1] } else { '' }); no = $_.작성번호; rmk = $_.적요 } })

# 계정: 분류 → (판관, 제조)
$ACC = @{ '복리후생비' = @('8110000', '5110000'); '여비교통비' = @('8120000', '5120000'); '접대비' = @('8130000', '5130000'); '교육훈련비' = @('8250000', '5250000'); '소모품비' = @('8300000', '5300000'); '지급수수료' = @('8310000', '5310000'); '해외출장비' = @('8340000', '8340000'); '차량유지비' = @('8220000', '5220000'); '통신비' = @('8140000', '8140000'); '도서인쇄비' = @('8260000', '5260000'); '광고선전비' = @('8330000', '8330000') }
$ACCNM = @{ '8110000' = '복리후생비(판)'; '5110000' = '복리후생비(제)'; '8120000' = '여비교통비(판)'; '5120000' = '여비교통비(제)'; '8130000' = '접대비(판)'; '5130000' = '접대비(제)'; '8250000' = '교육훈련비(판)'; '5250000' = '교육훈련비(제)'; '8300000' = '소모품비(판)'; '5300000' = '소모품비(제)'; '8310000' = '지급수수료(판)'; '5310000' = '지급수수료(제)'; '8340000' = '해외출장비'; '8220000' = '차량유지비(판)'; '5220000' = '차량유지비(제)'; '8140000' = '통신비(판)'; '8260000' = '도서인쇄비(판)'; '5260000' = '도서인쇄비(제)'; '8330000' = '광고선전비(판)'; '1350000' = '부가세대급금'; '2530000' = '미지급금' }

# 사용부서: 지출결의서 '작성자/작성부서'의 작성부서 이름 → 부서 코드 (2026-11-27 사용자 지정. 예: 품질관리본부_지점5 → 6600). 비용 줄에만 넣음
$deptByName = @{}; if (Test-Path -LiteralPath $DeptFile) { foreach ($dr0 in (Import-Csv -LiteralPath $DeptFile -Encoding UTF8)) { $deptByName["$($dr0.부서명)".Trim()] = "$($dr0.부서코드)".Trim() } } else { Write-Host "  부서 코드표가 없어 사용부서는 비웁니다: $DeptFile" -ForegroundColor Yellow }
$deptMiss = @{}
# 적요 이름 붙이기용: 이번 지출결의서에 나온 사람 이름(사원·작성자) + 사람 이름 모양 알아보기
$KNOWN = @{}; foreach ($d0 in $docs) { $w0 = "$($d0['info']['작성자/작성부서'])"; if ($w0.Contains('/')) { $n0 = $w0.Substring(0, $w0.IndexOf('/')).Trim(); if ($n0.Length -ge 2) { $KNOWN[$n0] = 1 } }; foreach ($i0 in $d0['items']) { $e0 = "$($i0['emp'])".Trim(); if ($e0.Length -ge 2 -and $e0 -ne '-') { $KNOWN[$e0] = 1 } } }
$NOTNAME = '영업|생산|품질|관리|본사|공장|외근|야근|출장|회의|점심|저녁|중식|석식|조식|간식|기타|일반|법인|개인|현금|국내|해외|주차|택시|대리|교통|숙박|식대|음료|다과|선물|공항|병원|문구|구매|구입|창원|마산|김해|부산|서울|대전|세종|울산|진주|사무실|임원|직원|손님|고객|업체|협력|신규|기존|정기|오전|오후|왕복|편도|신입|대표|회장|사장|담당|현장|경영|기획|재무|회계|인사|총무|구매팀|영업팀'
function Test-HasName([string]$s) {
  if ($s -match '외\s*\d+\s*(명|인)') { return $true }                                   # '정우석 팀장 외 1명'
  foreach ($k in $KNOWN.Keys) { if ($s.Contains($k)) { return $true } }                 # 이번 지출결의서의 사원·작성자 이름
  foreach ($m in [regex]::Matches($s, '([가-힣]{2,4})\s?(팀장|과장|차장|부장|수석|책임|선임|주임|실장|본부장|소장|공장장|사원|매니저|이사|상무|전무)')) { if ($m.Groups[1].Value -notmatch "^($NOTNAME)$") { return $true } }   # '서정수실장', '박준모 대표이사'
  foreach ($m in [regex]::Matches($s, '\(\s*([가-힣]{2,4})[\s,·/)]')) { if ($m.Groups[1].Value -notmatch "^($NOTNAME)$") { return $true } }   # '(정우석)', '(장성우,김항석)'
  return $false
}
# 적요 낱말 맞추기 (2026-11-27 사용자): ⓪ 야근식대 → 석식대, 주차비·전철비·택시비 → 교통비(항상) ① 식대·식사는 영수증 거래 시간으로 오후 4시 전 = 중식대, 4시부터 = 석식대 ② 음료·음료수는 '음료대'
#  · 내용에 식사 낱말(중식·석식·식대·식사·저녁식대 등)이 하나면 그 낱말을 시간에 맞는 중식대/석식대로 바꿈. 중식·석식이 같이 있으면(여러 끼) 그대로, 조식은 그대로
#  · 식사 낱말이 없어도 용도가 식대·식사면 뒤에 중식대/석식대를 붙임(예: 용도 '식대(영업)' + 내용 '고객업무협의' → '고객업무협의 석식대'). 영수증 시간이 없으면 안 바꿈
#  · 음료·음료수(구매·구입·값·비)는 '음료대'로. 다과·간식·식사와 같이 산 건은 그대로
function Fix-RemarkText([string]$desc, [string]$use, [string]$dtText) {
  $s = "$desc".Trim()
  $hh = $null; if ($dtText -match '\b(\d{1,2}):(\d{2})') { $hh = [int]$Matches[1] }
  # 2026-11-27 사용자 '항상': 야근식대 → 석식대(시간 상관없이), 주차비 → 교통비
  $night = ($s -match '야근\s?(?:식대|식사|석식대|석식|저녁)')
  $s = $s -replace '야근\s?(?:식대|식사|석식대|석식|저녁)', '석식대'
  $s = $s -replace '주차\s?(?:비|료|요금|대)', '교통비'
  $s = $s -replace '(?:전철|지하철)\s?(?:비|료|요금|대)', '교통비'   # 2026-11-27 사용자: 전철비도 교통비
  $s = $s -replace '택시\s?(?:비|료|요금|대|이용료)', '교통비'   # 2026-11-27 사용자: 택시비도 교통비
  $mealRx = '(?:점심|저녁|중식|석식)\s?(?:식대|식사)|중식대|석식대|식사대|중식|석식|식사|식대'
  $nonMeal = ($s -match '교통|택시|대리|주차|숙박|유류|주유|항공|렌트|통행') -or ("$use" -match '교통|유류|숙박|주차|항공|렌트')   # 택시·숙박 등에 들어간 '석식 후' 같은 말은 안 바꿈
  $drink = ($s -match '음료') -and ($s -notmatch '다과|간식|도시락|과자|디저트')
  if ($drink) {
    # 음료 건: '중식 음료' → '음료대' (사용자가 11/26 파일에서 '협의 중식 음료'를 '협의 음료대'로 고침)
    $s = $s -replace '(?:중식|석식|점심|저녁)\s?(?=음료)', ''
    $s = $s -replace '음료수?\s?(?:구매|구입|값|비|대)?', '음료대'
  }
  elseif ($null -ne $hh -and -not $nonMeal -and -not $night -and $s -notmatch '조식|아침') {
    $want = $(if ($hh -lt 16) { '중식대' } else { '석식대' })
    $both = ($s -match '중식|점심') -and ($s -match '석식|저녁')
    if ($s -match $mealRx) { if (-not $both) { $s = ([regex]$mealRx).Replace($s, $want, 1); $s = $s -replace "$want\s*대", $want } }
    elseif ("$use" -match '식대|식사|중식|석식' -and $s -notmatch '다과|간식|도시락|커피') { $s = "$s $want".Trim() }
  }
  return $s.Trim()
}
$VOU = New-Object Collections.ArrayList      # 전표
$SKIP = New-Object Collections.ArrayList   # 뺀 것
$CHK = New-Object Collections.ArrayList    # 확인할 것
[void]$CHK.Add([pscustomobject]@{ 문서 = '(전체)'; 내용 = '전표발행'; 확인 = '승인일은 업로드 엑셀에 칸이 없고 아마란스 자동전표처리에서 전표발행 때 정해짐 — 발행일자를 한 날짜로 정하지 말고 발행해야 승인일 = 각 전표 작성일자(카드 쓴 날). 10월 자동전표 342줄이 10/31로 승인된 것이 발행일자를 정한 경우' })
foreach ($d in $docs) { foreach ($it in $d['items']) {
  $rc = $it['receipt']; if (-not $rc) { [void]$SKIP.Add([pscustomobject]@{ 문서 = $d['docNo']; 내용 = $it['desc']; 사유 = '신용카드매출전표가 없어 넣지 않음' }); continue }
  $dt = [datetime](($rc['dt'] -split ' ')[0]); $total = Num $it['total']; $sup = Num $it['sup']; $vat = Num $it['vat']
  if ($sup + $vat -ne $total) { [void]$CHK.Add([pscustomobject]@{ 문서 = $d['docNo']; 내용 = $it['desc']; 확인 = "공급가액+부가세($($sup + $vat))가 합계($total)와 다름" }) }
  $alias = "$($rc['cardAlias'])"; $co = ($CARD.Keys | Where-Object { $alias -like "$_*" } | Select-Object -First 1)
  $l4 = if ("$($rc['cardNo'])" -match '(\d{4})\s*$') { $Matches[1] } else { '' }
  if (-not $co) { [void]$CHK.Add([pscustomobject]@{ 문서 = $d['docNo']; 내용 = $it['desc']; 확인 = "카드사를 알 수 없음: $alias" }); continue }
  # 이미 전표 있음? (같은 카드 끝4자리·같은 금액·날짜 ±1일)
  $dup = @($payLines | Where-Object { $_.l4 -eq $l4 -and $_.amt -eq $total -and [Math]::Abs(($_.d - $dt).TotalDays) -le 1 })
  if (-not $dup.Count -and ($it['desc'] + $rc['name']) -match 'OPENAI|CHATGPT') {   # ChatGPT는 월말에 '외 N건'으로 묶어 처리해 옴
    $mon = @($hist | Where-Object { $_.l4 -eq $l4 -and $_.tr -match 'OPENAI' -and ([datetime]$_.d).ToString('yyyyMM') -eq $dt.ToString('yyyyMM') })
    if ($mon.Count) { $dup = @([pscustomobject]@{ d = [datetime]$mon[-1].d; no = ''; rmk = "$($mon[-1].rmk) (같은 달 OPENAI $($mon.Count)줄, 금액은 카드 청구 기준이라 다름)" }) }
  }
  if ($dup.Count) { [void]$SKIP.Add([pscustomobject]@{ 문서 = $d['docNo']; 내용 = $it['desc']; 날짜 = $dt.ToString('yyyy-MM-dd'); 금액 = $total; 카드 = "$co$l4"; 사유 = "이미 전표 있음: $($dup[0].d.ToString('yyyy-MM-dd')) 번호 $($dup[0].no) [$($dup[0].rmk)]" }); continue }

  $use = "$($it['use'])"; $desc = "$($it['desc'])"; $txt = "$use $desc"
  # 분류
  $cat = $null; $why = ''
  # 2026-11-27 대표님 지출결의(업무지원실 박새별, 하나 #0073) 대비 — 1~11월 #0073 카드 전표 442장의 처리 방식에 맞춤(되돌려 보기로 확인)
  #  · 국내 항공권 = 여비교통비(부가세 없음), 해외 출장 항공·숙박·일비·취소수수료 = 해외출장비, 국내 항공권 발권 수수료 = 지급수수료
  #  · 렌트카·일일기사 = 여비교통비(부가세 있으면 대급금), 택시 가맹점 = 여비교통비, 대표 차량 세차·광택 = 차량유지비, 택배·우편 = 통신비
  #  · 도서·현수막 = 도서인쇄비, 회사 기념품·각인 = 광고선전비, 손님·VIP·고객사 선물 = 접대비, 퇴직·생일 등 내부 선물 = 복리후생비
  #  · 임원 건강식품·개인물품·사진인쇄·액자 = 복리후생비, 임원실 다과·음료 = 복리후생비(먹을 것을 '구매'보다 먼저 봄)
  $vendor = "$($rc['name'])"
  $air = ($txt -match '항공권|항공료|항공편|항공운임|항권공|발권|비행기|탑승권') -or ($vendor -match '(?<!공)항공(?!청)|에어부산|에어서울|진에어|티웨이|이스타항공|아시아나|에어로케이|에어프레미아|\bAIR(?!BNB)')
  $abroad = ($txt -match '해외|국제선|외국|스위스|일본|미국|중국|베트남|독일|유럽|태국|대만|홍콩|싱가포르|인도네시아|멕시코|오사카|도쿄|동경|후쿠오카|상하이|방콕')
  $inside = ($txt -match '직원|사내|임직원|소통|포상|시상|창립|명절|퇴직|퇴임|생일|결혼|출산|조의|부의')
  if ($abroad -and ($air -or $txt -match '출장|숙박|호텔|일비')) { $cat = '해외출장비' }
  elseif ($air -and $txt -match '수수료') { $cat = '지급수수료'; $why = '항공권 발권 수수료(예전 전표 지급수수료)' }
  elseif ($air) { $cat = '여비교통비'; $why = '국내 항공권(예전 대표님·회장님 항공권 = 여비교통비, 부가세 없음)'; if ($total -gt 400000) { [void]$CHK.Add([pscustomobject]@{ 문서 = $d['docNo']; 내용 = $desc; 확인 = "항공권 $('{0:N0}' -f $total)원 — 금액이 커서 해외 항공권일 수 있음(해외면 해외출장비). 국내로 보고 여비교통비로 넣음" }) } }
  elseif ($txt -match '렌트|렌탈|렌터카|쏘카|그린카|일일기사|기사 ?(이용|비)|의전 ?차량' -or $vendor -match '렌트카|렌터카|롯데렌탈|SK렌터카|쏘카|그린카') { $cat = '여비교통비'; $why = '렌트카·기사 이용료(예전 전표 여비교통비)' }
  elseif ($vendor -match '택시|카카오T|카카오모빌리티|티머니|이동의즐거움|대리운전|우티|UT ') { $cat = '여비교통비'; $why = '택시·대리 가맹점' }
  elseif ($txt -match '주차|택시|대리|유류|주유|기름|교통|철도|기차|열차|승차권|KTX|SRT|버스|숙박|모텔|호텔|통행' -or $vendor -match '한국철도공사|코레일(?!유통)|에스알|도로공사') { $cat = '여비교통비' }
  elseif ($txt -match '포장용품|포장지|포장 ?박스') { $cat = '소모품비' }
  elseif ($txt -match '세차|광택|타이어|엔진오일|차량 ?(정비|수리|점검)') { $cat = '차량유지비' }
  elseif ($txt -match '탁송') { $cat = '지급수수료' }
  elseif (($txt -match '택배|우편|소포|등기|발송|배송' -or $vendor -match '우정사업본부|우체국') -and $txt -notmatch '박스|상자') { $cat = '통신비' }
  elseif ($txt -match '회사 ?기념품|대신금속 ?기념품|기념품 ?각인|판촉') { $cat = '광고선전비'; $why = '회사 기념품·각인(예전 전표 광고선전비)' }
  elseif ($txt -match '양압기|건강|영양제|의료|약품|비타민|개인 ?물품') { $cat = '복리후생비'; $why = '임원 건강 관련(예전 임원 건강식품·회장님 개인물품 = 복리후생비)'; if ($txt -match '양압기|의료') { [void]$CHK.Add([pscustomobject]@{ 문서 = $d['docNo']; 내용 = $desc; 확인 = '임원 건강·의료 물품 — 예전 전표에 같은 물건이 없어 임원 건강식품처럼 복리후생비(판)로 넣음. 계정 확인' }) } }
  elseif ($txt -match '사진|인화|액자') { $cat = '복리후생비'; $why = '사진·액자(예전 임원 사진인쇄·사내액자 = 복리후생비)' }
  elseif ($txt -match '도서|서적|(?<!명함)책 ?(구입|구매)|현수막|인쇄|출력' -or $vendor -match '교보문고|영풍문고|알라딘|예스24|YES24') { $cat = '도서인쇄비' }
  elseif ($txt -match '교육') { $cat = '교육훈련비' }
  elseif ($txt -match '사용료|수수료|구독|CHATGPT|OPENAI') { $cat = '지급수수료' }
  elseif ($txt -match '식대|중식|석식|조식|점심|저녁|음료|간식|다과|회식|도시락|디저트|케이크|빵|과자|떡|과일|커피(?!포트|머신|메이커)') { $cat = '복리후생비' }
  elseif ($txt -match '마이크|구매|구입|비품|소모품|물품|용품|명함|봉투|정리함|수납|문구|필기') { $cat = '소모품비' }
  # 선물·기념품: 손님·고객사·VIP = 접대비(금액 상관없이, 예전 'vip선물'·'고객사 기념선물'), 직원·퇴직·생일 등 내부 = 복리후생비. '선물 포장용품'은 소모품비 그대로
  if ($cat -notin '광고선전비', '해외출장비' -and $txt -match '선물|기념품' -and $txt -notmatch '선물 ?포장') {
    if ($inside) { $cat = '복리후생비'; $why = '직원·내부 선물·기념품 = 복리후생비(예전 전표)' }
    else { $cat = '접대비'; $why = '손님·고객사 선물 = 접대비(부가세 불공제, 예전 VIP·고객사 선물 전표)' }
  }
  # 손님(고객) 건: 식대이거나, 용도로 계정을 못 정한 건(예: 용도 '기타' + 내용 '고객업무협의')도 같이 봄 (2026-11-26 사용자 지정: "고객업무 이런 거 나오면 접대, 12만원 넘는 건")
  #  2026-11-27: '손님·방문객·VIP·접식' 등도 손님으로 봄 (예: '손님 방문 제공 음료 및 다과'). '임원 및 손님' 같이 섞인 다과는 예전처럼 복리후생비
  if ((-not $cat -or $cat -eq '복리후생비') -and "$use $desc" -match '신코|머큐리|고객|접대|접식|거래처|회계법인|법무법인|세무법인|세무사|손님|방문객|내방|VIP|귀빈') {
    $how = if ($cat) { '식대' } else { "건(용도 '$use')" }
    if ($txt -match '임원|직원' -and $txt -match '손님') { $cat = '복리후생비'; $why = '임원·손님 같이 쓰는 다과·음료(예전 전표 복리후생비)' }
    elseif ($total -le 120000) { $cat = '복리후생비'; $why = "손님(고객) $how 이지만 12만원 이하라 복리후생비(2026-11-19 사용자 지정)" }
    else { $cat = '접대비'; $why = "손님(고객) $how, 12만원 초과라 접대비(부가세 불공제, 2026-11-26 사용자 지정)" }
  }
  if (-not $cat) { $cat = '복리후생비'; [void]$CHK.Add([pscustomobject]@{ 문서 = $d['docNo']; 내용 = $desc; 확인 = "용도 '$use'로 계정을 정하지 못해 복리후생비로 넣음 — 계정 확인" }) }
  # 프로젝트 (먼저 정함 — 판관/제조 판단에 씀)
  $pj = ''; $pjWhy = ''
  foreach ($cand in @("$($it['proj'])", "$($d['info']['프로젝트'])")) { if ($cand -and $cand -ne '-' -and $pjByName.ContainsKey($cand)) { $pj = $pjByName[$cand]; $pjWhy = "지출결의서 프로젝트 '$cand'"; break } }
  if (-not $pj -and $desc -match '창원공장') { $pj = '103'; $pjWhy = "내용에 '창원공장'" }
  if (-not $pj) { $hp = @($hist | Where-Object { $_.l4 -eq $l4 -and $_.pj } | Group-Object pj | Sort-Object Count -Descending); if ($hp.Count -and $hp[0].Count -ge 5 -and $hp[0].Count / (($hp | Measure-Object Count -Sum).Sum) -ge 0.6) { $pj = $hp[0].Name; $pjWhy = "지출결의서에 없어 이 카드(#$l4)가 주로 쓴 프로젝트" } }
  if (-not $pj) { $pj = '303'; $pjWhy = '지출결의서에 없어 대신금속(303)' }
  # 판관/제조: ① 이 줄 용도 ② 같은 문서 다른 줄 용도 ③ 부서(품질·생산 쪽 = 제조) ④ 프로젝트(303 = 판관, 공장·라인 = 제조)
  $docHint = $null; foreach ($o in $d['items']) { if ("$($o['use'])" -match '사무실|관리|영업') { $docHint = 0; break } elseif ("$($o['use'])" -match '품질|생산') { $docHint = 1; break } }
  if ($cat -eq '해외출장비') { $side = 0 }
  elseif ($use -match '사무실|관리|영업') { $side = 0 }
  elseif ($use -match '품질|생산') { $side = 1 }
  elseif ($null -ne $docHint) { $side = $docHint }
  elseif ("$($it['dept'])" -match '품질|생산|가공|금형|주조') { $side = 1 }
  elseif ($pj -eq '303') { $side = 0 }
  else { $side = 1 }
  $acct = $ACC[$cat][$side]
  # 부가세 줄
  $vatLine = ($vat -gt 0)
  if ($vatLine -and $cat -eq '접대비') { $vatLine = $false }
  if ($vatLine -and $txt -match '유류|주유|기름') { $vatLine = $false; $why = '주유비 부가세는 그동안 공제하지 않아(1~11월 130장 중 3장만 공제) 합계로 비용 처리' }
  $bz = ("$($rc['bizNo'])" -replace '\D', ''); $kindTx = $(if ($bizKind.ContainsKey($bz)) { $bizKind[$bz] } else { '' })
  if ($vatLine -and $kindTx -like '간이*') { $vatLine = $false; $why = "간이과세자($kindTx, 국세청 조회)라 전표유형 일반 — 부가세대급금 안 넣고 합계로 비용(2026-11-13 사용자 지정)" }
  if ($vatLine -and $txt -match '주차') { $vatLine = $false; $why = '주차비는 부가세대급금 없이 합계로, 전표유형 일반(2026-11-13 사용자 지정)' }
  # 가맹점 거래처
  # 같은 사업자번호(결제대행사는 여러 가맹점이 한 번호)에 등록된 거래처 중 이름이 맞는 것만 씀, 없으면 새 거래처
  $biz = ("$($rc['bizNo'])" -replace '\D', ''); $trCd = ''; $trNm = "$($rc['name'])"; $newTr = $true
  if ($biz.Length -eq 10 -and $trByBiz.ContainsKey($biz)) {
    $best = $null; $bestS = 0
    foreach ($cand in $trByBiz[$biz]) { $s0 = NameScore $rc['name'] $cand[1]; if ($s0 -gt $bestS) { $bestS = $s0; $best = $cand } }
    # 사업자번호에 등록 거래처가 하나뿐이고 결제대행 모양이 아니면 이름이 달라도 그 거래처(예: 철도승차권 = 한국철도공사)
    if (-not $best -and $trByBiz[$biz].Count -eq 1 -and "$($rc['name'])" -notmatch $PGRX -and "$($trByBiz[$biz][0][1])" -notmatch $PGRX) { $best = $trByBiz[$biz][0] }
    # 결제대행 번호가 아니면 사업자번호가 같으면 같은 사업자 → 새로 만들지 않고 기존 거래처를 씀 (2026-11-26 사용자: "사업자번호가 같으면 굳이 새로 등록할 필요 없음")
    #   여러 곳이면 ① 이름이 4글자 이상 겹치는 곳(예: '스타벅스코리아' → '스타벅스커피코리아') ② 아니면 전표에 가장 많이 쓴 곳(예: 코레일유통 번호의 모르는 가게 → 코레일유통(주))
    if (-not $best -and -not $pgBiz.ContainsKey($biz) -and "$($rc['name'])" -notmatch $PGRX) {
      $bk = -1; $bc = $null
      foreach ($cand in $trByBiz[$biz]) { $l0 = LcsLen $rc['name'] $cand[1]; $u0 = $(if ($trUse.ContainsKey("$($cand[0])")) { $trUse["$($cand[0])"] } else { 0 }); $k0 = $(if ($l0 -ge 4) { 1000000 + $l0 * 10000 } else { 0 }) + [Math]::Min($u0, 9999); if ($k0 -gt $bk) { $bk = $k0; $bc = $cand } }
      if ($bc) {
        $best = $bc
        if ((NameScore $rc['name'] $bc[1]) -eq 0) { [void]$CHK.Add([pscustomobject]@{ 문서 = $d['docNo']; 내용 = $desc; 확인 = "영수증 이름 '$($rc['name'])'은 다르지만 사업자번호($biz)가 같은 기존 거래처 [$($bc[0])] $($bc[1]) 를 씀(새로 등록 안 함)$(if (@($trByBiz[$biz]).Count -gt 1) { " — 같은 번호로 $(@($trByBiz[$biz]).Count)곳 등록돼 있음" })" }) }
      }
    }
    if ($best) { $trCd = $best[0]; $trNm = $best[1]; $newTr = $false }
  }
  # 새 거래처: 코드를 매겨 전표에 넣고 '새거래처' 업로드 파일에도 같은 코드로 넣음 (같은 가게가 여러 번 나오면 같은 코드)
  if ($newTr -and $null -ne $nextCode -and $biz.Length -eq 10) {
    $nk = "$biz|$(NormNm (ShopPart $rc['name']))"
    if (-not $TRADD.Contains($nk)) {
      $nm0 = "$($rc['name'])".Trim(); if ($nm0.Length -gt 60) { $nm0 = $nm0.Substring(0, 60) }
      $TRADD[$nk] = [pscustomobject]@{ 코드 = (New-TrCode); 이름 = $nm0; 사업자번호 = $biz; 대표자 = (CleanCeo $rc['ceo']); 주소 = "$($rc['addr'])".Trim(); 전화 = (CleanTel $rc['tel']) }
      if ($trByBiz.ContainsKey($biz)) { [void]$CHK.Add([pscustomobject]@{ 문서 = $d['docNo']; 내용 = $desc; 확인 = "새 거래처 '$nm0'(코드 $($TRADD[$nk].코드)) — 결제대행 사업자번호($biz)라 가게마다 따로 등록함. 이 번호로 이미 $(@($trByBiz[$biz]).Count)곳: $((@($trByBiz[$biz] | Select-Object -First 6 | ForEach-Object { "$($_[0]) $($_[1])" })) -join ', ')$(if (@($trByBiz[$biz]).Count -gt 6) { ' …' }). 그중 같은 가게가 있으면 업로드 파일에서 이 줄을 지우고 전표 코드를 그 코드로 바꾸세요" }) }
    }
    $trCd = $TRADD[$nk].코드; $trNm = $TRADD[$nk].이름
  }
  # 대리운전·택시는 적요 '_#' 앞에 '교통비'를 넣음(예: 고객응대 교통비_#하나8508, 2026-11-13 사용자 지정). 이미 '교통비'가 있으면 그대로
  $descR = Fix-RemarkText $desc $use "$($rc['dt'])"   # 적요용 내용: 중식대/석식대(영수증 시간), 음료대
  $tail = ''; if (($use -match '대리' -or $desc -match '대리|택시' -or "$($rc['name'])" -match '대리|택시') -and $descR -notmatch '교통비') { $tail = ' 교통비' }
  # 적요에 쓴 사람 이름: 내용에 사람 이름이 없으면 지출결의서 '사원'(없으면 '작성자') 이름을 붙임 → '업무수행 교통비(박새별)_#하나0073', '주유비용(전승우)_#하나8508'
  #  2026-11-27 사용자: 처음엔 대표님 하나카드(#0073 박새별)만 → "이름이 없는 것들은 작성자 이름을 박새별처럼 붙여 줘" 모든 카드로 넓힘
  #  안 붙이는 것: 이미 사람 이름이 있음(사원·작성자 등 이번 지출결의서에 나온 이름, '(정우석)'·'(장성우,김항석)' 같은 괄호 이름, '정우석 팀장'·'외 1명')
  #               대표님·회장님 항공권(예전 #0073 전표에 한 번도 안 붙임)
  $who = ''
  if ($AddName) {
    $emp = "$($it['emp'])".Trim(); if (-not $emp -or $emp -eq '-') { $wd0 = "$($d['info']['작성자/작성부서'])"; $emp = $(if ($wd0.Contains('/')) { $wd0.Substring(0, $wd0.IndexOf('/')).Trim() } else { '' }) }
    $airRmk = ($desc -match '항공|발권|항권공|비행기|탑승권') -and ($desc -match '대표|회장|사장')
    if ($emp -and $emp -ne '-' -and -not $airRmk -and -not (Test-HasName $descR)) { $who = "($emp)" }
  }
  $suffix = "$tail$who`_#$co$l4"; $rmk = "$descR".Trim(); if (($rmk + $suffix).Length -gt 80) { $rmk = $rmk.Substring(0, 80 - $suffix.Length).Trim() }; $rmk = $rmk + $suffix
  $wd = "$($d['info']['작성자/작성부서'])"; $deptNm = $(if ($wd.Contains('/')) { $wd.Substring($wd.IndexOf('/') + 1).Trim() } else { '' })
  if (-not $deptNm) { $deptNm = "$($it['dept'])".Trim() }; if (-not $deptNm) { $deptNm = "$($d['info']['기 안 부 서'])".Trim() }
  $deptCd = $(if ($deptNm -and $deptByName.ContainsKey($deptNm)) { $deptByName[$deptNm] } else { '' })
  if ($deptNm -and -not $deptCd -and -not $deptMiss.ContainsKey($deptNm)) { $deptMiss[$deptNm] = 1; [void]$CHK.Add([pscustomobject]@{ 문서 = $d['docNo']; 내용 = $desc; 확인 = "작성부서 '$deptNm'의 부서 코드를 부서코드.csv에서 못 찾아 사용부서를 비움 → 부서목록_받기.ps1로 코드표를 새로 받거나 아마란스에서 직접 넣으세요" }) }
  [void]$VOU.Add([pscustomobject]@{ 문서 = $d['docNo']; 문서제목 = $d['title']; 순번 = $it['no']; 날짜 = $dt; 용도 = $use; 내용 = $desc; 가맹점 = "$($rc['name'])"; 사업자번호 = $biz; 과세유형 = $kindTx; 거래처코드 = $trCd; 거래처명 = $trNm; 새거래처 = $newTr
    사용부서 = $deptCd; 사용부서명 = $deptNm
    대표자 = (CleanCeo $rc['ceo']); 주소 = "$($rc['addr'])"; 전화 = (CleanTel $rc['tel']); 카드사 = $co; 카드번호 = ("$($rc['cardNo'])" -replace '\D', ''); 끝4 = $l4; 승인번호 = "$($rc['approval'])"
    계정 = $acct; 분류 = $cat; 프로젝트 = $pj; 프로젝트근거 = $pjWhy; 공급가액 = $sup; 부가세 = $vat; 합계 = $total; 부가세줄 = $vatLine; 적요 = $rmk; 참고 = $why })
} }

# 작성번호 (날짜마다 StartNo부터, 문서 순서대로)
# 2026-11-27: 예전에 만든 카드 전표 업로드 파일(결과·지난결과 폴더, 백업 제외)에서 이미 쓴 날짜·번호는 건너뜀.
#  (2026-11-27: '해외출장_카드자동전표_*.xlsx' 같은 다른 카드 전표 파일도 봄)
#  11/13에 올려 발행한 11/06·11/09·11/11 3000번대와 겹쳐 "이미 발행되어 추가입력할 수 없습니다"로 업로드가 실패했었음.
$usedNo = @{}; $todayUp = "지출결의_카드자동전표_$(Get-Date -Format 'yyyyMMdd').xlsx"
$prevDirs = @(@($OutDir, (Join-Path (Split-Path $OutDir) '지난결과')) | Where-Object { Test-Path -LiteralPath $_ })
foreach ($pf in @(Get-ChildItem -LiteralPath $prevDirs -Recurse -Filter '*카드자동전표_*.xlsx' -ErrorAction SilentlyContinue | Where-Object { $_.Name -notmatch '검토|분개확인' -and $_.FullName -notmatch '\\백업\\' -and $_.Name -ne $todayUp })) {
  try { $px = Read-Xlsx $pf.FullName 1 } catch { Write-Host "  (예전 전표 파일을 못 읽음: $($pf.Name))" -ForegroundColor Yellow; continue }
  $r2 = $px.Rows | Where-Object { $_.R -eq 2 } | Select-Object -First 1; if (-not $r2) { continue }
  $cDt = $null; $cSq = $null; foreach ($k9 in $r2.Cells.Keys) { $v9 = "$($r2.Cells[$k9])".Trim(); if ($v9 -eq 'MENU_DT') { $cDt = $k9 } elseif ($v9 -eq 'MENU_SQ') { $cSq = $k9 } }
  if (-not $cDt -or -not $cSq) { continue }
  foreach ($pr in ($px.Rows | Where-Object { $_.R -ge 4 })) { $d0 = "$($pr.Cells[$cDt])".Trim(); $s0 = "$($pr.Cells[$cSq])".Trim(); if ($d0 -match '^\d{8}$' -and $s0 -match '^\d+$') { $usedNo["$d0|$([int]$s0)"] = $pf.Name } }
}
$seq = @{}; $skipped = @{}
foreach ($v in $VOU) {
  $k = $v.날짜.ToString('yyyyMMdd'); $n0 = if ($seq.ContainsKey($k)) { $seq[$k] + 1 } else { $StartNo }
  while ($usedNo.ContainsKey("$k|$n0")) { $skipped[$k] = 1; $n0++ }
  $seq[$k] = $n0; $v | Add-Member 작성번호 $n0
}
if ($skipped.Count) {
  $first = @{}; foreach ($v in $VOU) { $k = $v.날짜.ToString('yyyyMMdd'); if ($skipped.ContainsKey($k) -and -not $first.ContainsKey($k)) { $first[$k] = $v.작성번호 } }
  Write-Host "  예전에 올린 전표와 번호가 겹치지 않게 건너뜀: $((@($first.Keys | Sort-Object | ForEach-Object { "$_ 은 $($first[$_])번부터" })) -join ', ')"
}
# 업로드 줄
$LINES = New-Object Collections.ArrayList
foreach ($v in $VOU) {
  $dts = $v.날짜.ToString('yyyyMMdd'); $ty = $(if ($v.부가세줄) { '2' } else { '1' }); $ln = 0
  $base = @{ IN_DIV_CD = '1000'; MENU_DT = $dts; MENU_SQ = $v.작성번호; ISU_DOC = $v.적요; DOCU_TY = $ty; RMK_DC = $v.적요 }
  $ln++; $e = $base.Clone(); $e.MENU_LN_SQ = $ln; $e.DRCR_FG = '3'; $e.ACCT_CD = $v.계정; $e.TR_CD = $v.거래처코드; $e.TR_NM = $v.거래처명; $e.ACCT_AM = $(if ($v.부가세줄) { $v.공급가액 } else { $v.합계 }); $e.PJT_CD = $v.프로젝트; $e.CT_DEPT = $v.사용부서; $e.REG_NB = $v.사업자번호
  if ($v.새거래처) { $e.CEO_NM = $v.대표자; $e.DIV_ADDR1 = $v.주소; $e.TEL = $v.전화 }
  [void]$LINES.Add($e)
  if ($v.부가세줄) { $ln++; $t = $base.Clone(); $t.MENU_LN_SQ = $ln; $t.DRCR_FG = '3'; $t.ACCT_CD = '1350000'; $t.TR_CD = $v.거래처코드; $t.TR_NM = $v.거래처명; $t.ACCT_AM = $v.부가세; $t.REG_NB = $v.사업자번호
    $t.VAT_DIV_CD = '1000'; $t.ISS_DT = $dts; $t.TAX_FG = '27'; $t.CT_DEAL = '27'; $t.SUP_AM = $v.공급가액; $t.CT_NB = $CARD[$v.카드사][0]   # 결제카드 = 언제나 카드사 거래처코드 (기업 99706 / 경남 99709 / 농협 99708 / 현대 99903 / 하나 99905 / 국민 99701) — 2026-11-19 사용자 지정: 카드번호별 코드는 쓰지 않는다
    if ($v.새거래처) { $t.CEO_NM = $v.대표자; $t.DIV_ADDR1 = $v.주소; $t.TEL = $v.전화 }
    [void]$LINES.Add($t) }
  $ln++; $c = $base.Clone(); $c.MENU_LN_SQ = $ln; $c.DRCR_FG = '4'; $c.ACCT_CD = '2530000'; $c.TR_CD = $CARD[$v.카드사][0]; $c.TR_NM = $CARD[$v.카드사][1]; $c.ACCT_AM = $v.합계
  [void]$LINES.Add($c)
}
if ($TRADD.Count) {
  $cds = @($TRADD.Values | ForEach-Object { $_.코드 })
  $CHK.Insert(0, [pscustomobject]@{ 문서 = '(전체)'; 내용 = "새 거래처 $($TRADD.Count)곳 먼저 등록"; 확인 = "전표 올리기 전에 '지출결의_새거래처_$(Get-Date -Format 'yyyyMMdd').xlsx'를 아마란스 회계관리 > 회계기초정보관리 > 거래처/계좌/카드관리 > 거래처엑셀업로드에서 먼저 등록(기능모음 > 파일업로드 → 변환하기). 파일의 코드 $($cds[0])~$($cds[-1])는 임시 번호 — 아마란스가 코드를 직접 다시 매김(11/27: 00001~). 변환하기까지 끝낸 뒤 바탕화면\지출결의서 '4. 새 거래처 코드 맞추기'를 실행하면 실제 코드로 전표 파일을 고침. 그다음 전표 올리기" })
}
foreach ($v in ($VOU | Where-Object 새거래처)) {
  $msg = if ($v.거래처코드) { "새 거래처: '$($v.가맹점)' 사업자번호 $($v.사업자번호) — 아마란스에 없음 → 코드 $($v.거래처코드)로 새로 매김('새거래처' 파일로 먼저 등록)" }
         elseif ($master.Count) { "새 거래처: '$($v.가맹점)' — 사업자번호가 없어 코드를 못 매김. 아마란스에 직접 등록 후 코드 넣기" }
         else { "새 거래처: '$($v.가맹점)' 사업자번호 $($v.사업자번호) — 아마란스 거래처 목록을 못 받아 거래처코드 비움. 아마란스에 거래처 등록 후 코드 넣기" }
  [void]$CHK.Add([pscustomobject]@{ 문서 = $v.문서; 내용 = $v.내용; 확인 = $msg })
}
foreach ($v in ($VOU | Where-Object { -not $CARD.ContainsKey("$($_.카드사)") })) { [void]$CHK.Add([pscustomobject]@{ 문서 = $v.문서; 내용 = $v.내용; 확인 = "모르는 카드사 '$($v.카드사)' — 거래처코드를 못 넣었습니다. 카드사 코드를 알려주세요(기업 99706 / 경남 99709 / 농협 99708 / 현대 99903 / 하나 99905 / 국민 99701)" }) }
foreach ($v in ($VOU | Where-Object { $_.참고 })) { [void]$CHK.Add([pscustomobject]@{ 문서 = $v.문서; 내용 = $v.내용; 확인 = $v.참고 }) }
foreach ($v in ($VOU | Where-Object { $_.프로젝트근거 -notlike '지출결의서 프로젝트*' })) { [void]$CHK.Add([pscustomobject]@{ 문서 = $v.문서; 내용 = $v.내용; 확인 = "프로젝트 $($v.프로젝트): $($v.프로젝트근거)" }) }

# ── 엑셀 ──
New-Item -ItemType Directory -Force $OutDir | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd'
$upload = Join-Path $OutDir "지출결의_카드자동전표_$stamp.xlsx"; $review = Join-Path $OutDir "지출결의_카드자동전표_$($stamp)_검토.xlsx"
$work = Join-Path $env:TEMP '카드전표'; $tmpUp = Join-Path $work "업로드_임시.xlsx"
$xl = New-Object -ComObject Excel.Application; $xl.DisplayAlerts = $false
try {
  if (Test-Path -LiteralPath $tmpUp) { Remove-Item -LiteralPath $tmpUp -Force }
  $wb = $xl.Workbooks.Open($Template, 0, $true); $wb.SaveAs($tmpUp, 51)
  $sh = $wb.Worksheets.Item(1); $maxC = $sh.UsedRange.Columns.Count; $colOf = @{}; $isText = @{}
  for ($ci = 1; $ci -le $maxC; $ci++) { $cd = "$($sh.Cells(2, $ci).Value2)".Trim(); if ($cd -and -not $colOf.ContainsKey($cd)) { $colOf[$cd] = $ci; $isText[$ci] = ("$($sh.Cells(3, $ci).Value2)" -match '타입\s*:\s*문자') } }
  $codes = 'IN_DIV_CD', 'MENU_DT', 'MENU_SQ', 'MENU_LN_SQ', 'ISU_DOC', 'DOCU_TY', 'DRCR_FG', 'ACCT_CD', 'TR_CD', 'TR_NM', 'ACCT_AM', 'RMK_DC', 'VAT_DIV_CD', 'ISS_DT', 'TAX_FG', 'SUP_AM', 'CT_NB', 'PJT_CD', 'CT_DEPT', 'CT_DEAL', 'REG_NB', 'CEO_NM', 'DIV_ADDR1', 'TEL'
  $miss = @($codes | Where-Object { -not $colOf.ContainsKey($_) }); if ($miss.Count) { throw "양식에 칸이 없음: $($miss -join ', ')" }
  $n = $LINES.Count; $arr = New-Object 'object[,]' $n, $maxC
  for ($i = 0; $i -lt $n; $i++) { foreach ($cd in $codes) { $val = $LINES[$i][$cd]; if ($null -ne $val -and "$val" -ne '') { $arr[$i, ($colOf[$cd] - 1)] = $(if ($isText[$colOf[$cd]]) { [string]$val } else { $val }) } } }
  foreach ($cd in $codes) { if ($isText[$colOf[$cd]]) { $sh.Range($sh.Cells(4, $colOf[$cd]), $sh.Cells(3 + $n, $colOf[$cd])).NumberFormat = '@' } }
  $sh.Range($sh.Cells(4, 1), $sh.Cells(3 + $n, $maxC)).Value2 = $arr
  foreach ($cd in 'ACCT_AM', 'SUP_AM') { $sh.Range($sh.Cells(4, $colOf[$cd]), $sh.Cells(3 + $n, $colOf[$cd])).NumberFormat = '#,##0' }
  $wb.Save(); $wb.Close($false)
  $wbC = $xl.Workbooks.Open($tmpUp, 0, $true); $lastC = $wbC.Worksheets.Item(1).Cells($wbC.Worksheets.Item(1).Rows.Count, 1).End(-4162).Row; $wbC.Close($false)
  if ($lastC -ne 3 + $n) { throw "업로드 파일 저장 확인 실패(마지막 줄 $lastC, 예상 $(3 + $n))" }
  if (Test-Path -LiteralPath $upload) { Remove-Item -LiteralPath $upload -Force }
  Move-Item -LiteralPath $tmpUp $upload -Force

  # 새 거래처 업로드 엑셀 (아마란스 거래처엑셀업로드 양식: 1줄 이름 / 2줄 칸코드 / 3줄 설명 / 4줄부터 자료)
  # 2026-11-27 사용자가 준 양식(6칸: 회사코드·거래처구분·코드·거래처명·약칭·사업자번호)으로 바꿈.
  # 엑셀로 다시 저장하지 않고 양식 파일에 줄만 더함(새거래처_양식채우기.ps1). 양식에 없는 칸(대표자·주소·전화 등)은 넣지 않음.
  $newUp = Join-Path $OutDir "지출결의_새거래처_$stamp.xlsx"
  if (Test-Path -LiteralPath $newUp) { $bkD = Join-Path $OutDir '백업'; New-Item -ItemType Directory -Force $bkD | Out-Null; Move-Item -LiteralPath $newUp (Join-Path $bkD "지출결의_새거래처_$($stamp)_$(Get-Date -Format 'HHmmss').xlsx") -Force }
  if ($TRADD.Count) {
    $recT = @(foreach ($t0 in $TRADD.Values) { @{ CO_CD = '1000'; TR_FG = '1'; TR_CD = $t0.코드; TR_NM = $t0.이름; ATTR_NM = $t0.이름; CEO_NM = $t0.대표자; REG_NB = $t0.사업자번호; FOR_YN = '0'; DIV_ADDR1 = $t0.주소; TEL = $t0.전화; USE_YN = '1' } })
    & (Join-Path $PSScriptRoot '새거래처_양식채우기.ps1') -Template $TradeTemplate -Out $newUp -Records $recT
  }

  # 검토 엑셀
  $wb = $xl.Workbooks.Add(); while ($wb.Worksheets.Count -lt 3) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  function Fill($s, $name, $hdr, $rows) {
    $s.Name = $name; $nr = @($rows).Count; $dd = New-Object 'object[,]' ($nr + 1), $hdr.Count
    for ($c = 0; $c -lt $hdr.Count; $c++) { $dd[0, $c] = $hdr[$c] }
    $r = 1; foreach ($row in $rows) { for ($c = 0; $c -lt $hdr.Count; $c++) { $x = $row.($hdr[$c]); $dd[$r, $c] = $(if ($x -is [datetime]) { $x.ToString('yyyy-MM-dd') } elseif ($x -is [bool]) { $(if ($x) { 'O' } else { '' }) } else { $x }) }; $r++ }
    $s.Range($s.Cells(1, 1), $s.Cells($nr + 1, $hdr.Count)).Value2 = $dd
    $h = $s.Range($s.Cells(1, 1), $s.Cells(1, $hdr.Count)); $h.Font.Bold = $true; $h.Interior.Color = 0xF2E6DD
    [void]$s.Columns.AutoFit()
  }
  $s1rows = $VOU | ForEach-Object { [pscustomobject]@{ 날짜 = $_.날짜; 작성번호 = $_.작성번호; 문서 = $_.문서; 용도 = $_.용도; 내용 = $_.내용; 가맹점 = $_.가맹점; 사업자번호 = $_.사업자번호; 과세유형 = $_.과세유형; 거래처 = "[$($_.거래처코드)] $($_.거래처명)"; 새거래처 = $_.새거래처; 계정 = "[$($_.계정)] $($ACCNM[$_.계정])"; 프로젝트 = $_.프로젝트; 사용부서 = $(if ($_.사용부서) { "[$($_.사용부서)] $($_.사용부서명)" } elseif ($_.사용부서명) { "(코드 없음) $($_.사용부서명)" } else { '' }); 공급가액 = $_.공급가액; 부가세 = $_.부가세; 합계 = $_.합계; 부가세대급금 = $_.부가세줄; 카드 = "$($_.카드사)$($_.끝4) → [$($CARD[$_.카드사][0])] $($CARD[$_.카드사][1])"; 승인번호 = $_.승인번호; 적요 = $_.적요 } }
  Fill $wb.Worksheets.Item(1) '전표(카드 1건=1장)' @('날짜', '작성번호', '문서', '용도', '내용', '가맹점', '사업자번호', '과세유형', '거래처', '새거래처', '계정', '프로젝트', '사용부서', '공급가액', '부가세', '합계', '부가세대급금', '카드', '승인번호', '적요') $s1rows
  $wb.Worksheets.Item(1).Range('N:P').NumberFormat = '#,##0'
  if ($CHK.Count) { Fill $wb.Worksheets.Item(2) '확인할 것' @('문서', '내용', '확인') $CHK } else { $wb.Worksheets.Item(2).Name = '확인할 것'; $wb.Worksheets.Item(2).Range('A1').Value2 = '확인할 것 없음' }
  if ($SKIP.Count) { Fill $wb.Worksheets.Item(3) '뺀 것(이미 전표 있음)' @('문서', '내용', '날짜', '금액', '카드', '사유') $SKIP; $wb.Worksheets.Item(3).Range('D:D').NumberFormat = '#,##0' } else { $wb.Worksheets.Item(3).Name = '뺀 것'; $wb.Worksheets.Item(3).Range('A1').Value2 = '없음' }
  $wb.Worksheets.Item(1).Activate()
  if (Test-Path -LiteralPath $review) { Remove-Item -LiteralPath $review -Force }
  $wb.SaveAs($review, 51); $wb.Close($false)
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }

"전표 $($VOU.Count)장 ($($LINES.Count)줄) / 부가세대급금 $(@($VOU | Where-Object 부가세줄).Count)장 / 뺀 것 $($SKIP.Count)건 / 확인할 것 $($CHK.Count)건"
"차변 합계 {0:N0} / 대변 합계 {1:N0}" -f (($LINES | Where-Object { $_.DRCR_FG -eq '3' } | ForEach-Object { $_.ACCT_AM } | Measure-Object -Sum).Sum), (($LINES | Where-Object { $_.DRCR_FG -eq '4' } | ForEach-Object { $_.ACCT_AM } | Measure-Object -Sum).Sum)
"업로드: $upload"; "검토: $review"
if ($TRADD.Count) { "새 거래처 $($TRADD.Count)곳 (코드 $(@($TRADD.Values)[0].코드)~$(@($TRADD.Values)[-1].코드)): $newUp  ← 전표보다 먼저 아마란스 거래처엑셀업로드로 등록" }
