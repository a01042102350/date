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
  [int]$StartNo = 3000                                              # 작성번호: 날짜마다 3000부터(손 전표·해외 1000·국내 2000과 안 겹치게)
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
  $lastReg = $master | Where-Object { "$($_.trCd)" -match '^\d{5}$' } | Sort-Object { "$($_.insertDt)" } -Descending | Select-Object -First 1
  if ($lastReg) { $nextCode = [int]$lastReg.trCd + 1; Write-Host "  아마란스 거래처 $($master.Count)곳 / 가장 최근 등록 $($lastReg.trCd) $($lastReg.trNm) ($($lastReg.insertDt)) → 새 코드는 $('{0:D5}' -f $nextCode)부터" }
} else { Write-Host '  아마란스 거래처 목록(오늘 것)이 없어 새 거래처 코드는 비워 둡니다' -ForegroundColor Yellow }
function New-TrCode { while ($codeSet.ContainsKey(('{0:D5}' -f $script:nextCode))) { $script:nextCode++ }; $c0 = '{0:D5}' -f $script:nextCode; $codeSet[$c0] = 1; $script:nextCode++; $c0 }
$NEWTR = [ordered]@{}   # 이번에 코드를 새로 매긴 거래처 (키: 사업자번호|가게이름)
# 이름 맞추기: 괄호·(주)·띄어쓰기·기호를 빼고 비교. '결제대행/가맹점' 모양이면 '/' 뒤 가맹점 이름끼리 비교(앞 결제대행 이름이 같아도 다른 가게)
#  같음 = 100, 한쪽이 다른 쪽을 품음(2글자 이상) = 50+글자수, 겹친 글자가 4글자 이상이면서 짧은 이름의 60% 이상 = 겹친 글자수, 그 밖 = 0
function NormNm($s) { ("$s" -replace '주식회사|\(주\)|㈜|\(재\)|\(사\)', '' -replace '[\s\(\)\[\]\-_/\.,·&]', '').ToLower() }
function ShopPart($s) { $t = "$s"; if ($t.Contains('/')) { $t.Substring($t.LastIndexOf('/') + 1) } else { $t } }
function NameScore($a, $b) {
  $x = NormNm (ShopPart $a); $y = NormNm (ShopPart $b); if (-not $x -or -not $y) { return 0 }
  if ($x -eq $y) { return 100 }
  if ($x.Contains($y) -or $y.Contains($x)) { $m = [Math]::Min($x.Length, $y.Length); if ($m -ge 2) { return 50 + $m } }
  $best = 0; for ($i = 0; $i -lt $x.Length; $i++) { for ($k = $i + $best + 1; $k -le $x.Length; $k++) { if ($y.Contains($x.Substring($i, $k - $i))) { $best = $k - $i } else { break } } }
  if ($best -ge 4 -and $best -ge 0.6 * [Math]::Min($x.Length, $y.Length)) { return $best } else { return 0 }
}
$PGRX = '/|이니시스|KCP|KICC|나이스|NICE|TOSS|토스|카카오페이|네이버페이|인터넷상거래|결제|결재|요금창구|세외수입'   # 결제대행을 거친 가맹점 이름
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
$ACC = @{ '복리후생비' = @('8110000', '5110000'); '여비교통비' = @('8120000', '5120000'); '접대비' = @('8130000', '5130000'); '교육훈련비' = @('8250000', '5250000'); '소모품비' = @('8300000', '5300000'); '지급수수료' = @('8310000', '5310000'); '해외출장비' = @('8340000', '8340000') }
$ACCNM = @{ '8110000' = '복리후생비(판)'; '5110000' = '복리후생비(제)'; '8120000' = '여비교통비(판)'; '5120000' = '여비교통비(제)'; '8130000' = '접대비(판)'; '5130000' = '접대비(제)'; '8250000' = '교육훈련비(판)'; '5250000' = '교육훈련비(제)'; '8300000' = '소모품비(판)'; '5300000' = '소모품비(제)'; '8310000' = '지급수수료(판)'; '5310000' = '지급수수료(제)'; '8340000' = '해외출장비'; '1350000' = '부가세대급금'; '2530000' = '미지급금' }

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
  if ($txt -match '항공' -and $txt -match '해외|스위스|일본|미국|출장') { $cat = '해외출장비' }
  elseif ($txt -match '주차|택시|대리|유류|주유|기름|교통|철도|KTX|SRT|버스|숙박|모텔|호텔|항공|통행') { $cat = '여비교통비' }
  elseif ($txt -match '교육') { $cat = '교육훈련비' }
  elseif ($txt -match '사용료|수수료|구독|CHATGPT|OPENAI') { $cat = '지급수수료' }
  elseif ($txt -match '마이크|구매|구입|비품|소모품') { $cat = '소모품비' }
  elseif ($txt -match '식대|중식|석식|점심|저녁|음료|간식|다과|회식') { $cat = '복리후생비' }
  # 손님(고객) 건: 식대이거나, 용도로 계정을 못 정한 건(예: 용도 '기타' + 내용 '고객업무협의')도 같이 봄 (2026-11-26 사용자 지정: "고객업무 이런 거 나오면 접대, 12만원 넘는 건")
  if ((-not $cat -or $cat -eq '복리후생비') -and "$use $desc" -match '신코|머큐리|고객|접대|거래처|회계법인|법무법인|세무법인|세무사') {
    $how = if ($cat) { '식대' } else { "건(용도 '$use')" }
    if ($total -le 120000) { $cat = '복리후생비'; $why = "손님(고객) $how 이지만 12만원 이하라 복리후생비(2026-11-19 사용자 지정)" }
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
    if ($best) { $trCd = $best[0]; $trNm = $best[1]; $newTr = $false }
  }
  # 새 거래처: 코드를 매겨 전표에 넣고 '새거래처' 업로드 파일에도 같은 코드로 넣음 (같은 가게가 여러 번 나오면 같은 코드)
  if ($newTr -and $null -ne $nextCode -and $biz.Length -eq 10) {
    $nk = "$biz|$(NormNm (ShopPart $rc['name']))"
    if (-not $NEWTR.Contains($nk)) { $nm0 = "$($rc['name'])".Trim(); if ($nm0.Length -gt 60) { $nm0 = $nm0.Substring(0, 60) }; $NEWTR[$nk] = [pscustomobject]@{ 코드 = (New-TrCode); 이름 = $nm0; 사업자번호 = $biz; 대표자 = "$($rc['ceo'])".Trim(); 주소 = "$($rc['addr'])".Trim(); 전화 = "$($rc['tel'])".Trim() } }
    $trCd = $NEWTR[$nk].코드; $trNm = $NEWTR[$nk].이름
  }
  # 대리운전·택시는 적요 '_#' 앞에 '교통비'를 넣음(예: 고객응대 교통비_#하나8508, 2026-11-13 사용자 지정). 이미 '교통비'가 있으면 그대로
  $tail = ''; if (($use -match '대리' -or $desc -match '대리|택시' -or "$($rc['name'])" -match '대리|택시') -and $desc -notmatch '교통비') { $tail = ' 교통비' }
  $suffix = "$tail`_#$co$l4"; $rmk = "$desc".Trim(); if (($rmk + $suffix).Length -gt 80) { $rmk = $rmk.Substring(0, 80 - $suffix.Length).Trim() }; $rmk = $rmk + $suffix
  [void]$VOU.Add([pscustomobject]@{ 문서 = $d['docNo']; 문서제목 = $d['title']; 순번 = $it['no']; 날짜 = $dt; 용도 = $use; 내용 = $desc; 가맹점 = "$($rc['name'])"; 사업자번호 = $biz; 과세유형 = $kindTx; 거래처코드 = $trCd; 거래처명 = $trNm; 새거래처 = $newTr
    대표자 = "$($rc['ceo'])"; 주소 = "$($rc['addr'])"; 전화 = "$($rc['tel'])"; 카드사 = $co; 카드번호 = ("$($rc['cardNo'])" -replace '\D', ''); 끝4 = $l4; 승인번호 = "$($rc['approval'])"
    계정 = $acct; 분류 = $cat; 프로젝트 = $pj; 프로젝트근거 = $pjWhy; 공급가액 = $sup; 부가세 = $vat; 합계 = $total; 부가세줄 = $vatLine; 적요 = $rmk; 참고 = $why })
} }

# 작성번호 (날짜마다 StartNo부터, 문서 순서대로)
$seq = @{}; foreach ($v in $VOU) { $k = $v.날짜.ToString('yyyyMMdd'); if (-not $seq.ContainsKey($k)) { $seq[$k] = $StartNo } else { $seq[$k]++ }; $v | Add-Member 작성번호 $seq[$k] }
# 업로드 줄
$LINES = New-Object Collections.ArrayList
foreach ($v in $VOU) {
  $dts = $v.날짜.ToString('yyyyMMdd'); $ty = $(if ($v.부가세줄) { '2' } else { '1' }); $ln = 0
  $base = @{ IN_DIV_CD = '1000'; MENU_DT = $dts; MENU_SQ = $v.작성번호; ISU_DOC = $v.적요; DOCU_TY = $ty; RMK_DC = $v.적요 }
  $ln++; $e = $base.Clone(); $e.MENU_LN_SQ = $ln; $e.DRCR_FG = '3'; $e.ACCT_CD = $v.계정; $e.TR_CD = $v.거래처코드; $e.TR_NM = $v.거래처명; $e.ACCT_AM = $(if ($v.부가세줄) { $v.공급가액 } else { $v.합계 }); $e.PJT_CD = $v.프로젝트; $e.REG_NB = $v.사업자번호
  if ($v.새거래처) { $e.CEO_NM = $v.대표자; $e.DIV_ADDR1 = $v.주소; $e.TEL = $v.전화 }
  [void]$LINES.Add($e)
  if ($v.부가세줄) { $ln++; $t = $base.Clone(); $t.MENU_LN_SQ = $ln; $t.DRCR_FG = '3'; $t.ACCT_CD = '1350000'; $t.TR_CD = $v.거래처코드; $t.TR_NM = $v.거래처명; $t.ACCT_AM = $v.부가세; $t.REG_NB = $v.사업자번호
    $t.VAT_DIV_CD = '1000'; $t.ISS_DT = $dts; $t.TAX_FG = '27'; $t.CT_DEAL = '27'; $t.SUP_AM = $v.공급가액; $t.CT_NB = $CARD[$v.카드사][0]   # 결제카드 = 언제나 카드사 거래처코드 (기업 99706 / 경남 99709 / 농협 99708 / 현대 99903 / 하나 99905 / 국민 99701) — 2026-11-19 사용자 지정: 카드번호별 코드는 쓰지 않는다
    if ($v.새거래처) { $t.CEO_NM = $v.대표자; $t.DIV_ADDR1 = $v.주소; $t.TEL = $v.전화 }
    [void]$LINES.Add($t) }
  $ln++; $c = $base.Clone(); $c.MENU_LN_SQ = $ln; $c.DRCR_FG = '4'; $c.ACCT_CD = '2530000'; $c.TR_CD = $CARD[$v.카드사][0]; $c.TR_NM = $CARD[$v.카드사][1]; $c.ACCT_AM = $v.합계
  [void]$LINES.Add($c)
}
if ($NEWTR.Count) {
  $cds = @($NEWTR.Values | ForEach-Object { $_.코드 })
  $CHK.Insert(0, [pscustomobject]@{ 문서 = '(전체)'; 내용 = "새 거래처 $($NEWTR.Count)곳 먼저 등록"; 확인 = "전표 올리기 전에 '지출결의_새거래처_$(Get-Date -Format 'yyyyMMdd').xlsx'를 아마란스 회계관리 > 회계기초정보관리 > 거래처/계좌/카드관리 > 거래처엑셀업로드에서 먼저 등록(기능모음 > 파일업로드 → 변환하기). 코드 $($cds[0])~$($cds[-1]). 등록 후 일반거래처등록에서 코드가 파일과 같은지 확인(다르면 전표 파일 코드도 고쳐야 함)" })
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
  $codes = 'IN_DIV_CD', 'MENU_DT', 'MENU_SQ', 'MENU_LN_SQ', 'ISU_DOC', 'DOCU_TY', 'DRCR_FG', 'ACCT_CD', 'TR_CD', 'TR_NM', 'ACCT_AM', 'RMK_DC', 'VAT_DIV_CD', 'ISS_DT', 'TAX_FG', 'SUP_AM', 'CT_NB', 'PJT_CD', 'CT_DEAL', 'REG_NB', 'CEO_NM', 'DIV_ADDR1', 'TEL'
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
  $newUp = Join-Path $OutDir "지출결의_새거래처_$stamp.xlsx"
  if (Test-Path -LiteralPath $newUp) { $bkD = Join-Path $OutDir '백업'; New-Item -ItemType Directory -Force $bkD | Out-Null; Move-Item -LiteralPath $newUp (Join-Path $bkD "지출결의_새거래처_$($stamp)_$(Get-Date -Format 'HHmmss').xlsx") -Force }
  if ($NEWTR.Count) {
    if (-not (Test-Path -LiteralPath $TradeTemplate)) { throw "거래처 업로드 양식이 없습니다: $TradeTemplate" }
    $tmpTr = Join-Path $work '새거래처_임시.xlsx'; if (Test-Path -LiteralPath $tmpTr) { Remove-Item -LiteralPath $tmpTr -Force }
    $wbT = $xl.Workbooks.Open($TradeTemplate, 0, $true); $wbT.SaveAs($tmpTr, 51)
    $shT = $wbT.Worksheets.Item(1); $mxT = $shT.UsedRange.Columns.Count; $colT = @{}
    for ($ci = 1; $ci -le $mxT; $ci++) { $cd = "$($shT.Cells(2, $ci).Value2)".Trim(); if ($cd -and -not $colT.ContainsKey($cd)) { $colT[$cd] = $ci } }
    $needT = 'CO_CD', 'TR_FG', 'TR_CD', 'TR_NM', 'ATTR_NM', 'CEO_NM', 'REG_NB', 'FOR_YN', 'DIV_ADDR1', 'TEL', 'USE_YN'
    $missT = @($needT | Where-Object { -not $colT.ContainsKey($_) }); if ($missT.Count) { throw "거래처 양식에 칸이 없음: $($missT -join ', ')" }
    $nT = $NEWTR.Count; $arrT = New-Object 'object[,]' $nT, $mxT; $i = 0
    foreach ($t0 in $NEWTR.Values) {
      $vals = @{ CO_CD = '1000'; TR_FG = '1'; TR_CD = $t0.코드; TR_NM = $t0.이름; ATTR_NM = $t0.이름; CEO_NM = $t0.대표자; REG_NB = $t0.사업자번호; FOR_YN = '0'; DIV_ADDR1 = $t0.주소; TEL = $t0.전화; USE_YN = '1' }
      foreach ($k in $vals.Keys) { if ("$($vals[$k])") { $arrT[$i, ($colT[$k] - 1)] = [string]$vals[$k] } }; $i++
    }
    $rgT = $shT.Range($shT.Cells(4, 1), $shT.Cells(3 + $nT, $mxT)); $rgT.NumberFormat = '@'; $rgT.Value2 = $arrT
    $wbT.Save(); $wbT.Close($false)
    Move-Item -LiteralPath $tmpTr $newUp -Force
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
  $s1rows = $VOU | ForEach-Object { [pscustomobject]@{ 날짜 = $_.날짜; 작성번호 = $_.작성번호; 문서 = $_.문서; 용도 = $_.용도; 내용 = $_.내용; 가맹점 = $_.가맹점; 사업자번호 = $_.사업자번호; 과세유형 = $_.과세유형; 거래처 = "[$($_.거래처코드)] $($_.거래처명)"; 새거래처 = $_.새거래처; 계정 = "[$($_.계정)] $($ACCNM[$_.계정])"; 프로젝트 = $_.프로젝트; 공급가액 = $_.공급가액; 부가세 = $_.부가세; 합계 = $_.합계; 부가세대급금 = $_.부가세줄; 카드 = "$($_.카드사)$($_.끝4) → [$($CARD[$_.카드사][0])] $($CARD[$_.카드사][1])"; 승인번호 = $_.승인번호; 적요 = $_.적요 } }
  Fill $wb.Worksheets.Item(1) '전표(카드 1건=1장)' @('날짜', '작성번호', '문서', '용도', '내용', '가맹점', '사업자번호', '과세유형', '거래처', '새거래처', '계정', '프로젝트', '공급가액', '부가세', '합계', '부가세대급금', '카드', '승인번호', '적요') $s1rows
  $wb.Worksheets.Item(1).Range('M:O').NumberFormat = '#,##0'
  if ($CHK.Count) { Fill $wb.Worksheets.Item(2) '확인할 것' @('문서', '내용', '확인') $CHK } else { $wb.Worksheets.Item(2).Name = '확인할 것'; $wb.Worksheets.Item(2).Range('A1').Value2 = '확인할 것 없음' }
  if ($SKIP.Count) { Fill $wb.Worksheets.Item(3) '뺀 것(이미 전표 있음)' @('문서', '내용', '날짜', '금액', '카드', '사유') $SKIP; $wb.Worksheets.Item(3).Range('D:D').NumberFormat = '#,##0' } else { $wb.Worksheets.Item(3).Name = '뺀 것'; $wb.Worksheets.Item(3).Range('A1').Value2 = '없음' }
  $wb.Worksheets.Item(1).Activate()
  if (Test-Path -LiteralPath $review) { Remove-Item -LiteralPath $review -Force }
  $wb.SaveAs($review, 51); $wb.Close($false)
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }

"전표 $($VOU.Count)장 ($($LINES.Count)줄) / 부가세대급금 $(@($VOU | Where-Object 부가세줄).Count)장 / 뺀 것 $($SKIP.Count)건 / 확인할 것 $($CHK.Count)건"
"차변 합계 {0:N0} / 대변 합계 {1:N0}" -f (($LINES | Where-Object { $_.DRCR_FG -eq '3' } | ForEach-Object { $_.ACCT_AM } | Measure-Object -Sum).Sum), (($LINES | Where-Object { $_.DRCR_FG -eq '4' } | ForEach-Object { $_.ACCT_AM } | Measure-Object -Sum).Sum)
"업로드: $upload"; "검토: $review"
if ($NEWTR.Count) { "새 거래처 $($NEWTR.Count)곳 (코드 $(@($NEWTR.Values)[0].코드)~$(@($NEWTR.Values)[-1].코드)): $newUp  ← 전표보다 먼저 아마란스 거래처엑셀업로드로 등록" }
