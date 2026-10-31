param([string]$XlsPath = "C:\ai\기초자료\202609\mes9월.xls", [string]$Out = "C:\ai\기초자료\202609\ekfmsrj.html")
$ErrorActionPreference = 'Stop'
function HE($s) { [System.Net.WebUtility]::HtmlEncode([string]$s) }
function Fmt($v, $isDate, $hdr) {
  if ($v -eq $null) { return '' }
  if ($v -is [double]) {
    if ($isDate -and $v -gt 30000 -and $v -lt 60000) { return [DateTime]::FromOADate($v).ToString('yyyy-MM-dd') }
    if ($rateCols -contains $hdr) { return $(if ([math]::Abs($v) -lt 50) { $v.ToString('#,##0.0000') } else { $v.ToString('#,##0.00') }) }
    if ($hdr -match '달러|외화') { return $v.ToString('#,##0.00') }
    if ($unit.ContainsKey($hdr)) { return $(if ([math]::Abs($v - [math]::Round($v)) -lt 1e-9) { $v.ToString('#,##0') } else { $v.ToString('#,##0.##') }) }
    if ([math]::Abs($v - [math]::Round($v)) -lt 1e-9) { return ([math]::Round($v)).ToString('0') }
    return ([math]::Round($v, 4)).ToString('0.####')
  }
  return ([string]$v).Trim()
}
# 금액·수량 열(쉼표) / 단위 / 좁게 줄바꿈할 메모 열 — 품번·인보이스·NO 같은 번호 열에는 쉼표를 넣지 않음
$cats6 = '사형', '자동', '금형 주조', '가공1(2공장)', '가공2(5공장)', '가공3(3공장)', '스크랩(사형)', '스크랩(금형)'
$unit = @{ '출고수량' = 'EA'; 'MES 출고수량' = 'EA'; '주조비' = '원'; '가공비' = '원'; '단가' = '원'; '금액' = '원'; 'MES 금액' = '원'; '원화금액(회계)' = '원'
  '환율' = '원'; '기초자료환율' = '원'; '달러' = 'USD·엔'; '기초자료달러' = 'USD·엔'; '외화금액' = 'USD·엔'
  '환율(회계)' = '원'; '외화금액(회계)' = 'USD·엔'; '건수' = '행'
  '전자세금계산서 금액' = '원'; '공급가액(세금계산서)' = '원'; '세액' = '원'; '합계금액' = '원'; '차액(공급가액-MES)' = '원'; '세금계산서 매수' = '매'
  '출하기준 금액(MES)' = '원'; '확정 금액(회계)' = '원'; '원화 차액' = '원'; '출하기준 외화' = 'USD·엔'; '확정 외화(회계)' = 'USD·엔'; '외화 차액' = 'USD·엔' }
foreach ($c in $cats6) { $unit[$c] = '원' }
$rateCols = '환율', '기초자료환율', '환율(회계)'
$centerCols = 'NO', '출고일', 'BL날짜', '인보이스 번호', '수출면장 번호', '구분', '품목구분', '품번', '라인구분', '가공구분', '출고Lot', '담당자', '통화', '분류', '인보이스/Lot', '고객사', '수출면장 번호(회계)', 'BL날짜(회계)', '고객사(MES)', '사업자번호', '작성일', '세금계산서 일련번호'
$memoCols ='메모', '에러내용', '사유', '비교 출처', '출처 기준처', '품명', '기준값(회계·MES)', '비교값', 'MES 행', '항목', '이월 사유', '참고'
$sheets = @(
  @{ Key = 'mes'; Name = 'mes9월';     Title = 'MES 출고 데이터'; RedCols = 'BL날짜','인보이스 번호','환율','기초자료환율','달러','기초자료달러','수출면장 번호','출고Lot'
     Desc = 'MES 9월 출고 1행 = 1품목(맨 아래 TOTAL 제외). 출고Lot = 인보이스 번호. 빨간 칸 = 회계 기준과 다르거나 자료 없음(사유는 오류목록). 기초자료환율·기초자료달러 = 9월라인/PDF 값(달러는 MES 금액 비율로 행에 배분, 신코·코마츠는 엔화).' }
  @{ Key = 'sum'; Name = '해외매출'; Title = '해외매출'; RedCols = '*'
     Desc = '9월 해외매출(옛 면장별합계). 회계 면장 중 회색(다음 달 이월) 제외. 신코·코마츠는 PDF 매출내역서 기준(스크랩 별도 칸), 그 외는 MES, MES에 없는 전월 출고분은 9월라인 참고. 빨간 칸 = 회계와 다름(에러내용 열 참고).' }
  @{ Key = 'dom'; Name = '국내매출'; Title = '국내매출'; RedCols = '*'
     Desc = '9월 국내매출, 고객사별 1줄 (신코·코마츠는 해외매출에 있음). ① MES 출고 ↔ 일별마감 9월 출하: 반품·재출고로 MES 수량·금액이 달라졌는지 (다르면 빨간 칸). ② 9월 매출 = 일별마감 9월 출하 + 전월 출하분 + 반품·불량(-) + 단가 소급 (10월 이월은 빼서 이월매출 시트로). ③ 세금계산서 공급가액 ↔ 9월 매출 (다르면 빨간 칸). MES와 확정 매출 차이는 출하금액변동 시트에 이유별로.' }
  @{ Key = 'taxerr'; Name = '전자세금계산서 금액 오류'; Title = '전자세금계산서 금액 오류'; RedCols = '*'
     Desc = '국내매출(9월 매출, 일별마감 기준) 금액과 아마란스 전자세금계산서 공급가액이 다른 고객사, 또는 한쪽에만 있는 고객사. 세금계산서 파일이 없으면 안내 문구만 보임.' }
  @{ Key = 'carry'; Name = '이월매출'; Title = '이월매출'; RedCols = $null
     Desc = '10월로 넘어가는 매출. 회계 BL일 10/01(BL 미확정·9월라인에 없음) 또는 회계 면장번호 회색. 인보이스 1건 = 1줄, 금액은 회계 기준. 정상 건이라 오류목록에서는 뺐음.' }
  @{ Key = 'ship'; Name = '출하금액변동'; Title = '출하금액변동'; RedCols = $null
     Desc = '출하기준 금액변동 = MES 출하 기준 금액과 확정 금액(회계·인보이스) 차이. MES는 내부 고정환율이라 생기는 정상 차이로, 에러가 아님. 인보이스 1건 = 1줄. 국내는 전자세금계산서 자료를 받으면 구분=국내로 추가.' }
  @{ Key = 'err'; Name = '오류목록';   Title = '오류목록'; RedCols = $null
     Desc = "동일 인보이스는 1줄로 모음(건수 = 원래 행 수, 사유는 항목별). 에러 = 고쳐야 할 것 / 참고 = 내부 고정환율·단수 차이 등 정상 범위. 해결한 항목은 왼쪽 체크 후 '체크한 항목 삭제'(이 브라우저에 기억됨, '되살리기'로 복구)." })
$readme = @(
  @('자료', '원본 = MES(mes9월.xls). 비교 기준 = 회계 2026년 수출 (2026.09).xlsx (26.09(NEW) 시트)'),
  @('BL날짜', '해외 = 9월라인 bl 날짜 / 국내 = 출하일(MES 출고일) / 신코(SHINKO)·코마츠(JS TECH) = PDF 매출내역서. 회계와 다르거나 없으면 빨간색'),
  @('인보이스 번호', '해외 = 9월라인 인보이스 / 국내 = 기재 안 함(오류 표시 없음) / 신코·코마츠 = PDF'),
  @('환율', 'smbs.biz 매매기준율(BL날짜 기준, JPY는 1엔당). Bergen은 회계 1,350 고정'),
  @('기초자료환율·달러', '9월라인 / PDF 매출내역서 값. 회계 환율·외화금액과 비교'),
  @('회색(회계)', '회계 파일 면장번호 칸이 회색이면 무조건 다음 달 이월 → 해외매출에서 제외, 이월매출 시트에 표시'),
  @('이월매출', 'BL 미확정·9월라인에 없음(회계 BL일 10/01) 또는 회계 회색 면장 → 10월 매출. 인보이스별 1줄'),
  @('신코·코마츠', '매월 기초자료의 "N월신코, 코마츠 …xlsx" 엑셀에서 탭이 색칠된 시트를 자동으로 읽음(없으면 PDF 값). 인보이스 앞 "10-"처럼 숫자가 붙으면 그 달(마감 월보다 크면 이월). 엑셀에 없고 회계가 회색이면 이월'),
  @('알림', '에러는 아니지만 매달 꼭 확인할 것 (노란 줄). 예: 매출내역서 엑셀 엔화에 소수점이 있어 회계(반올림 금액)와 원화가 몇 원 다름 → 회계가 맞음'),
  @('출고Lot 오기', 'MES DSM2608-AL6 = DSM2608-AL06(8월 출고·10월 이월)로 인식. 확인 완료라 에러 표시 안 함(이월매출 참고 열에 기록)'),
  @('국내매출', 'MES 출고 · 일별마감 · 아마란스 전자세금계산서(매출) 3가지를 고객사별로 비교. 기초자료에 "N월세금계산서.xlsx"를 넣으면 세금계산서도 자동 비교. 수정세금계산서는 원본과 합친 금액'),
  @('전자세금계산서 금액(MES)', 'mes9월 시트 금액 옆 칸. 같은 거래처의 세금계산서를 작성일별로 묶고, MES 행은 출고일 이후 가장 가까운 작성일 묶음에 넣어 MES 금액 비율로 나눔. 맨 끝 칸에 그 세금계산서 일련번호(승인번호)'),
  @('일별마감 비고',"'10월이월' = 9월 매출에서 빼고 이월매출 시트로 / '불량반출' = 매출 아님, 참고 기록 / 마이너스 줄 = 반품·불량 차감 / '소급' = 단가 소급 / 날짜·품번 없는 소계 줄은 뺌"),
  @('수정세금계산서', '작성일 기준으로 월을 나눔. 작성일이 이번 달이면 다음 달에 발급돼도 이번 달에 합쳐 비교(표시함). 작성일이 전월이면 전월 정정분, 다음 달이면 다음 달 매출 → 비교에서 빼고 오류목록 참고로'),
  @('출하금액변동','MES 출하 기준 금액 ≠ 확정 금액(회계·인보이스, 국내는 전자세금계산서) → 에러가 아니라 출하기준 금액변동. 출하금액변동 시트에 모음'),
  @('Replacement','메모에 Replacement가 있는 행 = 교체품. 금액·달러·기초자료달러 0이 정상이라 에러 아님(빨간색 없음)'),
  @('불량대체','국내 MES 금액이 비어 있는 행 = 불량 대체품(국내 영업팀 확인). 정상이라 에러 아님 → mes 시트 전자세금계산서 금액 칸에 "불량대체". 일별마감 금액 칸에 "불량대체"라고 적힌 줄도 매출 아님(금액은 비고에 적힘)'),
  @('MES 금액','MES는 내부 고정환율(약 1,350원/USD) 기준이라 BL일 환율로 나눈 달러와 0.4~2.5% 차이는 정상(참고)'))

$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false
# 종류별 보기 체크박스: 줄 내용으로 종류를 붙임 (불량대체·이월 등)
$tagDefs = @(@('이월', '이월'), @('불량대체', '불량대체'), @('반품·불량', '반품·불량|반품'), @('단가 소급', '소급'), @('전월 출하', '전월 출하'), @('수정세금계산서', '수정세금계산서'), @('출하기준 금액변동', '출하기준 금액변동'))
$tagOrder = @('에러', '빨간 칸', '알림', '참고') + @($tagDefs | % { $_[0] }); $tagCnt = @{}
$chkCols = @('불량', '이월', '반품', '관세', '툴링비')   # mes 시트 체크 칸 (엑셀 ☐/☑ → HTML 체크박스)
$tables = @{}; $counts = @{}; $stats = @{}; $notes = @{}   # $notes: 표가 아니라 설명으로 보여 줄 [결과]·[안내] 줄
try {
  $wb = $xl.Workbooks.Open($XlsPath, 0, $true)
  foreach ($s in $sheets) {
    $ws = $wb.Worksheets.Item($s.Name); $u = $ws.UsedRange; $v = $u.Value2; $nr = $u.Rows.Count; $nc = $u.Columns.Count
    if ($s.Key -eq 'err') { $nc = 10 }   # 오른쪽 요약표 제외
    $hd = @(for ($j = 1; $j -le $nc; $j++) { ([string]$v[1,$j]) -replace '^[\s\u00A0]+|[\s\u00A0]+$', '' })
    foreach ($h0 in $hd) { if (-not $unit.ContainsKey($h0) -and $s.Key -in 'dom', 'carry', 'taxerr' -and $h0 -notmatch '사유|행|번호|거래처|작성일|출처|에러|고객사|구분') { if ($h0 -match '수량') { $unit[$h0] = 'EA' } elseif ($h0 -match '금액|매출|소급|반품|공급가액|차이|이월') { $unit[$h0] = '원' } } }
    $isDate = @($hd | % { $_ -match '날짜|출고일' })
    $last = $nr; if ("$($v[$nr,1])".Trim() -in 'TOTAL', '합계') { $last = $nr - 1 }
    $redIdx = if ($s.RedCols -eq '*') { 1..$nc } elseif ($s.RedCols) { @(for ($j = 1; $j -le $nc; $j++) { if ($s.RedCols -contains $hd[$j-1]) { $j } }) } else { @() }
    $t = New-Object System.Text.StringBuilder; $need = 0
    [void]$t.Append("<table id=`"t-$($s.Key)`"><thead><tr>")
    if ($s.Key -eq 'err') { [void]$t.Append("<th scope=`"col`">해결</th>") }
   foreach ($h in $hd) { $lab = if ($unit[$h]) { "$h ($($unit[$h]))" } else { $h }; [void]$t.Append("<th scope=`"col`">$(HE $lab)</th>") }
    [void]$t.Append("</tr></thead><tbody>")
    $skipN = 0; $tagCnt[$s.Key] = @{}
    $guIdx = [array]::IndexOf($hd, '구분') + 1; $prevGu = $null; $guCnt = @{}   # MES 탭: 국내(위)·해외(아래) 제목 줄
    if ($s.Key -eq 'mes' -and $guIdx -gt 0) { for ($q = 2; $q -le $last; $q++) { $guCnt["$($v[$q,$guIdx])".Trim()]++ } }
    for ($i = 2; $i -le $last; $i++) {
      if ("$($v[$i,1])" -match '^\[(결과|안내)\]' -or ($s.Key -eq 'taxerr' -and -not ("$(@(1..$nc | % { $v[$i,$_] }) -join '')".Trim()))) { if ("$($v[$i,1])".Trim()) { $notes[$s.Key] += @("$($v[$i,1])".Trim()) }; $skipN++; continue }
      if ($s.Key -eq 'mes' -and $guIdx -gt 0) { $gu0 = "$($v[$i,$guIdx])".Trim(); if ($gu0 -ne $prevGu) { $prevGu = $gu0; [void]$t.Append("<tr class=`"sep`"><td colspan=`"$nc`">▼ $(HE $gu0) ($($guCnt[$gu0])줄 · 품번 → 출고일 순)</td></tr>") } }
      $red = @{}; foreach ($j in $redIdx) { if ([int]$ws.Cells($i, $j).Interior.Color -eq 255) { $red[$j] = $true } }
      $cls = ''
      if ($s.Key -eq 'err') { $k = "$($v[$i,1])"; $cls = if ($k -eq '에러') { 'need' } elseif ($k -eq '알림') { 'warn' } else { '' } }
      elseif ($red.Count) { $cls = 'need' }
      if ($cls -eq 'need') { $need++ }
      $trA = if ($cls) { ' class="' + $cls + '"' } else { '' }
      $rt = (@(1..$nc | % { "$($v[$i,$_])" }) -join ' '); $tg = @()
      if ($cls -eq 'need') { $tg += $(if ($s.Key -eq 'err') { '에러' } else { '빨간 칸' }) } elseif ($cls -eq 'warn') { $tg += '알림' }; if ($s.Key -eq 'err' -and "$($v[$i,1])" -eq '참고') { $tg += '참고' }
      foreach ($td0 in $tagDefs) { if ($rt -match $td0[1]) { $tg += $td0[0] } }
      $tg = @($tg | select -Unique); foreach ($x0 in $tg) { $tagCnt[$s.Key][$x0]++ }; $trA += ' data-tags="' + (HE ($tg -join ',')) + '"'
      if ($s.Key -eq 'err') {
        $rid = "$($v[$i,2])|$($v[$i,4])|$($v[$i,5])"   # 인보이스/Lot|항목|MES 행 → 삭제 기억용 키
        $trA += ' data-id="' + (HE $rid) + '"'
        [void]$t.Append("<tr$trA><td class=`"c`"><input type=`"checkbox`" class=`"done`" aria-label=`"해결`"></td>")
      } else { [void]$t.Append("<tr$trA>") }
      for ($j = 1; $j -le $nc; $j++) {
        $hn = $hd[$j-1]; $c = Fmt $v[$i,$j] $isDate[$j-1] $hn
        $cc = @(); if ($red[$j]) { $cc += 'st-bad' }; if ($unit[$hn] -and $c -match '^-?[\d,]+(\.\d+)?$') { $cc += 'num' }; if ($hn -eq '세금계산서 일련번호') { $cc += 'ser' } elseif ($memoCols -contains $hn) { $cc += 'memo' } elseif ($centerCols -contains $hn) { $cc += 'c' }
        if ($s.Key -eq 'err' -and $j -eq 1 -and $c -eq '에러') { $cc += 'st-bad' }
        if ($s.Key -eq 'mes' -and $chkCols -contains $hn) { $mkId = "$($v[$i,1])|$hn"; [void]$t.Append("<td class=`"c`"><input type=`"checkbox`" class=`"mk`" data-k=`"$(HE $mkId)`" aria-label=`"$hn`"$(if ($c -in '☑', '1') { ' checked' })></td>"); continue }   # 사람이 직접 체크하는 칸 (엑셀 1 = ☑)
        $tdA = if ($cc) { ' class="' + ($cc -join ' ') + '"' } else { '' }
        [void]$t.Append("<td$tdA>$(HE $c)</td>")
      }
      [void]$t.Append("</tr>")
    }
    [void]$t.Append("</tbody></table>")
    $tables[$s.Key] = $t.ToString(); $counts[$s.Key] = $last - 1 - $skipN; $stats[$s.Key] = $need
    if ($s.Key -eq 'sum') { $ci = [array]::IndexOf($hd, '원화금액(회계)') + 1; $stats['krw'] = [double]$v[$nr,$ci] }
  }
  $wb.Close($false)
} finally { $xl.Quit() }

$gptPrompt = "이 페이지의 'README'를 먼저 읽고, 'MES 출고 데이터'의 빨간 칸(st-bad)과 '오류목록'의 분류 열을 기준으로 분석해줘. 비교 기준은 회계 파일이고, 신코·코마츠 행의 기초자료달러는 엔화 금액이야. '이월매출'(10월로 넘어가는 건)과 '참고'는 정상이니 '에러'만 정리해줘."
$css = @'
:root{--bg:#ffffff;--fg:#1e2a2a;--muted:#5b6b6a;--line:#d5dddc;--head:#0f3d3e;--headfg:#ffffff;--tint:#e6f0ee;--alt:#f6f8f8;--bad:#b42318;--badbg:#fdecea;--carry:#eceff1;--warnbg:#fff4c2}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--bg:#0f1716;--fg:#e4ecea;--muted:#9fb0ae;--line:#2a3837;--head:#1f4a4a;--headfg:#ffffff;--tint:#16302f;--alt:#141f1e;--bad:#ff8a80;--badbg:#3a1714;--carry:#1c2626;--warnbg:#3a3214}}
:root[data-theme="dark"]{--bg:#0f1716;--fg:#e4ecea;--muted:#9fb0ae;--line:#2a3837;--head:#1f4a4a;--headfg:#ffffff;--tint:#16302f;--alt:#141f1e;--bad:#ff8a80;--badbg:#3a1714;--carry:#1c2626;--warnbg:#3a3214}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);font:14px/1.55 "Malgun Gothic","맑은 고딕",system-ui,sans-serif}
header,main{max-width:1400px;margin:0 auto;padding:0 16px}header{padding-top:28px}
h1{font-size:24px;margin:0 0 4px}h2{font-size:18px;margin:28px 0 6px}.sub{color:var(--muted);margin:0 0 16px}
.stats{display:flex;flex-wrap:wrap;gap:12px;margin:8px 0 4px}.stat{background:var(--tint);border-radius:10px;padding:10px 14px;min-width:150px}
.stat b{display:block;font-size:20px}.stat span{color:var(--muted);font-size:12px}.stat.bad b{color:var(--bad)}
.prompt{background:var(--tint);border-radius:10px;padding:12px 14px;margin:12px 0}.prompt p{margin:0 0 8px}
button{font:inherit;cursor:pointer;border:1px solid var(--line);background:var(--bg);color:var(--fg);border-radius:8px;padding:6px 12px}
button[aria-selected="true"]{background:var(--head);color:var(--headfg);border-color:var(--head)}
nav.tabs{display:flex;flex-wrap:wrap;gap:6px;margin:20px 0 4px;position:sticky;top:0;background:var(--bg);padding:8px 0;z-index:3}
.tools{display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin:8px 0}.tools input[type=search]{font:inherit;padding:6px 10px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--fg);min-width:240px;max-width:100%}
.count{color:var(--muted);font-size:12px}
.wrap{overflow:auto;max-height:70vh;border:1px solid var(--line);border-radius:8px}
table{border-collapse:collapse;width:max-content;min-width:100%;font-size:13px}
th,td{border-bottom:1px solid var(--line);padding:5px 8px;white-space:nowrap;text-align:left;vertical-align:top}
thead th{position:sticky;top:0;background:var(--head);color:var(--headfg);z-index:2;text-align:center;vertical-align:middle}
td{vertical-align:middle}td.c{text-align:center;font-variant-numeric:tabular-nums}
th,td{border-right:1px solid var(--line)}tbody tr:hover{background:var(--tint)}
tbody tr:nth-child(even){background:var(--alt)}tbody tr.warn,tbody tr.warn:nth-child(even){background:var(--warnbg)}tbody tr.carry{background:var(--carry);color:var(--muted)}
td.num{text-align:right;font-variant-numeric:tabular-nums}td.rn{color:var(--muted);text-align:right}
td.st-bad{color:var(--bad);background:var(--badbg);font-weight:600}
td.memo{white-space:normal;min-width:140px;max-width:260px;font-size:12px;line-height:1.4;word-break:keep-all;overflow-wrap:anywhere}
#readme table td:first-child{font-weight:600;white-space:nowrap}#readme table td:last-child{white-space:normal;min-width:320px}
section[hidden]{display:none}td.ser{white-space:pre-line;font-size:12px;font-variant-numeric:tabular-nums;text-align:center;min-width:200px}.upload{background:var(--tint);border-radius:10px;padding:12px 14px;margin:12px 0}.upload p{margin:0 0 8px}a.btn{display:inline-block;padding:7px 14px;border-radius:8px;background:var(--head);color:var(--headfg);text-decoration:none;margin:4px 8px 0 0;font-weight:600}a.btn.ghost{background:transparent;color:var(--fg);border:1px solid var(--line);font-weight:400}p.note{background:var(--tint);border-radius:8px;padding:8px 12px;margin:6px 0}
.tagbar{flex-basis:100%;display:flex;flex-wrap:wrap;gap:6px 12px;align-items:center;padding:6px 10px;background:var(--tint);border-radius:8px}.tagbar label{white-space:nowrap;cursor:pointer}.tagn{color:var(--muted);font-size:12px}.taghint{color:var(--muted);font-size:12px}tbody tr.sep td{background:var(--head);color:var(--headfg);font-weight:700;padding:7px 10px;position:sticky;left:0}tbody tr.okd td{background:var(--carry);color:var(--muted)}tbody tr.okd td.st-bad{background:var(--carry)}input.ok,input.mk{width:16px;height:16px;cursor:pointer}.okcnt{color:var(--muted);font-size:12px}tbody tr.checked td{text-decoration:line-through;color:var(--muted)}tbody tr.checked td.st-bad{background:transparent}
input.done{width:16px;height:16px;cursor:pointer}.gonecnt{color:var(--muted);font-size:12px}
button.del{background:var(--bad);color:#fff;border-color:var(--bad)}button:disabled{opacity:.5;cursor:default}
@media print{nav.tabs,.tools{display:none}section[hidden]{display:block}.wrap{max-height:none;overflow:visible}}
'@
$js = @'
const tabs=[...document.querySelectorAll('nav.tabs button')];
function show(id){tabs.forEach(b=>b.setAttribute('aria-selected',b.dataset.t===id));document.querySelectorAll('main>section.sheet').forEach(s=>s.hidden=s.id!==id);try{localStorage.setItem('tab1028',id)}catch(e){}}
tabs.forEach(b=>b.addEventListener('click',()=>show(b.dataset.t)));
let start='readme';try{start=localStorage.getItem('tab1028')||'readme'}catch(e){}if(!document.getElementById(start))start='readme';show(start);
// 오류목록: 해결 체크 → '체크한 항목 삭제'로 숨김(이 브라우저에 기억), '되살리기'로 복구
const GKEY='resolved1028';let gone=[];try{gone=JSON.parse(localStorage.getItem(GKEY)||'[]')}catch(e){gone=[]}
function saveGone(){try{localStorage.setItem(GKEY,JSON.stringify(gone))}catch(e){}}
// 모든 탭: 줄마다 '확인' 체크 → 회색, 이 브라우저에 기억. '확인한 줄 숨기기'로 숨김
const OKEY='checked_ok';let okSet={};try{okSet=JSON.parse(localStorage.getItem(OKEY)||'{}')}catch(e){okSet={}}
function saveOk(){try{localStorage.setItem(OKEY,JSON.stringify(okSet))}catch(e){}}
document.querySelectorAll('tr[data-ok]').forEach(r=>{const cb=r.querySelector('input.ok');if(okSet[r.dataset.ok]){cb.checked=true;r.classList.add('okd')}
  cb.addEventListener('change',()=>{r.classList.toggle('okd',cb.checked);if(cb.checked)okSet[r.dataset.ok]=1;else delete okSet[r.dataset.ok];saveOk();applies.forEach(f=>f())})});
// MES 탭 체크 칸(불량·이월·반품·관세·툴링비): 누르면 이 브라우저에 기억 (엑셀 파일에는 저장되지 않음)
const MKEY='mes_checks';let mk={};try{mk=JSON.parse(localStorage.getItem(MKEY)||'{}')}catch(e){mk={}}
document.querySelectorAll('input.mk').forEach(cb=>{const k=cb.dataset.k;if(k in mk)cb.checked=!!mk[k];cb.addEventListener('change',()=>{mk[k]=cb.checked?1:0;try{localStorage.setItem(MKEY,JSON.stringify(mk))}catch(e){}})});
// '체크 저장': 엑셀과 달라진 체크만 체크_YYYYMM.json 으로 저장 → 월 폴더에 두면 다음 반영 때 엑셀에 합쳐짐
const sv=document.getElementById('mk-save');
if(sv){sv.addEventListener('click',async()=>{
  const o={};let n=0;document.querySelectorAll('input.mk').forEach(cb=>{if(cb.checked!==cb.defaultChecked){o[cb.dataset.k]=cb.checked?1:0;n++}});
  const msg=document.getElementById('mk-msg');if(!n){msg.textContent='엑셀과 달라진 체크가 없습니다.';return}
  const txt=JSON.stringify(o,null,1),name='체크_'+sv.dataset.ym+'.json';
  try{if(window.showSaveFilePicker){const fh=await showSaveFilePicker({suggestedName:name,types:[{description:'체크 파일',accept:{'application/json':['.json']}}]});const w=await fh.createWritable();await w.write(txt);await w.close();msg.textContent=n+'칸 저장함. 이 파일이 '+sv.dataset.dir+' 폴더에 있으면 다음 반영 때 엑셀에 들어갑니다.';return}}catch(e){if(e.name==='AbortError')return}
  const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([txt],{type:'application/json'}));a.download=name;a.click();
  msg.textContent=n+'칸 → 내려받기 폴더에 '+name+' 저장됨. '+sv.dataset.dir+' 폴더로 옮겨 주세요.'})}
const applies=[];
document.querySelectorAll('.tools').forEach(tool=>{
  const tb=document.getElementById(tool.dataset.table),q=tool.querySelector('input[type=search]'),only=tool.querySelector('input.only'),out=tool.querySelector('.count'),hideok=tool.querySelector('input.hideok'),okc=tool.querySelector('.okcnt'),tagfs=[...tool.querySelectorAll('input.tagf')];
  const rows=[...tb.tBodies[0].rows].filter(r=>!r.classList.contains('sep'));
  function apply(){const k=q.value.trim().toLowerCase();let n=0;rows.forEach(r=>{const hit=!gone.includes(r.dataset.id)&&(!k||r.textContent.toLowerCase().includes(k))&&(!only||!only.checked||r.classList.contains('need'))&&!(hideok&&hideok.checked&&r.classList.contains('okd'))&&(()=>{const sel=tagfs.filter(x=>x.checked).map(x=>x.value);if(!sel.length)return true;const tg=(r.dataset.tags||'').split(',');return sel.some(x=>tg.includes(x))})();r.hidden=!hit;if(hit)n++});out.textContent=n+' / '+rows.length+'행';if(okc)okc.textContent='확인 '+rows.filter(r=>r.classList.contains('okd')).length+' / '+rows.length}
  q.addEventListener('input',apply);if(only)only.addEventListener('change',apply);if(hideok)hideok.addEventListener('change',apply);tagfs.forEach(x=>x.addEventListener('change',apply));apply();applies.push(apply);
});
const et=document.getElementById('t-err');
if(et){
  const del=document.getElementById('del-done'),res=document.getElementById('restore'),gc=document.getElementById('gonecnt'),left=document.getElementById('errLeft');
  const erows=[...et.tBodies[0].rows];
  function refresh(){applies.forEach(f=>f());const n=erows.filter(r=>gone.includes(r.dataset.id)).length;gc.textContent=n?('삭제(해결)한 항목 '+n+'건'):'';res.disabled=!n;
    del.disabled=!erows.some(r=>!r.hidden&&r.querySelector('input.done').checked);
    if(left)left.textContent=erows.filter(r=>r.classList.contains('need')&&!gone.includes(r.dataset.id)).length+'건'}
  et.addEventListener('change',e=>{if(e.target.matches('input.done')){e.target.closest('tr').classList.toggle('checked',e.target.checked);refresh()}});
  del.addEventListener('click',()=>{const pick=erows.filter(r=>!r.hidden&&r.querySelector('input.done').checked);if(!pick.length)return;
    if(!confirm('체크한 '+pick.length+'건을 해결로 보고 목록에서 삭제할까요? (되살리기 가능)'))return;
    pick.forEach(r=>{if(!gone.includes(r.dataset.id))gone.push(r.dataset.id);r.querySelector('input.done').checked=false;r.classList.remove('checked')});saveGone();refresh()});
  res.addEventListener('click',()=>{gone=[];saveGone();refresh()});
  refresh();
}
document.getElementById('copy').addEventListener('click',async()=>{const t=document.getElementById('gpt').textContent;try{await navigator.clipboard.writeText(t);document.getElementById('copy').textContent='복사됨'}catch(e){}});
'@
$h = New-Object System.Text.StringBuilder
[void]$h.Append("<!doctype html><html lang=`"ko`"><head><meta charset=`"utf-8`"><meta name=`"viewport`" content=`"width=device-width,initial-scale=1`"><title>9월 수출 마감 대사</title><style>$css</style></head><body>")
[void]$h.Append("<header><h1>9월 수출 마감 대사 데이터</h1><p class=`"sub`">대신금속 2026년 9월 · mes9월.xls 기준 (2026-10-28 작업 8~12 반영) · 비교 기준 = 회계 파일</p>")
[void]$h.Append("<div class=`"stats`"><div class=`"stat`"><b>$($counts['mes'])행</b><span>MES 출고 데이터 (빨간 칸 있는 행 $($stats['mes']))</span></div><div class=`"stat`"><b>$($counts['sum'] - 0)건</b><span>해외매출</span></div><div class=`"stat`"><b>$('{0:N0}' -f $stats['krw'])원</b><span>해외매출 원화(회계) 합계</span></div><div class=`"stat`"><b>$($counts['carry'])건</b><span>이월매출 (10월로)</span></div><div class=`"stat bad`"><b id=`"errLeft`">$($stats['err'])건</b><span>남은 에러 (오류목록 $($counts['err'])줄 중)</span></div></div>")
# 자료 올리기: GitHub 업로드 화면(그달·다음 달 기초자료 폴더)으로 가는 버튼. 올린 파일은 10분마다 PC로 자동 동기화
$ymCur = Split-Path (Split-Path $XlsPath -Parent) -Leaf; if ($ymCur -notmatch '^\d{6}$') { $ymCur = '202609' }
$ymNext = [DateTime]::ParseExact($ymCur, 'yyyyMM', $null).AddMonths(1).ToString('yyyyMM')
$ghBase = "https://github.com/a01042102350/date"; $fold = [Uri]::EscapeDataString('기초자료')
[void]$h.Append("<div class=`"upload`"><p><b>자료 올리기</b> — 버튼을 누르면 GitHub 업로드 화면이 열립니다. 파일을 끌어다 놓고 아래 <b>Commit changes</b>를 누르면 10분 안에 PC의 기초자료 폴더로 자동으로 내려옵니다. (GitHub 로그인 필요) 다 올린 뒤 클로드에게 ""$ymNext 반영해줘""라고 하면 됩니다.</p><a class=`"btn`" href=`"$ghBase/upload/main/$fold/$ymNext`" target=`"_blank`" rel=`"noopener`">$ymNext 자료 올리기 (다음 달)</a><a class=`"btn ghost`" href=`"$ghBase/upload/main/$fold/$ymCur`" target=`"_blank`" rel=`"noopener`">$ymCur 자료 추가</a><a class=`"btn ghost`" href=`"$ghBase/tree/main/$fold`" target=`"_blank`" rel=`"noopener`">GitHub 폴더 보기</a></div>")
[void]$h.Append("<div class=`"prompt`"><p><b>GPT에 이 페이지를 줄 때 함께 붙일 문장</b></p><p id=`"gpt`">$(HE $gptPrompt)</p><button id=`"copy`" type=`"button`">문장 복사</button></div></header><main>")
[void]$h.Append("<nav class=`"tabs`" aria-label=`"시트`"><button data-t=`"readme`">README</button>")
foreach ($s in $sheets) { [void]$h.Append("<button data-t=`"s-$($s.Key)`">$(HE $s.Title) ($($counts[$s.Key]))</button>") }
[void]$h.Append("</nav><section class=`"sheet`" id=`"readme`"><h2>README (규칙)</h2><div class=`"wrap`"><table><thead><tr><th scope=`"col`">구분</th><th scope=`"col`">내용</th></tr></thead><tbody>")
foreach ($r in $readme) { [void]$h.Append("<tr><td>$(HE $r[0])</td><td>$(HE $r[1])</td></tr>") }
[void]$h.Append("</tbody></table></div></section>")
foreach ($s in $sheets) {
  $lbl = if ($s.Key -eq 'err') { '에러만 보기' } else { '빨간 칸 있는 행만 보기' }
  $tc0 = $tagCnt[$s.Key]; $tagBox = if ($tc0 -and $tc0.Count) { '<div class="tagbar"><b>보기:</b> ' + ((@($tagOrder | ? { $tc0[$_] } | % { "<label><input type=`"checkbox`" class=`"tagf`" value=`"$_`"> $_ <span class=`"tagn`">$($tc0[$_])</span></label>" })) -join '') + ' <span class="taghint">(아무것도 안 고르면 전체)</span></div>' } else { '' }
  $okBox = if ($s.Key -ne 'err') { '<label><input type="checkbox" class="hideok"> 확인한 줄 숨기기</label><span class="okcnt"></span>' } else { '' }
  $onlyBox = if ($s.Key -in 'carry', 'ship') { '' } else { "<label><input type=`"checkbox`" class=`"only`"> $lbl</label>" }
  $ymH = Split-Path (Split-Path $XlsPath -Parent) -Leaf; $dirH = Split-Path $XlsPath -Parent
  $saveBtn = if ($s.Key -eq 'mes') { "<button type=`"button`" id=`"mk-save`" class=`"del`" data-ym=`"$ymH`" data-dir=`"$(HE $dirH)`">체크 저장</button><span id=`"mk-msg`" class=`"gonecnt`"></span>" } else { '' }
  $errBtns = if ($s.Key -eq 'err') { "<button type=`"button`" id=`"del-done`" class=`"del`">체크한 항목 삭제(해결)</button><button type=`"button`" id=`"restore`">삭제한 항목 되살리기</button><span id=`"gonecnt`" class=`"gonecnt`"></span>" } else { '' }
  [void]$h.Append("<section class=`"sheet`" id=`"s-$($s.Key)`"><h2>$(HE $s.Title)</h2><p class=`"sub`">$(HE $s.Desc)</p>$(($notes[$s.Key] | % { ""<p class=`""note`""><b>$(HE $_)</b></p>"" }) -join '')<div class=`"tools`" data-table=`"t-$($s.Key)`"><input type=`"search`" placeholder=`"검색 (예: DSM2609-SK03, 신코, 에러)`" aria-label=`"$(HE $s.Title) 검색`">$onlyBox$errBtns$saveBtn<span class=`"count`"></span>$tagBox</div><div class=`"wrap`">$($tables[$s.Key])</div></section>")
}
[void]$h.Append("</main><script>$js</script></body></html>")
[IO.File]::WriteAllText($Out, $h.ToString(), (New-Object Text.UTF8Encoding $false))
"saved $Out ($([math]::Round((Get-Item $Out).Length / 1KB)) KB) / " + (($sheets | % { "$($_.Name) $($counts[$_.Key])행(표시 $($stats[$_.Key]))" }) -join ', ') + " / 원화합계 $($stats['krw'])"