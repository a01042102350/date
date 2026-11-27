# 해외출장 카드 자동전표 (지출결의서 없이): 천문영 팀장 11/9~17 스위스·노르웨이 출장, 경남카드 #7006 + 항공료 2건은 하나카드 #5839(99905)
# 자료: 해외출장여비정산내역서261117.xls(정산서) + 해외 매출표사본출력.pdf(BC카드 해외카드 매출내역, 청구금액 = 수수료 포함 확정금액)
# 사용자 지정(2026-11-27): 거래 1건 = 전표 1장 / 비용은 전부 해외출장비 8340000 / 한국에서 쓴 것도 부가세대급금 없이 전표유형 일반
#   적요 낱말(2026-11-27 사용자): 해외 식대도 전부 석식대, 야근식대 → 석식대, 주차비·전철비 → 교통비(항상), '(도시세)' 빼고 숙박비
#   작성일자: 해외 = 카드사 현지거래일, 국내 = 정산서 날짜 (2026-11-27 사용자)
#   미지급금 = 경남카드 99709 / 비용 줄 거래처는 코드 없이 거래처명에 가맹점명만 / 계정 834(해외출장비, 판관비) = 사용자가 말한 "800번대"
$ErrorActionPreference = 'Stop'
$Template = 'C:\ai\매월마감\해외자동전표\자동전표처리(일반) 양식.xlsx'
$OutDir = "$env:USERPROFILE\Desktop\지출결의서\결과"
$stamp = Get-Date -Format 'yyyyMMdd'
$upload = Join-Path $OutDir "해외출장_카드자동전표_천문영_$stamp.xlsx"
$review = Join-Path $OutDir "해외출장_카드자동전표_천문영_$($stamp)_검토.xlsx"
$prefix = '11/9~17 천문영팀장 스위스 노르웨이 출장'
$cardSfx = '_#경남7006'

# 해외: 정산서 줄(날짜·상호·적요·외화·정산서 원화) ↔ 카드사 매출내역(현지거래일·가맹점·현지금액·통화·미화·환율·수수료·청구금액)
$OVS = @(
  @{ d = '2026-11-10'; nm = 'IBIS 취리히';     item = '숙박비'; sFx = 17.5;    sKrw = 27804.7;     ld = '2026-11-09'; shop = 'IBIS B ZURICH CITY W';    amt = 17.5;    cur = 'CHF'; usd = 21.62;  rate = 1356.9; fee = 72;   bill = 29408 },
  @{ d = '2026-11-10'; nm = 'UBER';           item = '택시비';         sFx = 19.32;   sKrw = 30696.3888;  ld = '2026-11-09'; shop = 'UBER *TRIP';              amt = 19.32;   cur = 'CHF'; usd = 23.88;  rate = 1356.9; fee = 80;   bill = 32482 },
  @{ d = '2026-11-11'; nm = 'IBIS 취리히';     item = '어댑터 대여료';   sFx = 10;      sKrw = 15888.4;     ld = '2026-11-10'; shop = 'Ibis budget Zurich Cit';  amt = 10;      cur = 'CHF'; usd = 12.36;  rate = 1354.3; fee = 41;   bill = 16780 },
  @{ d = '2026-11-11'; nm = 'COOP';           item = '석식대';           sFx = 56.45;   sKrw = 89690.018;   ld = '2026-11-10'; shop = 'Coop-5240 ZH Pfingstwe';  amt = 72.62;   cur = 'USD'; usd = 73.35;  rate = 1354.3; fee = 245;  bill = 99582 },
  @{ d = '2026-11-12'; nm = 'KFC';            item = '석식대';           sFx = 19.9;    sKrw = 31617.916;   ld = '2026-11-11'; shop = 'KFC Welle 7';             amt = 19.9;    cur = 'CHF'; usd = 24.59;  rate = 1347.1; fee = 82;   bill = 33207 },
  @{ d = '2026-11-15'; nm = 'Walkthrough';    item = '석식대';           sFx = 53.5;    sKrw = 85002.94;    ld = '2026-11-14'; shop = 'Walkthrough Level 2 ZR';  amt = 53.5;    cur = 'CHF'; usd = 64.77;  rate = 1347.1; fee = 215;  bill = 87466 },
  @{ d = '2026-11-15'; nm = 'OST';            item = '교통비';         sFx = 51;      sKrw = 7068.09;     ld = '2026-11-14'; shop = 'OST*Skyss';               amt = 51;      cur = 'NOK'; usd = 5.42;   rate = 1347.1; fee = 18;   bill = 7319 },
  @{ d = '2026-11-15'; nm = 'Clarion admiral'; item = '호텔비';        sFx = 2842.38; sKrw = 393925.4442; ld = '2026-11-14'; shop = 'ClarionHotelAdmira';      amt = 2842.38; cur = 'NOK'; usd = 302.26; rate = 1347.1; fee = 1007; bill = 408181 },
  @{ d = '2026-11-15'; nm = 'Kiwi';           item = '석식대';           sFx = 237.5;   sKrw = 32915.125;   ld = '2026-11-14'; shop = '7080001270056 KIWI 826';  amt = 237.5;   cur = 'NOK'; usd = 25.26;  rate = 1347.1; fee = 84;   bill = 34111 },
  @{ d = '2026-11-16'; nm = 'FLIBANEN';       item = '교통비';         sFx = 145;     sKrw = 20095.55;    ld = '2026-11-15'; shop = 'FL IBANEN AS';            amt = 145;     cur = 'NOK'; usd = 15.42;  rate = 1349.7; fee = 51;   bill = 20863 },
  @{ d = '2026-11-17'; nm = 'NOR-WAY BUS';    item = '교통비';         sFx = 220;     sKrw = 30489.8;     ld = '2026-11-16'; shop = 'NOR-WAY Bussekspress';    amt = 220;     cur = 'NOK'; usd = 23.39;  rate = 1348.7; fee = 78;   bill = 31624 },
  @{ d = '2026-11-17'; nm = 'DUTYFREE';       item = '면세점 석식대';     sFx = 759.5;   sKrw = 105259.105;  ld = '2026-11-16'; shop = 'DUTYFREE 7401';           amt = 759.5;   cur = 'NOK'; usd = 80.77;  rate = 1348.7; fee = 269;  bill = 109203 }
)
# 한국에서 쓴 것(정산서 금액 그대로, 부가세대급금 없음)
$DOM = @(
  @{ d = '2026-11-06'; shop = '맘스터치';           item = '석식대'; amt = 40700 },
  @{ d = '2026-11-09'; shop = '에이치씨에너지';      item = '유류비';   amt = 100000 },
  @{ d = '2026-11-09'; shop = '롯데김해공항면세점';  item = '고객선물'; amt = 88550 },
  @{ d = '2026-11-09'; shop = '현대디에프';          item = '고객선물'; amt = 201300 },
  @{ d = '2026-11-17'; shop = '김해공항주차장';      item = '교통비';   amt = 105000 },
  @{ d = '2026-11-18'; shop = '서브웨이';            item = '석식대'; amt = 33600 }
)
# 하나카드 #5839로 산 항공료(정산서 비고 'VISA5839'): 미지급금 하나카드 99905 (2026-11-27 사용자 "5839카드 맞네, 하나카드 99905로 넣어서")
$HANA = @(
  @{ d = '2026-11-10'; shop = '항공㈜';     item = '항공료'; amt = 608100 },
  @{ d = '2026-11-10'; shop = '에어프랑스'; item = '항공료'; amt = 582000 }
)
$LEFT = @(
  [pscustomobject]@{ 날짜 = '11/9&17'; 상호 = '도로공사'; 금액 = 10700; 사유 = '개인카드·현금(천문영) — 법인카드가 아니라 카드 전표에서 뺌(개인 경비 지급으로 따로 처리)' }
)

# 확인: 카드사 청구금액 = 미화 × 환율(원 미만 버림) + 수수료, 합계 910,226
foreach ($o in $OVS) { $calc = [math]::Floor($o.usd * $o.rate) + $o.fee; if ([math]::Abs($calc - $o.bill) -gt 1) { throw "청구금액 계산이 안 맞음: $($o.shop) 계산 $calc / 청구 $($o.bill)" } }
$sumBill = ($OVS | ForEach-Object { $_.bill } | Measure-Object -Sum).Sum; if ($sumBill -ne 910226) { throw "카드사 합계가 910,226이 아님: $sumBill" }

# 전표 목록 (날짜순, 날짜마다 800부터)
$VLIST = New-Object Collections.ArrayList
foreach ($o in $OVS) { [void]$VLIST.Add([pscustomobject]@{ 날짜 = $o.ld; 정산서날짜 = $o.d; 가맹점 = $o.shop; 항목 = $o.item; 금액 = [double]$o.bill; 구분 = '해외'; 카드코드 = '99709'; 카드명 = '경남비씨카드(2113)'; 외화 = "$($o.cur) $($o.amt)"; 미화 = $o.usd; 환율 = $o.rate; 수수료 = $o.fee; 정산서 = [math]::Round($o.sKrw); 현지거래일 = $o.ld; 적요 = "$prefix $($o.item) USD$($o.usd)*@$($o.rate)$cardSfx" }) }
foreach ($o in $DOM) { [void]$VLIST.Add([pscustomobject]@{ 날짜 = $o.d; 가맹점 = $o.shop; 항목 = $o.item; 금액 = [double]$o.amt; 구분 = '국내'; 카드코드 = '99709'; 카드명 = '경남비씨카드(2113)'; 외화 = ''; 미화 = ''; 환율 = ''; 수수료 = ''; 정산서 = $o.amt; 현지거래일 = ''; 적요 = "$prefix $($o.item)$cardSfx" }) }
foreach ($o in $HANA) { [void]$VLIST.Add([pscustomobject]@{ 날짜 = $o.d; 가맹점 = $o.shop; 항목 = $o.item; 금액 = [double]$o.amt; 구분 = '국내(하나5839)'; 카드코드 = '99905'; 카드명 = '하나카드'; 외화 = ''; 미화 = ''; 환율 = ''; 수수료 = ''; 정산서 = $o.amt; 현지거래일 = ''; 적요 = "$prefix $($o.item)_#하나5839" }) }
$VLIST = @($VLIST | Sort-Object 날짜, @{ Expression = { if ($_.구분 -eq '국내') { 0 } elseif ($_.구분 -eq '해외') { 1 } else { 2 } } })
# 작성번호: 카드 전표처럼 날짜마다 3000번대. 아마란스 자동전표처리에 이미 있는 번호(아마란스_자동전표번호.txt)와 예전 카드 전표 파일 번호는 건너뜀
#  (처음엔 '800번대'를 작성번호로 잘못 알아들었음 — 사용자 뜻은 계정 834 해외출장비(판관비))
$usedNo = @{}
$amaFile = Join-Path $env:TEMP '카드전표\아마란스_자동전표번호.txt'; if (Test-Path -LiteralPath $amaFile) { foreach ($l9 in (Get-Content -LiteralPath $amaFile -Encoding UTF8)) { if ($l9 -match '^(\d{8})\|(\d+)$') { $usedNo["$($Matches[1])|$([int]$Matches[2])"] = 1 } } }
. 'C:\ai\도구\엑셀없이_xlsx읽기.ps1'; Add-Type -AssemblyName System.IO.Compression
foreach ($pf in @(Get-ChildItem -LiteralPath @($OutDir, (Join-Path (Split-Path $OutDir) '지난결과')) -Recurse -Filter '*카드자동전표_*.xlsx' -ErrorAction SilentlyContinue | Where-Object { $_.Name -notmatch '검토|분개확인' -and $_.FullName -notmatch '\\백업\\' -and $_.FullName -ne $upload })) {
  $px = Read-Xlsx $pf.FullName 1; $h9 = ($px.Rows | Where-Object { $_.R -eq 2 } | Select-Object -First 1).Cells; $cDt = $null; $cSq = $null
  foreach ($k9 in $h9.Keys) { if ("$($h9[$k9])".Trim() -eq 'MENU_DT') { $cDt = $k9 } elseif ("$($h9[$k9])".Trim() -eq 'MENU_SQ') { $cSq = $k9 } }
  if ($cDt -and $cSq) { foreach ($q9 in ($px.Rows | Where-Object { $_.R -ge 4 })) { $d9 = "$($q9.Cells[$cDt])".Trim(); $s9 = "$($q9.Cells[$cSq])".Trim(); if ($d9 -match '^\d{8}$' -and $s9 -match '^\d+$') { $usedNo["$d9|$([int]$s9)"] = 1 } } }
}
"이미 쓴 날짜·번호 $($usedNo.Count)개를 건너뜀"
$seq = @{}; foreach ($v in $VLIST) { $k = ([datetime]$v.날짜).ToString('yyyyMMdd'); $n9 = $(if ($seq.ContainsKey($k)) { $seq[$k] + 1 } else { 3000 }); while ($usedNo.ContainsKey("$k|$n9")) { $n9++ }; $seq[$k] = $n9; $v | Add-Member 작성번호 $n9 }
foreach ($v in $VLIST) { if ($v.적요.Length -gt 80) { throw "적요가 80자 넘음: $($v.적요)" } }

$LINES = New-Object Collections.ArrayList
foreach ($v in $VLIST) {
  $dts = ([datetime]$v.날짜).ToString('yyyyMMdd')
  $base = @{ IN_DIV_CD = '1000'; MENU_DT = $dts; MENU_SQ = $v.작성번호; ISU_DOC = $v.적요; DOCU_TY = '1'; RMK_DC = $v.적요 }
  $e = $base.Clone(); $e.MENU_LN_SQ = 1; $e.DRCR_FG = '3'; $e.ACCT_CD = '8340000'; $e.TR_NM = $v.가맹점; $e.ACCT_AM = $v.금액; $e.PJT_CD = '303'; $e.CT_DEPT = '6220'; [void]$LINES.Add($e)
  $c = $base.Clone(); $c.MENU_LN_SQ = 2; $c.DRCR_FG = '4'; $c.ACCT_CD = '2530000'; $c.TR_CD = $v.카드코드; $c.TR_NM = $v.카드명; $c.ACCT_AM = $v.금액; [void]$LINES.Add($c)
}

# 엑셀 (카드 자동전표 도구와 같은 방식: 양식 → 2줄 칸코드에 맞춰 4줄부터)
New-Item -ItemType Directory -Force $OutDir | Out-Null
$work = Join-Path $env:TEMP '카드전표'; New-Item -ItemType Directory -Force $work | Out-Null; $tmpUp = Join-Path $work '해외출장_업로드_임시.xlsx'
foreach ($f in $upload, $review) { if (Test-Path -LiteralPath (Join-Path $OutDir ('~$' + [IO.Path]::GetFileName($f)))) { throw "엑셀에 열려 있음: $f" } }
$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$xl = New-Object -ComObject Excel.Application; $mine = @(Get-Process EXCEL | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
try {
  $xl.DisplayAlerts = $false
  if (Test-Path -LiteralPath $tmpUp) { Remove-Item -LiteralPath $tmpUp -Force }
  $wb = $xl.Workbooks.Open($Template, 0, $true); $wb.SaveAs($tmpUp, 51)
  $sh = $wb.Worksheets.Item(1); $maxC = $sh.UsedRange.Columns.Count; $colOf = @{}; $isText = @{}
  for ($ci = 1; $ci -le $maxC; $ci++) { $cd = "$($sh.Cells(2, $ci).Value2)".Trim(); if ($cd -and -not $colOf.ContainsKey($cd)) { $colOf[$cd] = $ci; $isText[$ci] = ("$($sh.Cells(3, $ci).Value2)" -match '타입\s*:\s*문자') } }
  $codes = 'IN_DIV_CD', 'MENU_DT', 'MENU_SQ', 'MENU_LN_SQ', 'ISU_DOC', 'DOCU_TY', 'DRCR_FG', 'ACCT_CD', 'TR_CD', 'TR_NM', 'ACCT_AM', 'RMK_DC', 'PJT_CD', 'CT_DEPT'
  $miss = @($codes | Where-Object { -not $colOf.ContainsKey($_) }); if ($miss.Count) { throw "양식에 칸이 없음: $($miss -join ', ')" }
  $n = $LINES.Count; $arr = New-Object 'object[,]' $n, $maxC
  for ($i = 0; $i -lt $n; $i++) { foreach ($cd in $codes) { $val = $LINES[$i][$cd]; if ($null -ne $val -and "$val" -ne '') { $arr[$i, ($colOf[$cd] - 1)] = $(if ($isText[$colOf[$cd]]) { [string]$val } else { $val }) } } }
  foreach ($cd in $codes) { if ($isText[$colOf[$cd]]) { $sh.Range($sh.Cells(4, $colOf[$cd]), $sh.Cells(3 + $n, $colOf[$cd])).NumberFormat = '@' } }
  $sh.Range($sh.Cells(4, 1), $sh.Cells(3 + $n, $maxC)).Value2 = $arr
  $sh.Range($sh.Cells(4, $colOf['ACCT_AM']), $sh.Cells(3 + $n, $colOf['ACCT_AM'])).NumberFormat = '#,##0'
  $wb.Save(); $wb.Close($false)
  if (Test-Path -LiteralPath $upload) { $bk = Join-Path $OutDir '백업'; New-Item -ItemType Directory -Force $bk | Out-Null; Move-Item -LiteralPath $upload (Join-Path $bk ("해외출장_카드자동전표_천문영_$($stamp)_" + (Get-Date -Format 'HHmmss') + '.xlsx')) -Force }
  Move-Item -LiteralPath $tmpUp $upload -Force

  # 검토 엑셀: ① 전표 ② 정산서 ↔ 카드사 대조 ③ 뺀 것
  $wr = $xl.Workbooks.Add(); while ($wr.Worksheets.Count -lt 3) { [void]$wr.Worksheets.Add([Type]::Missing, $wr.Worksheets.Item($wr.Worksheets.Count)) }
  function Fill($s, $name, $hdr, $rows) {
    $s.Name = $name; $nr = @($rows).Count; $dd = New-Object 'object[,]' ($nr + 1), $hdr.Count
    for ($c = 0; $c -lt $hdr.Count; $c++) { $dd[0, $c] = $hdr[$c] }
    $r = 1; foreach ($row in $rows) { for ($c = 0; $c -lt $hdr.Count; $c++) { $dd[$r, $c] = $row.($hdr[$c]) }; $r++ }
    $s.Range($s.Cells(1, 1), $s.Cells($nr + 1, $hdr.Count)).Value2 = $dd
    $h = $s.Range($s.Cells(1, 1), $s.Cells(1, $hdr.Count)); $h.Font.Bold = $true; $h.Interior.Color = 0xF2E6DD
    [void]$s.Columns.AutoFit()
  }
  $r1 = $VLIST | ForEach-Object { [pscustomobject]@{ 날짜 = $_.날짜; 작성번호 = $_.작성번호; 구분 = $_.구분; 가맹점 = $_.가맹점; 항목 = $_.항목; '차변 해외출장비 8340000' = $_.금액; '대변 미지급금' = $_.금액; 카드 = "[$($_.카드코드)] $($_.카드명)"; 프로젝트 = '303'; 사용부서 = '[6220] 해외영업팀_지점5'; 적요 = $_.적요 } }
  Fill $wr.Worksheets.Item(1) '전표(거래 1건=1장)' @('날짜', '작성번호', '구분', '가맹점', '항목', '차변 해외출장비 8340000', '대변 미지급금', '카드', '프로젝트', '사용부서', '적요') $r1
  $wr.Worksheets.Item(1).Range('F:G').NumberFormat = '#,##0'
  $r2 = $VLIST | Where-Object 구분 -eq '해외' | ForEach-Object { [pscustomobject]@{ '정산서 날짜' = $_.정산서날짜; '카드사 현지거래일' = $_.현지거래일; 가맹점 = $_.가맹점; 항목 = $_.항목; '현지금액' = $_.외화; '미화($)' = $_.미화; '적용환율' = $_.환율; '수수료' = $_.수수료; '청구금액(전표)' = $_.금액; '정산서 원화' = $_.정산서; '차이(청구-정산서)' = $_.금액 - $_.정산서 } }
  Fill $wr.Worksheets.Item(2) '정산서-카드사 대조(해외)' @('정산서 날짜', '카드사 현지거래일', '가맹점', '항목', '현지금액', '미화($)', '적용환율', '수수료', '청구금액(전표)', '정산서 원화', '차이(청구-정산서)') $r2
  $wr.Worksheets.Item(2).Range('H:K').NumberFormat = '#,##0'
  Fill $wr.Worksheets.Item(3) '뺀 것' @('날짜', '상호', '금액', '사유') $LEFT
  $wr.Worksheets.Item(3).Range('C:C').NumberFormat = '#,##0'
  $wr.Worksheets.Item(1).Activate()
  if (Test-Path -LiteralPath $review) { Remove-Item -LiteralPath $review -Force }
  $wr.SaveAs($review, 51); $wr.Close($false)
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl); [GC]::Collect() }
Start-Sleep -Seconds 2
foreach ($p in @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $mine -contains $_.Id -and $_.MainWindowHandle -eq 0 })) { Stop-Process -Id $p.Id -Force }

$dr = ($LINES | Where-Object { $_.DRCR_FG -eq '3' } | ForEach-Object { $_.ACCT_AM } | Measure-Object -Sum).Sum; $cr = ($LINES | Where-Object { $_.DRCR_FG -eq '4' } | ForEach-Object { $_.ACCT_AM } | Measure-Object -Sum).Sum
"전표 $($VLIST.Count)장 ($($LINES.Count)줄): 해외 $(@($VLIST | ? 구분 -eq '해외').Count)장 {0:N0} + 국내 $(@($VLIST | ? 구분 -eq '국내').Count)장 {1:N0} + 하나5839 항공 $(@($VLIST | ? 구분 -eq '국내(하나5839)').Count)장 {2:N0}" -f $sumBill, (($DOM | % { $_.amt } | Measure-Object -Sum).Sum), (($HANA | % { $_.amt } | Measure-Object -Sum).Sum)
"차변 {0:N0} / 대변 {1:N0}" -f $dr, $cr
"업로드: $upload"; "검토: $review"
