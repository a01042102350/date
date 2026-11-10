# 해외매출 자동전표(아마란스10 '자동전표처리(일반)' 업로드 엑셀) 만들기 — 매달
# 사용: powershell -ExecutionPolicy Bypass -File 해외매출_자동전표.ps1 -YM 202610
# 근거(2026-11-10, 9월 실제 전표 72장 분석):
#  - 인보이스 1건 = 전표 1장. 작성일자 = 수출 파일 BL일자, 금액 = 수출 파일 원화(당월매출)
#  - 줄: 국외매출(4140000, 대변) 라인별 여러 줄(프로젝트) → 부가세예수금(2550000, 대변 0원) → 외상매출금(1080000, 차변, 증빙 9)
#  - 라인별 금액 = 해외영업팀 '인보이스별 상세'(신코·코마츠는 지출품의서 마감)의 사형·자동·금형·가공1·가공2·가공3·스크랩 칸
#  - 적요 = "인보이스 품명 $외화*@환율 -라인이름"(부가세·외상 줄은 라인이름 없음), 관세는 전표를 따로
#  - 회계단위 1000 고정, 부가세 사업장 = 수출 파일 공장(마산1·마산2…) → 해외전표_사업장.csv
param([Parameter(Mandatory = $true)][string]$YM, [string]$ResultFile, [string]$SheetName, [string]$DetailFile, [string[]]$JpSources, [string]$OutDir, [string]$CompareCsv, [string]$TaxType = '16')
$ErrorActionPreference = 'Stop'
$root = 'C:\ai\매월마감'; $toolDir = Split-Path $MyInvocation.MyCommand.Path
$Y = [int]$YM.Substring(0, 4); $M = [int]$YM.Substring(4, 2); $cur = [DateTime]::new($Y, $M, 1); $MM = $cur.ToString('MM'); $ym4 = $cur.ToString('yyMM')
$monthDir = Join-Path $root $YM
if (-not $SheetName) { $SheetName = $cur.ToString('yy') + '.' + $MM + '(NEW)' }
if (-not $ResultFile) { $ResultFile = Join-Path $monthDir "$($Y)년 수출 ($($Y).$($MM))_$($M)월반영.xlsx" }
if (-not $DetailFile) { $DetailFile = @(Get-ChildItem $monthDir -File | Where-Object { $_.Name -like '*인보이스별 상세*' -and $_.Name -notlike '~$*' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1 | ForEach-Object FullName)[0] }
if (-not $JpSources) { $JpSources = @("$monthDir\신코_지출품의서 마감.xlsx|^신코20$($ym4)$", "$monthDir\코마츠_지출품의서 마감.xlsx|^코마츠20$($ym4)$") }
if (-not $OutDir) { $OutDir = Join-Path $root "해외자동전표\$YM" }
$template = Join-Path $root '해외자동전표\자동전표처리(일반) 양식.xlsx'
New-Item -ItemType Directory -Force $OutDir | Out-Null
$workDir = if (Test-Path $monthDir) { Join-Path $monthDir '작업' } else { $OutDir }; New-Item -ItemType Directory -Force $workDir | Out-Null
$logF = Join-Path $workDir '전표_실행기록.txt'; $doneF = Join-Path $workDir '전표_결과.txt'; $U8 = New-Object Text.UTF8Encoding($true)
[IO.File]::WriteAllText($logF, "$($Y)년 $($M)월 해외매출 자동전표 ($(Get-Date -f 'yyyy-MM-dd HH:mm:ss'))`r`n", $U8); if (Test-Path $doneF) { Remove-Item $doneF }
function Say([string]$m, [string]$c = 'Gray') { Write-Host $m -ForegroundColor $c; [IO.File]::AppendAllText($logF, "$m`r`n", $U8) }
function Done([string]$st, [string]$msg) { [IO.File]::WriteAllText($doneF, "$st`r`n$msg", $U8) }
$CHK = New-Object Collections.Generic.List[object]
function Chk($lvl, $inv, $text) { $CHK.Add([pscustomobject]@{ 구분 = $lvl; 인보이스 = $inv; 내용 = $text }) | Out-Null; Say "  [$lvl] $inv : $text" $(if ($lvl -eq '에러') { 'Red' } else { 'Yellow' }) }
function NumV($v) { if ($null -eq $v -or "$v".Trim() -eq '') { return $null }; if ($v -is [double]) { return $v }; $d = 0.0; if ([double]::TryParse(("$v" -replace '[,₩\$¥\s]', ''), [ref]$d)) { $d } else { $null } }
function NormInv($s) { ("$s" -replace '\s', '').ToUpper() }
function NormHead($s) { ("$s" -replace '\s', '').ToLower() }
function SplitName($h) { switch -regex ($h) { '^사형$' { '사형' } '^자동$' { '자동' } '^금형(주조)?$' { '금형' } '^가공1' { '가공1' } '^가공2' { '가공2' } '^가공3' { '가공3' } '^가공$' { '가공3' } '^스크랩\(사형\)' { '사형스크랩' } '^스크랩\(금형\)' { '금형스크랩' } default { $null } } }

# ── 설정 ──
$CUST = @{}; Import-Csv (Join-Path $toolDir '해외전표_거래처.csv') -Encoding UTF8 | ForEach-Object { $CUST[$_.키] = $_ }
$SITE = @{}; Import-Csv (Join-Path $toolDir '해외전표_사업장.csv') -Encoding UTF8 | ForEach-Object { $SITE[$_.공장] = $_ }
$code = 1
try {
Say "══ $($Y)년 $($M)월 해외매출 자동전표 ══" 'Cyan'
foreach ($f in $ResultFile, $DetailFile, $template) { if (-not $f -or -not (Test-Path -LiteralPath $f)) { throw "파일이 없습니다: $f" } }
Say "  수출 파일: $([IO.Path]::GetFileName($ResultFile)) / 시트 $SheetName"; Say "  라인별 상세: $([IO.Path]::GetFileName($DetailFile))"

$xl = New-Object -ComObject Excel.Application; $xl.DisplayAlerts = $false
try {
  # ── ① 수출 파일: 이번 달 매출 줄 ──
  $wb = $xl.Workbooks.Open($ResultFile, 0, $true); $ws = $null; foreach ($s0 in $wb.Worksheets) { if (($s0.Name -replace '\s', '') -eq ($SheetName -replace '\s', '')) { $ws = $s0 } }
  if (-not $ws) { throw "수출 파일에 '$SheetName' 시트가 없습니다." }
  $hdr = 0; for ($q = 5; $q -le 80; $q++) { if ($ws.Cells($q, 1).Text -match '^◆') { $hdr = $q; break } }
  $SALES = New-Object Collections.Generic.List[object]
  for ($q = $hdr + 3; $q -le 800; $q++) { $a1 = "$($ws.Cells($q,1).Value2)".Trim(); if ($a1 -match '총\s*계') { break }
    $invRaw = "$($ws.Cells($q,3).Value2)".Trim(); if (-not $invRaw) { continue }
    if ("$($ws.Cells($q,11).Formula)".Trim()) { continue }   # 다음 달 이월(회색) 제외
    $key = if ($a1 -like 'ALLISON*') { 'ALLISON' } else { $a1 }
    if ($key -like 'ATLAS*') { continue }                    # ATLAS는 국내(전자세금계산서) 매출이라 해외 전표 아님
    $bl = $ws.Cells($q, 6).Value2
    $SALES.Add([pscustomobject]@{ 행 = $q; 키 = $key; 공장 = $ws.Cells($q, 2).Text.Trim(); 인보이스 = $invRaw; 면장 = $ws.Cells($q, 4).Text.Trim()
        BL = $(if ($bl -is [double]) { [DateTime]::FromOADate($bl) } else { $null }); 환율 = (NumV $ws.Cells($q, 7).Value2); 외화 = (NumV $ws.Cells($q, 8).Value2); 원화 = (NumV $ws.Cells($q, 9).Value2)
        관세 = ($invRaw -match '관세'); 툴링 = ($invRaw -match '툴링') }) | Out-Null }
  $wb.Close($false)
  Say "  이번 달 매출 줄(이월·ATLAS 제외): $($SALES.Count)"

  # ── ② 라인별 금액 ──
  $SPL = @{}
  function ReadSplitBook($path, $sheetRe, [switch]$TabColor, [string]$src) {
    $bk = $xl.Workbooks.Open($path, 0, $true)
    foreach ($sh in $bk.Worksheets) {
      if ($TabColor) { if ($sh.Tab.ColorIndex -eq -4142) { continue } } elseif ($sheetRe -and $sh.Name -notmatch $sheetRe) { continue }
      $v = $sh.Range('A1:AD250').Value2; $hr = 0
      for ($r = 1; $r -le 250 -and -not $hr; $r++) { for ($c = 1; $c -le 30; $c++) { if ("$($v[$r,$c])".Trim() -eq '인보이스') { $hr = $r; break } } }
      if (-not $hr) { continue }
      $cols = @{}; $splitCols = @{}
      for ($c = 1; $c -le 30; $c++) { $h = NormHead $v[$hr, $c]; if (-not $h) { continue }; if (-not $cols.ContainsKey($h)) { $cols[$h] = $c }; $sn = SplitName $h; if ($sn -and -not $splitCols.ContainsKey($sn)) { $splitCols[$sn] = $c } }
      $ci = $cols['인보이스']; $cw = if ($cols['원화']) { $cols['원화'] } elseif ($cols['계']) { $cols['계'] } else { $cols['총'] }
      for ($r = $hr + 1; $r -le 250; $r++) { $inv = "$($v[$r,$ci])".Trim(); if (-not $inv -or $inv -match '합\s*계') { break }
        $k = NormInv ($inv -replace '\(툴링비\)', '')
        $parts = [ordered]@{}; foreach ($sn in ' 사형 자동 금형 가공1 가공2 가공3 사형스크랩 금형스크랩'.Trim().Split(' ')) { if ($splitCols.ContainsKey($sn)) { $a = NumV $v[$r, $splitCols[$sn]]; if ($a -and [Math]::Abs($a) -ge 0.5) { $parts[$sn] = [Math]::Round($a, 0) } } }
        if ($SPL.ContainsKey($k)) { Chk '알림' $inv "라인별 금액이 두 군데에 있습니다('$($SPL[$k].시트)', '$($sh.Name)'). 앞의 것을 씁니다."; continue }
        $SPL[$k] = [pscustomobject]@{ 시트 = $sh.Name; 행 = $r; 원화 = (NumV $v[$r, $cw]); 라인 = $parts; 출처 = $src } } }
    $bk.Close($false) }
  ReadSplitBook $DetailFile $null -src ([IO.Path]::GetFileName($DetailFile))
  foreach ($js in $JpSources) { $pp = $js -split '\|', 2; if (-not (Test-Path -LiteralPath $pp[0])) { Chk '알림' '' "신코·코마츠 파일이 없습니다: $($pp[0])"; continue }
    if ($pp[1] -eq '색칠') { ReadSplitBook $pp[0] $null -TabColor -src ([IO.Path]::GetFileName($pp[0])) } else { ReadSplitBook $pp[0] $pp[1] -src ([IO.Path]::GetFileName($pp[0])) } }
  Say "  라인별 금액을 읽은 인보이스: $($SPL.Count)"

  # ── ③ 전표 만들기 ──
  $LINES = New-Object Collections.Generic.List[object]; $VCH = New-Object Collections.Generic.List[object]; $seqByDate = @{}
  $projFor = @{ 사형 = '101'; 자동 = '102'; 가공1 = '401'; 가공2 = '402'; 가공3 = '403' }
  foreach ($x in ($SALES | Sort-Object { $_.BL }, 행)) {
    $cu = $CUST[$x.키]; if (-not $cu) { Chk '에러' $x.인보이스 "거래처 설정(해외전표_거래처.csv)에 '$($x.키)'가 없습니다. 이 인보이스는 넣지 않았습니다."; continue }
    if (-not $x.BL) { Chk '에러' $x.인보이스 'BL일자가 없어 넣지 않았습니다.'; continue }
    if ($x.인보이스 -notmatch '^(DSM\d{4}-|\d{4}P)') { Chk '에러' $x.인보이스 "인보이스 형식이 아니라 정정 줄로 보여 전표를 만들지 않았습니다(원화 $('{0:N0}' -f $x.원화)). 필요하면 해당 전표에 직접 반영해 주세요."; continue }
    if ($null -eq $x.원화 -or [Math]::Abs([double]$x.원화) -lt 1) { Chk '알림' $x.인보이스 '원화가 0원(무상 샘플 등)이라 전표를 만들지 않았습니다.'; continue }
    if (-not $cu.거래처코드) { Chk '확인' $x.인보이스 "거래처 '$($cu.거래처명)'의 거래처코드가 없습니다(새 거래처). 업로드 전에 거래처코드를 넣어 주세요." }
    $d8 = $x.BL.ToString('yyyyMMdd'); if (-not $seqByDate.ContainsKey($d8)) { $seqByDate[$d8] = 0 }; $seqByDate[$d8]++; $no = $seqByDate[$d8]
    $invTxt = $x.인보이스 -replace '\s', ''
    $desc = if ($x.관세 -and $cu.관세품명) { $cu.관세품명 } elseif ($x.툴링 -and $cu.툴링품명) { $cu.툴링품명 } else { $cu.품명 }
    if ($x.툴링) { Chk '확인' $x.인보이스 "툴링비 전표는 9월에 예가 없어 적요 품명을 '$desc'(으)로 넣었습니다. 평소 쓰시는 품명으로 바꿔 주세요." }
    $krw = [Math]::Round([double]$x.원화, 0)
    $amtTxt = switch ($cu.통화) { 'JPY' { '¥' + ('{0:N2}' -f $x.외화) + '*@' + ([double]$x.환율).ToString('0.####') } 'KRW' { '\' + ('{0:N0}' -f $krw) } default { '$' + ('{0:N2}' -f $x.외화) + '*@' + ('{0:N2}' -f $x.환율) } }
    $base = ("$invTxt $desc $amtTxt" -replace '\s+', ' ').Trim()
    # 라인별 금액
    $sp0 = $SPL[(NormInv ($x.인보이스 -replace '\(툴링비\)', ''))]; $parts = [ordered]@{}
    if (-not $sp0 -and $x.인보이스 -match '^DSM\d{2}(\d{2}-.+)$') { $tail = $matches[1]; $cand = @($SPL.Keys | Where-Object { $_ -match ('^DSM\d{2}' + [regex]::Escape((NormInv $tail)) + '$') })
      if ($cand.Count -eq 1) { $sp0 = $SPL[$cand[0]]; Chk '알림' $x.인보이스 "라인별 금액을 연도만 다른 인보이스 '$($cand[0])'에서 가져왔습니다(수출 파일 인보이스 연도 오기로 보임)." } }
    if ($sp0) { foreach ($kk in $sp0.라인.Keys) { $parts[$kk] = $sp0.라인[$kk] }
      if ($null -ne $sp0.원화 -and [Math]::Abs($sp0.원화 - $krw) -ge 1) {
        if ($cu.통화 -eq 'JPY') { Chk '알림' $x.인보이스 "엔화 건이라 매출내역서 원화 $('{0:N0}' -f $sp0.원화)을(를) 썼습니다(수출 파일 원화 $('{0:N0}' -f $krw), 9월 SK03도 매출내역서 금액으로 전표 처리)."; $krw = [Math]::Round($sp0.원화, 0) }
        else { Chk '확인' $x.인보이스 "라인별 상세의 원화($('{0:N0}' -f $sp0.원화))가 수출 파일 원화($('{0:N0}' -f $krw))와 다릅니다." } } }
    if ($parts.Count -eq 0) { $parts['금형'] = $krw; Chk '에러' $x.인보이스 "라인별 금액을 찾지 못해 전액을 '$($cu.금형이름)'(프로젝트 $($cu.금형프로젝트)) 한 줄로 넣었습니다. 라인을 나눠 주세요." }
    $sum = 0.0; foreach ($kk in @($parts.Keys)) { $sum += $parts[$kk] }
    if ([Math]::Abs($sum - $krw) -ge 1) { $big = @($parts.Keys | Sort-Object { -$parts[$_] })[0]; $parts[$big] = $parts[$big] + ($krw - $sum)
      Chk '확인' $x.인보이스 "라인별 합계($('{0:N0}' -f $sum))가 원화($('{0:N0}' -f $krw))와 달라 차이 $('{0:N0}' -f ($krw - $sum))원을 '$big' 줄에 더했습니다." }
    $siteRow = $SITE[$x.공장]; if (-not $siteRow) { Chk '에러' $x.인보이스 "공장 '$($x.공장)'의 부가세 사업장 설정이 없습니다." }
    $ctNo = if ($x.면장 -match '^\d{5}-\d{2}-\d{6}[A-Z]?$') { $x.면장 } else { '' }
    if (-not $ctNo -and -not $x.관세) { Chk '알림' $x.인보이스 "수출신고번호(면장번호)가 없습니다(수출 파일 면장 칸: '$($x.면장)')." }
    $ln = 0
    foreach ($kk in $parts.Keys) { $ln++
      $proj = if ($kk -eq '금형' -or $kk -eq '금형스크랩') { $cu.금형프로젝트 } elseif ($kk -eq '사형스크랩') { '101' } else { $projFor[$kk] }
      $nm = switch ($kk) { '사형' { $cu.사형이름 } '금형' { $cu.금형이름 } '사형스크랩' { $cu.사형이름 + '스크랩' } '금형스크랩' { $cu.금형이름 + '스크랩' } default { $kk } }
      $LINES.Add([pscustomobject]@{ 회계단위 = '1000'; 작성일자 = $d8; 작성번호 = $no; 라인순번 = $ln; 차대구분 = 4; 계정과목 = 4140000; 거래처 = $cu.거래처코드; 거래처명 = $cu.거래처명; 계정금액 = [double]$parts[$kk]; 증빙코드 = ''; 적요 = "$base -$nm"; 부가세사업장 = ''; 세무구분 = ''; 공급가액 = $null; 환종 = ''; 외화금액 = $null; 환율 = $null; 수출신고번호 = ''; 프로젝트 = "$proj"; 인보이스 = $x.인보이스 }) | Out-Null }
    if ($cu.통화 -eq 'KRW') { Chk '확인' $x.인보이스 "원화수출이라 부가세 줄의 환종을 KRW, 외화금액 = 원화, 환율 = 1로 넣었습니다. 평소와 같은지 확인해 주세요." }
    $ln++; $cur2 = $cu.통화
    $LINES.Add([pscustomobject]@{ 회계단위 = '1000'; 작성일자 = $d8; 작성번호 = $no; 라인순번 = $ln; 차대구분 = 4; 계정과목 = 2550000; 거래처 = $cu.거래처코드; 거래처명 = $cu.거래처명; 계정금액 = [double]0; 증빙코드 = ''; 적요 = $base; 부가세사업장 = $(if ($siteRow) { $siteRow.부가세사업장코드 } else { '' }); 세무구분 = $TaxType; 공급가액 = [double]$krw; 환종 = $cur2; 외화금액 = $(if ($cur2 -eq 'KRW') { [double]$krw } else { [double]$x.외화 }); 환율 = $(if ($cur2 -eq 'KRW') { [double]1 } else { [double]$x.환율 }); 수출신고번호 = $ctNo; 프로젝트 = ''; 인보이스 = $x.인보이스 }) | Out-Null
    $ln++
    $LINES.Add([pscustomobject]@{ 회계단위 = '1000'; 작성일자 = $d8; 작성번호 = $no; 라인순번 = $ln; 차대구분 = 3; 계정과목 = 1080000; 거래처 = $cu.거래처코드; 거래처명 = $cu.거래처명; 계정금액 = [double]$krw; 증빙코드 = '9'; 적요 = $base; 부가세사업장 = ''; 세무구분 = ''; 공급가액 = $null; 환종 = ''; 외화금액 = $null; 환율 = $null; 수출신고번호 = ''; 프로젝트 = ''; 인보이스 = $x.인보이스 }) | Out-Null
    $VCH.Add([pscustomobject]@{ 작성일자 = $d8; 작성번호 = $no; 인보이스 = $x.인보이스; 거래처 = $cu.거래처명; 공장 = $x.공장; 부가세사업장 = $(if ($siteRow) { "$($siteRow.부가세사업장코드) $($siteRow.사업장명)" }); 원화 = $krw; 외화 = $x.외화; 환율 = $x.환율; 면장 = $ctNo
        라인 = (($parts.Keys | ForEach-Object { "$_ $('{0:N0}' -f $parts[$_])" }) -join ' / '); 라인출처 = $(if ($sp0) { "$($sp0.출처) '$($sp0.시트)' $($sp0.행)행" } else { '없음' }); 적요 = $base }) | Out-Null }
  $saleSum = ($SALES | Where-Object { $_.원화 } | Measure-Object 원화 -Sum).Sum; $vSum = ($VCH | Measure-Object 원화 -Sum).Sum
  Chk '참고' '' "외상매출금 합계 $('{0:N0}' -f $vSum)원 / 수출 파일 당월매출(다음 달 이월·ATLAS 제외) $('{0:N0}' -f $saleSum)원 / 차이 $('{0:N0}' -f ($vSum - $saleSum))원 (차이가 있으면 위의 0원·정정 줄·엔화 매출내역서 금액 때문)"
  # 차대 점검
  foreach ($g in ($LINES | Group-Object 작성일자, 작성번호)) { $dr = ($g.Group | Where-Object { $_.차대구분 -eq 3 } | Measure-Object 계정금액 -Sum).Sum; $cr = ($g.Group | Where-Object { $_.차대구분 -eq 4 } | Measure-Object 계정금액 -Sum).Sum
    if ([Math]::Abs($dr - $cr) -ge 0.5) { Chk '에러' $g.Group[0].인보이스 "차변 $dr 과 대변 $cr 이 다릅니다." } }
  $LINES | Export-Csv (Join-Path $OutDir "해외매출_자동전표_$($YM)_줄.csv") -NoTypeInformation -Encoding UTF8
  Say "  전표 $($VCH.Count)장, $($LINES.Count)줄, 외상매출금 합계 $('{0:N0}' -f (($VCH | Measure-Object 원화 -Sum).Sum))" 'Green'

  # ── ④ 업로드 엑셀 (양식 그대로, 4행부터) ──
  $upload = Join-Path $OutDir "해외매출_자동전표_$($YM).xlsx"
  Copy-Item -LiteralPath $template $upload -Force
  $wb = $xl.Workbooks.Open($upload); $sh = $wb.Worksheets.Item(1)
  $props = '회계단위', '작성일자', '작성번호', '라인순번', '차대구분', '계정과목', '거래처', '거래처명', '계정금액', '증빙코드', '적요', '부가세사업장', '세무구분', '공급가액', '환종', '외화금액', '환율', '수출신고번호', '프로젝트'
  $textCols = @(1, 2, 7, 8, 10, 11, 12, 13, 15, 18, 19)
  $n = $LINES.Count
  if ($n) {
    foreach ($c in $textCols) { $sh.Range($sh.Cells(4, $c), $sh.Cells(3 + $n, $c)).NumberFormat = '@' }
    $data = New-Object 'object[,]' $n, 22
    for ($i = 0; $i -lt $n; $i++) { $l = $LINES[$i]; for ($c = 0; $c -lt $props.Count; $c++) { $val = $l.($props[$c]); if ($null -eq $val -or "$val" -eq '') { continue }
        if ($textCols -contains ($c + 1)) { $data.SetValue([object][string]$val, [int]$i, [int]$c) } else { $data.SetValue([object][double]$val, [int]$i, [int]$c) } } }
    $sh.Range($sh.Cells(4, 1), $sh.Cells(3 + $n, 22)).Value2 = $data }
  $wb.Save(); $wb.Close($false)
  Say "  업로드 파일: $upload" 'Green'

  # ── ⑤ 9월 실제 전표와 대조(시험용) ──
  $CMP = @()
  if ($CompareCsv) {
    $act = Import-Csv $CompareCsv -Encoding UTF8 | Where-Object { $_.거래처분류 -eq '해외매출' }
    foreach ($v0 in $VCH) { $invK = $v0.인보이스 -replace '\(툴링비\)', '' -replace '\s', ''
      $a = @($act | Where-Object { ($_.품의내역 -split '\s')[0] -eq $invK })
      if (-not $a.Count) { $CMP += [pscustomobject]@{ 인보이스 = $v0.인보이스; 결과 = '실제 전표 없음'; 내용 = '' }; continue }
      $a = $a[0]; $aLines = @($a.구성 -split ' ; ' | ForEach-Object { $f = $_ -split '/'; [pscustomobject]@{ 계정 = $f[0]; 금액 = [double](($f[1] -replace '^[CD]', '') -replace '^$', '0'); 프로젝트 = $f[2] } })
      $aTot = ($aLines | Where-Object { $_.계정 -eq '1080000' } | Measure-Object 금액 -Sum).Sum
      $aPj = @{}; foreach ($al in ($aLines | Where-Object { $_.계정 -eq '4140000' })) { $aPj[$al.프로젝트] = $aPj[$al.프로젝트] + $al.금액 }
      $gPj = @{}; foreach ($gl in ($LINES | Where-Object { $_.인보이스 -eq $v0.인보이스 -and $_.계정과목 -eq 4140000 })) { $gPj[$gl.프로젝트] = $gPj[$gl.프로젝트] + $gl.계정금액 }
      $msg = @(); if ($a.작성일 -replace '-', '' -ne $v0.작성일자) { $msg += "날짜 실제 $($a.작성일)" }
      if ([Math]::Abs($aTot - $v0.원화) -ge 1) { $msg += "합계 실제 $('{0:N0}' -f $aTot) / 만든 것 $('{0:N0}' -f $v0.원화)" }
      $pjs = @($aPj.Keys) + @($gPj.Keys) | Select-Object -Unique
      foreach ($pj in $pjs) { if ([Math]::Abs([double]$aPj[$pj] - [double]$gPj[$pj]) -ge 1) { $msg += "프로젝트 $pj 실제 $('{0:N0}' -f [double]$aPj[$pj]) / 만든 것 $('{0:N0}' -f [double]$gPj[$pj])" } }
      $CMP += [pscustomobject]@{ 인보이스 = $v0.인보이스; 결과 = $(if ($msg) { '다름' } else { '같음' }); 내용 = ($msg -join ' / ') } }
    $same = @($CMP | Where-Object 결과 -eq '같음').Count
    Say "  실제 전표와 대조: 같음 $same / $($CMP.Count)" 'Cyan'
    $CMP | Where-Object 결과 -ne '같음' | ForEach-Object { Say "    $($_.인보이스) [$($_.결과)] $($_.내용)" } }

  # ── ⑥ 검토 엑셀 ──
  $review = Join-Path $OutDir "해외매출_자동전표_$($YM)_검토.xlsx"
  $wb = $xl.Workbooks.Add(); while ($wb.Worksheets.Count -lt 3) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  function FillSheet($sh, $hdrs, $rows, $props) { $nr = @($rows).Count; $dd = New-Object 'object[,]' ($nr + 1), $hdrs.Count
    for ($c = 0; $c -lt $hdrs.Count; $c++) { $dd.SetValue([object][string]$hdrs[$c], 0, $c) }
    $ri = 0; foreach ($rw in $rows) { $ri++; for ($c = 0; $c -lt $props.Count; $c++) { $val = $rw.($props[$c]); if ($null -eq $val) { continue }; $val = $val.psobject.BaseObject; if ($val -is [double] -or $val -is [int] -or $val -is [long] -or $val -is [decimal]) { $dd.SetValue([object][double]$val, $ri, $c) } else { $dd.SetValue([object][string]$val, $ri, $c) } } }
    $sh.Range($sh.Cells(1, 1), $sh.Cells($nr + 1, $hdrs.Count)).Value2 = $dd; $hh = $sh.Range($sh.Cells(1, 1), $sh.Cells(1, $hdrs.Count)); $hh.Font.Bold = $true; $hh.Interior.Color = 0xF2E6DD; [void]$sh.Columns.AutoFit() }
  $s1 = $wb.Worksheets.Item(1); $s1.Name = '전표 목록'
  FillSheet $s1 @('작성일자', '작성번호', '인보이스', '거래처', '공장', '부가세 사업장', '원화(외상매출금)', '외화', '환율', '수출신고번호', '라인별 금액', '라인 금액 출처', '적요') @($VCH | ForEach-Object { $_ }) @('작성일자', '작성번호', '인보이스', '거래처', '공장', '부가세사업장', '원화', '외화', '환율', '면장', '라인', '라인출처', '적요')
  $s1.Range('G:G').NumberFormat = '#,##0'; $s1.Range('H:H').NumberFormat = '#,##0.00'
  $s2 = $wb.Worksheets.Item(2); $s2.Name = '확인할 것'
  if ($CHK.Count) { FillSheet $s2 @('구분', '인보이스', '내용') @($CHK | ForEach-Object { $_ }) @('구분', '인보이스', '내용'); $s2.Columns('C').ColumnWidth = 120 } else { $s2.Range('A1').Value2 = '확인할 것 없음' }
  $s3 = $wb.Worksheets.Item(3); $s3.Name = '실제 전표 대조'
  if ($CMP.Count) { FillSheet $s3 @('인보이스', '결과', '내용') @($CMP | ForEach-Object { $_ }) @('인보이스', '결과', '내용') } else { $s3.Range('A1').Value2 = '대조 안 함' }
  $s1.Activate(); $wb.SaveAs($review, 51); $wb.Close($false)
  Say "  검토 파일: $review" 'Green'
}
finally { if ($xl) { try { foreach ($w0 in @($xl.Workbooks)) { $w0.Close($false) } } catch { }; $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } }
  Done 'OK' ("전표 $($VCH.Count)장, $($LINES.Count)줄, 외상매출금 합계 $('{0:N0}' -f (($VCH | Measure-Object 원화 -Sum).Sum))원`r`n확인할 것: $(@($CHK | Where-Object { $_.구분 -ne '참고' }).Count)건 (검토 파일 '확인할 것' 시트)`r`n업로드 파일: $upload"); $code = 0
}
catch { $m0 = "중간에 멈췄습니다: $($_.Exception.Message) (스크립트 $($_.InvocationInfo.ScriptLineNumber)행)"; Say $m0 'Red'; Done 'ERROR' $m0; $code = 1 }
exit $code
