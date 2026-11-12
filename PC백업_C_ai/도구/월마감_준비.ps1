# 월마감_준비.ps1 — 월마감 첫 단계 (2026-11-12 만듦, 10월부터)
#   9월(202609)에는 이 단계를 손으로 했음 → 9월 결과(백업\mes9월_백업_20261028.xls)와 같은 규칙으로 만듦 (matching.md 1~6·13·14)
#   실행:  & "C:\ai\도구\월마감_준비.ps1" -YM 202610      → 다음에  & "C:\ai\도구\월마감_반영.ps1" -YM 202610
#   하는 일
#   - MES에서 받은 원본(mesN월.xls, HTML 형식)은 백업\mesN월_MES원본.xls 로 보관, 그 달 출고분만 남김(다음 달 출고 줄은 뺌)
#   - 칸 추가: 출고일 뒤 BL날짜·인보이스 번호·환율·달러·수출면장 번호, 맨 끝 사형·자동·금형 주조·가공1~3 (6칸)
#     (BL날짜·인보이스 번호·기초자료환율·기초자료달러는 월마감_반영이 채움)
#   - 환율 = BL일 서울외국환중개 매매기준율(엔화는 100엔 값 ÷ 100), Bergen = 회계 환율(1,350), Atlas(전자세금계산서) = 원화 거래라 비움
#   - 달러 = MES 금액 ÷ 환율 (인보이스 합계가 회계 외화금액과 다르면 빨간색 = 출하기준 금액변동)
#   - 수출면장 번호 = 회계 면장번호 ↔ 해외영업팀 N월마감 '수출신고필증 NO.' 비교
#   - 6칸 = 수량×주조비 → 라인(사형·자동·금형 주조), 나머지 → 가공(가공1·2·3). 신코·코마츠는 매출내역서 금액을 줄별로 나눔(연회색)
#   - 확인 시트 4장: 면장불일치·환율불일치·달러불일치·라인구분불일치 (월마감_반영이 오류목록으로 합친 뒤 지움)
#   -Force : 이미 준비된 파일을 MES 원본부터 다시 만듦 (mes 시트에 사람이 넣은 체크·금액 수정은 사라짐 → 먼저 백업\에 복사해 둠)
#   MES 첫 출고일이 그달 4일 이후면 멈춤(받은 기간이 1일~말일이 아닌 것) → 정말 맞으면 -AllowLateStart
param([int]$YM = 0, [switch]$Force, [switch]$AllowLateStart)
$ErrorActionPreference = 'Stop'
if (-not $YM) { $YM = [int](@(Get-ChildItem 'C:\ai\기초자료' -Directory | ? { $_.Name -match '^\d{6}$' -and @(Get-ChildItem $_.FullName -File -Filter 'mes*월.xls' -ErrorAction SilentlyContinue).Count } | sort Name)[-1].Name) }
$yy = [int][math]::Floor($YM / 100); $mm = [int]($YM % 100); $nx = $mm % 12 + 1; $mm2 = '{0:00}' -f $mm; $yy2 = '{0:00}' -f ($yy % 100)
$mLab = "$($mm)월"; $nLab = "$($nx)월"
$base = "C:\ai\기초자료\$YM"
$mesPath = "$base\mes$($mm)월.xls"
$rawBak = "$base\백업\mes$($mm)월_MES원본.xls"
$sheetName = "mes$($mm)월"
$firstOA = (New-Object DateTime $yy, $mm, 1).ToOADate(); $nextOA = (New-Object DateTime $yy, $mm, 1).AddMonths(1).ToOADate()
$cats = @('사형', '자동', '금형 주조', '가공1(2공장)', '가공2(5공장)', '가공3(3공장)')
$RED = 255; $COLGRAY = 15921906; $HEAD = 14277081
"== $YM 준비 (MES 원본 → 기본 칸 채우기)"

function Num($x) { if ($x -is [double]) { return $x }; $d = 0.0; if ([double]::TryParse(("$x" -replace '[,\s₩$¥]', ''), [ref]$d)) { $d } else { $null } }
function K($x) { "$x".Trim() }
function DS($n) { if ($n) { [DateTime]::FromOADate($n).ToString('yyyy-MM-dd') } else { '없음' } }
function ToOA($x) {
  if ($x -is [double]) { if ($x -gt 1 -and $x -lt 100000) { return [math]::Floor($x) } else { return $null } }
  $t = "$x".Trim()
  if ($t -match '^(\d{4})[-./](\d{1,2})[-./](\d{1,2})') { return (New-Object DateTime ([int]$matches[1]), ([int]$matches[2]), ([int]$matches[3])).ToOADate() }
  if ($t -match '^(\d{1,2})\s*/\s*(\d{1,2})$') { return (New-Object DateTime $yy, ([int]$matches[1]), ([int]$matches[2])).ToOADate() }
  $null
}
function Head($p) { $fs = [IO.File]::OpenRead($p); try { $buf = New-Object byte[] 512; $n = $fs.Read($buf, 0, 512); [Text.Encoding]::ASCII.GetString($buf, 0, $n) } finally { $fs.Close() } }

if (-not (Test-Path $mesPath)) { throw "MES 파일이 없음: $mesPath" }
try { $fs0 = [IO.File]::Open($mesPath, 'Open', 'ReadWrite', 'None'); $fs0.Close() } catch { throw "mes$($mm)월.xls가 엑셀에서 열려 있음 → 닫고 다시 실행해 주세요" }
if (-not (Test-Path "$base\백업")) { New-Item -ItemType Directory "$base\백업" | Out-Null }
if ((Head $mesPath) -match '<html') { Copy-Item $mesPath $rawBak -Force; $srcFile = $rawBak; "MES 원본 보관: $rawBak" }
elseif ($Force) {
  if (-not (Test-Path $rawBak)) { throw "MES 원본(백업\mes$($mm)월_MES원본.xls)이 없어 다시 만들 수 없음" }
  $bk = "$base\백업\mes$($mm)월_백업_$(Get-Date -Format 'yyyyMMdd_HHmm')_준비다시전.xls"; Copy-Item $mesPath $bk; "지금 파일 백업: $bk"
  $srcFile = $rawBak
} else { "이미 준비된 파일입니다 (MES 원본 형식이 아님) → 건너뜀. 원본부터 다시 하려면 -Force"; return }

# ── 서울외국환중개 매매기준율 (매월마감 도구가 모아 둔 CSV, 없는 날짜는 smbs.biz에서 읽음. 휴일 = 앞 영업일)
$fxCsv = 'C:\ai\매월마감\도구\서울외국환중개_환율.csv'; $fx = @{}
if (Test-Path $fxCsv) { foreach ($r0 in @(Import-Csv $fxCsv -Encoding UTF8)) { $p0 = @($r0.PSObject.Properties); $fx["$($p0[0].Value)".Trim()] = @{ USD = (Num $r0.USD); JPY = (Num $r0.JPY) } } }
function Get-Fx($oa, $cur) {
  $d = [DateTime]::FromOADate($oa)
  for ($k = 0; $k -le 10; $k++) {
    $dk = $d.AddDays(-$k); $key = $dk.ToString('yyyy-MM-dd')
    if (-not $fx.ContainsKey($key) -and $dk -le (Get-Date)) {
      try { [Net.ServicePointManager]::SecurityProtocol = 'Tls12'
        $t = (Invoke-WebRequest "http://www.smbs.biz/Flash/TodayExRate_flash.jsp?tr_date=$key" -UseBasicParsing -TimeoutSec 20).Content
        $fx[$key] = @{ USD = $(if ($t -match 'USD=([\d,\.]+)') { [double]($matches[1] -replace ',', '') } else { $null }); JPY = $(if ($t -match 'JPY=([\d,\.]+)') { [double]($matches[1] -replace ',', '') / 100 } else { $null }) }
      } catch { $fx[$key] = @{ USD = $null; JPY = $null } }
    }
    if ($fx.ContainsKey($key) -and $fx[$key][$cur]) { return [pscustomobject]@{ Rate = [double]$fx[$key][$cur]; Date = $key } }
  }
  $null
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  # ── N월라인 (해외영업팀 인보이스별 상세): 인보이스 → BL·환율·달러·원화·6칸
  $line = @{}
  $lf = @(Get-ChildItem $base -File | ? { $_.Name -match "^$($mm)월라인" -and $_.Extension -match '^\.xls' -and $_.Name -notmatch '^~\$' })[0]
  if (-not $lf) { throw "$($mLab)라인 파일이 없음 ($base)" }
  $wbR = $xl.Workbooks.Open($lf.FullName, 0, $true)
  foreach ($s in @($wbR.Worksheets)) {
    $u = $s.UsedRange; $vv = $u.Value2; $nr = $u.Rows.Count; $ncx = $u.Columns.Count; $hr = $null; $col = @{}
    for ($i = 1; $i -le [math]::Min($nr, 25) -and -not $hr; $i++) { for ($j = 1; $j -le $ncx; $j++) { if ("$($vv[$i,$j])".Trim() -eq '인보이스') { $hr = $i } } }
    if (-not $hr) { continue }
    for ($j = 1; $j -le $ncx; $j++) { $h = ("$($vv[$hr,$j])" -replace '\s', ''); if ($h -and -not $col.ContainsKey($h)) { $col[$h] = $j } }
    $blC = if ($col['bl날짜']) { $col['bl날짜'] } else { $col['전자세금계산서발행일'] }
    for ($i = $hr + 1; $i -le $nr; $i++) {
      $inv = K $vv[$i,$col['인보이스']]; if (-not $inv -or $line.ContainsKey($inv)) { continue }
      $sp = @{}; foreach ($c0 in $cats) { $cc = $col[($c0 -replace '\s', '')]; if ($cc) { $sp[$c0] = [double](Num $vv[$i,$cc]) } }
      $line[$inv] = [pscustomobject]@{ Bl = $(if ($blC) { ToOA $vv[$i,$blC] }); Rate = $(if ($col['환율']) { Num $vv[$i,$col['환율']] }); Usd = $(if ($col['달러']) { Num $vv[$i,$col['달러']] }); Krw = $(if ($col['원화']) { Num $vv[$i,$col['원화']] }); Split = $sp }
    }
  }
  $wbR.Close($false)
  "$($lf.Name): 인보이스 $($line.Count)건"

  # ── 신코·코마츠 매출내역서 엑셀: 'YYYY신코/코마츠' 시트가 있으면 그 시트만, 없으면 탭이 색칠된 시트 (월마감_반영과 같은 규칙)
  $sk = [ordered]@{}
  foreach ($f in @(Get-ChildItem $base -File | ? { $_.Name -match '신코|코마츠' -and $_.Extension -match '^\.xls' -and $_.Name -notmatch '^~\$' })) {
    $wbS = $xl.Workbooks.Open($f.FullName, 0, $true)
    try {
      $shs = @(@($wbS.Worksheets) | ? { $_.Name -match "^(신코|코마츠)\s*$YM$" })
      if (-not $shs.Count) { $shs = @(@($wbS.Worksheets) | ? { $_.Tab.ColorIndex -ne -4142 -and $_.Visible -eq -1 }) }
      foreach ($sh in $shs) {
        $u = $sh.UsedRange; $nr = [math]::Min($u.Row + $u.Rows.Count - 1, 80); $ncx = [math]::Min($u.Column + $u.Columns.Count - 1, 60)
        $vv = $sh.Range($sh.Cells(1,1), $sh.Cells($nr,$ncx)).Value2; $hr = $null; $hc = $null; $col = @{}
        for ($i = 1; $i -le $nr -and -not $hr; $i++) { for ($j = 1; $j -le $ncx; $j++) { if ("$($vv[$i,$j])".Trim() -eq '인보이스') { $hr = $i; break } } }
        if (-not $hr) { continue }
        for ($j = 1; $j -le $ncx; $j++) { $h = ("$($vv[$hr,$j])" -replace '\s', ''); if ($h -and -not $col.ContainsKey($h)) { $col[$h] = $j } }
        if (-not $col['엔화금액']) { continue }
        for ($i = $hr + 1; $i -le [math]::Min($nr, $hr + 40); $i++) {
          $inv = K $vv[$i,$col['인보이스']]
          if ($inv -match '합\s*계') { break }; if (-not $inv) { continue }
          $pm = $null; if ($inv -match '^(\d{1,2})\s*-\s*(.+)$') { $pm = [int]$matches[1]; $inv = $matches[2].Trim() }
          if ($pm -and $pm -gt $mm) { continue }   # 'NN-' = 다음 달 → 이번 달 매출 아님
          $g0 = { param($nm) if ($col[$nm]) { [double](Num $vv[$i,$col[$nm]]) } else { 0.0 } }
          $sk[$inv] = [pscustomobject]@{ Bl = (ToOA $vv[$i,$col['BL날짜']]); Rate = (& $g0 '환율'); Fc = [math]::Round((& $g0 '엔화금액'))
            Split = @{ '사형' = (& $g0 '사형') + (& $g0 '스크랩(사형)'); '금형 주조' = (& $g0 '금형') + (& $g0 '스크랩(금형)'); '가공3(3공장)' = (& $g0 '가공') }
            Where = "$($f.Name) [$($sh.Name) 시트]" }
        }
      }
    } finally { $wbS.Close($false) }
  }
  "신코·코마츠 매출내역서: " + $(if ($sk.Count) { ($sk.Keys -join ', ') } else { '없음' })

  # ── 회계 수출 파일 'YY.MM(NEW)' 시트: 거래처·Invoice·면장번호·BL일자·환율·외화·원화 (면장번호 칸 회색 = 다음 달 이월)
  $acc = @{}; $accName = ''
  $af = @(Get-ChildItem $base -File | ? { $_.Name -match '^회계.*수출' -and $_.Name -match "\($yy\.$mm2\)" -and $_.Name -notmatch '^~\$' })[0]
  if (-not $af) { throw "회계 수출 파일이 없음 (회계 ... ($yy.$mm2).xlsx)" }
  $wbA = $xl.Workbooks.Open($af.FullName, 0, $true); $accName = "$yy2.$mm2(NEW)"; $sa = $wbA.Worksheets.Item($accName)
  $na = $sa.UsedRange.Row + $sa.UsedRange.Rows.Count - 1; $va = $sa.Range("A1:K$na").Value2; $ah = $null
  for ($i = 1; $i -le [math]::Min($na, 80) -and -not $ah; $i++) { for ($j = 1; $j -le 6; $j++) { if ("$($va[$i,$j])".Trim() -eq 'Invoice') { $ah = $i; break } } }
  if (-not $ah) { $wbA.Close($false); throw "회계 $accName 시트에서 'Invoice' 머리글을 못 찾음" }
  for ($i = $ah + 1; $i -le $na; $i++) {
    if ("$($va[$i,1])" -match '총\s*계') { break }
    $inv = K $va[$i,3]; if (-not $inv) { continue }
    $c0 = [int]$sa.Range("D$i").DisplayFormat.Interior.Color; $rr = $c0 -band 255; $gg = ($c0 -shr 8) -band 255; $bb = ($c0 -shr 16) -band 255
    if (-not $acc.ContainsKey($inv)) { $acc[$inv] = [pscustomobject]@{ R = $i; Cust = (K $va[$i,1]); Mj = (K $va[$i,4]); Bl = (ToOA $va[$i,6]); Rate = (Num $va[$i,7]); Fc = (Num $va[$i,8]); Krw = (Num $va[$i,9]); Gray = ($rr -eq $gg -and $gg -eq $bb -and $rr -lt 230) } }
  }
  $wbA.Close($false)
  "$($af.Name) [$accName]: 인보이스 $($acc.Count)건"

  # ── 해외영업팀 N월마감: 당월·이월 면장 시트의 INVOICE NO → 수출신고필증 NO.
  $mg = @{}
  $gf = @(Get-ChildItem $base -File | ? { $_.Name -match "^$($mm)월마감" -and $_.Extension -match '^\.xls' -and $_.Name -notmatch '^~\$' })[0]
  if ($gf) {
    $wbG = $xl.Workbooks.Open($gf.FullName, 0, $true)
    foreach ($s in @(@($wbG.Worksheets) | ? { $_.Name -match '면장' })) {
      $u = $s.UsedRange; $vv = $u.Value2; $nr = $u.Rows.Count; $ncx = [math]::Min($u.Columns.Count, 20); $hr = $null; $ci = $null; $cm = $null
      for ($i = 1; $i -le [math]::Min($nr, 15) -and -not $hr; $i++) { for ($j = 1; $j -le $ncx; $j++) { $t0 = ("$($vv[$i,$j])" -replace '\s', ''); if ($t0 -match '^INVOICE') { $hr = $i; $ci = $j }; if ($t0 -match '수출신고필증') { $cm = $j } } }
      if (-not $hr -or -not $cm) { continue }
      for ($i = $hr + 1; $i -le $nr; $i++) { $inv = K $vv[$i,$ci]; if ($inv -match '^(DSM|\d{4}P)' -and -not $mg.ContainsKey($inv)) { $mg[$inv] = [pscustomobject]@{ Mj = (K $vv[$i,$cm]); Where = "$($s.Name)!R$($u.Row + $i - 1)" } } }
    }
    $wbG.Close($false)
    "$($gf.Name): 인보이스 $($mg.Count)건"
  } else { "  (참고) $($mLab)마감 파일이 없어 면장번호는 회계 값만 씀" }

  # ── 일별마감 (국내 라인 코드 비교용, 없으면 건너뜀)
  $daily = @(); $dSheet = "$yy2$mm2"; $dName = ''
  $df = @(Get-ChildItem $base -File | ? { $_.Name -match '일별마감' -and $_.Extension -match '^\.xls' -and $_.Name -notmatch '^~\$' })[0]
  if ($df) {
    $wbD = $xl.Workbooks.Open($df.FullName, 0, $true)
    $sD = @(@($wbD.Worksheets) | ? { $_.Name -eq $dSheet })[0]
    if ($sD) { $dn = $sD.UsedRange.Rows.Count; $dv = $sD.Range("A1:I$dn").Value2; $dName = "일별마감 $dSheet"
      for ($i = 2; $i -le $dn; $i++) { $pn = K $dv[$i,4]; if (-not $pn) { continue }; $daily += [pscustomobject]@{ R = $i; D = (ToOA $dv[$i,1]); Cust = (K $dv[$i,2]); Pn = $pn; Ln = (K $dv[$i,5]); Q = (Num $dv[$i,6]) } } }
    $wbD.Close($false)
    "$($df.Name) [$dSheet]: $($daily.Count)줄"
  } else { "  (참고) 일별마감list 파일이 없어 국내 라인 코드 비교는 건너뜀 (파일을 넣고 -Force로 다시 하면 비교함)" }

  # ── MES 원본 열기 → 그 달 출고분만 남김
  $wb = $xl.Workbooks.Open($srcFile)
  $ws = $wb.Worksheets.Item(1); $ws.Name = $sheetName
  $xl.ScreenUpdating = $false
  function HdrMap() { $n0 = $ws.UsedRange.Column + $ws.UsedRange.Columns.Count - 1; $m0 = @{}; for ($j = 1; $j -le $n0; $j++) { $h = K $ws.Cells(1,$j).Value2; if ($h -and -not $m0.ContainsKey($h)) { $m0[$h] = $j } }; $m0 }
  $HM = HdrMap
  foreach ($h in 'NO', '출고일', '구분', '고객사', '라인구분', '가공구분', '출고수량', '주조비', '가공비', '금액', '출고Lot', '메모') { if (-not $HM[$h]) { throw "MES 원본에 '$h' 칸이 없음" } }
  $nr0 = $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1
  $tot = $null; for ($r = $nr0; $r -ge 2; $r--) { if ((K $ws.Cells($r,1).Value2) -eq 'TOTAL') { $tot = $r; break } }
  $lastD = if ($tot) { $tot - 1 } else { $nr0 }
  $dCol = $ws.Range($ws.Cells(1, $HM['출고일']), $ws.Cells($lastD, $HM['출고일'])).Value2
  $del = New-Object System.Collections.ArrayList; $other = @{}
  for ($r = 2; $r -le $lastD; $r++) { $oa = ToOA $dCol[$r,1]; if (-not ($oa -and $oa -ge $firstOA -and $oa -lt $nextOA)) { [void]$del.Add($r); $lab = if ($oa) { [DateTime]::FromOADate($oa).ToString('yyyy-MM') } else { '날짜 없음' }; $other[$lab]++ } }
  if ($del.Count) {
    $i = $del.Count - 1
    while ($i -ge 0) { $e0 = $del[$i]; $s0 = $e0; while ($i -gt 0 -and $del[$i - 1] -eq $s0 - 1) { $i--; $s0 = $del[$i] }; $ws.Range("A$($s0):A$($e0)").EntireRow.Delete() | Out-Null; $i-- }
    "다른 달 출고 줄 뺌: " + (($other.Keys | sort | % { "$_ $($other[$_])줄" }) -join ', ') + " (원본은 백업\에 그대로)"
  }
  $lastD -= $del.Count; if ($tot) { $tot -= $del.Count }
  "$($mLab) 출고: $($lastD - 1)줄"
  # MES를 '오늘 기준 한 달'로 받으면 그달 초가 빠짐 (2026-11-12: 10월 MES가 10/12~11/11이라 10/1~10/11 출고 누락) → 첫 출고일이 4일 이후면 멈춤
  $dKeep = $ws.Range($ws.Cells(2, $HM['출고일']), $ws.Cells([math]::Max(2, $lastD), $HM['출고일'])).Value2
  $minOA = @(@($dKeep) | % { ToOA $_ } | ? { $_ } | sort)[0]
  if ($minOA -and ($minOA - $firstOA) -ge 3 -and -not $AllowLateStart) {
    throw "MES 첫 출고일이 $(DS $minOA)임 → $($mm)/1~$($mm)/$([DateTime]::FromOADate($minOA).Day - 1) 출고가 빠졌을 수 있음. MES에서 기간을 $($mm)/1~말일로 다시 받아 mes$($mm)월.xls로 넣어 주세요 (정말 그날부터 출고가 맞으면 -AllowLateStart)"
  }

  # ── 칸 추가: 출고일 뒤 5칸, 맨 끝 6칸
  $cD = $HM['출고일']
  $ws.Range($ws.Cells(1, $cD + 1), $ws.Cells(1, $cD + 5)).EntireColumn.Insert() | Out-Null
  $newH = @('BL날짜', '인보이스 번호', '환율', '달러', '수출면장 번호'); for ($q = 0; $q -lt 5; $q++) { $ws.Cells(1, $cD + 1 + $q).Value2 = [string]$newH[$q] }
  $lastC = $ws.UsedRange.Column + $ws.UsedRange.Columns.Count - 1
  $ws.Cells(1, $lastC).Copy($ws.Range($ws.Cells(1, $lastC + 1), $ws.Cells(1, $lastC + 6))) | Out-Null
  for ($q = 0; $q -lt 6; $q++) { $ws.Cells(1, $lastC + 1 + $q).Value2 = [string]$cats[$q] }
  $HM = HdrMap; $ncA = $lastC + 6
  $v = $ws.Range($ws.Cells(1,1), $ws.Cells($lastD, $ncA)).Value2

  # ── 행 나누기·6칸 (MES 기준)
  $lineMap = @{ '사형' = '사형'; '자동' = '자동'; '금형' = '금형 주조'; '금형주조' = '금형 주조'; '금형 주조' = '금형 주조' }
  $procMap = @{ '가공1' = '가공1(2공장)'; '가공2' = '가공2(5공장)'; '가공3' = '가공3(3공장)' }
  $nRow = $lastD - 1
  $outE = New-Object 'object[,]' $nRow, 3      # 환율·달러·수출면장 번호
  $out6 = New-Object 'object[,]' $nRow, 6
  $redE = New-Object System.Collections.ArrayList   # @(행, 칸이름)
  $red6 = @{}; $gray6 = @{}
  $mjRows = @(); $fxRows = @(); $usdRows = @(); $lnRows = @()
  $six = @{}; $grp = [ordered]@{}; $kind = @{}; $repl = @{}
  for ($r = 2; $r -le $lastD; $r++) {
    $gu = K $v[$r,$HM['구분']]; $cu = K $v[$r,$HM['고객사']]; $lot = K $v[$r,$HM['출고Lot']]; $memo = K $v[$r,$HM['메모']]
    if ($memo -match 'replacement') { $repl[$r] = $true }
    $isJ = $sk.Contains($lot) -or $lot -match '^DSM\d{4}-(SK|JS)\d' -or ($lot -match '^DSM' -and $cu -match '신코|코마츠|SHINKO|JS\s*TECH')
    $kind[$r] = if ($isJ) { 'J' } elseif ($gu -eq '해외') { 'O' } else { 'D' }
    $q0 = [double](Num $v[$r,$HM['출고수량']]); $c1 = [double](Num $v[$r,$HM['주조비']]); $c2 = [double](Num $v[$r,$HM['가공비']]); $am = Num $v[$r,$HM['금액']]
    $ln = K $v[$r,$HM['라인구분']]; $pr = K $v[$r,$HM['가공구분']]
    $lc = $lineMap[$ln]; $pc = $procMap[$pr]; $res = @{}; $note = $null
    if (-not $am) { $note = 'MES 금액 없음 → 분해 불가' }
    elseif ($lc -and $pc) {
      $lv = [math]::Round($q0 * $c1); $res[$lc] = $lv; $res[$pc] = [double]$am - $lv
      $calc = $q0 * ($c1 + $c2); if ([math]::Abs($calc - $am) -ge 0.5) { $note = '수량×(주조비+가공비) {0:N0} ≠ 금액 {1:N0} → 차액 {2:N0}을 가공에 반영' -f $calc, $am, ($am - $calc) }
    }
    elseif ($lc) { $res[$lc] = [double]$am }
    elseif ($pc) { $res[$pc] = [double]$am }
    else { $note = "MES 라인구분('$ln')·가공구분('$pr')으로 6칸을 정할 수 없음 → MES 라인을 확인해 주세요" }
    $six[$r] = $res
    if ($note) { $red6[$r] = $true; $lnRows += ,@([double]$r, $cu, $lot, 'MES 자체', "$ln/$pr", '', $note) }
    if ($kind[$r] -ne 'D' -and $lot) { if (-not $grp.Contains($lot)) { $grp[$lot] = New-Object System.Collections.ArrayList }; [void]$grp[$lot].Add($r) }
  }

  # ── 인보이스별: 환율·달러·면장·(신코·코마츠) 6칸
  $mjPat = '^\d{5}-\d{2}-\d{6}X$'
  foreach ($lot in @($grp.Keys)) {
    $rows = @($grp[$lot]); $r1 = $rows[0]; $isJ = $kind[$r1] -eq 'J'
    $cu0 = K $v[$r1,$HM['고객사']]; $a = $acc[$lot]; $src = if ($isJ) { $sk[$lot] } else { $line[$lot] }
    $carryA = $a -and ($a.Gray -or ($a.Bl -and $a.Bl -ge $nextOA))
    $isBergen = $cu0 -match 'bergen' -or ($a -and $a.Cust -match 'BERGEN')
    $isKrw = -not $isJ -and (($a -and $a.Mj -match '전자세금계산서') -or $cu0 -match 'atlas' -or ($src -and -not $src.Rate -and -not $src.Usd -and -not $isBergen))
    $cur = if ($isJ) { 'JPY' } elseif ($isKrw) { 'KRW' } else { 'USD' }
    $blU = $null; $blFrom = ''
    if ($src -and $src.Bl) { $blU = $src.Bl; $blFrom = 'src' } elseif ($a -and $a.Bl) { $blU = $a.Bl; $blFrom = 'acc' }
    $noSrcTxt = if ($isJ) { '매출내역서' } else { "$($mLab)라인" }
    # 환율
    $rate = $null; $rw = @(); $rRed = $false
    if ($isKrw) { $rw += '원화 거래(회계 환율 없음) → 환율 미기재' }
    elseif ($isBergen) {
      if ($a -and $a.Rate) { $rate = [double]$a.Rate } else { $rRed = $true; $rw += '회계에 Bergen 환율 없음' }
      if ($blFrom -eq 'acc') { $rRed = $true; $rw += "BL날짜 미확정(BL불일치사유 시트 참고) → 회계 BL날짜 $(DS $blU) 기준 회계 환율 그대로(Bergen) 적용" }
    }
    elseif (-not $blU) { $rRed = $true; $rw += "BL날짜 없음($noSrcTxt·회계 모두 없음) → 환율 미기재" + $(if (-not $a) { " ($nLab 이월 건으로 보임)" } else { '' }) }
    else {
      $g0 = Get-Fx $blU $cur
      if (-not $g0) { $rRed = $true; $rw += "서울외국환중개 $cur 환율 자료 없음 (BL $(DS $blU))" }
      else {
        $rate = $g0.Rate
        if ($blFrom -eq 'acc') { $rRed = $true; $rw += "BL날짜 미확정(BL불일치사유 시트 참고) → 회계 BL날짜 $(DS $blU) 기준 smbs 환율 적용" }
        if ($a -and $a.Rate -and [math]::Abs($a.Rate - $rate) -ge 0.00005) { $rRed = $true; $rw += "서울외국환중개 매매기준율 $rate (BL $(DS $blU)) ≠ 회계 환율 $($a.Rate) → 회계 환율을 확인해 주세요 (알림)" }
      }
    }
    foreach ($r in $rows) {
      if ($rate) { $outE[($r - 2), 0] = [double]$rate }
      if ($rRed) { [void]$redE.Add(@($r, '환율')) }
      if ($rw) { $fxRows += ,@([double]$r, (K $v[$r,$HM['고객사']]), $lot, $(if ($blU) { [double]$blU } else { '' }), $cur, $(if ($rate) { [double]$rate } else { '' }), $(if ($a -and $a.Rate) { [double]$a.Rate } else { '' }), ($rw -join ' / ')) }
    }
    # 달러(엔) = MES 금액 ÷ 환율
    $unit = if ($isJ) { '엔' } else { '달러' }; $sumU = 0.0; $uv = @{}
    foreach ($r in $rows) { $am = Num $v[$r,$HM['금액']]; if ($rate -and $am) { $uv[$r] = [math]::Round([double]$am / $rate, 2); $sumU += $uv[$r]; $outE[($r - 2), 1] = $uv[$r] } }
    $dw = @()
    if ($isKrw) { $dw += '원화 거래(환율 없음) → 달러 미기재' }
    else {
      if ($rate -and $a -and $a.Fc) { $dd = [math]::Round($sumU - $a.Fc, 2); if ([math]::Abs($dd) -gt 0.01) { $dw += ('인보이스 합계 {0:N2}{4} ≠ 회계 외화금액 {1:N2}{4} (차 {2:N2}, {3:0.00}%) - MES 금액은 내부 고정환율 기준이라 BL일 환율로 나누면 인보이스 금액과 다름' -f $sumU, $a.Fc, $dd, ($dd / $a.Fc * 100), $unit) } }
      if ($rate -and $blFrom -eq 'acc') { $dw += '환율 빨간색(BL날짜 미확정 → 회계 BL날짜 기준 환율 사용, 환율불일치 시트 참고)' }
      if (-not $rate) { $dw += "환율이 없어 $unit 계산 안 함" }
    }
    foreach ($r in $rows) {
      $am = Num $v[$r,$HM['금액']]; $w0 = @($dw); if (-not $am -and -not $isKrw) { $w0 = @('MES 금액 없음') + @($dw | ? { $_ -notmatch '^환율이 없어' }) }
      if ($w0.Count) { $usdRows += ,@([double]$r, (K $v[$r,$HM['고객사']]), $lot, $cur, $(if ($am) { [double]$am } else { '' }), $(if ($rate) { [double]$rate } else { '' }), $(if ($uv.ContainsKey($r)) { $uv[$r] } else { '' }), ($w0 -join ' / '))
        if (-not $isKrw) { [void]$redE.Add(@($r, '달러')) } }
    }
    # 수출면장 번호 = 회계 ↔ N월마감
    $g1 = $mg[$lot]; $am0 = if ($a) { $a.Mj } else { '' }; $cmpMj = if ($g1) { $g1.Mj } else { '' }; $cmpW = if ($g1) { $g1.Where } else { '' }
    $mjV = $null; $mw = $null
    if (-not $a) { $mw = if ($isJ -and -not $src) { "회계·매출내역서에 인보이스 없음 → $nLab 이월 건으로 보임(BL 미발행)" } else { '회계에 인보이스 없음' } }
    elseif ($am0 -match '전자세금계산서') { $mw = if ($cmpMj) { "$($mLab)마감 번호 $cmpMj" } else { "$($mLab)마감 번호 비어 있음" } }
    elseif ($am0 -notmatch $mjPat) { $mw = "회계 면장번호가 비어 있음('$am0')" + $(if ($carryA) { " – $nLab 이월 건(면장 미발행)" } else { '' }) }
    elseif ($isJ -and -not $g1) { $mjV = $am0 }   # 신코·코마츠는 해외영업팀 마감 파일에 없음(국내영업 담당) → 회계 번호 그대로
    elseif (-not $g1 -or -not $cmpMj) { $mjV = $am0; $mw = $(if ($g1) { "$($mLab)마감 번호 비어 있음" } else { "$($mLab)마감 파일에 인보이스 없음" }) + $(if ($carryA) { " / $nLab 이월 건(회계 면장번호 회색·BL $(DS $a.Bl))" } else { '' }) }
    elseif ($cmpMj -ne $am0) { $mjV = $am0; $mw = "면장번호 다름: 회계 $am0 / $($mLab)마감 $cmpMj" }
    else { $mjV = $am0 }
    foreach ($r in $rows) {
      if ($mjV) { $outE[($r - 2), 2] = [string]$mjV }
      if ($mw) { [void]$redE.Add(@($r, '수출면장 번호')); $mjRows += ,@([double]$r, (K $v[$r,$HM['고객사']]), $lot, $am0, $(if ($a) { "$accName!R$($a.R)" } else { '' }), $cmpMj, $cmpW, $mw) }
    }
    # 신코·코마츠: 매출내역서 금액(사형+스크랩, 금형+스크랩, 가공3)을 MES 라인별 금액 비율로 줄에 나눔
    if ($isJ -and $src) {
      $orig6 = @{}; foreach ($r in $rows) { $orig6[$r] = $six[$r]; $six[$r] = @{ '__sk' = $true } }
      foreach ($cat in @($src.Split.Keys)) {
        $tot0 = [double]$src.Split[$cat]; if (-not $tot0) { continue }
        $wts = @($rows | % { [double]$orig6[$_][$cat] }); $ws0 = ($wts | measure -Sum).Sum
        if (-not $ws0) { $wts = @($rows | % { [double](Num $v[$_,$HM['금액']]) }); $ws0 = ($wts | measure -Sum).Sum }
        if (-not $ws0) { $wts = @(1.0) + @(1..($rows.Count) | select -Skip 1 | % { 0.0 }); $ws0 = 1.0 }
        $endI = $rows.Count - 1; while ($endI -gt 0 -and -not $wts[$endI]) { $endI-- }
        $acc0 = 0.0; $put = @{}
        for ($k = 0; $k -lt $rows.Count; $k++) { $val = if ($k -eq $endI) { $tot0 - $acc0 } elseif ($wts[$k]) { [math]::Round($tot0 * $wts[$k] / $ws0) } else { 0.0 }; $acc0 += $val; $put[$rows[$k]] = $val }
        foreach ($r in $rows) { $six[$r][$cat] = $put[$r] }
      }
      foreach ($r in $rows) { $gray6[$r] = $true; $red6.Remove($r) }
    }
    # 해외: MES 6칸 비율 ↔ N월라인 비율 (0.5%p 넘게 다르면 참고)
    if (-not $isJ -and $src -and $src.Split.Count) {
      $ms = @{}; foreach ($c0 in $cats) { $ms[$c0] = [double](($rows | ? { -not $repl[$_] } | % { [double]$six[$_][$c0] } | measure -Sum).Sum) }
      $tm = ($ms.Values | measure -Sum).Sum; $tl = ($src.Split.Values | measure -Sum).Sum
      if ($tm -and $tl) {
        $pm0 = @{}; $pl0 = @{}; foreach ($c0 in $cats) { $pm0[$c0] = 100 * $ms[$c0] / $tm; $pl0[$c0] = 100 * [double]$src.Split[$c0] / $tl }
        $dif = @($cats | sort | ? { [math]::Abs($pm0[$_] - $pl0[$_]) -gt 0.5 })
        if ($dif.Count) {
          $tx = { param($p) (@($cats | sort | ? { [math]::Round($p[$_], 1) -ne 0 } | % { '{0} {1:0.0}%' -f $_, $p[$_] })) -join ', ' }
          $lnRows += ,@('', $cu0, $lot, "$($mLab)라인", (& $tx $pm0), (& $tx $pl0), ('비율 차이(0.5%p 초과): ' + ((@($dif | % { '{0} MES {1:0.0}% vs 참고 {2:0.0}%' -f $_, $pm0[$_], $pl0[$_] })) -join ' / ')))
        }
      }
    }
  }
  $lnRows = @($lnRows | ? { -not ($_[0] -is [double] -and $gray6.ContainsKey([int]$_[0])) })
  $nSk = @($gray6.Keys).Count
  if ($nSk) { $lnRows += ,@('', '신코·코마츠', '', '매출내역서 기준 분해', '', '', "신코·코마츠 $nSk 행은 매출내역서 엑셀 합계 기준으로 6칸을 나눔(연회색 칸). 이 행들의 6칸 합은 MES 금액이 아니라 매출내역서(BL 환율) 금액") }

  # ── 국내: MES 라인구분 ↔ 일별마감 라인 코드(사·자·금)
  if ($daily.Count) {
    $kN = { param($x) ("$x" -replace '\(주\)|㈜|주식회사|\(유\)|\s', '').ToUpper() }
    $byPn = @{}; foreach ($d0 in $daily) { $kk = $d0.Pn.ToUpper(); if (-not $byPn.ContainsKey($kk)) { $byPn[$kk] = New-Object System.Collections.ArrayList }; [void]$byPn[$kk].Add($d0) }
    $code = @{ '사형' = '사'; '자동' = '자'; '금형' = '금' }
    for ($r = 2; $r -le $lastD; $r++) {
      if ($kind[$r] -ne 'D') { continue }
      $pn = (K $v[$r,$HM['품번']]).ToUpper(); if (-not $pn -or -not $byPn.ContainsKey($pn)) { continue }
      $cuK = & $kN $v[$r,$HM['고객사']]; $dt = ToOA $v[$r,$HM['출고일']]; $qq = [double](Num $v[$r,$HM['출고수량']])
      $cand = @($byPn[$pn] | ? { $k1 = & $kN $_.Cust; $k1 -and $cuK -and ($k1.StartsWith($cuK.Substring(0, [math]::Min(2, $cuK.Length))) -or $cuK.StartsWith($k1.Substring(0, [math]::Min(2, $k1.Length)))) })
      $hit = @($cand | ? { $_.D -eq $dt -and $_.Q -eq $qq })[0]; if (-not $hit) { $hit = @($cand | ? { $_.D -eq $dt })[0] }; if (-not $hit -and $cand.Count -eq 1) { $hit = $cand[0] }
      if (-not $hit -or -not $hit.Ln) { continue }
      $ml = K $v[$r,$HM['라인구분']]; $dl = $hit.Ln.Substring(0, 1)
      $why0 = if ('사', '자', '금' -notcontains $dl) { "일별마감 라인 코드 '$($hit.Ln)'는 사/자/금이 아님 → 비교 불가(MES 기준 분해, 참고)" } elseif ($code[$ml] -ne $dl) { "라인 다름: MES 라인구분 '$ml' / 일별마감 라인 '$($hit.Ln)' → 어느 쪽이 맞는지 확인해 주세요" } else { $null }
      if ($why0) { $lnRows += ,@([double]$r, (K $v[$r,$HM['고객사']]), (K $v[$r,$HM['출고Lot']]), "일별마감 $dSheet!R$($hit.R)", "MES 라인구분: $ml", "일별마감 라인: $($hit.Ln)", $why0) }
    }
  }

  # ── 시트에 쓰기
  for ($r = 2; $r -le $lastD; $r++) { $res = $six[$r]; for ($q = 0; $q -lt 6; $q++) { $x0 = $res[$cats[$q]]; if ($x0) { $out6[($r - 2), $q] = [double]$x0 } } }
  $cE = $HM['환율']
  $ws.Range($ws.Cells(2, $HM['수출면장 번호']), $ws.Cells($lastD, $HM['수출면장 번호'])).NumberFormat = '@'
  $ws.Range($ws.Cells(2, $cE), $ws.Cells($lastD, $cE + 2)).Value2 = $outE
  $c6 = $HM['사형']; $ws.Range($ws.Cells(2, $c6), $ws.Cells($lastD, $c6 + 5)).Value2 = $out6
  foreach ($e in $redE) { $ws.Cells($e[0], $HM[$e[1]]).Interior.Color = $RED }
  foreach ($r in @($red6.Keys)) { $ws.Range($ws.Cells($r, $c6), $ws.Cells($r, $c6 + 5)).Interior.Color = $RED }
  foreach ($r in @($gray6.Keys)) { $ws.Range($ws.Cells($r, $c6), $ws.Cells($r, $c6 + 5)).Interior.Color = $COLGRAY }
  $ws.Range($ws.Cells(2, $cE), $ws.Cells($lastD + 1, $cE)).NumberFormat = '[<50]#,##0.0000;#,##0.00'
  $ws.Range($ws.Cells(2, $HM['달러']), $ws.Cells($lastD + 1, $HM['달러'])).NumberFormat = '#,##0.00'
  $ws.Range($ws.Cells(2, $c6), $ws.Cells($lastD + 1, $c6 + 5)).NumberFormat = '#,##0'
  if ($tot) {
    foreach ($h in @('출고수량', '주조비', '가공비', '금액') + $cats) { $cc = $HM[$h]; $cl = ($ws.Cells(1,$cc).Address($false,$false) -replace '\d', ''); $ws.Cells($tot, $cc).Formula = "=SUM(${cl}2:${cl}$lastD)" }
    $ws.Range($ws.Cells($tot, $c6), $ws.Cells($tot, $c6 + 5)).NumberFormat = '#,##0'
  }
  foreach ($cc in $cE, $HM['달러'], $HM['수출면장 번호']) { $ws.Columns($cc).AutoFit() | Out-Null }

  function NewSheet($name, $titles, $rowsArr, $fmts) {
    $n = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)); $n.Name = $name
    for ($j = 0; $j -lt $titles.Count; $j++) { $n.Cells(1, $j + 1).Value2 = [string]$titles[$j] }
    $hh = $n.Range($n.Cells(1,1), $n.Cells(1,$titles.Count)); $hh.Font.Bold = $true; $hh.Interior.Color = $HEAD
    for ($j = 0; $j -lt $titles.Count; $j++) { $n.Range($n.Cells(2, $j + 1), $n.Cells([math]::Max(2, $rowsArr.Count + 1), $j + 1)).NumberFormat = $(if ($fmts[$j]) { $fmts[$j] } else { '@' }) }
    for ($i = 0; $i -lt $rowsArr.Count; $i++) { for ($j = 0; $j -lt $rowsArr[$i].Count; $j++) { $x = $rowsArr[$i][$j]; if ($x -eq $null -or "$x" -eq '') { continue }; if ($x -is [double]) { $n.Cells($i + 2, $j + 1).Value2 = [double]$x } else { $n.Cells($i + 2, $j + 1).Value2 = [string]$x } } }
    $n.Columns.AutoFit() | Out-Null; $last0 = $titles.Count; if ($n.Columns($last0).ColumnWidth -gt 80) { $n.Columns($last0).ColumnWidth = 80; $n.Columns($last0).WrapText = $true }
    $n
  }
  $null = NewSheet '면장불일치' @('MES 행', '고객사', '인보이스 번호(출고Lot)', '회계 면장번호(기준)', '회계 위치', "비교 번호($($mLab)마감)", '비교 위치', '사유') $mjRows @('0')
  $null = NewSheet '환율불일치' @('MES 행', '고객사', '인보이스 번호(출고Lot)', 'BL날짜', '통화', 'smbs 매매기준율', '회계 환율', '사유') $fxRows @('0', $null, $null, 'yyyy-mm-dd', $null, '#,##0.0000', '#,##0.0000')
  $null = NewSheet '달러불일치' @('MES 행', '고객사', '인보이스 번호(출고Lot)', '통화', 'MES 금액', '환율', '계산 달러(엔)', '사유') $usdRows @('0', $null, $null, $null, '#,##0', '#,##0.0000', '#,##0.00')
  $null = NewSheet '라인구분불일치' @('MES 행', '고객사', '인보이스/Lot', '비교 대상', 'MES', '참고자료', '사유') $lnRows @('0')
  $ws.Activate() | Out-Null; $ws.Range('A2').Select() | Out-Null
  $xl.ScreenUpdating = $true

  # 원본 자리에 .xls(엑셀 97-2003)로 저장 (임시 파일 → 옮김)
  $tmp = "$base\백업\_준비중_mes$($mm)월.xls"; if (Test-Path $tmp) { Remove-Item $tmp -Force }
  $wb.SaveAs($tmp, 56); $wb.Close($false); $wb = $null
  Move-Item $tmp $mesPath -Force
  "저장: $mesPath"
  "  해외 $(@($kind.Values | ? { $_ -eq 'O' }).Count)줄 / 신코·코마츠 $(@($kind.Values | ? { $_ -eq 'J' }).Count)줄 / 국내 $(@($kind.Values | ? { $_ -eq 'D' }).Count)줄, 인보이스 $($grp.Count)건"
  "  확인 시트: 면장불일치 $($mjRows.Count) / 환율불일치 $($fxRows.Count) / 달러불일치 $($usdRows.Count) / 라인구분불일치 $($lnRows.Count)"
  "다음: & `"C:\ai\도구\월마감_반영.ps1`" -YM $YM"
} finally {
  if ($wb) { try { $wb.Close($false) } catch {} }
  $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
