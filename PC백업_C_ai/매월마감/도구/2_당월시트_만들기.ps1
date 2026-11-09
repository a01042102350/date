$ErrorActionPreference = 'Stop'
function SetV($cell, $v) { $a = New-Object 'object[,]' 1, 1; if ($null -ne $v) { $v = $v.psobject.BaseObject }; $a[0, 0] = $v; $cell.Value2 = $a }

$dir = 'C:\ai\매월마감\202610'
$sp = Split-Path $MyInvocation.MyCommand.Path
$orig = Join-Path $dir '2026년 수출 (2026.10).xlsx'
$newF = Join-Path $dir '2026년 수출 (2026.10)_10월반영.xlsx'
New-Item -ItemType Directory -Force "$dir\백업" | Out-Null
Copy-Item $orig "$dir\백업\2026년 수출 (2026.10)_원본_$(Get-Date -f yyyyMMdd_HHmmss).xlsx"
$KEEP = @{}
if (Test-Path $newF) {
  Copy-Item $newF "$dir\백업\2026년 수출 (2026.10)_10월반영_다시만들기전_$(Get-Date -f yyyyMMdd_HHmmss).xlsx"
  $kx = New-Object -ComObject Excel.Application; $kx.DisplayAlerts = $false
  try { $kb = $kx.Workbooks.Open($newF, 0, $true); $ks = $kb.Worksheets.Item('26.10(NEW)')
    for ($q = 30; $q -le 300; $q++) { if ($ks.Cells($q,1).Text -like '총*계*') { break }; $ki = "$($ks.Cells($q,3).Value2)".Trim(); if (-not $ki) { continue }
      $KEEP[$ki] = @{ 면장 = $ks.Cells($q,4).Value2; 유니패스 = $ks.Cells($q,5).Value2; 메모 = $ks.Cells($q,13).Value2 } }
    $kb.Close($false) } finally { $kx.Quit() }
  "기존 파일에서 읽어 둔 줄: $($KEEP.Count) (유니패스 입력 $(@($KEEP.Values | ? { $_.유니패스 }).Count)건)" }
Copy-Item $orig $newF -Force
$LB = Import-Csv "$sp\c_B.csv" -Encoding UTF8 | ? { $_.인보이스 -and $_.인보이스 -notmatch '매출마감일' }
$SA = Import-Csv "$sp\c_A.csv" -Encoding UTF8
$P9 = Import-Csv "$sp\c_S9.csv" -Encoding UTF8 | ? { $_.인보이스 }
$RT = @{}; Import-Csv "$sp\c_FX.csv" -Encoding UTF8 | ? 환율 | % { $RT[$_.날짜] = [double]$_.환율 }
$prevInv = @{}; foreach ($p in $P9) { $prevInv[$p.인보이스] = $p }

function CustOf($x) { $i = $x.인보이스
  if ($x.거래처 -like '툴링비*' -or $x.거래처 -like '*관세*') { if ($i -match 'ME01') { return 'MEPPI' } elseif ($i -match '관세') { return 'ALLISON' } else { return 'MERCURY U' } }
  switch -regex ($x.거래처) { '^DOMETIC' { 'DOMETIC' } '(?i)^POLARIS' { 'POLARIS' } 'MERCURY CHINA' { 'MERCURY C'; break } '^MERCURY' { 'MERCURY U' } '^BERGEN' { 'BERGEN ENG' }
    '^ALLISON' { 'ALLISON' } '^SAM' { 'SAM' } '^EMERSON' { 'EMERSON' } 'AMAT' { 'AMAT' } '^ATLAS' { 'ATLAS KOREA' } default { "?$($x.거래처)" } } }

# ── 10월 줄 목록 ──
$rows = New-Object Collections.Generic.List[object]
foreach ($x in $LB) { $c = CustOf $x
  $kind = if ($x.구분 -eq '이월') { '이월' } elseif ($x.인보이스 -match '관세') { '관세' } elseif ($x.거래처 -like '툴링비*') { '툴링' } elseif ($prevInv.ContainsKey($x.인보이스)) { '전월' } else { '당월' }
  $inv = $x.인보이스; if ($kind -eq '툴링' -and $inv -notmatch '툴링') { $inv = "$inv(툴링비)" }
  $dno = if ($kind -eq '관세') { '관세' } elseif ($x.면장) { $x.면장 } elseif ($c -eq 'ATLAS KOREA') { '전자세금계산서' } else { '툴링비' }
  $rows.Add([pscustomobject]@{ 거래처 = $c; 종류 = $kind; 인보이스 = $inv; 원인보이스 = $x.인보이스; 면장 = $dno; BL = $x.BL; 환율 = $x.환율; 외화 = [double]$x.외화; 원화 = [double]$x.원화; 값고정 = ($c -eq 'BERGEN ENG'); 구월 = $false; 빈줄 = $false }) }
# 상세에만 있는 8월 오기재 관세 차액분 (관세 맨 앞)
$cor = $SA | ? { $_.인보이스 -like '*오기재*' } | Select -First 1
if ($false -and $cor) { $rows.Add([pscustomobject]@{ 거래처 = 'ALLISON'; 종류 = '관세'; 인보이스 = '8월 오기재 관세 차액분(DSM2608-AL05 관세)'; 원인보이스 = $cor.인보이스; 면장 = '관세'; BL = $cor.BL; 환율 = $cor.환율; 외화 = [double]$cor.외화; 원화 = [double]$cor.원화; 값고정 = $true; 구월 = $false; 빈줄 = $false; 앞 = $true }) }
# 9월 이월 중 해외 자료에 없는 것(코마츠 JS03) → 전월이월
foreach ($p in $P9) { if ($false -and -not ($rows | ? 원인보이스 -eq $p.인보이스)) { $rows.Add([pscustomobject]@{ 거래처 = $p.거래처; 종류 = '전월'; 인보이스 = $p.인보이스; 원인보이스 = $p.인보이스; 면장 = $p.면장; BL = '9월 시트'; 환율 = $null; 외화 = [double]$p.외화; 원화 = [double]$p.원화; 값고정 = $false; 구월 = $true; 빈줄 = $false }) } }
# 빈 줄(신코·코마츠 10월분 직접 입력용)
# 신코·코마츠: 10월 매출내역서 시트(신코202610, 코마츠202610)만 — 엔화, 환율은 기초자료 값
$JP = @(Import-Csv "$sp\c_JP.csv" -Encoding UTF8)
foreach ($nm in 'SHINKO', 'JS TECH(코마츠)') { $g = @($JP | ? 고객 -eq $nm)
  if ($g.Count -eq 0) { $rows.Add([pscustomobject]@{ 거래처 = $nm; 종류 = '당월'; 인보이스 = ''; 원인보이스 = ''; 면장 = ''; BL = ''; 환율 = $null; 외화 = $null; 원화 = $null; 값고정 = $false; 구월 = $false; 빈줄 = $true }) }
  foreach ($x in $g) { $kind = if ($x.인보이스 -match 'DSM(\d{4})' -and [int]$matches[1] -lt 2610) { '전월' } else { '당월' }
    $rows.Add([pscustomobject]@{ 거래처 = $nm; 종류 = $kind; 인보이스 = $x.인보이스; 원인보이스 = $x.인보이스; 면장 = ''; BL = $x.BL; 환율 = [double]$x.환율; 외화 = [double]$x.엔화; 원화 = [double]$x.원화; 값고정 = $false; 구월 = $false; 빈줄 = $false; 엔화 = $true }) } }
$xl = New-Object -ComObject Excel.Application; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
try {
  $wb = $xl.Workbooks.Open($newF)
  $src = $wb.Worksheets.Item('26.09(NEW)')
  $nameOf = @{}; $fac = @{}
  for ($r = 9; $r -le 21; $r++) { $n = "$($src.Cells($r,3).Value2)"; if ($n -and $src.Cells($r,2).Text -notmatch '소계') { $k = $n.Trim(); if ($k -like 'ALLISON*') { $k = 'ALLISON' }; $fac[$k] = $src.Cells($r,2).Text } }
  $hs = @{}; $detOrder = @()
  for ($r = 31; $r -le 132; $r++) { $n = "$($src.Cells($r,1).Value2)"; $k = $n.Trim(); if (-not $k) { continue }
    if (-not $nameOf[$k]) { $nameOf[$k] = $n }; if ($detOrder -notcontains $k) { $detOrder += $k }
    $m = "$($src.Cells($r,13).Value2)"; if ($m -and -not $hs[$k] -and $src.Cells($r,13).Interior.Color -eq 16777215) { $hs[$k] = $m } }
  $nameOf['MEPPI'] = 'MEPPI'; $fac['MEPPI'] = '마산2'
  $ai = [array]::IndexOf($detOrder, 'ALLISON'); $detOrder = @($detOrder[0..($ai - 1)]) + @('MEPPI') + @($detOrder[$ai..($detOrder.Count - 1)])
  $kOrd = @{ '전월' = 0; '당월' = 1; '툴링' = 2; '관세' = 3; '이월' = 4 }
  $i = 0; foreach ($x in $rows) { $x | Add-Member 순번 $(if ($x.앞) { -1 } else { $i }) -Force; $i++ }
  $bad = @($rows | ? { [array]::IndexOf($detOrder, $_.거래처) -lt 0 }); if ($bad) { throw "거래처 못 찾음: $(($bad | % 거래처) -join ',')" }
  $sorted = @($rows | Sort-Object { [array]::IndexOf($detOrder, $_.거래처) }, { $kOrd[$_.종류] }, 순번)

  # ── 시트 복사 ──
  $src.Copy([Type]::Missing, $src); $ws = $wb.Worksheets.Item($src.Index + 1); $ws.Name = '26.10(NEW)'
  SetV $ws.Cells(1,1) ('※ 2026년 10월 마감내역 ※')
  # 요약표에 MEPPI 줄 (ALLISON 아래)
  $alR = 0; for ($r = 9; $r -le 21; $r++) { if ($ws.Cells($r,3).Text -like 'ALLISON*') { $alR = $r } }
  $ws.Rows($alR + 1).Insert() | Out-Null; $ws.Rows($alR).Copy($ws.Rows($alR + 1)) | Out-Null
  SetV $ws.Cells($alR + 1, 2) ('마산2'); SetV $ws.Cells($alR + 1, 3) ('MEPPI')
  # 상세 영역
  $hdr = 0; for ($r = 20; $r -le 40; $r++) { if ($ws.Cells($r,1).Text -like '◆*') { $hdr = $r } }
  $d0 = $hdr + 3; $t0 = 0; for ($r = $d0; $r -le 200; $r++) { if ($ws.Cells($r,1).Text -like '총*계*') { $t0 = $r; break } }
  $ws.Range("A$($d0):A$($t0 - 1)").EntireRow.Delete() | Out-Null
  $tpl = @{ '전월¥' = 31; '당월¥' = 32; '이월¥' = 38; '전월$' = 51; '당월$' = 68; '툴링$' = 109; '관세$' = 109; '이월$' = 110 }
  $r = $d0
  foreach ($x in $sorted) { $yen = ($x.거래처 -eq 'SHINKO' -or $x.거래처 -eq 'JS TECH(코마츠)'); $cur = if ($yen) { '¥' } else { '$' }
    $src.Rows($tpl["$($x.종류)$cur"]).Copy() | Out-Null; $ws.Rows($r).Insert(-4121) | Out-Null
    $ws.Range("A$($r):T$r").ClearContents() | Out-Null; $ws.Range("M$($r):T$r").Interior.ColorIndex = -4142
    SetV $ws.Cells($r,1) ($nameOf[$x.거래처]); SetV $ws.Cells($r,2) ($fac[$x.거래처])
    if (-not $x.빈줄) {
      SetV $ws.Cells($r,3) ($x.인보이스); SetV $ws.Cells($r,4) ($x.면장)
      if ($x.구월) { for ($q = 31; $q -le 132; $q++) { if ("$($src.Cells($q,3).Value2)".Trim() -eq $x.인보이스) { $src.Range("E$($q):H$q").Copy() | Out-Null; $ws.Range("E$r").PasteSpecial(-4163) | Out-Null } } }
      else { SetV $ws.Cells($r,6) ([DateTime]::Parse($x.BL).ToOADate())
        if ($false) { }
        elseif ($x.엔화) { SetV $ws.Cells($r,7) ($x.환율) } elseif ($x.거래처 -ne 'ATLAS KOREA') { $ws.Cells($r,7).Formula = "=VLOOKUP(F$r,환율!A:B,2,0)" }
        SetV $ws.Cells($r,8) ($x.외화) } }
    if ($x.거래처 -eq 'ATLAS KOREA' -or $x.값고정) { SetV $ws.Cells($r,9) ($x.원화) } else { $ws.Cells($r,9).Formula = "=ROUNDDOWN(H$r*G$r,0)" }
    if ($x.종류 -eq '이월') { $ws.Cells($r,11).Formula = "=I$r" } else { $ws.Cells($r,10).Formula = "=I$r" }
    # HS코드는 기초자료에 없어 넣지 않음
    $kv = $KEEP[$x.인보이스]; if ($kv) { if ($kv.유니패스) { SetV $ws.Cells($r,5) ($kv.유니패스) }; if ($kv.메모) { SetV $ws.Cells($r,13) ($kv.메모) }; if ($kv.면장 -and -not $x.면장) { SetV $ws.Cells($r,4) ($kv.면장) } }
    $x | Add-Member 행 $r -Force; $r++ }
  $tot = $r
  $lastCol = 11; for ($cc = 11; $cc -le 13; $cc++) { if ($src.Cells(94, $cc).Borders(9).Weight -eq -4138) { $lastCol = $cc } }
  $ws.Range($ws.Cells($d0, 1), $ws.Cells($tot - 1, $lastCol)).Borders(12).Weight = 2
  $ws.Range($ws.Cells($d0, 1), $ws.Cells($tot - 1, $lastCol)).Borders(12).LineStyle = 1
  for ($q = $d0; $q -lt $tot; $q++) { $bb = $ws.Range($ws.Cells($q, 1), $ws.Cells($q, $lastCol)).Borders(9); $bb.LineStyle = 1
    if ("$($ws.Cells($q,1).Value2)".Trim() -ne "$($ws.Cells($q + 1,1).Value2)".Trim()) { $bb.Weight = -4138 } else { $bb.Weight = 2 } }
  foreach ($cc in 'I','J','K') { $ws.Range("$cc$tot").Formula = "=SUM($cc$($d0):$cc$($tot-1))" }
  $firstOv = ($sorted | ? { $_.거래처 -ne 'SHINKO' -and $_.거래처 -ne 'JS TECH(코마츠)' } | Select -First 1).행
  $ws.Cells($tot + 1, 10).Formula = "=SUM(I$($firstOv):I$($tot-1))-K$tot"
  $dm = ($sorted | ? 거래처 -eq 'DOMETIC' | Select -First 1).행; $k = 14; foreach ($t in '면장','IN','PK','BL','기타') { SetV $ws.Cells($dm,$k) ($t); $ws.Cells($dm,$k).Interior.Color = 65535; $k++ }

  # ── 요약표 ──
  function Rg($list, $col) { $rs = @($list | % 행 | Sort-Object); if ($rs.Count -eq 0) { return $null }; "SUM($col$($rs[0]):$col$($rs[-1]))" }
  $sumEnd = 0; for ($q = 9; $q -le 25; $q++) { if ($ws.Cells($q,2).Text -match '해외영업') { $sumEnd = $q } }
  $custRows = @()
  for ($q = 9; $q -lt $sumEnd; $q++) { $cn = "$($ws.Cells($q,3).Value2)"; if (-not $cn) { continue }; $key = $cn.Trim(); if ($key -like 'ALLISON*') { $key = 'ALLISON' }
    $g = @($sorted | ? 거래처 -eq $key); $custRows += [pscustomobject]@{ 행 = $q; 고객 = $key; 공장 = $ws.Cells($q,2).Text }
    $fD = Rg @($g | ? 종류 -eq '전월') 'J'; $fE = Rg @($g | ? { $_.종류 -in '당월','툴링','관세' }) 'J'; $fH = Rg @($g | ? 종류 -eq '이월') 'K'
    if ($fD) { $ws.Cells($q,4).Formula = "=$fD" } else { SetV $ws.Cells($q,4) (0) }
    if ($fE) { $ws.Cells($q,5).Formula = "=$fE" } else { SetV $ws.Cells($q,5) (0) }
    if ($fH) { $ws.Cells($q,8).Formula = "=$fH" } else { SetV $ws.Cells($q,8) (0) }
    $crit = if ($key -eq 'ALLISON') { '"ALLISON"' } else { "C$q" }; $ws.Cells($q,9).Formula = "=SUMIF(`$A`$$($d0):`$A`$$($tot),$crit,`$I`$$($d0):`$I`$$($tot))"
    $ws.Cells($q,10).Formula = "=SUM(D$($q):G$q)" }
  # 공장별 합계(K): 같은 공장이 이어지는 구간 마지막 줄
  foreach ($cr in $custRows) { if ($cr.고객 -ne 'ATLAS KOREA') { $ws.Cells($cr.행, 11).ClearContents() | Out-Null } }
  $runStart = $null
  for ($j = 0; $j -lt $custRows.Count; $j++) { $f = $custRows[$j]
    if (-not $runStart) { $runStart = $f.행 }; $nx = if ($j + 1 -lt $custRows.Count) { $custRows[$j + 1] } else { $null }
    if (-not $nx -or $nx.공장 -ne $f.공장 -or $nx.행 -ne $f.행 + 1) { $at2 = if ($f.고객 -eq 'ATLAS KOREA') { $f.행 - 1 } else { $f.행 }; $ws.Cells($at2, 11).Formula = $(if ($runStart -eq $f.행) { "=J$($f.행)" } else { "=SUM(J$($runStart):J$($f.행))" }); $runStart = $null } }
  $atl = ($custRows | ? 고객 -eq 'ATLAS KOREA').행; SetV $ws.Cells($atl, 11) (' 국내매출 ')
  # ── TOOLING 현황 ──
  $tt = $sumEnd + 2; $toolG = @($sorted | ? 종류 -eq '툴링'); $dutyG = @($sorted | ? 종류 -eq '관세')
  $lines = @([pscustomobject]@{ 공장 = '창원'; 고객 = $nameOf['SHINKO']; 건 = 0; 식 = $null })
  foreach ($gg in $toolG | Group-Object 거래처) { $lines += [pscustomobject]@{ 공장 = $fac[$gg.Name]; 고객 = $nameOf[$gg.Name]; 건 = $gg.Count; 식 = (Rg @($gg.Group) 'J') } }
  $lines += [pscustomobject]@{ 공장 = '마산2'; 고객 = 'ALLISON'; 건 = $dutyG.Count; 식 = (Rg $dutyG 'J') }
  if ($lines.Count -gt 4) { throw '툴링 표 줄 부족' }
  $tEnd = $tt + 3; 
  $q = $tt; foreach ($l in $lines) { SetV $ws.Cells($q,2) ($l.공장); SetV $ws.Cells($q,3) ($l.고객); SetV $ws.Cells($q,4) ($l.건); if ($l.식) { $ws.Cells($q,5).Formula = "=$($l.식)" } else { SetV $ws.Cells($q,5) (0) }; $q++ }
  $ws.Cells($tt,8).Formula = "=SUM(E$($tt):G$tEnd)"
  SetV $ws.Cells($tt,9) (" TOOLING $($toolG.Count)건 "); SetV $ws.Cells($tt+1,9) (" 관세 $($dutyG.Count)건 "); SetV $ws.Cells($tt+2,9) (' 선급,중자0건 '); SetV $ws.Cells($tt+3,9) (" 총 TOOLING-$($toolG.Count + $dutyG.Count)건 ")
  $ws.Cells($tt,11).Formula = "=J$atl"; $ws.Cells($tt+1,11).Formula = "=SUM(J9:J10,J12:J$($atl-1))"
  # 오른쪽 작은 표
  $yr = @($sorted | ? { $_.거래처 -eq 'SHINKO' -or $_.거래처 -eq 'JS TECH(코마츠)' } | % 행 | Sort-Object)
  $ws.Range('P21').Formula = "=SUM(H$($yr[0]):H$($yr[-1]),)"; $ws.Range('Q21').Formula = "=SUM(J$($yr[0]):J$($yr[-1]))"
  $at = @($sorted | ? 거래처 -eq 'ATLAS KOREA' | % 행 | Sort-Object); $ws.Range('P25').Formula = "=SUM(H$($at[0]):H$($at[-1]))"; $ws.Range('Q25').Formula = "=SUM(I$($at[0]):I$($at[-1]))"

  # ── 환율 시트: 비어 있는 날짜에 라인별 파일 환율 ──
  $hx = $wb.Worksheets.Item('환율'); $filled = @(); $hv = $hx.Range('A1:B3040').Value2
  for ($q = 3; $q -le 3040; $q++) { $dv = $hv[$q,1]; if ($dv -is [double]) { $k2 = [DateTime]::FromOADate($dv).ToString('yyyy-MM-dd'); if ($RT[$k2] -and ("$($hv[$q,2])".Trim() -eq '')) { SetV $hx.Cells($q,2) ($RT[$k2]); $filled += $k2 } } }
  "환율 넣은 날: $($filled.Count)일 ($($filled[0]) ~ $($filled[-1]))"
  $xl.Calculate()
  foreach ($x in $sorted) { $x | Add-Member 계산 ($ws.Cells($x.행, 9).Value2) -Force; $x | Add-Member 계산환율 ($ws.Cells($x.행, 7).Value2) -Force }
  $sorted | Select 행, 거래처, 종류, 인보이스, 원인보이스, 면장, BL, 환율, 계산환율, 외화, 원화, 계산 | Export-Csv "$sp\b_rows.csv" -NoTypeInformation -Encoding UTF8
  "요약표:"; for ($q = 8; $q -le $tEnd; $q++) { $l = @(); foreach ($c in 2,3,4,5,8,9,10,11) { $t = $ws.Cells($q,$c).Text.Trim(); if ($t) { $l += $t } }; "$q  " + ($l -join ' | ') }
  "상세 $d0 ~ $($tot-1), 합계줄 $tot : I=$($ws.Cells($tot,9).Text) J=$($ws.Cells($tot,10).Text) K=$($ws.Cells($tot,11).Text)  아래줄 J=$($ws.Cells($tot+1,10).Text)"
  $ws.Activate(); $ws.Range('A1').Select() | Out-Null
  $wb.Save(); $wb.Close($false); "저장: $newF"
} finally { $xl.ScreenUpdating = $true; $xl.Quit() }
