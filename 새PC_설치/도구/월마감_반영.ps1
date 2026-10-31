# 2026-10-28 작업 8·9·10
#  8) BL날짜/인보이스: 해외=9월라인, 국내=출고일(출하일), 신코·코마츠=PDF. 회계 기준으로 다르면 빨간색 + 사유 시트
#  9) 면장별합계 + 면장별합계2 → 면장별합계 하나로 (신코·코마츠는 PDF 기준), 회계 기준 비교
# 10) mes9월: E(환율) 뒤 '기초자료환율', 달러 뒤 '기초자료달러' 열 추가 (PDF / 9월라인 값)
$base = "C:\ai\기초자료\202609"   # 월별 폴더: C:\ai\기초자료\YYYYMM (받은 자료 + mes + 결과 HTML + 백업\)
$mesPath = "C:\ai\기초자료\202609\mes9월.xls"
$backup = "C:\ai\기초자료\202609\백업\mes9월_백업_20261028.xls"
function Num($x) { $d = 0.0; if ([double]::TryParse("$x", [ref]$d)) { $d } else { $null } }
function K($x) { "$x".Trim() }
function DS($n) { if ($n) { [DateTime]::FromOADate($n).ToString('yyyy-MM-dd') } else { '없음' } }
function NZ0($x) { ((("$x" -replace '[^A-Za-z0-9]', '').ToUpper()) -replace '(?<=[A-Z])0+(?=\d)', '') }
$RED = 255; $PINK = 13421823; $HEAD = 14277081

if (-not (Test-Path $backup)) { Copy-Item $mesPath $backup }

# PDF 매출내역서 판독값 (신코.pdf p1, js.pdf p1) — BL날짜는 각 B/L On board 일자와도 일치 확인
$pdf = [ordered]@{
  'DSM2608-SK05' = @{ Bl = '2026-09-04'; Rate = 8.712;  Fc = 10855290; Mj = '41072-26-702348X'; File = '신코.pdf'; Split = @{ '사형' = 76647650;  '금형 주조' = 556083;   '가공3(3공장)' = 17367553 }; Scrap = @(0, 0) }
  'DSM2609-SK01' = @{ Bl = '2026-09-07'; Rate = 8.6766; Fc = 12227374; Mj = '41072-26-702439X'; File = '신코.pdf'; Split = @{ '사형' = 76451055;  '금형 주조' = 6245740;  '가공3(3공장)' = 22950962 }; Scrap = @(95130, 349146) }
  'DSM2609-SK02' = @{ Bl = '2026-09-14'; Rate = 8.7625; Fc = 19165174; Mj = '41072-26-702583X'; File = '신코.pdf'; Split = @{ '사형' = 112534504; '금형 주조' = 20506534; '가공3(3공장)' = 32965488 }; Scrap = @(1928311, 0) }
  'DSM2609-SK03' = @{ Bl = '2026-09-20'; Rate = 8.8529; Fc = 9339380;  Mj = '41072-26-702665X'; File = '신코.pdf'; Split = @{ '사형' = 66373862;  '금형 주조' = 1264438;  '가공3(3공장)' = 14652504 }; Scrap = @(389788, 0) }
  'DSM2609-SK04' = @{ Bl = '2026-09-28'; Rate = 8.579;  Fc = 12873460; Mj = '41072-26-702741X'; File = '신코.pdf'; Split = @{ '사형' = 80695889;  '금형 주조' = 7864780;  '가공3(3공장)' = 21880744 }; Scrap = @(0, 0) }
  'DSM2609-JS01' = @{ Bl = '2026-09-15'; Rate = 8.7207; Fc = 7474920;  Mj = '41072-26-702465X'; File = 'js.pdf';   Split = @{ '사형' = 46473957;  '가공3(3공장)' = 18712577 }; Scrap = @(0, 0) }
  'DSM2609-JS02' = @{ Bl = '2026-09-23'; Rate = 8.6368; Fc = 7474920;  Mj = '41072-26-702629X'; File = 'js.pdf';   Split = @{ '사형' = 46026841;  '가공3(3공장)' = 18532548 }; Scrap = @(0, 0) }
  'DSM2609-JS03' = @{ Bl = '2026-09-30'; Rate = 8.6316; Fc = 5232444;  Mj = '41072-26-702700X'; File = 'js.pdf';   Split = @{ '사형' = 32199390;  '가공3(3공장)' = 12964973 }; Scrap = @(0, 0) }
}
$cats = @('사형', '자동', '금형 주조', '가공1(2공장)', '가공2(5공장)', '가공3(3공장)')
$closeYM = 202609   # 마감 월

# 16) 신코·코마츠 매출내역서 엑셀 (예: '9월신코, 코마츠 지출품의서마감.xlsx'): 탭이 색칠된 시트에서 '인보이스'·'엔화금액' 표를 '합계' 줄 전까지 읽음
#     인보이스 앞 'NN-' = 월 (마감 월보다 크면 이월, 작으면 전월분). 엔화 소수점은 반올림(회계 기준), 소수점이 있으면 '알림'
function Read-SkJs($xl, $path, $ym) {
  function N0($x) { $d = 0.0; if ([double]::TryParse(("$x" -replace '[,\s₩]', ''), [ref]$d)) { $d } else { 0.0 } }
  $res = [ordered]@{}
  $wbX = $xl.Workbooks.Open($path, 0, $true)
  try {
    foreach ($sh in @($wbX.Worksheets)) {
      if ($sh.Tab.ColorIndex -eq -4142 -or $sh.Visible -ne -1) { continue }
      $u = $sh.UsedRange; $nr = [math]::Min($u.Row + $u.Rows.Count - 1, 80); $ncx = [math]::Min($u.Column + $u.Columns.Count - 1, 60)
      $vv = $sh.Range($sh.Cells(1,1), $sh.Cells($nr,$ncx)).Value2
      $hr = $null; $col = @{}
      for ($i = 1; $i -le $nr -and -not $hr; $i++) { for ($j = 1; $j -le $ncx; $j++) { if ("$($vv[$i,$j])".Trim() -eq '인보이스') { $hr = $i; break } } }
      if (-not $hr) { continue }
      for ($j = 1; $j -le $ncx; $j++) { $h = ("$($vv[$hr,$j])" -replace '\s', ''); if ($h -and -not $col.ContainsKey($h)) { $col[$h] = $j } }
      if (-not $col['엔화금액']) { continue }
      for ($i = $hr + 1; $i -le $nr; $i++) {
        $inv = "$($vv[$i,$col['인보이스']])".Trim()
        if (-not $inv -or $inv -match '합\s*계') { break }
        $pm = $null; if ($inv -match '^(\d{1,2})\s*-\s*(.+)$') { $pm = [int]$matches[1]; $inv = $matches[2].Trim() }
        $blv = $vv[$i,$col['BL날짜']]
        $bl = if ($blv -is [double]) { [DateTime]::FromOADate($blv) } elseif ("$blv" -match '^(\d{1,2})/(\d{1,2})$') { New-Object DateTime ([int][math]::Floor($ym / 100)), ([int]$matches[1]), ([int]$matches[2]) } else { $null }
        $fcRaw = N0 $vv[$i,$col['엔화금액']]
        $res[$inv] = @{ Bl = $(if ($bl) { $bl.ToString('yyyy-MM-dd') } else { '' }); Rate = (N0 $vv[$i,$col['환율']]); Fc = [math]::Round($fcRaw); FcRaw = $fcRaw; PrefixM = $pm
          Split = @{ '사형' = (N0 $vv[$i,$col['사형']]); '금형 주조' = $(if ($col['금형']) { N0 $vv[$i,$col['금형']] } else { 0.0 }); '가공3(3공장)' = (N0 $vv[$i,$col['가공']]) }
          Scrap = @($(if ($col['스크랩(사형)']) { N0 $vv[$i,$col['스크랩(사형)']] } else { 0.0 }), $(if ($col['스크랩(금형)']) { N0 $vv[$i,$col['스크랩(금형)']] } else { 0.0 }))
          File = "$($sh.Name) 시트"; Mj = '' }
      }
    }
  } finally { $wbX.Close($false) }
  $res
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $skOut = @{}; $skAlert = @(); $skSrc = 'PDF(눈으로 읽은 값)'
  $skF = @(Get-ChildItem $base -File | ? { $_.Name -match '신코' -and $_.Extension -match '^\.xls' -and $_.Name -notmatch '^~\$' } | sort { if ($_.Name -match "^$($closeYM % 100)월") { 0 } else { 1 } }, Name) | select -First 1
  if ($skF) {
    $xr = Read-SkJs $xl $skF.FullName $closeYM
    if ($xr.Count) {
      foreach ($k in @($pdf.Keys)) { if (-not $xr.Contains($k)) { $skOut[$k] = "매출내역서 엑셀에 없음" } }   # 예: JS03 → 회계 회색이면 이월
      $pdf = [ordered]@{}
      foreach ($k in $xr.Keys) {
        $x = $xr[$k]
        if ($x.PrefixM -and $x.PrefixM -gt ($closeYM % 100)) { $skOut[$k] = "매출내역서 인보이스 앞 '$($x.PrefixM)-' → $($x.PrefixM)월 이월"; continue }
        $pdf[$k] = $x; $pdf[$k].File = "$($skF.Name) [$($x.File)]"
        if ($x.PrefixM -and $x.PrefixM -lt ($closeYM % 100)) { $pdf[$k].Prev = "$($x.PrefixM)월분" }
        if ([math]::Abs($x.FcRaw - $x.Fc) -ge 0.005) { $skAlert += [pscustomobject]@{ Inv = $k; Raw = $x.FcRaw; Fc = $x.Fc; Rate = $x.Rate } }
      }
      $skSrc = "엑셀 $($skF.Name)"
    }
  }
  foreach ($k in $pdf.Keys) { $p = $pdf[$k]; $p.BlN = [DateTime]::Parse($p.Bl).ToOADate(); $p.Krw = ($p.Split.Values | measure -Sum).Sum + $p.Scrap[0] + $p.Scrap[1] }
  "신코·코마츠 기준: $skSrc / 읽은 인보이스 $($pdf.Count)건" + $(if ($skOut.Count) { " / 빠진 건: " + (($skOut.Keys | % { "$_($($skOut[$_]))" }) -join ', ') } else { '' }) + $(if ($skAlert.Count) { " / 엔화 소수점 알림: " + (($skAlert | % Inv) -join ', ') } else { '' })
  # 9월라인: 인보이스별 bl 날짜·환율·달러·원화·6칸
  $line = @{}
  $wb = $xl.Workbooks.Open("$base\9월라인.xlsx", 0, $true)
  foreach ($s in $wb.Worksheets) {
    $u = $s.UsedRange; $v = $u.Value2; $nr = $u.Rows.Count; $nc = $u.Columns.Count; $hr = $null; $col = @{}
    for ($i = 1; $i -le [Math]::Min($nr, 25) -and -not $hr; $i++) { for ($j = 1; $j -le $nc; $j++) { if ("$($v[$i,$j])".Trim() -eq '인보이스') { $hr = $i } } }
    if (-not $hr) { continue }
    for ($j = 1; $j -le $nc; $j++) { $h = "$($v[$hr,$j])".Trim(); $h2 = $h -replace '\s', ''; if ($h -and -not $col.ContainsKey($h)) { $col[$h] = $j }; if ($h2 -and -not $col.ContainsKey($h2)) { $col[$h2] = $j } }
    $blC = if ($col['bl날짜']) { $col['bl날짜'] } else { $col['전자세금계산서발행일'] }
    for ($i = $hr + 1; $i -le $nr; $i++) {
      $inv = K $v[$i,$col['인보이스']]; if (-not $inv -or $line.ContainsKey($inv)) { continue }
      $sp = @{}; foreach ($c in $cats) { if ($col[$c]) { $sp[$c] = [double](Num $v[$i,$col[$c]]) } }
      $line[$inv] = [pscustomobject]@{ Bl = (Num $v[$i,$blC]); Rate = $(if ($col['환율']) { Num $v[$i,$col['환율']] }); Usd = $(if ($col['달러']) { Num $v[$i,$col['달러']] }); Krw = (Num $v[$i,$col['원화']]); Where = "$($s.Name)!R$($u.Row + $i - 1)"; Split = $sp }
    }
  }
  $wb.Close($false)

  # 회계 26.09(NEW): A 거래처, C 인보이스, D 면장번호, F BL일자, G 환율, H 외화, I 원화
  $acc = @{}; $accList = @()
  $wb = $xl.Workbooks.Open("$base\회계 2026년 수출 (2026.09).xlsx", 0, $true); $sa = $wb.Worksheets.Item("26.09(NEW)")
  $v = $sa.Range("A31:I132").Value2
  for ($i = 1; $i -le 102; $i++) {
    $inv = K $v[$i,3]; if (-not $inv) { continue }
    $a = [pscustomobject]@{ R = 30 + $i; Cust = (K $v[$i,1]); Inv = $inv; Mj = (K $v[$i,4]); Bl = (Num $v[$i,6]); Rate = (Num $v[$i,7]); Fc = (Num $v[$i,8]); Krw = (Num $v[$i,9]); Gray = $false; Green = $false }
    if ($a.Mj -match '^\d{5}-\d{2}-\d{6}X$') {
      $c = [int]$sa.Range("D$($a.R)").DisplayFormat.Interior.Color; $rr = $c -band 255; $gg = ($c -shr 8) -band 255; $bb = ($c -shr 16) -band 255
      $a.Gray = ($rr -eq $gg -and $gg -eq $bb -and $rr -lt 230); $a.Green = ($c -eq 10213316); $accList += $a
    }
    if (-not $acc.ContainsKey($inv)) { $acc[$inv] = $a }
  }
  $wb.Close($false)

  $wb = $xl.Workbooks.Open($mesPath)
  $ws = $wb.Worksheets.Item("mes9월")
  if ($wb.ReadOnly) { throw "mes9월.xls가 읽기 전용으로 열림(다른 Excel에서 열려 있음)" }
  $xl.ScreenUpdating = $false; $xl.Calculation = -4135   # 쓰는 동안 수동 계산
  # 10) 열 추가 (이미 있으면 건너뜀)
  $hd0a = @(1..($ws.UsedRange.Column + $ws.UsedRange.Columns.Count - 1) | % { "$($ws.Cells(1,$_).Value2)".Trim() })
  if ($hd0a -notcontains '기초자료환율') {
    if ("$($ws.Range('E1').Value2)".Trim() -ne '환율' -or "$($ws.Range('F1').Value2)".Trim() -ne '달러') { throw "E1/F1이 환율/달러가 아님: $($ws.Range('E1').Value2) / $($ws.Range('F1').Value2)" }
    $ws.Columns("F").Insert() | Out-Null; $ws.Range("F1").Value2 = '기초자료환율'
    $ws.Columns("H").Insert() | Out-Null; $ws.Range("H1").Value2 = '기초자료달러'
    "열 추가: F=기초자료환율, H=기초자료달러"
  }
  # 15) 금액 옆 '전자세금계산서 금액', 맨 끝 '세금계산서 일련번호' (이미 있으면 건너뜀)
  $nc0 = $ws.UsedRange.Column + $ws.UsedRange.Columns.Count - 1; $hd0 = @(1..$nc0 | % { "$($ws.Cells(1,$_).Value2)".Trim() })
  if ($hd0 -notcontains '전자세금계산서 금액') {
    $amtC0 = [array]::IndexOf($hd0, '금액') + 1; if ($amtC0 -lt 1) { throw "mes9월에 '금액' 열이 없음" }
    $ws.Columns($amtC0 + 1).Insert() | Out-Null; $ws.Cells(1, $amtC0 + 1).Value2 = '전자세금계산서 금액'; $nc0++; "열 추가: 전자세금계산서 금액 (금액 옆)"
  }
  if (@(1..$nc0 | % { "$($ws.Cells(1,$_).Value2)".Trim() }) -notcontains '세금계산서 일련번호') {
    $ws.Cells(1, $nc0 + 1).Value2 = '세금계산서 일련번호'; $ws.Cells(1, $nc0).Copy() | Out-Null; $ws.Cells(1, $nc0 + 1).PasteSpecial(-4122) | Out-Null; "열 추가: 세금계산서 일련번호 (맨 끝)"
  }
  # 사람이 직접 체크하는 칸 (☐/☑) — 전자세금계산서 금액 앞. 반영을 다시 해도 체크 내용은 지우지 않음 (사용자 요청 2026-10-31)
  $chkCols = @('불량', '이월', '반품', '관세', '툴링비')
  $hd1 = @(1..($ws.UsedRange.Column + $ws.UsedRange.Columns.Count - 1) | % { "$($ws.Cells(1,$_).Value2)".Trim() })
  if ($hd1 -notcontains $chkCols[0]) {
    $pos = [array]::IndexOf($hd1, '전자세금계산서 금액') + 1; if ($pos -lt 1) { throw "'전자세금계산서 금액' 열이 없음" }
    for ($k = $chkCols.Count - 1; $k -ge 0; $k--) { $ws.Columns($pos).Insert() | Out-Null; $ws.Cells(1, $pos).Value2 = [string]$chkCols[$k] }
    "열 추가: 체크 칸 $($chkCols -join ', ') (전자세금계산서 금액 앞)"
  }
  # 열 위치 정리 (사용자 요청 2026-10-31): 체크 칸(불량·이월·반품·관세·툴링비) → 인보이스 번호 뒤(수출면장 번호 앞),
  #   환율·기초자료환율·달러·기초자료달러 → 금액 뒤(전자세금계산서 금액 앞). 이미 그 순서면 건너뜀
  function ColOf($name) { $n0 = $ws.UsedRange.Column + $ws.UsedRange.Columns.Count - 1; for ($j = 1; $j -le $n0; $j++) { if ("$($ws.Cells(1,$j).Value2)".Trim() -eq $name) { return $j } }; 0 }
  function PutBefore($names, $anchor) {
    $a0 = ColOf $anchor; if (-not $a0) { throw "'$anchor' 열이 없음" }
    $ok = $true; for ($q = 0; $q -lt $names.Count; $q++) { if ((ColOf $names[$q]) -ne ($a0 - $names.Count + $q)) { $ok = $false } }
    if ($ok) { return $false }
    foreach ($nm in $names) { $s0 = ColOf $nm; if (-not $s0) { throw "'$nm' 열이 없음" }; $ws.Columns($s0).Cut() | Out-Null; $ws.Columns((ColOf $anchor)).Insert() | Out-Null }
    $true
  }
  # TOTAL 줄 병합(예: A~O "TOTAL") 때문에 열을 옮길 수 없음 → 잠시 풀고, 옮긴 뒤 "TOTAL"을 출고수량 앞까지 다시 합침
  $trw = $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1; $totMerged = $false
  if ("$($ws.Cells($trw,1).Value2)".Trim() -eq 'TOTAL') { $rg0 = $ws.Range($ws.Cells($trw,1), $ws.Cells($trw, $ws.UsedRange.Column + $ws.UsedRange.Columns.Count - 1)); if ($rg0.MergeCells -ne $false) { $rg0.UnMerge() | Out-Null; $totMerged = $true } }
  if (PutBefore @('전자세금계산서 금액', '환율', '기초자료환율', '달러', '기초자료달러') '사형') { "열 이동: 금액 뒤 = 전자세금계산서 금액 → 환율·기초자료환율·달러·기초자료달러" }   # 사용자 요청 2026-10-31
  $cMemo = ColOf '메모'; $cEnd = $ws.UsedRange.Column + $ws.UsedRange.Columns.Count - 1   # 메모는 맨 끝 열 (사용자 요청 2026-10-31)
  if ($cMemo -and $cMemo -ne $cEnd) { $ws.Columns($cMemo).Cut() | Out-Null; $ws.Columns($cEnd + 1).Insert() | Out-Null; "열 이동: 메모 → 맨 끝" }
  if (PutBefore $chkCols '수출면장 번호') { "열 이동: 체크 칸 → 인보이스 번호 뒤" }
  if (PutBefore @('담당자', '출고Lot') '구분') { "열 이동: 담당자·출고Lot → 구분 앞" }   # 사용자 요청 2026-10-31
  if ($totMerged) { $q0 = (ColOf '출고수량') - 1; if ($q0 -ge 2) { $m0 = $ws.Range($ws.Cells($trw,1), $ws.Cells($trw,$q0)); $m0.Merge() | Out-Null; $m0.HorizontalAlignment = -4108 } }
  # 줄 정렬 (사용자 요청 2026-10-31): 국내 위 · 해외 아래, 각각 품번 순 → 출고일 순. NO는 MES 원래 번호 그대로(체크 기록 기준)
  #   줄 위치가 바뀌면 오류원본의 'MES 행' 번호도 새 위치로 바꿈. 국내/해외 사이 굵은 선
  $lastS = $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1; if ("$($ws.Cells($lastS,1).Value2)".Trim() -eq 'TOTAL') { $lastS-- }
  $cGuS = ColOf '구분'; $cPnS = ColOf '품번'; $cDtS = ColOf '출고일'; $cNoS = ColOf 'NO'; $ncS = $ws.UsedRange.Column + $ws.UsedRange.Columns.Count - 1
  if ($cGuS -and $cPnS -and $cDtS -and $cNoS) {
    $noOld = @{}; for ($r0 = 2; $r0 -le $lastS; $r0++) { $noOld["$($ws.Cells($r0,$cNoS).Value2)".Trim()] = $r0 }
    $dataS = $ws.Range($ws.Cells(2,1), $ws.Cells($lastS,$ncS))
    for ($r0 = 2; $r0 -le $lastS; $r0++) { $bd0 = $ws.Range($ws.Cells($r0,1), $ws.Cells($r0,$ncS)).Borders.Item(8); if ($bd0.Weight -eq -4138) { $bd0.Weight = 2 } }   # 예전 굵은 구분선만 가늘게
    $dataS.Sort($ws.Cells(2,$cGuS), 1, $ws.Cells(2,$cPnS), [Type]::Missing, 1, $ws.Cells(2,$cDtS), 1, 2) | Out-Null
    $mapS = @{}; $moved = 0; for ($r0 = 2; $r0 -le $lastS; $r0++) { $n0 = "$($ws.Cells($r0,$cNoS).Value2)".Trim(); $old0 = $noOld[$n0]; if ($old0) { $mapS["$old0"] = $r0; if ($old0 -ne $r0) { $moved++ } } }
    $firstOv = 0; for ($r0 = 2; $r0 -le $lastS; $r0++) { if ("$($ws.Cells($r0,$cGuS).Value2)".Trim() -eq '해외') { $firstOv = $r0; break } }
    if ($firstOv -gt 2) { $bd = $ws.Range($ws.Cells($firstOv,1), $ws.Cells($firstOv,$ncS)).Borders.Item(8); $bd.LineStyle = 1; $bd.Weight = -4138; $bd.Color = 0 }
    if ($moved) {
      foreach ($sh0 in @($wb.Worksheets)) { if ($sh0.Name -eq '오류원본') {
        $rn0 = $sh0.UsedRange.Rows.Count
        for ($r0 = 2; $r0 -le $rn0; $r0++) { $s0 = "$($sh0.Cells($r0,2).Value2)"; if (-not $s0 -or $s0 -match '^(해외|국내)매출') { continue }
          $s1 = [regex]::Replace($s0, '\d+', { param($m) if ($mapS.ContainsKey($m.Value)) { "$($mapS[$m.Value])" } else { $m.Value } }); if ($s1 -ne $s0) { $sh0.Cells($r0,2).NumberFormat = "@"; $sh0.Cells($r0,2).Value2 = [string]$s1 } }
      } }
      "줄 정렬: 국내 위 · 해외 아래(품번 → 출고일), $($moved)줄 위치 바뀜 (해외 시작 $($firstOv)행)"
    }
  }
  $nc = $ws.UsedRange.Columns.Count; $hdr = $ws.Range($ws.Cells(1,1), $ws.Cells(1,$nc)).Value2
  $C = @{}; for ($j = 1; $j -le $nc; $j++) { $h = "$($hdr[1,$j])".Trim(); if ($h -and -not $C.ContainsKey($h)) { $C[$h] = $j } }
  foreach ($h in '출고일','BL날짜','인보이스 번호','환율','기초자료환율','달러','기초자료달러','수출면장 번호','구분','고객사','금액','출고Lot','출고수량') { if (-not $C[$h]) { throw "mes9월에 '$h' 열이 없음" } }
  $last = $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1
  if ("$($ws.Cells($last,1).Value2)".Trim() -eq 'TOTAL') { $tot = $last; $last-- } else { $tot = $null }
  $cTaxA = $C['전자세금계산서 금액']; $cTaxN = $C['세금계산서 일련번호']   # 매번 새로 채움
  # 체크 칸: 숫자 1 = ☑, 0 = ☐ (서식으로 표시). 칸 선택 → 1 입력 → Enter, 여러 칸은 선택 → 1 → Ctrl+Enter. 예전 ☑/☐ 글자는 1/0으로 바꿈. 체크 내용은 지우지 않음
  # HTML '체크 저장'으로 만든 체크_YYYYMM*.json 이 월 폴더에 있으면 여기서 합친 뒤 이름을 '체크_반영됨_…'으로 바꿈
  $chkJson = @(Get-ChildItem $base -File -Filter '체크_*.json' -ErrorAction SilentlyContinue | ? { $_.Name -notmatch '반영됨' } | sort LastWriteTime)
  $noCol = $C['NO']; $noRow = @{}; if ($noCol) { for ($r0 = 2; $r0 -le $last; $r0++) { $k0 = "$($ws.Cells($r0,$noCol).Value2)".Trim(); if ($k0) { $noRow[$k0] = $r0 } } }
  foreach ($h0 in $chkCols) {
    $cc = $C[$h0]; if (-not $cc) { continue }
    $rg = $ws.Range($ws.Cells(2,$cc), $ws.Cells($last,$cc))
    try { $rg.Validation.Delete() } catch {}
    $rg.Validation.Add(1, 1, 1, "0", "1") | Out-Null; $rg.Validation.InputTitle = "$h0 체크"; $rg.Validation.InputMessage = "1 = 체크(☑), 0 = 해제(☐)"; $rg.Validation.ErrorMessage = "1(체크) 또는 0(해제)만 입력해 주세요"
    $rg.NumberFormat = '[=1]"☑";"☐"'; $rg.HorizontalAlignment = -4108; $ws.Columns($cc).ColumnWidth = 6; $ws.Cells(1,$cc).HorizontalAlignment = -4108
    $cv = $rg.Value2
    for ($r0 = 1; $r0 -le ($last - 1); $r0++) { $x0 = "$($cv[$r0,1])".Trim(); $n0 = if ($x0 -eq '☑' -or $x0 -eq '1') { 1 } else { 0 }; if ("$($cv[$r0,1])" -ne "$n0") { $ws.Cells($r0 + 1, $cc).Value2 = [double]$n0 } }
  }
  $nJ = 0
  foreach ($jf in $chkJson) {
    $js = Get-Content $jf.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($pp in $js.PSObject.Properties) { $kk = $pp.Name -split '\|'; if ($kk.Count -ne 2) { continue }; $r0 = $noRow[$kk[0]]; $cc = $C[$kk[1]]; if ($r0 -and $cc -and ($chkCols -contains $kk[1])) { $ws.Cells($r0, $cc).Value2 = [double]$(if ([int]$pp.Value) { 1 } else { 0 }); $nJ++ } }
    Rename-Item $jf.FullName ("체크_반영됨_" + (Get-Date -Format 'yyyyMMdd_HHmm') + "_" + $jf.Name.Substring(3)) -ErrorAction SilentlyContinue
  }
  if ($chkJson.Count) { "HTML 체크 파일 $($chkJson.Count)개 → 엑셀에 $($nJ)칸 반영" }

  foreach ($cc in $cTaxA, $cTaxN) { $ws.Range($ws.Cells(2,$cc), $ws.Cells($last,$cc)).ClearContents() | Out-Null }
  $ws.Range($ws.Cells(2,$cTaxA), $ws.Cells($last + 1,$cTaxA)).NumberFormat = "#,##0"
  $ws.Columns($cTaxN).ColumnWidth = 27; $ws.Columns($cTaxN).WrapText = $true   # 일련번호(24자리)는 여러 개면 한 줄에 하나씩
  if ($tot) { $clT = ($ws.Cells(1,$cTaxA).Address($false,$false) -replace '\d',''); $ws.Cells($tot, $cTaxA).Formula = "=SUM(${clT}2:${clT}$last)" }
  $v = $ws.Range($ws.Cells(1,1), $ws.Cells($last,$nc)).Value2
  # 메모에 Replacement = 무상 교체품 → 금액·달러·기초자료달러 0이 정상, 에러 아님 (작업 13)
  $repl = @{}; if ($C['메모']) { for ($r = 2; $r -le $last; $r++) { if ("$($v[$r,$C['메모']])" -match 'replacement') { $repl[$r] = $true } } }
  # 국내 MES 금액이 빈 행(Replacement 아님) = 불량대체 → 정상 (국내 영업팀 확인 2026-10-30). 전자세금계산서 금액 칸에 '불량대체'
  $bulD = @{}; for ($r = 2; $r -le $last; $r++) { if ("$($v[$r,$C['구분']])".Trim() -eq '국내' -and -not $repl[$r] -and -not (Num $v[$r,$C['금액']])) { $bulD[$r] = $true } }

  # 출고Lot 오기: 어느 자료에도 없고 비슷한 번호가 하나뿐이면 그 번호로 인식(예: DSM2608-AL6 → AL06, 8월 이월 건). MES 칸은 그대로 두고 빨간색
  $allKeys = @(@($line.Keys) + @($acc.Keys) + @($pdf.Keys) | sort -Unique)
  $lots = @{}; $alias = @{}
  for ($r = 2; $r -le $last; $r++) {
    $o = K $v[$r,$C['출고Lot']]; $lots[$r] = $o
    if (-not $o -or $o -notmatch '^(DSM|\d{4}P)' -or $line[$o] -or $acc[$o] -or $pdf[$o]) { continue }
    $cand = @($allKeys | ? { (NZ0 $_) -eq (NZ0 $o) })
    if ($cand.Count -eq 1) { $alias[$r] = $o; $lots[$r] = [string]$cand[0]   # 사용자 확인 완료(12) → 에러 표시 안 함, 행 배경색 그대로
      $bg0 = $ws.Cells($r, $C['출고일']).Interior; if ($bg0.ColorIndex -eq -4142) { $ws.Cells($r, $C['출고Lot']).Interior.ColorIndex = -4142 } else { $ws.Cells($r, $C['출고Lot']).Interior.Color = [int]$bg0.Color } }
  }

  $cBl = $C['BL날짜']; $cInv = $C['인보이스 번호']; $cFx = $C['기초자료환율']; $cUsd = $C['기초자료달러']
  foreach ($cc in $cBl, $cInv, $cFx, $cUsd) { $rg = $ws.Range($ws.Cells(2,$cc), $ws.Cells($last,$cc)); $rg.ClearContents() | Out-Null; $rg.ClearComments() | Out-Null }
  $blRows = @(); $invRows = @(); $cnt = @{ PDF = 0; LINE = 0; DOM = 0; Red = 0; InvRed = 0; Carry = 0 }
  $grp = [ordered]@{}   # 인보이스별 MES 행 (기초자료달러 배분용)
  for ($r = 2; $r -le $last; $r++) {
    $lot = $lots[$r]; $kind = "$($v[$r,$C['구분']])".Trim(); $cust = "$($v[$r,$C['고객사']])".Trim()
    $bgI = $ws.Cells($r, $C['출고일']).Interior   # 원래 행 배경(SK05 추가행 주황 등) 유지
    $bgNone = ($bgI.ColorIndex -eq -4142); $bgCol = [int]$bgI.Color
    foreach ($cc in $cBl, $cInv, $cFx, $cUsd) { if ($bgNone) { $ws.Cells($r,$cc).Interior.ColorIndex = -4142 } else { $ws.Cells($r,$cc).Interior.Color = $bgCol } }
    $src = $null
    if ($skOut[$lot]) {   # 신코·코마츠인데 매출내역서 엑셀에 없음(또는 'NN-' 다음 달) → 회계 회색이면 이월
      $a0 = $acc[$lot]; $gray0 = @($accList | ? { $_.Inv -eq $lot -and $_.Gray }).Count -gt 0
      $ws.Cells($r,$cBl).NumberFormat = "@"; $ws.Cells($r,$cBl).Value2 = '이월'; $ws.Cells($r,$cBl).Interior.Color = 15921906; $cnt.Carry++
      $blRows += ,@($r, $cust, $lot, $(if ($a0) { $a0.Bl }), '신코·코마츠 매출내역서', $null, "$($skOut[$lot]) / $(if ($gray0 -or $skOut[$lot] -match '이월') { "회계 면장번호 회색 → $(($closeYM % 100) % 12 + 1)월 이월 건" } else { '회계 면장번호가 회색이 아님 → 매출내역서에서 빠졌는지 확인해 주세요' })")
      continue
    }
    if ($pdf[$lot]) { $p = $pdf[$lot]; $src = [pscustomobject]@{ Name = "$($p.File) 매출내역서"; Bl = $p.BlN; Inv = $lot; Rate = $p.Rate; Fc = $p.Fc } ; $cnt.PDF++ }
    elseif ($kind -eq '해외') {
      $cnt.LINE++
      if ($line[$lot]) { $l = $line[$lot]; $src = [pscustomobject]@{ Name = "9월라인 $($l.Where)"; Bl = $l.Bl; Inv = $lot; Rate = $l.Rate; Fc = $l.Usd } }
    }
    elseif ($kind -eq '국내') {
      # 국내: BL날짜 = 출하일(MES 출고일), 인보이스 번호는 기재하지 않음, 오류 표시 없음
      $d = Num $v[$r,$C['출고일']]; if ($d) { $ws.Cells($r,$cBl).Value2 = $d; $ws.Cells($r,$cBl).NumberFormat = "yyyy-mm-dd" }
      $cnt.DOM++; continue
    } else { continue }

    $a = $acc[$lot]
    $al = if ($alias[$r]) { " / MES 출고Lot 오기: '$($alias[$r])' → '$lot'로 인식" } else { '' }
    # BL날짜
    if ($src -and $src.Bl) { $ws.Cells($r,$cBl).Value2 = [double]$src.Bl; $ws.Cells($r,$cBl).NumberFormat = "yyyy-mm-dd" }
    $why = @()
    if (-not $src) { $why += "9월라인에 인보이스 없음" } elseif (-not $src.Bl) { $why += "$($src.Name)에 BL날짜 없음" }
    if (-not $a) { $why += "회계에 인보이스 없음(기준 없음)" }
    elseif ($src -and $src.Bl -and $a.Bl -and [math]::Floor($src.Bl) -ne [math]::Floor($a.Bl)) { $why += "BL날짜 다름: 기초자료 $(DS $src.Bl) / 회계 $(DS $a.Bl)" }
    if ($a -and $a.Bl -ge 46296 -and -not $src) { $why += "회계 BL일 $(DS $a.Bl) → 10월 이월 건(9월라인은 당월 BL분만 있음)" }
    $carry = $a -and $a.Bl -ge 46296 -and -not $src   # 회계 BL일 10/01 = 다음 달 이월 → 날짜 대신 '이월', 회색
    if ($carry) { $ws.Cells($r,$cBl).NumberFormat = "@"; $ws.Cells($r,$cBl).Value2 = '이월'; $ws.Cells($r,$cBl).Interior.Color = 15921906; $cnt.Carry++
      $blRows += ,@($r, $cust, $lot, $a.Bl, '9월라인', $null, (($why -join ' / ') + $al)) }
    elseif ($why) { $ws.Cells($r,$cBl).Interior.Color = $RED; $cnt.Red++
      $blRows += ,@($r, $cust, $lot, $(if ($a) { $a.Bl }), $(if ($src) { $src.Name } else { '9월라인' }), $(if ($src) { $src.Bl }), (($why -join ' / ') + $al)) }
    # 인보이스 번호
    if ($src) { $ws.Cells($r,$cInv).NumberFormat = "@"; $ws.Cells($r,$cInv).Value2 = [string]$src.Inv }
    $why2 = @(); if (-not $src) { $why2 += "9월라인에 인보이스 없음" }; if (-not $a) { $why2 += "회계에 인보이스 없음(기준 없음)" }
    if ($a -and $a.Bl -ge 46296 -and -not $src) { $why2 += "10월 이월 건" }
    if ($why2) { $ws.Cells($r,$cInv).Interior.Color = $RED; $cnt.InvRed++
      $invRows += ,@($r, $cust, $lot, $(if ($a) { $a.Inv }), $(if ($src) { $src.Name } else { '9월라인' }), $(if ($src) { $src.Inv } else { '없음' }), (($why2 -join ' / ') + $al)) }
    if (-not $grp.Contains($lot)) { $grp[$lot] = [pscustomobject]@{ Src = $src; Acc = $a; Cust = $cust; Rows = New-Object System.Collections.ArrayList } }
    [void]$grp[$lot].Rows.Add($r)
  }

  # 기초자료환율·기초자료달러 (인보이스 단위 값 → 달러는 MES 금액 비율로 행에 배분, 합계 = 기초자료 금액)
  $cmpRows = @()
  foreach ($lot in $grp.Keys) {
    $g = $grp[$lot]; $src = $g.Src; $a = $g.Acc; $rows = @($g.Rows)
    $jpy = [bool]$pdf[$lot]; $fmtR = if ($jpy) { "0.0000" } else { "#,##0.00" }
    $krwOnly = $src -and -not $src.Rate -and -not $src.Fc   # Bergen·Atlas: 9월라인 원화 거래
    $why = @()
    if (-not $src) { $why += "9월라인에 인보이스 없음 → 기초자료 환율·달러 없음"; if ($a -and $a.Bl -ge 46296) { $why += "10월 이월 건" } }
    elseif ($krwOnly) { }
    else {
      foreach ($r in $rows) { $ws.Cells($r,$cFx).Value2 = [double]$src.Rate; $ws.Cells($r,$cFx).NumberFormat = $fmtR }
      $amts = @($rows | % { [double](Num $ws.Cells($_, $C['금액']).Value2) }); $sum = ($amts | measure -Sum).Sum
      $acc2 = 0.0
      $endI = $rows.Count - 1; while ($endI -gt 0 -and -not $amts[$endI]) { $endI-- }   # 단수 차이는 금액 있는 마지막 행이 흡수(Replacement 0원 행은 0)
      for ($i = 0; $i -lt $rows.Count; $i++) {
        $val = if ($i -gt $endI -or ($sum -and -not $amts[$i])) { 0.0 } elseif ($i -eq $endI) { [math]::Round($src.Fc - $acc2, 2) } elseif ($sum) { [math]::Round($src.Fc * $amts[$i] / $sum, 2) } else { 0.0 }
        $acc2 += $val; $ws.Cells($rows[$i],$cUsd).Value2 = $val; $ws.Cells($rows[$i],$cUsd).NumberFormat = "#,##0.00"
      }
      if (-not $a) { $why += "회계에 인보이스 없음(기준 없음)" }
      else {
        if (-not $a.Rate -or [math]::Abs($a.Rate - $src.Rate) -ge 0.00005) { $why += "환율 다름: 기초자료 $($src.Rate) / 회계 $($a.Rate)"; foreach ($r in $rows) { $ws.Cells($r,$cFx).Interior.Color = $RED } }
        if (-not $a.Fc -or [math]::Abs($a.Fc - $src.Fc) -gt 0.01) { $why += "외화금액 다름: 기초자료 {0:N2} / 회계 {1:N2}" -f $src.Fc, $a.Fc; foreach ($r in $rows) { $ws.Cells($r,$cUsd).Interior.Color = $RED } }
      }
    }
    if (-not $src) { foreach ($r in $rows) { $ws.Cells($r,$cFx).Interior.Color = $RED; $ws.Cells($r,$cUsd).Interior.Color = $RED } }
    if ($why) { $cmpRows += ,@($lot, $g.Cust, $(if ($src) { $src.Name } else { '9월라인' }), $(if ($src) { $src.Rate }), $(if ($a) { $a.Rate }), $(if ($src) { $src.Fc }), $(if ($a) { $a.Fc }), (($rows | % { "R$_" }) -join ','), ($why -join ' / ')) }
  }
  # Replacement 행: 달러·기초자료달러 0, 달러·기초자료달러·6칸 빨간색 해제(행 배경색으로)
  foreach ($r in @($repl.Keys)) {
    $ws.Cells($r, $C['달러']).Value2 = 0.0; if ($ws.Cells($r, $cUsd).Value2 -ne $null) { $ws.Cells($r, $cUsd).Value2 = 0.0 }
    $bg0 = $ws.Cells($r, $C['출고일']).Interior
    foreach ($cc in @($C['달러'], $cUsd) + @($cats | % { $C[$_] } | ? { $_ })) { if ($bg0.ColorIndex -eq -4142) { $ws.Cells($r, $cc).Interior.ColorIndex = -4142 } else { $ws.Cells($r, $cc).Interior.Color = [int]$bg0.Color } }
  }
  # 불량대체 행: 6칸 빨간색 해제(행 배경색으로)
  foreach ($r in @($bulD.Keys)) {
    $bg0 = $ws.Cells($r, $C['출고일']).Interior
    foreach ($cc in @($cats | % { $C[$_] } | ? { $_ })) { if ($bg0.ColorIndex -eq -4142) { $ws.Cells($r, $cc).Interior.ColorIndex = -4142 } else { $ws.Cells($r, $cc).Interior.Color = [int]$bg0.Color } }
  }
  # 숫자 쉼표 서식 (품번·NO 등 번호 열은 건드리지 않음). 환율은 엔화(50 미만) 소수 4자리, 달러 2자리
  $endR = $last + 1
  foreach ($h in '출고수량', '주조비', '가공비', '단가', '금액', '사형', '자동', '금형 주조', '가공1(2공장)', '가공2(5공장)', '가공3(3공장)') { if ($C[$h]) { $ws.Range($ws.Cells(2,$C[$h]), $ws.Cells($endR,$C[$h])).NumberFormat = "#,##0" } }
  foreach ($h in '환율', '기초자료환율') { $ws.Range($ws.Cells(2,$C[$h]), $ws.Cells($endR,$C[$h])).NumberFormat = "[<50]#,##0.0000;#,##0.00" }
  foreach ($h in '달러', '기초자료달러') { $ws.Range($ws.Cells(2,$C[$h]), $ws.Cells($endR,$C[$h])).NumberFormat = "#,##0.00" }
  foreach ($cc in $cBl, $cInv, $cFx, $cUsd, $C['달러'], $C['환율']) { $ws.Columns($cc).AutoFit() | Out-Null }

  function NewSheet($name, $titles) {
    foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq $name) { $s.Delete() } }
    $n = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)); $n.Name = $name
    for ($j = 0; $j -lt $titles.Count; $j++) { $n.Cells(1, $j + 1).Value2 = [string]$titles[$j] }
    $h = $n.Range($n.Cells(1,1), $n.Cells(1,$titles.Count)); $h.Font.Bold = $true; $h.Interior.Color = $HEAD
    $n
  }
  function Put($sh, $rowsArr) {
    for ($i = 0; $i -lt $rowsArr.Count; $i++) { for ($j = 0; $j -lt $rowsArr[$i].Count; $j++) {
      $x = $rowsArr[$i][$j]; if ($x -eq $null -or "$x" -eq '') { continue }
      if ($x -is [double] -or $x -is [int]) { $sh.Cells($i + 2, $j + 1).Value2 = [double]$x } else { $sh.Cells($i + 2, $j + 1).NumberFormat = "@"; $sh.Cells($i + 2, $j + 1).Value2 = [string]$x } } }
  }
  # 사유 시트 (실행할 때마다 새로 작성)
  $s1 = NewSheet "BL불일치사유" @('MES 행', '고객사', '인보이스 번호(출고Lot)', '회계 BL날짜(기준)', '기초자료 출처', '기초자료 BL날짜', '사유')
  Put $s1 $blRows
  for ($i = 0; $i -lt $blRows.Count; $i++) { $b = $blRows[$i][3]; $o = $blRows[$i][5]; if (-not $b) { $s1.Cells($i + 2, 4).Interior.Color = $PINK }; if (-not $o -or ($b -and [math]::Floor($o) -ne [math]::Floor($b))) { $s1.Cells($i + 2, 6).Interior.Color = $PINK } }
  $s1.Range("D:D").NumberFormat = "yyyy-mm-dd"; $s1.Range("F:F").NumberFormat = "yyyy-mm-dd"; $s1.Columns("A:G").AutoFit() | Out-Null
  $s2 = NewSheet "인보이스불일치" @('MES 행', '고객사', '출고Lot', '회계 인보이스 번호(기준)', '기초자료 출처', '기초자료 인보이스', '사유')
  Put $s2 $invRows
  for ($i = 0; $i -lt $invRows.Count; $i++) { if (-not $invRows[$i][3]) { $s2.Cells($i + 2, 4).Interior.Color = $PINK }; if ($invRows[$i][5] -eq '없음') { $s2.Cells($i + 2, 6).Interior.Color = $PINK } }
  $s2.Columns("A:G").AutoFit() | Out-Null
  $s3 = NewSheet "기초자료불일치" @('인보이스', '고객사', '기초자료 출처', '기초자료 환율', '회계 환율', '기초자료 달러(엔)', '회계 외화금액', 'MES 행', '사유')
  Put $s3 $cmpRows
  $s3.Range("D:E").NumberFormat = "#,##0.0000"; $s3.Range("F:G").NumberFormat = "#,##0.00"; $s3.Columns("A:I").AutoFit() | Out-Null

  # 9) 면장별합계 하나로: 회계 면장 목록(회색=10월 이월 제외). 신코·코마츠 = PDF 매출내역서, 그 외 = MES(SUMIFS) / MES에 없으면 9월라인 참고
  foreach ($s in @($wb.Worksheets)) { if ($s.Name -in '면장별합계', '면장별합계2', '해외매출') { $s.Delete() } }
  $gs = $wb.Worksheets.Add([Type]::Missing, $ws); $gs.Name = "해외매출"   # 12) 면장별합계 → 해외매출
  $ColL = { param($c) ($ws.Cells(1,$c).Address($false,$false) -replace '\d','') }
  $mesName = "'mes9월'"; $mjL = & $ColL $C['수출면장 번호']
  $gh = @('수출면장 번호', '인보이스 번호', 'BL날짜', '고객사', '구분', '통화', '환율', '외화금액', '원화금액(회계)', 'MES 행 수', 'MES 출고수량', 'MES 금액') + $cats + @('스크랩(사형)', '스크랩(금형)', '출처 기준처', '에러내용')
  for ($j = 0; $j -lt $gh.Count; $j++) { $gs.Cells(1, $j + 1).Value2 = [string]$gh[$j] }
  $mesMj = @{}; for ($r = 2; $r -le $last; $r++) { $m0 = "$($ws.Cells($r,$C['수출면장 번호']).Value2)".Trim(); if ($m0) { $mesMj[$m0]++ } }
  $row = 2; $nPdf = 0; $nMes = 0; $nRef = 0; $nRed = 0; $gsNotes = @()
  foreach ($t in @($accList | ? { -not $_.Gray })) {
    $note = @()
    $gs.Cells($row, 1).NumberFormat = "@"; $gs.Cells($row, 1).Value2 = [string]$t.Mj
    $gs.Cells($row, 2).Value2 = [string]$t.Inv
    if ($t.Bl) { $gs.Cells($row, 3).Value2 = [double]$t.Bl; $gs.Cells($row, 3).NumberFormat = "yyyy-mm-dd" }
    $gs.Cells($row, 4).Value2 = [string]$t.Cust
    $gs.Cells($row, 5).Value2 = [string]$(if ($t.Green) { '전월 출하·당월 선적' } else { '당월 출하·당월 선적' })
    $gs.Cells($row, 6).Value2 = [string]$(if (-not $t.Rate) { 'KRW' } elseif ($t.Rate -lt 50) { 'JPY' } elseif ($t.Cust -match 'BERGEN') { 'KRW(1,350 고정)' } else { 'USD' })
    if ($t.Rate) { $gs.Cells($row, 7).Value2 = [double]$t.Rate }; if ($t.Fc) { $gs.Cells($row, 8).Value2 = [double]$t.Fc }; if ($t.Krw) { $gs.Cells($row, 9).Value2 = [double]$t.Krw }
    $gs.Cells($row, 10).Formula = "=COUNTIFS($mesName!${mjL}:${mjL},`$A$row)"
    $q = 11; foreach ($sc0 in @($C['출고수량'], $C['금액'])) { $sc = & $ColL $sc0; $gs.Cells($row, $q).Formula = "=SUMIFS($mesName!${sc}:${sc},$mesName!${mjL}:${mjL},`$A$row)"; $q++ }
    # 회계 기준 비교 대상 (PDF 또는 9월라인)
    $p = $pdf[$t.Inv]; $l = $line[$t.Inv]
    if ($p) {
      $cmp = [pscustomobject]@{ N = $p.File; Bl = $p.BlN; Rate = $p.Rate; Fc = $p.Fc; Krw = $p.Krw; Mj = $p.Mj }
      $k0 = 13; foreach ($k in $cats) { $vv = [double]$p.Split[$k]; if ($vv) { $gs.Cells($row, $k0).Value2 = $vv }; $k0++ }
      if ($p.Scrap[0]) { $gs.Cells($row, 19).Value2 = [double]$p.Scrap[0] }; if ($p.Scrap[1]) { $gs.Cells($row, 20).Value2 = [double]$p.Scrap[1] }
      $gs.Range($gs.Cells($row,13), $gs.Cells($row,20)).Interior.Color = 15921906
      $gs.Cells($row, 21).Value2 = [string]"$($p.File) 매출내역서(스크랩 별도 칸, 가공은 가공3)"; $nPdf++
      if (-not $mesMj[$t.Mj]) { $note += "MES 9월 출고에 없음(전월 출고분)" }
    } else {
      $cmp = if ($l) { [pscustomobject]@{ N = '9월라인'; Bl = $l.Bl; Rate = $l.Rate; Fc = $l.Usd; Krw = $l.Krw; Mj = $null } } else { $null }
      if ($mesMj[$t.Mj]) {
        $k0 = 13; foreach ($k in $cats) { $sc = & $ColL $C[$k]; $gs.Cells($row, $k0).Formula = "=SUMIFS($mesName!${sc}:${sc},$mesName!${mjL}:${mjL},`$A$row)"; $k0++ }
        $gs.Cells($row, 21).Value2 = "MES"; $nMes++
      } elseif ($l) {
        $k0 = 13; foreach ($k in $cats) { if ($l.Split[$k]) { $gs.Cells($row, $k0).Value2 = [double]$l.Split[$k] }; $k0++ }
        $gs.Range($gs.Cells($row,13), $gs.Cells($row,18)).Interior.Color = 10092543; $gs.Cells($row, 21).Value2 = "9월라인(참고)"; $nRef++
        $note += "MES 9월 출고에 없음(전월 출고분)"
      } else { $gs.Range($gs.Cells($row,13), $gs.Cells($row,18)).Interior.Color = $RED; $note += "6칸 분해 자료 없음"; $note += "MES 9월 출고에 없음(전월 출고분)" }
    }
    # 회계와 비교 → 다르면 해당 칸 빨간색
    if (-not $cmp) { $note += "9월라인에 인보이스 없음(비교 불가)" }
    else {
      if ($cmp.Bl -and $t.Bl -and [math]::Floor($cmp.Bl) -ne [math]::Floor($t.Bl)) { $gs.Cells($row, 3).Interior.Color = $RED; $note += "BL날짜: $($cmp.N) $(DS $cmp.Bl) ≠ 회계 $(DS $t.Bl)" }
      if ($cmp.Mj -and $cmp.Mj -ne $t.Mj) { $gs.Cells($row, 1).Interior.Color = $RED; $note += "면장번호: $($cmp.N) $($cmp.Mj) ≠ 회계 $($t.Mj)" }
      if ($cmp.Rate -and (-not $t.Rate -or [math]::Abs($cmp.Rate - $t.Rate) -ge 0.00005)) { $gs.Cells($row, 7).Interior.Color = $RED; $note += "환율: $($cmp.N) $($cmp.Rate) ≠ 회계 $($t.Rate)" }
      if ($cmp.Fc -and (-not $t.Fc -or [math]::Abs($cmp.Fc - $t.Fc) -gt 0.01)) { $gs.Cells($row, 8).Interior.Color = $RED; $note += "외화: {0} {1:N2} ≠ 회계 {2:N2}" -f $cmp.N, $cmp.Fc, $t.Fc }
      if ($cmp.Krw -and (-not $t.Krw -or [math]::Abs($cmp.Krw - $t.Krw) -gt 0.5)) { $gs.Cells($row, 9).Interior.Color = $RED; $note += "원화 금액 다름: {0} 합계 {1:N0}원 / 회계 {2:N0}원 → {3:N0}원 {4}{5}" -f $cmp.N, $cmp.Krw, $t.Krw, [math]::Abs($cmp.Krw - $t.Krw), $(if ($cmp.Krw -lt $t.Krw) { '적음' } else { '많음' }), $(if ($p) { $al0 = @($skAlert | ? { $_.Inv -eq $t.Inv })[0]; if ($al0) { ". 매출내역서 엑셀 엔화가 {0:N2}엔(소수점)이라 원화가 달라짐. 회계 {1:N0}엔이 맞음 → 이번 달에도 국내팀 엑셀 엔화 소수점을 확인해 주세요 (알림)" -f $al0.Raw, $t.Fc } else { ". 회계(엔화×환율)가 맞음 → 매출내역서 항목별 금액(사형·금형·가공·스크랩) 확인 필요" } } else { '' }) }
    }
    if ($note | ? { $_ -match '≠|없음\(비교|금액 다름' }) { $nRed++; $gsNotes += "  $($t.Inv)`t$($note -join ' / ')" }
    if ($note) { $gs.Cells($row, 22).Value2 = [string]($note -join ' / ') }
    $row++
  }
  $gs.Cells($row, 1).Value2 = "합계"; $gs.Range($gs.Cells($row,1), $gs.Cells($row,22)).Font.Bold = $true
  foreach ($cc in 9..20) { $cl = ($gs.Cells(1,$cc).Address($false,$false) -replace '\d',''); $gs.Cells($row, $cc).Formula = "=SUM(${cl}2:${cl}$($row-1))" }
  $gs.Range("G2:G$row").NumberFormat = "#,##0.0000"; $gs.Range("H2:H$row").NumberFormat = "#,##0.00"; $gs.Range($gs.Cells(2,9), $gs.Cells($row,20)).NumberFormat = "#,##0"
  $h1 = $gs.Range($gs.Cells(1,1), $gs.Cells(1,22)); $h1.Font.Bold = $true; $h1.Interior.Color = $HEAD; $h1.WrapText = $true
  $gs.Columns(10).Delete() | Out-Null   # 'MES 행 수' 열 삭제 (사용자 지시 11)
  $gs.Range($gs.Cells(1,1), $gs.Cells($row,21)).Columns.AutoFit() | Out-Null
  $gs.Activate() | Out-Null; $xl.ActiveWindow.SplitRow = 1; $xl.ActiveWindow.SplitColumn = 0; $xl.ActiveWindow.FreezePanes = $true

  # 15) 국내매출 시트: MES 국내 출고(신코·코마츠 PDF 건 제외) 고객사별 1줄 + 아마란스 전자세금계산서(매출) 공급가액 비교
  $kNorm = { param($x) (("$x" -replace '\(주\)|㈜|주식회사|\(유\)|\s', '' -replace '에스앤티', 'SNT' -replace '아트라스콥코코리아|아트라스콥코', 'ATLASKOREA').ToUpper()) }   # 거래처명 대응표
  $taxF = @(Get-ChildItem $base -File | ? { $_.Name -match '세금계산서|taxData' -and $_.Extension -match '^\.xls' -and $_.Name -notmatch '^~\$' } | sort { if ($_.Name -match '^9월') { 0 } else { 1 } }, Name) | select -First 1
  $tax = @(); $closeYM = 202609   # 마감 월 (작성일 기준으로 이 달만 MES와 비교)
  $toDate = { param($x) if ($x -is [double]) { [DateTime]::FromOADate($x) } else { $p0 = [DateTime]::MinValue; if ([DateTime]::TryParse("$x".Trim(), [ref]$p0)) { $p0 } else { $null } } }
  if ($taxF) {
    $wbT = $xl.Workbooks.Open($taxF.FullName, 0, $true); $sT = $wbT.Worksheets.Item(1); $tv = $sT.UsedRange.Value2; $tn = $sT.UsedRange.Rows.Count; $tc = $sT.UsedRange.Columns.Count
    $th = $null; for ($i = 1; $i -le [math]::Min($tn, 15) -and -not $th; $i++) { $rowTxt = @(1..$tc | % { "$($tv[$i,$_])".Trim() }); if ($rowTxt -contains '거래처명' -and $rowTxt -contains '공급가액') { $th = $i } }
    if (-not $th) { $wbT.Close($false); throw "세금계산서 파일 머리글(거래처명·공급가액)을 못 찾음: $($taxF.Name)" }
    $tcol = @{}; for ($j = 1; $j -le $tc; $j++) { $h = "$($tv[$th,$j])".Trim(); if ($h -and -not $tcol.ContainsKey($h)) { $tcol[$h] = $j } }
    $bizC = @($tcol.Keys | ? { $_ -match '사업자' })[0]; $serC = @($tcol.Keys | ? { $_ -match '승인번호|일련번호' })[0]
    for ($i = $th + 1; $i -le $tn; $i++) {
      $nm = "$($tv[$i,$tcol['거래처명']])".Trim(); if (-not $nm) { continue }
      if ($tcol['구분'] -and "$($tv[$i,$tcol['구분']])" -notmatch '매출') { continue }   # 매입 제외
      $dtW = & $toDate $tv[$i,$tcol['작성일']]; $dtI = if ($tcol['발급일']) { & $toDate $tv[$i,$tcol['발급일']] } else { $null }
      $dts = if ($dtW) { $dtW.ToString('MM-dd') } else { "$($tv[$i,$tcol['작성일']])".Trim() }
      $tax += [pscustomobject]@{ Name = $nm; Key = (& $kNorm $nm); Biz = $(if ($bizC) { "$($tv[$i,$tcol[$bizC]])".Trim() }); D = $dts; Kind = $(if ($tcol['분류']) { "$($tv[$i,$tcol['분류']])".Trim() })
        YM = $(if ($dtW) { $dtW.Year * 100 + $dtW.Month } else { $closeYM }); IssueYM = $(if ($dtI) { $dtI.Year * 100 + $dtI.Month } else { $null }); Issue = $(if ($dtI) { $dtI.ToString('yyyy-MM-dd') } else { '' }); Wrote = $(if ($dtW) { $dtW.ToString('yyyy-MM-dd') } else { '' })
        Item = $(if ($tcol['대표품목']) { "$($tv[$i,$tcol['대표품목']])".Trim() } else { '' })
        Serial = $(if ($serC) { "$($tv[$i,$tcol[$serC]])".Trim() } else { '' })
        Sup = [double](Num $tv[$i,$tcol['공급가액']]); Vat = $(if ($tcol['세액']) { [double](Num $tv[$i,$tcol['세액']]) } else { 0.0 }); Tot = $(if ($tcol['합계금액']) { [double](Num $tv[$i,$tcol['합계금액']]) } else { 0.0 }) }
    }
    $wbT.Close($false)
  }
  # 일별마감(예: 2609 시트): 9월 출하 / 전월 출하(9월 매출) / 반품·불량(-) / 단가 소급 / 10월 이월(매출 제외) / 불량반출(참고만). 소계 줄(날짜·품번 없음) 제외
  $mNo = $closeYM % 100; $nextNo = $mNo % 12 + 1; $mLab = "$($mNo)월"; $nLab = "$($nextNo)월"
  $pdfCustK = @{}; for ($r = 2; $r -le $last; $r++) { if ($pdf[$lots[$r]]) { $pdfCustK[(& $kNorm $v[$r,$C['고객사']])] = $true } }
  $daily = @(); $dPath = "$base\2026년 일별마감list.xlsx"; $dSheet = "$($closeYM - 200000)"
  if (Test-Path $dPath) {
    $wbD = $xl.Workbooks.Open($dPath, 0, $true); $sD = $wbD.Worksheets.Item($dSheet); $dn = $sD.UsedRange.Rows.Count; $dv = $sD.Range("A1:I$dn").Value2; $dfm = $sD.Range("H1:H$dn").Formula; $wbD.Close($false)
    # 고객사별 하늘색 소계 줄(=SUM(H a:H b)) 범위 = 이번 달 매출(세금계산서와 같음). 소계가 있는 고객사의 범위 밖 줄은 비고가 없어도 이월 (사용자 확인 2026-10-29)
    $covD = @{}; $subK = @{}
    for ($i = 2; $i -le $dn; $i++) { if ("$($dfm[$i,1])" -match '^=SUM\(H(\d+):H(\d+)\)$') { $a1 = [int]$matches[1]; $b1 = [int]$matches[2]; for ($k1 = $a1; $k1 -le $b1; $k1++) { $covD[$k1] = $i }; $subK[(& $kNorm $dv[$a1,2])] = $true } }
    for ($i = 2; $i -le $dn; $i++) {
      $c0 = "$($dv[$i,2])".Trim(); if (-not $c0) { continue }
      $d0 = $dv[$i,1]; $pn = "$($dv[$i,4])".Trim()
      if (-not ($d0 -is [double]) -and -not $pn) { continue }   # 소계 줄
      $k0 = & $kNorm $c0; if ($pdfCustK[$k0]) { continue }       # 신코·코마츠 = 해외매출(PDF 기준)
      $ym = if ($d0 -is [double]) { $dd = [DateTime]::FromOADate($d0); $dd.Year * 100 + $dd.Month } else { 0 }
      $b0 = "$($dv[$i,9])".Trim(); $q0 = [double](Num $dv[$i,6]); $a0 = [double](Num $dv[$i,8])
      $cat = if ($b0 -match '불량반출') { 'bad' } elseif ($b0 -match "^$($nextNo)월\s*이월") { 'carry' } elseif ($subK[$k0] -and -not $covD[$i]) { if (-not $b0) { $b0 = '(비고 없음 · 일별마감 소계 범위 밖)' }; 'carry' } elseif ($b0 -match '소급') { 'retro' } elseif ($a0 -lt 0 -or $q0 -lt 0) { 'ret' } elseif ($ym -eq $closeYM) { 'cur' } elseif ($ym) { 'prev' } else { 'nodate' }
      # 불량대체: 일별마감 금액 칸(또는 비고)에 '불량대체', 원래 금액은 비고에 적힘 (예: 맥시스 금액 '불량대체' / 비고 9219000) → 매출 아님, MES와 차이 설명용
      $bd0 = if ("$($dv[$i,8])$b0" -match '불량대체') { $n0 = Num ($b0 -replace '[^\d\.]', ''); if ($n0) { [double]$n0 } else { 0.0 } } else { $null }
      $daily += [pscustomobject]@{ R = $i; Cust = $c0; Key = $k0; YM = $ym; Cat = $cat; Q = $q0; A = $a0; B = $b0; BD = $bd0 }
    }
  }
  # 고객사 묶기: MES · 일별마감 · 세금계산서를 거래처명으로 맞춤(㈜·(주)·주식회사·띄어쓰기 빼고 같으면 → 아니면 앞부분이 같은 곳이 하나뿐일 때)
  $dom = [ordered]@{}
  $newG = { param($name, $key) [pscustomobject]@{ Cust = $name; Key = $key; Rows = (New-Object System.Collections.ArrayList); Tax = (New-Object System.Collections.ArrayList); Day = (New-Object System.Collections.ArrayList) } }
  $findG = { param($key) $h0 = @($dom.Keys | ? { $dom[$_].Key -eq $key }); if (-not $h0.Count -and $key.Length -ge 2) { $h0 = @($dom.Keys | ? { $dom[$_].Key.Length -ge 2 -and ($dom[$_].Key.StartsWith($key) -or $key.StartsWith($dom[$_].Key)) }) }; if ($h0.Count -eq 1) { $h0[0] } else { $null } }
  for ($r = 2; $r -le $last; $r++) {
    if ("$($v[$r,$C['구분']])".Trim() -ne '국내' -or $pdf[$lots[$r]] -or $skOut[$lots[$r]]) { continue }
    $cu = "$($v[$r,$C['고객사']])".Trim(); if (-not $cu) { $cu = '(고객사 없음)' }
    if (-not $dom.Contains($cu)) { $dom[$cu] = & $newG $cu (& $kNorm $cu) }
    [void]$dom[$cu].Rows.Add($r)
  }
  foreach ($a in @($acc.Values | ? { $_.Mj -match '전자세금계산서' } | sort Inv)) {   # 국·해외: MES는 해외 출고지만 전자세금계산서 거래 → 세금계산서와 비교, 이번 달 매출 = 회계 원화
    $gk = "$($a.Cust) (국·해외)"
    if (-not $dom.Contains($gk)) { $dom[$gk] = & $newG $gk (& $kNorm $a.Cust); $dom[$gk] | Add-Member NoteProperty Acc (New-Object System.Collections.ArrayList) }
    [void]$dom[$gk].Acc.Add($a); for ($r = 2; $r -le $last; $r++) { if ($lots[$r] -eq $a.Inv) { [void]$dom[$gk].Rows.Add($r) } }
  }
  foreach ($x in $daily) { $gk = & $findG $x.Key; if (-not $gk) { $gk = "일별마감:$($x.Cust)"; if (-not $dom.Contains($gk)) { $dom[$gk] = & $newG $x.Cust $x.Key } }; [void]$dom[$gk].Day.Add($x) }
  # 작성일이 마감 월이 아닌 세금계산서는 비교에서 빼고 참고로 (전월 = 전월 매출 정정분, 다음 달 = 다음 달 매출)
  $domNote = @()
  foreach ($t in @($tax | ? { $_.YM -ne $closeYM })) {
    $ym0 = "{0}-{1:00}" -f [math]::Floor($t.YM / 100), ($t.YM % 100)
    $why0 = if ($t.YM -lt $closeYM) { "작성일 $($t.Wrote) = 전월($ym0) 매출 정정분 → 이번 달 비교에서 제외 (참고)" } else { "작성일 $($t.Wrote) = $ym0 매출 → 이번 달 비교에서 제외 (참고)" }
    $domNote += ,@('국내매출', '', $t.Name, "세금계산서 $($t.Wrote) $($t.Kind)", "작성일 $($t.Wrote) / 발급일 $($t.Issue)", ('공급가액 {0:N0} ({1})' -f $t.Sup, $t.Kind), "아마란스 세금계산서 ($($taxF.Name))", $why0)
  }
  $tax = @($tax | ? { $_.YM -eq $closeYM })
  $taxNon = @()
  foreach ($t in $tax) {
    if ($t.Item -match '외주가공비|고철|AL\s*재|차량\s*매각|임차료') { $taxNon += $t; continue }
    $gk = & $findG $t.Key
    if (-not $gk) { $cand = @($dom.Keys | ? { $dom[$_].Key.Length -ge 2 -and ($dom[$_].Key.StartsWith($t.Key) -or $t.Key.StartsWith($dom[$_].Key)) } | ? { $ms = [double](($dom[$_].Rows | % { [double](Num $v[$_, $C['금액']]) } | measure -Sum).Sum); [math]::Abs($ms - $t.Sup) -lt 1 }); if ($cand.Count -eq 1) { $gk = $cand[0] } }
    if (-not $gk) { $taxNon += $t; continue }   # 국내매출·국해외 어디에도 없는 거래처 = 제품 외 매출로 보고 비교 안 함
    [void]$dom[$gk].Tax.Add($t)
  }

  foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq '국내매출') { $s.Delete() } }
  $ks = $wb.Worksheets.Add([Type]::Missing, $gs); $ks.Name = '국내매출'
  $kh = @('고객사', 'MES 출고수량', 'MES 금액', "일별마감 $mLab 출하 수량", "일별마감 $mLab 출하 금액", '출하 수량 차이(일별마감-MES)', '출하 금액 차이(일별마감-MES)',
          "전월 출하분($mLab 매출)", '반품·불량(-)', '단가 소급', "$nLab 이월(매출 제외)", "$mLab 매출(일별마감)",
          '거래처명(세금계산서)', '작성일', '세금계산서 매수', '공급가액(세금계산서)', '세액', '합계금액', "세금계산서 차이(공급가액-$mLab 매출)") + $cats + @('출처 기준처', '에러내용')
  for ($j = 0; $j -lt $kh.Count; $j++) { $ks.Cells(1, $j + 1).Value2 = [string]$kh[$j] }
  $nK = $kh.Count; $krow = 2; $domShip = @(); $domErr = @(); $domCarry = @(); $taxErr = @()
  $sumCat = { param($xs, $cats0, $f) [double](($xs | ? { $cats0 -contains $_.Cat } | measure $f -Sum).Sum) }
  foreach ($g in @($dom.Values)) {
    $note = @(); $rowsG = @($g.Rows); $tx = @($g.Tax); $dy = @($g.Day)
    $ks.Cells($krow, 1).Value2 = [string]$g.Cust
    $qty = 0.0; $amt = 0.0
    if ($rowsG.Count) {
      $qty = [double](($rowsG | % { [double](Num $v[$_, $C['출고수량']]) } | measure -Sum).Sum); $amt = [double](($rowsG | % { [double](Num $v[$_, $C['금액']]) } | measure -Sum).Sum)
      $ks.Cells($krow, 2).Value2 = $qty; $ks.Cells($krow, 3).Value2 = $amt
      $k0 = 20; foreach ($k in $cats) { $sv0 = [double](($rowsG | % { [double](Num $v[$_, $C[$k]]) } | measure -Sum).Sum); if ($sv0) { $ks.Cells($krow, $k0).Value2 = $sv0 }; $k0++ }
    }
    # 일별마감 나눠 담기
    $shipQ = [double](($dy | ? { $_.YM -eq $closeYM -and ($_.Cat -eq 'cur' -or $_.Cat -eq 'carry') } | measure Q -Sum).Sum)
    $shipA = [double](($dy | ? { $_.YM -eq $closeYM -and ($_.Cat -eq 'cur' -or $_.Cat -eq 'carry') } | measure A -Sum).Sum)
    $carry9 = [double](($dy | ? { $_.YM -eq $closeYM -and $_.Cat -eq 'carry' } | measure A -Sum).Sum)
    $prevA = & $sumCat $dy @('prev') 'A'; $retA = & $sumCat $dy @('ret') 'A'; $retroA = & $sumCat $dy @('retro') 'A'; $carryA = & $sumCat $dy @('carry') 'A'; $nodA = & $sumCat $dy @('nodate') 'A'
    $saleA = (& $sumCat $dy @('cur') 'A') + $prevA + $retA + $retroA + $nodA
    $hasDay = @($dy | ? { $_.Cat -ne 'bad' }).Count -gt 0
    if ($hasDay) {
      $ks.Cells($krow, 4).Value2 = $shipQ; $ks.Cells($krow, 5).Value2 = $shipA
      foreach ($pair in @(@(8, $prevA), @(9, $retA), @(10, $retroA), @(11, -$carryA))) { if ($pair[1]) { $ks.Cells($krow, $pair[0]).Value2 = [double]$pair[1] } }
      $ks.Cells($krow, 12).Value2 = $saleA
      if ($nodA) { $note += "일별마감에 출하일이 비어 있는 줄 {0:N0}원은 {1} 매출에 넣음 (참고)" -f $nodA, $mLab }
      if ($retA) { $note += "반품·불량으로 {0:N0}원 차감됨 (일별마감)" -f [math]::Abs($retA) }
      if ($retroA) { $note += "단가 소급 {0:N0}원 (일별마감)" -f $retroA }
      if ($carryA) { $note += "$nLab 이월 {0:N0}원은 {1} 매출에서 뺌 → 이월매출 시트" -f $carryA, $mLab
        $domCarry += [pscustomobject]@{ Cust = $g.Cust; Q = (& $sumCat $dy @('carry') 'Q'); A = $carryA; Rows = (@($dy | ? { $_.Cat -eq 'carry' } | % R) -join ', '); B = ((@($dy | ? { $_.Cat -eq 'carry' } | % B | select -Unique)) -join ' / '); Mes = (@($rowsG | sort) -join ', ') } }
      $nBad = @($dy | ? { $_.Cat -eq 'bad' }).Count; if ($nBad) { $note += "불량반출 기록 $($nBad)건은 매출이 아니라 참고로만 둠" }
    }
    # 불량대체: MES 금액 빈 줄 / 일별마감 '불량대체' 줄 → 정상
    $bdM = @($rowsG | ? { $bulD[$_] }); $bdD = @($dy | ? { $_.BD -ne $null -and $_.YM -eq $closeYM }); $bdDA = [double](($bdD | measure BD -Sum).Sum)
    if ($bdM.Count) { $note += "MES 금액이 빈 {0}줄(MES {1}행, 수량 {2:N0}개)은 불량대체 → 정상 (국내 영업팀 확인)" -f $bdM.Count, ((@($bdM | sort)) -join ', '), [double](($bdM | % { [double](Num $v[$_, $C['출고수량']]) } | measure -Sum).Sum) }
    if ($bdD.Count) { $note += "일별마감 {0} 시트 {1}행은 불량대체(금액 {2:N0}원) → 매출 아님, 정상" -f $mLab, ((@($bdD | % R)) -join ', '), $bdDA }
    $isKH = [bool]$g.PSObject.Properties['Acc']; $khMiss = 0.0
    if ($isKH) { $saleA = [double](($g.Acc | measure Krw -Sum).Sum); $ks.Cells($krow, 12).Value2 = $saleA
      $mesLots = @($rowsG | % { $lots[$_] }); $khMiss = [double](($g.Acc | ? { $mesLots -notcontains $_.Inv } | measure Krw -Sum).Sum)
      $note += "국·해외: MES는 해외 출고지만 전자세금계산서로 거래 → 회계 원화 {0:N0}원({1})을 {2} 매출로 봄" -f $saleA, ((@($g.Acc | % Inv)) -join ', '), $mLab }
    # ① MES 출고 vs 일별마감 출하 (반품·재출고로 MES 수량·금액이 바뀌었는지)
    if ($rowsG.Count -and $hasDay) {
      $dq = $shipQ - $qty; $da = $shipA - $amt
      if ([math]::Abs($dq) -ge 0.5) { $ks.Cells($krow, 6).Value2 = $dq }; if ([math]::Abs($da) -ge 1) { $ks.Cells($krow, 7).Value2 = $da }
      if ([math]::Abs($dq) -ge 0.5 -or [math]::Abs($da) -ge 1) {
        # 사용자 규칙(2026-10-29): 일별마감이 MES보다 적으면(-) 반품·불량, 많으면(+) 단가 소급으로 봄 → 에러 아님, 출하금액변동으로
        $minus = if ([math]::Abs($da) -ge 1) { $da -lt 0 } else { $dq -lt 0 }
        $estLab = if ($minus) { '반품·불량' } else { '단가 소급' }
        foreach ($cc in 6, 7) { $ks.Cells($krow, $cc).Interior.Color = 10284031 }   # 연한 노랑 = 판단해서 분류한 칸
        if (($bdM.Count -or $bdD.Count) -and [math]::Abs($da + $bdDA) -lt 1) {   # 금액 차이가 불량대체로 다 설명됨(MES 금액 빈 줄은 금액 0이라 수량만 차이)
          $note += $(if ($bdDA) { "MES 출고보다 일별마감 $mLab 출하 금액이 {0:N0}원 적음 = 일별마감 불량대체 금액 → 불량대체(정상)" -f $bdDA }
                     else { "MES 출고와 일별마감 $mLab 출하가 수량만 {0:N0}개 다르고 금액은 같음 → MES 금액이 빈 불량대체 줄 때문(정상)" -f $dq })
        } else {
          $note += "MES 출고보다 일별마감 $mLab 출하가 {0} (수량 {1:N0}개 / 금액 {2:N0}원) → {3}(으)로 봄 (MES와 일별마감 차이로 판단{4})" -f $(if ($minus) { '적음' } else { '많음' }), $dq, $da, $estLab, $(if ($bdM.Count -or $bdD.Count) { ', 불량대체 포함' } else { '' })
        }
      }
    } elseif ($rowsG.Count -and $daily.Count -and -not $isKH) {
      $ks.Cells($krow, 5).Interior.Color = $RED; $msg = "MES에는 $mLab 출고가 있는데 일별마감에 없음 → 일별마감 누락인지 확인해 주세요"
      $note += $msg; $domErr += ,@('국내매출', "국내매출 R$krow", $g.Cust, '', ('MES 금액 {0:N0}원' -f $amt), '일별마감 없음', "일별마감 $mLab 시트", $msg)
    } elseif ($hasDay -and $shipA) {
      $ks.Cells($krow, 3).Interior.Color = $RED; $msg = "{0} 매출은 맞음(일별마감 {1:N0}원, {0} 매출에 반영함). 그런데 MES에 출고 입력이 빠져 있음 → MES에 출고를 입력해 주세요" -f $mLab, $shipA
      $note += $msg; $domErr += ,@('국내매출', "국내매출 R$krow", $g.Cust, '', 'MES 없음', ('일별마감 금액 {0:N0}원' -f $shipA), "일별마감 $mLab 시트", $msg)
    }
    # ② 세금계산서 vs 일별마감 매출 → 다르면 '전자세금계산서 금액 오류' 시트 (15)
    $sup = 0.0
    $serAll = ((@($tx | % Serial | ? { $_ } | select -Unique)) -join "`n")
    if ($tx.Count) {
      $ks.Cells($krow, 13).Value2 = [string]((@($tx | % Name | select -Unique)) -join ', ')
      $ks.Cells($krow, 14).NumberFormat = "@"; $ks.Cells($krow, 14).Value2 = [string]((@($tx | % D | select -Unique)) -join ', ')
      $ks.Cells($krow, 15).Value2 = [double]$tx.Count
      $sup = [double](($tx | measure Sup -Sum).Sum); $ks.Cells($krow, 16).Value2 = $sup
      $ks.Cells($krow, 17).Value2 = [double](($tx | measure Vat -Sum).Sum); $ks.Cells($krow, 18).Value2 = [double](($tx | measure Tot -Sum).Sum)
      if (@($tx | ? { $_.Kind -match '수정' }).Count) { $note += "수정세금계산서 포함(원본과 합친 금액)" }
      $late = @($tx | ? { $_.Kind -match '수정' -and $_.IssueYM -and $_.IssueYM -gt $closeYM })
      if ($late.Count) { $note += "다음 달 발급 수정세금계산서 포함(발급일 $((@($late | % Issue | select -Unique)) -join ', '), 작성일 기준 이번 달 반영)" }
      # mes9월 '전자세금계산서 금액'·'세금계산서 일련번호': 작성일이 같은 세금계산서끼리 묶고, MES 행은 출고일 이후 가장 가까운 작성일 묶음에 넣어 MES 금액 비율로 나눔
      if ($rowsG.Count) {
        $bk = @($tx | group Wrote | sort Name | % { [pscustomobject]@{ D = $_.Name; Sup = [double](($_.Group | measure Sup -Sum).Sum); Ser = ((@($_.Group | % Serial | ? { $_ })) -join "`n"); Rows = (New-Object System.Collections.ArrayList) } })
        foreach ($r0 in @($rowsG | sort { [double](Num $v[$_, $C['출고일']]) })) {
          $sd = [double](Num $v[$r0, $C['출고일']]); $sds = if ($sd) { [DateTime]::FromOADate($sd).ToString('yyyy-MM-dd') } else { '' }
          $b1 = @($bk | ? { $_.D -ge $sds })[0]; if (-not $b1) { $b1 = $bk[-1] }; [void]$b1.Rows.Add($r0)
        }
        foreach ($b1 in $bk) {
          $rs1 = @($b1.Rows); if (-not $rs1.Count) { continue }
          $am1 = @($rs1 | % { [double](Num $v[$_, $C['금액']]) }); $s1 = [double](($am1 | measure -Sum).Sum); $acc1 = 0.0
          $okSum = ($hasDay -or $isKH) -and [math]::Abs($sup - $saleA) -lt 1   # (가) 고객사 합계가 세금계산서와 맞으면 행에는 MES 금액 그대로 (사용자 선택 2026-10-29)
          for ($i1 = 0; $i1 -lt $rs1.Count; $i1++) {
            $val1 = if ($okSum) { $am1[$i1] } elseif ($i1 -eq $rs1.Count - 1) { $b1.Sup - $acc1 } elseif ($s1) { [math]::Round($b1.Sup * $am1[$i1] / $s1) } else { 0.0 }
            $acc1 += $val1; $ws.Cells($rs1[$i1], $cTaxA).Value2 = [double]$val1
            if ($b1.Ser) { $ws.Cells($rs1[$i1], $cTaxN).NumberFormat = "@"; $ws.Cells($rs1[$i1], $cTaxN).Value2 = [string]$b1.Ser }
          }
        }
      }
      if ($hasDay -or $isKH) {
        $td = $sup - $saleA; if ([math]::Abs($td) -ge 1) { $ks.Cells($krow, 19).Value2 = $td; $ks.Cells($krow, 19).Interior.Color = $RED
          $msg = "세금계산서 공급가액 {0:N0}원이 {1} 매출(일별마감) {2:N0}원과 {3:N0}원 다름 → 계산서 금액 확인(수정세금계산서 필요 여부)" -f $sup, $mLab, $saleA, $td
          $note += "$msg (전자세금계산서 금액 오류 시트)"; $taxErr += [pscustomobject]@{ Cust = $g.Cust; Sale = $saleA; Mes = $amt; Sup = $sup; Diff = $td; N = $tx.Count; D = ((@($tx | % D | select -Unique)) -join ', '); Ser = $serAll; Why = $msg } }
      } else { $ks.Cells($krow, 12).Interior.Color = $RED; $msg = "세금계산서는 있는데 일별마감 $mLab 매출에 없음 → 거래처명이 다르거나 전월분인지 확인해 주세요"
        $note += "$msg (전자세금계산서 금액 오류 시트)"; $taxErr += [pscustomobject]@{ Cust = $g.Cust; Sale = $null; Mes = $amt; Sup = $sup; Diff = $sup; N = $tx.Count; D = ((@($tx | % D | select -Unique)) -join ', '); Ser = $serAll; Why = $msg } }
    } elseif ($taxF -and ($hasDay -or $isKH) -and [math]::Abs($saleA) -ge 1) {
      $ks.Cells($krow, 16).Interior.Color = $RED; $msg = "{0} 매출(일별마감) {1:N0}원이 있는데 세금계산서 없음 → 발행 전·거래처명 확인해 주세요" -f $mLab, $saleA
      $note += "$msg (전자세금계산서 금액 오류 시트)"; $taxErr += [pscustomobject]@{ Cust = $g.Cust; Sale = $saleA; Mes = $amt; Sup = $null; Diff = -$saleA; N = 0; D = ''; Ser = ''; Why = $msg }
    }
    # ③ 출하기준 금액변동: MES 출하 기준(이월분 제외) vs 확정 매출(세금계산서, 없으면 일별마감)
    $fix = if ($tx.Count) { $sup } elseif ($hasDay -or $isKH) { $saleA } else { $null }
    if ($fix -ne $null -and $rowsG.Count) {
      $baseA = $amt - $carry9; $chg = $fix - $baseA
      if ([math]::Abs($chg) -ge 1) {
        $other = $chg - ($prevA + $retA + $retroA + $nodA + $khMiss - $bdDA)
        $parts = @(); if ($prevA) { $parts += "전월 출하분 +{0:N0}" -f $prevA }; if ($retA) { $parts += "반품·불량 {0:N0}" -f $retA }; if ($retroA) { $parts += "단가 소급 {0:+#,##0;-#,##0}" -f $retroA }; if ($nodA) { $parts += "출하일 없는 줄 +{0:N0}" -f $nodA }
        if ($bdDA) { $parts += "불량대체(정상) -{0:N0}" -f $bdDA }
        if ($khMiss) { $parts += "MES {0} 출고에 없는 회계 건 +{1:N0}" -f $mLab, $khMiss }
        if ([math]::Abs($other) -ge 1) { $parts += "{0}(MES와 일별마감 차이로 판단) {1:+#,##0;-#,##0}" -f $(if ($other -lt 0) { '반품·불량' } else { '단가 소급' }), $other }
        $domShip += ,@('출하기준 금액변동', '국내', $(if ($tx.Count) { "세금계산서 $($tx.Count)매" } else { '일별마감' }), $g.Cust, 'KRW', ((@($rowsG | sort)) -join ', '), $qty, $baseA, $fix, $chg, $null, $null, $null, $null, $(if ($baseA) { '{0:+0.00;-0.00}%' -f ($chg / $baseA * 100) } else { '' }),
          $(if ($tx.Count) { "아마란스 세금계산서" } else { "일별마감 $mLab 시트" }), ("MES 출하 기준(이월분 제외) → 확정 매출 차이: " + ($parts -join ' / ')))
      }
    }
    $ks.Cells($krow, $nK - 1).Value2 = [string]((@('MES') + @($(if ($hasDay) { '일별마감' })) + @($(if ($tx.Count) { '세금계산서' }))) -join ' + ')
    if ($note) { $ks.Cells($krow, $nK).Value2 = [string]($note -join ' / ') }
    $krow++
  }
  $ks.Cells($krow, 1).Value2 = "합계"; $ks.Range($ks.Cells($krow,1), $ks.Cells($krow,$nK)).Font.Bold = $true
  foreach ($cc in @(2..12) + @(15..($nK - 2))) { $cl = ($ks.Cells(1,$cc).Address($false,$false) -replace '\d',''); $ks.Cells($krow, $cc).Formula = "=SUM(${cl}2:${cl}$($krow-1))" }
  $ks.Range($ks.Cells(2,2), $ks.Cells($krow,$nK - 2)).NumberFormat = "#,##0"
  $h1 = $ks.Range($ks.Cells(1,1), $ks.Cells(1,$nK)); $h1.Font.Bold = $true; $h1.Interior.Color = $HEAD; $h1.WrapText = $true; $ks.Rows(1).RowHeight = 45
  for ($cc = 1; $cc -lt $nK; $cc++) { $ks.Columns($cc).ColumnWidth = 13 }; $ks.Columns(1).ColumnWidth = 22; $ks.Columns(13).ColumnWidth = 20; $ks.Columns($nK - 1).ColumnWidth = 22
  $ks.Columns($nK).ColumnWidth = 70; $ks.Columns($nK).WrapText = $true
  $ks.Activate() | Out-Null; $xl.ActiveWindow.SplitRow = 1; $xl.ActiveWindow.SplitColumn = 1; $xl.ActiveWindow.FreezePanes = $true
  $nDom = $krow - 2
  foreach ($r in @($bulD.Keys)) { $ws.Cells($r, $cTaxA).NumberFormat = "@"; $ws.Cells($r, $cTaxA).Value2 = '불량대체' }   # MES 금액 빈 행 = 불량대체(정상)

  # 15) '전자세금계산서 금액 오류' 시트: 국내매출(9월 매출) ≠ 세금계산서 공급가액, 또는 한쪽에만 있는 고객사
  $tsName = '전자세금계산서 금액 오류'
  foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq $tsName) { $s.Delete() } }
  $ts = $wb.Worksheets.Add([Type]::Missing, $ks); $ts.Name = $tsName
  $th0 = @('고객사', "$mLab 매출(국내매출·일별마감)", 'MES 금액', '공급가액(세금계산서)', "차이(공급가액-$mLab 매출)", '세금계산서 매수', '작성일', '세금계산서 일련번호', '사유')
  for ($j = 0; $j -lt $th0.Count; $j++) { $ts.Cells(1, $j + 1).Value2 = [string]$th0[$j] }
  $trow = 2
  # 오류가 없으면 [결과]·[안내] 문구만 (빈 줄 없이) → HTML은 이런 줄을 표가 아닌 설명으로 보여 주고 개수에서 뺌
  if (-not $taxF) { $ts.Cells(2, 1).Value2 = "[안내] 아직 세금계산서 파일이 없음 → 기초자료 월 폴더에 아마란스 매출 세금계산서(작성일 그달 1일~말일)를 넣으면 자동으로 비교합니다"; $trow = 3 }
  elseif (-not $taxErr.Count) { $ts.Cells(2, 1).Value2 = "[결과] 오류 0건 — 국내매출 금액과 세금계산서 금액이 모두 맞음"; $trow = 3 }
  else {
    foreach ($x in $taxErr) {
      $ts.Cells($trow, 1).Value2 = [string]$x.Cust
      foreach ($pair in @(@(2, $x.Sale), @(3, $x.Mes), @(4, $x.Sup), @(5, $x.Diff), @(6, $x.N))) { if ($pair[1] -ne $null) { $ts.Cells($trow, $pair[0]).Value2 = [double]$pair[1] } }
      $ts.Cells($trow, 7).NumberFormat = "@"; $ts.Cells($trow, 7).Value2 = [string]$x.D
      $ts.Cells($trow, 8).NumberFormat = "@"; $ts.Cells($trow, 8).Value2 = [string]$x.Ser
      $ts.Cells($trow, 9).Value2 = [string]$x.Why; $ts.Cells($trow, 5).Interior.Color = $RED
      $trow++
    }
    $ts.Cells($trow, 1).Value2 = '합계'; foreach ($cc in 2..5) { $cl = ($ts.Cells(1,$cc).Address($false,$false) -replace '\d',''); $ts.Cells($trow, $cc).Formula = "=SUM(${cl}2:${cl}$($trow-1))" }; $ts.Range($ts.Cells($trow,1), $ts.Cells($trow,9)).Font.Bold = $true
  }
  if ($taxNon.Count) { if ($taxErr.Count) { $trow++ }; $ts.Cells($trow, 1).Value2 = [string]("[안내] 제품 외 매출 {0}매 {1:N0}원은 비교하지 않음: {2}" -f $taxNon.Count, ($taxNon | measure Sup -Sum).Sum, ((@($taxNon | % { "$($_.Name)($($_.Item))" } | select -Unique)) -join ', ')) }
  $ts.Range("B2:F$trow").NumberFormat = "#,##0"
  $h1 = $ts.Range($ts.Cells(1,1), $ts.Cells(1,9)); $h1.Font.Bold = $true; $h1.Interior.Color = $HEAD; $h1.WrapText = $true
  $ts.Columns("A:H").AutoFit() | Out-Null; foreach ($cc in 1, 7, 8) { if ($ts.Columns($cc).ColumnWidth -gt 30) { $ts.Columns($cc).ColumnWidth = 30; $ts.Columns($cc).WrapText = $true } }; $ts.Columns(9).ColumnWidth = 70; $ts.Columns(9).WrapText = $true
  $ts.Activate() | Out-Null; $xl.ActiveWindow.SplitRow = 1; $xl.ActiveWindow.SplitColumn = 0; $xl.ActiveWindow.FreezePanes = $true
  $nTaxErr = $taxErr.Count

  # 11) 오류 시트 합치기 → '오류목록' 한 장 (분류: 에러 / 10월 이월 / 참고)
  function T($x, $date) { if ($x -eq $null -or "$x" -eq '') { '' } elseif ($date -and $x -is [double]) { [DateTime]::FromOADate($x).ToString('yyyy-MM-dd') } elseif ($x -is [double]) { $x.ToString('#,##0.####') } else { "$x".Trim() } }
  # 시트명 → 열 매핑: 행, 고객사, 인보이스, 기준값, 비교값, 출처, 사유 (숫자=열 번호, 문자열=고정값, 스크립트블록=계산)
  $maps = [ordered]@{
    'BL불일치사유'   = @{ Item = 'BL날짜';   Map = { param($v,$i) @((T $v[$i,1]), (T $v[$i,2]), (T $v[$i,3]), (T $v[$i,4] 1), (T $v[$i,6] 1), (T $v[$i,5]), (T $v[$i,7])) } }
    '인보이스불일치' = @{ Item = '인보이스'; Map = { param($v,$i) @((T $v[$i,1]), (T $v[$i,2]), (T $v[$i,3]), (T $v[$i,4]), (T $v[$i,6]), (T $v[$i,5]), (T $v[$i,7])) } }
    '면장불일치'     = @{ Item = '수출면장'; Map = { param($v,$i) @((T $v[$i,1]), (T $v[$i,2]), (T $v[$i,3]), (T $v[$i,4]), (T $v[$i,6]), (T $v[$i,7]), (T $v[$i,8])) } }
    '환율불일치'     = @{ Item = '환율';     Map = { param($v,$i) @((T $v[$i,1]), (T $v[$i,2]), (T $v[$i,3]), (T $v[$i,7]), (T $v[$i,6]), "smbs.biz $(T $v[$i,5]) (BL $(T $v[$i,4] 1))", (T $v[$i,8])) } }
    '달러불일치'     = @{ Item = '달러';     Map = { param($v,$i) @((T $v[$i,1]), (T $v[$i,2]), (T $v[$i,3]), '', (T $v[$i,7]), "MES 금액 $(T $v[$i,5]) ÷ 환율 $(T $v[$i,6]) ($(T $v[$i,4]))", (T $v[$i,8])) } }
    '라인구분불일치' = @{ Item = '라인구분'; Map = { param($v,$i) @((T $v[$i,1]), (T $v[$i,2]), (T $v[$i,3]), (T $v[$i,5]), (T $v[$i,6]), (T $v[$i,4]), (T $v[$i,7])) } }
    '기초자료불일치' = @{ Item = '기초자료 환율·달러'; Map = { param($v,$i) @((T $v[$i,8]), (T $v[$i,2]), (T $v[$i,1]), "환율 $(T $v[$i,5]) / 외화 $(T $v[$i,7])", "환율 $(T $v[$i,4]) / 외화 $(T $v[$i,6])", (T $v[$i,3]), (T $v[$i,9])) } }
  }
  $errs = New-Object System.Collections.ArrayList
  # 행 단위 원자료는 숨김 시트 '오류원본'에 보관(없으면 예전 형식 '오류목록'에서). 이번에 원본 시트가 없는 항목(예전 스크립트가 만든 환율·달러 등)은 그대로 이어받음
  $have = @{}; foreach ($s in @($wb.Worksheets)) { $have[$s.Name] = $s }
  $regen = @($maps.Keys | ? { $have[$_] } | % { $maps[$_].Item }) + @('면장별합계', '해외매출', '국내매출')
  $old = if ($have['오류원본']) { $have['오류원본'] } elseif ($have['오류목록'] -and "$($have['오류목록'].Range('B1').Value2)" -eq '항목') { $have['오류목록'] } else { $null }
  if ($old) {
    $off = if ($old.Name -eq '오류원본') { 1 } else { 2 }   # 예전 오류목록은 A열이 분류
    $on = $old.UsedRange.Rows.Count; $ov = $old.Range($old.Cells(1,1), $old.Cells($on, 9)).Value2
    for ($i = 2; $i -le $on; $i++) { $it = "$($ov[$i,$off])"; if ($it -and $regen -notcontains $it) { [void]$errs.Add(@($off..($off + 7) | % { "$($ov[$i,$_])" })) } }
  } elseif ($have['오류목록']) { throw "오류목록이 새 형식인데 숨김 시트 '오류원본'이 없음 — 백업에서 복구 필요" }
  foreach ($sn in $maps.Keys) {
    $s = $have[$sn]; if (-not $s) { continue }
    $n = $s.UsedRange.Rows.Count
    if ($n -ge 2) { $sv = $s.Range($s.Cells(1,1), $s.Cells($n, 9)).Value2
      for ($i = 2; $i -le $n; $i++) { if ("$($sv[$i,1])$($sv[$i,3])" -eq '') { continue }; [void]$errs.Add(@($maps[$sn].Item) + @(& $maps[$sn].Map $sv $i)) } }
    $s.Delete()
  }
  # 해외매출 에러내용도 포함
  $gv = $gs.UsedRange.Value2; $gn = $gs.UsedRange.Rows.Count
  for ($i = 2; $i -le $gn; $i++) { $e = "$($gv[$i,21])"; if ($e -match '≠|비교 불가|금액 다름') { [void]$errs.Add(@('해외매출', "해외매출 R$i", "$($gv[$i,4])", "$($gv[$i,2])", '회계', '', "$($gv[$i,20])", $e)) } }
  foreach ($x in @($domErr) + @($domNote)) { [void]$errs.Add($x) }   # 국내매출: 세금계산서·MES 한쪽에만 있는 건
  # 원자료 → 숨김 시트 '오류원본' (다음 실행 때 다시 읽음)
  if ($have['오류원본']) { $have['오류원본'].Visible = -1; $have['오류원본'].Delete() }
  $rs = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)); $rs.Name = '오류원본'
  $rs.Range($rs.Cells(1,1), $rs.Cells($errs.Count + 1, 8)).NumberFormat = "@"
  $rh = @('항목', 'MES 행', '고객사', '인보이스/Lot', '기준값(회계·MES)', '비교값', '비교 출처', '사유'); for ($j = 0; $j -lt 8; $j++) { $rs.Cells(1, $j + 1).Value2 = [string]$rh[$j] }
  for ($ri = 0; $ri -lt $errs.Count; $ri++) { for ($ci = 0; $ci -lt 8; $ci++) { if ("$($errs[$ri][$ci])") { $rs.Cells(2 + $ri, 1 + $ci).Value2 = [string]$errs[$ri][$ci] } } }
  $rs.Visible = 2   # xlSheetVeryHidden
  # 분류 (MES 출고Lot 오기 DSM2608-AL6은 사용자 확인 완료 → 에러 아님, 이월로 분류)
  $nRepl = 0; $nShip = 0; $shipInv = [ordered]@{}
  $recs = @(foreach ($e in $errs) {
    $why = "$($e[7])"
    # 14) MES 출하 기준 금액 ≠ 확정 금액(회계·인보이스) → 에러 아님, '출하금액변동' 시트로
    if ($e[0] -eq '달러' -and $why -match '≠ 회계 외화금액') {
      if ("$($e[3])".Trim()) { $shipInv["$($e[3])".Trim()] = $true }
      $why = @($why -split ' / ' | ? { $_.Trim() -and $_ -notmatch '≠ 회계 외화금액' }) -join ' / '
      if (-not $why) { $nShip++; continue }
    }
    # Replacement 행(메모)만 가리키는 달러·라인구분 항목은 0이 정상 → 목록에서 제외
    $rn0 = @([regex]::Matches("$($e[1])", '\d+') | % { [int]$_.Value })
    if ("$($e[1])" -notmatch '^(해외|국내)매출' -and $rn0.Count -and @($rn0 | ? { -not $repl[$_] }).Count -eq 0 -and $e[0] -in '달러', '라인구분') { $nRepl++; continue }
    if ($e[0] -eq '수출면장' -and "$($e[3])" -match '-AT\d' -and $why -notmatch '전자세금계산서') { $why = "$why (Atlas는 전자세금계산서 거래라 면장 없음)" }
    $why = $why -replace '\s*\([^()]*시트 참고\)', '' -replace ',\s*[^,()]*시트 참고', ''   # 지금은 없는 옛 시트 안내 제거
    $cls = if ($why -match '\(알림\)') { '알림' } elseif ($why -match '이월') { '이월' }
           elseif ($e[0] -eq '달러' -or $e[0] -eq '환율' -or $why -match '참고|차액 .*가공에 반영|전자세금계산서') { '참고' }
           else { '에러' }
    # 라인 코드 줄임말(사/자/금) → 전체 이름으로 (사용자 요청)
    # 다른 사람이 봐도 바로 알 수 있게: 줄임말·셀 주소·축약 문구를 쉬운 말로 (사용자 요청)
    $baseT = "$($e[4])"; $cmpT = "$($e[5])"
    if ($e[0] -eq '라인구분') {
      $full = @{ '사' = '사형'; '자' = '자동'; '금' = '금형' }
      $mesL = ($baseT -replace '^MES 라인구분:\s*', '').Trim(); $dayL = if ($cmpT -match '일별마감 라인:\s*(\S+)') { $matches[1].Trim() } else { '' }
      if ($why -match '사/자/금이 아님') {
        $baseT = "MES: $(if ($mesL) { $mesL } else { '라인 없음' })"; $cmpT = "일별마감: '$dayL' (사형·자동·금형이 아님)"
        $why = if ($mesL) { "일별마감 파일에 라인이 '$dayL'(으)로 적혀 있어 맞는지 확인할 수 없음. MES에 적힌 '$mesL'(으)로 금액을 나눠 둠. → 일별마감 라인을 사형·자동·금형 중 하나로 고쳐 주세요 (참고)" }
                else { "일별마감 파일에는 '$dayL', MES에는 라인이 비어 있어 어느 라인(사형·자동·금형)인지 알 수 없음. → 라인을 확인해 주세요 (참고)" }
      } else {
        if ($dayL) { $cmpT = "일별마감: $(if ($full[$dayL]) { $full[$dayL] } else { $dayL })" }
        if ($baseT -match '^MES 라인구분:') { $baseT = "MES: $(if ($mesL) { $mesL } else { '라인 없음' })" }
      }
      if ($why -match 'MES 금액 없음 → 분해 불가') {   # 국내 영업팀 확인(2026-10-30): MES 금액 빈 행 = 불량대체 → 정상(참고). (10-29의 '담당자 입력 오류' 판단을 바꿈)
        $why = $why -replace 'MES 금액 없음 → 분해 불가', '불량대체: MES 금액이 비어 있는 것은 불량 대체품이라 정상 (국내 영업팀 확인). 금액이 없어 라인별로 나누지 않음'
        $cls = '참고'
      }
      $why = $why -replace '비율 차이\(0\.5%p 초과\):\s*', 'MES와 9월라인의 라인별 금액 비율이 0.5% 넘게 다름 → ' -replace ' vs 참고 ', ' / 9월라인 '
      $why = [regex]::Replace($why, '수량×\(주조비\+가공비\) ([\d,]+) ≠ 금액 ([\d,]+) → 차액 (-?[\d,]+)을 가공에 반영', '수량×단가(주조비+가공비) = $1원, MES 금액 = $2원 → 차이 $3원은 단수 차이라 가공 금액에 넣음')
      if ($baseT -match '^([^/%,\s]+)/([^/%,\s]+)$') { $baseT = "MES 라인: $($matches[1]) / 가공: $(if ($matches[2] -eq '소재') { '없음(소재)' } else { $matches[2] })" }
      $why = $why -replace '사/자/금', '사형·자동·금형'
    }
    $why = [regex]::Replace($why, '원화 오류: (\S+) ([\d,]+) ≠ 회계 ([\d,]+) \(차 (-?[\d,]+)\) - 회계\(엔화×환율\)가 맞음, PDF 사형·금형·가공·스크랩 합계가 틀림', { param($m) $dd = [double]($m.Groups[4].Value -replace ',', ''); "$($m.Groups[1].Value) 항목별 금액(사형·금형·가공·스크랩) 합계 $($m.Groups[2].Value)원이 회계 $($m.Groups[3].Value)원(엔화×환율)보다 $('{0:N0}' -f [math]::Abs($dd))원 $(if ($dd -lt 0) { '적음' } else { '많음' }) → 회계가 맞음. PDF 금액 확인 필요" })
    $why = $why -replace '9월마감 번호 비어 있음', '9월마감 파일에 면장번호가 비어 있음' -replace '원화 거래\(환율 없음\) → 달러 미기재', '원화 거래라 환율이 없어 달러를 적지 않음' -replace '원화 거래\(회계 환율 없음\) → 환율 미기재', '원화 거래라 회계에 환율이 없어 환율을 적지 않음'
    $srcT = "$($e[6])" -replace '^MES 자체$', 'MES' -replace '\(스크랩 별도 칸, 가공은 가공3\)', ''
    $srcT = [regex]::Replace($srcT, '일별마감 (\d{2})(\d{2})!R(\d+)', { param($m) "일별마감 $([int]$m.Groups[2].Value)월 시트 $($m.Groups[3].Value)행" })
    $srcT = $srcT -replace '(\S+)!R(\d+)', '''$1'' 시트 $2행'
    [pscustomobject]@{ Cls = $cls; Item = "$($e[0])"; Row = "$($e[1])".Trim(); Cust = "$($e[2])".Trim(); Inv = "$($e[3])".Trim(); Base = $baseT; Cmp = $cmpT; Src = $srcT; Why = $why }
  })
  function U($xs) { @($xs | ? { "$_".Trim() } | % { "$_".Trim() } | select -Unique) }
  function RowList($rs0) { $ns = @($rs0 | ? { $_ -notmatch '^(해외|국내)매출' } | % { [regex]::Matches($_, '\d+') | % { [int]$_.Value } } | sort -Unique); (@($ns | % { "$_" }) + @(U ($rs0 | ? { $_ -match '^(해외|국내)매출' }))) -join ', ' }

  # 12) 이월매출 시트: BL 미확정·9월라인에 없음(회계 BL일 10월) + 회계 회색 면장 → 인보이스 1건 = 1줄
  $carry = [ordered]@{}
  foreach ($x in @($recs | ? { $_.Cls -eq '이월' -and $_.Inv } | sort Cust, Inv)) { $carry[$x.Inv] = $true }
  foreach ($t in @($accList | ? { $_.Gray })) { $carry[$t.Inv] = $true }
  if ($have['이월매출']) { $have['이월매출'].Delete() }
  $cs = $wb.Worksheets.Add([Type]::Missing, $ts); $cs.Name = '이월매출'
  $ch = @('인보이스 번호', '고객사', '수출면장 번호(회계)', 'BL날짜(회계)', '환율(회계)', '외화금액(회계)', '원화금액(회계)', 'MES 행', 'MES 출고수량', 'MES 금액', '이월 사유', '참고')
  for ($j = 0; $j -lt $ch.Count; $j++) { $cs.Cells(1, $j + 1).Value2 = [string]$ch[$j] }
  $boiler = '9월라인에 인보이스 없음|10월 이월 건|BL날짜 미확정|환율 빨간색|회계에 인보이스 없음'
  $crow = 2
  foreach ($inv in @($carry.Keys)) {
    $a = $acc[$inv]; $mine = @($recs | ? { $_.Inv -eq $inv -and $_.Cls -ne '에러' })
    $mr = if ($grp.Contains($inv)) { @($grp[$inv].Rows) } else { @() }
    $cust = if ($a -and $a.Cust) { $a.Cust } else { @($mine | % Cust | ? { $_ })[0] }
    $main = @(); if ($a -and $a.Bl -ge 46296) { $main += "회계 BL일 $(DS $a.Bl) → 10월 매출" }
    if ($a -and @($accList | ? { $_.Inv -eq $inv -and $_.Gray }).Count) { $main += "회계 면장번호 회색(다음 달 이월)" }
    if (-not $line[$inv] -and -not $pdf[$inv]) { $main += "9월라인에 없음(BL 미확정)" }
    $extra = U ($mine | % { $_.Why -split ' / ' } | ? { $_ -notmatch $boiler -and $_ -notmatch '^회계 BL일' })
    $cs.Cells($crow, 1).NumberFormat = "@"; $cs.Cells($crow, 1).Value2 = [string]$inv
    if ($cust) { $cs.Cells($crow, 2).Value2 = [string]$cust }
    if ($a) {
      if ($a.Mj) { $cs.Cells($crow, 3).NumberFormat = "@"; $cs.Cells($crow, 3).Value2 = [string]$a.Mj }
      if ($a.Bl) { $cs.Cells($crow, 4).Value2 = [double]$a.Bl }; if ($a.Rate) { $cs.Cells($crow, 5).Value2 = [double]$a.Rate }
      if ($a.Fc) { $cs.Cells($crow, 6).Value2 = [double]$a.Fc }; if ($a.Krw) { $cs.Cells($crow, 7).Value2 = [double]$a.Krw }
    }
    $rl = if ($mr.Count) { (@($mr | sort) -join ', ') } else { RowList ($mine | % Row) }
    if ($rl) { $cs.Cells($crow, 8).NumberFormat = "@"; $cs.Cells($crow, 8).Value2 = [string]$rl } else { $cs.Cells($crow, 8).Value2 = 'MES 9월 출고 없음' }
    if ($mr.Count) { $cs.Cells($crow, 9).Value2 = [double](($mr | % { [double](Num $v[$_, $C['출고수량']]) } | measure -Sum).Sum); $cs.Cells($crow, 10).Value2 = [double](($mr | % { [double](Num $v[$_, $C['금액']]) } | measure -Sum).Sum) }
    $cs.Cells($crow, 11).Value2 = [string]($main -join ' / ')
    if ($extra) { $cs.Cells($crow, 12).Value2 = [string]($extra -join ' / ') }
    $crow++
  }
  $nCarryOv = $crow - 2
  foreach ($dc in $domCarry) {   # 국내: 일별마감 비고 '10월이월' → 다음 달 매출
    $cs.Cells($crow, 1).NumberFormat = "@"; $cs.Cells($crow, 1).Value2 = '-'
    $cs.Cells($crow, 2).Value2 = [string]$dc.Cust
    $cs.Cells($crow, 7).Value2 = [double]$dc.A
    $cs.Cells($crow, 8).NumberFormat = "@"; $cs.Cells($crow, 8).Value2 = [string]"일별마감 $mLab 시트 $($dc.Rows)행"
    $cs.Cells($crow, 9).Value2 = [double]$dc.Q; $cs.Cells($crow, 10).Value2 = [double]$dc.A
    $cs.Cells($crow, 11).Value2 = [string]"일별마감 $mLab 소계에 안 들어감($nLab 이월) → $nLab 매출"
    if ($dc.B) { $cs.Cells($crow, 12).Value2 = [string]"일별마감 비고: $($dc.B)" }
    $crow++
  }
  $cs.Cells($crow, 1).Value2 = '합계'; foreach ($cc in 7, 9, 10) { $cl = ($cs.Cells(1,$cc).Address($false,$false) -replace '\d',''); $cs.Cells($crow, $cc).Formula = "=SUM(${cl}2:${cl}$($crow-1))" }
  $cs.Range($cs.Cells($crow,1), $cs.Cells($crow,12)).Font.Bold = $true
  $cs.Range("D2:D$crow").NumberFormat = "yyyy-mm-dd"; $cs.Range("E2:E$crow").NumberFormat = "[<50]#,##0.0000;#,##0.00"; $cs.Range("F2:F$crow").NumberFormat = "#,##0.00"
  $cs.Range("G2:G$crow").NumberFormat = "#,##0"; $cs.Range("I2:J$crow").NumberFormat = "#,##0"
  $h1 = $cs.Range($cs.Cells(1,1), $cs.Cells(1,12)); $h1.Font.Bold = $true; $h1.Interior.Color = $HEAD
  $cs.Columns("A:J").AutoFit() | Out-Null; if ($cs.Columns("H").ColumnWidth -gt 28) { $cs.Columns("H").ColumnWidth = 28 }; $cs.Columns("H").WrapText = $true
  $cs.Columns("K").ColumnWidth = 40; $cs.Columns("L").ColumnWidth = 60; $cs.Columns("K:L").WrapText = $true
  # 맨 앞에 '구분'(해외/국내) 열, 머리글은 해외·국내 둘 다 맞게
  $cs.Columns(1).Insert() | Out-Null; $cs.Cells(1, 1).Value2 = '구분'
  for ($i = 2; $i -lt $crow; $i++) { $cs.Cells($i, 1).Value2 = [string]$(if ($i - 2 -lt $nCarryOv) { '해외' } else { '국내' }) }
  $cs.Cells(1, 8).Value2 = '원화금액(회계·일별마감)'; $cs.Cells(1, 9).Value2 = 'MES 행 / 일별마감 행'; $cs.Cells(1, 10).Value2 = '출고수량'; $cs.Cells(1, 11).Value2 = '출고 금액'
  $h1 = $cs.Range($cs.Cells(1,1), $cs.Cells(1,13)); $h1.Font.Bold = $true; $h1.Interior.Color = $HEAD; $cs.Columns(1).AutoFit() | Out-Null
  $cs.Activate() | Out-Null; $xl.ActiveWindow.SplitRow = 1; $xl.ActiveWindow.SplitColumn = 0; $xl.ActiveWindow.FreezePanes = $true
  $nCarry = $crow - 2

  # 14) 출하금액변동 시트: 출하 기준(MES 금액) vs 확정(회계) — 인보이스 1건 = 1줄. 국내는 전자세금계산서 자료가 오면 같은 형식으로 추가
  if ($have['출하금액변동']) { $have['출하금액변동'].Delete() }
  $ss = $wb.Worksheets.Add([Type]::Missing, $cs); $ss.Name = '출하금액변동'
  $sh0 = @('분류', '구분', '인보이스 번호', '고객사', '통화', 'MES 행', 'MES 출고수량', '출하기준 금액(MES)', '확정 금액(회계)', '원화 차액', '환율(회계)', '출하기준 외화', '확정 외화(회계)', '외화 차액', '차액률', '확정 출처', '사유')
  for ($j = 0; $j -lt $sh0.Count; $j++) { $ss.Cells(1, $j + 1).Value2 = [string]$sh0[$j] }
  $srow = 2
  foreach ($inv in @($shipInv.Keys | sort)) {
    if ($carry.Contains($inv) -or -not $grp.Contains($inv)) { continue }   # 이월 건은 이월매출 시트에서 관리
    $a = $acc[$inv]; if (-not $a -or -not $a.Rate) { continue }
    $mr = @($grp[$inv].Rows | ? { -not $repl[$_] })
    $mAmt = [double](($mr | % { [double](Num $v[$_, $C['금액']]) } | measure -Sum).Sum); $mQty = [double](($mr | % { [double](Num $v[$_, $C['출고수량']]) } | measure -Sum).Sum)
    $mFc = [math]::Round($mAmt / $a.Rate, 2); $dFc = [math]::Round($a.Fc - $mFc, 2)
    $cur = if ($a.Rate -lt 50) { 'JPY' } else { 'USD' }
    $vals0 = @('출하기준 금액변동', '해외', $inv, $a.Cust, $cur, ((@($mr | sort)) -join ', '), $mQty, $mAmt, $a.Krw, ($a.Krw - $mAmt), $a.Rate, $mFc, $a.Fc, $dFc, $(if ($mFc) { '{0:+0.00;-0.00}%' -f ($dFc / $mFc * 100) } else { '' }),
      $(if ($pdf[$inv]) { "회계 (= $($pdf[$inv].File) 매출내역서)" } else { '회계 (= 9월라인)' }), "MES 금액은 내부 고정환율(약 1,350원/USD) 기준 → BL일 환율·인보이스 금액과 차이 (정상)")
    for ($j = 0; $j -lt $vals0.Count; $j++) { $x0 = $vals0[$j]; if ($x0 -eq $null -or "$x0" -eq '') { continue }
      if ($x0 -is [double] -or $x0 -is [int]) { $ss.Cells($srow, $j + 1).Value2 = [double]$x0 } else { $ss.Cells($srow, $j + 1).NumberFormat = "@"; $ss.Cells($srow, $j + 1).Value2 = [string]$x0 } }
    $srow++
  }
  foreach ($vals0 in $domShip) {   # 국내: 세금계산서 공급가액 ≠ MES 금액
    for ($j = 0; $j -lt $vals0.Count; $j++) { $x0 = $vals0[$j]; if ($x0 -eq $null -or "$x0" -eq '') { continue }
      if ($x0 -is [double] -or $x0 -is [int]) { $ss.Cells($srow, $j + 1).Value2 = [double]$x0 } else { $ss.Cells($srow, $j + 1).NumberFormat = "@"; $ss.Cells($srow, $j + 1).Value2 = [string]$x0 } }
    $srow++
  }
  $ss.Cells($srow, 1).Value2 = '합계'; foreach ($cc in 7, 8, 9, 10) { $cl = ($ss.Cells(1,$cc).Address($false,$false) -replace '\d',''); $ss.Cells($srow, $cc).Formula = "=SUM(${cl}2:${cl}$($srow-1))" }
  $ss.Range($ss.Cells($srow,1), $ss.Cells($srow,17)).Font.Bold = $true
  $ss.Range("G2:J$srow").NumberFormat = "#,##0"; $ss.Range("K2:K$srow").NumberFormat = "[<50]#,##0.0000;#,##0.00"; $ss.Range("L2:N$srow").NumberFormat = "#,##0.00"
  $h1 = $ss.Range($ss.Cells(1,1), $ss.Cells(1,17)); $h1.Font.Bold = $true; $h1.Interior.Color = $HEAD
  $ss.Columns("A:P").AutoFit() | Out-Null; if ($ss.Columns("F").ColumnWidth -gt 28) { $ss.Columns("F").ColumnWidth = 28 }; $ss.Columns("F").WrapText = $true; $ss.Columns("Q").ColumnWidth = 50; $ss.Columns("Q").WrapText = $true
  $ss.Activate() | Out-Null; $xl.ActiveWindow.SplitRow = 1; $xl.ActiveWindow.SplitColumn = 0; $xl.ActiveWindow.FreezePanes = $true
  $nShipInv = $srow - 2

  # 12) 오류목록: 이월 건 제외, 동일 인보이스 = 1줄(사유는 항목별로 모아서). 인보이스 없는 행은 행별 1줄
  $rest = @($recs | ? { -not ($_.Inv -and $carry.Contains($_.Inv) -and $_.Cls -ne '에러') -and $_.Cls -ne '이월' })
  $groups = [ordered]@{}
  foreach ($x in $rest) { $key = if ($x.Inv) { $x.Inv } else { "$($x.Item)|$($x.Row)" }; if (-not $groups.Contains($key)) { $groups[$key] = New-Object System.Collections.ArrayList }; [void]$groups[$key].Add($x) }
  $rowsOut = @(foreach ($key in $groups.Keys) {
    $g = @($groups[$key]); $cls = if (@($g | ? { $_.Cls -eq '에러' }).Count) { '에러' } elseif (@($g | ? { $_.Cls -eq '알림' }).Count) { '알림' } else { '참고' }
    $its = U ($g | % Item)
    $vals = { param($p) $u0 = U ($g | % $p); if ($u0.Count -le 2) { $u0 -join ' / ' } else { "행별 상이($($u0.Count)건, mes9월 참고)" } }
    $why = @(foreach ($it in $its) { $ps = U ($g | ? { $_.Item -eq $it } | % { $_.Why -split ' / ' }); $(if ($its.Count -gt 1) { "[$it] " } else { '' }) + ($ps -join ' / ') }) -join "`n"
    ,@($cls, $g[0].Inv, (U ($g | % Cust) | select -First 1), ($its -join ', '), (RowList ($g | % Row)), [string]$g.Count, (& $vals 'Base'), (& $vals 'Cmp'), $(if (@(U ($g | % Src)).Count -le 2) { (U ($g | % Src)) -join ' / ' } else { (@(U ($g | % { $_.Src -replace '[\d,\.]+\s*÷', '÷' })) -join ' / ') + ' (행별)' }), $why)
  })
  $order = @{ '에러' = 0; '알림' = 1; '참고' = 2 }
  $rowsOut = @($rowsOut | sort { $order[$_[0]] }, { $_[3] }, { $_[1] })
  if ($have['오류목록']) { $have['오류목록'].Delete() }
  $es = $wb.Worksheets.Add([Type]::Missing, $ss); $es.Name = '오류목록'
  $eh = @('분류', '인보이스/Lot', '고객사', '항목', 'MES 행', '건수', '기준값(회계·MES)', '비교값', '비교 출처', '사유'); $ncE = $eh.Count
  $es.Range($es.Cells(1,1), $es.Cells($rowsOut.Count + 1, $ncE)).NumberFormat = "@"
  for ($j = 0; $j -lt $ncE; $j++) { $es.Cells(1, $j + 1).Value2 = [string]$eh[$j] }
  for ($ri = 0; $ri -lt $rowsOut.Count; $ri++) { for ($ci = 0; $ci -lt $ncE; $ci++) {
    $txt = [string]$rowsOut[$ri][$ci]
    if ($txt) { try { $es.Cells(2 + $ri, 1 + $ci).Value2 = [string]$rowsOut[$ri][$ci] } catch { throw "오류목록 쓰기 실패 R$($ri + 2) C$($ci + 1): $($_.Exception.Message)" } } } }
  $es.Range($es.Cells(1,6), $es.Cells($rowsOut.Count + 1, 6)).NumberFormat = "0"; for ($ri = 0; $ri -lt $rowsOut.Count; $ri++) { $es.Cells(2 + $ri, 6).Value2 = [double]$rowsOut[$ri][5] }
  $hdrE = $es.Range($es.Cells(1,1), $es.Cells(1,$ncE)); $hdrE.Font.Bold = $true; $hdrE.Interior.Color = $HEAD
  for ($i = 0; $i -lt $rowsOut.Count; $i++) { if ($rowsOut[$i][0] -eq '에러') { $es.Range($es.Cells($i + 2, 1), $es.Cells($i + 2, $ncE)).Interior.Color = [int]$PINK } elseif ($rowsOut[$i][0] -eq '알림') { $es.Range($es.Cells($i + 2, 1), $es.Cells($i + 2, $ncE)).Interior.Color = 10284031 } }   # ($c는 열 맵 $C와 겹치므로 사용 금지)
  $es.Columns("A:I").AutoFit() | Out-Null; $es.Columns("J").ColumnWidth = 80; $es.Columns("J").WrapText = $true
  foreach ($cl in 'C', 'D', 'E', 'G', 'H', 'I') { if ($es.Columns($cl).ColumnWidth -gt 28) { $es.Columns($cl).ColumnWidth = 28; $es.Columns($cl).WrapText = $true } }
  $es.Range($es.Cells(1,1), $es.Cells(1,$ncE)).AutoFilter() | Out-Null
  # 오른쪽 요약표
  $sc0 = $ncE + 2
  $sumRows = @(@('분류', '줄 수(인보이스)', '원래 행 수'), @('에러', [double]@($rowsOut | ? { $_[0] -eq '에러' }).Count, [double](@($rest | ? { $_.Cls -eq '에러' }).Count)), @('알림', [double]@($rowsOut | ? { $_[0] -eq '알림' }).Count, [double](@($rest | ? { $_.Cls -eq '알림' }).Count)), @('참고', [double]@($rowsOut | ? { $_[0] -eq '참고' }).Count, [double](@($rest | ? { $_.Cls -eq '참고' }).Count)), @('이월매출 시트', [double]$nCarry, [double](@($recs).Count - @($rest).Count)), @('출하금액변동 시트', [double]$nShipInv, [double]$nShip), @('전자세금계산서 금액 오류 시트', [double]$nTaxErr, [double]$nTaxErr))
  for ($i = 0; $i -lt $sumRows.Count; $i++) { for ($j = 0; $j -lt 3; $j++) { $val0 = $sumRows[$i][$j]; if ($val0 -is [double]) { $es.Cells(1 + $i, $sc0 + $j).Value2 = [double]$val0 } else { $es.Cells(1 + $i, $sc0 + $j).Value2 = [string]$val0 } } }
  $sum = $es.Range($es.Cells(1,$sc0), $es.Cells($sumRows.Count,$sc0 + 2)); $sum.Borders.LineStyle = 1; $sum.Columns.AutoFit() | Out-Null
  $es.Range($es.Cells(1,$sc0), $es.Cells(1,$sc0 + 2)).Font.Bold = $true; $es.Range($es.Cells(1,$sc0), $es.Cells(1,$sc0 + 2)).Interior.Color = $HEAD; $es.Range($es.Cells(2,$sc0), $es.Cells(2,$sc0 + 2)).Font.Color = $RED
  $es.Activate() | Out-Null; $xl.ActiveWindow.SplitRow = 1; $xl.ActiveWindow.SplitColumn = 0; $xl.ActiveWindow.FreezePanes = $true
  $errCnt = @($rowsOut | ? { $_[0] -eq '에러' }); $errList = @($errCnt | % { "  [$($_[3])] $($_[1]) MES행 $($_[4]) : $($_[9] -replace "`n", ' | ')" })
  "국내매출: 고객사 $($nDom)줄, 세금계산서 " + $(if ($taxF) { "$($taxF.Name) 매출 $($tax.Count)매" } else { "파일 없음(기초자료에 '9월세금계산서.xlsx' 넣으면 자동 비교)" })
  "Replacement 행 $($repl.Count)개 → 오류에서 제외 $($nRepl)건 / 불량대체 행 $($bulD.Count)개(MES $((@($bulD.Keys | sort)) -join ', ')행) → 참고 / 출하금액변동 $($nShipInv)건(원래 행 $($nShip))"
  "원자료 $(@($recs).Count)행 → 이월매출 $($nCarry)건, 오류목록 $($rowsOut.Count)줄 (에러 $($errCnt.Count) / 참고 $(@($rowsOut | ? { $_[0] -eq '참고' }).Count))"
  $ws.Activate() | Out-Null
  $xl.Calculation = -4105; $xl.ScreenUpdating = $true
  $wb.Save(); $wb.Close($false)

  "mes9월: 신코·코마츠(PDF) $($cnt.PDF)행 / 해외(9월라인) $($cnt.LINE)행 / 국내(출고일) $($cnt.DOM)행"
  "  BL날짜 '이월' $($cnt.Carry)행, BL날짜 빨간색 $($cnt.Red)행, 인보이스 빨간색 $($cnt.InvRed)행, 기초자료 환율·달러 사유 $($cmpRows.Count)건(인보이스)"
  "해외매출: $($row - 2)건 (PDF $nPdf / MES $nMes / 9월라인 참고 $nRef), 회계와 불일치·비교불가 $($nRed)건"
  "`n== 해외매출 에러내용(불일치)"; $gsNotes
  "`n== 오류목록 '에러' 분류"; $errList
} finally { $xl.Quit() }