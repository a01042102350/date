# 그쪽(MES매칭 통합) 거래처별 월별 원장 대사 ↔ 내 1~10월 결과 대조 → cmp.csv
$ErrorActionPreference = 'Stop'
$W = 'C:\Users\user\AppData\Local\Temp\claude\d3df15a30ffc05d5\ac480345-177e-4ecb-8db7-e819891326b4\scratchpad\mesmatch'
function N([string]$s) { $x = $s -replace '[^\d\.\-]', ''; if ($x -match '^-?\d+(\.\d+)?$') { [decimal]$x } else { [decimal]0 } }
$ALIAS = @{
  '귀뚜라미범양냉방' = '귀뚜라미범양냉방'; '귀뚜라미범양' = '귀뚜라미범양냉방'; '부경테크윈' = '부경테크원'; 'ATLASKOREA(국·해외)' = '아트라스콥코'; 'DN솔루션즈' = '디엔솔루션즈'
  'STX엔진제1사업장' = 'STX엔진'; 'STX엔진민수' = 'STX엔진'; 'STX엔진방산' = 'STX엔진'; '명화' = '명화공업'; '엘에스엠트론' = 'LS엠트론'; '현대로템방산' = '현대로템'; '현대로템철차' = '현대로템'; '현대트랜시스중기' = '현대트랜시스'
  'ALLISON' = 'ALLISON'; 'BERGENENG' = 'BERGEN ENGINE'; 'BERGENENGINE' = 'BERGEN ENGINE'; 'JSTECH(코마츠)' = '코마츠(JS TECH)'; '코마츠(JSTECH)' = '코마츠(JS TECH)'; 'MERCURYC' = 'MERCURY'; 'MERCURYCHINA' = 'MERCURY'; 'MERCURYU' = 'MERCURY'; 'MERCURY' = 'MERCURY'
  'SHINKO' = '신코'; 'SINKOINDLTD' = '신코'; 'EMERSON' = 'EMERSON'; 'POLARIS' = 'POLARIS'; 'DOMETIC' = 'DOMETIC'; 'AMAT' = 'AMAT'; 'SAM' = 'SAM'; 'MEPPI' = 'MEPPI'
}
function Key([string]$nm) { $k = ($nm -replace '㈜|\(주\)|주식회사|\s', '').ToUpper(); if ($ALIAS.ContainsKey($k)) { return $ALIAS[$k] }; return $k }
# 그쪽
$L = Get-Content (Join-Path $W 's1.tsv') -Encoding UTF8
$their = @{}; $dept = @{}
foreach ($line in $L | Select-Object -Skip 5) {
  $c = $line -split "`t"; if ($c[0] -notin '국내영업', '해외영업' -or $c[1] -like '*소계') { continue }
  $k = Key $c[1]; $dept[$k] = $c[0]
  for ($m = 1; $m -le 10; $m++) { $b = 2 + ($m - 1) * 5; $their["$k|$m"] = [pscustomobject]@{ 제품 = (N $c[$b]); 기타 = (N $c[$b + 1]); 합계 = (N $c[$b + 2]); 원장 = (N $c[$b + 3]); 차이 = (N $c[$b + 4]) } }
}
# 내 것
$mine = Import-Csv (Join-Path $W 'mine_rows.csv') -Encoding UTF8
$agg = @{}; $names = @{}
foreach ($r in $mine) {
  if ($r.출처 -eq '이월매출') { continue }
  $k = Key $r.고객사; $id = "$k|$($r.월)"
  if (-not $agg.ContainsKey($id)) { $agg[$id] = [pscustomobject]@{ 해외 = [decimal]0; 국내 = [decimal]0; 계산서 = [decimal]0; 계산서있음 = $false; 면장 = 0 } }
  $a = $agg[$id]
  if ($r.출처 -eq '해외매출') { $a.해외 += [decimal]$r.금액; $a.면장++ } else { $a.국내 += [decimal]$r.금액; if ($r.세금계산서 -ne '') { $a.계산서 += [decimal]$r.세금계산서; $a.계산서있음 = $true } }
  if (-not $names.ContainsKey($k)) { $names[$k] = New-Object Collections.Generic.HashSet[string] }; [void]$names[$k].Add($r.고객사)
}
$keys = @($their.Keys + $agg.Keys | ForEach-Object { ($_ -split '\|')[0] } | Select-Object -Unique)
$res = foreach ($k in $keys) {
  for ($m = 1; $m -le 10; $m++) {
    $t = $their["$k|$m"]; $a = $agg["$k|$m"]
    $tw = if ($t) { $t.원장 } else { [decimal]0 }; $tt = if ($t) { $t.합계 } else { [decimal]0 }
    $my = if ($a) { $a.해외 + $a.국내 } else { [decimal]0 }
    if ($tw -eq 0 -and $tt -eq 0 -and $my -eq 0) { continue }
    [pscustomobject]@{
      부서 = $(if ($dept.ContainsKey($k)) { $dept[$k] } else { '(그쪽에 없음)' }); 거래처 = $k; 월 = $m
      내이름 = $(if ($names.ContainsKey($k)) { ($names[$k] -join ', ') } else { '' })
      그쪽제품 = $(if ($t) { $t.제품 } else { 0 }); 그쪽기타 = $(if ($t) { $t.기타 } else { 0 }); 그쪽합계 = $tt; 원장 = $tw
      내구분 = $(if ($a -and $a.해외 -ne 0) { '해외매출' } elseif ($a) { '국내매출' } else { '' })
      내금액 = $my; 내계산서 = $(if ($a -and $a.계산서있음) { $a.계산서 } else { '' })
      차이_내금액_원장 = $my - $tw; 차이_내금액_그쪽제품 = $my - $(if ($t) { $t.제품 } else { 0 })
      차이_계산서_원장 = $(if ($a -and $a.계산서있음) { $a.계산서 - $tw } else { '' })
    }
  }
}
$res | Sort-Object 부서, 거래처, 월 | Export-Csv (Join-Path $W 'cmp.csv') -NoTypeInformation -Encoding UTF8
"대조 줄: $(@($res).Count)"
'--- 이름 짝(내 이름 → 그쪽 거래처)'; $names.Keys | Sort-Object | ForEach-Object { "  $_ ← $($names[$_] -join ', ')$(if (-not $dept.ContainsKey($_)) { '   ※그쪽에 없음' })" }
'--- 그쪽에만 있는 거래처'; $dept.Keys | Where-Object { -not $names.ContainsKey($_) } | ForEach-Object { "  $_ ($($dept[$_])) 1~10월 원장 합계 " + ('{0:N0}' -f (($res | Where-Object 거래처 -eq $_ | Measure-Object 원장 -Sum).Sum)) }
'--- 월별 합계 (원장 / 그쪽 합계 / 내 금액 / 내-원장)'
for ($m = 1; $m -le 10; $m++) { $g = @($res | Where-Object 월 -eq $m); '{0,2}월  원장 {1,16:N0}  그쪽합계 {2,16:N0}  내금액 {3,16:N0}  내-원장 {4,15:N0}' -f $m, ($g | Measure-Object 원장 -Sum).Sum, ($g | Measure-Object 그쪽합계 -Sum).Sum, ($g | Measure-Object 내금액 -Sum).Sum, ($g | Measure-Object 차이_내금액_원장 -Sum).Sum }
$okN = @($res | Where-Object { [math]::Abs($_.차이_내금액_원장) -le 10 }).Count; "내 금액 = 원장(±10원): $okN / $(@($res).Count) 칸"
