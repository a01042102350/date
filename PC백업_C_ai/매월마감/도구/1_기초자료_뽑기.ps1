$ErrorActionPreference = 'Stop'
$dir = 'C:\ai\매월마감\202610'
$fA = Get-ChildItem $dir -Filter '*인보이스별 상세*' | Select -Expand FullName
$fB = Get-ChildItem $dir -Filter '*라인별 매출*' | Select -Expand FullName
$fX = Get-ChildItem $dir -Filter '2026년 수출*' | Select -Expand FullName
$out = Split-Path $MyInvocation.MyCommand.Path
function D($v) { if ($v -is [double] -and $v -gt 30000 -and $v -lt 60000) { [DateTime]::FromOADate($v).ToString('yyyy-MM-dd') } else { "$v" } }
function N($v) { if ($null -eq $v -or "$v" -eq '') { $null } else { [double]$v } }
$xl = New-Object -ComObject Excel.Application; $xl.DisplayAlerts = $false
try {
  # ── 라인별 파일 ──
  $wb = $xl.Workbooks.Open($fB, 0, $true)
  $B = New-Object Collections.Generic.List[object]
  foreach ($sn in '26.10月당월_면장', '26.10月이월_면장') { $s = $wb.Worksheets.Item($sn); $v = $s.Range('A1:K40').Value2; $cur = ''
    $v = $s.Range('A1:K100').Value2
    for ($r = 1; $r -le 100; $r++) { $a = "$($v[$r,1])".Trim(); if ($a -match '합\s*계') { break }
      if ($a) { $cur = ($a -replace '\s+', ' ') }
      $inv = "$($v[$r,3])".Trim(); if (-not $inv -or $inv -match 'INVOICE') { continue }
      $B.Add([pscustomobject]@{ 구분 = $(if ($sn -like '*이월*') { '이월' } else { '당월' }); 거래처 = $cur; 공장 = "$($v[$r,2])"; 인보이스 = $inv; 면장 = "$($v[$r,4])".Trim(); BL = (D $v[$r,5]); 환율 = (N $v[$r,6]); 외화 = (N $v[$r,7]); 원화 = (N $v[$r,8]); 비고 = ("$($v[$r,11])" -replace '\s+', ' ')
        초록 = ($s.Cells($r,3).Interior.Color -eq 15523812); 환율수식 = $s.Cells($r,6).HasFormula; 행 = $r }) } }
  $h = $wb.Worksheets.Item('환율'); $hv = $h.Range('A1:B3040').Value2; $RATE = [ordered]@{}
  for ($r = 3; $r -le 3040; $r++) { $k = D $hv[$r,1]; if ($k -ge '2026-09-25' -and $k -le '2026-11-30') { $RATE[$k] = $hv[$r,2] } }
  $wb.Close($false)
  # ── 인보이스별 상세 ──
  $wb = $xl.Workbooks.Open($fA, 0, $true); $A = New-Object Collections.Generic.List[object]
  foreach ($s in $wb.Worksheets) { $v = $s.Range('O1:Z60').Value2; $hr = 0
    for ($r = 1; $r -le 60; $r++) { for ($c = 1; $c -le 12; $c++) { if ("$($v[$r,$c])".Trim() -eq '인보이스') { $hr = $r; $ic = $c } } ; if ($hr) { break } }
    if (-not $hr) { "헤더 없음: $($s.Name)"; continue }
    $col = @{}; for ($c = 1; $c -le 12; $c++) { $col[("$($v[$hr,$c])" -replace '\s+', '').ToLower()] = $c }
    $cBL = @($col.Keys | ? { $_ -match 'bl날짜|발행일' })[0]; $cR = $col['환율']; $cU = $col['달러']; $cW = $col['원화']; $cC = $col['거래처']
    for ($r = $hr + 1; $r -le 60; $r++) { $inv = ("$($v[$r,$ic])" -replace '\s+', ' ').Trim(); if (-not $inv) { if ("$($v[$r,$cW])") { } ; continue }
      $A.Add([pscustomobject]@{ 시트 = $s.Name; 거래처 = $(if ($cC) { "$($v[$r,$cC])".Trim() } else { $s.Name }); 인보이스 = $inv; BL = (D $v[$r, $col[$cBL]]); 환율 = $(if ($cR) { N $v[$r,$cR] }); 외화 = $(if ($cU) { N $v[$r,$cU] }); 원화 = (N $v[$r,$cW]); 행 = $r + 0 }) } }
  $wb.Close($false)
  # ── 수출 파일 9월 시트 이월(회색) 줄 ──
  $wb = $xl.Workbooks.Open($fX, 0, $true); $s = $wb.Worksheets.Item('26.09(NEW)'); $v = $s.Range('A1:K140').Value2; $S9 = @()
  for ($r = 31; $r -le 133; $r++) { if ($null -ne $v[$r,11] -and "$($v[$r,11])".Trim() -ne '') { $S9 += [pscustomobject]@{ 거래처 = "$($v[$r,1])".Trim(); 인보이스 = "$($v[$r,3])".Trim(); 면장 = "$($v[$r,4])".Trim(); 외화 = (N $v[$r,8]); 원화 = (N $v[$r,11]) } } }
  $wb.Close($false)
} finally { $xl.Quit() }
$A | Export-Csv "$out\c_A.csv" -NoTypeInformation -Encoding UTF8
$B | Export-Csv "$out\c_B.csv" -NoTypeInformation -Encoding UTF8
$S9 | Export-Csv "$out\c_S9.csv" -NoTypeInformation -Encoding UTF8
$RATE.GetEnumerator() | % { [pscustomobject]@{ 날짜 = $_.Key; 환율 = $_.Value } } | Export-Csv "$out\c_FX.csv" -NoTypeInformation -Encoding UTF8
"상세 $($A.Count)줄, 라인별 $($B.Count)줄, 9월 이월 $($S9.Count)줄, 환율 $($RATE.Count)일"
