$ErrorActionPreference = 'Stop'
$sp = Split-Path $MyInvocation.MyCommand.Path
$dir = 'C:\ai\매월마감\202610'
$fA = (Get-ChildItem $dir -Filter '*인보이스별 상세*').FullName; $fB = (Get-ChildItem $dir -Filter '*라인별 매출*').FullName
$fW = "$dir\2026년 수출 (2026.10)_10월반영.xlsx"
# 서울외국환중개 환율 (휴일 = 바로 앞 영업일)
$sm = Import-Csv "$sp\smbs.csv" -Encoding UTF8; $rate = @{}; $last = $null
foreach ($x in $sm) { if ($x.USD) { $last = [double]$x.USD }; $rate[$x.날짜] = [pscustomobject]@{ 환율 = $last; 고시 = [bool]$x.USD } }
function D($v) { if ($v -is [double] -and $v -gt 30000 -and $v -lt 60000) { [DateTime]::FromOADate($v).ToString('yyyy-MM-dd') } else { $null } }
$RED = 255
# 사용자 엑셀에 열어 둔 기초자료(읽기 전용) 닫기
try { $ux = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application'); foreach ($w in @($ux.Workbooks)) { if ($w.FullName -in $fA, $fB, $fW) { $w.Close($false) } } } catch { }
New-Item -ItemType Directory -Force "$dir\백업" | Out-Null
$ts = Get-Date -f yyyyMMdd_HHmmss
foreach ($f in $fA, $fB) { Copy-Item $f "$dir\백업\$([IO.Path]::GetFileNameWithoutExtension($f))_빨간칠전_$ts.xlsx" }
$res = @()
$xl = New-Object -ComObject Excel.Application; $xl.DisplayAlerts = $false
try {
  # ① 라인별 매출: 당월·이월 면장 시트 (BL일자 E, 환율 F)
  $wb = $xl.Workbooks.Open($fB)
  foreach ($sn in '26.10月당월_면장', '26.10月이월_면장') { $s = $wb.Worksheets.Item($sn); $v = $s.Range('A1:H100').Value2; $cu = ''
    for ($r = 1; $r -le 100; $r++) { if ("$($v[$r,1])" -match '합\s*계') { break }; if ("$($v[$r,1])".Trim()) { $cu = ("$($v[$r,1])" -replace '\s+', ' ').Trim() }
      $inv = "$($v[$r,3])".Trim(); $bl = D $v[$r,5]; $g = $v[$r,6]; if (-not $inv -or -not $bl -or $null -eq $g -or $cu -like 'BERGEN*') { continue }
      if (-not $rate[$bl]) { $res += [pscustomobject]@{ 파일 = '범위 밖'; 시트 = "$r"; 행 = $r; 인보이스 = $inv; BL = $bl; 적힌환율 = $g; 서울외국환중개 = $null; 맞음 = $true }; continue }; $want = $rate[$bl].환율; $res += [pscustomobject]@{ 파일 = '라인별 매출'; 시트 = $sn; 행 = $r; 인보이스 = $inv; BL = $bl; 적힌환율 = [double]$g; 서울외국환중개 = $want; 맞음 = ([Math]::Abs([double]$g - $want) -lt 0.001) }
      if ([Math]::Abs([double]$g - $want) -ge 0.001) { $s.Cells($r, 6).Interior.Color = $RED } } }
  $wb.Save(); $wb.Close($false)
  # ② 인보이스별 상세: 각 시트 '인보이스' 머리글 아래 (bl 날짜, 환율 칸)
  $wb = $xl.Workbooks.Open($fA)
  foreach ($s in $wb.Worksheets) { $v = $s.Range('O1:Z60').Value2; $hr = 0
    for ($r = 1; $r -le 60 -and -not $hr; $r++) { for ($c = 1; $c -le 12; $c++) { if ("$($v[$r,$c])".Trim() -eq '인보이스') { $hr = $r; $ic = $c } } }
    if (-not $hr) { continue }; $cB = 0; $cR = 0; for ($c = 1; $c -le 12; $c++) { $h = ("$($v[$hr,$c])" -replace '\s+', '').ToLower(); if ($h -eq 'bl날짜') { $cB = $c }; if ($h -eq '환율') { $cR = $c } }
    if (-not $cB -or -not $cR) { continue }
    for ($r = $hr + 1; $r -le 60; $r++) { $inv = "$($v[$r,$ic])".Trim(); $bl = D $v[$r,$cB]; $g = $v[$r,$cR]; if (-not $inv -or -not $bl -or -not ($g -is [double])) { continue }
      if (-not $rate[$bl]) { $res += [pscustomobject]@{ 파일 = '범위 밖'; 시트 = "$r"; 행 = $r; 인보이스 = $inv; BL = $bl; 적힌환율 = $g; 서울외국환중개 = $null; 맞음 = $true }; continue }; $want = $rate[$bl].환율; $ok = [Math]::Abs($g - $want) -lt 0.001
      $res += [pscustomobject]@{ 파일 = '인보이스별 상세'; 시트 = $s.Name; 행 = $r; 인보이스 = $inv; BL = $bl; 적힌환율 = $g; 서울외국환중개 = $want; 맞음 = $ok }
      if (-not $ok) { $s.Cells($r, 14 + $cR).Interior.Color = $RED } } }
  $wb.Save(); $wb.Close($false)
  # ③ 작업 파일 10월 시트 (BL일자 F, 환율 G) — 칠하기만, 값은 안 바꿈
  $wb = $xl.Workbooks.Open($fW); $s = $wb.Worksheets.Item('26.10(NEW)')
  for ($r = 32; $r -le 140; $r++) { $a = $s.Cells($r,1).Text; if ($a -like '총*계*') { break }; $bl = D $s.Cells($r,6).Value2; $g = $s.Cells($r,7).Value2
    if (-not $bl -or -not ($g -is [double]) -or $a -like 'BERGEN*' -or $a -like 'SHINKO*' -or $a -like 'JS TECH*') { continue }
    if (-not $rate[$bl]) { $res += [pscustomobject]@{ 파일 = '범위 밖'; 시트 = "$r"; 행 = $r; 인보이스 = $inv; BL = $bl; 적힌환율 = $g; 서울외국환중개 = $null; 맞음 = $true }; continue }; $want = $rate[$bl].환율; $ok = [Math]::Abs($g - $want) -lt 0.001
    $res += [pscustomobject]@{ 파일 = '10월반영(작업 파일)'; 시트 = '26.10(NEW)'; 행 = $r; 인보이스 = $s.Cells($r,3).Text; BL = $bl; 적힌환율 = $g; 서울외국환중개 = $want; 맞음 = $ok }
    if (-not $ok) { $s.Cells($r, 7).Interior.Color = $RED } }
  $wb.Save(); $wb.Close($false)
} finally { $xl.Quit() }
$res | Export-Csv "$sp\redmark.csv" -NoTypeInformation -Encoding UTF8
$res | Group-Object 파일 | % { "$($_.Name): 비교 $($_.Count)칸, 틀림 $(@($_.Group | ? { -not $_.맞음 }).Count)칸" }
"--- 틀린 칸"; $res | ? { -not $_.맞음 } | % { "$($_.파일) / $($_.시트) $($_.행)행 $($_.인보이스) BL $($_.BL): 적힌 $($_.적힌환율) → 서울외국환중개 $($_.서울외국환중개)" }
