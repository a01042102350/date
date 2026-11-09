$ErrorActionPreference = 'Stop'
$sp = Split-Path $MyInvocation.MyCommand.Path
$dir = 'C:\ai\매월마감\202610'; $ym = '2610'; $year = 2026
$files = @(@{ 고객 = 'SHINKO'; 파일 = '신코_지출품의서 마감.xlsx'; 시트 = "신코20$ym" }, @{ 고객 = 'JS TECH(코마츠)'; 파일 = '코마츠_지출품의서 마감.xlsx'; 시트 = "코마츠20$ym" })
$out = @()
$xl = New-Object -ComObject Excel.Application; $xl.DisplayAlerts = $false
try {
  foreach ($f in $files) { $wb = $xl.Workbooks.Open((Join-Path $dir $f.파일), 0, $true); $s = $wb.Worksheets.Item($f.시트)
    $v = $s.Range('A1:AB40').Value2; $hr = 0; $col = @{}
    for ($r = 1; $r -le 40 -and -not $hr; $r++) { for ($c = 1; $c -le 28; $c++) { if ("$($v[$r,$c])".Trim() -eq '인보이스') { $hr = $r } } }
    for ($c = 1; $c -le 28; $c++) { $h = ("$($v[$hr,$c])" -replace '\s+', '').ToLower(); if ($h) { $col[$h] = $c } }
    for ($r = $hr + 1; $r -le 40; $r++) { $inv = "$($v[$r,$col['인보이스']])".Trim(); if ($inv -match '합\s*계' -or -not $inv) { break }
      $bl = $v[$r, $col['bl날짜']]; $blT = $s.Cells($r, $col['bl날짜']).Text.Trim()
      $date = if ($bl -is [double] -and $bl -gt 30000) { [DateTime]::FromOADate($bl) } elseif ($blT -match '^(\d{1,2})/(\d{1,2})$') { Get-Date -Year $year -Month ([int]$matches[1]) -Day ([int]$matches[2]) } else { throw "BL날짜 못 읽음: $inv $blT" }
      $out += [pscustomobject]@{ 고객 = $f.고객; 파일 = $f.파일; 시트 = $f.시트; 행 = $r; 인보이스 = $inv; BL = $date.ToString('yyyy-MM-dd'); 환율 = [double]$v[$r, $col['환율']]; 엔화 = [double]$v[$r, $col['엔화금액']]; 원화 = [double]$v[$r, $col['계']] } }
    $wb.Close($false) }
} finally { $xl.Quit() }
$sm = @{}; $last = $null; Import-Csv "$sp\smbs.csv" -Encoding UTF8 | % { if ($_.JPY) { $last = [double]$_.JPY }; $sm[$_.날짜] = $last }
foreach ($x in $out) { $x | Add-Member 서울외국환중개 $sm[$x.BL]; $x | Add-Member 계산원화 ([Math]::Floor($x.엔화 * $x.환율 + 1e-7)) }
$out | Export-Csv "$sp\c_JP.csv" -NoTypeInformation -Encoding UTF8
$out | % { "{0} {1} {2}행 {3} BL {4} 환율 {5} (서울외국환중개 {6}) 엔 {7:N0} 원화 {8:N0} (엔×환율 내림 {9:N0})" -f $_.고객, $_.시트, $_.행, $_.인보이스, $_.BL, $_.환율, $_.서울외국환중개, $_.엔화, $_.원화, $_.계산원화 }
