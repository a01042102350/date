# 수출 마감 자동 — "2026년 수출 (YYYY.MM)_M월반영.xlsx" 만들기
# 사용: powershell -ExecutionPolicy Bypass -File 수출마감_자동.ps1 -YM 202611 [-CheckOnly] [-Folder 다른폴더]
# 엑셀 버튼: C:\ai\매월마감\매월마감_실행.xlsm
# 순서: ① 파일 확인 ② 서울외국환중개 환율 받기 ③ 기초자료 읽기 ④ 당월 시트 만들기(사용자 입력 되살림)
#       ⑤ 오른쪽 확인표 ⑥ 기초자료 검증 시트 ⑦ 환율 틀린 칸 빨간칠(기초자료·결과) ⑧ 저장
param([Parameter(Mandatory = $true)][string]$YM, [switch]$CheckOnly, [string]$Folder)
$ErrorActionPreference = 'Stop'
$root = 'C:\ai\매월마감'
$toolDir = Split-Path $MyInvocation.MyCommand.Path
$tplFile = Join-Path $toolDir '서식틀_수출상세.xlsx'
$rateCache = Join-Path $toolDir '서울외국환중개_환율.csv'

if ($YM -notmatch '^20\d{2}(0[1-9]|1[0-2])$') { Write-Host "마감월은 6자리 숫자로 넣어 주세요 (예: 202611). 받은 값: $YM" -ForegroundColor Red; exit 2 }
$Y = [int]$YM.Substring(0, 4); $M = [int]$YM.Substring(4, 2)
$cur = [DateTime]::new($Y, $M, 1); $prv = $cur.AddMonths(-1); $nxt = $cur.AddMonths(1)
$MM = $cur.ToString('MM'); $tag = $cur.ToString('yy') + '.' + $MM; $ptag = $prv.ToString('yy') + '.' + $prv.ToString('MM')
$ym4 = $cur.ToString('yyMM')
$newSheet = "$($tag)(NEW)"; $prevSheet = "$($ptag)(NEW)"; $verName = "$($M)월 기초자료 검증"
$dir = if ($Folder) { $Folder } else { Join-Path $root $YM }
$origName = "$($Y)년 수출 ($($Y).$($MM)).xlsx"
$newF = Join-Path $dir "$($Y)년 수출 ($($Y).$($MM))_$($M)월반영.xlsx"
$ts = Get-Date -f yyyyMMdd_HHmmss

New-Item -ItemType Directory -Force $dir | Out-Null
$work = Join-Path $dir '작업'; $bak = Join-Path $dir '백업'
New-Item -ItemType Directory -Force $work, $bak | Out-Null
$logF = Join-Path $work '실행기록.txt'; $doneF = Join-Path $work '결과.txt'
$U8 = New-Object Text.UTF8Encoding($true)
[IO.File]::WriteAllText($logF, "$($Y)년 $($M)월 수출 마감 실행기록 ($(Get-Date -f 'yyyy-MM-dd HH:mm:ss'))`r`n", $U8)
if (Test-Path $doneF) { Remove-Item $doneF }
function Log([string]$m, [string]$c = 'Gray') { Write-Host $m -ForegroundColor $c; [IO.File]::AppendAllText($logF, "$m`r`n", $U8) }
function Done([string]$st, [string]$msg) { [IO.File]::WriteAllText($doneF, "$st`r`n$msg", $U8) }
$NOTE = New-Object Collections.Generic.List[object]
function Note($lvl, $what, $amt, $text, $todo) { $NOTE.Add(@($lvl, $what, $amt, $text, $todo)) | Out-Null
  Log "  [$lvl] $what : $text" $(switch ($lvl) { '에러' { 'Red' } '알림' { 'Yellow' } default { 'DarkGray' } }) }
$BAD = New-Object Collections.Generic.List[string]
function Bad([string]$m) { $BAD.Add($m) | Out-Null; Log "  ✖ $m" 'Red' }

function SetV($cell, $v) { $a = New-Object 'object[,]' 1, 1; if ($null -ne $v) { $v = $v.psobject.BaseObject }; $a[0, 0] = $v; $cell.Value2 = $a }
function Put($s, $r, $vals, $c0 = 1) { $a = New-Object 'object[,]' 1, $vals.Count
  for ($i = 0; $i -lt $vals.Count; $i++) { $x = $vals[$i]; if ($null -ne $x) { $x = $x.psobject.BaseObject }; $a[0, $i] = $x }
  $s.Range($s.Cells($r, $c0), $s.Cells($r, $c0 + $vals.Count - 1)).Value2 = $a }
function D($v) { if ($v -is [double] -and $v -gt 30000 -and $v -lt 60000) { [DateTime]::FromOADate($v).ToString('yyyy-MM-dd') } else { "$v".Trim() } }
function N($v) { if ($null -eq $v -or "$v".Trim() -eq '') { return $null }; if ($v -is [double]) { return $v }; $d = 0.0
  if ([double]::TryParse(("$v" -replace '[,₩\$¥\s]', ''), [ref]$d)) { $d } else { $null } }
function Locked($f) { if (-not $f -or -not (Test-Path -LiteralPath $f)) { return $false }; try { $h = [IO.File]::Open($f, 'Open', 'ReadWrite', 'None'); $h.Close(); $false } catch { $true } }
function SheetOf($wb, $name) { $k = $name -replace '\s', ''; foreach ($s in $wb.Worksheets) { if (($s.Name -replace '\s', '') -eq $k) { return $s } }; $null }
function FindRow($sh, $col, $pat, $from, $to) { for ($q = $from; $q -le $to; $q++) { if ($sh.Cells($q, $col).Text -match $pat) { return $q } }; 0 }
function Pick($pat, $what, [switch]$Opt) {
  $c = @(Get-ChildItem -LiteralPath $dir -File | ? { $_.Name -like $pat -and $_.Name -notlike '~$*' -and $_.Extension -in '.xlsx', '.xls', '.xlsm' -and $_.Name -notlike '*반영*' } | Sort-Object LastWriteTime -Descending)
  if ($c.Count -eq 0) { if ($Opt) { Note '알림' $what $null "폴더에 '$what' 파일이 없습니다." '있어야 하는 달이면 폴더에 넣고 다시 눌러 주세요.' } else { Bad "'$what' 파일이 없습니다. 이름에 '$($pat.Trim('*'))'가 들어간 엑셀 파일을 폴더에 넣어 주세요." }; return $null }
  if ($c.Count -gt 1) { Note '알림' "$what 파일이 여러 개" $null ("가장 최근에 저장된 '" + $c[0].Name + "'을(를) 썼습니다. 안 쓴 파일: " + (($c | Select-Object -Skip 1 | % Name) -join ', ')) '다른 파일을 써야 하면 안 쓰는 파일을 백업 폴더로 옮기고 다시 눌러 주세요.' }
  Log "  ✔ $what : $($c[0].Name)" 'Green'; $c[0].FullName }

$xl = $null; $code = 1
try {
  Log "══ $($Y)년 $($M)월 수출 마감 ══  폴더: $dir" 'Cyan'
  # ───────── ① 파일 확인 ─────────
  Log '① 파일 확인' 'Cyan'
  if (-not (Test-Path $tplFile)) { Bad "서식 틀 파일이 없습니다: $tplFile" }
  $fA = Pick '*인보이스별 상세*' '인보이스별 상세(해외영업팀)'
  $fB = Pick '*라인별 매출*' '라인별 매출(해외영업팀)'
  $fS = Pick '신코*지출품의서*' '신코 지출품의서 마감' -Opt
  $fK = Pick '코마츠*지출품의서*' '코마츠 지출품의서 마감' -Opt
  foreach ($f in $fA, $fB) { if ($f -and ([IO.Path]::GetFileName($f) -notmatch "(?<!\d)$($M)월")) { Bad "'$([IO.Path]::GetFileName($f))' 파일 이름에 '$($M)월'이 없습니다. 다른 달 파일인지 확인해 주세요." } }
  $orig = Join-Path $dir $origName; $origNote = $null
  if (-not (Test-Path -LiteralPath $orig)) {
    $alt = @(Get-ChildItem -LiteralPath $dir -File | ? { $_.Name -like "*수출 ($($Y).$($MM))*.xls*" -and $_.Name -notlike '*반영*' -and $_.Name -notlike '~$*' } | Sort-Object LastWriteTime -Descending)
    $pdir = Join-Path $root $prv.ToString('yyyyMM')
    $pres = if (Test-Path $pdir) { @(Get-ChildItem -LiteralPath $pdir -File | ? { $_.Name -like "*수출*_$($prv.Month)월반영.xlsx" -and $_.Name -notlike '~$*' }) } else { @() }
    if ($alt) { $orig = $alt[0].FullName }
    elseif ($pres) { $orig = $pres[0].FullName; $origNote = "폴더에 '$origName' 파일이 없어 지난달 결과 파일 '$($pres[0].Name)'을(를) 바탕으로 만들었습니다." }
    else { Bad "수출 파일 '$origName'이(가) 없습니다. 회계 수출 파일을 이 이름으로 폴더에 넣어 주세요."; $orig = $null } }
  if ($orig) { Log "  ✔ 수출 파일(바탕) : $([IO.Path]::GetFileName($orig))" 'Green'; if ($origNote) { Note '알림' '바탕 수출 파일' $null $origNote '이번 달 회계 수출 파일을 따로 받았다면 폴더에 넣고 다시 눌러 주세요.' } }
  if (Locked $newF) { Bad "결과 파일 '$([IO.Path]::GetFileName($newF))'이(가) 엑셀에서 열려 있습니다. 저장하고 닫은 뒤 다시 눌러 주세요." }
  foreach ($f in $fA, $fB) { if (Locked $f) { Bad "'$([IO.Path]::GetFileName($f))'이(가) 엑셀에서 열려 있습니다. 닫고 다시 눌러 주세요(환율 틀린 칸에 빨간칠을 해야 해서)." } }

  $xl = New-Object -ComObject Excel.Application; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.AskToUpdateLinks = $false
  # 라인별 매출
  $LINE = New-Object Collections.Generic.List[object]; $LRATE = [ordered]@{}
  if ($fB) { $wbB = $xl.Workbooks.Open($fB, 0, $true)
    foreach ($sn in "$($tag)月당월_면장", "$($tag)月이월_면장") { $s = SheetOf $wbB $sn
      if (-not $s) { if ($sn -like '*당월*') { Bad "라인별 매출 파일에 '$sn' 시트가 없습니다. (있는 시트: $((@($wbB.Worksheets) | % Name) -join ', '))" } else { Note '알림' '이월 시트 없음' $null "라인별 매출 파일에 '$sn' 시트가 없어 다음 달 이월은 넣지 않았습니다." '이월이 있으면 확인해 주세요.' }; continue }
      $v = $s.Range('A1:K400').Value2; $cu = ''
      for ($r = 1; $r -le 400; $r++) { $a1 = "$($v[$r,1])".Trim(); if ($a1 -match '합\s*계') { break }
        if ($a1) { $cu = ($a1 -replace '\s+', ' ') }
        $inv = "$($v[$r,3])".Trim(); if (-not $inv -or $inv -match 'INVOICE' -or $inv -match '매출마감일') { continue }
        $LINE.Add([pscustomobject]@{ 구분 = $(if ($sn -like '*이월*') { '이월' } else { '당월' }); 거래처 = $cu; 공장 = "$($v[$r,2])".Trim(); 인보이스 = $inv; 면장 = "$($v[$r,4])".Trim(); BL = (D $v[$r, 5]); 환율 = (N $v[$r, 6]); 외화 = (N $v[$r, 7]); 원화 = (N $v[$r, 8]); 비고 = ("$($v[$r,11])" -replace '\s+', ' ').Trim(); 시트 = $s.Name; 행 = $r }) | Out-Null } }
    $h = SheetOf $wbB '환율'
    if ($h) { $hv = $h.Range('A1:B5000').Value2; $lo = [DateTime]::new($prv.Year, $prv.Month, 20).ToString('yyyy-MM-dd'); $hi = $nxt.AddMonths(1).AddDays(-1).ToString('yyyy-MM-dd')
      for ($r = 1; $r -le 5000; $r++) { $k = D $hv[$r, 1]; if ($k -match '^\d{4}-' -and $k -ge $lo -and $k -le $hi -and (N $hv[$r, 2])) { $LRATE[$k] = N $hv[$r, 2] } } }
    $wbB.Close($false)
    Log "  라인별 매출: 당월 $(@($LINE | ? 구분 -eq '당월').Count)줄, 이월 $(@($LINE | ? 구분 -eq '이월').Count)줄, 환율 시트 $($LRATE.Count)일" }
  # 인보이스별 상세
  $DET = New-Object Collections.Generic.List[object]
  if ($fA) { $wbA = $xl.Workbooks.Open($fA, 0, $true)
    foreach ($s in $wbA.Worksheets) { $v = $s.Range('O1:Z80').Value2; $hr = 0; $ic = 0
      for ($r = 1; $r -le 80 -and -not $hr; $r++) { for ($c = 1; $c -le 12; $c++) { if ("$($v[$r,$c])".Trim() -eq '인보이스') { $hr = $r; $ic = $c; break } } }
      if (-not $hr) { continue }
      $col = @{}; for ($c = 1; $c -le 12; $c++) { $hh = ("$($v[$hr,$c])" -replace '\s+', '').ToLower(); if ($hh -and -not $col.ContainsKey($hh)) { $col[$hh] = $c } }
      $kBL = @($col.Keys | ? { $_ -match 'bl날짜|발행일' })[0]; $cBL = if ($kBL) { $col[$kBL] } else { 0 }
      $cR = [int]$col['환율']; $cU = [int]$col['달러']; $cW = [int]$col['원화']; $cC = [int]$col['거래처']
      if (-not $cW) { continue }
      for ($r = $hr + 1; $r -le 80; $r++) { $inv = ("$($v[$r,$ic])" -replace '\s+', ' ').Trim(); if (-not $inv) { continue }
        $DET.Add([pscustomobject]@{ 시트 = $s.Name; 거래처 = $(if ($cC) { "$($v[$r,$cC])".Trim() } else { $s.Name }); 인보이스 = $inv; BL = $(if ($cBL) { D $v[$r, $cBL] }); 환율 = $(if ($cR) { N $v[$r, $cR] }); 외화 = $(if ($cU) { N $v[$r, $cU] }); 원화 = (N $v[$r, $cW]); 행 = $r; 환율열 = $(if ($cR) { 14 + $cR } else { 0 }) }) | Out-Null } }
    $wbA.Close($false)
    if ($DET.Count -eq 0) { Bad "인보이스별 상세 파일에서 '인보이스' 머리글이 있는 표를 찾지 못했습니다." } else { Log "  인보이스별 상세: $($DET.Count)줄 ($((@($DET | % 시트 | Select-Object -Unique)) -join ', '))" } }
  # 신코·코마츠 (그 달 시트만)
  $JPL = New-Object Collections.Generic.List[object]
  foreach ($jf in @(@{ 고객 = 'SHINKO'; 파일 = $fS; 시트 = "신코20$ym4" }, @{ 고객 = 'JS TECH(코마츠)'; 파일 = $fK; 시트 = "코마츠20$ym4" })) {
    if (-not $jf.파일) { continue }
    $wbJ = $xl.Workbooks.Open($jf.파일, 0, $true); $s = SheetOf $wbJ $jf.시트
    if (-not $s) { Note '알림' "$($jf.고객) 이번 달 시트 없음" $null "'$([IO.Path]::GetFileName($jf.파일))'에 '$($jf.시트)' 시트가 없어 넣지 않았습니다(빈 줄만 둠)." '이번 달 매출이 있으면 시트를 확인해 주세요.'; $wbJ.Close($false); continue }
    $v = $s.Range('A1:AB60').Value2; $hr = 0; $col = @{}
    for ($r = 1; $r -le 60 -and -not $hr; $r++) { for ($c = 1; $c -le 28; $c++) { if ("$($v[$r,$c])".Trim() -eq '인보이스') { $hr = $r } } }
    if ($hr) { for ($c = 1; $c -le 28; $c++) { $hh = ("$($v[$hr,$c])" -replace '\s+', '').ToLower(); if ($hh -and -not $col.ContainsKey($hh)) { $col[$hh] = $c } }
      foreach ($need in '인보이스', 'bl날짜', '환율', '엔화금액', '계') { if (-not $col[$need]) { Bad "'$($jf.시트)' 시트에 '$need' 머리글이 없습니다." } }
      if ($col['인보이스'] -and $col['bl날짜'] -and $col['환율'] -and $col['엔화금액'] -and $col['계']) {
        for ($r = $hr + 1; $r -le 60; $r++) { $inv = "$($v[$r,$col['인보이스']])".Trim(); if ($inv -match '합\s*계' -or -not $inv) { break }
          $bl = $v[$r, $col['bl날짜']]; $blT = $s.Cells($r, $col['bl날짜']).Text.Trim()
          $date = if ($bl -is [double] -and $bl -gt 30000) { [DateTime]::FromOADate($bl) } elseif ($blT -match '^(\d{1,2})/(\d{1,2})$') { [DateTime]::new($Y, [int]$matches[1], [int]$matches[2]) } else { $null }
          if (-not $date) { Bad "'$($jf.시트)' 시트 $($r)행 $inv 의 BL날짜('$blT')를 읽지 못했습니다."; continue }
          $JPL.Add([pscustomobject]@{ 고객 = $jf.고객; 파일 = [IO.Path]::GetFileName($jf.파일); 시트 = $jf.시트; 행 = $r; 인보이스 = $inv; BL = $date.ToString('yyyy-MM-dd'); 환율 = (N $v[$r, $col['환율']]); 엔화 = (N $v[$r, $col['엔화금액']]); 원화 = (N $v[$r, $col['계']]) }) | Out-Null } } }
    else { Bad "'$($jf.시트)' 시트에서 '인보이스' 머리글을 찾지 못했습니다." }
    $wbJ.Close($false) }
  Log "  신코·코마츠: $($JPL.Count)줄"
  # 수출 파일(바탕): 지난달 시트·환율 시트 확인, 지난달 이월(회색) 줄
  $PREVC = New-Object Collections.Generic.List[object]
  if ($orig) { $wbO = $xl.Workbooks.Open($orig, 0, $true); $so = SheetOf $wbO $prevSheet
    if (-not $so) { Bad "수출 파일에 지난달 시트 '$prevSheet'이(가) 없습니다. 이 시트를 복사해서 이번 달 시트를 만듭니다." }
    if (-not (SheetOf $wbO '환율')) { Bad "수출 파일에 '환율' 시트가 없습니다." }
    if (SheetOf $wbO $newSheet) { Bad "수출 파일에 이미 '$newSheet' 시트가 있습니다. 이번 달 시트가 없는 원본 수출 파일을 넣어 주세요." }
    if ($so) { $sH = FindRow $so 1 '^◆' 5 80
      if (-not $sH) { Bad "'$prevSheet' 시트에서 '◆ 국외매출 업체별 상세내역' 줄을 찾지 못했습니다." }
      else { $v = $so.Range("A1:K600").Value2
        for ($r = $sH + 3; $r -le 600; $r++) { if ("$($v[$r,1])" -match '총\s*계') { break }; $kk = N $v[$r, 11]
          if ($null -ne $kk -and "$($v[$r,3])".Trim()) { $PREVC.Add([pscustomobject]@{ 거래처 = "$($v[$r,1])".Trim(); 인보이스 = "$($v[$r,3])".Trim(); 면장 = "$($v[$r,4])".Trim(); 외화 = (N $v[$r, 8]); 원화 = $kk }) | Out-Null } } } }
    $wbO.Close($false); Log "  지난달 시트 '$prevSheet' 이월(회색) $($PREVC.Count)줄" }

  if ($BAD.Count) { throw "파일 확인에서 문제 $($BAD.Count)건" }
  Log '  파일 확인 끝 — 문제 없음' 'Green'
  if ($CheckOnly) { Done 'OK' "파일 확인만 했습니다. 문제 없음."; Log '파일 확인만 했습니다.' 'Green'; $code = 0; return }

  # ───────── ② 서울외국환중개 환율 ─────────
  Log '② 서울외국환중개 매매기준율 받기' 'Cyan'
  [Net.ServicePointManager]::SecurityProtocol = 'Tls12'
  $cache = @{}; if (Test-Path $rateCache) { Import-Csv $rateCache -Encoding UTF8 | % { $cache[$_.날짜] = $_ } }
  $d0r = [DateTime]::new($prv.Year, $prv.Month, 15); $today = (Get-Date).Date; $d1r = $nxt.AddMonths(1).AddDays(-1); if ($d1r -gt $today) { $d1r = $today }
  $got = 0; $fail = 0
  for ($d = $d0r; $d -le $d1r; $d = $d.AddDays(1)) { $k = $d.ToString('yyyy-MM-dd'); $c0 = $cache[$k]
    if ($c0 -and $c0.USD -ne 'ERR' -and ($c0.USD -or $d -lt $today.AddDays(-1))) { continue }
    try { $t = (Invoke-WebRequest "http://www.smbs.biz/Flash/TodayExRate_flash.jsp?tr_date=$k" -UseBasicParsing -TimeoutSec 20).Content
      $usd = if ($t -match 'USD=([\d,\.]+)') { [double]($matches[1] -replace ',', '') } else { $null }
      $jpy = if ($t -match 'JPY=([\d,\.]+)') { [double]($matches[1] -replace ',', '') / 100 } else { $null }
      $cache[$k] = [pscustomobject]@{ 날짜 = $k; 요일 = $d.ToString('ddd'); USD = $usd; JPY = $jpy }; $got++ }
    catch { $cache[$k] = [pscustomobject]@{ 날짜 = $k; 요일 = $d.ToString('ddd'); USD = 'ERR'; JPY = '' }; $fail++ } }
  $cache.Values | Sort-Object 날짜 | Export-Csv $rateCache -NoTypeInformation -Encoding UTF8
  $SMU = @{}; $SMJ = @{}; $lu = $null; $lj = $null
  foreach ($x in ($cache.Values | Sort-Object 날짜)) { if ($x.USD -and $x.USD -ne 'ERR') { $lu = [double]$x.USD; if ($x.JPY) { $lj = [double]$x.JPY } }; if ($x.USD -ne 'ERR') { $SMU[$x.날짜] = $lu; $SMJ[$x.날짜] = $lj } }
  Log "  새로 받은 날 $($got)일, 실패 $($fail)일 (저장: $rateCache)"
  if ($fail) { Note '알림' '서울외국환중개 환율 일부 못 받음' $null "인터넷 문제로 $($fail)일의 환율을 받지 못했습니다. 그날 BL은 환율 비교를 못 했습니다." '인터넷 연결을 확인하고 다시 눌러 주세요.' }

  # ───────── ④ 당월 시트 만들기 ─────────
  Log "④ '$newSheet' 시트 만들기" 'Cyan'
  $KEEP = @{}
  if (Test-Path -LiteralPath $newF) {
    Copy-Item -LiteralPath $newF (Join-Path $bak "$([IO.Path]::GetFileNameWithoutExtension($newF))_다시만들기전_$ts.xlsx")
    $kb = $xl.Workbooks.Open($newF, 0, $true); $ks = SheetOf $kb $newSheet
    if ($ks) { for ($q = 30; $q -le 600; $q++) { if ($ks.Cells($q, 1).Text -like '총*계*') { break }; $ki = "$($ks.Cells($q,3).Value2)".Trim(); if (-not $ki) { continue }
        $KEEP[$ki] = @{ 면장 = $ks.Cells($q, 4).Value2; 유니패스 = $ks.Cells($q, 5).Value2; 메모 = $ks.Cells($q, 13).Value2 } } }
    $kb.Close($false)
    Log "  기존 결과 파일을 백업하고 사용자 입력을 읽었습니다: $($KEEP.Count)줄 (유니패스 $(@($KEEP.Values | ? { $_.유니패스 }).Count)건)" 'Yellow' }
  Copy-Item -LiteralPath $orig (Join-Path $bak "$([IO.Path]::GetFileNameWithoutExtension($orig))_바탕_$ts.xlsx")
  $tmpF = Join-Path $work "만드는중_$ts.xlsx"; Copy-Item -LiteralPath $orig $tmpF -Force   # 끝까지 성공해야 결과 파일과 바꿈
  $prevInv = @{}; foreach ($p in $PREVC) { $prevInv[$p.인보이스] = $p }

  $wb = $xl.Workbooks.Open($tmpF); $src = SheetOf $wb $prevSheet
  # 지난달 시트 요약표·상세 구조 읽기
  $sHdr = FindRow $src 3 '^\s*거래처\s*$' 3 40; $sEnd = FindRow $src 2 '해외영업' ($sHdr + 1) 80
  if (-not $sHdr -or -not $sEnd) { throw "'$prevSheet' 시트 요약표(거래처 머리글·해외영업 소계 줄)를 찾지 못했습니다." }
  function Key($n) { $k = "$n".Trim(); if ($k -like 'ALLISON*') { 'ALLISON' } else { $k } }
  $fac = @{}; $nameOf = @{}; $detOrder = New-Object Collections.Generic.List[string]
  for ($q = $sHdr + 1; $q -lt $sEnd; $q++) { $n = "$($src.Cells($q,3).Value2)"; if ($n.Trim() -and $src.Cells($q, 2).Text -notmatch '소계') { $fac[(Key $n)] = $src.Cells($q, 2).Text.Trim() } }
  $sH = FindRow $src 1 '^◆' 5 80
  for ($q = $sH + 3; $q -le 600; $q++) { $n = "$($src.Cells($q,1).Value2)"; if ($n -match '총\s*계') { break }; $k = $n.Trim(); if (-not $k) { continue }
    if (-not $nameOf[$k]) { $nameOf[$k] = $n }; if (-not $detOrder.Contains($k)) { $detOrder.Add($k) } }
  foreach ($k in @($fac.Keys)) { if (-not $detOrder.Contains($k)) { $detOrder.Add($k); $nameOf[$k] = $k } }

  function CustOf($x) { $i = $x.인보이스; $n = $x.거래처
    if ($n -like '툴링비*' -or $n -like '*관세*') {
      if ($i -match '관세' -or $n -match '앨리슨|ALLISON') { return 'ALLISON' }
      if ($i -match 'DSM\d{4}-ME' -or $n -match 'MEPPI') { return 'MEPPI' }
      if ($n -match '머큐리|MERCURY') { if ($n -match '중국|CHINA') { return 'MERCURY C' }; return 'MERCURY U' }
      if ($i -match 'DSM\d{4}-AL') { return 'ALLISON' }
      return "?툴링비 거래처를 정하지 못했습니다: 거래처 칸 '$n', 인보이스 $i" }
    switch -regex ($n) { '^DOMETIC' { return 'DOMETIC' } '(?i)^POLARIS' { return 'POLARIS' } 'MERCURY CHINA' { return 'MERCURY C' } '^MERCURY' { return 'MERCURY U' } '^BERGEN' { return 'BERGEN ENG' }
      '^ALLISON' { return 'ALLISON' } '^SAM' { return 'SAM' } '^EMERSON' { return 'EMERSON' } 'AMAT' { return 'AMAT' } '^ATLAS' { return 'ATLAS KOREA' } '^MEPPI' { return 'MEPPI' } }
    foreach ($k in $detOrder) { if ($n -like "$k*") { return $k } }
    return "+$($n.Trim())" }
  $knownFac = @{ 'MEPPI' = '마산2' }

  # 이번 달 줄 목록
  $rows = New-Object Collections.Generic.List[object]; $newCust = New-Object Collections.Generic.List[string]
  foreach ($x in $LINE) { $c = CustOf $x
    if ($c.StartsWith('?')) { throw $c.Substring(1) }
    if ($c.StartsWith('+')) { $c = $c.Substring(1) }
    if (-not $detOrder.Contains($c) -and -not $newCust.Contains($c)) { $newCust.Add($c)
      $fac[$c] = if ($knownFac[$c]) { $knownFac[$c] } elseif ($x.공장 -match '^(창원|마산\d)$') { $x.공장 } else { '확인필요' }; $nameOf[$c] = $c
      $after = -1; for ($j = 0; $j -lt $detOrder.Count; $j++) { if ($fac[$detOrder[$j]] -eq $fac[$c]) { $after = $j } }
      if ($after -lt 0) { $after = [Math]::Max(0, $detOrder.IndexOf('ATLAS KOREA') - 1) }
      $detOrder.Insert($after + 1, $c) }
    $kind = if ($x.구분 -eq '이월') { '이월' } elseif ($x.인보이스 -match '관세') { '관세' } elseif ($x.거래처 -like '툴링비*') { '툴링' } elseif ($prevInv.ContainsKey($x.인보이스)) { '전월' } else { '당월' }
    $inv = $x.인보이스; if ($kind -eq '툴링' -and $inv -notmatch '툴링') { $inv = "$inv(툴링비)" }
    $dno = if ($kind -eq '관세') { '관세' } elseif ($x.면장) { $x.면장 } elseif ($c -eq 'ATLAS KOREA') { '전자세금계산서' } else { '툴링비' }
    if ($null -eq $x.외화 -and $c -ne 'ATLAS KOREA') { Note '알림' "외화 금액 없음: $($x.인보이스)" $null "라인별 매출 '$($x.시트)' 시트 $($x.행)행에 외화 금액이 비어 있습니다." '기초자료를 확인해 주세요.' }
    $rows.Add([pscustomobject]@{ 거래처 = $c; 종류 = $kind; 인보이스 = $inv; 원인보이스 = $x.인보이스; 면장 = $dno; BL = $x.BL; 환율 = $x.환율; 외화 = $x.외화; 원화 = $x.원화; 값고정 = ($c -eq 'BERGEN ENG'); 빈줄 = $false; 엔화 = $false }) | Out-Null }
  foreach ($nc in $newCust) { Note '알림' "새 거래처 $nc" $null "지난달 시트에 없는 거래처입니다. 요약표와 상세에 '$($fac[$nc])' 공장으로 줄을 새로 넣었습니다." '공장 구분과 줄 위치가 맞는지 확인해 주세요.' }
  foreach ($nm in 'SHINKO', 'JS TECH(코마츠)') { $g = @($JPL | ? 고객 -eq $nm)
    if ($g.Count -eq 0) { $rows.Add([pscustomobject]@{ 거래처 = $nm; 종류 = '당월'; 인보이스 = ''; 원인보이스 = ''; 면장 = ''; BL = ''; 환율 = $null; 외화 = $null; 원화 = $null; 값고정 = $false; 빈줄 = $true; 엔화 = $true }) | Out-Null }
    foreach ($x in $g) { $kind = if ($x.인보이스 -match 'DSM(\d{4})' -and [int]$matches[1] -lt [int]$ym4) { '전월' } else { '당월' }
      $rows.Add([pscustomobject]@{ 거래처 = $nm; 종류 = $kind; 인보이스 = $x.인보이스; 원인보이스 = $x.인보이스; 면장 = ''; BL = $x.BL; 환율 = $x.환율; 외화 = $x.엔화; 원화 = $x.원화; 값고정 = $false; 빈줄 = $false; 엔화 = $true }) | Out-Null } }
  foreach ($r0 in $rows) { if (-not $detOrder.Contains($r0.거래처)) { throw "상세 순서에 없는 거래처: $($r0.거래처)" } }
  $kOrd = @{ '전월' = 0; '당월' = 1; '툴링' = 2; '관세' = 3; '이월' = 4 }
  $i = 0; foreach ($x in $rows) { $x | Add-Member 순번 $i -Force; $i++ }
  $sorted = @($rows | Sort-Object { $detOrder.IndexOf($_.거래처) }, { $kOrd[$_.종류] }, 순번)

  # 시트 복사
  $src.Copy([Type]::Missing, $src); $ws = $wb.Worksheets.Item($src.Index + 1); $ws.Name = $newSheet
  SetV $ws.Cells(1, 1) ("※ $($Y)년 $($M)월 마감내역 ※")
  # 요약표에 새 거래처 줄
  foreach ($nc in $newCust) { $hEnd = FindRow $ws 2 '해외영업' ($sHdr + 1) 80; $side = FindRow $ws 14 '부\s*서' 3 80
    $pos = 0; for ($q = $sHdr + 1; $q -lt $hEnd; $q++) { if ($ws.Cells($q, 2).Text.Trim() -eq $fac[$nc] -and "$($ws.Cells($q,3).Value2)".Trim()) { $pos = $q + 1 } }
    if (-not $pos) { $pos = FindRow $ws 3 '^ATLAS' ($sHdr + 1) $hEnd; if (-not $pos) { $pos = $hEnd } }
    if ($side -and $pos -gt $side) { $pos = $side }
    $ws.Rows($pos).Insert() | Out-Null; $ws.Range("B$($pos - 1):K$($pos - 1)").Copy($ws.Range("B$pos")) | Out-Null
    SetV $ws.Cells($pos, 2) ($fac[$nc]); SetV $ws.Cells($pos, 3) ($nc) }
  # 상세 영역 지우고 새로 채우기
  $hdr = FindRow $ws 1 '^◆' 5 80; $d0 = $hdr + 3; $t0 = FindRow $ws 1 '총\s*계' $d0 800
  if ($t0 -gt $d0) { $ws.Range("A$($d0):A$($t0 - 1)").EntireRow.Delete() | Out-Null }
  $tb = $xl.Workbooks.Open($tplFile, 0, $true); $tsh = SheetOf $tb '서식틀'
  $tpl = @{ '전월¥' = 1; '당월¥' = 2; '이월¥' = 3; '전월$' = 4; '당월$' = 5; '툴링$' = 6; '관세$' = 6; '이월$' = 7; '툴링¥' = 2; '관세¥' = 2 }
  $restored = 0; $r = $d0
  foreach ($x in $sorted) { $cy = if ($x.엔화) { '¥' } else { '$' }
    $tsh.Rows($tpl["$($x.종류)$cy"]).Copy() | Out-Null; $ws.Rows($r).Insert(-4121) | Out-Null
    $ws.Range("A$($r):T$r").ClearContents() | Out-Null; $ws.Range("M$($r):T$r").Interior.ColorIndex = -4142
    SetV $ws.Cells($r, 1) ($nameOf[$x.거래처]); SetV $ws.Cells($r, 2) ($fac[$x.거래처])
    if (-not $x.빈줄) {
      SetV $ws.Cells($r, 3) ($x.인보이스); SetV $ws.Cells($r, 4) ($x.면장)
      if ($x.BL -match '^\d{4}-') { SetV $ws.Cells($r, 6) ([DateTime]::Parse($x.BL).ToOADate()) } else { SetV $ws.Cells($r, 6) ($x.BL) }
      if ($x.엔화) { SetV $ws.Cells($r, 7) ($x.환율) } elseif ($x.거래처 -ne 'ATLAS KOREA') { $ws.Cells($r, 7).Formula = "=VLOOKUP(F$r,환율!A:B,2,0)" }
      SetV $ws.Cells($r, 8) ($x.외화) }
    if ($x.거래처 -eq 'ATLAS KOREA' -or $x.값고정) { SetV $ws.Cells($r, 9) ($x.원화) } else { $ws.Cells($r, 9).Formula = "=ROUNDDOWN(H$r*G$r,0)" }
    if ($x.종류 -eq '이월') { $ws.Cells($r, 11).Formula = "=I$r" } else { $ws.Cells($r, 10).Formula = "=I$r" }
    $kv = if ($x.인보이스) { $KEEP[$x.인보이스] } else { $null }
    if ($kv) { $hit = $false; if ($kv.유니패스) { SetV $ws.Cells($r, 5) ($kv.유니패스); $hit = $true }; if ($kv.메모) { SetV $ws.Cells($r, 13) ($kv.메모); $hit = $true }; if ($kv.면장 -and -not $x.면장) { SetV $ws.Cells($r, 4) ($kv.면장); $hit = $true }; if ($hit) { $restored++ } }
    $x | Add-Member 행 $r -Force; $r++ }
  $tb.Close($false)
  $tot = $r; $lastCol = 11
  $ws.Range($ws.Cells($d0, 1), $ws.Cells($tot - 1, $lastCol)).Borders(12).LineStyle = 1
  $ws.Range($ws.Cells($d0, 1), $ws.Cells($tot - 1, $lastCol)).Borders(12).Weight = 2
  for ($q = $d0; $q -lt $tot; $q++) { $bb = $ws.Range($ws.Cells($q, 1), $ws.Cells($q, $lastCol)).Borders(9); $bb.LineStyle = 1
    if ("$($ws.Cells($q,1).Value2)".Trim() -ne "$($ws.Cells($q + 1,1).Value2)".Trim()) { $bb.Weight = -4138 } else { $bb.Weight = 2 } }
  foreach ($cc in 'I', 'J', 'K') { $ws.Range("$cc$tot").Formula = "=SUM($cc$($d0):$cc$($tot-1))" }
  $firstOv = ($sorted | ? { -not $_.엔화 } | Select-Object -First 1).행
  if ($firstOv) { $ws.Cells($tot + 1, 10).Formula = "=SUM(I$($firstOv):I$($tot-1))-K$tot" }
  $dm = ($sorted | ? 거래처 -eq 'DOMETIC' | Select-Object -First 1).행
  if ($dm) { $k = 14; foreach ($t in '면장', 'IN', 'PK', 'BL', '기타') { SetV $ws.Cells($dm, $k) ($t); $ws.Cells($dm, $k).Interior.Color = 65535; $k++ } }
  if ($KEEP.Count) { Log "  사용자 입력(유니패스·메모·면장) 되살린 줄: $restored" 'Yellow' }

  # 요약표
  function Rg($list, $col) { $rs = @($list | % 행 | Sort-Object); if ($rs.Count -eq 0) { return $null }; "SUM($col$($rs[0]):$col$($rs[-1]))" }
  $sumEnd = FindRow $ws 2 '해외영업' ($sHdr + 1) 80; $domSub = FindRow $ws 2 '국내영업' ($sHdr + 1) $sumEnd
  $custRows = @()
  for ($q = $sHdr + 1; $q -lt $sumEnd; $q++) { $cn = "$($ws.Cells($q,3).Value2)"; if (-not $cn.Trim() -or $ws.Cells($q, 2).Text -match '소계') { continue }; $key = Key $cn
    $g = @($sorted | ? 거래처 -eq $key); $custRows += [pscustomobject]@{ 행 = $q; 고객 = $key; 공장 = $ws.Cells($q, 2).Text }
    $fD = Rg @($g | ? 종류 -eq '전월') 'J'; $fE = Rg @($g | ? { $_.종류 -in '당월', '툴링', '관세' }) 'J'; $fH = Rg @($g | ? 종류 -eq '이월') 'K'
    if ($fD) { $ws.Cells($q, 4).Formula = "=$fD" } else { SetV $ws.Cells($q, 4) (0) }
    if ($fE) { $ws.Cells($q, 5).Formula = "=$fE" } else { SetV $ws.Cells($q, 5) (0) }
    if ($fH) { $ws.Cells($q, 8).Formula = "=$fH" } else { SetV $ws.Cells($q, 8) (0) }
    $crit = if ($key -eq 'ALLISON') { '"ALLISON"' } else { "C$q" }; $ws.Cells($q, 9).Formula = "=SUMIF(`$A`$$($d0):`$A`$$($tot),$crit,`$I`$$($d0):`$I`$$($tot))"
    $ws.Cells($q, 10).Formula = "=SUM(D$($q):G$q)" }
  foreach ($cr in $custRows) { if ($cr.고객 -ne 'ATLAS KOREA') { $ws.Cells($cr.행, 11).ClearContents() | Out-Null } }
  $runStart = $null
  for ($j = 0; $j -lt $custRows.Count; $j++) { $f = $custRows[$j]
    if (-not $runStart) { $runStart = $f.행 }; $nx = if ($j + 1 -lt $custRows.Count) { $custRows[$j + 1] } else { $null }
    if (-not $nx -or $nx.공장 -ne $f.공장 -or $nx.행 -ne $f.행 + 1) { $at2 = if ($f.고객 -eq 'ATLAS KOREA') { $f.행 - 1 } else { $f.행 }; $ws.Cells($at2, 11).Formula = $(if ($runStart -eq $f.행) { "=J$($f.행)" } else { "=SUM(J$($runStart):J$($f.행))" }); $runStart = $null } }
  $atl = ($custRows | ? 고객 -eq 'ATLAS KOREA').행; if ($atl) { SetV $ws.Cells($atl, 11) (' 국내매출 ') }
  # TOOLING 현황
  $tt = $sumEnd + 2; $toolG = @($sorted | ? 종류 -eq '툴링'); $dutyG = @($sorted | ? 종류 -eq '관세')
  $tl = @([pscustomobject]@{ 공장 = '창원'; 고객 = $nameOf['SHINKO']; 건 = 0; 식 = $null })
  foreach ($gg in $toolG | Group-Object 거래처) { $tl += [pscustomobject]@{ 공장 = $fac[$gg.Name]; 고객 = $nameOf[$gg.Name]; 건 = $gg.Count; 식 = (Rg @($gg.Group) 'J') } }
  $tl += [pscustomobject]@{ 공장 = '마산2'; 고객 = 'ALLISON'; 건 = $dutyG.Count; 식 = (Rg $dutyG 'J') }
  if ($tl.Count -gt 4) { throw "TOOLING 현황 표는 4줄인데 이번 달은 $($tl.Count)줄이 필요합니다(툴링 거래처: $(($toolG | % 거래처 | Select-Object -Unique) -join ', ')). 표 모양을 어떻게 할지 정해야 합니다." }
  $tEnd = $tt + 3
  for ($q = $tt; $q -le $tEnd; $q++) { foreach ($c in 2..5) { SetV $ws.Cells($q, $c) ($null) } }
  $q = $tt; foreach ($l in $tl) { SetV $ws.Cells($q, 2) ($l.공장); SetV $ws.Cells($q, 3) ($l.고객); SetV $ws.Cells($q, 4) ($l.건); if ($l.식) { $ws.Cells($q, 5).Formula = "=$($l.식)" } else { SetV $ws.Cells($q, 5) (0) }; $q++ }
  $ws.Cells($tt, 8).Formula = "=SUM(E$($tt):G$tEnd)"
  SetV $ws.Cells($tt, 9) (" TOOLING $($toolG.Count)건 "); SetV $ws.Cells($tt + 1, 9) (" 관세 $($dutyG.Count)건 "); SetV $ws.Cells($tt + 2, 9) (' 선급,중자0건 '); SetV $ws.Cells($tt + 3, 9) (" 총 TOOLING-$($toolG.Count + $dutyG.Count)건 ")
  if ($atl) { $ws.Cells($tt, 11).Formula = "=J$atl"
    $ov1 = $domSub + 1; $ws.Cells($tt + 1, 11).Formula = "=SUM(J$($sHdr + 1):J$($domSub - 1),J$($ov1):J$($atl - 1))" }

  # 환율 시트: 비어 있는 날짜에 라인별 매출 파일 환율 시트 값
  $hx = SheetOf $wb '환율'; $filled = @(); $hv = $hx.Range('A1:B5000').Value2
  for ($q = 1; $q -le 5000; $q++) { $dv = $hv[$q, 1]; if ($dv -is [double] -and $dv -gt 30000) { $k2 = [DateTime]::FromOADate($dv).ToString('yyyy-MM-dd'); if ($LRATE[$k2] -and ("$($hv[$q,2])".Trim() -eq '')) { SetV $hx.Cells($q, 2) ($LRATE[$k2]); $filled += $k2 } } }
  if ($filled) { Note '알림' "환율 시트 $($filled[0]) ~ $($filled[-1])" $null "수출 파일 환율 시트에 비어 있던 $($filled.Count)일을 라인별 매출 파일의 환율 시트 값으로 채웠습니다." '서울외국환중개 매매기준율과 다른 날은 이번 달 시트 환율 칸에 빨간색으로 칠해 두었습니다.' }

  # ───────── ⑤ 오른쪽 확인표 ─────────
  Log '⑤ 오른쪽 확인표' 'Cyan'
  function Blocks($list, $col) { $rs = @($list | % { [int]$_.행 } | Sort-Object); if ($rs.Count -eq 0) { return '0' }; $o = @(); $st = $rs[0]; $pv = $rs[0]
    foreach ($x in ($rs | Select-Object -Skip 1)) { if ($x -ne $pv + 1) { $o += "$col$($st):$col$pv"; $st = $x }; $pv = $x }; $o += "$col$($st):$col$pv"; $o -join ',' }
  $h = FindRow $ws 14 '부\s*서' 3 80
  if ($h) { $yen = @($sorted | ? 엔화); $tool = @($sorted | ? { $_.종류 -in '툴링', '관세' }); $atlR = @($sorted | ? 거래처 -eq 'ATLAS KOREA')
    $prod = @($sorted | ? { $_.종류 -in '전월', '당월' -and -not $_.엔화 -and $_.거래처 -ne 'ATLAS KOREA' })
    $ws.Cells($h + 1, 16).Formula = "=SUM($(Blocks $yen 'H'))"; $ws.Cells($h + 1, 17).Formula = "=SUM($(Blocks $yen 'J'))"
    $ws.Cells($h + 3, 16).Formula = "=SUM($(Blocks $prod 'H'))"; $ws.Cells($h + 3, 17).Formula = "=SUM($(Blocks $prod 'J'))"
    $ws.Cells($h + 4, 16).Formula = "=SUM($(Blocks $tool 'H'))"; $ws.Cells($h + 4, 17).Formula = "=SUM($(Blocks $tool 'J'))"
    $ws.Cells($h + 5, 16).Formula = "=SUM($(Blocks $atlR 'H'))"; $ws.Cells($h + 5, 17).Formula = "=SUM($(Blocks $atlR 'J'))"
    SetV $ws.Cells($h + 7, 14) ("$($M)월 수출매출(국내영업+해외영업) 합계") }
  else { Note '알림' '오른쪽 확인표' $null "'부 서' 머리글을 찾지 못해 오른쪽 확인표는 고치지 못했습니다." '표 위치를 확인해 주세요.' }
  $ws.Cells($tot, 11).NumberFormat = $ws.Cells($tot, 10).NumberFormat
  $xl.Calculate()
  foreach ($x in $sorted) { $x | Add-Member 계산 ($ws.Cells($x.행, 9).Value2) -Force; $x | Add-Member 계산환율 ($ws.Cells($x.행, 7).Value2) -Force }
  $sorted | Select-Object 행, 거래처, 종류, 인보이스, 원인보이스, 면장, BL, 환율, 계산환율, 외화, 원화, 계산 | Export-Csv (Join-Path $work '이번달_줄목록.csv') -NoTypeInformation -Encoding UTF8

  # ───────── ⑥ 검증 ─────────
  Log '⑥ 기초자료 검증' 'Cyan'
  $HD = 0xF2E6DD; $RED = 0xC7CEFF; $YEL = 0x99FFFF; $GRY = 0xF2F2F2
  $kName = @{ '전월' = "$($prv.Month)월에서 넘어온 매출(녹색)"; '당월' = "$($M)월 매출"; '툴링' = '툴링비'; '관세' = '관세'; '이월' = "$($nxt.Month)월로 이월(회색)" }
  $PAINT = New-Object Collections.Generic.List[object]   # 기초자료 빨간칠 목록
  foreach ($x in $LINE) { if ($x.BL -match '^\d{4}-' -and $null -ne $x.환율 -and $SMU[$x.BL] -and [Math]::Abs($x.환율 - $SMU[$x.BL]) -ge 0.001) { $PAINT.Add([pscustomobject]@{ 파일 = $fB; 시트 = $x.시트; 행 = $x.행; 열 = 6; 인보이스 = $x.인보이스; BL = $x.BL; 적힌 = $x.환율; 맞는 = $SMU[$x.BL] }) | Out-Null } }
  foreach ($x in $DET) { if ($x.BL -match '^\d{4}-' -and $null -ne $x.환율 -and $x.환율열 -and $SMU[$x.BL] -and [Math]::Abs($x.환율 - $SMU[$x.BL]) -ge 0.001) { $PAINT.Add([pscustomobject]@{ 파일 = $fA; 시트 = $x.시트; 행 = $x.행; 열 = $x.환율열; 인보이스 = $x.인보이스; BL = $x.BL; 적힌 = $x.환율; 맞는 = $SMU[$x.BL] }) | Out-Null } }
  $cmp = foreach ($wr in $sorted) { $inv = $wr.원인보이스
    $dA = $DET | ? 인보이스 -eq $inv | Select-Object -First 1; $dB = $LINE | ? 인보이스 -eq $inv | Select-Object -First 1
    $jp = if ($inv) { $JPL | ? 인보이스 -eq $inv | Select-Object -First 1 }
    $msg = @(); $lvl = '맞음'; $want = $null
    if ($jp) { $want = $SMJ[$jp.BL]; $dA = $null; $dB = $null
      if ($null -eq $want) { $msg += '서울외국환중개 엔화 환율을 아직 못 받음' }
      elseif ([Math]::Abs($jp.환율 - $want) -gt 0.00001) { $msg += "엔화 환율 $($jp.환율)이 서울외국환중개 $want 과 다름"; $lvl = '에러' }
      $calc = [Math]::Floor($jp.엔화 * $jp.환율 + 1e-7); if ([Math]::Abs($calc - $jp.원화) -ge 1) { $msg += "매출내역서 원화 $($jp.원화.ToString('N0'))가 엔화×환율 내림($($calc.ToString('N0')))과 다름(엔화 소수점 등 — 매출내역서 금액 사용)"; if ($lvl -eq '맞음') { $lvl = '확인' } }
      if ([Math]::Abs((N $wr.계산) - $jp.원화) -ge 1) { $msg += "결과 원화가 매출내역서 $($jp.원화.ToString('N0'))와 다름"; if ($lvl -eq '맞음') { $lvl = '확인' } }
      $msg += "$($jp.파일) '$($jp.시트)' 시트 $($jp.행)행 (엔화)" }
    elseif (-not $inv) { $msg += "$($M)월 매출내역서 시트가 없어 빈 줄로 둠"; $lvl = '참고' }
    else {
      if (-not $dA -and $wr.종류 -ne '이월') { $msg += '인보이스별 상세에 없음'; $lvl = '확인' }
      if ($dA -and $dB) {
        if ($dA.BL -ne $dB.BL) { $msg += "BL날짜 다름(상세 $($dA.BL), 라인별 $($dB.BL))"; $lvl = '확인' }
        if ($null -ne $dA.환율 -and $null -ne $dB.환율 -and [Math]::Abs($dA.환율 - $dB.환율) -gt 0.001) { $msg += '환율 다름'; $lvl = '확인' }
        if ($null -ne $dA.외화 -and $null -ne $dB.외화 -and [Math]::Abs($dA.외화 - $dB.외화) -gt 0.005) { $msg += '외화 다름'; $lvl = '확인' }
        if ($null -ne $dA.원화 -and $null -ne $dB.원화 -and [Math]::Abs($dA.원화 - $dB.원화) -ge 1) { $msg += '원화 다름'; $lvl = '확인' } }
      if ($wr.종류 -eq '이월' -and -not $dA) { $msg += '다음 달 이월이라 인보이스별 상세에는 없음(정상)' }
      if ($wr.BL -match '^\d{4}-' -and $wr.거래처 -ne 'ATLAS KOREA') { $want = $SMU[$wr.BL]
        if ($null -eq $want) { $msg += '서울외국환중개 환율을 아직 못 받음(미래 날짜 등)' }
        else {
          if ($null -ne $wr.환율 -and [Math]::Abs($wr.환율 - $want) -ge 0.001) { $msg += "기초자료 환율 $($wr.환율.ToString('N2'))원이 서울외국환중개 BL일 환율 $($want.ToString('N2'))원과 다름(기초자료에 빨간칠)"; $lvl = '에러' }
          if ($wr.계산환율 -is [double] -and [Math]::Abs($wr.계산환율 - $want) -ge 0.001) { $msg += "결과 시트 환율 $($wr.계산환율.ToString('N2'))원(환율 시트 값)이 서울외국환중개와 다름(빨간칠)"; $lvl = '에러' } } }
      if ($null -ne $wr.원화 -and [Math]::Abs((N $wr.계산) - $wr.원화) -ge 1) { $msg += "결과 원화가 기초자료보다 $(((N $wr.계산) - $wr.원화).ToString('N0'))원"; if ($lvl -eq '맞음') { $lvl = '확인' } } }
    if ($wr.값고정) { $msg += '원화수출이라 원화는 기초자료 금액 그대로' }
    [pscustomobject]@{ 행 = [int]$wr.행; 거래처 = $wr.거래처; 종류 = $wr.종류; 구분 = $kName[$wr.종류]; 인보이스 = $wr.인보이스; BL상세 = $(if ($dA) { $dA.BL } elseif ($jp) { $jp.BL }); BL라인 = $(if ($dB) { $dB.BL }); 환율상세 = $(if ($dA) { $dA.환율 } elseif ($jp) { $jp.환율 }); 환율라인 = $(if ($dB) { $dB.환율 }); 서울환율 = $want
      외화상세 = $(if ($dA) { $dA.외화 } elseif ($jp) { $jp.엔화 }); 외화라인 = $(if ($dB) { $dB.외화 }); 원화상세 = $(if ($dA) { $dA.원화 } elseif ($jp) { $jp.원화 }); 원화라인 = $(if ($dB) { $dB.원화 }); 작업원화 = (N $wr.계산); 결과 = $lvl; 내용 = ($msg -join ' / ') } }
  $cmp = @($cmp)
  $miss = @($DET | ? { $iv = $_.인보이스; $iv -and -not ($sorted | ? 원인보이스 -eq $iv) })
  # 자동 알림
  $errs = @($cmp | ? 결과 -eq '에러'); $chks = @($cmp | ? 결과 -eq '확인')
  if ($errs) { Note '에러' "에러 $($errs.Count)줄" $null ((($errs | Select-Object -First 15) | % { "$($_.인보이스): $($_.내용)" }) -join "`n") '해외영업팀에 확인해 주세요. 아래 5번 표의 분홍 줄입니다.' }
  else { Note '참고' '환율' $null '모든 줄의 환율이 서울외국환중개 매매기준율(휴일은 앞 영업일)과 같습니다.' '' }
  if ($chks) { Note '알림' "확인할 줄 $($chks.Count)줄" $null ((($chks | Select-Object -First 15) | % { "$($_.인보이스): $($_.내용)" }) -join "`n") '아래 5번 표의 노란 줄입니다.' }
  if ($miss) { Note '알림' "상세에만 있는 인보이스 $($miss.Count)건" (($miss | Measure-Object 원화 -Sum).Sum) (($miss | % { "$($_.시트) 시트 $($_.행)행 $($_.인보이스)" }) -join "`n") '라인별 매출에 없어 결과 시트에 넣지 않았습니다. 넣어야 하는지 확인해 주세요.' }
  $pmiss = @($PREVC | ? { $iv = $_.인보이스; -not ($sorted | ? 원인보이스 -eq $iv) })
  if ($pmiss) { Note '알림' "$($prv.Month)월 이월인데 이번 달에 없음 $($pmiss.Count)건" (($pmiss | Measure-Object 원화 -Sum).Sum) (($pmiss | % { "$($_.거래처) $($_.인보이스)" }) -join ', ') '기초자료에 없어 넣지 않았습니다. 넣어야 하면 기초자료에 넣어 달라고 하거나 알려 주세요.' }
  $pchg = @($PREVC | ? { $iv = $_.인보이스; $t = $LINE | ? 인보이스 -eq $iv | Select-Object -First 1; $t -and $null -ne $_.외화 -and $null -ne $t.외화 -and [Math]::Abs($t.외화 - $_.외화) -gt 0.005 })
  if ($pchg) { Note '알림' "이월 금액이 바뀐 건 $($pchg.Count)건" $null (($pchg | % { $iv = $_.인보이스; $t = $LINE | ? 인보이스 -eq $iv | Select-Object -First 1; "$iv : 지난달 이월 $($_.외화.ToString('N2')) → 이번 달 $($t.외화.ToString('N2'))$(if ($t.비고) { " (비고: $($t.비고))" })" }) -join "`n") '단가 변경 등이 맞는지 확인해 주세요.' }
  $uni = @($sorted | ? { -not $_.빈줄 -and -not "$($ws.Cells($_.행, 5).Value2)".Trim() }).Count
  Note '알림' '유니패스 BL일자체크 칸' $null "비어 있는 줄 $uni 개$(if ($KEEP.Count) { ", 전에 입력한 값 $restored 줄은 되살림" })." '직접 채워 주세요. BL날짜가 다르면 BL일자 칸을 고치면 환율·원화가 따라 바뀝니다.'
  if ($PAINT.Count) { Note '에러' "기초자료 환율 틀린 칸 $($PAINT.Count)곳" $null (($PAINT | % { "$([IO.Path]::GetFileName($_.파일)) '$($_.시트)' 시트 $($_.행)행 $($_.인보이스) BL $($_.BL): 적힌 $($_.적힌) → 서울외국환중개 $($_.맞는)" }) -join "`n") '값은 고치지 않고 빨간색만 칠했습니다. 해외영업팀에 확인해 주세요.' }

  try { $old = SheetOf $wb $verName; if ($old) { $old.Delete() } } catch { }
  $s = $wb.Worksheets.Add([Type]::Missing, $ws); $s.Name = $verName
  $s.Cells.Font.Name = '맑은 고딕'; $s.Cells.Font.Size = 10
  Put $s 1 @("$($Y)년 $($M)월 해외 기초자료 검증 (자동, $(Get-Date -f 'yyyy-MM-dd HH:mm'))"); $s.Cells(1, 1).Font.Size = 14; $s.Cells(1, 1).Font.Bold = $true
  Put $s 2 @("쓴 파일: 인보이스별 상세 = $([IO.Path]::GetFileName($fA)) / 라인별 매출 = $([IO.Path]::GetFileName($fB)) / 신코 = $(if ($fS) { [IO.Path]::GetFileName($fS) } else { '없음' }) / 코마츠 = $(if ($fK) { [IO.Path]::GetFileName($fK) } else { '없음' }) / 바탕 = $([IO.Path]::GetFileName($orig))")
  Put $s 3 @("결과 시트($newSheet) 원화 = ROUNDDOWN(외화×환율,0), 환율은 BL일자로 환율 시트에서 찾음. 환율 비교는 모두 서울외국환중개 매매기준율(휴일은 앞 영업일). BERGEN은 원화수출이라 원화는 기초자료 금액 그대로, ATLAS는 전자세금계산서 원화.")
  # 1. 합계
  $r = 5; Put $s $r @('1. 합계 비교'); $s.Cells($r, 1).Font.Bold = $true; $s.Cells($r, 1).Font.Size = 12; $r++
  Put $s $r @('항목', '인보이스별 상세', '라인별 매출', "$($M)월 시트", '설명'); $s.Range("A$($r):E$r").Interior.Color = $HD; $s.Range("A$($r):E$r").Font.Bold = $true; $r1 = $r; $r++
  function SumOf($list, $col) { $t = 0.0; $any = $false; foreach ($x in $list) { if ($null -ne $x.$col) { $t += [double]$x.$col; $any = $true } }; if ($any) { $t } else { $null } }
  $cProd = @($cmp | ? { $_.종류 -in '전월', '당월' -and $_.거래처 -notin 'SHINKO', 'JS TECH(코마츠)' }); $cPrev = @($cProd | ? 종류 -eq '전월')
  $cTool = @($cmp | ? 종류 -eq '툴링'); $cDuty = @($cmp | ? 종류 -eq '관세'); $cCar = @($cmp | ? 종류 -eq '이월'); $cJp = @($cmp | ? { $_.거래처 -in 'SHINKO', 'JS TECH(코마츠)' })
  function Same($a, $b, $c) { $v = @($a, $b, $c | ? { $null -ne $_ }); if ($v.Count -lt 2) { return '' }; if (($v | % { [Math]::Round($_) } | Select-Object -Unique).Count -eq 1) { '모두 같음' } else { '서로 다름 — 아래 5번 표 확인' } }
  foreach ($l in @(
      @("제품 매출 ($($prv.Month)월 이월분 + $($M)월분, 툴링·관세 제외)", (SumOf $cProd '원화상세'), (SumOf $cProd '원화라인'), (SumOf $cProd '작업원화')),
      @("  그중 $($prv.Month)월에서 넘어온 매출(녹색)", (SumOf $cPrev '원화상세'), (SumOf $cPrev '원화라인'), (SumOf $cPrev '작업원화')),
      @("툴링비 ($($cTool.Count)건)", (SumOf $cTool '원화상세'), (SumOf $cTool '원화라인'), (SumOf $cTool '작업원화')),
      @("관세 ($($cDuty.Count)건)", (SumOf $cDuty '원화상세'), (SumOf $cDuty '원화라인'), (SumOf $cDuty '작업원화')),
      @("$($nxt.Month)월로 이월(회색, $($cCar.Count)건)", $null, (SumOf $cCar '원화라인'), (SumOf $cCar '작업원화')),
      @("신코·코마츠 (엔화, $(@($cJp | ? 인보이스).Count)건)", (SumOf $cJp '원화상세'), $null, (SumOf $cJp '작업원화')))) {
    $ex = if ($l[0] -like '*이월(회색*') { '이월은 라인별 매출 이월 시트에만 있음. 매출이 아니라 이월 칸에만 넣음' } elseif ($l[0] -like '신코*') { '왼쪽 칸 = 매출내역서 금액' } else { Same $l[1] $l[2] $l[3] }
    Put $s $r @($l[0], $l[1], $l[2], $l[3], $ex); $r++ }
  $totR = FindRow $ws 2 '^\s*TOTAL\s*$' 5 80
  Put $s $r @("$($M)월 시트 당월매출 합계(TOTAL)", $null, $null, $(if ($totR) { $ws.Cells($totR, 10).Value2 }), '제품 매출 + 툴링비 + 관세 + 신코·코마츠 + ATLAS'); $s.Range("A$($r):E$r").Font.Bold = $true; $r++
  $s.Range("B$($r1):D$r").NumberFormat = '#,##0;-#,##0;'; $s.Range("A$($r1):E$($r-1)").Borders.LineStyle = 1; $s.Range("A$($r1):E$($r-1)").Borders.Color = 0xBFBFBF
  $r++
  # 2. 확인할 것
  Put $s $r @('2. 확인할 것'); $s.Cells($r, 1).Font.Bold = $true; $s.Cells($r, 1).Font.Size = 12; $r++
  Put $s $r @('구분', '대상', '금액', '내용', '할 일'); $s.Range("A$($r):E$r").Interior.Color = $HD; $s.Range("A$($r):E$r").Font.Bold = $true; $r2 = $r; $r++
  foreach ($n in ($NOTE | Sort-Object { @{ '에러' = 0; '알림' = 1; '참고' = 2 }[$_[0]] })) { Put $s $r $n; $s.Range("A$($r):E$r").Interior.Color = $(switch ($n[0]) { '에러' { $RED } '알림' { $YEL } default { $GRY } }); $r++ }
  $s.Range("C$($r2):C$r").NumberFormat = '#,##0;-#,##0;'; $s.Range("D$($r2):E$r").WrapText = $true; $s.Range("A$($r2):E$($r-1)").VerticalAlignment = -4160
  $s.Range("A$($r2):E$($r-1)").Borders.LineStyle = 1; $s.Range("A$($r2):E$($r-1)").Borders.Color = 0xBFBFBF
  $r++
  # 3. ALLISON 매출·관세 짝
  $alS = @($DET | ? 시트 -eq 'ALLISON'); $alD = @($DET | ? 시트 -eq 'ALLISON 관세')
  if ($alS -and $alD) {
    Put $s $r @('3. ALLISON 매출·관세 짝 맞춤 (인보이스별 상세 기준)'); $s.Cells($r, 1).Font.Bold = $true; $s.Cells($r, 1).Font.Size = 12; $r++
    Put $s $r @('인보이스', '매출 BL날짜', '매출 환율', '매출 외화', '매출 원화', '관세 BL날짜', '관세 환율', '관세 외화', '관세 원화', '관세/매출', "관세 원화 ($($M)월 시트)", '결과'); $s.Range("A$($r):L$r").Interior.Color = $HD; $s.Range("A$($r):L$r").Font.Bold = $true; $s.Range("A$($r):L$r").WrapText = $true; $ra = $r; $r++
    foreach ($x in $alS) { $ad = $alD | ? { $_.인보이스 -eq "$($x.인보이스)(관세)" } | Select-Object -First 1; $z = $cmp | ? 인보이스 -eq "$($x.인보이스)(관세)" | Select-Object -First 1
      $ok = if (-not $ad) { '관세 없음' } elseif ($null -ne $x.환율 -and $null -ne $ad.환율 -and [Math]::Abs($x.환율 - $ad.환율) -gt 0.001) { '관세 환율이 매출 환율과 다름' } else { '맞음' }
      Put $s $r @($x.인보이스, $x.BL, $x.환율, $x.외화, $x.원화, $ad.BL, $ad.환율, $ad.외화, $ad.원화, $(if ($ad -and $x.외화) { $ad.외화 / $x.외화 }), $z.작업원화, $ok)
      if ($ok -ne '맞음') { $s.Range("A$($r):L$r").Interior.Color = $(if ($ok -eq '관세 없음') { $GRY } else { $RED }) }; $r++ }
    foreach ($cc in 'C', 'D', 'G', 'H') { $s.Range("$cc$($ra):$cc$r").NumberFormat = '#,##0.00' }; foreach ($cc in 'E', 'I', 'K') { $s.Range("$cc$($ra):$cc$r").NumberFormat = '#,##0' }; $s.Range("J$($ra):J$r").NumberFormat = '0.00%'
    $s.Range("A$($ra):L$($r-1)").Borders.LineStyle = 1; $s.Range("A$($ra):L$($r-1)").Borders.Color = 0xBFBFBF; $r++ }
  # 4. 지난달 이월 → 이번 달
  Put $s $r @("4. $($prv.Month)월 시트 이월(회색) → $($M)월 매출(녹색) 대조"); $s.Cells($r, 1).Font.Bold = $true; $s.Cells($r, 1).Font.Size = 12; $r++
  Put $s $r @('거래처', '인보이스', "$($prv.Month)월 면장번호", "$($M)월 면장번호", "$($prv.Month)월 외화", "$($M)월 외화", "$($prv.Month)월 이월 원화", "$($M)월 원화", '결과'); $s.Range("A$($r):I$r").Interior.Color = $HD; $s.Range("A$($r):I$r").Font.Bold = $true; $r3 = $r; $r++
  foreach ($p in $PREVC) { $iv = $p.인보이스; $t = $cmp | ? 인보이스 -eq $iv | Select-Object -First 1; $lb = $LINE | ? 인보이스 -eq $iv | Select-Object -First 1
    $m10 = if ($lb) { $lb.면장 } else { '' }; $u10 = if ($lb) { $lb.외화 } else { $null }
    $res = if (-not $t) { "$($M)월 기초자료에 없음" } elseif ($null -ne $u10 -and $null -ne $p.외화 -and [Math]::Abs($u10 - $p.외화) -gt 0.005) { '외화 금액 바뀜' } elseif ($lb -and $m10 -ne $p.면장) { '면장번호 바뀜(지난달 비었으면 정상)' } else { '맞음 (원화는 이번 달 BL일 환율로 다시 계산)' }
    Put $s $r @($p.거래처, $iv, $p.면장, $m10, $p.외화, $u10, $p.원화, $(if ($t) { $t.작업원화 }), $res)
    if ($res -notlike '맞음*') { $s.Range("A$($r):I$r").Interior.Color = $YEL }; $r++ }
  $s.Range("E$($r3):F$r").NumberFormat = '#,##0.00'; $s.Range("G$($r3):H$r").NumberFormat = '#,##0'
  $s.Range("A$($r3):I$($r-1)").Borders.LineStyle = 1; $s.Range("A$($r3):I$($r-1)").Borders.Color = 0xBFBFBF; $r++
  # 5. 줄별 비교
  Put $s $r @("5. 줄별 비교 ($($M)월 시트 순서)"); $s.Cells($r, 1).Font.Bold = $true; $s.Cells($r, 1).Font.Size = 12; $r++
  Put $s $r @("$($M)월 시트 행", '거래처', '구분', '인보이스', 'BL날짜 (상세)', 'BL날짜 (라인별)', '환율 (상세)', '환율 (라인별)', '서울외국환중개 BL일 환율', '외화 (상세)', '외화 (라인별)', '원화 (상세)', '원화 (라인별)', "원화 ($($M)월 시트)", '결과', '내용')
  $s.Range("A$($r):P$r").Interior.Color = $HD; $s.Range("A$($r):P$r").Font.Bold = $true; $s.Range("A$($r):P$r").WrapText = $true; $r4 = $r; $r++
  foreach ($c in $cmp) { Put $s $r @("$($c.행)행", $c.거래처, $c.구분, $c.인보이스, $c.BL상세, $c.BL라인, $c.환율상세, $c.환율라인, $c.서울환율, $c.외화상세, $c.외화라인, $c.원화상세, $c.원화라인, $c.작업원화, $c.결과, $c.내용)
    $cl = switch ($c.결과) { '에러' { $RED } '확인' { $YEL } '참고' { $GRY } default { $null } }; if ($cl) { $s.Range("A$($r):P$r").Interior.Color = $cl }; $r++ }
  foreach ($mi in $miss) { Put $s $r @('없음', $mi.거래처, '', $mi.인보이스, $mi.BL, '', $mi.환율, $null, $(if ($mi.BL) { $SMU[$mi.BL] }), $mi.외화, $null, $mi.원화, $null, $null, '확인', "인보이스별 상세에 있는데 $($M)월 시트에 없음"); $s.Range("A$($r):P$r").Interior.Color = $YEL; $r++ }
  $s.Range("G$($r4):K$r").NumberFormat = '#,##0.00##'; $s.Range("L$($r4):N$r").NumberFormat = '#,##0;-#,##0'
  $s.Range("A$($r4):P$($r-1)").Borders.LineStyle = 1; $s.Range("A$($r4):P$($r-1)").Borders.Color = 0xBFBFBF
  $s.Range("A$($r4):P$($r-1)").AutoFilter() | Out-Null
  $wd = @{ A = 44; B = 26; C = 18; D = 60; E = 48; F = 14; G = 14; H = 14; I = 22; J = 13; K = 13; L = 14; M = 14; N = 14; O = 7; P = 70 }
  foreach ($k in $wd.Keys) { $s.Columns($k).ColumnWidth = $wd[$k] }
  $s.Range("P$($r4):P$r").WrapText = $true

  # ───────── ⑦ 결과 시트 환율 틀린 칸 빨간칠 ─────────
  $redW = 0
  foreach ($x in $sorted) { if ($x.엔화 -or $x.빈줄 -or $x.거래처 -eq 'ATLAS KOREA' -or $x.BL -notmatch '^\d{4}-') { continue }; $want = $SMU[$x.BL]; $g = $x.계산환율
    if ($want -and $g -is [double] -and [Math]::Abs($g - $want) -ge 0.001) { $ws.Cells($x.행, 7).Interior.Color = 255; $redW++ } }
  # 오류 칸 검사
  $badCells = @(); foreach ($sh in $ws, $s) { $ur = $sh.UsedRange; $vv = $ur.Value2; for ($i1 = 1; $i1 -le $vv.GetLength(0); $i1++) { for ($i2 = 1; $i2 -le $vv.GetLength(1); $i2++) { if ($vv[$i1, $i2] -is [int]) { $badCells += "$($sh.Name) $($ur.Row + $i1 - 1)행 $($ur.Column + $i2 - 1)열" } } } }
  $s.Activate(); $xl.ActiveWindow.Zoom = 90; $s.Range('A1').Select() | Out-Null
  $ws.Activate(); $ws.Range('A1').Select() | Out-Null
  $wb.Save(); $wb.Close($false)
  Move-Item -LiteralPath $tmpF $newF -Force
  Log "  저장: $newF" 'Green'

  # 기초자료 빨간칠 (값은 그대로)
  if ($PAINT.Count) {
    foreach ($f in @($PAINT | % 파일 | Select-Object -Unique)) { Copy-Item -LiteralPath $f (Join-Path $bak "$([IO.Path]::GetFileNameWithoutExtension($f))_빨간칠전_$ts.xlsx")
      $wbP = $xl.Workbooks.Open($f); foreach ($p in ($PAINT | ? 파일 -eq $f)) { (SheetOf $wbP $p.시트).Cells($p.행, $p.열).Interior.Color = 255 }; $wbP.Save(); $wbP.Close($false) }
    Log "  기초자료 환율 틀린 칸 $($PAINT.Count)곳 빨간칠 (원래 파일은 백업 폴더에)" 'Yellow' }

  # 요약
  Log '══ 끝 ══' 'Cyan'
  $sumTxt = @("결과 파일: $([IO.Path]::GetFileName($newF))", "시트: $newSheet, $verName", "줄: $($sorted.Count)개 (에러 $($errs.Count), 확인 $($chks.Count))", "결과 시트 환율 빨간칠: $($redW)칸, 기초자료 빨간칠: $($PAINT.Count)칸")
  if ($badCells) { $sumTxt += "수식 오류 칸 $($badCells.Count)곳: $(($badCells | Select-Object -First 10) -join ', ')" }
  if ($KEEP.Count) { $sumTxt += "사용자 입력 되살림: $restored 줄" }
  $sumTxt += ''; $sumTxt += '확인할 것:'; foreach ($n in $NOTE) { if ($n[0] -ne '참고') { $sumTxt += "[$($n[0])] $($n[1])" } }
  $sumTxt | % { Log $_ }
  Done 'OK' ($sumTxt -join "`r`n")
  $code = 0
}
catch {
  $m = "$($_.Exception.Message)"
  if ($BAD.Count) { $m = "파일 확인에서 문제가 있습니다:`r`n- " + ($BAD -join "`r`n- ") }
  else { $m = "중간에 멈췄습니다: $m`r`n(스크립트 $($_.InvocationInfo.ScriptLineNumber)행)" }
  Log $m 'Red'; Done 'ERROR' $m; $failed = $true
  $code = 1
}
finally { if ($xl) { try { foreach ($w in @($xl.Workbooks)) { $w.Close($false) } } catch { }; $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } }
if ($failed -and $tmpF -and (Test-Path -LiteralPath $tmpF)) { Remove-Item -LiteralPath $tmpF -Force -ErrorAction SilentlyContinue }
exit $code
