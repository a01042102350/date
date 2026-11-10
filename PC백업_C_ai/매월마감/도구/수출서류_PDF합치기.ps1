# 해외 수출서류(면장·인보이스·패킹리스트·BL) PDF를 한 파일로 합치기 — 매달
# 사용: powershell -ExecutionPolicy Bypass -File 수출서류_PDF합치기.ps1 -YM 202611
# 대상: 그 달 결과 파일("…_M월반영.xlsx")의 이번 달 시트에서 전월이월(녹색)·당월·툴링 줄 (다음 달 이월(회색)·관세 줄·국내(신코·코마츠·ATLAS) 제외)
# PDF: C:\ai\매월마감\해외pdf 아래 모든 달 폴더에서 인보이스 이름으로 폴더/파일을 찾음 ((X)로 시작하는 폴더, 인쇄용_ 폴더 제외)
# 순서: 수출 파일 줄 순서, 인보이스마다 면장 → 인보이스 → 패킹리스트 → BL
#  - 인보이스·패킹이 한 파일이면 쪽마다 윈도우 글자 인식(OCR)으로 제목을 읽어 인보이스 쪽을 앞으로
#  - 같은 인보이스 파일이 여러 버전이면 OCR로 금액을 읽어 수출 파일 외화 금액과 같은 버전을 고름(없으면 REV·정정·'-' 규칙)
# 결과: 해외pdf\인쇄용_YYYYMM\YYYYMM_해외_수출서류_인쇄용.pdf + …_인쇄순서.xlsx(인쇄 순서·빠진 서류·확인할 것)
param([Parameter(Mandatory = $true)][string]$YM, [string]$ResultFile, [string]$PdfRoot = 'C:\ai\매월마감\해외pdf')
$ErrorActionPreference = 'Stop'
$root = 'C:\ai\매월마감'
if ($YM -notmatch '^20\d{2}(0[1-9]|1[0-2])$') { Write-Host "마감월은 6자리 숫자로 (예: 202611). 받은 값: $YM" -ForegroundColor Red; exit 2 }
$Y = [int]$YM.Substring(0, 4); $M = [int]$YM.Substring(4, 2); $cur = [DateTime]::new($Y, $M, 1); $MM = $cur.ToString('MM')
$sheetName = $cur.ToString('yy') + '.' + $MM + '(NEW)'
$work = Join-Path $root "$YM\작업"; New-Item -ItemType Directory -Force $work | Out-Null
$logF = Join-Path $work 'PDF_실행기록.txt'; $doneF = Join-Path $work 'PDF_결과.txt'
$U8 = New-Object Text.UTF8Encoding($true)
[IO.File]::WriteAllText($logF, "$($Y)년 $($M)월 수출서류 PDF 합치기 ($(Get-Date -f 'yyyy-MM-dd HH:mm:ss'))`r`n", $U8); if (Test-Path $doneF) { Remove-Item $doneF }
function Log([string]$m, [string]$c = 'Gray') { Write-Host $m -ForegroundColor $c; [IO.File]::AppendAllText($logF, "$m`r`n", $U8) }
function Done([string]$st, [string]$msg) { [IO.File]::WriteAllText($doneF, "$st`r`n$msg", $U8) }
$outDir = Join-Path $PdfRoot "인쇄용_$YM"
$outPdf = Join-Path $outDir "$($YM)_해외_수출서류_인쇄용.pdf"; $outXls = Join-Path $outDir "$($YM)_해외_수출서류_인쇄순서.xlsx"
$qpdfDir = Join-Path $env:LOCALAPPDATA 'qpdf'
$CHECK = New-Object Collections.Generic.List[object]   # 확인할 것
function Chk($inv, $what, $text) { $CHECK.Add([pscustomobject]@{ 인보이스 = $inv; 구분 = $what; 내용 = $text }) | Out-Null; Log "  [확인] $inv $what : $text" 'Yellow' }
$code = 1; $xl = $null; $tmp = $null
try {
  Log "══ $($Y)년 $($M)월 수출서류 PDF 합치기 ══" 'Cyan'
  # ───── ① 결과 파일 ─────
  if (-not $ResultFile) { $ResultFile = Join-Path $root "$YM\$($Y)년 수출 ($($Y).$($MM))_$($M)월반영.xlsx" }
  if (-not (Test-Path -LiteralPath $ResultFile)) { throw "결과 파일이 없습니다: $ResultFile`r`n먼저 '③ 결과 만들기'로 이번 달 결과 파일을 만들어 주세요." }
  if (-not (Test-Path $PdfRoot)) { throw "PDF 폴더가 없습니다: $PdfRoot" }
  Log "  결과 파일: $([IO.Path]::GetFileName($ResultFile))"
  # ───── ② qpdf ─────
  $qpdf = @(Get-ChildItem $qpdfDir -Recurse -Filter qpdf.exe -ErrorAction SilentlyContinue | Select-Object -First 1)
  if (-not $qpdf) {
    Log '  PDF 합치는 도구(qpdf)가 없어 GitHub 공식 배포본을 받습니다.' 'Yellow'
    [Net.ServicePointManager]::SecurityProtocol = 'Tls12'
    $rel = Invoke-RestMethod 'https://api.github.com/repos/qpdf/qpdf/releases/latest' -Headers @{ 'User-Agent' = 'ps' }
    $asset = $rel.assets | ? { $_.name -match 'msvc64\.zip$' } | Select-Object -First 1
    $dl = Join-Path $env:TEMP $asset.name; Invoke-WebRequest $asset.browser_download_url -OutFile $dl -UseBasicParsing
    $h = 'sha256:' + (Get-FileHash $dl -Algorithm SHA256).Hash.ToLower()
    if ($asset.digest -and $asset.digest -ne $h) { Remove-Item $dl; throw "받은 qpdf 파일 검증값이 GitHub과 다릅니다. 쓰지 않았습니다." }
    Expand-Archive $dl $qpdfDir -Force; Remove-Item $dl
    $qpdf = @(Get-ChildItem $qpdfDir -Recurse -Filter qpdf.exe | Select-Object -First 1) }
  $qpdf = $qpdf[0].FullName
  # ───── ③ 대상 인보이스 ─────
  Log '① 대상 인보이스 읽기' 'Cyan'
  $xl = New-Object -ComObject Excel.Application; $xl.DisplayAlerts = $false
  $wb = $xl.Workbooks.Open($ResultFile, 0, $true); $ws = $null; foreach ($s in $wb.Worksheets) { if (($s.Name -replace '\s', '') -eq $sheetName) { $ws = $s } }
  if (-not $ws) { $wb.Close($false); throw "결과 파일에 '$sheetName' 시트가 없습니다." }
  $hdr = 0; for ($q = 5; $q -le 80; $q++) { if ($ws.Cells($q, 1).Text -match '^◆') { $hdr = $q; break } }
  if (-not $hdr) { $wb.Close($false); throw "'$sheetName' 시트에서 '◆ 국외매출 업체별 상세내역' 줄을 못 찾았습니다." }
  $TG = New-Object Collections.Generic.List[object]; $skipN = @{ 이월 = 0; 관세 = 0; 국내 = 0 }
  for ($q = $hdr + 3; $q -le 800; $q++) { $a1 = "$($ws.Cells($q,1).Value2)".Trim(); if ($a1 -match '총\s*계') { break }
    $inv = "$($ws.Cells($q,3).Value2)".Trim(); if (-not $inv) { continue }
    if ("$($ws.Cells($q,11).Formula)".Trim()) { $skipN.이월++; continue }
    if ($inv -match '관세') { $skipN.관세++; continue }
    if ($a1 -match '^(SHINKO|JS TECH|ATLAS)') { $skipN.국내++; continue }
    $col = $ws.Cells($q, 3).Interior.Color
    $kind = if ($col -eq 10213316) { '전월이월' } elseif ($inv -match '툴링') { '툴링' } else { '당월' }
    $amt = $ws.Cells($q, 8).Value2; $amt = if ($amt -is [double]) { [Math]::Round($amt, 2) } else { $null }
    $TG.Add([pscustomobject]@{ 행 = $q; 거래처 = $a1; 구분 = $kind; 인보이스 = ($inv -replace '\(툴링비\)', '').Trim(); 외화 = $amt }) | Out-Null }
  $wb.Close($false)
  Log "  대상 $($TG.Count)건 (빼는 것: 다음 달 이월 $($skipN.이월), 관세 줄 $($skipN.관세), 국내 $($skipN.국내))"
  if ($TG.Count -eq 0) { throw '대상 인보이스가 없습니다.' }

  # ───── OCR 준비 (윈도우 내장) ─────
  Add-Type -AssemblyName System.Runtime.WindowsRuntime
  [void][Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
  [void][Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
  [void][Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
  [void][Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
  [void][Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType = WindowsRuntime]
  $asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | ? { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
  $asTaskA = ([System.WindowsRuntimeSystemExtensions].GetMethods() | ? { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' })[0]
  function Await($op, [Type]$t) { $task = $asTask.MakeGenericMethod($t).Invoke($null, @($op)); $task.Wait(-1) | Out-Null; $task.Result }
  function AwaitA($op) { $task = $asTaskA.Invoke($null, @($op)); $task.Wait(-1) | Out-Null }
  $ocr = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages(); if (-not $ocr) { throw '윈도우 글자 인식(OCR)을 쓸 수 없습니다.' }
  $DOCS = @{}; $OCRC = @{}
  function Doc($path) { if (-not $DOCS.ContainsKey($path)) { $sf = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($path)) ([Windows.Storage.StorageFile]); $DOCS[$path] = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($sf)) ([Windows.Data.Pdf.PdfDocument]) }; $DOCS[$path] }
  function Pages($path) { [int](Doc $path).PageCount }
  function OcrText($path, [int]$pg, [double]$frac) { $key = "$path|$pg|$frac"
    if (-not $OCRC.ContainsKey($key)) { $d = Doc $path; $page = $d.GetPage($pg); $ms = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
      $opt = New-Object Windows.Data.Pdf.PdfPageRenderOptions; $opt.DestinationWidth = 1500; $opt.SourceRect = New-Object Windows.Foundation.Rect 0, 0, $page.Size.Width, ($page.Size.Height * $frac)
      AwaitA ($page.RenderToStreamAsync($ms, $opt))
      $dec = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($ms)) ([Windows.Graphics.Imaging.BitmapDecoder])
      $bmp = Await ($dec.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
      $OCRC[$key] = ((Await ($ocr.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult])).Text -replace '\s+', ' ').ToUpper() }
    $OCRC[$key] }
  function PageKind($path, [int]$pg) { $t = OcrText $path $pg 0.3; if ($t -match 'PACKING') { '패킹' } elseif ($t -match 'INVOICE') { '인보이스' } else { '모름' } }
  function HasAmount($path, $amt) { if ($null -eq $amt) { return $null }; $want = ('{0:0.00}' -f $amt) -replace '\.', ''
    for ($pg = 0; $pg -lt (Pages $path); $pg++) { $d = ((OcrText $path $pg 1.0) -replace 'O', '0' -replace '[IL]', '1') -replace '[^0-9]', ''; if ($d.Contains($want)) { return $true } }; $false }
  if ($env:OCRTEST) { $tp = 'C:\ai\매월마감\해외pdf\202610해외pdf\10월\Allison\DSM2610-AL06_22일(배)\DSM2610-AL06.pdf'; $tx = OcrText $tp 0 0.3
    Log "OCR 시험: 글자 $($tx.Length), 종류 $(PageKind $tp 0)/$(PageKind $tp 1), 금액 $(HasAmount $tp 350088.68)"
    $w = [pscustomobject]@{ 외화 = 350088.68 }; $fs = @(Get-Item -LiteralPath $tp); $cc = @($fs | % { [pscustomobject]@{ f = $_ } }); $hit = @($cc | ? { HasAmount $_.f.FullName $w.외화 }); Log "Where 안에서: $($hit.Count)"
    throw '시험 끝' }

  # ───── ④ PDF 찾기·고르기 ─────
  Log '② PDF 찾아서 서류 종류 나누기' 'Cyan'
  $allDirs = @(Get-ChildItem $PdfRoot -Recurse -Directory | ? { $_.Name -notlike '(X)*' -and $_.FullName -notmatch '\\인쇄용_' })
  $allPdf = @(Get-ChildItem $PdfRoot -Recurse -File | ? { $_.Extension -match '^\.pdf$' -and $_.FullName -notmatch '\\인쇄용_' -and $_.FullName -notmatch '\\\(X\)' })
  function Kind($n) {
    if ($n -match 'CERT|TransportLabel|SHIPPER_LETTER|INSTRUCTION|INV\(청구\)|Label|cpk|사유서|Booking|병합') { return '제외' }
    if ($n -match '(?i)^\d[\d,]*[_ ].*\d{6,} I\.pdf$|^\d{9} I\.pdf$|^DFS\d+ INV|^\d{6}대신금속') { return '제외' }   # 운송사 청구서·배차 안내
    if ($n -match '면장|신고필증') { return '면장' }
    if ($n -match '(?i)BL[_ \.]|SURBL|SUR BL|\bSUR\b|HBL|HAWB|AWB|Waybill|_WAY|Ocean Bill|Airway|Bill of Lading|DKCHI|^nnrk|^AASEL') { return 'BL' }
    if ($n -match '^[0-9a-f]{8}-[0-9a-f]{4}-') { return 'BL' }   # 특송사(UPS) 인보이스·운송장
    return '인보이스' }
  function RevScore($n) { if ($n -match '(?i)final rev ?(\d*)') { return 50 + [int]('0' + $matches[1]) }; if ($n -match '(?i)\brev(?:ision)? ?(\d*)') { return 10 + [int]('0' + $matches[1]) }; if ($n -match '정정') { return 10 }; if ($n -match '-\.pdf$') { return 5 }; 0 }
  $PARTS = New-Object Collections.Generic.List[object]; $MISS = New-Object Collections.Generic.List[object]
  $i = 0
  foreach ($w in $TG) { $i++; $inv = $w.인보이스; $re = '^' + [regex]::Escape($inv) + '(?![0-9A-Za-z]|-[A-Za-z])'
    Write-Progress -Activity 'PDF 고르는 중' -Status "$inv ($i/$($TG.Count))" -PercentComplete ($i * 100 / $TG.Count)
    $dirs = @($allDirs | ? { $_.Name -match $re })
    $files = if ($dirs.Count) { @($dirs | % { Get-ChildItem $_.FullName -File -Filter '*.pdf' }) } else { @($allPdf | ? { $_.BaseName -match $re }) }
    $c = @($files | % { [pscustomobject]@{ f = $_; k = (Kind $_.Name); s = (RevScore $_.Name) } })
    $mj = @($c | ? k -eq '면장'); if (@($mj | ? { $_.f.Name -match '정정' }).Count) { $mj = @($mj | ? { $_.f.Name -match '정정' }) }
    $ivc = @($c | ? k -eq '인보이스'); $duty = @($ivc | ? { $_.f.Name -match '관세' }); $ivc = @($ivc | ? { $_.f.Name -notmatch '관세' })
    $sepI = @($ivc | ? { $_.f.Name -match '_INV' }); $sepP = @($ivc | ? { $_.f.Name -match '_PCKL' })
    $bl = @($c | ? k -eq 'BL')
    $pick = @(); $how = ''
    if ($sepI.Count -or $sepP.Count) { $pick = @($sepI | Select-Object -First 1) + @($sepP | Select-Object -First 1); $how = '인보이스·패킹 따로 된 파일' }
    elseif ($ivc.Count -eq 1) { $pick = @($ivc[0]); $how = '파일 1개' }
    elseif ($ivc.Count -gt 1) {
      $hit = @($ivc | ? { HasAmount $_.f.FullName $w.외화 })
      if ($hit.Count -ge 1) { $pick = @($hit | Sort-Object @{ e = 's'; Descending = $true }, @{ e = { $_.f.LastWriteTime }; Descending = $true } | Select-Object -First 1); $how = "금액 $('{0:N2}' -f $w.외화) 일치" }
      else { $pick = @($ivc | Sort-Object @{ e = 's'; Descending = $true }, @{ e = { $_.f.LastWriteTime }; Descending = $true } | Select-Object -First 1); $how = '금액으로 못 가려 이름(REV·정정·-) 규칙'
        Chk $inv '버전' "인보이스 파일 $($ivc.Count)개 중 금액($('{0:N2}' -f $w.외화))이 맞는 파일을 못 찾아 '$($pick[0].f.Name)'을(를) 넣었습니다. 맞는지 확인해 주세요." }
      Log "  $inv : 인보이스 파일 $($ivc.Count)개 중 '$($pick[0].f.Name)' ($how), 안 씀: $((@($ivc | ? { $_ -ne $pick[0] }) | % { $_.f.Name }) -join ' / ')" }
    if ($pick.Count -and $how -eq '파일 1개' -and $null -ne $w.외화) { if (-not (HasAmount $pick[0].f.FullName $w.외화)) { Chk $inv '금액' "인보이스 '$($pick[0].f.Name)'에서 수출 파일 금액 $('{0:N2}' -f $w.외화)을(를) 읽지 못했습니다(글자 인식 실패이거나 금액이 다름)." } }
    $lack = @(); if (-not $mj.Count) { $lack += '면장' }; if (-not $pick.Count) { $lack += '인보이스' }; if (-not $bl.Count) { $lack += 'BL' }
    if ($lack.Count) { $MISS.Add([pscustomobject]@{ 행 = $w.행; 구분 = $w.구분; 거래처 = $w.거래처; 인보이스 = $inv; 없는서류 = ($lack -join '·'); 폴더 = $(if ($dirs.Count) { $dirs[0].FullName.Replace($PdfRoot, '') } else { '폴더 없음' }) }) | Out-Null }
    # 쪽 단위 순서
    $seq = New-Object Collections.Generic.List[object]
    foreach ($x in $mj) { $seq.Add(@(1, $x.f.FullName, '1-z', "면장 ($($x.f.Name))", (Pages $x.f.FullName))) | Out-Null }
    foreach ($x in $pick) { $np = Pages $x.f.FullName
      if ($x.f.Name -match '_PCKL') { $seq.Add(@(3, $x.f.FullName, '1-z', "패킹리스트 ($($x.f.Name))", $np)) | Out-Null; continue }
      if ($x.f.Name -match '_INV') { $seq.Add(@(2, $x.f.FullName, '1-z', "인보이스 ($($x.f.Name))", $np)) | Out-Null; continue }
      for ($pg = 0; $pg -lt $np; $pg++) { $pk = PageKind $x.f.FullName $pg
        if ($pk -eq '모름') { Chk $inv '쪽 종류' "'$($x.f.Name)' $($pg + 1)쪽이 인보이스인지 패킹리스트인지 글자 인식으로 못 가려 인보이스 쪽에 넣었습니다." }
        $ord = if ($pk -eq '패킹') { 3 } else { 2 }; $nm = if ($pk -eq '패킹') { '패킹리스트' } else { '인보이스' }
        $seq.Add(@($ord, $x.f.FullName, "$($pg + 1)", "$nm ($($x.f.Name) $($pg + 1)쪽)", 1)) | Out-Null } }
    foreach ($x in $duty) { $seq.Add(@(2.5, $x.f.FullName, '1-z', "관세 인보이스 ($($x.f.Name))", (Pages $x.f.FullName))) | Out-Null }
    foreach ($x in $bl) { $seq.Add(@(4, $x.f.FullName, '1-z', "BL ($($x.f.Name))", (Pages $x.f.FullName))) | Out-Null
      if ($x.f.Name -match '^[0-9a-f]{8}-') { Chk $inv 'BL' "'$($x.f.Name)'(특송사 인보이스·운송장으로 보임)을 BL 자리에 넣었습니다." } }
    $first = $true
    foreach ($ord in 1, 2, 2.5, 3, 4) { foreach ($p in $seq) { if ($p[0] -ne $ord) { continue }
        $PARTS.Add([pscustomobject]@{ 행 = $w.행; 구분 = $w.구분; 거래처 = $w.거래처; 인보이스 = $(if ($first) { $inv } else { '' }); 서류 = $p[3]; 경로 = $p[1]; 쪽 = $p[2]; 쪽수 = [int]$p[4] }) | Out-Null; $first = $false } } }
  Write-Progress -Activity 'PDF 고르는 중' -Completed
  foreach ($m0 in $MISS) { Log "  [빠짐] $($m0.인보이스): $($m0.없는서류) 없음 ($($m0.폴더))" 'Yellow' }

  # ───── ⑤ 합치기 ─────
  Log '③ 한 파일로 합치기' 'Cyan'
  New-Item -ItemType Directory -Force $outDir | Out-Null
  foreach ($f in $outPdf, $outXls) { if (Test-Path -LiteralPath $f) { try { $h = [IO.File]::Open($f, 'Open', 'ReadWrite', 'None'); $h.Close() } catch { throw "'$([IO.Path]::GetFileName($f))'이(가) 열려 있습니다. 닫고 다시 실행해 주세요." } } }
  $tmp = Join-Path $env:TEMP "수출서류합치기_$(Get-Date -f yyyyMMddHHmmss)"; New-Item -ItemType Directory $tmp | Out-Null
  $map = @{}; $n = 0; $argl = New-Object Collections.Generic.List[string]; $page = 0
  foreach ($p in $PARTS) { if (-not $map.ContainsKey($p.경로)) { $n++; $d = Join-Path $tmp ("s{0:D4}.pdf" -f $n); Copy-Item -LiteralPath $p.경로 $d; $map[$p.경로] = $d }
    $argl.Add($map[$p.경로]) | Out-Null; $argl.Add($p.쪽) | Out-Null
    $p | Add-Member 시작쪽 ($page + 1) -Force; $page += $p.쪽수 }
  $argFile = Join-Path $tmp 'args.txt'
  [IO.File]::WriteAllLines($argFile, [string[]](@('--empty', '--pages') + @($argl) + @('--', (Join-Path $tmp 'out.pdf'))), (New-Object Text.UTF8Encoding($false)))
  & $qpdf --warning-exit-0 "@$argFile"; if ($LASTEXITCODE -ne 0) { throw "PDF 합치기(qpdf) 실패: 종료코드 $LASTEXITCODE" }
  $got = [int]("$(& $qpdf --show-npages (Join-Path $tmp 'out.pdf') 2>$null | Select-Object -Last 1)".Trim())
  Copy-Item (Join-Path $tmp 'out.pdf') $outPdf -Force
  Log "  저장: $outPdf (총 $($got)쪽, 계산 $($page)쪽)" 'Green'
  if ($got -ne $page) { Chk '' '쪽수' "합친 PDF 쪽수($got)가 계산($page)과 다릅니다." }

  # ───── ⑥ 목록 엑셀 ─────
  function Fill($sh, $hdrs, $rows, $props) { $data = New-Object 'object[,]' ($rows.Count + 1), $hdrs.Count
    for ($c = 0; $c -lt $hdrs.Count; $c++) { $data[0, $c] = $hdrs[$c] }
    for ($r = 0; $r -lt $rows.Count; $r++) { for ($c = 0; $c -lt $props.Count; $c++) { $v = $rows[$r].($props[$c]); if ($null -ne $v) { $v = $v.psobject.BaseObject }; $data[($r + 1), $c] = $v } }
    $sh.Range($sh.Cells(1, 1), $sh.Cells($rows.Count + 1, $hdrs.Count)).Value2 = $data
    $hr = $sh.Range($sh.Cells(1, 1), $sh.Cells(1, $hdrs.Count)); $hr.Font.Bold = $true; $hr.Interior.Color = 0xF2E6DD; $sh.Columns.AutoFit() | Out-Null }
  $wb2 = $xl.Workbooks.Add(); while ($wb2.Worksheets.Count -lt 3) { $wb2.Worksheets.Add([Type]::Missing, $wb2.Worksheets.Item($wb2.Worksheets.Count)) | Out-Null }
  $s1 = $wb2.Worksheets.Item(1); $s1.Name = '인쇄 순서'
  $ord1 = @($PARTS | % { $k = [pscustomobject]@{ 구분 = $_.구분; 거래처 = $_.거래처; 인보이스 = $_.인보이스; 서류 = $_.서류; 시작쪽 = $_.시작쪽; 쪽수 = $_.쪽수 }; $k })
  Fill $s1 @('구분', '거래처', '인보이스', '서류', '시작 쪽', '쪽수') $ord1 @('구분', '거래처', '인보이스', '서류', '시작쪽', '쪽수')
  $s1.Columns('D').ColumnWidth = 75; for ($r = 0; $r -lt $ord1.Count; $r++) { if ($ord1[$r].인보이스) { $s1.Range("A$($r + 2):F$($r + 2)").Borders(8).Weight = -4138 } }
  $s2 = $wb2.Worksheets.Item(2); $s2.Name = '빠진 서류'
  if ($MISS.Count) { Fill $s2 @('구분', '거래처', '인보이스', '없는 서류', '찾아본 폴더') @($MISS) @('구분', '거래처', '인보이스', '없는서류', '폴더') } else { $s2.Range('A1').Value2 = '빠진 서류 없음' }
  $s3 = $wb2.Worksheets.Item(3); $s3.Name = '확인할 것'
  if ($CHECK.Count) { Fill $s3 @('인보이스', '구분', '내용') @($CHECK) @('인보이스', '구분', '내용'); $s3.Columns('C').ColumnWidth = 110 } else { $s3.Range('A1').Value2 = '확인할 것 없음' }
  $s1.Activate(); $wb2.SaveAs($outXls, 51); $wb2.Close($false)
  Log "  목록: $outXls" 'Green'
  $sum = @("PDF: $([IO.Path]::GetFileName($outPdf)) (총 $($got)쪽)", "대상 인보이스 $($TG.Count)건, 서류 $($PARTS.Count)개", "빠진 서류가 있는 인보이스: $($MISS.Count)건", "확인할 것: $($CHECK.Count)건")
  $sum | % { Log $_ }
  Done 'OK' ($sum -join "`r`n"); $code = 0
}
catch { $m = "중간에 멈췄습니다: $($_.Exception.Message)`r`n(스크립트 $($_.InvocationInfo.ScriptLineNumber)행)"; Log $m 'Red'; Done 'ERROR' $m; $code = 1 }
finally {
  if ($xl) { try { foreach ($w0 in @($xl.Workbooks)) { $w0.Close($false) } } catch { }; $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
  if ($tmp -and (Test-Path $tmp)) { Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue } }
exit $code
