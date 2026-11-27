param(
  [string]$Voucher = '',     # 카드 자동전표 업로드 파일 (비우면 결과 폴더의 가장 최근 것)
  [string]$OutDir = "$env:USERPROFILE\Desktop\지출결의서\결과",
  [string]$TradeList = "$env:TEMP\카드전표\거래처목록.json",
  [switch]$Pause             # 바로가기로 실행할 때: 끝나면 창을 바로 닫지 않고 기다림
)
trap { Write-Host "`n멈춤: $_" -ForegroundColor Red; if ($Pause) { [void](Read-Host '창을 닫으려면 Enter') }; exit 1 }
# 새 거래처를 아마란스 거래처엑셀업로드로 등록한 뒤, 아마란스가 실제로 매긴 거래처코드로 전표 파일의 코드를 바꿈.
#  (2026-11-27: 파일에 05724~05730을 적어 올렸는데 아마란스가 00001~00007로 자동으로 매김 → 전표 검증 '오류', 사업자번호 빈칸)
#  1) 아마란스 거래처 전체를 새로 받음(거래처목록_받기.ps1, 작업용 엣지·읽기만)
#  2) '지출결의_새거래처_날짜.xlsx'의 거래처마다 사업자번호(+이름)로 아마란스에 등록된 코드를 찾음
#  3) 전표 업로드 파일과 검토 파일에서 그 거래처코드 칸만 바꿈(적요 등 손으로 고친 칸은 그대로). 바꾸기 전 파일은 결과\백업에.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
. 'C:\ai\도구\엑셀없이_xlsx읽기.ps1'

if (-not $Voucher) {
  $Voucher = (Get-ChildItem -LiteralPath $OutDir -Filter '지출결의_카드자동전표_*.xlsx' | Where-Object { $_.Name -match '^지출결의_카드자동전표_\d{8}\.xlsx$' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
  if (-not $Voucher) { throw "결과 폴더에 전표 파일이 없습니다: $OutDir" }
}
$stamp = [regex]::Match([IO.Path]::GetFileName($Voucher), '(\d{8})').Groups[1].Value
$newTr = Join-Path (Split-Path $Voucher) "지출결의_새거래처_$stamp.xlsx"
$review = Join-Path (Split-Path $Voucher) "지출결의_카드자동전표_$($stamp)_검토.xlsx"
if (-not (Test-Path -LiteralPath $newTr)) { throw "새 거래처 파일이 없습니다: $newTr (이번 전표에는 새 거래처가 없었을 수 있음)" }
foreach ($f in $Voucher, $review) { $lk = Join-Path (Split-Path $f) ('~$' + [IO.Path]::GetFileName($f)); if (Test-Path -LiteralPath $lk) { throw "엑셀에 열려 있습니다. 닫고 다시 실행해 주세요: $f" } }
Write-Host "전표 파일: $Voucher"

# 1) 아마란스 거래처 전체 새로 받기
Write-Host '아마란스 거래처 목록을 새로 받는 중... (작업용 엣지)'
& (Join-Path $PSScriptRoot '거래처목록_받기.ps1') -Out $TradeList | Write-Host
$master = @(([IO.File]::ReadAllText($TradeList, [Text.Encoding]::UTF8) | ConvertFrom-Json).rows)
$codeSet = @{}; foreach ($m in $master) { $codeSet["$($m.trCd)"] = $m }

# 2) 새 거래처 파일의 거래처 → 실제 코드
function Norm($s) { ("$s" -replace '\(주\)|㈜|주식회사|\s|[^\p{L}\p{Nd}]', '').ToLower() }
$x = Read-Xlsx $newTr 1
$hdr = ($x.Rows | Where-Object { $_.R -eq 2 }).Cells; $col = @{}; foreach ($k in $hdr.Keys) { $col["$($hdr[$k])".Trim()] = $k }
$map = [ordered]@{}; $problems = @()
foreach ($r in ($x.Rows | Where-Object { $_.R -ge 4 })) {
  $plan = "$($r.Cells[$col['TR_CD']])".Trim(); $nm = "$($r.Cells[$col['TR_NM']])".Trim(); $biz = ("$($r.Cells[$col['REG_NB']])" -replace '\D', '')
  if (-not $plan) { continue }
  $cand = @($master | Where-Object { ("$($_.regNb)" -replace '\D', '') -eq $biz -and $biz })
  if ($cand.Count -gt 1) { $byName = @($cand | Where-Object { (Norm $_.trNm) -eq (Norm $nm) }); if ($byName.Count) { $cand = $byName } }
  if ($cand.Count -gt 1) { $cand = @($cand | Sort-Object { "$($_.insertDt)" } -Descending | Select-Object -First 1) }
  if ($cand.Count -eq 1) { $map[$plan] = @("$($cand[0].trCd)", "$($cand[0].trNm)") }
  elseif ($codeSet.ContainsKey($plan) -and (Norm $codeSet[$plan].trNm) -eq (Norm $nm)) { $map[$plan] = @($plan, $nm) }
  else { $problems += "$plan $nm ($biz) — 아마란스에서 못 찾음(아직 변환하기 안 했을 수 있음)" }
}
Write-Host '--- 새 거래처 코드 (파일에 적은 코드 → 아마란스 실제 코드)'
foreach ($k in $map.Keys) { Write-Host ("  {0} → {1}  {2}" -f $k, $map[$k][0], $map[$k][1]) }
if ($problems.Count) { $problems | ForEach-Object { Write-Host "  ! $_" -ForegroundColor Yellow }; throw '못 찾은 거래처가 있어 전표 파일은 고치지 않았습니다. 거래처엑셀업로드에서 변환하기까지 했는지 확인해 주세요.' }
$chg = @{}; foreach ($k in $map.Keys) { if ($map[$k][0] -ne $k) { $chg[$k] = $map[$k][0] } }
if (-not $chg.Count) { Write-Host '코드가 파일과 같습니다. 고칠 것 없음.' -ForegroundColor Green; if ($Pause) { [void](Read-Host '창을 닫으려면 Enter') }; return }

# 3) 백업 후 전표·검토 파일의 거래처코드만 바꿈
# 임시 폴더 사본에서 고치고, 바뀐 것을 확인한 뒤 원래 자리에 덮어씀
$work = Join-Path $env:TEMP '카드전표'; New-Item -ItemType Directory -Force $work | Out-Null
$wV = Join-Path $work '코드맞추기_전표.xlsx'; $wR = Join-Path $work '코드맞추기_검토.xlsx'
Copy-Item -LiteralPath $Voucher $wV -Force; if (Test-Path -LiteralPath $review) { Copy-Item -LiteralPath $review $wR -Force }
$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$xl = New-Object -ComObject Excel.Application; $mine = @(Get-Process EXCEL | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
try {
  $xl.DisplayAlerts = $false
  $wb = $xl.Workbooks.Open($wV); if ($wb.ReadOnly) { throw '전표 파일이 읽기 전용으로 열렸습니다' }
  $sh = $wb.Worksheets.Item(1); $mc = $sh.UsedRange.Columns.Count; $last = $sh.UsedRange.Rows.Count; $cTr = 0
  for ($c = 1; $c -le $mc; $c++) { if ("$($sh.Cells(2, $c).Value2)".Trim() -eq 'TR_CD') { $cTr = $c } }
  if (-not $cTr) { throw '전표 파일에서 TR_CD 칸을 못 찾음' }
  $nV = 0
  for ($r = 4; $r -le $last; $r++) { $cell = $sh.Cells($r, $cTr); $v = "$($cell.Text)".Trim(); if ($chg.ContainsKey($v)) { $cell.NumberFormat = '@'; $cell.Value2 = [string]$chg[$v]; $nV++ } }
  $wb.Save(); $wb.Close($false)
  Write-Host "전표 파일: 거래처코드 $nV 칸 바꿈"
  if (Test-Path -LiteralPath $review) {
    $wb = $xl.Workbooks.Open($wR); $sh = $wb.Worksheets.Item(1); $mc = $sh.UsedRange.Columns.Count; $last = $sh.UsedRange.Rows.Count; $cT = 0
    for ($c = 1; $c -le $mc; $c++) { if ("$($sh.Cells(1, $c).Value2)".Trim() -eq '거래처') { $cT = $c } }
    $nR = 0
    if ($cT) { for ($r = 2; $r -le $last; $r++) { $v = "$($sh.Cells($r, $cT).Value2)"; if ($v -match '^\[(\d+)\]') { $old = $Matches[1]; if ($chg.ContainsKey($old)) { $sh.Cells($r, $cT).Value2 = $v -replace "^\[$old\]", "[$($chg[$old])]"; $nR++ } } } }
    $wb.Save(); $wb.Close($false)
    Write-Host "검토 파일: 거래처 $nR 칸 바꿈"
  }
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl); [GC]::Collect() }
Start-Sleep -Seconds 2
foreach ($p in @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $mine -contains $_.Id -and $_.MainWindowHandle -eq 0 })) { Stop-Process -Id $p.Id -Force }
# 확인: 사본에 옛 코드가 남지 않았는지, 새 코드가 들어갔는지 → 원래 자리에 덮어씀
$chk = Read-Xlsx $wV 1; $h2 = ($chk.Rows | Where-Object { $_.R -eq 2 }).Cells; $cc = ($h2.Keys | Where-Object { "$($h2[$_])".Trim() -eq 'TR_CD' } | Select-Object -First 1)
$vals = @($chk.Rows | Where-Object { $_.R -ge 4 } | ForEach-Object { "$($_.Cells[$cc])".Trim() })
$left = @($vals | Where-Object { $chg.ContainsKey($_) }); $got = @($vals | Where-Object { $chg.Values -contains $_ })
if ($nV -eq 0) {
  if ($got.Count) { Write-Host "전표 파일이 이미 실제 코드로 맞춰져 있습니다(새 거래처 $($got.Count)칸). 고칠 것 없음." -ForegroundColor Green }
  else { Write-Host '전표 파일에서 새 거래처 코드를 찾지 못했습니다. 파일을 확인해 주세요.' -ForegroundColor Yellow }
  if ($Pause) { [void](Read-Host '창을 닫으려면 Enter') }; return
}
if ($left.Count -or $got.Count -lt $nV) { throw "고친 파일 확인 실패(옛 코드 $($left.Count)칸 남음, 새 코드 $($got.Count)칸) → 원래 파일은 그대로 둠" }
$bk = Join-Path (Split-Path $Voucher) '백업'; New-Item -ItemType Directory -Force $bk | Out-Null; $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
foreach ($f in $Voucher, $review) { if (Test-Path -LiteralPath $f) { Copy-Item -LiteralPath $f (Join-Path $bk ([IO.Path]::GetFileNameWithoutExtension($f) + "_코드맞추기전_$ts.xlsx")) -Force } }
Copy-Item -LiteralPath $wV $Voucher -Force; if (Test-Path -LiteralPath $wR) { Copy-Item -LiteralPath $wR $review -Force }
Write-Host '끝. 이제 전표 파일을 자동전표처리에 올리면 됩니다.' -ForegroundColor Green
if ($Pause) { [void](Read-Host '창을 닫으려면 Enter') }
