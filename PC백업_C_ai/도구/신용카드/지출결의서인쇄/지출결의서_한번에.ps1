param(
  [string]$InboxUrl = 'https://gw.ds-al.com/#/UB/UB/UBA0000?specialLnb=Y&moduleCode=UB&menuCode=UBA&pageCode=UBA2020',
  [switch]$NoAsk,          # 묻지 않고 전부 처리 (시험용)
  [switch]$DoPrint,        # NoAsk 일 때: 종이로 뽑기
  [switch]$DoVoucher,      # NoAsk 일 때: 전표 만들기
  [switch]$RealPrint,      # NoAsk 일 때: 미리보기 말고 진짜 인쇄
  [switch]$PrintOnly,      # 인쇄만 하기
  [switch]$VoucherOnly,    # 전표만 만들기
  [switch]$Preview,        # 인쇄 대신 미리보기 그림만
  [int]$Port = 9333
)
# 미결함에 있는 지출결의서를 읽어 와서, 고른 것들을
#   1) 결재문서 + 영수증 + 첨부 를 순서대로 인쇄하고
#   2) 카드 자동전표(아마란스 업로드 엑셀)까지 한 번에 만든다
$ErrorActionPreference = 'Stop'
$S = $PSScriptRoot
. "$S\cdp.ps1"
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework
$work = "$env:TEMP\지출결의서인쇄"
$cardDir = "$env:TEMP\카드전표"
$outDir = "$env:USERPROFILE\Desktop\지출결의서\결과"
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
New-Item -ItemType Directory -Force $work, $cardDir, $outDir | Out-Null

function Say($m, $c = 'Cyan') { Write-Host $m -ForegroundColor $c }
function Box($m, $t = '지출결의서 한번에') { [System.Windows.MessageBox]::Show($m, $t) | Out-Null }
function Die($m) { Box $m; exit 1 }

# ── 작업용 엣지 ──
$ed = 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
$ud = "$env:LOCALAPPDATA\영수증_작업용엣지"
try { $null = Invoke-RestMethod "http://127.0.0.1:$Port/json/version" -TimeoutSec 2 }
catch {
  Say '작업용 엣지를 띄웁니다...'
  Start-Process $ed -ArgumentList "--user-data-dir=`"$ud`"", "--remote-debugging-port=$Port", "--remote-allow-origins=http://127.0.0.1:$Port", '--no-first-run', '--window-size=1300,950', 'https://gw.ds-al.com/'
  Start-Sleep 7
}
Add-Type 'using System; using System.Runtime.InteropServices; public class WinH { [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h); }'
function Show-EdgeWindow {
  try {
    Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | Where-Object { $_.CommandLine -match '영수증_작업용엣지' -and $_.CommandLine -notmatch '--type=' } | ForEach-Object {
      $h = (Get-Process -Id $_.ProcessId -ErrorAction SilentlyContinue).MainWindowHandle
      if ($h -and $h -ne 0) { [WinH]::SetForegroundWindow($h) | Out-Null } }
  } catch {}
}
$tab = New-Tab 'about:blank'
Connect-Tab $tab
$null = Cdp 'Page.enable'
function Test-GwLogin {
  $null = Cdp 'Page.navigate' @{ url = 'about:blank' }; Start-Sleep -Milliseconds 300
  $null = Cdp 'Page.navigate' @{ url = 'https://gw.ds-al.com/#/' }
  Start-Sleep 4
  $u = "$(Js 'location.href')"
  -not (($u -match '#/login') -or (Js "!!document.querySelector('input[type=password]')"))
}
Say '그룹웨어 로그인 상태를 확인하는 중...'
if (-not (Test-GwLogin)) {
  while ($true) {
    Show-EdgeWindow
    $r = [System.Windows.MessageBox]::Show(
      "작업용 엣지 창에 그룹웨어 로그인 화면이 떠 있습니다.`n그 창에서 로그인해 주세요.`n`n로그인을 마치셨으면 [확인]을 누르세요.",
      '지출결의서 한번에 — 로그인이 필요합니다', 'OKCancel', 'Information')
    if ($r -ne 'OK') { Disconnect-Tab; Close-Tab $tab.id; exit }
    Say '로그인 다시 확인하는 중...'
    if (Test-GwLogin) { break }
  }
}
Say '로그인 확인됨.' 'Green'

# ── 미결함 열고 목록 읽기 ──
Say '미결함(결재할 문서함)을 여는 중...'
$null = Cdp 'Page.navigate' @{ url = 'about:blank' }; Start-Sleep -Milliseconds 300
$null = Cdp 'Page.navigate' @{ url = $InboxUrl }
if (-not (WaitJs "document.querySelectorAll('li[id^=readchk_]').length > 0" 40)) {
  Disconnect-Tab; Close-Tab $tab.id
  Die "미결함을 읽지 못했습니다.`n`n작업용 엣지 창에서 미결함(결재할 문서)을 직접 열어 보시고,`n문서가 보이는지 확인해 주세요."
}
Start-Sleep 2
$listJs = @'
(() => {
  const li = document.querySelector('li[id^=readchk_]'); if (!li) return null;
  const k = Object.keys(li).find(k => k.startsWith('__reactInternalInstance') || k.startsWith('__reactFiber'));
  let f = li[k];
  for (let i = 0; i < 40 && f; i++, f = f.return) { const p = f.memoizedProps; if (p && Array.isArray(p.listSet)) {
    return JSON.stringify({ rows: p.listSet.map((r, idx) => ({ idx, DOC_ID: r.DOC_ID, FORM_ID: r.FORM_ID, DOC_NO: r.DOC_NO, DOC_TITLE: r.DOC_TITLE, FORM_NM: r.FORM_NM, FILE_CNT: r.FILE_CNT, USER_NM: r.USER_NM, ARRIVED_DT: r.ARRIVED_DT })) }); } }
  return null;
})()
'@
$json = Js $listJs
Disconnect-Tab
Close-Tab $tab.id
if (-not $json) { Die '미결함 목록을 읽지 못했습니다. 화면 구조가 바뀌었을 수 있습니다.' }
$all = @(($json | ConvertFrom-Json).rows)
$docs = @($all | Where-Object { $_.FORM_NM -eq '지출결의서' })
Say "미결함 문서 $($all.Count)건 중 지출결의서 $($docs.Count)건" 'Green'
$docs | ForEach-Object { "   $($_.DOC_NO)  |  $($_.DOC_TITLE)  |  첨부 $($_.FILE_CNT)개" }
if (-not $docs.Count) { Die "미결함에 지출결의서가 없습니다.`n(미결함 문서 $($all.Count)건은 다른 양식입니다)" }

# ── 무엇을 할지 정하기 (창을 띄우지 않는다) ──
$sel = $docs
$wantPrint = $true
$wantVou = $true
$realPrint = $true
if ($PSBoundParameters.ContainsKey('DoPrint') -or $PSBoundParameters.ContainsKey('DoVoucher') -or $NoAsk) {
  $wantPrint = [bool]$DoPrint; $wantVou = [bool]$DoVoucher; $realPrint = [bool]$RealPrint
}
if ($PrintOnly) { $wantPrint = $true; $wantVou = $false }
if ($VoucherOnly) { $wantPrint = $false; $wantVou = $true }
if ($Preview) { $realPrint = $false }

# ── 고른 것만 목록으로 저장 ──
$listFile = Join-Path $work '목록.json'
if (Test-Path $listFile) { Copy-Item $listFile (Join-Path $work "목록_백업_$stamp.json") -Force }
$rows = New-Object Collections.Generic.List[object]
$i = 0
foreach ($d in $sel) {
  $rows.Add([ordered]@{ idx = $i; DOC_ID = $d.DOC_ID; FORM_ID = $d.FORM_ID; DOC_NO = $d.DOC_NO
      DOC_TITLE = $d.DOC_TITLE; FORM_NM = $d.FORM_NM; FILE_CNT = $d.FILE_CNT; USER_NM = $d.USER_NM; ARRIVED_DT = $d.ARRIVED_DT })
  $i++
}
[IO.File]::WriteAllText($listFile, ((@{ rows = $rows }) | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding $false))
''
Say "고른 지출결의서 $($sel.Count)건" 'Green'
$sel | ForEach-Object { "   $($_.DOC_NO)  $($_.DOC_TITLE)" }

$step = 0; $total = 0
if ($wantPrint) { $total += 2 }
if ($wantVou) { $total += 5 }
$done = @()

# ── 종이로 뽑기 ──
if ($wantPrint) {
  ''
  $step++; Say "[$step/$total] 결재문서 / 영수증 / 첨부를 모으는 중... (문서 수에 따라 몇 분 걸립니다)"
  & "$S\결의서_모으기.ps1"
  $step++
  if ($realPrint) {
    Say "[$step/$total] 인쇄를 보내는 중..."
    & "$S\결의서_인쇄.ps1" -PerDoc
    $done += "종이로 인쇄했습니다 ($($sel.Count)건)"
  } else {
    Say "[$step/$total] 미리보기 그림을 만드는 중..."
    & "$S\결의서_인쇄.ps1" -PerDoc -Preview
    $pv = Join-Path $work '미리보기'
    if (Test-Path $pv) { Start-Process $pv }
    $done += "미리보기 그림을 만들었습니다 (폴더를 열어 두었습니다)"
  }
}

# ── 카드 자동전표 ──
if ($wantVou) {
  ''
  $step++; Say "[$step/$total] 지출내역과 영수증을 읽는 중..."
  $dataFile = Join-Path $cardDir '자료.json'
  if (Test-Path $dataFile) { Copy-Item $dataFile (Join-Path $cardDir "자료_백업_$stamp.json") -Force }
  & "$S\지출결의_카드자료.ps1"
  if (-not (Test-Path $dataFile)) { Die '지출내역을 읽지 못했습니다.' }

  ''
  $step++; Say "[$step/$total] 국세청에서 가맹점 사업자 상태를 조회하는 중..."
  $bizFile = Join-Path $cardDir '사업자상태.csv'
  $old = @()
  if (Test-Path $bizFile) { $old = @(Import-Csv $bizFile -Encoding UTF8); Copy-Item $bizFile (Join-Path $cardDir "사업자상태_백업_$stamp.csv") -Force }
  & "$S\사업자상태_조회.ps1"
  if (Test-Path $bizFile) {
    $new = @(Import-Csv $bizFile -Encoding UTF8)
    # 이번에 없던 곳 + 이번에 '조회 실패'인데 전에는 조회됐던 곳 → 전 결과를 씀 (실패로 덮어쓰지 않게, 2026-11-26)
    $newNo = @($new | ForEach-Object { $_.사업자번호 }); $newOk = @($new | Where-Object { $_.구분 -ne '조회 실패' } | ForEach-Object { $_.사업자번호 })
    $keep = @($old | Where-Object { $_.사업자번호 -notin $newOk -and ($_.구분 -ne '조회 실패' -or $_.사업자번호 -notin $newNo) })
    if ($keep.Count) {
      $keepNo = @($keep | ForEach-Object { $_.사업자번호 })
      (@($new | Where-Object { $_.사업자번호 -notin $keepNo }) + @($keep)) | Sort-Object 사업자번호 -Unique | Export-Csv $bizFile -NoTypeInformation -Encoding UTF8
      Say "   전에 조회한 $($keep.Count)곳과 합쳤습니다"
    }
  }

  ''
  $step++; Say "[$step/$total] 아마란스 거래처 목록을 받는 중... (새 거래처 코드 매기기용)"
  try { & "$S\거래처목록_받기.ps1" } catch { Say "   거래처 목록을 못 받았습니다: $($_.Exception.Message) — 새 거래처 코드는 비워 둡니다" 'Yellow' }

  ''
  $step++; Say "[$step/$total] 자동전표를 만드는 중..."
  & "$S\지출결의_카드자동전표.ps1" -OutDir $outDir

  ''
  $step++; Say "[$step/$total] 카드 거래처코드를 점검하는 중..."
  & "$S\카드코드_점검.ps1"
  $made = @(Get-ChildItem $outDir -Filter '지출결의_*.xlsx' | Where-Object { $_.Name -match '^지출결의_(카드자동전표|새거래처)_' } | Where-Object { $_.LastWriteTime -gt (Get-Date).AddMinutes(-15) } | ForEach-Object { $_.Name })
  $done += "카드 자동전표를 만들었습니다:`n  " + ($made -join "`n  ")
  if ($made -match '새거래처') { $done += "※ 새 거래처가 있습니다. 전표를 올리기 전에 지출결의_새거래처 파일을 먼저 아마란스 [거래처엑셀업로드]에서 파일업로드 → 변환하기 하세요." }
}

''
Say '다 됐습니다.' 'Green'
Start-Process $outDir
Box ("지출결의서 $($sel.Count)건을 처리했습니다.`n`n" + ($done -join "`n`n") + "`n`n검토 파일의 '확인할 것' 시트를 꼭 봐 주세요.")