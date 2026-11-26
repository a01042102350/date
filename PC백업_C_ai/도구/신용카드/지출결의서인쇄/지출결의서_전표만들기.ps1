param(
  [string[]]$Urls = @(),      # 지출결의서 주소. 비우면 붙여넣는 창이 뜸
  [switch]$UseLastList,       # 주소 대신, 바로 전에 인쇄한 문서 목록 그대로 쓰기
  [int]$Port = 9333
)
# 지출결의서 주소 -> 카드 자동전표(업로드 엑셀 + 검토 엑셀)를 한 번에
#   1) 문서에서 지출내역·영수증 읽기   2) 국세청 사업자 상태 조회
#   3) 전표 만들기                     4) 카드 거래처코드 점검
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
function Box($m, $t = '지출결의서 전표 만들기') { [System.Windows.MessageBox]::Show($m, $t) | Out-Null }
function Die($m) { Box $m; exit 1 }

# ── 작업용 엣지 ──
$ed = 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
$ud = "$env:LOCALAPPDATA\영수증_작업용엣지"
try { $null = Invoke-RestMethod "http://127.0.0.1:$Port/json/version" -TimeoutSec 2 }
catch {
  Say '작업용 엣지를 띄웁니다...'
  Start-Process $ed -ArgumentList "--user-data-dir=`"$ud`"", "--remote-debugging-port=$Port", "--remote-allow-origins=http://127.0.0.1:$Port", '--no-first-run', '--window-size=1200,900', 'https://gw.ds-al.com/'
  Start-Sleep 6
}
Add-Type 'using System; using System.Runtime.InteropServices; public class WinT { [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h); }'
function Show-EdgeWindow {
  try {
    Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | Where-Object { $_.CommandLine -match '영수증_작업용엣지' -and $_.CommandLine -notmatch '--type=' } | ForEach-Object {
      $h = (Get-Process -Id $_.ProcessId -ErrorAction SilentlyContinue).MainWindowHandle
      if ($h -and $h -ne 0) { [WinT]::SetForegroundWindow($h) | Out-Null } }
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
      '지출결의서 전표 만들기 — 로그인이 필요합니다', 'OKCancel', 'Information')
    if ($r -ne 'OK') { Close-Tab $tab.id; exit }
    Say '로그인 다시 확인하는 중...'
    if (Test-GwLogin) { break }
  }
}
Say '로그인 확인됨.' 'Green'

# ── 주소 받기 ──
# 창을 띄우기 전에, 바로 전에 뽑은 문서 목록이 있는지 본다
$listFile = Join-Path $work '목록.json'
$lastDocs = @()
if (Test-Path $listFile) {
  try { $lastDocs = @(((Get-Content $listFile -Raw -Encoding UTF8) | ConvertFrom-Json).rows | Where-Object { $_.FORM_NM -eq '지출결의서' }) } catch { $lastDocs = @() }
}
if ($lastDocs.Count) {
  $script:lastLabel = "바로 앞에 뽑은 지출결의서 $($lastDocs.Count)건으로 만들기 (위에 주소를 넣지 않아도 됩니다)"
  $script:lastDetail = '-> ' + (($lastDocs | ForEach-Object { $_.DOC_NO }) -join ',  ')
  $script:lastOk = $true
} else {
  $script:lastLabel = '바로 앞에 뽑은 지출결의서로 만들기'
  $script:lastDetail = '-> 바로 앞에 뽑은 문서가 없습니다. 위에 주소를 넣어 주세요.'
  $script:lastOk = $false
}

function Ask-Links {
  [xml]$x = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="지출결의서 전표 만들기" Height="500" Width="800" WindowStartupLocation="CenterScreen">
  <Grid Margin="16">
    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
    <TextBlock Grid.Row="0" TextWrapping="Wrap" Margin="0,0,0,10" FontSize="13">지출결의서 문서 주소를 아래에 붙여넣으세요. 여러 개면 한 줄에 하나씩.&#10;카드 자동전표(아마란스에 올리는 엑셀)를 만들어 드립니다.</TextBlock>
    <TextBox Grid.Row="1" Name="Box" AcceptsReturn="True" TextWrapping="NoWrap" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="12" Padding="6"/>
    <Border Grid.Row="2" Margin="0,14,0,0" Padding="10,8" Background="#F4F6F8" CornerRadius="4">
      <StackPanel>
        <CheckBox Name="Chk" FontSize="13" FontWeight="Bold"/>
        <TextBlock Name="Detail" FontSize="12" Foreground="#555" Margin="22,4,0,0" TextWrapping="Wrap"/>
      </StackPanel>
    </Border>
    <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0">
      <Button Name="Ok" Content="시작" Width="100" Height="32" Margin="0,0,8,0"/>
      <Button Name="No" Content="취소" Width="100" Height="32" IsCancel="True"/>
    </StackPanel>
  </Grid>
</Window>
"@
  $script:win = [Windows.Markup.XamlReader]::Load((New-Object Xml.XmlNodeReader $x))
  $script:box = $script:win.FindName('Box')
  $script:chk = $script:win.FindName('Chk')
  $script:chk.Content = $script:lastLabel
  $script:chk.IsEnabled = $script:lastOk
  $script:win.FindName('Detail').Text = $script:lastDetail
  $script:ret = $null
  $script:win.FindName('Ok').Add_Click({ $script:ret = @{ text = $script:box.Text; last = $script:chk.IsChecked }; $script:win.Close() })
  $script:win.FindName('No').Add_Click({ $script:ret = $null; $script:win.Close() })
  $script:win.Add_ContentRendered({ $script:box.Focus() })
  $null = $script:win.ShowDialog()
  $script:ret
}
if (-not $Urls.Count -and -not $UseLastList) {
  $a = Ask-Links
  if (-not $a) { Close-Tab $tab.id; exit }
  if ($a.last) { $UseLastList = $true }
  $Urls = @($a.text -split '[\s\r\n]+' | Where-Object { $_ -match '^https?://' })
  if (-not $Urls.Count -and -not $UseLastList) { Close-Tab $tab.id; exit }
}

# ── 1단계: 목록 만들기 ──
if ($UseLastList -and -not $Urls.Count) {
  if (-not (Test-Path $listFile)) { Disconnect-Tab; Close-Tab $tab.id; Die "바로 전에 뽑은 문서 목록이 없습니다.`n주소를 넣어 주세요." }
  $rows = @(((Get-Content $listFile -Raw -Encoding UTF8) | ConvertFrom-Json).rows | Where-Object { $_.FORM_NM -eq '지출결의서' })
  Say "바로 전 목록을 씁니다: 지출결의서 $($rows.Count)건"
  $rows | ForEach-Object { "   $($_.DOC_NO)" }
} else {
  Say '문서 정보를 읽는 중...'
  $list = New-Object Collections.Generic.List[object]
  $idx = 0
  foreach ($u in $Urls) {
    $docId = ''; $formId = ''
    if ($u -match '[?&]docI[Dd]=(\d+)') { $docId = $matches[1] }
    if ($u -match '[?&]formId=(\d+)') { $formId = $matches[1] }
    if (-not $docId) { Write-Warning "문서번호를 못 찾아 건너뜁니다 -> $u"; continue }
    if (-not $formId) { $formId = '175' }
    if ($formId -ne '175') { Write-Warning "지출결의서가 아니라 건너뜁니다 (formId=$formId)"; continue }
    $du = "https://gw.ds-al.com/#/popup?MicroModuleCode=eap&docAuth=0&docID=$docId&formId=$formId&pageCode=UBA1100&callComp=UBAP002&popupUUID=$([guid]::NewGuid())"
    $null = Cdp 'Page.navigate' @{ url = 'about:blank' }; Start-Sleep -Milliseconds 300
    $null = Cdp 'Page.navigate' @{ url = $du }
    if (-not (WaitJs "document.body && (document.body.innerText||'').includes('지출내역')" 30)) { Write-Warning "문서를 못 읽어 건너뜁니다 (docId=$docId)"; continue }
    Start-Sleep 1
    $info = Js @'
(() => {
  const t = document.body.innerText || '';
  const no = (t.match(/[^\s]*\(주\)-\d{4}-\d+/) || t.match(/\d{4}-\d{4,}/) || [''])[0];
  let cnt = 0;
  [...document.querySelectorAll('div.UpDownLoader')].forEach(b => {
    const m = (b.innerText || '').match(/첨부파일\s*(\d+)\s*개/); if (m) cnt = Math.max(cnt, parseInt(m[1]));
  });
  let title = '';
  const ti = [...document.querySelectorAll('input,textarea')].map(e => (e.value || '').trim()).filter(v => v.length > 2 && v.length < 60);
  if (ti.length) title = ti[0];
  return JSON.stringify({ no: no, cnt: cnt, title: title });
})()
'@
    $o = $info | ConvertFrom-Json
    $docNo = if ($o.no) { $o.no } else { "문서$docId" }
    $list.Add([ordered]@{ idx = $idx; DOC_ID = [int]$docId; FORM_ID = [int]$formId; DOC_NO = $docNo
        DOC_TITLE = $(if ($o.title) { $o.title } else { '지출결의서' }); FORM_NM = '지출결의서'
        FILE_CNT = [int]$o.cnt; USER_NM = ''; ARRIVED_DT = '' })
    "   읽음: $docNo"
    $idx++
  }
  if (-not $list.Count) { Disconnect-Tab; Close-Tab $tab.id; Die '처리할 지출결의서가 없습니다.' }
  if (Test-Path $listFile) { Copy-Item $listFile (Join-Path $work "목록_백업_$stamp.json") -Force }
  [IO.File]::WriteAllText($listFile, ((@{ rows = $list }) | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding $false))
  Say "지출결의서 $($list.Count)건" 'Green'
}
Disconnect-Tab
Close-Tab $tab.id

# ── 2단계: 지출내역·영수증 읽기 ──
''
Say '[1/4] 지출결의서에서 지출내역과 영수증을 읽는 중...'
$dataFile = Join-Path $cardDir '자료.json'
if (Test-Path $dataFile) { Copy-Item $dataFile (Join-Path $cardDir "자료_백업_$stamp.json") -Force }
& "$S\지출결의_카드자료.ps1"
if (-not (Test-Path $dataFile)) { Die '지출내역을 읽지 못했습니다.' }

# ── 3단계: 국세청 사업자 상태 (기존 것과 합치기) ──
''
Say '[2/4] 국세청에서 가맹점 사업자 상태를 조회하는 중...'
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
    Say "   전에 조회한 $($keep.Count)곳과 합쳤습니다 (총 $((Import-Csv $bizFile -Encoding UTF8).Count)곳)"
  }
}

# ── 4단계: 전표 만들기 ──
''
Say '[3/4] 자동전표를 만드는 중...'
& "$S\지출결의_카드자동전표.ps1" -OutDir $outDir

# ── 5단계: 카드 코드 점검 ──
''
Say '[4/4] 카드 거래처코드를 점검하는 중...'
& "$S\카드코드_점검.ps1"

''
$made = Get-ChildItem $outDir -Filter '지출결의_카드자동전표_*.xlsx' | Where-Object { $_.LastWriteTime -gt (Get-Date).AddMinutes(-10) } | Sort-Object LastWriteTime -Descending
Say '다 됐습니다.' 'Green'
$made | ForEach-Object { "   $($_.Name)" }
Start-Process $outDir
$names = ($made | ForEach-Object { $_.Name }) -join "`n"
Box "자동전표를 만들었습니다.`n`n$names`n`n검토 파일의 '확인할 것' 시트를 꼭 봐 주세요.`n폴더를 열어 두었습니다."
