param(
  [string[]]$Urls = @(),       # 지출결의서 문서 주소(여러 개 가능). 비우면 붙여넣는 창이 뜸
  [switch]$Print,              # 바로 인쇄 (없으면 미리보기 그림만)
  [string]$Printer = '',       # 비우면 기본 프린터
  [int]$Port = 9333
)
# 지출결의서 문서 주소를 주면 → 결재문서 + 영수증 + 첨부를 모아 문서마다 차례로 인쇄
# 기존 도구(결의서_모으기 / 결의서_인쇄)를 그대로 쓰되, 미결함 목록 대신 '주신 주소'로 목록을 만든다
$ErrorActionPreference = 'Stop'
$S = Join-Path $PSScriptRoot '지출결의서인쇄'
if (-not (Test-Path $S)) { $S = $PSScriptRoot }   # 이 파일이 지출결의서인쇄 안에 있는 경우
. "$S\cdp.ps1"
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework

function Fail($msg) {
  [System.Windows.MessageBox]::Show($msg, '지출결의서 인쇄') | Out-Null
  exit 1
}

# ── 작업용 엣지 연결 (없으면 띄움) ──
$ed = 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
$ud = "$env:LOCALAPPDATA\영수증_작업용엣지"
try { $null = Invoke-RestMethod "http://127.0.0.1:$Port/json/version" -TimeoutSec 2 }
catch {
  Write-Host '작업용 엣지를 띄웁니다...' -ForegroundColor Cyan
  Start-Process $ed -ArgumentList "--user-data-dir=`"$ud`"", "--remote-debugging-port=$Port", "--remote-allow-origins=http://127.0.0.1:$Port", '--no-first-run', '--window-size=1200,900', 'https://gw.ds-al.com/'
  Start-Sleep 6
}

Add-Type 'using System; using System.Runtime.InteropServices; public class WinG { [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h); }'
function Show-EdgeWindow {
  try {
    Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | Where-Object { $_.CommandLine -match '영수증_작업용엣지' -and $_.CommandLine -notmatch '--type=' } | ForEach-Object {
      $h = (Get-Process -Id $_.ProcessId -ErrorAction SilentlyContinue).MainWindowHandle
      if ($h -and $h -ne 0) { [WinG]::SetForegroundWindow($h) | Out-Null } }
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

Write-Host '그룹웨어 로그인 상태를 확인하는 중...' -ForegroundColor Cyan
if (-not (Test-GwLogin)) {
  while ($true) {
    Show-EdgeWindow
    $r = [System.Windows.MessageBox]::Show(
      "작업용 엣지 창에 그룹웨어 로그인 화면이 떠 있습니다.`n그 창에서 로그인해 주세요.`n`n로그인을 마치셨으면 [확인]을 누르세요.",
      '지출결의서 인쇄 — 로그인이 필요합니다', 'OKCancel', 'Information')
    if ($r -ne 'OK') { Close-Tab $tab.id; exit }
    Write-Host '로그인 다시 확인하는 중...' -ForegroundColor Cyan
    if (Test-GwLogin) { break }
  }
}
Write-Host '로그인 확인됨.' -ForegroundColor Green

# ── 주소 받기 ──
function Ask-Links {
  [xml]$x = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="지출결의서 인쇄" Height="460" Width="780" WindowStartupLocation="CenterScreen">
  <Grid Margin="14">
    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
    <TextBlock Grid.Row="0" TextWrapping="Wrap" Margin="0,0,0,10" FontSize="13">지출결의서 문서 주소를 아래에 붙여넣으세요.&#10;여러 개면 한 줄에 하나씩 넣으시면 됩니다.&#10;결재문서 + 영수증 + 첨부서류를 문서마다 차례로 뽑습니다.</TextBlock>
    <TextBox Grid.Row="1" Name="Box" AcceptsReturn="True" TextWrapping="NoWrap" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="12" Padding="6"/>
    <CheckBox Grid.Row="2" Name="Chk" Margin="2,12,0,0" FontSize="13">바로 인쇄하기 (체크 안 하면 미리보기 그림만 만듭니다)</CheckBox>
    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
      <Button Name="Ok" Content="시작" Width="100" Height="32" Margin="0,0,8,0"/>
      <Button Name="No" Content="취소" Width="100" Height="32" IsCancel="True"/>
    </StackPanel>
  </Grid>
</Window>
"@
  $script:win = [Windows.Markup.XamlReader]::Load((New-Object Xml.XmlNodeReader $x))
  $script:box = $script:win.FindName('Box')
  $script:chk = $script:win.FindName('Chk')
  $script:ret = $null
  $script:win.FindName('Ok').Add_Click({ $script:ret = @{ text = $script:box.Text; print = $script:chk.IsChecked }; $script:win.Close() })
  $script:win.FindName('No').Add_Click({ $script:ret = $null; $script:win.Close() })
  $script:win.Add_ContentRendered({ $script:box.Focus() })
  $null = $script:win.ShowDialog()
  $script:ret
}

if (-not $Urls.Count) {
  $a = Ask-Links
  if (-not $a) { Close-Tab $tab.id; exit }
  $Urls = @($a.text -split '[\s\r\n]+' | Where-Object { $_ -match '^https?://' })
  if ($a.print) { $Print = $true }
  if (-not $Urls.Count) { Close-Tab $tab.id; exit }
}

# ── 주소에서 문서번호 뽑고, 문서를 열어 문서번호·첨부 수 읽기 ──
$rows = New-Object Collections.Generic.List[object]
$idx = 0
foreach ($u in $Urls) {
  $docId = ''; $formId = ''
  if ($u -match '[?&]docI[Dd]=(\d+)') { $docId = $matches[1] }
  if ($u -match '[?&]formId=(\d+)') { $formId = $matches[1] }
  if (-not $docId) { Write-Warning "문서번호를 못 찾은 주소라 건너뜁니다 -> $u"; continue }
  if (-not $formId) { $formId = '175' }
  if ($formId -ne '175') { Write-Warning "지출결의서(175)가 아닌 문서라 건너뜁니다 (formId=$formId)"; continue }

  $du = "https://gw.ds-al.com/#/popup?MicroModuleCode=eap&docAuth=0&docID=$docId&formId=$formId&pageCode=UBA1100&callComp=UBAP002&popupUUID=$([guid]::NewGuid())"
  $null = Cdp 'Page.navigate' @{ url = 'about:blank' }; Start-Sleep -Milliseconds 300
  $null = Cdp 'Page.navigate' @{ url = $du }
  if (-not (WaitJs "document.body && (document.body.innerText||'').includes('지출내역')" 30)) {
    Write-Warning "문서를 못 읽어 건너뜁니다 (docId=$docId)"; continue
  }
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
  $rows.Add([ordered]@{
      idx = $idx; DOC_ID = [int]$docId; FORM_ID = [int]$formId; DOC_NO = $docNo
      DOC_TITLE = $(if ($o.title) { $o.title } else { '지출결의서' })
      FORM_NM = '지출결의서'; FILE_CNT = [int]$o.cnt; USER_NM = ''; ARRIVED_DT = ''
    })
  "읽음: $docNo  (첨부 $($o.cnt)개)  $($o.title)"
  $idx++
}
Disconnect-Tab
Close-Tab $tab.id
if (-not $rows.Count) { Fail '처리할 지출결의서가 없습니다.' }

# ── 목록.json 만들기 (기존 것은 백업) ──
$work = "$env:TEMP\지출결의서인쇄"
New-Item -ItemType Directory -Force $work | Out-Null
$listFile = Join-Path $work '목록.json'
if (Test-Path $listFile) { Copy-Item $listFile (Join-Path $work ("목록_백업_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.json')) -Force }
$json = (@{ rows = $rows }) | ConvertTo-Json -Depth 6
[IO.File]::WriteAllText($listFile, $json, (New-Object Text.UTF8Encoding $false))
"목록 만듦: 지출결의서 $($rows.Count)건"

# ── 기존 도구로 모으기 → 인쇄 ──
Write-Host ''
Write-Host '결재문서 / 영수증 / 첨부를 모으는 중... (문서 수에 따라 몇 분 걸립니다)' -ForegroundColor Cyan
& "$S\결의서_모으기.ps1"
if ($LASTEXITCODE -and $LASTEXITCODE -ne 0) { Fail '모으기 단계에서 멈췄습니다.' }

Write-Host ''
if ($Print) {
  Write-Host '인쇄를 보냅니다...' -ForegroundColor Cyan
  if ($Printer) { & "$S\결의서_인쇄.ps1" -PerDoc -Printer $Printer } else { & "$S\결의서_인쇄.ps1" -PerDoc }
  [System.Windows.MessageBox]::Show("지출결의서 $($rows.Count)건을 인쇄로 보냈습니다.", '지출결의서 인쇄') | Out-Null
} else {
  Write-Host '미리보기 그림을 만듭니다...' -ForegroundColor Cyan
  & "$S\결의서_인쇄.ps1" -PerDoc -Preview
  $pv = Join-Path $work '미리보기'
  if (Test-Path $pv) { Start-Process $pv }
  [System.Windows.MessageBox]::Show("미리보기 그림을 만들었습니다.`n확인하시고 인쇄하시려면 다시 실행하면서 '바로 인쇄하기'에 체크해 주세요.", '지출결의서 인쇄') | Out-Null
}
