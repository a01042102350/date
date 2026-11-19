param(
  [string[]]$Urls = @(),          # 영수증 링크들 (또는 지출결의 문서 링크)
  [string]$UrlFile = '',          # 링크를 한 줄에 하나씩 적은 텍스트 파일
  [string]$OutDir = "$env:USERPROFILE\Desktop\지출결의서\결과",
  [int]$Port = 9333,
  [switch]$Probe,                 # 화면 구조만 보고 끝냄 (개발용)
  [string]$Eval = ''              # Probe 때 실행해 볼 스크립트
)
# 작업용 엣지(로그인된 창)로 링크를 열어 영수증 부분만 그림으로 찍고, 엑셀에 A4 한 장 8장으로 정리
$ErrorActionPreference = 'Stop'
trap { Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show("영수증을 모으지 못했습니다.`n$($_.Exception.Message)", '영수증 모아 8장 정리') | Out-Null; break }
$ed = 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'; $ud = "$env:LOCALAPPDATA\영수증_작업용엣지"
if ($UrlFile) { $Urls += Get-Content $UrlFile -Encoding UTF8 | ? { $_ -match '^https?://' } }
$Urls = @($Urls | % { $_.Trim() } | ? { $_ })

# ── 작업용 엣지 연결 (없으면 띄움) ──
try { $null = Invoke-RestMethod "http://127.0.0.1:$Port/json/version" -TimeoutSec 2 }
catch { Start-Process $ed -ArgumentList "--user-data-dir=`"$ud`"", "--remote-debugging-port=$Port", "--remote-allow-origins=http://127.0.0.1:$Port", '--no-first-run', '--window-size=1200,900', 'about:blank'; Start-Sleep 4 }
# 작업용 엣지 창이 최소화돼 있으면 화면을 안 그려서 찍기가 멈춤 → 창을 살짝 복원
Add-Type 'using System; using System.Runtime.InteropServices; public class WinR { [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c); [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h); }'
Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | ? { $_.CommandLine -match '영수증_작업용엣지' -and $_.CommandLine -notmatch '--type=' } | % {
  $h = (Get-Process -Id $_.ProcessId).MainWindowHandle; if ($h -ne 0 -and [WinR]::IsIconic($h)) { [WinR]::ShowWindow($h, 4) | Out-Null } }
  $tab = Invoke-RestMethod -Method Put "http://127.0.0.1:$Port/json/new?about:blank"
$ws = New-Object System.Net.WebSockets.ClientWebSocket
$ws.ConnectAsync([Uri]$tab.webSocketDebuggerUrl, [Threading.CancellationToken]::None).Wait()
$script:id = 0
function Cdp($method, $params = @{}) {
  $script:id++; $myId = $script:id
  $msg = @{ id = $myId; method = $method; params = $params } | ConvertTo-Json -Depth 10 -Compress
  $b = [Text.Encoding]::UTF8.GetBytes($msg)
  $ws.SendAsync([ArraySegment[byte]]$b, 'Text', $true, [Threading.CancellationToken]::None).Wait()
  $buf = New-Object byte[] 1048576
  while ($true) {
    $ms = New-Object IO.MemoryStream
    do { $r = $ws.ReceiveAsync([ArraySegment[byte]]$buf, [Threading.CancellationToken]::None); if (-not $r.Wait(60000)) { throw "응답 없음: $method" }; $ms.Write($buf, 0, $r.Result.Count) } while (-not $r.Result.EndOfMessage)
    $o = [Text.Encoding]::UTF8.GetString($ms.ToArray()) | ConvertFrom-Json
    if ($o.id -eq $myId) { if ($o.error) { throw "$method 실패: $($o.error.message)" }; return $o.result }
  }
}
function Js($expr) { $r = Cdp 'Runtime.evaluate' @{ expression = $expr; returnByValue = $true; awaitPromise = $true }; if ($r.exceptionDetails) { throw "스크립트 오류: $($r.exceptionDetails.exception.description)" }; $r.result.value }
function WaitFor($cond, $sec = 15) { for ($i = 0; $i -lt $sec * 4; $i++) { if (Js $cond) { return $true }; Start-Sleep -Milliseconds 250 }; $false }

# 영수증 상자 찾기 (승인번호·거래일시가 있고, 금액 '원'까지 포함하는 가장 작은 상자) → 화면 좌표
$findBox = @'
(() => {
  const has = (e, ws) => ws.every(w => (e.innerText || '').includes(w));
  const all = [...document.querySelectorAll('body *')].filter(e => e.offsetParent !== null || e.tagName === 'BODY');
  let c = all.filter(e => has(e, ['승인번호', '거래일시']));
  if (!c.length) return null;
  let b = c[c.length - 1];
  while (b.parentElement && !/\d[\d,]*\s*원/.test(b.innerText)) b = b.parentElement;
  // 가게 이름 줄까지: 위 형제가 있으면 한 단계 더
  if (b.parentElement && b.parentElement.innerText.length < b.innerText.length * 1.4) b = b.parentElement;
  // 실제 내용(글자·아이콘·선)이 있는 곳만: 안쪽 요소들의 영역을 합침 (아래 빈 공간 제외)
  let L = 1e9, T = 1e9, R = -1e9, B = -1e9;
  [b, ...b.querySelectorAll('*')].forEach(e => {
    const leaf = !e.children.length || e.tagName === 'svg' || e.tagName === 'IMG';
    const line = getComputedStyle(e).borderBottomStyle !== 'none' || getComputedStyle(e).borderTopStyle !== 'none';
    if (!leaf && !line) return;
    if (leaf && e.tagName !== 'svg' && e.tagName !== 'IMG' && !(e.innerText || '').trim() && !line) return;
    const q = e.getBoundingClientRect(); if (!q.width || !q.height) return;
    L = Math.min(L, q.left); T = Math.min(T, q.top); R = Math.max(R, q.right); B = Math.max(B, q.bottom);
  });
  const r0 = b.getBoundingClientRect(); if (R < L) { L = r0.left; T = r0.top; R = r0.right; B = r0.bottom; }
  return { x: L + scrollX, y: T + scrollY, w: R - L, h: B - T, text: b.innerText.slice(0, 80) };
})()
'@

if ($Probe) {
  foreach ($u in $Urls) {
    $null = Cdp 'Page.navigate' @{ url = $u }; Start-Sleep 4
    "주소: $(Js 'location.href')"; "제목: $(Js 'document.title')"
    "로그인 화면?: $(Js "!!document.querySelector('input[type=password]')")"
    "영수증 상자: $((Js $findBox) | ConvertTo-Json -Compress)"
    if ($Eval) { Start-Sleep 2; "검사: $(Js $Eval)" }
    "본문 앞부분: $((Js 'document.body.innerText').Substring(0, [Math]::Min(600, (Js 'document.body.innerText').Length)))"
    "매출전표 링크 수: $(Js "[...document.querySelectorAll('body *')].filter(e=>(e.innerText||'').trim()==='신용카드매출전표' && ![...e.children].some(c=>(c.innerText||'').trim()==='신용카드매출전표')).length")"
  }
  $null = Invoke-RestMethod "http://127.0.0.1:$Port/json/close/$($tab.id)"; return
}
# ── 그룹웨어 로그인 확인 (안 돼 있으면 로그인할 때까지 기다림) ──
Add-Type -AssemblyName PresentationFramework
Add-Type 'using System; using System.Runtime.InteropServices; public class WinF { [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h); }'
function Test-GwLogin {
  $null = Cdp 'Page.navigate' @{ url = 'about:blank' }; Start-Sleep -Milliseconds 300
  $null = Cdp 'Page.navigate' @{ url = 'https://gw.ds-al.com/#/' }
  Start-Sleep 4
  -not ((("$(Js 'location.href')") -match '#/login') -or (Js "!!document.querySelector('input[type=password]')"))
}
function Show-EdgeWindow {
  try {
    $null = Cdp 'Page.bringToFront'
    Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | ? { $_.CommandLine -match '영수증_작업용엣지' -and $_.CommandLine -notmatch '--type=' } | % {
      $h = (Get-Process -Id $_.ProcessId -ErrorAction SilentlyContinue).MainWindowHandle
      if ($h -and $h -ne 0) { [WinF]::SetForegroundWindow($h) | Out-Null } }
  } catch {}
}
Write-Host '그룹웨어 로그인 상태를 확인하는 중...' -ForegroundColor Cyan
if (-not (Test-GwLogin)) {
  while ($true) {
    Show-EdgeWindow
    $r = [System.Windows.MessageBox]::Show(
      "작업용 엣지 창에 그룹웨어 로그인 화면이 떠 있습니다.`n그 창에서 로그인해 주세요.`n`n로그인을 마치셨으면 [확인]을 누르세요.`n그만두시려면 [취소]를 누르세요.",
      '영수증 모아 8장 정리 — 로그인이 필요합니다', 'OKCancel', 'Information')
    if ($r -ne 'OK') { $null = Invoke-RestMethod "http://127.0.0.1:$Port/json/close/$($tab.id)"; return }
    Write-Host '로그인 다시 확인하는 중...' -ForegroundColor Cyan
    if (Test-GwLogin) { break }
  }
}
Write-Host '로그인 확인됨.' -ForegroundColor Green

# ── 링크 받기 (여러 줄 붙여넣기 가능) ──
function Ask-Links {
  [xml]$x = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="영수증 모아 8장 정리" Height="420" Width="760" WindowStartupLocation="CenterScreen">
  <Grid Margin="14">
    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
    <TextBlock Grid.Row="0" TextWrapping="Wrap" Margin="0,0,0,10" FontSize="13">영수증(신용카드 매출전표) 주소나 지출결의 문서 주소를 아래에 붙여넣으세요.&#10;여러 개면 한 줄에 하나씩 넣으시면 됩니다. 띄어쓰기로 구분해도 됩니다.</TextBlock>
    <TextBox Grid.Row="1" Name="Box" AcceptsReturn="True" TextWrapping="NoWrap" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="12" Padding="6"/>
    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
      <Button Name="Ok" Content="시작" Width="100" Height="32" Margin="0,0,8,0"/>
      <Button Name="No" Content="취소" Width="100" Height="32" IsCancel="True"/>
    </StackPanel>
  </Grid>
</Window>
"@
  $script:win = [Windows.Markup.XamlReader]::Load((New-Object Xml.XmlNodeReader $x))
  $script:box = $script:win.FindName('Box')
  $script:ret = ''
  $script:win.FindName('Ok').Add_Click({ $script:ret = $script:box.Text; $script:win.Close() })
  $script:win.FindName('No').Add_Click({ $script:ret = ''; $script:win.Close() })
  $script:win.Add_ContentRendered({ $script:box.Focus() })
  $null = $script:win.ShowDialog()
  $script:ret
}
$fromIcon = $false
if (-not $Urls.Count) {   # 아이콘으로 실행: 링크 붙여넣는 창
  $fromIcon = $true
  $in = Ask-Links
  $Urls = @($in -split '[\s\r\n]+' | ? { $_ -match '^https?://' })
  if (-not $Urls.Count) { $null = Invoke-RestMethod "http://127.0.0.1:$Port/json/close/$($tab.id)"; return }
  Write-Host "주소 $($Urls.Count)개를 받았습니다. 영수증을 찍습니다..." -ForegroundColor Cyan
}

# ── 영수증 찍기 ──
$stamp = Get-Date -Format 'yyyyMMdd_HHmm'; $imgDir = Join-Path $env:TEMP "영수증그림_$stamp"; New-Item -ItemType Directory -Force $imgDir | Out-Null
$null = Cdp 'Page.enable'
$null = Cdp 'Page.bringToFront'                       # 창이 가려져 있으면 화면을 안 그려서 찍기가 멈춤
$null = Cdp 'Emulation.setFocusEmulationEnabled' @{ enabled = $true }
$null = Cdp 'Emulation.setDeviceMetricsOverride' @{ width = 380; height = 900; deviceScaleFactor = 3; mobile = $false }
# 지출결의 문서 링크면 그 안의 '신용카드매출전표' 링크 주소를 꺼냄 (문서 순서대로)
$list = New-Object Collections.Generic.List[string]
foreach ($u in $Urls) {
  if ($u -match 'APB1020PCard') { $list.Add($u); continue }
  $null = Cdp 'Page.navigate' @{ url = $u }
  if (-not (WaitFor "[...document.querySelectorAll('a')].some(a => (a.innerText||'').includes('매출전표'))" 25)) {
    if (Js "!!document.querySelector('input[type=password]')") { throw '작업용 엣지 창에서 그룹웨어 로그인이 풀렸습니다. 그 창에서 다시 로그인해 주세요.' }
    Write-Warning "문서에서 매출전표 링크를 못 찾음 → $u"; continue
  }
  Start-Sleep 1
  $hrefs = @(Js "[...document.querySelectorAll('a')].filter(a => (a.innerText||'').includes('매출전표') && a.getAttribute('href')).map(a => a.href)")
  "문서 '$(Js 'document.title')': 매출전표 $($hrefs.Count)개"
  $hrefs | % { $list.Add($_) }
}
$Urls = @($list | Select-Object -Unique)
if (-not $Urls.Count) { throw '찍을 영수증 링크가 없습니다.' }
$shots = New-Object Collections.Generic.List[string]; $n = 0
foreach ($u in $Urls) {
  $n++; $null = Cdp 'Page.navigate' @{ url = 'about:blank' }; Start-Sleep -Milliseconds 300   # '#' 뒤만 다른 주소는 새로 안 불러와서 빈 화면을 거침
  $null = Cdp 'Page.navigate' @{ url = $u }
  if (-not (WaitFor "document.body && document.body.innerText.includes('승인번호') && document.body.innerText.includes('거래일시')" 20)) {
    if (Js "!!document.querySelector('input[type=password]')") { throw '작업용 엣지 창에서 그룹웨어 로그인이 풀렸습니다. 그 창에서 다시 로그인해 주세요.' }
    Write-Warning "$n 번째 링크: 영수증을 못 찾아 건너뜀 → $u"; continue
  }
  Start-Sleep -Milliseconds 600
  $box = Js $findBox
  $shot = Cdp 'Page.captureScreenshot' @{ format = 'png'; captureBeyondViewport = $true; clip = @{ x = [Math]::Max(0, $box.x - 6); y = [Math]::Max(0, $box.y - 6); width = $box.w + 12; height = $box.h + 12; scale = 1 } }
  $f = Join-Path $imgDir ('{0:000}.png' -f $n); [IO.File]::WriteAllBytes($f, [Convert]::FromBase64String($shot.data)); $shots.Add($f)
  "$n/$($Urls.Count) 찍음: $(($box.text -split "`n")[0..1] -join ' / ')"
}
$null = Invoke-RestMethod "http://127.0.0.1:$Port/json/close/$($tab.id)"
if (-not $shots.Count) { throw '찍은 영수증이 없습니다.' }

# ── 엑셀에 넣고 8장씩 정리 ──
$out = Join-Path $OutDir "영수증_8장_$stamp.xlsx"
$xl = $null; for ($k = 0; $k -lt 6 -and -not $xl; $k++) { try { $xl = New-Object -ComObject Excel.Application } catch { Start-Sleep 3 } }   # 엑셀이 바쁘면 다시 시도
if (-not $xl) { throw '엑셀을 열 수 없습니다. 엑셀에 열려 있는 저장 창·알림 창을 닫고 다시 해 주세요.' }
$xl.DisplayAlerts = $false; $wb = $xl.Workbooks.Add(); $sh = $wb.Worksheets.Item(1); $sh.Name = '영수증'
$i = 0; foreach ($f in $shots) { $sh.Shapes.AddPicture($f, 0, -1, 10, 10 + $i * 50, -1, -1) | Out-Null; $i++ }
$wb.SaveAs($out, 51); $wb.Close(); $xl.Quit()
& (Join-Path $PSScriptRoot '영수증_8장정리.ps1') -Workbook $out | Out-Null
"완료: 영수증 $($shots.Count)장 → $out"
if ($fromIcon) { Start-Process $out }
